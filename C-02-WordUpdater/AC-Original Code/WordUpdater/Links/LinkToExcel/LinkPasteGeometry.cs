using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;

namespace WordUpdater
{
    // =====================================================================
    // 「链接Excel」的几何与序列工具 —— 纯逻辑、无 UI，便于离线探针自检。
    //
    // 核心模型：把「源」和「目标」都展成一维序列，按位一一对应。
    //   · 源序列  ：子区域按出现顺序展开，每个区域内先行后列（行优先）
    //   · 目标序列：选中矩形先行后列（行优先）展开
    //   · 两序列长度必须严格相等才写入；不等即拒绝，不做任何扩展/兜底
    //
    // 这一条规则同时覆盖了：单格、连续区域、Ctrl 多区域三种源形态，
    // 且连续区域的展开结果与改造前的逐格偏移写法完全一致。
    // =====================================================================

    /// <summary>目标矩形的网格行列号（1 起始、含端点）；不保存会随文档编辑漂移的 Range 偏移。</summary>
    public sealed class LinkTargetRect
    {
        public int Row1;   // 左上角行
        public int Col1;   // 左上角列
        public int Row2;   // 右下角行
        public int Col2;   // 右下角列

        public int Rows { get { return Row2 - Row1 + 1; } }
        public int Cols { get { return Col2 - Col1 + 1; } }
        public int CellCount { get { return Rows * Cols; } }

        public override string ToString()
        {
            if (CellCount == 1)
                return "第 " + Row1 + " 行第 " + Col1 + " 列";
            return "第 " + Row1 + "~" + Row2 + " 行 × 第 " + Col1 + "~" + Col2 + " 列";
        }
    }

    /// <summary>目标侧的一个席位（绝对行列号）。</summary>
    public sealed class LinkCellPos
    {
        public int Row;
        public int Col;

        public LinkCellPos(int row, int col)
        {
            Row = row;
            Col = col;
        }
    }

    /// <summary>源侧的一个单元格（已行优先展开）。</summary>
    public sealed class LinkSourceCell
    {
        public string SheetPart = "";   // "Sheet1!"；DDE 形态为空串
        public int Row;
        public int Col;
        public bool A1Style;            // true=A1 记法，false=R1C1 记法

        public string Ref()
        {
            if (A1Style) return CellAddress.ColumnLetters(Col) + Row.ToString();
            return "R" + Row.ToString() + "C" + Col.ToString();
        }

        public override string ToString()
        {
            return SheetPart + Ref();
        }
    }

    /// <summary>解析「用户选中的单元格矩形」的结果。</summary>
    public enum TargetRectStatus
    {
        Ok,          // 规则矩形
        ReadError,   // COM 读取失败（接口缺失或疑似合并单元格）
        NoCells,     // 选区里一个单元格都没有
        OutOfTable,  // 选区跨出表格边界（含表格外正文）
        MultiTable,  // 选区横跨多张表格
        NotRect      // 不是规则矩形
    }

    internal sealed class LinkTargetCell
    {
        internal int Row, Column, GridColumn;
        internal object Cell;
    }

    public static class LinkPasteGeometry
    {
        /// <summary>
        /// 保留公开几何入口；引擎通过内部入口取得同一批已验证的实际单元格。
        /// </summary>
        public static TargetRectStatus ResolveTargetRect(object selection, object table,
                                                         out LinkTargetRect rect, out string detail)
        {
            List<LinkTargetCell> targets = null;
            try { return ResolveTargetCells(selection, table, out rect, out targets, out detail); }
            finally { ReleaseTargets(targets); }
        }

        private sealed class TableLayout
        {
            internal int Start, End;
            internal XmlNodeList Rows;
            internal XmlNamespaceManager Namespaces;
        }

        internal static TargetRectStatus ResolveTargetCells(object selection, object table,
            out LinkTargetRect rect, out List<LinkTargetCell> targets, out string detail)
        {
            rect = null; targets = null; detail = null;
            object selectionRange = null, cells = null;
            List<LinkTargetCell> found = new List<LinkTargetCell>();
            try
            {
                if (selection == null || table == null) throw new InvalidOperationException("无法获取选区或目标表格。");
                TableLayout layout = ReadLayout(table);
                selectionRange = ((dynamic)selection).Range;
                int start = (int)((dynamic)selectionRange).Start, end = (int)((dynamic)selectionRange).End;
                if (start < layout.Start || end > layout.End || end < start)
                { detail = "选区超出了这张表格的范围（包含了表格外的正文）。"; return TargetRectStatus.OutOfTable; }

                int count = -1;
                try { cells = ((dynamic)selection).Cells; count = (int)((dynamic)cells).Count; } catch { }
                if (count < 0)
                {
                    Release(cells); cells = null;
                    try { cells = ((dynamic)selectionRange).Cells; count = (int)((dynamic)cells).Count; } catch { }
                }
                if (count < 0) throw new InvalidOperationException("无法读取选区中的单元格集合。");
                if (count == 0) { detail = "没有选中任何表格单元格。"; return TargetRectStatus.NoCells; }

                var coordinates = new Dictionary<long, long>();
                var physical = new HashSet<long>();
                int minRow = int.MaxValue, maxRow = 0, minCol = int.MaxValue, maxCol = 0;
                for (int i = 1; i <= count; i++)
                {
                    object cell = null;
                    try
                    {
                        cell = WordComCollection.Item(cells, i);
                        if (cell == null) throw new InvalidOperationException("无法读取第 " + i + " 个选中单元格。");
                        LinkTargetCell target; long identity;
                        TargetRectStatus status = ReadCell(cell, layout, out target, out identity, out detail);
                        if (status != TargetRectStatus.Ok) return status;
                        long key = Pair(target.Row, target.GridColumn), previous;
                        if (coordinates.TryGetValue(key, out previous))
                        {
                            if (previous == identity) continue;
                            detail = "同一行列位置对应了多个单元格。"; return TargetRectStatus.NotRect;
                        }
                        if (!physical.Add(identity))
                        { detail = "多个目标位置指向同一个合并单元格。"; return TargetRectStatus.NotRect; }
                        coordinates.Add(key, identity);
                        found.Add(target); cell = null; // 本次取得的 Cell 引用交由 targets 持有。
                        minRow = Math.Min(minRow, target.Row); maxRow = Math.Max(maxRow, target.Row);
                        minCol = Math.Min(minCol, target.GridColumn); maxCol = Math.Max(maxCol, target.GridColumn);
                    }
                    finally { Release(cell); }
                }
                long area = (long)(maxRow - minRow + 1) * (maxCol - minCol + 1);
                if (area != found.Count)
                {
                    detail = "所选区域的单元格未形成连续且列边界对齐的矩形。";
                    return TargetRectStatus.NotRect;
                }
                found.Sort(delegate(LinkTargetCell a, LinkTargetCell b) {
                    int row = a.Row.CompareTo(b.Row); return row != 0 ? row : a.GridColumn.CompareTo(b.GridColumn);
                });
                rect = new LinkTargetRect { Row1 = minRow, Row2 = maxRow, Col1 = minCol, Col2 = maxCol };
                targets = found; found = null;
                return TargetRectStatus.Ok;
            }
            catch (Exception error)
            { detail = "无法确认所选区域的单元格结构：" + error.Message; return TargetRectStatus.ReadError; }
            finally { ReleaseTargets(found); Release(cells); Release(selectionRange); }
        }

        internal static TargetRectStatus RevalidateTargets(object table, List<LinkTargetCell> targets, out string detail)
        {
            detail = null;
            try
            {
                TableLayout layout = ReadLayout(table);
                var physical = new HashSet<long>();
                foreach (LinkTargetCell expected in targets)
                {
                    LinkTargetCell current; long identity;
                    TargetRectStatus status = ReadCell(expected.Cell, layout, out current, out identity, out detail);
                    if (status != TargetRectStatus.Ok) return status;
                    if (current.Row != expected.Row || current.Column != expected.Column
                        || current.GridColumn != expected.GridColumn || !physical.Add(identity))
                    { detail = "目标单元格在剪贴板探测后已改变，请重新选择区域后重试。"; return TargetRectStatus.NotRect; }
                }
                return TargetRectStatus.Ok;
            }
            catch (Exception error)
            { detail = "写入前无法重新确认目标单元格：" + error.Message; return TargetRectStatus.ReadError; }
        }

        private static TableLayout ReadLayout(object table)
        {
            object range = null;
            try
            {
                range = ((dynamic)table).Range;
                int start = (int)((dynamic)range).Start, end = (int)((dynamic)range).End;
                if (start < 0 || end <= start) throw new InvalidOperationException("目标表格边界不可读。");
                // WPS 的单格 WordOpenXML 会重建网格并丢掉 gridSpan；父表 XML 只用于定位所选 tc。
                string xml = (string)((dynamic)range).WordOpenXML;
                var document = new XmlDocument { XmlResolver = null };
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
                using (var input = new StringReader(xml))
                using (var reader = XmlReader.Create(input, settings)) document.Load(reader);
                var ns = new XmlNamespaceManager(document.NameTable);
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                XmlNode root = document.DocumentElement;
                if (root == null || root.LocalName != "tbl" || root.NamespaceURI != ns.LookupNamespace("w"))
                {
                    XmlNodeList tables = document.SelectNodes("//w:body/w:tbl", ns);
                    if (tables.Count != 1) throw new InvalidOperationException("无法唯一定位目标表格结构。");
                    root = tables[0];
                }
                XmlNodeList rows = root.SelectNodes("w:tr", ns);
                if (rows.Count == 0) throw new InvalidOperationException("目标表格没有可读取的行结构。");
                return new TableLayout { Start = start, End = end, Rows = rows, Namespaces = ns };
            }
            finally { Release(range); }
        }

        private static TargetRectStatus ReadCell(object value, TableLayout layout,
            out LinkTargetCell target, out long identity, out string detail)
        {
            target = null; identity = 0; detail = null;
            object range = null, tables = null, parent = null, parentRange = null;
            try
            {
                dynamic cell = value;
                int row = (int)cell.RowIndex, column = (int)cell.ColumnIndex;
                if (row <= 0 || column <= 0) throw new InvalidOperationException("所选单元格的行列号不可读。");
                range = cell.Range;
                tables = ((dynamic)range).Tables;
                int tableCount = (int)((dynamic)tables).Count;
                if (tableCount < 1) throw new InvalidOperationException("无法定位所选单元格所属表格。");
                parent = WordComCollection.Item(tables, 1); parentRange = ((dynamic)parent).Range;
                if ((int)((dynamic)parentRange).Start != layout.Start)
                { detail = "所选单元格不属于同一张目标表格。"; return TargetRectStatus.MultiTable; }
                int start = (int)((dynamic)range).Start, end = (int)((dynamic)range).End;
                if (start < layout.Start || end > layout.End || end <= start)
                { detail = "所选单元格的实际范围超出了目标表格。"; return TargetRectStatus.OutOfTable; }
                if (row > layout.Rows.Count)
                { detail = "所选单元格的行号超出了目标表格。"; return TargetRectStatus.OutOfTable; }
                XmlNode tr = layout.Rows[row - 1];
                XmlNodeList rowCells = tr.SelectNodes("w:tc", layout.Namespaces);
                if (column > rowCells.Count)
                { detail = "所选单元格的列号超出了该行。"; return TargetRectStatus.OutOfTable; }
                int gridColumn = checked(1 + Number(tr.SelectSingleNode("w:trPr/w:gridBefore", layout.Namespaces), 0));
                for (int i = 0; i < column - 1; i++)
                    gridColumn = checked(gridColumn + Number(rowCells[i].SelectSingleNode("w:tcPr/w:gridSpan", layout.Namespaces), 1));
                XmlNode tc = rowCells[column - 1];
                if (Number(tc.SelectSingleNode("w:tcPr/w:gridSpan", layout.Namespaces), 1) != 1
                    || tc.SelectSingleNode("w:tcPr/w:vMerge|w:tcPr/w:hMerge|.//w:tbl", layout.Namespaces) != null
                    || tableCount > 1)
                { detail = "所选区域包含合并单元格或嵌套表格，无法逐格对应。"; return TargetRectStatus.NotRect; }
                identity = Pair(start, end);
                target = new LinkTargetCell { Row = row, Column = column, GridColumn = gridColumn, Cell = value };
                return TargetRectStatus.Ok;
            }
            finally { Release(parentRange); Release(parent); Release(tables); Release(range); }
        }

        private static int Number(XmlNode node, int fallback)
        {
            if (node == null) return fallback;
            int value;
            XmlAttribute attribute = node.Attributes["val", node.NamespaceURI];
            if (attribute == null || !int.TryParse(attribute.Value, out value) || value < 0
                || (fallback == 1 && value == 0)) throw new InvalidOperationException("表格网格跨度不可读。");
            return value;
        }

        private static long Pair(int a, int b) { return ((long)a << 32) | (uint)b; }

        internal static void ReleaseTargets(List<LinkTargetCell> targets)
        {
            if (targets == null) return;
            foreach (LinkTargetCell target in targets) { Release(target.Cell); target.Cell = null; }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        /// <summary>目标矩形按行优先（行内先换列，再换行）展成一维序列。</summary>
        public static List<LinkCellPos> FlattenTarget(LinkTargetRect rect)
        {
            List<LinkCellPos> list = new List<LinkCellPos>();
            if (rect == null) return list;
            for (int r = rect.Row1; r <= rect.Row2; r++)
            {
                for (int c = rect.Col1; c <= rect.Col2; c++)
                    list.Add(new LinkCellPos(r, c));
            }
            return list;
        }
    }
}
