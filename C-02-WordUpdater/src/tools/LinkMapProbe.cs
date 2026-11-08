using System;
using System.Collections.Generic;
using System.Threading;
using System.Reflection;
using System.Collections;
using WordUpdater;

// =====================================================================
// 「链接到EXCEL」映射规则离线自检（不需要 WPS）
//
// 验证两件事：
//   1. 源引用 → 一维单元格序列（含 Ctrl 多区域拆分）
//   2. 选中区域 → 目标矩形（含规则矩形 / 越界 / 跨表 判定）与行优先展开
//
// 假对象模型三要点（本仓惯例）：
//   · 假类全部 public —— dynamic 绑定受可访问性限制
//   · C# 索引器默认发出方法名 Item，同类不能既有 this[int] 又有 Item(int)
//   · 断言写「正确行为应该是什么」，而不是复述当前实现
// =====================================================================

public class FakeRange
{
    public int Start;
    public int End;
    public FakeTables Tables = new FakeTables(new List<FakeTableRef>());
    public string WordOpenXML;
    public FakeDocument Document;
    public FakeCellCollection Cells;

    public FakeRange(int s, int e)
    {
        Start = s;
        End = e;
    }
    public bool Information(int kind) { return false; }
    public void Select() { }
}

public class FakeDocument
{
    public FakeRange Content = new FakeRange(0, 1001);
    public FakeRange Range(int start, int end) { return new FakeRange(start, end); }
}

public class FakeDocuments
{
    public int Count { get { return 1; } }
    public object Add() { throw new InvalidOperationException("离线探针没有剪贴板来源。"); }
}

public class FakeMapApplication
{
    public FakeDocuments Documents = new FakeDocuments();
    public FakeSelection Selection;
    public object DisplayAlerts = 1;
    public object ScreenUpdating = true;
}

public class FakeTableRef
{
    public FakeRange Range;

    public FakeTableRef(int start)
    {
        Range = new FakeRange(start, start + 100);
    }
}

public class FakeTables
{
    private List<FakeTableRef> _list;

    public FakeTables(List<FakeTableRef> list) { _list = list; }
    public int Count { get { return _list.Count; } }
    public FakeTableRef this[int i] { get { return _list[i - 1]; } }
}

public class FakeCell
{
    public int RowIndex;
    public int ColumnIndex;
    public FakeRange Range;
}

public class FakeCellCollection
{
    private List<FakeCell> _list;
    public int FailIndex;

    public FakeCellCollection(List<FakeCell> list) { _list = list; }
    public int Count { get { return _list.Count; } }
    public FakeCell this[int i] { get { if (i == FailIndex) throw new InvalidOperationException("单元格读取失败"); return _list[i - 1]; } }
}

public class FakeCount
{
    private int _n;
    public bool ThrowOnRead;
    public FakeCount(int n) { _n = n; }
    public int Count { get { if (ThrowOnRead) throw new InvalidOperationException("整表尺寸不可读"); return _n; } }
}

public class FakeTable : FakeTableRef
{
    public FakeCount Rows;
    public FakeCount Columns;
    public bool Uniform = true;
    public FakeTable() : base(0) { }
}

public class FakeSelection
{
    public FakeRange Range;
    public FakeCellCollection Cells;
    public FakeTables Tables;
    public bool Information(int kind) { return true; }
    public void PasteSpecial(int IconIndex, bool Link, int Placement, bool DisplayAsIcon, int DataType)
    { throw new InvalidOperationException("离线探针没有剪贴板来源。"); }
}

internal static class LinkMapProbe
{
    private static int _pass;
    private static int _fail;

    private static void Check(string name, bool ok, string detail)
    {
        if (ok)
        {
            _pass++;
            Console.WriteLine("[PASS] " + name);
        }
        else
        {
            _fail++;
            Console.WriteLine("[FAIL] " + name + "   " + detail);
        }
    }

    private static void CheckRefs(string name, string fieldCode, string expected)
    {
        List<string> refs;
        string err;
        bool ok = LinkPasteEngine.TryListSourceRefs(fieldCode, out refs, out err);
        if (!ok)
        {
            Check(name, false, "解析失败：" + err);
            return;
        }
        string actual = string.Join(",", refs.ToArray());
        Check(name, actual == expected, "期望 [" + expected + "] 实际 [" + actual + "]");
    }

    private static void CheckRefsFail(string name, string fieldCode)
    {
        List<string> refs;
        string err;
        bool ok = LinkPasteEngine.TryListSourceRefs(fieldCode, out refs, out err);
        Check(name, !ok, "本应拒绝，却解析出了 " + refs.Count + " 条");
    }

    // ---------- 假对象构造 ----------

    private static FakeTable MakeTable(int rows, int cols)
    {
        FakeTable t = new FakeTable();
        t.Rows = new FakeCount(rows);
        t.Columns = new FakeCount(cols);
        t.Range = new FakeRange(0, 1000);
        t.Range.Document = new FakeDocument();
        string xml = "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
        for (int row = 0; row < rows; row++)
        {
            string[] properties = new string[cols];
            xml += RowXml(properties);
        }
        t.Range.WordOpenXML = xml + "</w:tbl>";
        return t;
    }

    private static string RowXml(params string[] properties)
    {
        string xml = "<w:tr>";
        foreach (string property in properties)
            xml += "<w:tc><w:tcPr>" + property + "</w:tcPr><w:p/></w:tc>";
        return xml + "</w:tr>";
    }

    private static void SetRows(FakeTable table, params string[] rows)
    {
        table.Range.WordOpenXML = "<w:tbl xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + string.Join("", rows) + "</w:tbl>";
    }

    private static FakeSelection MakeSel(int[] rowIdx, int[] colIdx, int[] parentStart,
                                         int selStart, int selEnd)
    {
        List<FakeCell> cs = new List<FakeCell>();
        for (int i = 0; i < rowIdx.Length; i++)
        {
            FakeCell c = new FakeCell();
            c.RowIndex = rowIdx[i];
            c.ColumnIndex = colIdx[i];
            int ps = (parentStart != null && i < parentStart.Length) ? parentStart[i] : 0;
            c.Range = new FakeRange(10 + i * 2, 11 + i * 2);
            List<FakeTableRef> tl = new List<FakeTableRef>();
            tl.Add(new FakeTableRef(ps));
            c.Range.Tables = new FakeTables(tl);
            cs.Add(c);
        }
        FakeSelection s = new FakeSelection();
        s.Cells = new FakeCellCollection(cs);
        s.Range = new FakeRange(selStart, selEnd);
        s.Range.Cells = s.Cells;
        return s;
    }

    private static void CheckEngine(string name, FakeSelection selection, FakeTable table, LinkPasteStatus expected)
    {
        selection.Tables = new FakeTables(new List<FakeTableRef> { table });
        LinkPasteResult result = LinkPasteEngine.Execute(new FakeMapApplication { Selection = selection });
        Check(name, result.Status == expected && result.CellsLinked == 0,
            "期望 " + expected + " 实际 " + result.Status + " " + result.ErrorMessage);
    }

    private static void SelectedStructureCases()
    {
        FakeTable header = MakeTable(3, 4);
        header.Uniform = false;
        header.Columns.ThrowOnRead = true;
        SetRows(header, RowXml("<w:gridSpan w:val=\"4\"/>"), RowXml("", "", "", ""), RowXml("", "", "", ""));
        FakeSelection four = MakeSel(new[] { 2, 2, 3, 3 }, new[] { 1, 2, 1, 2 }, null, 10, 20);
        CheckRect("合并表头下方四格仍为2x2", four, header, TargetRectStatus.Ok, "2,1,3,2 (2x2)");
        CheckEngine("整表不规则且Columns不可读仍进入剪贴板检查", four, header, LinkPasteStatus.ClipboardInvalid);
        header.Rows.ThrowOnRead = true;
        CheckEngine("整表Rows/Columns均不可读不阻止局部选区", four, header, LinkPasteStatus.ClipboardInvalid);
        CheckRect("合并表头本身即使只选一格也拒绝",
            MakeSel(new[] { 1 }, new[] { 1 }, null, 10, 12), header, TargetRectStatus.NotRect, null);

        FakeTable vertical = MakeTable(3, 3);
        vertical.Uniform = false;
        SetRows(vertical, RowXml("<w:vMerge w:val=\"restart\"/>", "", ""),
            RowXml("<w:vMerge/>", "", ""), RowXml("", "", ""));
        CheckRect("选区外纵向合并不阻止右侧2x2",
            MakeSel(new[] { 1, 1, 2, 2 }, new[] { 2, 3, 2, 3 }, null, 10, 20), vertical, TargetRectStatus.Ok, "1,2,2,3 (2x2)");
        CheckRect("选区内纵向合并起点拒绝",
            MakeSel(new[] { 1 }, new[] { 1 }, null, 10, 12), vertical, TargetRectStatus.NotRect, null);
        CheckRect("选区内纵向合并延续拒绝",
            MakeSel(new[] { 2 }, new[] { 1 }, null, 10, 12), vertical, TargetRectStatus.NotRect, null);

        FakeTable shifted = MakeTable(2, 4);
        SetRows(shifted, RowXml("<w:gridSpan w:val=\"2\"/>", "", ""), RowXml("", "", "", ""));
        CheckRect("选区左侧横向合并导致物理列号错位仍按网格对齐",
            MakeSel(new[] { 1, 1, 2, 2 }, new[] { 2, 3, 3, 4 }, null, 10, 20), shifted, TargetRectStatus.Ok, "1,3,2,4 (2x2)");
        CheckRect("物理列号相同但实际列边界错位应拒绝",
            MakeSel(new[] { 1, 1, 2, 2 }, new[] { 2, 3, 2, 3 }, null, 10, 20), shifted, TargetRectStatus.NotRect, null);

        FakeTable nested = MakeTable(2, 3);
        string inside = "</w:tcPr><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:tcPr>";
        SetRows(nested, RowXml(inside, "", ""), RowXml("", "", ""));
        CheckRect("选区外嵌套不阻止右侧2x2",
            MakeSel(new[] { 1, 1, 2, 2 }, new[] { 2, 3, 2, 3 }, null, 10, 20), nested, TargetRectStatus.Ok, "1,2,2,3 (2x2)");
        CheckRect("选区内嵌套在写入前拒绝",
            MakeSel(new[] { 1 }, new[] { 1 }, null, 10, 12), nested, TargetRectStatus.NotRect, null);

        FakeTable split = MakeTable(2, 3);
        SetRows(split, RowXml("", "", "", ""), RowXml("", "", "<w:gridSpan w:val=\"2\"/>"));
        CheckRect("选区外拆分不阻止左侧2x2",
            MakeSel(new[] { 1, 1, 2, 2 }, new[] { 1, 2, 1, 2 }, null, 10, 20), split, TargetRectStatus.Ok, "1,1,2,2 (2x2)");
        CheckRect("选区内拆分造成缺格应拒绝",
            MakeSel(new[] { 1, 1, 2 }, new[] { 2, 3, 2 }, null, 10, 20), split, TargetRectStatus.NotRect, null);

        FakeSelection broken = MakeSel(new[] { 1, 1 }, new[] { 1, 2 }, null, 10, 20);
        broken.Cells.FailIndex = 2;
        CheckRect("中途单元格读取失败不能缩成单格后放行", broken, MakeTable(2, 2), TargetRectStatus.ReadError, null);
        FakeSelection alias = MakeSel(new[] { 1, 1 }, new[] { 1, 2 }, null, 10, 20);
        alias.Cells[2].Range.Start = alias.Cells[1].Range.Start;
        alias.Cells[2].Range.End = alias.Cells[1].Range.End;
        CheckRect("不同坐标指向同一物理单元格应拒绝", alias, MakeTable(2, 2), TargetRectStatus.NotRect, null);
        FakeSelection outside = MakeSel(new[] { 1 }, new[] { 1 }, null, 10, 20);
        outside.Cells[1].Range.End = 2000;
        CheckRect("单元格实际范围越界应拒绝", outside, MakeTable(2, 2), TargetRectStatus.OutOfTable, null);

        FakeTable unreadable = MakeTable(2, 2); unreadable.Range.WordOpenXML = "<broken";
        CheckRect("表格结构无法解析时明确失败", four, unreadable, TargetRectStatus.ReadError, null);
        FakeSelection fallback = MakeSel(new[] { 1, 1 }, new[] { 1, 2 }, null, 10, 20); fallback.Cells = null;
        CheckRect("Selection.Cells不可用时Range.Cells兜底", fallback, MakeTable(2, 2), TargetRectStatus.Ok, "1,1,1,2 (1x2)");

        foreach (int count in new[] { 500, 501 })
        {
            var row = new int[count]; var col = new int[count];
            for (int i = 0; i < count; i++) { row[i] = 1; col[i] = i + 1; }
            FakeTable large = MakeTable(1, count); large.Range.End = 10000;
            CheckEngine(count + "格上限检查", MakeSel(row, col, null, 10, 2000), large,
                count == 500 ? LinkPasteStatus.ClipboardInvalid : LinkPasteStatus.TableTooLarge);
        }
    }

    private static void RevalidationCases()
    {
        Type geometry = typeof(LinkPasteGeometry);
        MethodInfo resolve = geometry.GetMethod("ResolveTargetCells", BindingFlags.NonPublic | BindingFlags.Static);
        MethodInfo validate = geometry.GetMethod("RevalidateTargets", BindingFlags.NonPublic | BindingFlags.Static);
        MethodInfo release = geometry.GetMethod("ReleaseTargets", BindingFlags.NonPublic | BindingFlags.Static);
        FakeTable table = MakeTable(2, 2);
        FakeSelection selection = MakeSel(new[] { 1, 1, 2, 2 }, new[] { 1, 2, 1, 2 }, null, 10, 20);
        object[] parameters = { selection, table, null, null, null };
        TargetRectStatus initial = (TargetRectStatus)resolve.Invoke(null, parameters);
        Check("内部入口返回规则矩形及实际单元格", initial == TargetRectStatus.Ok && parameters[3] != null, "实际 " + initial);
        object targets = parameters[3];
        try
        {
            for (int i = 1; i <= 4; i++) { selection.Cells[i].Range.Start += 100; selection.Cells[i].Range.End += 100; }
            CheckRevalidation("文档偏移漂移后仍使用原Cell复核", validate, table, targets, TargetRectStatus.Ok);
            selection.Cells[4].ColumnIndex = 1;
            CheckRevalidation("写入前单元格坐标改变应拒绝", validate, table, targets, TargetRectStatus.NotRect);
            selection.Cells[4].ColumnIndex = 2;
            int start = selection.Cells[4].Range.Start, end = selection.Cells[4].Range.End;
            selection.Cells[4].Range.Start = selection.Cells[1].Range.Start;
            selection.Cells[4].Range.End = selection.Cells[1].Range.End;
            CheckRevalidation("写入前目标变成同一物理格应拒绝", validate, table, targets, TargetRectStatus.NotRect);
            selection.Cells[4].Range.Start = start; selection.Cells[4].Range.End = end;
            string xml = table.Range.WordOpenXML;
            SetRows(table, RowXml("<w:hMerge w:val=\"restart\"/>", ""), RowXml("", ""));
            CheckRevalidation("写入前新增合并结构应拒绝", validate, table, targets, TargetRectStatus.NotRect);
            table.Range.WordOpenXML = "<broken";
            CheckRevalidation("写入前结构无法读取应拒绝", validate, table, targets, TargetRectStatus.ReadError);
            table.Range.WordOpenXML = xml;
            CheckRevalidation("写入前所有目标完整可读才放行", validate, table, targets, TargetRectStatus.Ok);
        }
        finally { release.Invoke(null, new[] { targets }); }
    }

    private static void CheckRevalidation(string name, MethodInfo validate, FakeTable table, object targets, TargetRectStatus expected)
    {
        object[] parameters = { table, targets, null };
        TargetRectStatus actual = (TargetRectStatus)validate.Invoke(null, parameters);
        Check(name, actual == expected, "期望 " + expected + " 实际 " + actual + " " + parameters[2]);
    }

    private static string RectText(LinkTargetRect r)
    {
        if (r == null) return "(null)";
        return r.Row1 + "," + r.Col1 + "," + r.Row2 + "," + r.Col2 + " (" + r.Rows + "x" + r.Cols + ")";
    }

    private static object MultiAreaInfo(string path, string sheet, params int[][] rectangles)
    {
        Assembly assembly = typeof(LinkPasteEngine).Assembly;
        Type infoType = assembly.GetType("WordUpdater.EtSelectionInfo", true);
        Type areaType = assembly.GetType("WordUpdater.LinkFieldCode+SourceArea", true);
        object info = Activator.CreateInstance(infoType, true);
        infoType.GetField("WorkbookPath").SetValue(info, path); infoType.GetField("SheetName").SetValue(info, sheet);
        IList areas = (IList)infoType.GetField("Areas").GetValue(info);
        foreach (int[] rectangle in rectangles)
        {
            object area = Activator.CreateInstance(areaType, true);
            string[] names = { "R1", "C1", "R2", "C2" };
            for (int i = 0; i < names.Length; i++) areaType.GetField(names[i]).SetValue(area, rectangle[i]);
            areas.Add(area);
        }
        return info;
    }

    private static void CheckMultiAreaCandidate(string name, bool expected, params object[] candidates)
    {
        Type engine = typeof(LinkPasteEngine);
        MethodInfo parse = typeof(LinkPasteEngine).Assembly.GetType("WordUpdater.LinkFieldCode", true).GetMethod("ParseLinkCode", BindingFlags.Static | BindingFlags.NonPublic);
        object[] parseArguments = { "LINK Excel.Sheet.12 \"C:\\sample\\source.xlsx\" \"Sheet1!R2C1:R4C4\" \\t \\a", false, null };
        object template = parse.Invoke(null, parseArguments);
        Type info = engine.Assembly.GetType("WordUpdater.EtSelectionInfo", true);
        IList list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(info));
        foreach (object candidate in candidates) list.Add(candidate);
        object[] arguments = { template, list, 8, null };
        bool applied = (bool)engine.GetMethod("TryApplyEtSelections", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, arguments);
        IList cells = (IList)typeof(LinkPasteEngine).Assembly.GetType("WordUpdater.LinkFieldCode", true).GetMethod("ExpandSource", BindingFlags.NonPublic | BindingFlags.Static).Invoke(null, new[] { template });
        Check(name, applied == expected && cells.Count == (expected ? 8 : 12), "applied=" + applied + " cells=" + cells.Count + " " + arguments[3]);
        if (applied)
        {
            string[] actual = new string[cells.Count]; for (int i = 0; i < cells.Count; i++) actual[i] = cells[i].ToString();
            Check("第2/4行映射不包含第3行", string.Join(",", actual) == "Sheet1!R2C1,Sheet1!R2C2,Sheet1!R2C3,Sheet1!R2C4,Sheet1!R4C1,Sheet1!R4C2,Sheet1!R4C3,Sheet1!R4C4", string.Join(",", actual));
        }
    }

    private static void MultiAreaCandidateCases()
    {
        string path = "C:\\sample\\source.xlsx";
        object correct = MultiAreaInfo(path, "Sheet1", new[] { 2, 1, 2, 4 }, new[] { 4, 1, 4, 4 });
        CheckMultiAreaCandidate("第2/4行12格包围盒恢复为8格", true, correct);
        CheckMultiAreaCandidate("跳过格数不匹配窗口候选后继续寻找", true,
            MultiAreaInfo(path, "Sheet1", new[] { 2, 1, 2, 2 }, new[] { 4, 1, 4, 2 }), correct);
        CheckMultiAreaCandidate("不同工作簿窗口不可冒充来源", false,
            MultiAreaInfo("C:\\other.xlsx", "Sheet1", new[] { 2, 1, 2, 4 }, new[] { 4, 1, 4, 4 }));
        CheckMultiAreaCandidate("不同工作表窗口不可冒充来源", false,
            MultiAreaInfo(path, "Sheet2", new[] { 2, 1, 2, 4 }, new[] { 4, 1, 4, 4 }));
        CheckMultiAreaCandidate("仅位于包围盒内部但边界不同的陈旧选区拒绝", false,
            MultiAreaInfo(path, "Sheet1", new[] { 2, 1, 2, 4 }, new[] { 3, 1, 3, 4 }));
        CheckMultiAreaCandidate("8格含重叠的候选不能算8个独立格", false,
            MultiAreaInfo(path, "Sheet1", new[] { 2, 1, 2, 4 }, new[] { 2, 2, 4, 2 }, new[] { 4, 4, 4, 4 }));
        CheckMultiAreaCandidate("缺少来源路径时不猜工作簿", false,
            MultiAreaInfo(null, "Sheet1", new[] { 2, 1, 2, 4 }, new[] { 4, 1, 4, 4 }));
        CheckMultiAreaCandidate("没有可验证候选时保留剪贴板范围", false);
        CheckMultiAreaCandidate("相同窗口选区重复发现可以去重", true, correct, correct);
        CheckMultiAreaCandidate("多个窗口具有不同但均匹配的8格选区时不猜来源", false, correct,
            MultiAreaInfo(path, "Sheet1", new[] { 2, 1, 4, 1 }, new[] { 2, 4, 4, 4 }, new[] { 2, 2, 2, 3 }));
    }

    private static void CheckRect(string name, FakeSelection sel, FakeTable tbl,
                                  TargetRectStatus expectStatus, string expectRect)
    {
        LinkTargetRect rect;
        string detail;
        TargetRectStatus st = LinkPasteGeometry.ResolveTargetRect(sel, tbl, out rect, out detail);
        if (st != expectStatus)
        {
            Check(name, false, "期望状态 " + expectStatus + " 实际 " + st + "：" + detail);
            return;
        }
        if (expectRect == null)
        {
            Check(name, true, "");
            return;
        }
        string actual = RectText(rect);
        Check(name, actual == expectRect, "期望 [" + expectRect + "] 实际 [" + actual + "]");
    }

    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine("== LinkMapProbe：链接映射规则离线自检 ==");
        Console.WriteLine();

        // ============ 1. 源引用 → 单元格序列 ============
        Console.WriteLine("--- 源引用展开（行优先）---");

        CheckRefs("连续区域 R1C1:R2C2 展成 4 格（先行后列）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1:R2C2\" \\t \\a",
            "Sheet1!R1C1,Sheet1!R1C2,Sheet1!R2C1,Sheet1!R2C2");

        CheckRefs("连续区域 A1:B2 展成 4 格（先行后列）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:B2\" \\a",
            "Sheet1!A1,Sheet1!B1,Sheet1!A2,Sheet1!B2");

        CheckRefs("单行区域 A1:C1 展成 3 格",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:C1\" \\a",
            "Sheet1!A1,Sheet1!B1,Sheet1!C1");

        CheckRefs("Ctrl 多区域 R1C1,R3C1（第二片省略工作表名，应继承）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1,R3C1\" \\a",
            "Sheet1!R1C1,Sheet1!R3C1");

        CheckRefs("Ctrl 多区域 每片都带工作表名",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R1C1,Sheet1!R3C1\" \\a",
            "Sheet1!R1C1,Sheet1!R3C1");

        CheckRefs("Ctrl 多区域 A1,B2 视为两个单格（不是 A1:B2 区域）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1,B2\" \\a",
            "Sheet1!A1,Sheet1!B2");

        CheckRefs("Ctrl 多区域 H1,H3 映射为两个单格",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!H1,H3\" \\a",
            "Sheet1!H1,Sheet1!H3");

        CheckRefs("分号分隔同样视为多区域",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1;C1\" \\a",
            "Sheet1!A1,Sheet1!C1");

        CheckRefs("单个单元格 R11C12（L11 场景）",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R11C12\" \\a",
            "Sheet1!R11C12");

        CheckRefs("DDE 形态无工作表前缀",
            "DDE Excel \"[C:\\a.xlsx]Sheet1\" \"R1C1:R1C2\"",
            "R1C1,R1C2");

        CheckRefsFail("无法识别的引用应被拒绝",
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"ZZZ\" \\a");

        CheckRefsFail("空代码应被拒绝", "");

        // ============ 2. 目标矩形解析 ============
        Console.WriteLine();
        Console.WriteLine("--- 目标矩形解析 ---");

        FakeTable t3 = MakeTable(3, 3);

        // 只选 (3,3) 一格
        CheckRect("3x3 表只选 (3,3) 一格 → 1x1",
            MakeSel(new int[] { 3 }, new int[] { 3 }, null, 10, 12), t3,
            TargetRectStatus.Ok, "3,3,3,3 (1x1)");

        // 选 (1,1)-(2,2)
        CheckRect("选中 (1,1)-(2,2) → 2x2",
            MakeSel(new int[] { 1, 1, 2, 2 }, new int[] { 1, 2, 1, 2 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,2,2 (2x2)");

        // 选中第 1 列第 1~2 行（用户场景：H1/H3 → 两行一列）
        CheckRect("选中第 1 列的第 1~2 行 → 2x1",
            MakeSel(new int[] { 1, 2 }, new int[] { 1, 1 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,2,1 (2x1)");

        // 选中第 1 行第 1~2 列（用户场景：A1/C1 → 一行两列）
        CheckRect("选中第 1 行的第 1~2 列 → 1x2",
            MakeSel(new int[] { 1, 1 }, new int[] { 1, 2 }, null, 10, 20), t3,
            TargetRectStatus.Ok, "1,1,1,2 (1x2)");

        // 全表
        int[] allR = new int[9];
        int[] allC = new int[9];
        int k = 0;
        for (int r = 1; r <= 3; r++)
            for (int c = 1; c <= 3; c++) { allR[k] = r; allC[k] = c; k++; }
        CheckRect("选中整张 3x3 表 → 3x3",
            MakeSel(allR, allC, null, 0, 1000), t3,
            TargetRectStatus.Ok, "1,1,3,3 (3x3)");

        // L 形（不是矩形）
        CheckRect("L 形选区应判定不是规则矩形",
            MakeSel(new int[] { 1, 1, 2, 3 }, new int[] { 1, 2, 1, 1 }, null, 10, 20), t3,
            TargetRectStatus.NotRect, null);

        // 越出表格
        CheckRect("选区跨出表格边界应判定越界",
            MakeSel(new int[] { 3 }, new int[] { 3 }, null, 10, 5000), t3,
            TargetRectStatus.OutOfTable, null);

        // 行号超出表格行数
        CheckRect("单元格行号超出表格行数应判定越界",
            MakeSel(new int[] { 9 }, new int[] { 1 }, null, 10, 20), t3,
            TargetRectStatus.OutOfTable, null);

        // 跨两张表
        CheckRect("选区横跨两张表格应判定跨表",
            MakeSel(new int[] { 1, 1 }, new int[] { 1, 2 }, new int[] { 0, 500 }, 10, 20), t3,
            TargetRectStatus.MultiTable, null);

        // 没有单元格
        CheckRect("选区没有单元格应判定 NoCells",
            MakeSel(new int[0], new int[0], null, 10, 20), t3,
            TargetRectStatus.NoCells, null);

        SelectedStructureCases();
        RevalidationCases();
        MultiAreaCandidateCases();

        // ============ 3. 目标行优先展开顺序 ============
        Console.WriteLine();
        Console.WriteLine("--- 目标行优先展开顺序 ---");

        LinkTargetRect rect22 = new LinkTargetRect();
        rect22.Row1 = 1; rect22.Col1 = 1; rect22.Row2 = 2; rect22.Col2 = 2;
        List<LinkCellPos> flat = LinkPasteGeometry.FlattenTarget(rect22);
        string order = "";
        for (int i = 0; i < flat.Count; i++)
        {
            if (i > 0) order += ",";
            order += "(" + flat[i].Row + "," + flat[i].Col + ")";
        }
        Check("2x2 展开顺序为 (1,1),(1,2),(2,1),(2,2)",
            order == "(1,1),(1,2),(2,1),(2,2)", "实际 " + order);

        // 用户场景：源 H1,H3 → 目标第 1 列第 1~2 行
        LinkTargetRect col2 = new LinkTargetRect();
        col2.Row1 = 1; col2.Col1 = 1; col2.Row2 = 2; col2.Col2 = 1;
        List<LinkCellPos> f2 = LinkPasteGeometry.FlattenTarget(col2);
        Check("H1,H3 → 第1列1~2行：目标序列长度=2 且为 (1,1),(2,1)",
            f2.Count == 2 && f2[0].Row == 1 && f2[0].Col == 1 && f2[1].Row == 2 && f2[1].Col == 1,
            "实际 " + f2.Count + " 格");

        // 用户场景：源 A1,C1 → 目标第 1 行第 1~2 列
        LinkTargetRect row2 = new LinkTargetRect();
        row2.Row1 = 1; row2.Col1 = 1; row2.Row2 = 1; row2.Col2 = 2;
        List<LinkCellPos> f3 = LinkPasteGeometry.FlattenTarget(row2);
        Check("A1,C1 → 第1行1~2列：目标序列长度=2 且为 (1,1),(1,2)",
            f3.Count == 2 && f3[0].Row == 1 && f3[0].Col == 1 && f3[1].Row == 1 && f3[1].Col == 2,
            "实际 " + f3.Count + " 格");

        // ============ 4. 数量匹配（用户定义的严格规则）============
        Console.WriteLine();
        Console.WriteLine("--- 数量严格匹配 ---");

        List<string> refsA;
        string errA;
        LinkPasteEngine.TryListSourceRefs(
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!A1:B2\" \\a", out refsA, out errA);
        LinkTargetRect one = new LinkTargetRect();
        one.Row1 = 3; one.Col1 = 3; one.Row2 = 3; one.Col2 = 3;
        Check("源 2x2(4 格) 对目标 (3,3) 单格(1 格) → 数量不等，应拒绝",
            refsA.Count == 4 && Flatten(one) == 1 && refsA.Count != Flatten(one),
            "源=" + refsA.Count + " 目标=" + Flatten(one));

        List<string> refsB;
        string errB;
        LinkPasteEngine.TryListSourceRefs(
            "LINK Excel.Sheet.12 \"C:\\a.xlsx\" \"Sheet1!R11C12\" \\a", out refsB, out errB);
        Check("源单格 L11(1 格) 对目标 (3,3) 单格(1 格) → 数量相等，应放行",
            refsB.Count == 1 && Flatten(one) == 1 && refsB.Count == Flatten(one),
            "源=" + refsB.Count + " 目标=" + Flatten(one));

        Console.WriteLine();
        Console.WriteLine("结果：PASS=" + _pass + "  FAIL=" + _fail);
        Console.Out.Flush();
        return _fail == 0 ? 0 : 1;
    }

    private static int Flatten(LinkTargetRect r)
    {
        return LinkPasteGeometry.FlattenTarget(r).Count;
    }
}
