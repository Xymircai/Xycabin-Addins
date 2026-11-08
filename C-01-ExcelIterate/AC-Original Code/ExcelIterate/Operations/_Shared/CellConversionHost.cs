using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal static class ConversionRangeAddress
    {
        internal const int MaximumCells = 100000;
        internal static int[] Coordinates(string token) { return CellAddress.Coordinates(token); }
        private static IList<int[]> Areas(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) throw new FormatException("请选择或输入单元格区域。");
            var areas = new List<int[]>(); long count = 0;
            foreach (string area in address.Split(','))
            {
                string[] ends = area.Trim().Split(':');
                if (ends.Length > 2) throw new FormatException("区域地址无效。");
                int[] a = Coordinates(ends[0]), b = Coordinates(ends[ends.Length - 1]);
                int top = Math.Min(a[0], b[0]), bottom = Math.Max(a[0], b[0]), left = Math.Min(a[1], b[1]), right = Math.Max(a[1], b[1]);
                long size = (long)(bottom - top + 1) * (right - left + 1);
                if (size > MaximumCells) throw new FormatException("单批最多处理 100000 个单元格，请缩小区域。");
                count += size; areas.Add(new[] { top, left, bottom, right });
            }
            if (count > MaximumCells)
            {
                var seen = new HashSet<long>();
                foreach (var area in areas)
                    for (int row = area[0]; row <= area[2]; row++) for (int col = area[1]; col <= area[3]; col++)
                        if (seen.Add(((long)row << 15) | (uint)col) && seen.Count > MaximumCells)
                            throw new FormatException("单批最多处理 100000 个单元格，请缩小区域。");
            }
            return areas;
        }
        internal static IList<string> Cells(string address)
        {
            var names = new List<string>(); var seen = new HashSet<long>();
            foreach (var area in Areas(address))
                for (int row = area[0]; row <= area[2]; row++) for (int col = area[1]; col <= area[3]; col++)
                    if (seen.Add(((long)row << 15) | (uint)col)) names.Add(CellAddress.Name(row, col));
            return names;
        }
        internal static string Normalize(string address)
        {
            var result = new List<string>();
            foreach (var area in Areas(address))
            {
                string first = CellAddress.Name(area[0], area[1]), last = CellAddress.Name(area[2], area[3]);
                result.Add(first == last ? first : first + ":" + last);
            }
            return string.Join(",", result);
        }
        internal static string FirstCell(string address)
        { var area = Areas(address)[0]; return CellAddress.Name(area[0], area[1]); }
        // Compress adjacent cells before constructing a calculation union in the host.
        internal static IList<string> Rectangles(IList<string> addresses)
        {
            var points = new List<int[]>(); foreach (string address in addresses) points.Add(Coordinates(address));
            points.Sort(delegate(int[] a, int[] b) { int row = a[0].CompareTo(b[0]); return row != 0 ? row : a[1].CompareTo(b[1]); });
            var blocks = new List<int[]>(); var last = new Dictionary<string, int[]>();
            for (int i = 0; i < points.Count;)
            {
                int row = points[i][0], left = points[i][1], right = left; i++;
                while (i < points.Count && points[i][0] == row && points[i][1] <= right + 1) { right = Math.Max(right, points[i][1]); i++; }
                string key = left + ":" + right; int[] block;
                if (last.TryGetValue(key, out block) && block[2] == row - 1) block[2] = row;
                else { block = new[] { row, left, row, right }; blocks.Add(block); last[key] = block; }
            }
            var result = new List<string>();
            foreach (var block in blocks) { string first = CellAddress.Name(block[0], block[1]), end = CellAddress.Name(block[2], block[3]); result.Add(first == end ? first : first + ":" + end); }
            return result;
        }
    }

    // Immutable strings and flags only; no COM object or host closure crosses to the editor STA.
    internal sealed class ConversionPreview
    {
        internal readonly string Address, CellAddress, Display, Formula;
        internal readonly bool IsFormula, IsNumeric, IsBlank;
        internal ConversionPreview(string address, ConversionCellState state)
        {
            Address = ConversionRangeAddress.Normalize(address); CellAddress = Address.Split(',')[0].Split(':')[0];
            Display = state.Display ?? ""; Formula = state.Formula; IsFormula = state.IsFormula; IsNumeric = CellConversionEngine.IsNumeric(state.Value);
            IsBlank = !state.IsFormula && state.Value == null;
        }
    }

    internal interface IConversionSessionHost : IConversionHost
    {
        string Address { get; }
        string Description { get; }
        void SetAddress(string address);
        string PickRange();
        ConversionPreview ReadPreview(string address);
    }
    // All methods and owned COM references stay on the original host STA.
    internal sealed partial class ComConversionHost : IConversionSessionHost
    {
        private readonly object application;
        private object worksheet;
        private WorksheetIdentity identity;
        private readonly List<object> ownedCells = new List<object>();
        private IList<IConversionCell> cells;
        private bool disposed;
        public string Address { get; private set; }
        public string Description { get; private set; }
        internal ComConversionHost(object app)
        {
            application = app;
            try
            {
                worksheet = ((dynamic)app).ActiveSheet; identity = WorksheetIdentity.Capture(worksheet);
                Description = "工作簿：" + identity.Path + "\r\n工作表：" + identity.Name;
                RecheckScope();
                object range = ((dynamic)app).Selection;
                try { Address = RangeAddress(range); } finally { Release(range); }
            }
            catch { Dispose(); throw; }
        }
        internal static bool Boolean(object value, string property)
        { if (!(value is bool)) throw new InvalidOperationException("无法可靠读取 " + property + "。"); return (bool)value; }
        private static bool Missing(Exception error)
        {
            var com = error as COMException;
            return error is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException || error is MissingMemberException || com != null && (com.ErrorCode == unchecked((int)0x80020003) || com.ErrorCode == unchecked((int)0x80020006));
        }
        internal static void Release(object value)
        { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        internal string RangeAddress(object range)
        {
            object sheet = null, areas = null;
            try
            {
                if (range == null || range is bool) throw new InvalidOperationException("请选择工作表单元格区域。");
                sheet = ((dynamic)range).Worksheet;
                if (!identity.Matches(sheet)) throw new InvalidOperationException("请选择窗口所示工作簿和工作表中的区域。");
                areas = ((dynamic)range).Areas; int count = System.Convert.ToInt32(((dynamic)areas).Count); var addresses = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    object area = null, rows = null, columns = null;
                    try
                    {
                        area = ((dynamic)areas).Item(i); rows = ((dynamic)area).Rows; columns = ((dynamic)area).Columns;
                        int row = System.Convert.ToInt32(((dynamic)area).Row), col = System.Convert.ToInt32(((dynamic)area).Column);
                        int height = System.Convert.ToInt32(((dynamic)rows).Count), width = System.Convert.ToInt32(((dynamic)columns).Count);
                        if ((long)height * width > ConversionRangeAddress.MaximumCells) throw new InvalidOperationException("单批最多处理 100000 个单元格，请缩小区域。");
                        addresses.Add(CellAddress.Name(row, col) + ":" + CellAddress.Name(row + height - 1, col + width - 1));
                    }
                    finally { Release(columns); Release(rows); Release(area); }
                }
                return ConversionRangeAddress.Normalize(string.Join(",", addresses));
            }
            finally { Release(areas); Release(sheet); }
        }
        public void SetAddress(string address)
        {
            string normalized = ConversionRangeAddress.Normalize(address);
            if (normalized == Address) return;
            ClearCells(); Address = normalized;
        }
        public string PickRange()
        {
            RecheckScope(); object range = null;
            try
            {
                range = ((dynamic)application).InputBox("请框选本工作表中要转换的区域。取消会保留原地址，框选完成后仍需点击执行。", "转换 · 选择区域", Address, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (range == null || range is bool) return null;
                return RangeAddress(range);
            }
            finally { Release(range); }
        }
        public void RecheckScope()
        {
            if (disposed) throw new ObjectDisposedException("ComConversionHost");
            object active = null, window = null, selected = null;
            try
            {
                active = ((dynamic)application).ActiveSheet;
                if (!identity.Matches(active)) throw new InvalidOperationException("目标工作表已关闭、替换、改名或已切换，请返回原表或重新打开转换窗口。");
                window = ((dynamic)application).ActiveWindow; selected = ((dynamic)window).SelectedSheets;
                if (System.Convert.ToInt32(((dynamic)selected).Count) != 1) throw new InvalidOperationException("请先取消工作表组合，再执行转换。");
            }
            finally { Release(selected); Release(window); Release(active); }
        }
        public bool EnableEvents { get { return Boolean((object)((dynamic)application).EnableEvents, "EnableEvents"); } set { ((dynamic)application).EnableEvents = value; } }
        public bool ScreenUpdating { get { return Boolean((object)((dynamic)application).ScreenUpdating, "ScreenUpdating"); } set { ((dynamic)application).ScreenUpdating = value; } }
        public ConversionPreview ReadPreview(string address)
        {
            RecheckScope(); string normalized = ConversionRangeAddress.Normalize(address); object cell = null;
            try
            {
                cell = ((dynamic)worksheet).Range[ConversionRangeAddress.FirstCell(normalized)];
                return new ConversionPreview(normalized, new ComCell(cell, worksheet, "预览首格").Original);
            }
            finally { Release(cell); }
        }
        public void Calculate(IList<IConversionCell> targets)
        {
            RecheckScope(); var addresses = new List<string>();
            foreach (var item in targets)
            {
                var target = item as ComCell; var batch = item as BatchComCell;
                if (target == null && batch == null) throw new InvalidOperationException("计算目标不属于本次转换。");
                addresses.Add(target == null ? batch.LocalAddress : target.LocalAddress);
            }
            object combined = null;
            try
            {
                foreach (string address in ConversionRangeAddress.Rectangles(addresses))
                {
                    object part = ((dynamic)worksheet).Range[address];
                    try
                    {
                        if (combined == null) { combined = part; part = null; }
                        else { object joined = ((dynamic)application).Union(combined, part); Release(combined); combined = joined; }
                    }
                    finally { Release(part); }
                }
                if (combined != null) ((dynamic)combined).Calculate();
            }
            finally { Release(combined); }
        }
        public IList<IConversionCell> Collect()
        {
            RecheckScope(); if (cells != null) return cells;
            var result = new List<IConversionCell>();
            foreach (string address in ConversionRangeAddress.Cells(Address))
            {
                object cell = ((dynamic)worksheet).Range[address]; bool keep = false;
                try
                {
                    if (Boolean((object)((dynamic)cell).MergeCells, "MergeCells"))
                    {
                        object merge = ((dynamic)cell).MergeArea;
                        try
                        {
                            if (System.Convert.ToInt32(((dynamic)cell).Row) != System.Convert.ToInt32(((dynamic)merge).Row) || System.Convert.ToInt32(((dynamic)cell).Column) != System.Convert.ToInt32(((dynamic)merge).Column)) continue;
                        }
                        finally { Release(merge); }
                    }
                    var item = new ComCell(cell, worksheet, Description + "\r\n单元格：" + address, address);
                    result.Add(item); ownedCells.Add(cell); keep = true;
                }
                finally { if (!keep) Release(cell); }
            }
            cells = result; return cells;
        }
        private void ClearCells()
        {
            cells = null;
            var errors = new List<string>();
            for (int i = ownedCells.Count - 1; i >= 0; i--) try { Release(ownedCells[i]); } catch (Exception ex) { errors.Add(ex.Message); }
            ownedCells.Clear(); if (errors.Count > 0) throw new InvalidOperationException(string.Join("；", errors));
        }
        public void Dispose()
        {
            if (disposed) return; disposed = true;
            try { ClearCells(); }
            finally { try { if (identity != null) identity.Dispose(); } finally { identity = null; Release(worksheet); worksheet = null; } }
        }
        private sealed class ComCell : IConversionCell
        {
            private readonly object cell, worksheet;
            private bool formula2 = true;
            public string Address { get; private set; }
            internal readonly string LocalAddress;
            public ConversionCellState Original { get; private set; }
            internal ComCell(object range, object sheet, string address, string localAddress = null)
            { cell = range; worksheet = sheet; Address = address; LocalAddress = localAddress; Original = Read(); }
            public ConversionCellState Read()
            {
                var state = new ConversionCellState(); dynamic range = cell;
                state.IsFormula = ComConversionHost.Boolean((object)range.HasFormula, "HasFormula");
                if (state.IsFormula)
                {
                    object formula;
                    try { formula = formula2 ? range.Formula2 : range.Formula; }
                    catch (Exception ex) { if (!formula2 || !Missing(ex)) throw; formula2 = false; formula = range.Formula; }
                    state.Formula = formula as string; if (state.Formula == null) throw new InvalidOperationException(Address + " 无法读取公式。");
                }
                state.Value = range.Value2; state.NumberFormat = range.NumberFormat as string; state.Display = range.Text as string;
                if (state.NumberFormat == null || state.Display == null) throw new InvalidOperationException(Address + " 无法读取数字格式或显示文字。");
                return state;
            }
            public void CheckWritable()
            {
                try
                {
                    if (ComConversionHost.Boolean((object)((dynamic)cell).HasArray, "HasArray")) throw new InvalidOperationException("数组公式不能转换。");
                    if (ComConversionHost.Boolean((object)((dynamic)cell).HasSpill, "HasSpill")) throw new InvalidOperationException("动态溢出区域不能转换。");
                    if (ComConversionHost.Boolean((object)((dynamic)worksheet).ProtectContents, "ProtectContents") && ComConversionHost.Boolean((object)((dynamic)cell).Locked, "Locked")) throw new InvalidOperationException("单元格受保护且已锁定。");
                }
                catch (Exception ex) { throw new InvalidOperationException(Address + "：" + ex.Message, ex); }
            }
            public void Write(ConversionCellState state)
            {
                dynamic range = cell;
                // Assign the expression/quote prefix under General, then restore the desired
                // format. A pre-existing Text format must not turn a formula into literal text.
                range.NumberFormat = state.IsFormula || state.Value is string ? "General" : state.NumberFormat;
                if (state.IsFormula)
                {
                    try
                    {
                        try { if (formula2) range.Formula2 = state.Formula; else range.Formula = state.Formula; }
                        catch (Exception ex) { if (!formula2 || !Missing(ex)) throw; formula2 = false; range.Formula = state.Formula; }
                    }
                    catch (COMException error)
                    {
                        // Excel and WPS surface formula input rejection at this setter;
                        // transport/member errors and format writes must not enter this path.
                        if (BatchTextEditEngine.IsSyntaxRejection(error)) throw new FormulaInputRejectedException(error);
                        throw;
                    }
                }
                else if (state.Value == null) range.ClearContents();
                else if (state.Value is string) range.Formula = "'" + (string)state.Value;
                else range.Value2 = state.Value;
                // Some hosts normalize a format while assigning Formula/Value2.
                range.NumberFormat = state.NumberFormat;
            }
        }
    }
}
