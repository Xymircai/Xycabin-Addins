using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    internal sealed class ColumnChoice
    {
        internal int Number;
        internal string Letter, Header;
        internal bool OriginalHidden, Visible;
    }
    internal interface IColumnVisibilityHost : IDisposable
    {
        IList<ColumnChoice> Collect();
        void Recheck();
        bool ReadHidden(int column);
        void WriteHidden(int column, bool hidden);
        bool EnableEvents { get; set; }
        bool ScreenUpdating { get; set; }
    }
    internal static class ColumnVisibilityEngine
    {
        internal static FormulaCommandResult Apply(IColumnVisibilityHost host, IList<ColumnChoice> choices, string recoveryDirectory)
        {
            FormulaCommandResult result = new FormulaCommandResult();
            List<ColumnChoice> attempted = new List<ColumnChoice>();
            bool captured = false, events = false, screen = false;
            try
            {
                IList<ColumnChoice> original = host.Collect();
                if (choices.Count != original.Count) throw new InvalidOperationException("列清单已变化，请重新打开窗口。");
                bool anyVisible = false;
                for (int i = 0; i < choices.Count; i++)
                {
                    if (choices[i].Number != original[i].Number || choices[i].OriginalHidden != original[i].OriginalHidden)
                        throw new InvalidOperationException("列身份或快照不一致，请重新打开窗口。");
                    anyVisible |= choices[i].Visible;
                }
                if (!anyVisible) throw new InvalidOperationException("请至少保留一列勾选。");
                host.Recheck(); // Includes all headers, original visibility, bounds and protection.
                events = host.EnableEvents; screen = host.ScreenUpdating; captured = true;
                host.EnableEvents = false; host.ScreenUpdating = false;
                if (host.EnableEvents || host.ScreenUpdating) throw new InvalidOperationException("无法关闭事件或屏幕刷新，未修改列。");
                host.Recheck();
                foreach (ColumnChoice choice in choices)
                {
                    bool target = !choice.Visible;
                    if (target == choice.OriginalHidden) continue;
                    if (host.ReadHidden(choice.Number) != choice.OriginalHidden) throw new InvalidOperationException(choice.Letter + " 列状态已变化。");
                    attempted.Add(choice);
                    host.WriteHidden(choice.Number, target);
                    if (host.ReadHidden(choice.Number) != target) throw new InvalidOperationException(choice.Letter + " 列写入核对失败。");
                }
                result.Changed = attempted.Count;
                result.Details = result.Changed == 0 ? "列显示状态无需修改。" : "已修改 " + result.Changed + " 列的显示状态。";
            }
            catch (Exception ex)
            {
                result.Failed = true; result.Changed = 0;
                result.Details = "按列筛选未完成：" + ex.Message;
                List<string> failures = new List<string>();
                for (int i = attempted.Count - 1; i >= 0; i--)
                {
                    ColumnChoice column = attempted[i]; string writeError = null;
                    try { host.WriteHidden(column.Number, column.OriginalHidden); } catch (Exception error) { writeError = error.Message; }
                    try { if (host.ReadHidden(column.Number) != column.OriginalHidden) throw new InvalidOperationException("原状态核对失败"); }
                    catch (Exception error) { failures.Add(column.Letter + " 列；原隐藏状态：" + column.OriginalHidden + "；" + writeError + "；" + error.Message); }
                }
                if (failures.Count == 0) result.Details += attempted.Count == 0 ? "\r\n本批未修改任何列。" : "\r\n本批已尝试修改的列均已还原并核对。";
                else
                {
                    result.RollbackIncomplete = true;
                    result.Details += "\r\n\r\n以下列未能确认还原：\r\n" + string.Join("\r\n", failures);
                    try
                    {
                        Directory.CreateDirectory(recoveryDirectory);
                        result.RecoveryLog = Path.Combine(recoveryDirectory, "ColumnRecovery-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N") + ".txt");
                        File.WriteAllText(result.RecoveryLog, result.Details, System.Text.Encoding.UTF8);
                        result.Details += "\r\n恢复日志：" + result.RecoveryLog;
                    }
                    catch (Exception error) { result.Details += "\r\n日志保存失败：" + error.Message + "。请复制本窗口详情。"; }
                }
            }
            finally
            {
                if (captured)
                {
                    try { host.ScreenUpdating = screen; if (host.ScreenUpdating != screen) throw new Exception("恢复核对失败"); }
                    catch (Exception ex) { result.Failed = true; result.Details += "\r\n屏幕刷新恢复失败：" + ex.Message; }
                    try { host.EnableEvents = events; if (host.EnableEvents != events) throw new Exception("恢复核对失败"); }
                    catch (Exception ex) { result.Failed = true; result.Details += "\r\n事件恢复失败：" + ex.Message; }
                }
            }
            return result;
        }
    }

    // Bound to the opening selection; all calls run on the host's UI thread.
    internal sealed class ComColumnVisibilityHost : IColumnVisibilityHost
    {
        private readonly object app;
        private object sheet;
        private WorksheetIdentity identity;
        private int row, firstColumn, count, anchorRow, anchorColumn;
        private string tableSignature;
        internal string RangeDescription { get; private set; }
        private IList<ColumnChoice> collected;
        private readonly List<string> signatures = new List<string>();
        private bool disposed;
        internal ComColumnVisibilityHost(object application) { app = application; }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        public bool EnableEvents { get { return ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents"); } set { ((dynamic)app).EnableEvents = value; } }
        public bool ScreenUpdating { get { return ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating"); } set { ((dynamic)app).ScreenUpdating = value; } }
        private void CheckSheetGroup()
        {
            object window = ((dynamic)app).ActiveWindow, sheets = null;
            try { if (window == null) throw new InvalidOperationException("请先打开工作簿。"); sheets = ((dynamic)window).SelectedSheets; if (Convert.ToInt32(((dynamic)sheets).Count) != 1) throw new InvalidOperationException("请先取消工作表组合选择。"); }
            finally { Release(sheets); Release(window); }
        }
        private void CheckProtection()
        {
            if (!ComFormulaHost.Boolean(((dynamic)sheet).ProtectContents, "ProtectContents")) return;
            object protection = ((dynamic)sheet).Protection;
            try { if (!ComFormulaHost.Boolean(((dynamic)protection).AllowFormattingColumns, "AllowFormattingColumns")) throw new InvalidOperationException("工作表保护不允许设置列格式，请先取消保护。"); }
            finally { Release(protection); }
        }
        private static bool Missing(Exception ex)
        {
            COMException com = ex as COMException;
            return ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException || (com != null && (com.ErrorCode == unchecked((int)0x80020003) || com.ErrorCode == unchecked((int)0x80020006)));
        }
        private static int Size(object range, bool columns)
        {
            object items = columns ? ((dynamic)range).Columns : ((dynamic)range).Rows;
            try { return Convert.ToInt32(((dynamic)items).Count); } finally { Release(items); }
        }
        private bool Contains(object range)
        {
            int r = Convert.ToInt32(((dynamic)range).Row), c = Convert.ToInt32(((dynamic)range).Column);
            return anchorRow >= r && anchorRow < r + Size(range, false) && anchorColumn >= c && anchorColumn < c + Size(range, true);
        }
        private string Resolve(out int headerRow, out int start, out int width)
        {
            using (ComBag bag = new ComBag())
            {
                object range = null, headers = null, tables = null; string kind = "连续区域", name = "";
                try { tables = bag.Own(((dynamic)sheet).ListObjects); } catch (Exception ex) { if (!Missing(ex)) throw; }
                if (tables != null)
                    for (int i = 1; i <= Convert.ToInt32(((dynamic)tables).Count); i++)
                    {
                        object table = bag.Own(((dynamic)tables).Item(i)), candidate = bag.Own(((dynamic)table).Range);
                        if (!Contains(candidate)) continue;
                        range = candidate; headers = bag.Own(((dynamic)table).HeaderRowRange);
                        if (headers == null) throw new InvalidOperationException("该正式表格没有可读取的表头行。");
                        kind = "正式表格"; name = Convert.ToString(((dynamic)table).Name); break;
                    }
                if (range == null)
                {
                    object filter = null;
                    try { filter = bag.Own(((dynamic)sheet).AutoFilter); } catch (Exception ex) { if (!Missing(ex)) throw; }
                    if (filter != null)
                    {
                        object candidate = bag.Own(((dynamic)filter).Range);
                        if (candidate != null && Contains(candidate)) { range = candidate; kind = "自动筛选"; }
                    }
                }
                if (range == null)
                {
                    object cell = bag.Own(((dynamic)sheet).Range[CellAddress.Name(anchorRow, anchorColumn)]);
                    range = bag.Own(((dynamic)cell).CurrentRegion);
                }
                int r = Convert.ToInt32(((dynamic)range).Row), height = Size(range, false);
                start = Convert.ToInt32(((dynamic)range).Column); width = Size(range, true);
                headerRow = headers == null ? r : Convert.ToInt32(((dynamic)headers).Row);
                if (start < 1 || width < 1 || start + width - 1 > 16384 || height < 1 || !Contains(range)) throw new InvalidOperationException("无法识别有效表格，请重新选择表格中的单元格。");
                if (kind == "连续区域" && height == 1 && width == 1 && ((dynamic)range).Value2 == null) throw new InvalidOperationException("所选单元格不在有效数据区域内，请重新选择。");
                string address = CellAddress.Name(r, start) + ":" + CellAddress.Name(r + height - 1, start + width - 1);
                RangeDescription = kind + "：" + address;
                return kind + "|" + name + "|" + address + "|" + headerRow;
            }
        }
        public IList<ColumnChoice> Collect()
        {
            if (disposed) throw new ObjectDisposedException("ComColumnVisibilityHost");
            if (collected != null) return collected;
            CheckSheetGroup();
            sheet = ((dynamic)app).ActiveSheet;
            if (sheet == null) throw new InvalidOperationException("请先选中工作表中的一行。");
            identity = WorksheetIdentity.Capture(sheet);
            object selection = ((dynamic)app).ActiveCell, selectedSheet = null;
            try
            {
                if (selection == null) throw new InvalidOperationException("请选中一行单元格。");
                selectedSheet = ((dynamic)selection).Worksheet;
                if (!identity.Matches(selectedSheet)) throw new InvalidOperationException("选区必须位于当前工作表。");
                anchorRow = Convert.ToInt32(((dynamic)selection).Row); anchorColumn = Convert.ToInt32(((dynamic)selection).Column);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException ex) { throw new InvalidOperationException("请选择工作表单元格，不能选择图形或其他对象。", ex); }
            finally { Release(selectedSheet); Release(selection); }
            CheckProtection(); tableSignature = Resolve(out row, out firstColumn, out count);
            List<ColumnChoice> choices = new List<ColumnChoice>();
            for (int c = firstColumn; c < firstColumn + count; c++)
            {
                string header, signature; Header(c, out header, out signature); signatures.Add(signature);
                bool hidden = ReadHidden(c);
                choices.Add(new ColumnChoice { Number = c, Letter = CellAddress.Name(1, c).TrimEnd('1'), Header = header, OriginalHidden = hidden, Visible = !hidden });
            }
            collected = choices; return collected;
        }
        private void Header(int column, out string text, out string signature)
        {
            object cell = ((dynamic)sheet).Range[CellAddress.Name(row, column)], merge = null, anchor = null;
            try
            {
                dynamic source = cell; string annotation = "";
                if (ComFormulaHost.Boolean(source.MergeCells, "MergeCells"))
                {
                    merge = source.MergeArea;
                    int r = Convert.ToInt32(((dynamic)merge).Row), c = Convert.ToInt32(((dynamic)merge).Column);
                    anchor = ((dynamic)sheet).Range[CellAddress.Name(r, c)]; source = anchor;
                    string address = ((dynamic)merge).Address as string;
                    if (address == null) throw new InvalidOperationException("无法读取合并标题区域。");
                    annotation = " [合并 " + address + "]";
                }
                object value = source.Value2, formula = source.Formula;
                string display = source.Text as string;
                if (display == null) throw new InvalidOperationException("无法可靠读取标题显示内容。");
                text = (display.Length == 0 ? "（空标题）" : display) + annotation;
                signature = (value == null ? "null" : value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture)) + "\n" + Convert.ToString(formula, CultureInfo.InvariantCulture) + "\n" + text;
            }
            finally { Release(anchor); Release(merge); Release(cell); }
        }
        public void Recheck()
        {
            if (disposed || collected == null) throw new InvalidOperationException("窗口快照不可用。");
            CheckSheetGroup(); object active = ((dynamic)app).ActiveSheet;
            try { if (!identity.Matches(active)) throw new InvalidOperationException("工作表身份已变化，请重新打开窗口。"); }
            finally { Release(active); }
            CheckProtection(); int start, width, headerRow; string current = Resolve(out headerRow, out start, out width);
            if (current != tableSignature || headerRow != row || start != firstColumn || width != count) throw new InvalidOperationException("表格范围已变化，请重新打开窗口。");
            for (int i = 0; i < count; i++)
            {
                string text, signature; Header(firstColumn + i, out text, out signature);
                if (signature != signatures[i] || ReadHidden(firstColumn + i) != collected[i].OriginalHidden) throw new InvalidOperationException(collected[i].Letter + " 列标题或显示状态已变化，请重新打开窗口。");
            }
        }
        public bool ReadHidden(int column)
        {
            object cell = ((dynamic)sheet).Range[CellAddress.Name(1, column)], entire = null;
            try { entire = ((dynamic)cell).EntireColumn; return ComFormulaHost.Boolean(((dynamic)entire).Hidden, "整列 Hidden"); }
            finally { Release(entire); Release(cell); }
        }
        public void WriteHidden(int column, bool hidden)
        {
            object cell = ((dynamic)sheet).Range[CellAddress.Name(1, column)], entire = null;
            try { entire = ((dynamic)cell).EntireColumn; ((dynamic)entire).Hidden = hidden; }
            finally { Release(entire); Release(cell); }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { if (identity != null) identity.Dispose(); }
            finally { identity = null; Release(sheet); sheet = null; collected = null; }
        }
    }
}
