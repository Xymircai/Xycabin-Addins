using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    internal interface IFormulaCell
    {
        string Address { get; }
        string Snapshot { get; }
        void CheckWritable();
        string Read();
        void Write(string formula);
    }
    internal interface IFormulaValueCell
    {
        bool OriginalIsFormula { get; }
        object OriginalValue { get; }
        bool MatchesOriginal();
        void RestoreOriginal();
    }
    internal interface IFormulaScopeHost { void RecheckScope(); }
    internal interface IFormulaHost : IDisposable
    {
        IList<IFormulaCell> Collect();
        bool EnableEvents { get; set; }
        bool ScreenUpdating { get; set; }
    }
    internal sealed class FormulaCommandResult
    {
        internal int Changed;
        internal string Details;
        internal bool Failed;
        internal bool RollbackIncomplete;
        internal string RecoveryLog;
    }

    internal static class FormulaCommandEngine
    {
        private static bool MatchesOriginal(Change change)
        { IFormulaValueCell typed = change.Cell as IFormulaValueCell; return typed != null ? typed.MatchesOriginal() : change.Cell.Read() == change.Original; }
        private sealed class Change
        {
            internal IFormulaCell Cell;
            internal string Original, Updated;
        }
        internal static FormulaCommandResult Run(IFormulaHost host, FormulaOperation operation, Action<string> log, string recoveryDirectory)
        { return Run(host, FormulaRequest.Default(operation), log, recoveryDirectory); }

        internal static FormulaCommandResult Run(IFormulaHost host, FormulaRequest request, Action<string> log, string recoveryDirectory)
        {
            FormulaCommandResult result = new FormulaCommandResult();
            List<Change> changes = new List<Change>(), attempted = new List<Change>();
            bool events = false, screen = false, captured = false, disabling = false;
            try
            {
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (IFormulaCell cell in host.Collect())
                {
                    if (!seen.Add(cell.Address)) continue;
                    string original = cell.Snapshot;
                    IFormulaValueCell typed = cell as IFormulaValueCell;
                    if (typed != null && !typed.OriginalIsFormula && !request.ProcessesValues) continue;
                    string updated;
                    if (typed != null && !typed.OriginalIsFormula)
                        updated = request.Operation == FormulaOperation.Custom
                            ? "=" + request.ApplyCustom(FormulaRequest.SerializeValue(typed.OriginalValue))
                            : "=" + (FormulaRequest.NumericText(typed.OriginalValue).StartsWith("-") ? "(" + FormulaRequest.NumericText(typed.OriginalValue) + ")" : FormulaRequest.NumericText(typed.OriginalValue)) + request.Operator + request.Argument;
                    else
                        updated = FormulaWrapper.Transform(original, request);
                    if (updated != original) changes.Add(new Change { Cell = cell, Original = original, Updated = updated });
                }
                // Compute the whole batch before checking any write or touching host state.
                IFormulaScopeHost scope = host as IFormulaScopeHost;
                if (scope != null) scope.RecheckScope();
                foreach (Change change in changes) change.Cell.CheckWritable();
                if (changes.Count == 0)
                {
                    if (request.Operation == FormulaOperation.Custom)
                        result.Details = "当前选区没有可处理的公式或文本/数值常量。\r\n空白、布尔、错误常量及不适用的内容已跳过。";
                    else if (request.IsArithmetic)
                        result.Details = "当前选区没有可处理的公式或数值常量。\r\n文本、布尔、错误常量、空白及不适用的内容已跳过。";
                    else
                        result.Details = "当前选区没有可处理的公式。\r\n文本、布尔、错误常量、空白及不适用的内容已跳过。\r\n公式封装与移除仅处理已有公式；处理数值请使用「一键公式 → 数值计算」。";
                    return result;
                }
                events = host.EnableEvents; screen = host.ScreenUpdating; captured = true;
                disabling = true;
                host.EnableEvents = false; host.ScreenUpdating = false;
                if (host.EnableEvents || host.ScreenUpdating) throw new InvalidOperationException("宿主未接受关闭事件或屏幕刷新，未写入公式。");
                if (scope != null) scope.RecheckScope();
                // Full second check before the first write, plus a check at each write.
                foreach (Change change in changes)
                    if (!MatchesOriginal(change)) throw new InvalidOperationException(change.Cell.Address + " 的公式已变化，请重新执行。");
                foreach (Change change in changes)
                {
                    if (!MatchesOriginal(change)) throw new InvalidOperationException(change.Cell.Address + " 的公式已变化。");
                    attempted.Add(change); // A failing setter can already have mutated the cell.
                    change.Cell.Write(change.Updated);
                    // WPS stores recognized built-in names in upper case, so a lower-case
                    // template comes back upper-cased. Function names are case-insensitive
                    // in formula semantics, so only a real content difference should fail.
                    if (!string.Equals(change.Cell.Read(), change.Updated, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException(change.Cell.Address + " 写入后的公式与目标不一致。");
                }
                result.Changed = changes.Count;
            }
            catch (Exception error)
            {
                result.Failed = true; result.Changed = 0;
                result.Details = "一键公式未完成：" + error.Message;
                List<string> failures = new List<string>();
                for (int i = attempted.Count - 1; i >= 0; i--)
                {
                    Change change = attempted[i];
                    string writeError = null;
                    try { IFormulaValueCell typed = change.Cell as IFormulaValueCell; if (typed != null) typed.RestoreOriginal(); else change.Cell.Write(change.Original); } catch (Exception ex) { writeError = ex.Message; }
                    try
                    {
                        if (!MatchesOriginal(change)) throw new InvalidOperationException("核对结果与原公式不一致");
                    }
                    catch (Exception ex)
                    {
                        IFormulaValueCell typed = change.Cell as IFormulaValueCell;
                        string snapshot = typed != null && !typed.OriginalIsFormula ? "原数值（" + typed.OriginalValue.GetType().FullName + "）：" + FormulaRequest.NumericText(typed.OriginalValue) : "原公式：" + change.Original;
                        failures.Add("地址：" + change.Cell.Address + "\r\n" + snapshot + "\r\n错误：" + (writeError == null ? "" : writeError + "；") + ex.Message);
                    }
                }
                if (failures.Count != 0)
                {
                    result.RollbackIncomplete = true;
                    result.Details += "\r\n\r\n以下单元格未能确认还原，请按原内容手动恢复：\r\n" + string.Join("\r\n\r\n", failures);
                    try
                    {
                        Directory.CreateDirectory(recoveryDirectory);
                        result.RecoveryLog = Path.Combine(recoveryDirectory, "FormulaRecovery-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N") + ".txt");
                        File.WriteAllText(result.RecoveryLog, result.Details, System.Text.Encoding.UTF8);
                        result.Details += "\r\n\r\n恢复日志：" + result.RecoveryLog;
                    }
                    catch (Exception ex) { result.Details += "\r\n恢复日志保存失败：" + ex.Message + "。请复制本窗口详情。"; }
                }
                else result.Details += attempted.Count == 0 ? "\r\n本批未写入任何单元格。" : "\r\n本批已尝试写入的单元格均已还原并核对。";
                if (log != null) log(result.Details);
            }
            finally
            {
                List<string> cleanup = new List<string>();
                if (captured && disabling)
                {
                    try { host.ScreenUpdating = screen; if (host.ScreenUpdating != screen) throw new Exception("恢复核对失败"); }
                    catch (Exception ex) { cleanup.Add("屏幕刷新恢复失败：" + ex.Message); }
                    try { host.EnableEvents = events; if (host.EnableEvents != events) throw new Exception("恢复核对失败"); }
                    catch (Exception ex) { cleanup.Add("事件状态恢复失败：" + ex.Message); }
                }
                try { host.Dispose(); } catch (Exception ex) { cleanup.Add("临时引用释放失败：" + ex.Message); }
                if (cleanup.Count > 0)
                {
                    result.Failed = true;
                    result.Details = (result.Details ?? ("已修改 " + result.Changed + " 个单元格。")) + "\r\n" + string.Join("\r\n", cleanup);
                    if (log != null) log(result.Details);
                }
            }
            return result;
        }
    }

    // Late binding stays on the Ribbon host thread. Only references acquired here
    // are released; the shared Application is borrowed and never released.
    internal enum FormulaCapture { Formulas, Numbers, Values }

    internal sealed class ComFormulaHost : IFormulaHost, IFormulaScopeHost
    {
        private readonly object app;
        private readonly FormulaCapture capture;
        // Replaces Application.Selection when set, so the custom function can act on
        // the range the user picked after the editor hid itself.
        private readonly object picked;
        private WorksheetIdentity identity;
        private string context;
        private IList<IFormulaCell> collected;
        private bool disposed;
        private readonly List<object> owned = new List<object>();
        internal ComFormulaHost(object application) : this(application, FormulaCapture.Formulas) { }
        internal ComFormulaHost(object application, bool numbers) : this(application, numbers ? FormulaCapture.Numbers : FormulaCapture.Formulas) { }
        internal ComFormulaHost(object application, FormulaCapture mode) : this(application, mode, null) { }
        internal ComFormulaHost(object application, FormulaCapture mode, object range)
        { app = application; capture = mode; picked = range; }
        private bool KeepValue(object number)
        {
            if (capture == FormulaCapture.Numbers) return FormulaRequest.IsNumeric(number);
            if (capture == FormulaCapture.Values) return FormulaRequest.IsScalarValue(number);
            return false;
        }
        private object Own(object value) { if (value != null) owned.Add(value); return value; }
        public bool EnableEvents { get { return Boolean(((dynamic)app).EnableEvents, "EnableEvents"); } set { ((dynamic)app).EnableEvents = value; } }
        public bool ScreenUpdating { get { return Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating"); } set { ((dynamic)app).ScreenUpdating = value; } }
        internal static bool Boolean(object value, string name)
        {
            if (!(value is bool)) throw new InvalidOperationException("无法可靠读取 " + name + "，整批停止。");
            return (bool)value;
        }
        internal static bool Missing(Exception exception)
        {
            COMException com = exception as COMException;
            return exception is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException || exception is MissingMemberException
                || (com != null && (com.ErrorCode == unchecked((int)0x80020003) || com.ErrorCode == unchecked((int)0x80020006)));
        }
        public IList<IFormulaCell> Collect()
        {
            if (disposed) throw new ObjectDisposedException("ComFormulaHost");
            if (collected != null) return collected;
            List<IFormulaCell> result = new List<IFormulaCell>();
            dynamic window = Own(((dynamic)app).ActiveWindow);
            if (window == null) throw new InvalidOperationException("请先打开工作簿并选中单元格。");
            dynamic selectedSheets = Own(window.SelectedSheets);
            if (Convert.ToInt32(selectedSheets.Count) != 1) throw new InvalidOperationException("工作表正在成组选择，请先取消工作表组合。");
            dynamic sheet = Own(((dynamic)app).ActiveSheet);
            dynamic selection = picked != null ? Own(picked) : Own(((dynamic)app).Selection);
            if (sheet == null || selection == null) throw new InvalidOperationException("请先选中工作表单元格。");
            dynamic selectionSheet = Own(selection.Worksheet);
            identity = WorksheetIdentity.Capture((object)sheet);
            if (!identity.Matches((object)selectionSheet))
                throw new InvalidOperationException("只能处理当前工作表的选区。");
            context = "工作簿：" + identity.Path + "；工作表：" + identity.Name + "；单元格：";
            dynamic used = Own(sheet.UsedRange);
            dynamic areas = Own(selection.Areas);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int count = Convert.ToInt32(areas.Count);
            for (int a = 1; a <= count; a++)
            {
                dynamic area = Own(areas.Item(a));
                dynamic intersection = null;
                try
                {
                    intersection = Own(((dynamic)app).Intersect(area, used));
                    if (intersection == null) continue;
                    int row = Convert.ToInt32(intersection.Row), col = Convert.ToInt32(intersection.Column);
                    dynamic rows = Own(intersection.Rows), columns = Own(intersection.Columns);
                    int rowCount = Convert.ToInt32(rows.Count), colCount = Convert.ToInt32(columns.Count);
                    for (int x = 0; x < colCount; x += 256)
                    {
                        int width = Math.Min(256, colCount - x), rowStep = Math.Max(1, 4096 / width);
                        for (int y = 0; y < rowCount; y += rowStep)
                        {
                            int height = Math.Min(rowStep, rowCount - y);
                            CollectBlock(sheet, row + y, col + x, height, width, seen, result);
                        }
                    }
                }
                finally { Drop(intersection); Drop(area); }
            }
            collected = result;
            return result;
        }
        public void RecheckScope()
        {
            dynamic window = Own(((dynamic)app).ActiveWindow), selected = null, sheet = null;
            try
            {
                if (window == null) throw new InvalidOperationException("工作簿窗口已关闭，请重新执行。");
                selected = Own(window.SelectedSheets);
                if (Convert.ToInt32(selected.Count) != 1) throw new InvalidOperationException("工作表正在成组选择，请先取消组合。");
                sheet = Own(((dynamic)app).ActiveSheet);
                if (identity == null || !identity.Matches((object)sheet))
                    throw new InvalidOperationException("当前工作表与原目标不同，请返回打开参数窗时的工作表；目标已关闭、替换或改名时请重新执行。");
            }
            finally { Drop(sheet); Drop(selected); Drop(window); }
        }
        private void CollectBlock(dynamic sheet, int row, int col, int height, int width, HashSet<string> seen, List<IFormulaCell> result)
        {
            dynamic block = Own(sheet.Range[CellAddress.Name(row, col) + ":" + CellAddress.Name(row + height - 1, col + width - 1)]);
            try
            {
                object has = block.HasFormula;
                if (capture == FormulaCapture.Formulas && has is bool && !(bool)has) return;
                if (!(has is bool) && has != null && has != DBNull.Value) throw new InvalidOperationException("HasFormula 返回无法识别的值。");
                bool formula2 = true; object formulas;
                try { formulas = block.Formula2; }
                catch (Exception ex) { if (!Missing(ex)) throw; formula2 = false; formulas = block.Formula; }
                object constants = capture != FormulaCapture.Formulas ? (object)block.Value2 : null;
                Array numericValues = constants as Array;
                if (capture != FormulaCapture.Formulas && height * width > 1 && (numericValues == null || numericValues.Rank != 2 || numericValues.GetLength(0) != height || numericValues.GetLength(1) != width)) throw new InvalidOperationException("无法可靠批量读取选区数值。");
                Array values = formulas as Array;
                if (height * width > 1 && (values == null || values.Rank != 2 || values.GetLength(0) != height || values.GetLength(1) != width))
                    throw new InvalidOperationException("无法可靠批量读取选区公式。");
                for (int r = 0; r < height; r++) for (int c = 0; c < width; c++)
                {
                    string address = CellAddress.Name(row + r, col + c);
                    if (!seen.Add(address)) continue;
                    dynamic cell = Own(sheet.Range[address]);
                    bool keep = false;
                    try
                    {
                        bool isFormula = Boolean(cell.HasFormula, address + " HasFormula");
                        object number = numericValues == null ? constants : numericValues.GetValue(r + numericValues.GetLowerBound(0), c + numericValues.GetLowerBound(1));
                        if (!isFormula && !KeepValue(number)) continue;
                        if (Boolean(cell.MergeCells, address + " MergeCells"))
                        {
                            dynamic merge = Own(cell.MergeArea);
                            try { if (Convert.ToInt32(merge.Row) != row + r || Convert.ToInt32(merge.Column) != col + c) continue; }
                            finally { Drop(merge); }
                        }
                        object value = values == null ? formulas : values.GetValue(r + values.GetLowerBound(0), c + values.GetLowerBound(1));
                        string formula = value as string;
                        if (!isFormula) formula = "=" + FormulaRequest.SerializeValue(number);
                        if (formula == null) throw new InvalidOperationException(address + " 公式读取失败。");
                        result.Add(new ComCell(cell, sheet, context + address, formula, formula2, isFormula, number)); keep = true;
                    }
                    finally { if (!keep) Drop(cell); }
                }
            }
            finally { Drop(block); }
        }
        private void Drop(object value)
        {
            if (value == null) return;
            for (int i = owned.Count - 1; i >= 0; i--)
                if (ReferenceEquals(owned[i], value)) { owned.RemoveAt(i); Release(value); break; }
        }
        private static void Release(object value) { if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true; collected = null;
            List<string> errors = new List<string>();
            try { if (identity != null) identity.Dispose(); } catch (Exception ex) { errors.Add(ex.Message); }
            identity = null;
            for (int i = owned.Count - 1; i >= 0; i--) try { Release(owned[i]); } catch (Exception ex) { errors.Add(ex.Message); }
            owned.Clear();
            if (errors.Count > 0) throw new InvalidOperationException(string.Join("；", errors));
        }
        private sealed class ComCell : IFormulaCell, IFormulaValueCell
        {
            private readonly object cell;
            private readonly bool formula2;
            private readonly object worksheet;
            public string Address { get; private set; }
            public string Snapshot { get; private set; }
            public bool OriginalIsFormula { get; private set; }
            public object OriginalValue { get; private set; }
            public bool MatchesOriginal()
            {
                if (ComFormulaHost.Boolean((object)((dynamic)cell).HasFormula, "HasFormula") != OriginalIsFormula) return false;
                if (OriginalIsFormula) return Read() == Snapshot;
                object actual = ((dynamic)cell).Value2;
                return actual != null && OriginalValue != null && actual.GetType() == OriginalValue.GetType() && actual.Equals(OriginalValue);
            }
            public void RestoreOriginal()
            { if (OriginalIsFormula) Write(Snapshot); else ((dynamic)cell).Value2 = OriginalValue; }
            internal ComCell(object range, object sheet, string address, string snapshot, bool useFormula2, bool isFormula, object value)
            { cell = range; worksheet = sheet; Address = address; Snapshot = snapshot; formula2 = useFormula2; OriginalIsFormula = isFormula; OriginalValue = value; }
            public void CheckWritable()
            {
                try
                {
                    dynamic range = cell;
                    if (ComFormulaHost.Boolean((object)range.HasFormula, "HasFormula") != OriginalIsFormula) throw new InvalidOperationException("单元格已不再是公式，请重新执行");
                    if (ComFormulaHost.Boolean((object)range.HasArray, "HasArray")) throw new InvalidOperationException("数组公式");
                    if (ComFormulaHost.Boolean((object)range.HasSpill, "HasSpill")) throw new InvalidOperationException("动态溢出公式");
                    if (ComFormulaHost.Boolean((object)((dynamic)worksheet).ProtectContents, "ProtectContents") && ComFormulaHost.Boolean((object)range.Locked, "Locked")) throw new InvalidOperationException("受保护且锁定的单元格");
                    if (!MatchesOriginal()) throw new InvalidOperationException("普通公式核对失败");
                }
                catch (Exception ex) { throw new InvalidOperationException(Address + "：" + ex.Message, ex); }
            }
            public string Read()
            {
                if (!ComFormulaHost.Boolean((object)((dynamic)cell).HasFormula, "HasFormula")) throw new InvalidOperationException(Address + " 已不是公式。");
                object value = formula2 ? ((dynamic)cell).Formula2 : ((dynamic)cell).Formula;
                string formula = value as string;
                if (formula == null) throw new InvalidOperationException(Address + " 无法读取公式。");
                return formula;
            }
            public void Write(string formula)
            { if (formula2) ((dynamic)cell).Formula2 = formula; else ((dynamic)cell).Formula = formula; }
        }
    }
}
