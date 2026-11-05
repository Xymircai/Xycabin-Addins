using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Excel_Iterate
{
    internal enum CellConversionKind { TextFilter, ReferenceStyle, FormulaAsText, ResultAsText, ValuesOnly, RestoreTextFormula }
    internal enum ReferenceStyle { Absolute, RelativeColumnAbsoluteRow, AbsoluteColumnRelativeRow, Relative }
    internal sealed class ConversionRequest
    {
        internal CellConversionKind Kind;
        internal string Address;
        internal bool Extract = true;
        internal CellCharacterKinds Kinds = CellCharacterKinds.Chinese;
        internal ReferenceStyle Style = ReferenceStyle.Relative;
        internal void Validate()
        {
            if (Kind < CellConversionKind.TextFilter || Kind > CellConversionKind.RestoreTextFormula) throw new ArgumentException("转换命令无效。");
            if (Kind == CellConversionKind.TextFilter && (Kinds == CellCharacterKinds.None || (Kinds & ~CellCharacterKinds.All) != 0)) throw new ArgumentException("请至少勾选一种字符类型。");
            if (Kind == CellConversionKind.ReferenceStyle && (Style < ReferenceStyle.Absolute || Style > ReferenceStyle.Relative)) throw new ArgumentException("请选择有效引用方式。");
        }
    }
    internal sealed class ConversionCellState
    {
        internal bool IsFormula;
        internal string Formula, NumberFormat, Display;
        internal object Value;
        internal ConversionCellState Copy() { return (ConversionCellState)MemberwiseClone(); }
        internal static bool SameValue(object a, object b)
        {
            var ea = a as ErrorWrapper; var eb = b as ErrorWrapper;
            if (ea != null || eb != null) return ea != null && eb != null && ea.ErrorCode == eb.ErrorCode;
            return a == null ? b == null : b != null && a.GetType() == b.GetType() && a.Equals(b);
        }
        internal static bool SameFormula(string expected, string actual)
        {
            if (expected == actual) return true;
            if (expected == null || actual == null || expected.Length != actual.Length) return false;
            bool literal = false;
            for (int i = 0; i < expected.Length; i++)
            {
                char a = expected[i], b = actual[i];
                if (a == '"' || b == '"') { if (a != b) return false; literal = !literal; }
                else if (literal ? a != b : char.ToUpperInvariant(a) != char.ToUpperInvariant(b)) return false;
            }
            return true;
        }
        internal bool Matches(ConversionCellState actual, bool compareResult, bool allowFormulaCase = false)
        {
            return actual != null && IsFormula == actual.IsFormula && NumberFormat == actual.NumberFormat &&
                (IsFormula ? (allowFormulaCase ? SameFormula(Formula, actual.Formula) : Formula == actual.Formula) && (!compareResult || SameValue(Value, actual.Value)) : SameValue(Value, actual.Value));
        }
        internal string Describe()
        { return "公式：" + (IsFormula ? Formula : "无") + "；值类型：" + (Value == null ? "空白" : Value.GetType().FullName) + "；原值：" + (Value is ErrorWrapper ? ((ErrorWrapper)Value).ErrorCode.ToString(CultureInfo.InvariantCulture) : System.Convert.ToString(Value, CultureInfo.InvariantCulture)) + "；数字格式：" + NumberFormat; }
    }
    internal interface IConversionCell
    {
        string Address { get; }
        ConversionCellState Original { get; }
        void CheckWritable();
        ConversionCellState Read();
        void Write(ConversionCellState state);
    }
    internal interface IConversionHost : IDisposable
    {
        IList<IConversionCell> Collect();
        void RecheckScope();
        void Calculate(IList<IConversionCell> targets);
        bool EnableEvents { get; set; }
        bool ScreenUpdating { get; set; }
    }
    internal sealed class CellConversionResult
    {
        internal int Changed;
        internal bool Failed, RollbackIncomplete;
        internal string Details, RecoveryLog;
    }
    internal static class CellConversionEngine
    {
        private sealed class Change { internal IConversionCell Cell; internal ConversionCellState Before, After; }
        internal static ConversionCellState Transform(ConversionCellState original, ConversionRequest request)
        {
            request.Validate(); var next = original.Copy();
            if (request.Kind == CellConversionKind.ReferenceStyle)
            { if (!original.IsFormula) return null; next.Formula = ReferenceStyleConverter.Convert(original.Formula, request.Style); return next; }
            if (request.Kind == CellConversionKind.FormulaAsText)
            { if (!original.IsFormula) return null; next.IsFormula = false; next.Value = original.Formula; next.Formula = null; return next; }
            if (request.Kind == CellConversionKind.RestoreTextFormula)
            {
                string expression = original.Value as string;
                if (original.IsFormula || expression == null || !expression.StartsWith("=", StringComparison.Ordinal)) return null;
                next.IsFormula = true; next.Formula = expression;
                if (next.NumberFormat == "@") next.NumberFormat = "General";
                return next;
            }
            if (request.Kind == CellConversionKind.ValuesOnly)
            { if (!original.IsFormula) return null; next.IsFormula = false; next.Formula = null; if (next.Value is string && ((string)next.Value).Length == 0) next.Value = null; return next; }
            if (request.Kind == CellConversionKind.TextFilter && original.Value == null && !original.IsFormula) return null;
            string text = original.Display ?? "";
            if (IsNumeric(original.Value) && text.Trim().Length > 0 && text.Trim().Trim('#').Length == 0)
                throw new InvalidOperationException("数值显示为 ####，请先调宽该列后重试。");
            if (request.Kind == CellConversionKind.TextFilter) text = CellCharacterClassifier.Filter(text, request.Kinds, request.Extract);
            next.IsFormula = false; next.Formula = null; next.NumberFormat = "@"; next.Value = text.Length == 0 ? null : (object)text; return next;
        }
        internal static bool IsNumeric(object value)
        {
            if (value is int && ((uint)(int)value & 0xffff0000u) == 0x800a0000u || value is uint && ((uint)value & 0xffff0000u) == 0x800a0000u) return false;
            return value is double || value is float || value is decimal || value is int || value is long || value is short || value is byte || value is uint || value is ulong || value is ushort || value is sbyte;
        }
        internal static CellConversionResult Run(IConversionHost host, ConversionRequest request, string recoveryDirectory)
        {
            var result = new CellConversionResult(); var changes = new List<Change>(); var attempted = new List<Change>();
            bool captured = false, events = false, screen = false;
            try
            {
                request.Validate(); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in host.Collect())
                {
                    if (!seen.Add(cell.Address)) continue;
                    ConversionCellState next;
                    try { next = Transform(cell.Original, request); }
                    catch (Exception error) { throw new InvalidOperationException(cell.Address + "：" + error.Message, error); }
                    if (next != null && !next.Matches(cell.Original, false)) changes.Add(new Change { Cell = cell, Before = cell.Original, After = next });
                }
                host.RecheckScope();
                foreach (var change in changes) change.Cell.CheckWritable();
                if (changes.Count == 0) { result.Details = "当前选区没有需要转换的单元格。"; return result; }
                events = host.EnableEvents; screen = host.ScreenUpdating; captured = true;
                host.EnableEvents = false; host.ScreenUpdating = false;
                if (host.EnableEvents || host.ScreenUpdating) throw new InvalidOperationException("宿主未接受事件或屏幕刷新设置。");
                host.RecheckScope();
                // Verify all cached results BEFORE writes; later results may recalculate as this batch changes predecessors.
                foreach (var change in changes)
                {
                    change.Cell.CheckWritable();
                    if (!change.Before.Matches(change.Cell.Read(), true)) throw new InvalidOperationException(change.Cell.Address + " 的内容、结果或格式已改变，请重新执行。");
                }
                foreach (var change in changes)
                {
                    change.Cell.CheckWritable();
                    if (!change.Before.Matches(change.Cell.Read(), false)) throw new InvalidOperationException(change.Cell.Address + " 的内容或格式已改变。");
                    attempted.Add(change);
                    change.Cell.Write(change.After);
                    var written = change.Cell.Read(); bool restoring = request.Kind == CellConversionKind.RestoreTextFormula;
                    if (!change.After.Matches(written, false, restoring)) throw new InvalidOperationException(change.Cell.Address + " 写入核验失败。");
                    if (restoring) change.After.Formula = written.Formula; // Accept host function/name casing; literal strings remain exact.
                    change.Cell.CheckWritable();
                }
                if (request.Kind == CellConversionKind.RestoreTextFormula)
                {
                    host.RecheckScope();
                    var targets = new List<IConversionCell>(); foreach (var change in changes) targets.Add(change.Cell);
                    host.Calculate(targets);
                    foreach (var change in changes)
                    {
                        change.Cell.CheckWritable();
                        if (!change.After.Matches(change.Cell.Read(), false)) throw new InvalidOperationException(change.Cell.Address + " 计算后公式或格式核验失败。");
                    }
                }
                result.Changed = changes.Count; result.Details = "已转换 " + result.Changed + " 个单元格。";
            }
            catch (Exception error)
            {
                result.Failed = true; result.Changed = 0; result.Details = "转换未完成：" + error.Message;
                var failures = new List<string>();
                for (int i = attempted.Count - 1; i >= 0; i--)
                {
                    var change = attempted[i]; string writeError = "";
                    try { change.Cell.Write(change.Before); } catch (Exception ex) { writeError = ex.Message; }
                    try { if (!change.Before.Matches(change.Cell.Read(), false)) throw new InvalidOperationException("原内容或数字格式未恢复"); }
                    catch (Exception ex) { failures.Add(change.Cell.Address + "\r\n" + change.Before.Describe() + "\r\n" + writeError + "；" + ex.Message); }
                }
                result.RollbackIncomplete = failures.Count > 0;
                if (result.RollbackIncomplete)
                {
                    result.Details += "\r\n以下单元格需手工恢复：\r\n" + string.Join("\r\n", failures);
                    try
                    {
                        Directory.CreateDirectory(recoveryDirectory);
                        result.RecoveryLog = Path.Combine(recoveryDirectory, "ConversionRecovery-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N") + ".txt");
                        File.WriteAllText(result.RecoveryLog, result.Details, Encoding.UTF8); result.Details += "\r\n恢复日志：" + result.RecoveryLog;
                    }
                    catch (Exception ex) { result.Details += "\r\n恢复日志保存失败，请复制本窗口详情：" + ex.Message; }
                }
                else result.Details += attempted.Count == 0 ? "\r\n本批未写入任何单元格。" : "\r\n本批已尝试的内容及格式均已还原并核对。";
            }
            finally
            {
                var cleanup = new List<string>();
                if (captured)
                {
                    try { host.ScreenUpdating = screen; if (host.ScreenUpdating != screen) throw new InvalidOperationException("核验失败"); } catch (Exception ex) { cleanup.Add("屏幕刷新恢复失败：" + ex.Message); }
                    try { host.EnableEvents = events; if (host.EnableEvents != events) throw new InvalidOperationException("核验失败"); } catch (Exception ex) { cleanup.Add("事件恢复失败：" + ex.Message); }
                }
                try { host.Dispose(); } catch (Exception ex) { cleanup.Add("临时引用释放失败：" + ex.Message); }
                if (cleanup.Count > 0) { result.Failed = true; result.Details += "\r\n" + string.Join("\r\n", cleanup); }
            }
            return result;
        }
    }
}
