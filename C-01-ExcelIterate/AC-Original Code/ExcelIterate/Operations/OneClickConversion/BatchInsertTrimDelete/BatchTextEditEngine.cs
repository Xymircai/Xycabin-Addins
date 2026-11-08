using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Excel_Iterate
{
    internal interface IBatchTextPosition { string LocalAddress { get; } }
    internal interface IConversionCalculationControl { int Calculation { get; set; } }
    internal interface IBatchTextAccess
    {
        IList<IConversionCell> CollectTextEdit();
        void CheckTextBlocks(IList<BatchEditBlock> blocks, bool compareResults);
        void WriteTextBlock(BatchEditBlock block);
    }
    internal interface IBatchFormulaSyntax
    { IList<BatchEditChange> ValidateFormulaSyntax(IList<BatchEditChange> candidates); }
    internal sealed class BatchSyntaxReport
    {
        internal readonly int Count;
        internal readonly string Message;
        internal BatchSyntaxReport(IList<BatchEditChange> invalid)
        {
            Count = invalid.Count;
            Message = "本次有 " + Count + " 个单元格的公式在修改后语法无效，不能作为公式保存。\r\n继续操作将使用单引号前缀，将这些修改后的表达式保存为文本格式。其他结果按原定方式处理。\r\n\r\n"
                + string.Join("\r\n", invalid.Take(5).Select(c => c.LocalAddress + "：" + Short(c.Before.Formula) + " → " + Short(c.After.Formula)).ToArray())
                + (Count > 5 ? "\r\n另有 " + (Count - 5) + " 个。" : "") + "\r\n\r\n是否继续并转为文本？取消不会写入原区域。";
        }
        private static string Short(string value) { return value != null && value.Length > 160 ? value.Substring(0, 160) + "…" : value; }
    }
    internal sealed class BatchEditChange
    {
        internal IConversionCell Cell;
        internal string LocalAddress;
        internal int Row, Column;
        internal ConversionCellState Before, After;
        internal bool ConfirmedText;
    }
    internal sealed class BatchEditBlock
    {
        internal readonly List<BatchEditChange> Changes = new List<BatchEditChange>();
        internal int Top, Bottom, Left, Right;
        internal bool Restore, QuoteText;
        internal string NumberFormat;
        internal bool IsFormula;
        internal string Address { get { string first = CellAddress.Name(Top, Left), last = CellAddress.Name(Bottom, Right); return first == last ? first : first + ":" + last; } }
        internal ConversionCellState State(BatchEditChange change) { return Restore ? change.Before : change.After; }
        internal static IList<BatchEditBlock> Create(IEnumerable<BatchEditChange> changes, bool restore)
        {
            var result = new List<BatchEditBlock>();
            foreach (var group in changes.GroupBy(c => (restore ? c.Before : c.After).NumberFormat + "\0" + ((restore ? c.Before : c.After).IsFormula ? "F" : !restore && c.ConfirmedText ? "Q" : "V")))
            {
                var sorted = group.OrderBy(c => c.Row).ThenBy(c => c.Column).ToList();
                var previous = new Dictionary<string, BatchEditBlock>();
                for (int i = 0; i < sorted.Count;)
                {
                    int start = i, row = sorted[i].Row, left = sorted[i].Column, right = left; i++;
                    while (i < sorted.Count && sorted[i].Row == row && sorted[i].Column == right + 1 && i - start < 1024) { right++; i++; }
                    string key = left + ":" + right; BatchEditBlock block;
                    if (!previous.TryGetValue(key, out block) || block.Bottom != row - 1 || block.Changes.Count + i - start > 1024)
                    {
                        var state = restore ? sorted[start].Before : sorted[start].After;
                        block = new BatchEditBlock { Top = row, Bottom = row, Left = left, Right = right, Restore = restore, NumberFormat = state.NumberFormat, IsFormula = state.IsFormula, QuoteText = !restore && sorted[start].ConfirmedText };
                        previous[key] = block; result.Add(block);
                    }
                    block.Bottom = row; for (int at = start; at < i; at++) block.Changes.Add(sorted[at]);
                }
            }
            return result;
        }
        internal static IList<BatchEditBlock> Addresses(IEnumerable<string> names)
        {
            return Create(names.Select(name => { var point = ConversionRangeAddress.Coordinates(name); var state = new ConversionCellState { NumberFormat = "" }; return new BatchEditChange { LocalAddress = name, Row = point[0], Column = point[1], Before = state, After = state }; }), false);
        }
    }
    internal static class BatchTextEditEngine
    {
        internal static bool IsSyntaxRejection(COMException error)
        { return error.ErrorCode == unchecked((int)0x800a03ec) || error.ErrorCode == unchecked((int)0x80020009) || error.ErrorCode == unchecked((int)0x8fe3000b); }
        private sealed class ApplicationState : IDisposable
        {
            private readonly IConversionHost host;
            private readonly IConversionCalculationControl control;
            private readonly bool events, screen;
            private readonly int calculation;
            private bool active;
            internal ApplicationState(IConversionHost value)
            {
                host = value; control = host as IConversionCalculationControl;
                events = host.EnableEvents; screen = host.ScreenUpdating; if (control != null) calculation = control.Calculation;
                active = true;
                try
                {
                    host.EnableEvents = false; host.ScreenUpdating = false;
                    if (control != null) { control.Calculation = -4135; if (control.Calculation != -4135) throw new InvalidOperationException("宿主未接受临时手动计算模式。"); }
                    if (host.EnableEvents || host.ScreenUpdating) throw new InvalidOperationException("宿主未接受事件或屏幕刷新设置。");
                }
                catch { Dispose(); throw; }
            }
            public void Dispose()
            {
                if (!active) return; active = false; var errors = new List<string>();
                try { if (control != null) { control.Calculation = calculation; if (control.Calculation != calculation) throw new InvalidOperationException("计算模式恢复核验失败"); } } catch (Exception error) { errors.Add(error.Message); }
                try { host.ScreenUpdating = screen; if (host.ScreenUpdating != screen) throw new InvalidOperationException("屏幕刷新恢复核验失败"); } catch (Exception error) { errors.Add(error.Message); }
                try { host.EnableEvents = events; if (host.EnableEvents != events) throw new InvalidOperationException("事件恢复核验失败"); } catch (Exception error) { errors.Add(error.Message); }
                if (errors.Count > 0) throw new InvalidOperationException(string.Join("；", errors.ToArray()));
            }
        }
        private static void Check(IConversionHost host, IList<BatchEditBlock> blocks, bool results)
        {
            host.RecheckScope(); var access = host as IBatchTextAccess;
            if (access != null) { access.CheckTextBlocks(blocks, results); return; }
            foreach (var block in blocks) foreach (var change in block.Changes)
            {
                change.Cell.CheckWritable(); var expected = block.State(change); var actual = change.Cell.Read();
                if (!expected.Matches(actual, results, !block.Restore)) throw new InvalidOperationException(change.LocalAddress + " 内容、公式或格式核验失败。");
                if (!block.Restore && expected.IsFormula) expected.Formula = actual.Formula;
            }
        }
        private static void Write(IConversionHost host, BatchEditBlock block)
        {
            var access = host as IBatchTextAccess;
            if (access != null) access.WriteTextBlock(block);
            else foreach (var change in block.Changes) change.Cell.Write(block.State(change));
        }
        internal static CellConversionResult Run(IConversionHost host, ConversionRequest request, string recoveryDirectory, Func<BatchSyntaxReport, bool> confirm)
        {
            var result = new CellConversionResult(); var changes = new List<BatchEditChange>(); var attempted = new List<BatchEditChange>(); ApplicationState scope = null;
            var clock = Stopwatch.StartNew(); long read = 0, syntax = 0, write = 0, verify = 0, restore = 0;
            try
            {
                request.Validate(); var access = host as IBatchTextAccess; var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var cell in access == null ? host.Collect() : access.CollectTextEdit())
                {
                    if (!seen.Add(cell.Address)) continue;
                    var after = BatchTextEdit.TransformState(cell.Original, request.TextEdit); if (after == null || after.Matches(cell.Original, false)) continue;
                    var position = cell as IBatchTextPosition; string address = position == null ? cell.Address : position.LocalAddress; var point = ConversionRangeAddress.Coordinates(address);
                    changes.Add(new BatchEditChange { Cell = cell, LocalAddress = address, Row = point[0], Column = point[1], Before = cell.Original, After = after });
                }
                read = clock.ElapsedMilliseconds;
                if (changes.Count == 0) { result.Details = "当前选区没有需要转换的单元格。"; return result; }
                scope = new ApplicationState(host); var originals = BatchEditBlock.Create(changes, true); Check(host, originals, true);
                var candidates = changes.Where(c => c.After.IsFormula).ToList();
                if (candidates.Count > 0)
                {
                    var validator = host as IBatchFormulaSyntax; if (validator == null) throw new InvalidOperationException("宿主不支持隔离公式预检，原区域未写入。");
                    var invalid = validator.ValidateFormulaSyntax(candidates); syntax = clock.ElapsedMilliseconds - read;
                    if (invalid.Count > 0)
                    {
                        long at = clock.ElapsedMilliseconds; scope.Dispose(); scope = null; restore += clock.ElapsedMilliseconds - at;
                        if (confirm == null || !confirm(new BatchSyntaxReport(invalid))) { result.Cancelled = true; result.Details = "已取消，本批未写入原区域。"; return result; }
                        foreach (var change in invalid)
                        {
                            if (!candidates.Contains(change)) throw new InvalidOperationException("公式预检返回了无效的目标。");
                            string expression = change.After.Formula; change.After.IsFormula = false; change.After.Formula = null; change.After.Value = expression; change.After.NumberFormat = "@"; change.ConfirmedText = true;
                        }
                        scope = new ApplicationState(host); Check(host, originals, true);
                    }
                }
                var blocks = BatchEditBlock.Create(changes, false);
                foreach (var block in blocks)
                {
                    Check(host, BatchEditBlock.Create(block.Changes, true), false);
                    attempted.AddRange(block.Changes); long at = clock.ElapsedMilliseconds; Write(host, block); write += clock.ElapsedMilliseconds - at;
                    at = clock.ElapsedMilliseconds; Check(host, new[] { block }, false); verify += clock.ElapsedMilliseconds - at;
                }
                var formulas = changes.Where(c => c.After.IsFormula).Select(c => c.Cell).ToList();
                if (formulas.Count > 0) { host.RecheckScope(); host.Calculate(formulas); Check(host, blocks, false); }
                result.Changed = changes.Count; result.FormulasKept = formulas.Count; result.FormulasAsText = changes.Count(c => c.ConfirmedText); result.PlainResults = changes.Count - result.FormulasKept - result.FormulasAsText;
                result.Details = "已处理 " + result.Changed + " 个单元格；普通结果 " + result.PlainResults + " 个，保留公式 " + result.FormulasKept + " 个，确认转文本 " + result.FormulasAsText + " 个。";
            }
            catch (Exception error)
            {
                result.Failed = true; result.Changed = result.FormulasKept = result.FormulasAsText = result.PlainResults = 0;
                result.Details = "转换未完成：" + error.Message; UiDiagnostics.Record("batch-text-edit", error);
                var failures = new List<string>();
                foreach (var block in BatchEditBlock.Create(attempted, true).Reverse())
                {
                    string writeError = "";
                    try { Write(host, block); } catch (Exception rollback) { writeError = rollback.Message; }
                    try { Check(host, new[] { block }, false); }
                    catch (Exception rollback) { failures.Add(block.Address + "：" + writeError + "；" + rollback.Message + "\r\n" + string.Join("\r\n", block.Changes.Select(c => c.LocalAddress + " " + c.Before.Describe()).ToArray())); }
                }
                result.RollbackIncomplete = failures.Count > 0;
                if (failures.Count == 0) result.Details += attempted.Count == 0 ? "\r\n本批未写入任何单元格。" : "\r\n本批已尝试的内容及格式均已还原并核对。";
                else
                {
                    result.Details += "\r\n以下目标需手工恢复：\r\n" + string.Join("\r\n", failures.ToArray());
                    try { Directory.CreateDirectory(recoveryDirectory); result.RecoveryLog = Path.Combine(recoveryDirectory, "ConversionRecovery-" + Guid.NewGuid().ToString("N") + ".txt"); File.WriteAllText(result.RecoveryLog, result.Details, Encoding.UTF8); result.Details += "\r\n恢复日志：" + result.RecoveryLog; }
                    catch (Exception logError) { result.Details += "\r\n恢复日志写入失败：" + logError.Message; }
                }
            }
            finally
            {
                try { if (scope != null) { long at = clock.ElapsedMilliseconds; scope.Dispose(); restore += clock.ElapsedMilliseconds - at; } } catch (Exception error) { result.Failed = true; result.Details += "\r\n" + error.Message; }
                try { host.Dispose(); } catch (Exception error) { result.Failed = true; result.Details += "\r\n临时引用释放失败：" + error.Message; }
                result.Timings = string.Format(CultureInfo.InvariantCulture, "read={0}ms; syntax={1}ms; write={2}ms; verify={3}ms; restore_calculation={4}ms; total={5}ms", read, syntax, write, verify, restore, clock.ElapsedMilliseconds);
            }
            return result;
        }
    }
}
