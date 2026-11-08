using System;
using System.Linq;
using System.Runtime.InteropServices;
using Excel_Iterate;

internal static partial class CellConversionProbe
{
    private static BatchFixtureApp BatchApp(int count, string formula, string format = "General")
    {
        var app = new BatchFixtureApp();
        for (int i = 1; i <= count; i++) { var cell = app.ActiveSheet.Cell("A" + i); cell.Formula = formula; cell.Value = 12.0; cell.Format = format; }
        return app;
    }
    private static CellConversionResult BatchRun(BatchFixtureApp app, ConversionRequest request, string address = "A1", Func<BatchSyntaxReport, bool> confirm = null)
    { var host = new ComConversionHost(app); host.SetAddress(address); return CellConversionEngine.Run(host, request, recovery, confirm); }
    private static void BatchTextAdapterChecks()
    {
        foreach (int count in new[] { 300, 3000 })
        {
            var app = BatchApp(count, "=BS!$L$16"); int prompts = 0;
            var result = BatchRun(app, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 2), "A1:A" + count, report => { prompts++; return true; });
            Check(!result.Failed && result.Changed == count && app.ActiveSheet.Cell("A1").Value as string == "S!$L$16" && app.ActiveSheet.Cell("A" + count).Format == "General", "ordinary formula-marker removal preserves value and format for " + count);
            Check(app.FormulaWrites == 0 && app.ProbeFormulaWrites == 0 && app.Copies == 0 && app.ExplicitCalculations == 0 && prompts == 0 && app.TextReads == 0, "ordinary fast path has no formula setters, sandbox, calculation, prompt or formula Text reads for " + count);
            Check(app.ValueWrites == (count == 300 ? 1 : 3) && app.AutoRecalculations == 0, "bounded 1024-cell matrix writes avoid per-cell recalc for " + count);
            Check(app.EnableEvents && app.ScreenUpdating && app.Calculation == -4105 && result.PlainResults == count && result.FormulasAsText == 0, "ordinary result summary and host state restored for " + count);
            Console.WriteLine("BATCH_TIMING count=" + count + " " + result.Timings + "; value_writes=" + app.ValueWrites + "; ranges=" + app.RangeCalls + "; reads=" + app.ReadCalls);
        }
        foreach (BatchTextEditOperation operation in new[] { BatchTextEditOperation.Delete, BatchTextEditOperation.Trim, BatchTextEditOperation.Insert })
        {
            var app = BatchApp(2, "=BS!$L$16", "0.00");
            var request = operation == BatchTextEditOperation.Delete ? Edit(operation, BatchTextCountOrigin.Middle, 0, "", 2, 3)
                : operation == BatchTextEditOperation.Trim ? Edit(operation, BatchTextCountOrigin.Start, 4) : Edit(operation, BatchTextCountOrigin.Start, 1, "!");
            string expected = BatchTextEdit.Transform("=BS!$L$16", request.TextEdit); app.Reject.Add(expected); int prompts = 0;
            var result = BatchRun(app, request, "A1:A2", report => { prompts++; Check(report.Count == 2 && app.ValueWrites == 0 && app.FormulaWrites == 0 && app.Copies == app.ClosedCopies && app.Calculation == -4105 && app.EnableEvents, "prompt sees closed sandbox and untouched source " + operation); return false; });
            Check(result.Cancelled && !result.Failed && prompts == 1 && app.Source.Worksheets.Items[0].Cell("A1").Formula == "=BS!$L$16" && app.FormatWrites == 0, "single batch cancellation is zero-write " + operation + "; prompts=" + prompts + "; " + result.Details);
            prompts = 0;
            result = BatchRun(app, request, "A1:A2", report => { prompts++; return true; });
            var cell = app.ActiveSheet.Cell("A1");
            Check(!result.Failed && prompts == 1 && cell.Formula == null && (string)cell.Value == expected && cell.Format == "@" && app.LastQuote == "'" + expected && result.FormulasAsText == 2, "0x8FE3000B syntax refusal converts only after confirmation " + operation);
            Check(app.Copies == app.ClosedCopies && app.ProbeCalculations == 0 && app.DisplayAlerts && app.AutomationSecurity == 1 && app.AskToUpdateLinks, "all sandbox and security state restored " + operation);
        }
        var valid = BatchApp(1, "=BS!$L$16", "0.00"); valid.MissingFormula2 = true;
        valid.AfterCalculate = delegate { valid.ActiveSheet.Cell("A1").Value = new ErrorWrapper(unchecked((int)0x800a07d7)); };
        var ok = BatchRun(valid, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"), "A1", report => { throw new Exception("valid error value must not prompt"); });
        Check(!ok.Failed && valid.ActiveSheet.Cell("A1").Formula == "=BS!$L$16+1" && valid.ActiveSheet.Cell("A1").Value is ErrorWrapper && valid.ExplicitCalculations == 1, "Formula2 member fallback preserves formula error value without prompt");
        var prefix = BatchApp(1, "=\"=1\"");
        ok = BatchRun(prefix, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 2));
        Check(!ok.Failed && (string)prefix.ActiveSheet.Cell("A1").Value == "=1\"" && prefix.Copies == 0 && prefix.FormulaWrites == 0, "removed original marker stays ordinary even if another equal sign becomes first");
        var mixed = BatchApp(4, "=BS!$L$16"); mixed.ActiveSheet.Cell("A2").Format = "0.00"; mixed.ActiveSheet.Cell("A3").Format = "@";
        ok = BatchRun(mixed, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 1), "A1:A2,A2:A4");
        Check(!ok.Failed && ok.Changed == 4 && mixed.ActiveSheet.Cell("A2").Format == "0.00" && mixed.ActiveSheet.Cell("A3").Format == "@", "mixed formats and overlapping areas preserve unique targets");
        var strings = new BatchFixtureApp(); strings.ActiveSheet.Cell("A1").Value = "00123"; strings.ActiveSheet.Cell("A2").Value = "'abc";
        ok = BatchRun(strings, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "X"), "A1:A2");
        Check(!ok.Failed && (string)strings.ActiveSheet.Cell("A1").Value == "00123X" && (string)strings.ActiveSheet.Cell("A2").Value == "'abcX" && strings.FormulaWrites == 0, "leading zeroes and existing apostrophes stay exact without new prefixes");
        strings = new BatchFixtureApp { EatsLeadingQuote = false }; strings.ActiveSheet.Cell("A1").Value = "'abc";
        ok = BatchRun(strings, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "X"));
        Check(!ok.Failed && (string)strings.ActiveSheet.Cell("A1").Value == "'abcX" && strings.ValueWrites == 1, "quote-preserving hosts do not receive an unnecessary escape");
        var race = BatchApp(1, "=BS!$L$16"); race.Reject.Add("=!$L$16");
        ok = BatchRun(race, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", 2, 3), "A1", report => { race.ActiveSheet.Cell("A1").Value = 777.0; return true; });
        Check(ok.Failed && race.ValueWrites == 0 && race.FormulaWrites == 0 && (double)race.ActiveSheet.Cell("A1").Value == 777.0, "edits made while confirming prevent stale commit");
        var partial = BatchApp(300, "=BS!$L$16"); partial.FailValueWrite = 1;
        ok = BatchRun(partial, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 1), "A1:A300");
        Check(ok.Failed && !ok.RollbackIncomplete && partial.ActiveSheet.Cell("A1").Formula == "=BS!$L$16" && partial.ActiveSheet.Cell("A300").Formula == "=BS!$L$16" && partial.Calculation == -4105 && partial.EnableEvents, "partial matrix write rolls back entire attempted block and restores calculation");
        var incomplete = BatchApp(1, "=BS!$L$16"); incomplete.IncompleteContext = true;
        ok = BatchRun(incomplete, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"));
        Check(ok.Failed && incomplete.Copies == incomplete.ClosedCopies && incomplete.ValueWrites == 0 && incomplete.FormulaWrites == 0, "incomplete syntax context stops without touching source or leaking workbook");
        var transport = BatchApp(1, "=BS!$L$16"); transport.Reject.Add("=!$L$16"); transport.RejectionCode = unchecked((int)0x80010108);
        ok = BatchRun(transport, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", 2, 3), "A1", report => { throw new Exception("transport failure must not prompt"); });
        Check(ok.Failed && transport.ValueWrites == 0 && transport.FormulaWrites == 0 && transport.Copies == transport.ClosedCopies, "transport errors do not become syntax warnings");
        var copyFailure = BatchApp(1, "=BS!$L$16"); copyFailure.FailAfterCopy = true;
        ok = BatchRun(copyFailure, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"));
        Check(ok.Failed && copyFailure.Copies == copyFailure.ClosedCopies && copyFailure.Workbooks.Count == 1 && copyFailure.FormulaWrites == 0, "copy that creates a workbook then throws is cleaned without closing the source");
        var surprise = BatchApp(1, "=BS!$L$16");
        surprise.AfterCopyClosed = delegate { surprise.Reject.Add("=BS!$L$16+1"); surprise.UnexpectedCommitRejection = true; };
        ok = BatchRun(surprise, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"), "A1", report => { throw new Exception("unpredicted syntax failure must not auto-confirm"); });
        Check(ok.Failed && !ok.RollbackIncomplete && surprise.ActiveSheet.Cell("A1").Formula == "=BS!$L$16" && surprise.ActiveSheet.Cell("A1").Format == "General", "syntax rejection first encountered at commit rolls back without unapproved conversion");
        var formatFailure = BatchApp(1, "=BS!$L$16"); formatFailure.FailFormat = true;
        ok = BatchRun(formatFailure, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 1));
        Check(ok.Failed && !ok.RollbackIncomplete && formatFailure.ActiveSheet.Cell("A1").Formula == "=BS!$L$16", "0x8FE3000B outside a formula setter remains a real format failure");
        var sample = new BatchEditChange { LocalAddress = "A1", Before = new ConversionCellState { IsFormula = true, Formula = "=BS!$L$16" }, After = new ConversionCellState { IsFormula = true, Formula = "=!$L$16" } };
        using (var prompt = BatchTextEditConfirmation.Create(new BatchSyntaxReport(new[] { sample })))
            Check(prompt.Positive.Text == "继续并转为文本" && prompt.Negative.Text == "取消" && prompt.CancelButton == prompt.Negative && prompt.AcceptButton == null && prompt.DismissResult == System.Windows.Forms.DialogResult.Cancel, "real confirmation control uses explicit continuation and cancel-safe default");
    }
}
