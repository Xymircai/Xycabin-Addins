using System;
using System.Runtime.InteropServices;
using Excel_Iterate;

internal static partial class CellConversionProbe
{
    private static ConversionRequest Edit(BatchTextEditOperation operation, BatchTextCountOrigin origin, int count = 0, string text = "X", int first = 1, int last = 1)
    { return new ConversionRequest { Kind = CellConversionKind.BatchInsertTrimDelete, TextEdit = new BatchTextEditOptions { Operation = operation, Origin = origin, Count = count, Text = text, Start = first, End = last } }; }
    private static string EditText(string input, BatchTextEditOperation operation, BatchTextCountOrigin origin, int count = 0, string text = "X", int first = 1, int last = 1)
    { return BatchTextEdit.Transform(input, Edit(operation, origin, count, text, first, last).TextEdit); }
    private static void InvalidEdit(Action action, string name)
    { bool rejected = false; try { action(); } catch (ArgumentException) { rejected = true; } Check(rejected, name); }
    private static void BatchTextEditChecks()
    {
        var removedMarker = BatchTextEdit.TransformState(State(12.0, "12", "General", "=BS!$L$16"), Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 2).TextEdit);
        Check(removedMarker != null && !removedMarker.IsFormula && (string)removedMarker.Value == "S!$L$16" && removedMarker.NumberFormat == "General", "regression: removing original equal sign produces an ordinary value with original format");
        // Five Chinese code points plus three digits; expected strings stay independent of the algorithm.
        const string input = "甲乙丙丁戊123";
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, 0), "X" + input, "insert at beginning zero");
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0), input + "X", "insert at end zero");
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, 3), "甲乙丙X丁戊123", "insert after three code points");
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 3), "甲乙丙丁戊X123", "insert before last three code points");
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, int.MaxValue), input + "X", "clamp insertion to end");
        Equal(EditText(input, BatchTextEditOperation.Insert, BatchTextCountOrigin.End, int.MaxValue), "X" + input, "clamp insertion to beginning");
        Equal(EditText(input, BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 3), "丁戊123", "delete prefix");
        Equal(EditText(input, BatchTextEditOperation.Delete, BatchTextCountOrigin.End, 3), "甲乙丙丁戊", "delete suffix");
        Equal(EditText(input, BatchTextEditOperation.Trim, BatchTextCountOrigin.Start, 3), "甲乙丙", "keep prefix");
        Equal(EditText(input, BatchTextEditOperation.Trim, BatchTextCountOrigin.End, 3), "123", "keep suffix");
        foreach (BatchTextCountOrigin origin in new[] { BatchTextCountOrigin.Start, BatchTextCountOrigin.End })
        {
            Equal(EditText(input, BatchTextEditOperation.Delete, origin, 0), input, "delete zero " + origin);
            Equal(EditText(input, BatchTextEditOperation.Trim, origin, 0), "", "trim zero " + origin);
            Equal(EditText(input, BatchTextEditOperation.Delete, origin, int.MaxValue), "", "overlong delete " + origin);
            Equal(EditText(input, BatchTextEditOperation.Trim, origin, int.MaxValue), input, "overlong trim " + origin);
        }
        Equal(EditText(input, BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", 3, 6), "甲乙23", "inclusive middle delete");
        Equal(EditText(input, BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 3, 6), "丙丁戊1", "inclusive middle trim");
        Equal(EditText(input, BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 7, int.MaxValue), "23", "middle end clamps");
        Equal(EditText(input, BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", int.MaxValue, int.MaxValue), "", "middle past end trim");
        Equal(EditText(input, BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", int.MaxValue, int.MaxValue), input, "middle past end delete");
        Equal(EditText("A🙂𠀀B", BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 2, 3), "🙂𠀀", "surrogate pairs count once");
        Equal(EditText("A🙂𠀀B", BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", 2, 3), "AB", "surrogate deletion boundaries");
        Equal(EditText("A🙂𠀀B", BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, 2), "A🙂X𠀀B", "surrogate insertion boundaries");
        Equal(EditText("= \t中\r\n!", BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 2, 6), " \t中\r\n", "whitespace and line breaks count");
        InvalidEdit(() => Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.Middle).Validate(), "insert middle rejected");
        InvalidEdit(() => Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, -1).Validate(), "negative position rejected");
        InvalidEdit(() => Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, 0, "").Validate(), "empty insertion rejected");
        InvalidEdit(() => Edit(BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 0, 1).Validate(), "middle zero rejected");
        InvalidEdit(() => Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Middle, 0, "", 6, 3).Validate(), "reversed middle rejected");
        InvalidEdit(() => Edit((BatchTextEditOperation)99, BatchTextCountOrigin.Start).Validate(), "invalid operation rejected");
        var request = Edit(BatchTextEditOperation.Trim, BatchTextCountOrigin.Middle, 0, "", 3, 5);
        var formula = new Cell("A1", State(10.0, "####", "0.00", "=BS!$G$6")); var result = Run(new Host(formula), request);
        Check(!result.Failed && !formula.Current.IsFormula && (string)formula.Current.Value == "S!$" && formula.Current.NumberFormat == "0.00" && result.PlainResults == 1 && result.FormulasAsText == 0, "formula trim removes original marker without forced text format");
        formula = new Cell("A1", State(10.0, "10", "0.00", "=BS!$G$6")); var host = new Host(formula);
        host.OnCalculate = delegate { formula.Current.Value = new ErrorWrapper(unchecked((int)0x800a07d7)); };
        request.TextEdit.Operation = BatchTextEditOperation.Delete; result = Run(host, request);
        Check(!result.Failed && formula.Current.Formula == "=BG$6" && formula.Current.NumberFormat == "0.00" && result.FormulasKept == 1 && host.Calculations == 1 && formula.Current.Value is ErrorWrapper, "valid expression survives calculation error value");
        formula = new Cell("A1", State(3.0, "3", "@", "=1+2")) { RejectFormula = "=1+" }; host = new Host(formula);
        request = Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.End, 1); result = CellConversionEngine.Run(host, request, recovery, report => true);
        Check(!result.Failed && (string)formula.Current.Value == "=1+" && !formula.Current.IsFormula && formula.Current.NumberFormat == "@" && result.FormulasAsText == 1 && host.Calculations == 0 && formula.Writes == 1, "preflight-confirmed syntax rejection writes literal expression once");
        formula = new Cell("A1", State(3.0, "3", "0.00", "=1+2"));
        formula.AfterWrite = delegate { if (formula.Current.IsFormula && formula.Current.Formula == "=1+") { formula.Current.IsFormula = false; formula.Current.Value = "=1+"; formula.Current.Formula = null; } };
        result = Run(new Host(formula), request);
        Check(result.Failed && !result.RollbackIncomplete && formula.Original.Matches(formula.Current, false), "unpredicted host literal acceptance rolls back without unapproved conversion");
        formula = new Cell("A1", State(3.0, "3", "0.00", "=1+2")) { RejectFormula = "=1+", FailTextFallback = true }; host = new Host(formula); result = CellConversionEngine.Run(host, request, recovery, report => true);
        Check(result.Failed && !result.RollbackIncomplete && formula.Original.Matches(formula.Current, false) && host.EnableEvents && host.ScreenUpdating, "failed literal retry rolls back original formula and format");
        formula = new Cell("A1", State(3.0, "3", "0.00", "=1+2")) { FailWrite = true }; result = Run(new Host(formula), request);
        Check(result.Failed && !result.RollbackIncomplete && formula.Writes == 2 && formula.Original.Matches(formula.Current, false), "ordinary write exception does not downgrade formula");
        formula = new Cell("A1", State(3.0, "3", "0.00", "=1+2")); host = new Host(formula); host.OnCalculate = delegate { throw new InvalidOperationException("calculation failed"); };
        result = Run(host, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"));
        Check(result.Failed && !result.RollbackIncomplete && formula.Original.Matches(formula.Current, false) && result.FormulasKept == 0, "calculation call failure rolls back rather than literal fallback");
        formula = new Cell("A1", State(3.0, "3", "@", "=1+2")); result = Run(new Host(formula), Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 0));
        Check(!result.Failed && formula.Writes == 0 && formula.Current.NumberFormat == "@", "unchanged formula is untouched");
        var text = new Cell("A1", State("=1+2")); result = Run(new Host(text), request);
        Check(!result.Failed && !text.Current.IsFormula && (string)text.Current.Value == "=1+" && result.FormulasAsText == 0, "equal-sign text never becomes a formula");
        var numeric = new Cell("A1", State(12.0, "0012", "0000")); var blank = new Cell("B1", State(null)); host = new Host(numeric, blank); host.Cells.Add(numeric);
        result = Run(host, Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.Start, 0, "X"));
        Check(!result.Failed && result.Changed == 1 && (string)numeric.Current.Value == "X0012" && numeric.Writes == 1 && blank.Writes == 0, "display zeroes, true blanks and overlapping cells");
        var first = new Cell("A1", State("small")); var longCell = new Cell("B1", State(new string('x', 32767)));
        result = Run(new Host(first, longCell), Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.Start));
        Check(result.Failed && first.Writes == 0 && longCell.Writes == 0 && result.Details.Contains("32767"), "overlong result rejects whole batch before writes");
        formula = new Cell("A1", State(3.0, "3", "General", "=1+2")); var locked = new Cell("B1", State("abc")) { Locked = true };
        result = Run(new Host(formula, locked), Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, "+1"));
        Check(result.Failed && formula.Writes == 0 && locked.Writes == 0, "protected member prevents all writes");
        BatchTextAdapterChecks();
    }
}
