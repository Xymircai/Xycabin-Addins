using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;

// Late-bound in-memory workbook model; no COM application is started or attached.
public sealed class BatchFixtureData
{
    public string Formula, Format = "General";
    public object Value;
    public bool Array, Spill, Locked;
    public BatchFixtureData Copy() { return (BatchFixtureData)MemberwiseClone(); }
}
public static class BatchFixtureNames
{
    public static int[] Point(string name)
    { int column = 0, i = 0; while (i < name.Length && char.IsLetter(name[i])) column = column * 26 + char.ToUpperInvariant(name[i++]) - 'A' + 1; return new[] { int.Parse(name.Substring(i), CultureInfo.InvariantCulture), column }; }
    public static string Name(int row, int column)
    { string name = ""; while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; } return name + row; }
    public static List<string> Cells(string address)
    {
        var result = new List<string>();
        foreach (string area in address.Split(',')) { var ends = area.Split(':'); var a = Point(ends[0]); var b = Point(ends[ends.Length - 1]); for (int row = a[0]; row <= b[0]; row++) for (int col = a[1]; col <= b[1]; col++) result.Add(Name(row, col)); }
        return result;
    }
}
public sealed class BatchFixtureCount { public int Count { get; set; } }
public sealed class BatchFixtureEmpty
{
    public int Count { get { return 0; } }
    public object Item(object key) { throw new InvalidOperationException("fixture collection is empty"); }
}
public sealed class BatchFixtureWindow { public bool Visible = true; public BatchFixtureCount SelectedSheets = new BatchFixtureCount { Count = 1 }; }
public sealed class BatchFixtureWindows
{
    private readonly BatchFixtureWindow window = new BatchFixtureWindow();
    public int Count { get { return 1; } }
    public object Item(int index) { return window; }
}
public sealed class BatchFixtureBook
{
    public readonly BatchFixtureApp App;
    public string FullName;
    public bool Owned, Saved;
    public BatchFixtureSheets Worksheets;
    public BatchFixtureEmpty Names = new BatchFixtureEmpty();
    public BatchFixtureWindows Windows = new BatchFixtureWindows();
    public BatchFixtureBook(BatchFixtureApp app, bool owned)
    { App = app; Owned = owned; FullName = owned ? "isolated-fixture.xlsx" : "batch-fixture.xlsx"; Worksheets = new BatchFixtureSheets(this); }
    public void Close(bool save)
    { if (!Owned || save) throw new InvalidOperationException("refusing to close or save the source fixture"); App.ClosedCopies++; App.Workbooks.Items.Remove(this); App.ActiveWorkbook = App.Source; App.ActiveSheet = App.Source.Worksheets.Items[0]; if (App.AfterCopyClosed != null) App.AfterCopyClosed(); }
}
public sealed class BatchFixtureBooks
{
    public readonly List<BatchFixtureBook> Items = new List<BatchFixtureBook>();
    public int Count { get { return Items.Count; } }
    public object Item(int index) { return Items[index - 1]; }
}
public sealed class BatchFixtureSheets
{
    private readonly BatchFixtureBook book;
    public readonly List<BatchFixtureSheet> Items = new List<BatchFixtureSheet>();
    public BatchFixtureSheets(BatchFixtureBook owner) { book = owner; }
    public int Count { get { return Items.Count - (book.Owned && book.App.IncompleteContext ? 1 : 0); } }
    public object Item(object key) { return key is string ? Items.Single(s => s.Name == (string)key) : Items[Convert.ToInt32(key) - 1]; }
    public void Copy()
    {
        book.App.Copies++; var clone = new BatchFixtureBook(book.App, true);
        foreach (var original in Items) { var sheet = new BatchFixtureSheet(clone, original.Name); foreach (var pair in original.Cells) sheet.Cells.Add(pair.Key, pair.Value.Copy()); clone.Worksheets.Items.Add(sheet); }
        book.App.ActiveWorkbook = clone; book.App.ActiveSheet = clone.Worksheets.Items[0];
        book.App.Workbooks.Items.Add(clone);
        if (book.App.FailAfterCopy) throw new COMException("fixture copy mutated then failed", unchecked((int)0x8fe3000b));
    }
}
public sealed class BatchFixtureSheet
{
    public readonly BatchFixtureBook Parent;
    public readonly Dictionary<string, BatchFixtureData> Cells = new Dictionary<string, BatchFixtureData>();
    public string Name;
    public bool ProtectContents;
    public BatchFixtureRanges Range;
    public BatchFixtureEmpty ListObjects = new BatchFixtureEmpty();
    public BatchFixtureSheet(BatchFixtureBook book, string name) { Parent = book; Name = name; Range = new BatchFixtureRanges(this); }
    public BatchFixtureData Cell(string name) { BatchFixtureData cell; if (!Cells.TryGetValue(name, out cell)) Cells.Add(name, cell = new BatchFixtureData()); return cell; }
    public void Activate() { Parent.App.ActiveWorkbook = Parent; Parent.App.ActiveSheet = this; }
}
public sealed class BatchFixtureRanges
{
    private readonly BatchFixtureSheet sheet;
    public BatchFixtureRanges(BatchFixtureSheet owner) { sheet = owner; }
    public BatchFixtureRange this[string address] { get { sheet.Parent.App.RangeCalls++; return new BatchFixtureRange(sheet, address); } }
}
public sealed class BatchFixtureAreas
{
    private readonly BatchFixtureRange range;
    public BatchFixtureAreas(BatchFixtureRange value) { range = value; }
    public int Count { get { return 1; } }
    public object Item(int index) { return range; }
}
public sealed class BatchFixtureRange
{
    public readonly BatchFixtureSheet Worksheet;
    public readonly string Address;
    public readonly List<string> Members;
    private BatchFixtureApp App { get { return Worksheet.Parent.App; } }
    public BatchFixtureRange(BatchFixtureSheet sheet, string address) { Worksheet = sheet; Address = address; Members = BatchFixtureNames.Cells(address); }
    public int Row { get { return BatchFixtureNames.Point(Members[0])[0]; } }
    public int Column { get { return BatchFixtureNames.Point(Members[0])[1]; } }
    public BatchFixtureCount Rows { get { return new BatchFixtureCount { Count = BatchFixtureNames.Point(Members[Members.Count - 1])[0] - Row + 1 }; } }
    public BatchFixtureCount Columns { get { return new BatchFixtureCount { Count = BatchFixtureNames.Point(Members[Members.Count - 1])[1] - Column + 1 }; } }
    public BatchFixtureAreas Areas { get { return new BatchFixtureAreas(this); } }
    public bool MergeCells { get { App.ReadCalls++; return false; } }
    private object Uniform(Func<BatchFixtureData, object> get)
    { App.ReadCalls++; object value = get(Worksheet.Cell(Members[0])); return Members.All(name => Equals(value, get(Worksheet.Cell(name)))) ? value : null; }
    public object HasFormula { get { return Uniform(c => c.Formula != null); } }
    public object HasArray { get { return Uniform(c => c.Array); } }
    public object HasSpill { get { return Uniform(c => c.Spill); } }
    public object Locked { get { return Uniform(c => c.Locked); } }
    public object NumberFormat
    {
        get { return Uniform(c => c.Format); }
        set { if (!Worksheet.Parent.Owned) { App.FormatWrites++; if (App.FailFormat) throw new COMException("fixture format failure", unchecked((int)0x8fe3000b)); } foreach (string name in Members) Worksheet.Cell(name).Format = (string)value; }
    }
    private object Read(Func<BatchFixtureData, object> get)
    {
        App.ReadCalls++; if (Members.Count == 1) return get(Worksheet.Cell(Members[0]));
        var array = Array.CreateInstance(typeof(object), new[] { Rows.Count, Columns.Count }, new[] { 1, 1 });
        foreach (string name in Members) { var point = BatchFixtureNames.Point(name); array.SetValue(get(Worksheet.Cell(name)), point[0] - Row + 1, point[1] - Column + 1); } return array;
    }
    private object At(object value, string name)
    { var array = value as Array; if (array == null) return value; var point = BatchFixtureNames.Point(name); return array.GetValue(point[0] - Row + array.GetLowerBound(0), point[1] - Column + array.GetLowerBound(1)); }
    public object Value2
    {
        get { return Read(c => c.Value); }
        set
        {
            if (!Worksheet.Parent.Owned) { App.ValueWrites++; if (App.Calculation != -4135) App.AutoRecalculations++; }
            foreach (string name in Members)
            {
                var cell = Worksheet.Cell(name); cell.Formula = null; object raw = At(value, name);
                string text = raw as string; cell.Value = App.EatsLeadingQuote && text != null && text.StartsWith("'", StringComparison.Ordinal) ? text.Substring(1) : raw;
                if (!Worksheet.Parent.Owned && App.FailValueWrite == App.ValueWrites) throw new COMException("fixture partial value write", unchecked((int)0x80010108));
            }
        }
    }
    public object Formula { get { return Read(c => c.Formula == null ? c.Value : (object)c.Formula); } set { WriteFormula(value); } }
    public object Formula2 { get { if (App.MissingFormula2) throw new MissingMemberException("Formula2"); return Formula; } set { if (App.MissingFormula2) throw new MissingMemberException("Formula2"); WriteFormula(value); } }
    private void WriteFormula(object value)
    {
        if (!Worksheet.Parent.Owned) { App.FormulaWrites++; if (App.Calculation != -4135) App.AutoRecalculations++; } else App.ProbeFormulaWrites++;
        foreach (string name in Members)
        {
            string expression = (string)At(value, name); var cell = Worksheet.Cell(name);
            if (expression.StartsWith("'", StringComparison.Ordinal)) { cell.Formula = null; cell.Value = expression.Substring(1); App.LastQuote = expression; continue; }
            cell.Formula = expression;
            if (App.Reject.Contains(expression) && (Worksheet.Parent.Owned || App.UnexpectedCommitRejection)) throw new COMException("fixture syntax rejection", App.RejectionCode);
            if (!expression.StartsWith("=", StringComparison.Ordinal)) { cell.Formula = null; cell.Value = expression; }
        }
    }
    public string Text { get { App.TextReads++; var cell = Worksheet.Cell(Members[0]); string text = Convert.ToString(cell.Value, CultureInfo.InvariantCulture); return cell.Value is double && cell.Format.All(c => c == '0') ? ((double)cell.Value).ToString(cell.Format, CultureInfo.InvariantCulture) : text; } }
    public void ClearContents() { foreach (string name in Members) { Worksheet.Cell(name).Formula = null; Worksheet.Cell(name).Value = null; } }
    public void Select() { Worksheet.Activate(); App.SelectionAddress = Address; }
    public void Calculate() { if (Worksheet.Parent.Owned) App.ProbeCalculations++; else App.ExplicitCalculations++; if (App.AfterCalculate != null) App.AfterCalculate(); }
}
public sealed class BatchFixtureApp
{
    public readonly BatchFixtureBook Source;
    public readonly BatchFixtureBooks Workbooks = new BatchFixtureBooks();
    public BatchFixtureBook ActiveWorkbook;
    public BatchFixtureSheet ActiveSheet;
    public BatchFixtureWindow ActiveWindow = new BatchFixtureWindow();
    public bool EnableEvents = true, ScreenUpdating = true, DisplayAlerts = true, AskToUpdateLinks = true;
    public int AutomationSecurity = 1;
    private int calculation = -4105;
    public int Calculation { get { return calculation; } set { if (calculation == -4135 && value != -4135) RestoredCalculations++; calculation = value; } }
    public string SelectionAddress = "A1", LastQuote;
    public int RangeCalls, ReadCalls, TextReads, ValueWrites, FormulaWrites, FormatWrites, Copies, ClosedCopies, ProbeCalculations, ProbeFormulaWrites, ExplicitCalculations, AutoRecalculations, RestoredCalculations;
    public bool MissingFormula2, IncompleteContext, FailFormat, UnexpectedCommitRejection, FailAfterCopy;
    public bool EatsLeadingQuote = true;
    public int FailValueWrite;
    public int RejectionCode = unchecked((int)0x8fe3000b);
    public readonly HashSet<string> Reject = new HashSet<string>(StringComparer.Ordinal);
    public Action AfterCalculate;
    public Action AfterCopyClosed;
    public object Selection { get { return ActiveSheet.Range[SelectionAddress]; } }
    public BatchFixtureApp()
    { Source = new BatchFixtureBook(this, false); Source.Worksheets.Items.Add(new BatchFixtureSheet(Source, "数据")); Source.Worksheets.Items.Add(new BatchFixtureSheet(Source, "BS")); Workbooks.Items.Add(Source); ActiveWorkbook = Source; ActiveSheet = Source.Worksheets.Items[0]; }
    public object Union(object first, object second)
    { return new BatchFixtureRange(((BatchFixtureRange)first).Worksheet, ((BatchFixtureRange)first).Address + "," + ((BatchFixtureRange)second).Address); }
}
