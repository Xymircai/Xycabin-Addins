using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Excel_Iterate;

// Public managed objects exercise the same late-bound calls without starting a host.
public sealed class CompanyPickRange
{
    public object Parent { get; set; } public int Row { get; set; } public int Column { get; set; }
    public bool MergeCells { get; set; }
    public CompanyFakeCount Areas { get { return new CompanyFakeCount { Count = 1 }; } }
    public CompanyFakeCount Rows { get { return new CompanyFakeCount { Count = 1 }; } }
    public CompanyFakeCount Columns { get { return new CompanyFakeCount { Count = 1 }; } }
}
public sealed class CompanyFakeCell
{
    public object Value2 { get; set; }
    public object NumberFormat { get; set; }
    public bool MergeCells { get; set; }
    public bool HasArray { get; set; }
    public bool HasFormula { get { return false; } }
}
public sealed class CompanyFakeCells
{
    private readonly Dictionary<string, CompanyFakeCell> data = new Dictionary<string, CompanyFakeCell>();
    public CompanyFakeCell Item(int row, int column)
    {
        string key = row + ":" + column;
        if (!data.ContainsKey(key)) data[key] = new CompanyFakeCell { NumberFormat = "0.00" };
        return data[key];
    }
}
public sealed class CompanyFakeCount { public int Count { get; set; } }
public sealed class CompanyFakeUsed
{
    public int Column { get { return 1; } }
    public CompanyFakeCount Columns { get { return new CompanyFakeCount { Count = 3 }; } }
}
public sealed class CompanyFakeSheet
{
    public object Parent { get; set; }
    public string Name { get; set; }
    public bool ProtectContents { get { return false; } }
    public CompanyFakeCount Rows { get { return new CompanyFakeCount { Count = 65536 }; } }
    public CompanyFakeCount Columns { get { return new CompanyFakeCount { Count = 256 }; } }
    public void Activate() { }
    public CompanyFakeUsed UsedRange { get { return new CompanyFakeUsed(); } }
    public CompanyFakeCells Cells { get; private set; }
    public CompanyFakeSheet() { Cells = new CompanyFakeCells(); }
}
public sealed class CompanyFakeSheets
{
    public readonly CompanyFakeSheet Source = new CompanyFakeSheet { Name = "source" }, Lead = new CompanyFakeSheet { Name = "lead" };
    public int Count { get { return 2; } }
    public CompanyFakeSheet Item(object key) { return Convert.ToString(key) == "1" || Convert.ToString(key) == "source" ? Source : Lead; }
}
public sealed class CompanyFakeWindow { public bool Visible { get; set; } public CompanyFakeWindow() { Visible = true; } }
public sealed class CompanyFakeWindows
{
    public readonly CompanyFakeWindow Window = new CompanyFakeWindow();
    public int Count { get { return 1; } }
    public CompanyFakeWindow Item(int index) { return Window; }
}
public sealed class CompanyFakeBook
{
    public readonly CompanyFakeWindows Windows = new CompanyFakeWindows();
    public CompanyFakeSheets Worksheets { get; private set; }
    public bool ReadOnly { get; set; }
    public bool Closed, FailSave;
    public int Saved;
    public CompanyFakeBook() { Worksheets = new CompanyFakeSheets(); Worksheets.Source.Parent = Worksheets.Lead.Parent = this; }
    public void Save() { if (FailSave) throw new IOException("injected save failure"); Saved++; }
    public void Close(bool save) { Closed = true; }
    public void Activate() { if (Closed) throw new InvalidOperationException("cannot activate closed book"); }
}
public sealed class CompanyFakeBooks
{
    public CompanyFakeBook LastSelection;
    public readonly CompanyFakeBook Template = new CompanyFakeBook();
    public readonly Dictionary<string, CompanyFakeBook> Outputs = new Dictionary<string, CompanyFakeBook>();
    public bool FailSave;
    public CompanyFakeBook Open(string path, int links, bool readOnly)
    {
        if (Path.GetFileName(path).StartsWith(".company-select-", StringComparison.Ordinal))
        {
            LastSelection = new CompanyFakeBook { ReadOnly = readOnly };
            return LastSelection;
        }
        if (!Path.GetFileName(path).StartsWith(".company-", StringComparison.Ordinal)) return Template;
        if (!Outputs.ContainsKey(path)) Outputs[path] = new CompanyFakeBook();
        var book = Outputs[path]; book.ReadOnly = readOnly; book.Closed = false; book.FailSave = FailSave;
        book.Worksheets.Lead.Cells.Item(1, 2).Value2 = "code"; book.Worksheets.Lead.Cells.Item(1, 3).Value2 = "full";
        return book;
    }
}
public sealed class CompanyFakeApp
{
    public CompanyFakeBooks Workbooks { get; private set; }
    public object ActiveWorkbook { get; set; }
    public object ActiveSheet { get { return null; } }
    public bool DisplayAlerts { get; set; }
    public bool EnableEvents { get; set; }
    public bool ScreenUpdating { get; set; }
    public bool AskToUpdateLinks { get; set; }
    public int AutomationSecurity { get; set; }
    public CompanyFakeApp() { Workbooks = new CompanyFakeBooks(); DisplayAlerts = EnableEvents = ScreenUpdating = AskToUpdateLinks = true; AutomationSecurity = 2; }
}
internal static class CompanyWorkbookHostChecks
{
    internal static void Run(string folder, Action<bool, string> check, Action<Action, string> reject)
    {
        var request = new CompanyWorkbookRequest { SourceSheet = "source", TargetSheet = "lead", NameField = "name", CodeField = "code", FullNameField = "full", TargetCodeField = "code", TargetNameField = "full", Output = folder, HeaderRow = 1, FirstRow = 3, LastRow = 5, TargetHeaderRow = 1, TargetRow = 8, UpdateLead = true };
        var app = new CompanyFakeApp(); var cells = app.Workbooks.Template.Worksheets.Source.Cells;
        cells.Item(1, 1).Value2 = "name"; cells.Item(1, 2).Value2 = "code"; cells.Item(1, 3).Value2 = "full";
        for (int row = 3; row <= 5; row++) { cells.Item(row, 1).Value2 = "Fixture" + row + ".xlsx"; cells.Item(row, 2).Value2 = "=literal" + row; cells.Item(row, 3).Value2 = "Full" + row; }
        var lead = app.Workbooks.Template.Worksheets.Lead.Cells; lead.Item(1, 2).Value2 = "code"; lead.Item(1, 3).Value2 = "full";
        byte[] bytes = new byte[] { 1, 2, 3, 99, 0 }; string template = Path.Combine(folder, "template.xls"); File.WriteAllBytes(template, bytes);
        string snapshot;
        using (var host = new CompanyWorkbookHost(app, template))
        {
            snapshot = host.Snapshot;
            check(!app.Workbooks.Template.Windows.Window.Visible, "background template reader stays hidden");
            cells.Item(2, 1).Value2 = "newname"; cells.Item(2, 2).Value2 = "newcode"; cells.Item(2, 3).Value2 = "newfull";
            using (var form = new CompanyWorkbookForm())
            {
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                Func<string, object> field = delegate(string key) { return typeof(CompanyWorkbookForm).GetField(key, flags).GetValue(form); };
                typeof(CompanyWorkbookForm).GetField("host", flags).SetValue(form, host);
                try
                {
                    var sourceChoice = (System.Windows.Forms.ComboBox)field("source"); sourceChoice.Items.AddRange(new object[] { "source", "lead" }); sourceChoice.SelectedIndex = 0;
                    var targetChoice = (System.Windows.Forms.ComboBox)field("target"); targetChoice.Items.AddRange(new object[] { "source", "lead" }); targetChoice.SelectedIndex = 1;
                    ((System.Windows.Forms.NumericUpDown)field("first")).Value = 50; ((System.Windows.Forms.NumericUpDown)field("last")).Value = 80;
                    form.ApplyPicked("data", new CompanyWorkbookSelection { Sheet = "source", Row = 3, LastRow = 9, Column = 1 });
                    var built = (CompanyWorkbookRequest)typeof(CompanyWorkbookForm).GetMethod("BuildRequest", flags).Invoke(form, new object[] { false });
                    check(built.FirstRow == 3 && built.LastRow == 9, "new selected interval replaces previously typed row numbers");
                    form.ApplyPicked("code", new CompanyWorkbookSelection { Sheet = "source", Row = 2, LastRow = 2, Column = 2, Field = "newcode" });
                    built = (CompanyWorkbookRequest)typeof(CompanyWorkbookForm).GetMethod("BuildRequest", flags).Invoke(form, new object[] { false });
                    check(built.HeaderRow == 2 && built.CodeField == "newcode" && built.FirstRow == 0 && built.LastRow == 0, "new field selection replaces header and field rather than retaining old manual values");
                    form.ApplyPicked("code", new CompanyWorkbookSelection { Sheet = "lead", Row = 1, LastRow = 1, Column = 2, Field = "code" });
                    built = (CompanyWorkbookRequest)typeof(CompanyWorkbookForm).GetMethod("BuildRequest", flags).Invoke(form, new object[] { false });
                    check(built.SourceSheet == "lead" && built.HeaderRow == 1 && built.CodeField == "code", "new field can replace old sheet context inside snapshot");
                    ((System.Windows.Forms.NumericUpDown)field("leadRow")).Value = 99;
                    form.ApplyPicked("leadRow", new CompanyWorkbookSelection { Sheet = "lead", Row = 8, LastRow = 8, Column = 2 });
                    check(((System.Windows.Forms.NumericUpDown)field("leadRow")).Value == 8 && ((System.Windows.Forms.NumericUpDown)field("leadRow")).Text == "8", "selected target row replaces displayed and stored previous values");
                    var preset = form.CapturePreset(); preset.Name = "fixture";
                    ((System.Windows.Forms.NumericUpDown)field("first")).Value = 50;
                    form.ApplyPreset(preset);
                    check(((System.Windows.Forms.NumericUpDown)field("first")).Value == decimal.Parse(preset.Values["first"]), "company preset restores configured rows");
                    check(!((System.Windows.Forms.CheckBox)field("confirmed")).Checked, "company preset never restores execution confirmation");
                    preset.Values["naming"] = "missing-field";
                    reject(delegate { form.ApplyPreset(preset); }, "company preset missing header rejected before apply");
                }
                finally { typeof(CompanyWorkbookForm).GetField("host", flags).SetValue(form, null); }
            }
            var stages = new List<string>();
            var selected = host.Pick("sourceHeader", request, delegate(object b) { check(stages.Count == 3 && stages[2] == null, "preparation ends before native selection"); return new CompanyPickRange { Parent = ((CompanyFakeBook)b).Worksheets.Source, Row = 1, Column = 1 }; }, stages.Add);
            check(selected.Sheet == "source" && selected.Row == 1 && app.Workbooks.LastSelection.ReadOnly && app.Workbooks.LastSelection.Closed && app.Workbooks.LastSelection.Saved == 0, "snapshot picker returns header coordinates and closes only its readonly copy without saving");
            selected = host.Pick("code", request, delegate(object b) { return new CompanyPickRange { Parent = ((CompanyFakeBook)b).Worksheets.Source, Row = 1, Column = 2 }; });
            check(selected.Field == "code" && !app.Workbooks.Template.Closed, "field picker resolves column from immutable reader headers");
            check(host.Pick("data", request, delegate { return false; }) == null && app.Workbooks.LastSelection.Closed, "native cancellation closes selection copy and preserves reader");
            var otherSheet = host.Pick("code", request, delegate(object b) { return new CompanyPickRange { Parent = ((CompanyFakeBook)b).Worksheets.Lead, Row = 1, Column = 2 }; });
            check(otherSheet.Sheet == "lead" && otherSheet.Field == "code", "field picker can replace old sheet context within the same snapshot");
            reject(delegate { host.Pick("data", request, delegate(object b) { return new CompanyPickRange { Parent = ((CompanyFakeBook)b).Worksheets.Lead, Row = 3, Column = 2 }; }); }, "data still requires configured source sheet");
            reject(delegate { host.Pick("code", request, delegate { return new CompanyPickRange { Parent = app.Workbooks.Template.Worksheets.Source, Row = 1, Column = 2 }; }); }, "original or other workbook selection rejected");
            check(Directory.GetFiles(Path.GetDirectoryName(snapshot), ".company-select-*").Length == 0 && app.DisplayAlerts && app.ScreenUpdating && app.EnableEvents, "picker cancellation and invalid selections clean temporary files and restore host settings");
            var rows = host.Preview(request, delegate { }); check(rows.Count == 3 && rows[0].Row == 3, "custom start row");
            check(Path.GetFileName(rows[0].OutputPath) == "Fixture3.xls" && rows[0].Name == "Fixture3.xlsx" && app.Workbooks.Outputs.Count == 0, "real preview pipeline normalizes names without generating workbooks");
            cells.Item(1, 3).Value2 = "code"; reject(delegate { host.Headers("source", 1); }, "duplicate header"); cells.Item(1, 3).Value2 = "full";
            cells.Item(4, 1).Value2 = "Fixture3"; reject(delegate { host.Preview(request, delegate { }); }, "duplicate company"); cells.Item(4, 1).Value2 = "Fixture4";
            cells.Item(4, 2).Value2 = ""; reject(delegate { host.Preview(request, delegate { }); }, "empty code"); cells.Item(4, 2).Value2 = "=literal4";
            lead.Item(8, 2).MergeCells = true; reject(delegate { host.Preview(request, delegate { }); }, "merged target"); lead.Item(8, 2).MergeCells = false;
            foreach (var row in rows) host.Generate(request, row, CompanyWorkbookHost.UpdateLead);
            check(app.Workbooks.Outputs.Count == 3 && app.Workbooks.Outputs.Values.All(b => b.Closed && b.Saved == 1), "owned books saved closed");
            check(app.Workbooks.Outputs.Values.Select(b => Convert.ToString(b.Worksheets.Lead.Cells.Item(8, 2).Value2)).Distinct().Count() == 3, "per-row association");
            check(app.Workbooks.Outputs.Values.All(b => Convert.ToString(b.Worksheets.Lead.Cells.Item(8, 2).NumberFormat) == "0.00" && b.Worksheets.Lead.Cells.Item(7, 2).Value2 == null), "format and other rows preserved");
            request.UpdateLead = false; var copy = new CompanyWorkbookRecord { Row = 3, Name = "Copy", OutputPath = Path.Combine(folder, "Copy.xls") };
            host.Generate(request, copy, null); check(File.ReadAllBytes(copy.OutputPath).SequenceEqual(bytes), "copy-only byte preservation");
            request.UpdateLead = true; app.Workbooks.FailSave = true;
            string failedPath = Path.Combine(folder, "Failure.xls");
            try { host.Generate(request, new CompanyWorkbookRecord { OutputPath = failedPath, Code = "failed", FullName = "failed" }, CompanyWorkbookHost.UpdateLead); throw new Exception("expected save failure"); } catch (IOException) { check(true, "save fault observed"); }
            check(!File.Exists(failedPath) && Directory.GetFiles(folder, ".company-*").Length == 0 && app.Workbooks.Outputs.Values.All(b => b.Closed), "failed output closed unpublished cleaned");
            check(app.DisplayAlerts && app.EnableEvents && app.ScreenUpdating && app.AskToUpdateLinks && app.AutomationSecurity == 2, "settings restored after failure");
            request.UpdateLead = false;
            try { host.Generate(request, copy, null); throw new Exception("expected collision"); } catch (IOException) { check(File.ReadAllBytes(copy.OutputPath).SequenceEqual(bytes), "publish race refuses overwrite"); }
            app.ActiveWorkbook = app.Workbooks.Template;
        }
        check(!File.Exists(snapshot) && app.Workbooks.Template.Closed && File.ReadAllBytes(template).SequenceEqual(bytes), "snapshot cleaned original unchanged");
    }
}
