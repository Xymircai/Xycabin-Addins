using System;
using System.Collections.Generic;
using System.IO;
using Excel_Iterate;

internal static class CompanyWorkbookProbe
{
    private static int checks;
    private static void Check(bool value, string name) { checks++; if (!value) throw new Exception("FAIL: " + name); }
    private static void Reject(Action action, string name) { try { action(); } catch (InvalidOperationException) { checks++; return; } throw new Exception("FAIL reject: " + name); }
    [STAThread] private static int Main()
    {
        string folder = Path.Combine(Path.GetTempPath(), "CompanyWorkbookProbe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            foreach (string invalid in new[] { "", " ", "A ", " A", "A.", "A/B", "A:B", "CON", "nul.txt", "COM1", "LPT²", ".." }) Reject(delegate { CompanyWorkbookRules.FileName(invalid); }, invalid);
            Check(CompanyWorkbookRules.FileName("Company（01）") == "Company（01）", "legal name");
            var request = new CompanyWorkbookRequest { SourceSheet = "source", TargetSheet = "lead", NameField = "name", CodeField = "code", FullNameField = "full", TargetCodeField = "code", TargetNameField = "full", Output = folder, HeaderRow = 1, FirstRow = 3, LastRow = 5, TargetHeaderRow = 1, TargetRow = 8, UpdateLead = true };
            request.Validate(); request.FirstRow = 0; Reject(request.Validate, "explicit source rows"); request.FirstRow = 3;
            request.TargetRow = 0; Reject(request.Validate, "explicit lead row"); request.TargetRow = 8;
            request.LastRow = 2; Reject(request.Validate, "reversed source rows"); request.LastRow = 5;
            request.Output = Path.Combine(folder, "missing"); Reject(request.Validate, "missing output"); request.Output = folder;
            var headers = new Dictionary<string, int> { { "code", 3 }, { "full", 9 } };
            Check(CompanyWorkbookRules.Column(headers, "code") == 3, "field locates physical column");
            Reject(delegate { CompanyWorkbookRules.Column(headers, "missing"); }, "missing field");
            var rows = new List<CompanyWorkbookRecord> { new CompanyWorkbookRecord { Row = 3, Name = "CompanyA", Code = "01", FullName = "FullA" }, new CompanyWorkbookRecord { Row = 5, Name = "CompanyB", Code = "02", FullName = "FullB" }, new CompanyWorkbookRecord { Row = 8, Name = "CompanyC", Code = "03", FullName = "FullC" } };
            CompanyWorkbookRules.ValidateRecords(rows, folder, ".xls");
            Check(rows.Count == 3 && rows[2].OutputPath == Path.Combine(folder, "CompanyC.xls") && rows[1].Code == "02", "row association and extension");
            rows[1].Name = "companya"; Reject(delegate { CompanyWorkbookRules.ValidateRecords(rows, folder, ".xls"); }, "case insensitive duplicate"); rows[1].Name = "CompanyB";
            File.WriteAllText(rows[0].OutputPath, "original"); Reject(delegate { CompanyWorkbookRules.ValidateRecords(rows, folder, ".xls"); }, "existing file");
            Check(File.ReadAllText(rows[0].OutputPath) == "original", "existing bytes untouched"); File.Delete(rows[0].OutputPath);
            Directory.CreateDirectory(rows[0].OutputPath); Reject(delegate { CompanyWorkbookRules.ValidateRecords(rows, folder, ".xls"); }, "directory collision"); Directory.Delete(rows[0].OutputPath);
            rows[1].Name = ""; Reject(delegate { CompanyWorkbookRules.ValidateRecords(rows, folder, ".xls"); }, "blank name"); rows[1].Name = new string('x', 256);
            Reject(delegate { CompanyWorkbookRules.ValidateRecords(rows, folder, ".xlsx"); }, "long path");
            Reject(delegate { CompanyWorkbookRules.ValidateRecords(new List<CompanyWorkbookRecord>(), folder, ".xlsx"); }, "empty batch");
            Reject(delegate { CompanyWorkbookRules.Text(new System.Runtime.InteropServices.ErrorWrapper(1)); }, "Excel error variant");
            Check(CompanyWorkbookRules.Text("！literal") == "！literal", "literal exclamation preserved");
            request.UpdateLead = false; request.TargetRow = 0; request.TargetSheet = ""; request.Validate(); checks++;
            CompanyWorkbookHostChecks.Run(folder, Check, Reject);
            using (var form = new CompanyWorkbookForm()) Check(!form.Visible, "production form constructs offline");
            Console.WriteLine("PASS: " + checks + " offline validation checks; no Excel/VSTO runtime validation.");
            Directory.Delete(folder, true); return 0;
        }
        catch (Exception error) { Console.WriteLine(error); Console.WriteLine("Retained fixture: " + folder); return 1; }
    }
}
