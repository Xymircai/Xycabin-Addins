using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif

// Self-owned files only; no COM, preferences or user documents.
internal static class PdfSelectionProbe
{
    private static int checks;
    private static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    private static int Main()
    {
        string root = Path.Combine(Path.GetTempPath(), "PdfSelectionProbe-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try
        {
            string a = Path.Combine(root, "a", "项目"), b = Path.Combine(root, "b", "项目");
            Directory.CreateDirectory(Path.Combine(a, "部门")); Directory.CreateDirectory(b);
            string first = Path.Combine(a, "报告.DOCX"), nested = Path.Combine(a, "部门", "报告.xlsx"), other = Path.Combine(b, "报告.docx");
            File.WriteAllText(first, "source"); File.WriteAllText(nested, "source"); File.WriteAllText(other, "source");
            File.WriteAllText(Path.Combine(a, "readme.xyz"), "untouched");
            var model = new PdfFileSelection(); var warnings = new List<string>();
            model.AddInput(a); model.AddInput(b); model.AddInput(first);
            model.Commit(model.Scan(false, () => false, warnings));
            Check(model.Sources.Count == 3, "First-level scan/dedup failed");
            Check(model.Sources.Select(s => s.Folder).Distinct().Count() == 2, "Same-name roots must have distinct aliases");
            Check(model.Tasks(new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".docx" }, true).Length == 2, "Case-insensitive suffix selection failed");
            Check(model.Tasks(new HashSet<string>(), true).Length == 0 && model.Sources.Count == 3, "Filtering changed sources");
            Check(model.Tasks(new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".docx" }, true).Select(s => s.Path).SequenceEqual(new[] { first, other }), "Unhide order changed");
            model.Commit(model.Scan(true, () => false, warnings));
            Check(model.Sources.Count == 4, "Recursive scan did not discover nested file");
            Check(model.Sources.Single(s => s.Path == nested).Folder == Path.Combine("项目", "部门"), "Relative output hierarchy wrong");
            try { model.Scan(true, () => true, warnings); throw new Exception("Expected cancellation"); } catch (OperationCanceledException) { }
            Check(model.Sources.Count == 4, "Cancelled scan changed source list");
            model.Remove(model.Sources.Single(s => s.Path == first));
            model.Commit(model.Scan(true, () => false, warnings));
            Check(!model.Sources.Any(s => s.Path == first), "Removed file reappeared on rescan");
            model.AddInput(first); model.Commit(model.Scan(true, () => false, warnings));
            Check(model.Sources.Any(s => s.Path == first), "Explicit re-add failed");
            string output = Path.Combine(a, "生成"); Directory.CreateDirectory(output); File.WriteAllText(Path.Combine(output, "new.pdf"), "generated");
            model.OutputFolders.Add(output); model.Commit(model.Scan(true, () => false, warnings));
            Check(!model.Sources.Any(s => PdfFileSelection.IsWithin(s.Path, output)), "Output tree rescanned");
            model.Generated.Add(first); model.Commit(model.Scan(true, () => false, warnings));
            Check(!model.Sources.Any(s => s.Path == first), "Generated result rescanned");
            Check(model.Tasks(new HashSet<string> { ".docx" }, false).Length == 0, "PDF edit accepted Office input");
            model.Commit(model.Scan(false, () => false, warnings));
            Check(!model.Sources.Any(s => s.Path == nested), "Disabling recursion retained nested file");
            model.OutputFolders.Add(root);
            model.Commit(model.Scan(false, () => false, warnings));
            Check(model.Sources.Any(s => s.Path == other), "An ancestor output directory excluded the source tree");
            model.AddInput(Path.Combine(root, "missing")); model.Scan(false, () => false, warnings);
            Check(warnings.Any(s => s.Contains("missing")), "Missing directory was not reported");
            Check(File.ReadAllText(first) == "source" && File.ReadAllText(Path.Combine(a, "readme.xyz")) == "untouched", "Source contents changed");
            model.Clear(); Check(model.Sources.Count == 0 && model.Inputs.Count == 0, "Clear retained sources");
            Directory.Delete(root, true);
            Console.WriteLine("PASS: " + checks + " PDF selection checks."); return 0;
        }
        catch (Exception error) { Console.WriteLine(error + "\nRetained fixture: " + root); return 1; }
    }
}
