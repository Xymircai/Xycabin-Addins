using System;
using System.Linq;
using Excel_Iterate;

internal static partial class CellConversionProbe
{
    private static void WorksheetBatchChecks()
    {
        var request = new WorksheetBatchRequest { SheetName = "Notes", Address = "$A$1:A2,A2:B2", Find = "a", Replacement = "00" };
        Check(request.Validate().SequenceEqual(new[] { "A1", "A2", "B2" }), "batch range normalizes and deduplicates overlap");
        Check(request.Replace("aAx") == "0000x", "case-insensitive literal replacement preserves zeros");
        request.MatchCase = true;
        Check(request.Replace("aAx") == "00Ax", "case-sensitive replacement");
        request.WholeCell = true; request.Replacement = "";
        Check(request.Replace("a") == "" && request.Replace("ab") == "ab", "whole-cell matching and deletion");
        var replaceNumber = new WorksheetBatchRequest { Find = "1", Replacement = "99" };
        string digits = replaceNumber.Replace(WorksheetBatchRequest.SearchText(12344.0));
        Check(digits == "992344" && WorksheetBatchRequest.ReplacementValue(12344.0, digits) is double &&
            (double)WorksheetBatchRequest.ReplacementValue(12344.0, digits) == 992344.0, "numeric token replacement produces a calculable number");
        Check((string)WorksheetBatchRequest.ReplacementValue("12344", digits) == "992344" &&
            (string)WorksheetBatchRequest.ReplacementValue(12344.0, "00123") == "00123" &&
            (string)WorksheetBatchRequest.ReplacementValue(12344.0, "1234567890123456") == "1234567890123456" &&
            (string)WorksheetBatchRequest.ReplacementValue(1.0, "=A1") == "=A1", "text type, leading zeros, precision and literal formula-like results are preserved");
        Check(WorksheetBatchRequest.SearchText(0.0125) == "0.0125" &&
            WorksheetBatchRequest.SearchText(1e20) == "100000000000000000000" &&
            WorksheetBatchRequest.SearchText(true) == null && WorksheetBatchRequest.SearchText(double.PositiveInfinity) == null &&
            WorksheetBatchRequest.ReplacementValue(1.0, "") == null, "numeric search is based on a finite decimal value and supports clearing");
        Check(new[] { "yyyy-mm-dd", "h:mm:ss", "[h]:mm", "[$-F800]" }.All(WorksheetBatchRequest.IsTemporalFormat) &&
            !new[] { "General", "0.0%", "#,##0", "0.00E+00", "0 \"hours\"" }.Any(WorksheetBatchRequest.IsTemporalFormat),
            "date and elapsed-time formats are skipped without treating numbers, scientific notation or literals as dates");
        request.Source = WorksheetBatchSource.Files;
        request.Files = new[] { "Sample.xlsx", "sample.XLSX", "~$cache.xlsx", "note.txt", "wps.et" };
        Check(request.FilePaths().Select(System.IO.Path.GetFileName).SequenceEqual(new[] { "Sample.xlsx", "wps.et" }), "file selection deduplicates and skips lock files");
        var initial = new WorksheetBatchReport(); var changed = new WorksheetBatchReport();
        initial.Rows.Add(new WorksheetBatchRow { Workbook = "a.xlsx", Changed = 1, Detail = "可处理", Fingerprint = "before" });
        changed.Rows.Add(new WorksheetBatchRow { Workbook = "a.xlsx", Changed = 1, Detail = "可处理", Fingerprint = "different" });
        Check(!initial.Matches(changed), "same changed count cannot authorize a stale batch");
        var format = new WorksheetBatchRequest { SheetName = "Notes", Address = "A1" };
        Check(!format.FormatsCells, "keep all defaults issues no formatting requests");
        format.ClearFormatting = true;
        Check(format.FormatsCells && format.Validate().Single() == "A1", "clear works without a text replacement");
        format.Horizontal = WorksheetBatchHorizontal.Right;
        bool rejected = false;
        try { format.Validate(); } catch (ArgumentException) { rejected = true; }
        Check(rejected, "clear cannot overlap explicit setting requests");
        Check(WorksheetBatchFormat.SameFontName("SimSun", "宋体") && WorksheetBatchFormat.SameFontName("楷体", "KaiTi") &&
            !WorksheetBatchFormat.SameFontName("SimSun", "NSimSun"), "font aliases match without merging distinct fonts");
    }
}
