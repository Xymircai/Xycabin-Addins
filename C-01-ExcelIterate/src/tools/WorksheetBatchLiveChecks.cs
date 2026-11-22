using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Excel_Iterate;

internal static class WorksheetBatchLiveChecks
{
    private static void Check(bool condition, string message)
    { if (!condition) throw new InvalidOperationException(message); Console.WriteLine("PASS " + message); }

    internal static int Run()
    {
        using (var session = new ProbeSession("worksheet-batch-live"))
        {
            try
            {
                if (Process.GetProcessesByName("et").Length != 0)
                { Console.WriteLine("An existing WPS table host is open. Isolated multi-workbook probe skipped."); return session.Finish(2); }
                string[] ids = { "KET.Application", "et.Application" };
                object app = session.Take(ProbeSession.Connect(ids, false, id => null, id =>
                {
                    Type type = Type.GetTypeFromProgID(id);
                    return type == null ? null : Activator.CreateInstance(type);
                }, Console.WriteLine));
                return RunChecks(session, app, false);
            }
            catch (Exception error) { Console.WriteLine(error); return session.Finish(1); }
        }
    }

    internal static int RunPrivate()
    {
        int result = 1;
        using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(90)))
        {
            try
            {
                foreach (var progress in MigrationBackgroundSession.Run(app => CheckPrivateInstance(app, code => result = code), timeout.Token))
                    Thread.Sleep(25);
                return result;
            }
            catch (MigrationBackgroundUnavailableException error) { Console.WriteLine("Private WPS unavailable: " + error.Message); return 2; }
            catch (Exception error) { Console.WriteLine(error); return 1; }
        }
    }

    private static IEnumerable<MigrationProgress> CheckPrivateInstance(object app, Action<int> complete)
    {
        using (var session = new ProbeSession("worksheet-batch-private"))
            complete(RunChecks(session, app, true));
        yield break;
    }

    private static int RunChecks(ProbeSession session, object app, bool privateInstance)
    {
            try
            {
                if (app == null) return session.Finish(2);
                dynamic host = app;
                session.Override(() => (object)host.Visible, value => host.Visible = value, (object)host.Visible, "Visible");
                session.Override(() => (object)host.DisplayAlerts, value => host.DisplayAlerts = value, false, "DisplayAlerts");
                session.Override(() => (object)host.ScreenUpdating, value => host.ScreenUpdating = value, (object)host.ScreenUpdating, "ScreenUpdating");
                session.Override(() => (object)host.EnableEvents, value => host.EnableEvents = value, false, "EnableEvents");
                dynamic books = session.Take((object)host.Workbooks);
                int startingBooks = (int)books.Count;
                if (!privateInstance && startingBooks != 0) return session.Finish(2);
                // A private WPS launch can start with a blank book. Use a
                // unique sheet identity so only workbooks created below match.
                string targetSheet = "EI_Notes_" + Guid.NewGuid().ToString("N").Substring(0, 16);
                object first = AddBook(session, books, targetSheet), second = AddBook(session, books, targetSheet);
                AddBook(session, books, "Other");
                dynamic a = Seed(session, first), b = Seed(session, second);
                string recovery = session.DirectoryPath;
                var batchHost = new WorksheetBatchHost(app);
                var request = new WorksheetBatchRequest { SheetName = targetSheet, Address = "A1:B1", Find = "draft", Replacement = "posted",
                    TextStyle = WorksheetBatchTextStyle.Bold, FontColor = 0x623C12, FillColor = 0xCCF0FA, NumberFormat = "#,##0.00" };
                var preview = batchHost.Run(request, null, recovery);
                Check(preview.Rows.Count >= 3 && preview.Workbooks == 2 && preview.Cells == 4 &&
                    preview.Rows.Any(r => r.Detail.Contains("缺少工作表")) && (string)a.Value2 == "Draft", "open-workbook preview includes both matches and the missing sheet without writing: " + string.Join(" | ", preview.Rows.Select(r => r.Detail)));
                a.Value2 = "draft";
                bool stale = false;
                try { batchHost.Run(request, preview, recovery); } catch (InvalidOperationException) { stale = true; }
                Check(stale && (string)b.Value2 == "Draft", "stale preview prevents changes across all target books");
                a.Value2 = "Draft";
                using (var undo = batchHost.Run(request, batchHost.Run(request, null, recovery), recovery).Undo)
                {
                    Check((string)a.Value2 == "posted" && (string)b.Value2 == "posted" && (string)a.NumberFormat == "#,##0.00", "one apply edits both workbooks");
                    dynamic font = session.Take((object)a.Font), fill = session.Take((object)a.Interior);
                    Check((bool)font.Bold && Convert.ToInt32((object)font.Color) == 0x623C12 && Convert.ToInt32((object)fill.Color) == 0xCCF0FA,
                        "applied bold, font color and fill match the requested style");
                    a.Value2 = "kept manual edit";
                    stale = false;
                    try { undo.Execute(app, false, recovery); } catch (InvalidOperationException) { stale = true; }
                    Check(stale && (string)b.Value2 == "posted" && (string)a.Value2 == "kept manual edit", "undo leaves later edits and remaining targets intact on conflict");
                    a.Value2 = "posted";
                    undo.Execute(app, false, recovery);
                    Check((string)a.Value2 == "Draft" && (string)b.Value2 == "Draft" && !(bool)font.Bold && (int)fill.Pattern == -4142 &&
                        (int)font.ThemeColor == 2 && (string)a.NumberFormat == "0.000", "one undo restores both workbooks, theme font, no fill and numeric format");
                }
                CheckNewFormats(session, app, batchHost, first, second, targetSheet);
                CheckNumericReplacements(session, app, batchHost, first, second, targetSheet);

                string diskName = targetSheet + ".xlsx";
                string workbookPath = session.Output(diskName);
                object closed = null; bool alreadyClosed = false;
                closed = session.Own((object)books.Add(), () => { if (!alreadyClosed) ((dynamic)closed).Close(false); });
                dynamic fileRange = Seed(session, closed);
                ((dynamic)session.Take((object)fileRange.Worksheet)).Name = targetSheet;
                ((dynamic)closed).SaveAs(workbookPath, 51);
                ((dynamic)closed).Close(false); alreadyClosed = true;
                byte[] original = File.ReadAllBytes(workbookPath);
                request.Source = WorksheetBatchSource.Folder; request.Folder = session.DirectoryPath;
                FileAttributes attrs = File.GetAttributes(workbookPath);
                try
                {
                    File.SetAttributes(workbookPath, attrs | FileAttributes.ReadOnly);
                    preview = batchHost.Run(request, null, recovery);
                    Check(preview.Cells == 0 && preview.Rows.Single().Detail.Contains("只读"), "folder preview reports a read-only file before opening it");
                }
                finally { File.SetAttributes(workbookPath, attrs); }
                preview = batchHost.Run(request, null, recovery);
                Check(preview.Workbooks == 1 && (int)books.Count == startingBooks + 3, "folder preview opens and closes its own sample");
                request.Source = WorksheetBatchSource.Files; request.Files = new[] { workbookPath, workbookPath };
                using (var undo = batchHost.Run(request, preview, recovery).Undo)
                {
                    object opened = null; bool openedClosed = false;
                    opened = session.Own((object)books.Item(diskName), () => { if (!openedClosed) ((dynamic)opened).Close(false); });
                    Check((int)books.Count == startingBooks + 4 && !(bool)((dynamic)opened).Saved, "selected files are deduplicated and leave the workbook open with unsaved changes");
                    undo.Execute(app, false, recovery);
                    dynamic sh = session.Take((object)((dynamic)opened).Worksheets.Item(targetSheet));
                    dynamic restored = session.Take((object)sh.Range["A1"]);
                    Check((string)restored.Value2 == "Draft", "the workbook opened from disk participates in undo");
                    ((dynamic)opened).Close(false); openedClosed = true;
                    Check(original.SequenceEqual(File.ReadAllBytes(workbookPath)), "the batch never saves the file automatically");
                }
                CheckConversions(session, app, first, a, targetSheet);
                session.Finish(0);
            }
            catch (Exception error) { Console.WriteLine(error); session.Finish(1); }
        return session.ExitCode;
    }

    private static void CheckNumericReplacements(ProbeSession session, object app, WorksheetBatchHost host, object first, object second, string sheetName)
    {
        dynamic firstSheet = session.Take((object)((dynamic)first).Worksheets.Item(sheetName));
        dynamic secondSheet = session.Take((object)((dynamic)second).Worksheets.Item(sheetName));
        dynamic thousands = session.Take((object)firstSheet.Range["A9"]);
        dynamic textId = session.Take((object)firstSheet.Range["B9"]);
        dynamic date = session.Take((object)firstSheet.Range["C9"]);
        dynamic formula = session.Take((object)firstSheet.Range["D9"]);
        dynamic percent = session.Take((object)firstSheet.Range["E9"]);
        dynamic other = session.Take((object)secondSheet.Range["A9"]);
        thousands.NumberFormat = "#,##0"; thousands.Value2 = 12344.0;
        textId.NumberFormat = "@"; textId.Value2 = "00123";
        date.NumberFormat = "yyyy-mm-dd"; date.Value2 = 12344.0;
        formula.Formula = "=12344";
        percent.NumberFormat = "0.0%"; percent.Value2 = 0.125;
        other.Value2 = 12344.0;
        string recovery = session.DirectoryPath;
        var request = new WorksheetBatchRequest { SheetName = sheetName, Address = "A9:E9", Find = "1", Replacement = "99" };
        var preview = host.Run(request, null, recovery);
        Check(preview.Workbooks == 2 && preview.Cells == 4 && preview.Rows.Sum(row => row.SkippedDates) == 1 &&
            (double)thousands.Value2 == 12344.0, "numeric preview is read-only and reports one skipped date: " + string.Join(" | ", preview.Rows.Select(row => row.Detail)));
        const string source = "worksheet-batch";
        string unrelated = ConversionInteraction.UndoSource(CellConversionKind.TextFilter);
        Connect.RememberWorksheetEdit(source, host.Run(request, preview, recovery).Undo);
        var state = Connect.UndoState(source);
        Check(state != null && Connect.UndoState(unrelated) == null &&
            (object)thousands.Value2 is double && (double)thousands.Value2 == 992344.0 && (string)thousands.NumberFormat == "#,##0" &&
            (string)textId.Value2 == "009923" && Math.Abs((double)percent.Value2 - 0.9925) < 0.0000001 &&
            (double)other.Value2 == 992344.0 && (double)date.Value2 == 12344.0 && (string)formula.Formula == "=12344",
            "numeric apply preserves type and format across workbooks and leaves date and formula contents intact");
        Connect.RememberWorksheetEdit(source, null);
        bool wrongSourceRejected = false;
        try { Connect.UndoWorksheetEdit(unrelated, state.Id, app, recovery); } catch (InvalidOperationException) { wrongSourceRejected = true; }
        Check(wrongSourceRejected && Connect.UndoState(source).Id == state.Id && (double)thousands.Value2 == 992344.0,
            "unrelated worksheet window cannot consume the batch and no-op keeps the previous record");
        other.Value2 = 7.0;
        bool conflict = false;
        try { Connect.UndoWorksheetEdit(source, state.Id, app, recovery); } catch (InvalidOperationException) { conflict = true; }
        Check(conflict && Connect.UndoState(source).Id == state.Id && (double)thousands.Value2 == 992344.0 && (double)other.Value2 == 7.0,
            "later manual edit rejects numeric undo as a whole and keeps the attempt available");
        other.Value2 = 992344.0;
        Connect.UndoWorksheetEdit(source, state.Id, app, recovery);
        Check(Connect.UndoState(source) == null && (double)thousands.Value2 == 12344.0 && (double)other.Value2 == 12344.0 &&
            (string)textId.Value2 == "00123" && (double)percent.Value2 == 0.125, "undo restores original numeric and text types across workbooks");

        request.Address = "A9"; request.Find = "12344"; request.Replacement = "00123"; request.WholeCell = true;
        preview = host.Run(request, null, recovery);
        Check(preview.Rows.Sum(row => row.ConvertedToText) == 2, "numeric preview counts values that will become text");
        Connect.RememberWorksheetEdit(source, host.Run(request, preview, recovery).Undo);
        var next = Connect.UndoState(source);
        bool staleRejected = false;
        try { Connect.UndoWorksheetEdit(source, state.Id, app, recovery); } catch (InvalidOperationException) { staleRejected = true; }
        Check(next != null && next.Id != state.Id && staleRejected && (string)thousands.Value2 == "00123" && (string)other.Value2 == "00123" &&
            (string)thousands.NumberFormat == "#,##0", "later batch preserves leading zeros as literal text and rejects the older batch ID");
        Connect.UndoWorksheetEdit(source, next.Id, app, recovery);
        Check((object)thousands.Value2 is double && (double)thousands.Value2 == 12344.0 && (string)thousands.NumberFormat == "#,##0",
            "undo of number-to-text restores calculable value and number format");
    }

    private static void CheckNewFormats(ProbeSession session, object app, WorksheetBatchHost batchHost, object first, object second, string targetSheet)
    {
        const int doubleUnderline = -4119, singleUnderline = 2, noUnderline = -4142;
        dynamic aSheet = session.Take((object)((dynamic)first).Worksheets.Item(targetSheet));
        dynamic bSheet = session.Take((object)((dynamic)second).Worksheets.Item(targetSheet));
        dynamic a = session.Take((object)aSheet.Range["A3"]), b = session.Take((object)bSheet.Range["A3"]);
        a.Value2 = "Draft text"; b.Formula = "=2+3";
        var cells = new object[] { a, b };
        foreach (dynamic cell in cells)
        {
            dynamic font = session.Take((object)cell.Font), fill = session.Take((object)cell.Interior);
            font.Name = "Arial"; font.Size = 13.0; font.Bold = true; font.Underline = doubleUnderline;
            font.ThemeColor = 4; font.TintAndShade = 0.2;
            fill.Pattern = 1; fill.Color = 0xEECC99;
            cell.NumberFormat = "0.000"; cell.HorizontalAlignment = (int)WorksheetBatchHorizontal.Left;
            cell.VerticalAlignment = (int)WorksheetBatchVertical.Top; cell.IndentLevel = 1;
            cell.RowHeight = 28.5; cell.Locked = false;
        }
        dynamic aFont = session.Take((object)a.Font), bFont = session.Take((object)b.Font);
        dynamic border = session.Take((object)((dynamic)session.Take((object)a.Borders)).Item(9));
        border.LineStyle = 1;
        border.Weight = 2;
        dynamic conditions = session.Take((object)b.FormatConditions);
        session.Take((object)conditions.Add(1, 5, "0"));
        int conditionsBefore = (int)conditions.Count;
        double initialHeight = (double)a.RowHeight;
        var request = new WorksheetBatchRequest { SheetName = targetSheet, Address = "A3",
            TextStyle = WorksheetBatchTextStyle.Italic, FontName = "Times New Roman",
            Horizontal = WorksheetBatchHorizontal.Justify, Vertical = WorksheetBatchVertical.Center };
        using (var undo = Apply(batchHost, request, session.DirectoryPath))
        {
            Check((bool)aFont.Bold && (bool)aFont.Italic && (int)aFont.Underline == doubleUnderline,
                "italic applies without clearing existing bold and double underline");
            Check(WorksheetBatchFormat.SameFontName((string)aFont.Name, "Times New Roman") &&
                WorksheetBatchFormat.SameFontName((string)bFont.Name, "Times New Roman") &&
                (int)a.HorizontalAlignment == (int)WorksheetBatchHorizontal.Justify &&
                (int)b.VerticalAlignment == (int)WorksheetBatchVertical.Center, "font and both alignments apply across books");
            double appliedHeight = (double)a.RowHeight;
            a.RowHeight = appliedHeight + 6;
            bool staleUndo = false;
            try { undo.Execute(app, false, session.DirectoryPath); } catch (InvalidOperationException) { staleUndo = true; }
            Check(staleUndo && (bool)bFont.Italic, "later row height edit prevents a partial undo");
            a.RowHeight = appliedHeight;
            undo.Execute(app, false, session.DirectoryPath);
            Check(!(bool)aFont.Italic && (string)aFont.Name == "Arial" && (int)a.HorizontalAlignment == (int)WorksheetBatchHorizontal.Left &&
                (int)a.IndentLevel == 1 && Math.Abs((double)a.RowHeight - initialHeight) < .01, "undo restores alignment, font, indent and row height");
        }

        request = new WorksheetBatchRequest { SheetName = targetSheet, Address = "A3", TextStyle = WorksheetBatchTextStyle.Underline };
        using (var undo = Apply(batchHost, request, session.DirectoryPath))
        {
            Check((int)aFont.Underline == singleUnderline, "underline option applies single underline");
            undo.Execute(app, false, session.DirectoryPath);
            Check((int)aFont.Underline == doubleUnderline, "undo restores the exact double underline");
        }

        dynamic aStyles = session.Take((object)((dynamic)first).Styles), bStyles = session.Take((object)((dynamic)second).Styles);
        dynamic aDefault = session.Take((object)aStyles.Item("Normal")), bDefault = session.Take((object)bStyles.Item("Normal"));
        dynamic aDefaultFont = session.Take((object)aDefault.Font), bDefaultFont = session.Take((object)bDefault.Font);
        aDefaultFont.Name = "Times New Roman";
        bDefaultFont.Name = "KaiTi";
        aDefaultFont.Color = 0x123456;
        bDefaultFont.Color = 0x543210;
        request = new WorksheetBatchRequest { SheetName = targetSheet, Address = "A3", ClearFormatting = true,
            Find = "Draft", Replacement = "Posted" };
        var preview = batchHost.Run(request, null, session.DirectoryPath);
        Check(preview.Workbooks == 2, "clear preview plans both workbooks: " + string.Join(" | ", preview.Rows.Select(row => row.Detail)));
        aDefaultFont.Name = "SimSun";
        Check(WorksheetBatchFormat.SameFontName((string)aDefaultFont.Name, "SimSun") &&
            WorksheetBatchFormat.SameFontName((string)bDefaultFont.Name, "KaiTi"),
            "test workbooks have separately configurable default fonts");
        bool stale = false;
        try { batchHost.Run(request, preview, session.DirectoryPath); }
        catch (InvalidOperationException error) { if (!error.Message.Contains("请重新预览")) throw; stale = true; }
        Check(stale && (string)a.Value2 == "Draft text" && (bool)bFont.Bold, "changed workbook default invalidates a clear preview");
        int originalTheme = (int)bFont.ThemeColor;
        double originalTint = (double)bFont.TintAndShade;
        int originalIndent = (int)b.IndentLevel;
        Check(originalTheme == 4 && originalTint >= -1 && originalTint <= 1 && originalIndent == 1,
            "seeded theme, tint and indent: " + originalTheme + ", " + originalTint + ", " + originalIndent);
        using (var undo = Apply(batchHost, request, session.DirectoryPath))
        {
            Check(WorksheetBatchFormat.SameFontName((string)aFont.Name, (string)aDefaultFont.Name) &&
                WorksheetBatchFormat.SameFontName((string)bFont.Name, (string)bDefaultFont.Name) &&
                (int)aFont.Color == (int)aDefaultFont.Color && (int)bFont.Color == (int)bDefaultFont.Color,
                "clear uses each workbook's own default font and font color");
            Check((string)a.Value2 == "Posted text" && !(bool)aFont.Bold && (int)aFont.Underline == noUnderline &&
                (string)a.NumberFormat == "General" && (int)a.HorizontalAlignment == 1 && (int)b.VerticalAlignment == -4107 &&
                (int)((dynamic)session.Take((object)a.Interior)).Pattern == -4142, "clear resets common formats and combines with text replacement");
            Check((string)b.Formula == "=2+3" && (int)border.LineStyle == 1 && (int)conditions.Count == conditionsBefore &&
                (double)aFont.Size == 13.0 && !(bool)a.Locked, "clear preserves formula, border, conditional format, size and protection flags");
            undo.Execute(app, false, session.DirectoryPath);
            Check((string)a.Value2 == "Draft text" && (int)bFont.ThemeColor == originalTheme && Math.Abs((double)bFont.TintAndShade - originalTint) < .00001 &&
                (string)b.NumberFormat == "0.000" && (int)bFont.Underline == doubleUnderline && (int)b.IndentLevel == originalIndent,
                "undo restores original text, theme tint, numeric format and alignment: " + a.Value2 + ", " + bFont.ThemeColor + ", " +
                bFont.TintAndShade + ", " + b.NumberFormat + ", " + bFont.Underline + ", " + b.IndentLevel);
        }

        dynamic aFill = session.Take((object)a.Interior);
        int originalFill = (int)aFill.Color;
        aFill.Pattern = -4124;
        preview = batchHost.Run(request, null, session.DirectoryPath);
        Check(preview.Workbooks == 1 && preview.Rows.Any(row => row.Detail.Contains("图案或渐变填充")),
            "clear preview skips a patterned fill it cannot restore");
        aFill.Pattern = 1; aFill.Color = originalFill;

        aSheet.Protect();
        try
        {
            preview = batchHost.Run(request, null, session.DirectoryPath);
            Check(preview.Workbooks == 1 && preview.Rows.Any(row => row.Detail.Contains("受保护")), "clear preview reports protected sheets");
        }
        finally { aSheet.Unprotect(); }
    }

    private static WorksheetEditBatch Apply(WorksheetBatchHost host, WorksheetBatchRequest request, string recovery)
    {
        var preview = host.Run(request, null, recovery);
        Check(preview.Workbooks == 2, "both new format targets pass preview: " + string.Join(" | ", preview.Rows.Select(row => row.Detail)));
        var result = host.Run(request, preview, recovery);
        if (result.Undo == null) throw new InvalidOperationException("New format batch has no undo record.");
        return result.Undo;
    }

    private static object AddBook(ProbeSession session, dynamic books, string name)
    {
        object book = null;
        book = session.Own((object)books.Add(), () => ((dynamic)book).Close(false));
        ((dynamic)session.Take((object)((dynamic)book).Worksheets.Item(1))).Name = name;
        return book;
    }

    private static object Seed(ProbeSession session, object book)
    {
        dynamic sheet = session.Take((object)((dynamic)book).Worksheets.Item(1));
        dynamic cell = session.Take((object)sheet.Range["A1"]);
        cell.Value2 = "Draft"; cell.NumberFormat = "0.000";
        dynamic font = session.Take((object)cell.Font);
        font.ThemeColor = 2;
        ((dynamic)session.Take((object)sheet.Range["B1"])).Formula = "=1+2";
        return cell;
    }

    private static void CheckConversions(ProbeSession session, object app, object book, dynamic cell, string targetSheet)
    {
        dynamic sheet = session.Take((object)((dynamic)book).Worksheets.Item(targetSheet));
        ((dynamic)book).Activate(); sheet.Activate(); cell.Select();
        cell.NumberFormat = "General"; cell.Value2 = 123456.0;
        var request = new ConversionRequest { Kind = CellConversionKind.BatchInsertTrimDelete,
            TextEdit = new BatchTextEditOptions { Operation = BatchTextEditOperation.Insert, Origin = BatchTextCountOrigin.Start, Count = 0, Text = "000" } };
        var result = CellConversionEngine.Run(new ComConversionHost(app), request, session.DirectoryPath, recordUndo: true);
        using (var undo = result.Undo)
        {
            Check(!result.Failed && (string)cell.Value2 == "000123456" && undo != null, "batch insertion retains the three leading zeros: " + result.Details);
            undo.Execute(app, false, session.DirectoryPath);
            Check((object)cell.Value2 is double && (double)cell.Value2 == 123456.0 && (string)cell.NumberFormat == "General", "batch insertion undo restores original number and format");
        }
        cell.Formula = "'123456";
        result = CellConversionEngine.Run(new ComConversionHost(app), request, session.DirectoryPath, recordUndo: true);
        using (var undo = result.Undo)
        {
            Check(!result.Failed && (string)cell.Value2 == "000123456", "batch insertion preserves leading zeros for a text constant");
            undo.Execute(app, false, session.DirectoryPath);
            Check((string)cell.Value2 == "123456", "batch insertion undo keeps a text constant as text");
        }
        cell.Value2 = "Ab123";
        result = CellConversionEngine.Run(new ComConversionHost(app), new ConversionRequest { Kind = CellConversionKind.TextFilter, Kinds = CellCharacterKinds.Digits }, session.DirectoryPath, recordUndo: true);
        using (var undo = result.Undo)
        {
            Check(!result.Failed && (string)cell.Value2 == "123" && undo != null, "text extraction commits before undo");
            undo.Execute(app, false, session.DirectoryPath);
            Check((string)cell.Value2 == "Ab123" && (string)cell.NumberFormat == "General", "text extraction undo restores input text and numeric format");
        }
        cell.Formula = "=-2.5";
        var formula = FormulaCommandEngine.Run(new ComFormulaHost(app), FormulaRequest.Default(FormulaOperation.Abs), Console.WriteLine, session.DirectoryPath, recordUndo: true);
        using (var undo = formula.Undo)
        {
            Check(!formula.Failed && ((string)cell.Formula).Contains("ABS") && undo != null, "one-click formula has an undo record");
            undo.Execute(app, false, session.DirectoryPath);
            Check((string)cell.Formula == "=-2.5", "one-click formula undo restores the original expression");
        }
    }
}
