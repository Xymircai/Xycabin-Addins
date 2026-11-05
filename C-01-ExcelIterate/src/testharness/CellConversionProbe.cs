using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using Excel_Iterate;

internal static class CellConversionProbe
{
    private static int checks;
    private static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    private static void Equal(string actual, string expected, string message) { Check(actual == expected, message + ": " + actual + " != " + expected); }
    private static ConversionCellState State(object value, string display = null, string format = "General", string formula = null)
    { return new ConversionCellState { IsFormula = formula != null, Formula = formula, Value = value, Display = display ?? System.Convert.ToString(value, CultureInfo.InvariantCulture), NumberFormat = format }; }
    private static ConversionRequest Request(CellConversionKind kind) { return new ConversionRequest { Kind = kind }; }
    private sealed class Cell : IConversionCell
    {
        public string Address { get; set; }
        public ConversionCellState Original { get; set; }
        internal ConversionCellState Current;
        internal int Writes;
        internal bool FailWrite, FailRestore, Locked;
        internal Action AfterWrite;
        internal Cell(string address, ConversionCellState state) { Address = address; Original = state.Copy(); Current = state.Copy(); }
        public void CheckWritable() { if (Locked) throw new InvalidOperationException("protected"); }
        public ConversionCellState Read() { return Current.Copy(); }
        public void Write(ConversionCellState state)
        {
            Writes++;
            if (FailRestore && Writes > 1) throw new InvalidOperationException("restore denied");
            Current = state.Copy(); if (AfterWrite != null) AfterWrite();
            if (FailWrite && Writes == 1) throw new InvalidOperationException("setter mutated then threw");
        }
    }
    private sealed class Host : IConversionHost
    {
        internal readonly List<IConversionCell> Cells = new List<IConversionCell>();
        internal int Disposals, Rechecks;
        internal Action Recheck;
        internal int Calculations;
        internal Action<IList<IConversionCell>> OnCalculate;
        public bool EnableEvents { get; set; }
        public bool ScreenUpdating { get; set; }
        internal Host(params Cell[] cells) { Cells.AddRange(cells); EnableEvents = ScreenUpdating = true; }
        public IList<IConversionCell> Collect() { return Cells; }
        public void RecheckScope() { Rechecks++; if (Recheck != null) Recheck(); }
        public void Calculate(IList<IConversionCell> targets) { Calculations++; Check(!EnableEvents && !ScreenUpdating, "calculation remains inside transaction"); if (OnCalculate != null) OnCalculate(targets); }
        public void Dispose() { Disposals++; }
    }
    private static string recovery;
    private static CellConversionResult Run(Host host, ConversionRequest request) { return CellConversionEngine.Run(host, request, recovery); }
    private static void TextChecks()
    {
        string sample = "文化旅游开发有限公司1231ADSF！！！";
        Equal(CellCharacterClassifier.Filter(sample, CellCharacterKinds.Chinese, true), "文化旅游开发有限公司", "extract Chinese");
        Equal(CellCharacterClassifier.Filter(sample, CellCharacterKinds.Chinese | CellCharacterKinds.English, true), "文化旅游开发有限公司ADSF", "extract combined");
        Equal(CellCharacterClassifier.Filter(sample, CellCharacterKinds.Chinese, false), "1231ADSF！！！", "remove Chinese");
        string[] fragments = { " \u3000\t\r\n", "AaＺｂ", "中𠀀", "19０２", "！🙂é" };
        string mixed = string.Concat(fragments);
        for (int mask = 1; mask <= 31; mask++)
        {
            string keep = "", remove = "";
            for (int bit = 0; bit < fragments.Length; bit++) { if ((mask & (1 << bit)) != 0) keep += fragments[bit]; else remove += fragments[bit]; }
            Equal(CellCharacterClassifier.Filter(mixed, (CellCharacterKinds)mask, true), keep, "extract mask " + mask);
            Equal(CellCharacterClassifier.Filter(mixed, (CellCharacterKinds)mask, false), remove, "filter mask " + mask);
        }
        var text = Request(CellConversionKind.TextFilter); text.Kinds = CellCharacterKinds.Digits;
        var source = new Cell("A1", State(123.0, "00123", "00000")); var host = new Host(source); var result = Run(host, text);
        Check(!result.Failed && result.Changed == 1 && (string)source.Current.Value == "00123" && source.Current.NumberFormat == "@", "formatted numeric input stays text with zeroes");
        source = new Cell("A1", State(1.0, "=1", "General", "=TEXT(1,\"=0\")")); text.Kinds = CellCharacterKinds.All;
        result = Run(new Host(source), text); Check(!result.Failed && !source.Current.IsFormula && (string)source.Current.Value == "=1", "equal-sign result is literal");
        source = new Cell("A1", State("英文", "英文")); text.Kinds = CellCharacterKinds.English;
        result = Run(new Host(source), text); Check(!result.Failed && source.Current.Value == null && source.Current.NumberFormat == "@", "empty filter result clears contents");
        source = new Cell("A1", State(null)); result = Run(new Host(source), text); Check(result.Changed == 0 && source.Writes == 0, "empty input untouched by filter");
        source = new Cell("A1", State(100.0, "####")); result = Run(new Host(source), text); Check(result.Failed && source.Writes == 0 && result.Details.Contains("调宽"), "numeric hashes refuse before write");
        source = new Cell("A1", State("####")); text.Kinds = CellCharacterKinds.Other;
        result = Run(new Host(source), text); Check(!result.Failed && (string)source.Current.Value == "####", "literal hashes remain data");
        text.Kinds = CellCharacterKinds.None; result = Run(new Host(source), text); Check(result.Failed && source.Writes == 1, "empty flags rejected without another write");
    }
    private static void ReferenceChecks()
    {
        string[] expected = { "=$G$16+$G$12", "=G$16+G$12", "=$G16+$G12", "=G16+G12" };
        for (int i = 0; i < 4; i++) Equal(ReferenceStyleConverter.Convert("=G16+G$12", (ReferenceStyle)i), expected[i], "reference style " + i);
        Equal(ReferenceStyleConverter.Convert("=SUM(A1:$B$2,A:$C,1:$3)", ReferenceStyle.Absolute), "=SUM($A$1:$B$2,$A:$C,$1:$3)", "range endpoints and whole axes");
        Equal(ReferenceStyleConverter.Convert("='A1''表'!B2+[Book1.xlsx]Sheet1!C3+'E:\\A1\\[B2.xlsx]C3'!D4", ReferenceStyle.Absolute), "='A1''表'!$B$2+[Book1.xlsx]Sheet1!$C$3+'E:\\A1\\[B2.xlsx]C3'!$D$4", "quoted and external prefixes");
        Equal(ReferenceStyleConverter.Convert("=IF(A1=\"$B$2\",INDIRECT(\"C3\"),Table1[[#All],[A1]])+LOG10 (B2)+1.5E10+命名A1+XFE1+XFD1048576", ReferenceStyle.Absolute), "=IF($A$1=\"$B$2\",INDIRECT(\"C3\"),Table1[[#All],[A1]])+LOG10 ($B$2)+1.5E10+命名A1+XFE1+$XFD$1048576", "literals table names functions scientific numbers and boundaries");
        Equal(ReferenceStyleConverter.Convert("=A1:B2!C3+@D4+SUM(E5#)", ReferenceStyle.Absolute), "=A1:B2!$C$3+@$D$4+SUM($E$5#)", "3D prefixes and dynamic reference suffixes");
        string longFormula = "=" + string.Join("+", new string[300]).Replace("+", "A1+") + "B2";
        string converted = ReferenceStyleConverter.Convert(longFormula, ReferenceStyle.Absolute);
        Check(converted.Length > 255 && converted.StartsWith("=$A$1+") && converted.EndsWith("$B$2"), "long formula conversion");
        var cell = new Cell("A1", State(10.0, "10", "0.00", "=G16+G$12")); var request = Request(CellConversionKind.ReferenceStyle); request.Style = ReferenceStyle.Absolute;
        var result = Run(new Host(cell), request); Check(!result.Failed && cell.Current.Formula == expected[0] && cell.Current.NumberFormat == "0.00", "reference mode uses transaction");
        cell = new Cell("A1", State("=A1")); result = Run(new Host(cell), request); Check(result.Changed == 0 && cell.Writes == 0, "equals text is not a formula");
    }
    private static void ConversionChecks()
    {
        var formula = new Cell("A1", State(100.0, "100.00", "0.00", "=H23+H24+H26")); var constant = new Cell("B1", State(3.0));
        var result = Run(new Host(formula, constant), Request(CellConversionKind.FormulaAsText));
        Check(!result.Failed && (string)formula.Current.Value == "=H23+H24+H26" && !formula.Current.IsFormula && formula.Current.NumberFormat == "0.00" && constant.Writes == 0, "freeze formula text only");
        foreach (object value in new object[] { 100.0, true, "00123", new ErrorWrapper(unchecked((int)0x800a07d7)), "" })
        {
            formula = new Cell("A1", State(value, System.Convert.ToString(value, CultureInfo.InvariantCulture), "0.00", "=A2"));
            result = Run(new Host(formula), Request(CellConversionKind.ValuesOnly));
            Check(!result.Failed && !formula.Current.IsFormula && formula.Current.NumberFormat == "0.00" && (value is string && (string)value == "" ? formula.Current.Value == null : ConversionCellState.SameValue(value, formula.Current.Value)), "keep typed formula result " + value.GetType().Name);
        }
        var date = new Cell("A1", State(46331.0, "2026/11/05", "yyyy/mm/dd")); var zeros = new Cell("B1", State(7.0, "0007", "0000")); var blank = new Cell("C1", State(null));
        var boolean = new Cell("D1", State(true, "TRUE")); var error = new Cell("E1", State(new ErrorWrapper(unchecked((int)0x800a07d7)), "#DIV/0!"));
        result = Run(new Host(date, zeros, blank, boolean, error), Request(CellConversionKind.ResultAsText));
        Check(!result.Failed && result.Changed == 5 && (string)date.Current.Value == "2026/11/05" && (string)zeros.Current.Value == "0007" && blank.Current.Value == null && blank.Current.NumberFormat == "@" && (string)boolean.Current.Value == "TRUE" && (string)error.Current.Value == "#DIV/0!", "whole selection text conversion");
        var a = new Cell("A1", State(1.0, "1")); var b = new Cell("B1", State(2.0, "2", "General", "=SUM(A1)*2"));
        a.AfterWrite = delegate { b.Current.Value = 0.0; b.Current.Display = "0"; };
        result = Run(new Host(a, b), Request(CellConversionKind.ResultAsText));
        Check(!result.Failed && (string)b.Current.Value == "2", "all display results captured before dependency recalculation");
    }
    private static void TransactionChecks()
    {
        var a = new Cell("A1", State(1.0, "1", "0.00", "=B1")); var b = new Cell("B1", State(2.0)); b.FailWrite = true;
        var host = new Host(a, b); var result = Run(host, Request(CellConversionKind.ResultAsText));
        Check(result.Failed && !result.RollbackIncomplete && a.Original.Matches(a.Current, false) && b.Original.Matches(b.Current, false), "late setter rolls back formula value and format");
        Check(host.EnableEvents && host.ScreenUpdating && host.Disposals == 1, "host flags and references restored once");
        a = new Cell("A1", State(1.0)); b = new Cell("B1", State(2.0)) { Locked = true }; result = Run(new Host(a, b), Request(CellConversionKind.ResultAsText));
        Check(result.Failed && a.Writes == 0 && b.Writes == 0, "whole batch preflight blocks protected cell");
        a = new Cell("A1", State(1.0)); host = new Host(a); host.Recheck = delegate { if (host.Rechecks == 2) a.Current.NumberFormat = "0.00"; };
        result = Run(host, Request(CellConversionKind.ResultAsText)); Check(result.Failed && a.Writes == 0, "concurrent format change refuses stale snapshot");
        a = new Cell("A1", State(1.0)); host = new Host(a); host.Cells.Add(a); result = Run(host, Request(CellConversionKind.ResultAsText)); Check(result.Changed == 1 && a.Writes == 1, "overlap writes once");
        a = new Cell("A1", State(1.0)) { FailWrite = true, FailRestore = true }; result = Run(new Host(a), Request(CellConversionKind.ResultAsText));
        Check(result.Failed && result.RollbackIncomplete && File.Exists(result.RecoveryLog) && File.ReadAllText(result.RecoveryLog).Contains("A1"), "incomplete rollback records original value and format");
        Check(ConversionRangeAddress.Cells("A1:B2,B2:C2").Count == 5, "range union deduplicates");
        Equal(ConversionRangeAddress.Normalize("$B$2:a1,c3"), "A1:B2,C3", "normalize typed A1 ranges");
        Check(ConversionRangeAddress.Cells("A1:A100000").Count == 100000, "maximum supported range");
        foreach (string invalid in new[] { "", "A1:A100001", "XFE1", "A1048577", "Other!A1", "[Book.xlsx]Sheet1!A1", "A:A" })
        { bool refused = false; try { ConversionRangeAddress.Normalize(invalid); } catch (FormatException) { refused = true; } Check(refused, "reject invalid or oversized address " + invalid); }
    }
    private static void RestoreChecks()
    {
        var a = new Cell("A1", State("=B1*2", "=B1*2", "@")); var b = new Cell("B1", State("=25+25", "=25+25", "0.00"));
        var ordinary = new Cell("C1", State("'字面文本")); var live = new Cell("D1", State(3.0, "3", "General", "=1+2"));
        var host = new Host(a, b, ordinary, live);
        host.OnCalculate = delegate(IList<IConversionCell> targets)
        {
            Check(targets.Count == 2 && a.Current.IsFormula && b.Current.IsFormula, "all restored expressions written before calculating only restored targets");
            b.Current.Value = 50.0; a.Current.Value = 100.0;
        };
        var result = Run(host, Request(CellConversionKind.RestoreTextFormula));
        Check(!result.Failed && result.Changed == 2 && host.Calculations == 1 && (double)a.Current.Value == 100.0 && a.Current.NumberFormat == "General" && b.Current.NumberFormat == "0.00", "restore calculates dependencies and handles text format");
        Check(ordinary.Writes == 0 && live.Writes == 0, "restore skips ordinary text and existing formulas");
        foreach (object constant in new object[] { null, 12.0, true, "", "abc", "'=A1", " =A1" })
        { var cell = new Cell("A1", State(constant)); result = Run(new Host(cell), Request(CellConversionKind.RestoreTextFormula)); Check(!result.Failed && cell.Writes == 0, "restore leaves non-expression constant unchanged"); }
        a = new Cell("A1", State("='Sheet 1'!$A$1+1", null, "0.00")); host = new Host(a); result = Run(host, Request(CellConversionKind.RestoreTextFormula));
        Check(!result.Failed && a.Current.Formula == "='Sheet 1'!$A$1+1" && a.Current.NumberFormat == "0.00", "restore preserves quotes inside formula");
        a = new Cell("A1", State("=1+2", null, "@")); b = new Cell("B1", State("=3+4")); host = new Host(a, b);
        host.OnCalculate = delegate { throw new InvalidOperationException("calculation failed"); }; result = Run(host, Request(CellConversionKind.RestoreTextFormula));
        Check(result.Failed && !result.RollbackIncomplete && a.Original.Matches(a.Current, false) && b.Original.Matches(b.Current, false) && host.EnableEvents && host.ScreenUpdating, "calculation failure restores all formula texts and formats");
        a = new Cell("A1", State("=1+2", null, "@")); b = new Cell("B1", State("=3+4")) { FailWrite = true }; host = new Host(a, b); result = Run(host, Request(CellConversionKind.RestoreTextFormula));
        Check(result.Failed && !result.RollbackIncomplete && a.Original.Matches(a.Current, false) && b.Original.Matches(b.Current, false) && host.Calculations == 0, "restore setter mutates then throws rolls back without calculation");
        a = new Cell("A1", State("=1+2")); b = new Cell("B1", State("=3+4")) { Locked = true }; result = Run(new Host(a, b), Request(CellConversionKind.RestoreTextFormula));
        Check(result.Failed && a.Writes == 0 && b.Writes == 0, "restore protected target stops whole batch");
        a = new Cell("A1", State("=1+2")); b = new Cell("B1", State("=3+4")); host = new Host(a, b);
        host.Recheck = delegate { if (host.Rechecks == 2) b.Locked = true; }; result = Run(host, Request(CellConversionKind.RestoreTextFormula));
        Check(result.Failed && a.Writes == 0 && b.Writes == 0, "changed protection after host flag setup stops before the first write");
        a = new Cell("A1", State("=sum(B1:B2)")); a.AfterWrite = delegate { if (a.Current.IsFormula) a.Current.Formula = "=SUM(B1:B2)"; };
        result = Run(new Host(a), Request(CellConversionKind.RestoreTextFormula));
        Check(!result.Failed && a.Current.IsFormula && a.Current.Formula == "=SUM(B1:B2)", "restore accepts host function casing normalization");
        a = new Cell("A1", State("=\"CaseSensitive\"")); a.AfterWrite = delegate { if (a.Current.IsFormula) a.Current.Formula = "=\"CASESENSITIVE\""; };
        result = Run(new Host(a), Request(CellConversionKind.RestoreTextFormula));
        Check(result.Failed && !result.RollbackIncomplete && a.Original.Matches(a.Current, false), "restore does not accept changes inside literal strings");
        Equal(ConversionRangeAddress.FirstCell("G11:G14,G16:G18"), "G11", "first preview cell keeps input area order");
        Equal(ConversionRangeAddress.FirstCell("J5:J1,A1"), "J1", "first preview cell normalizes rectangle bounds");
        var rectangles = ConversionRangeAddress.Rectangles(new[] { "A1", "B1", "A2", "B2", "D5", "D6" });
        Check(rectangles.Count == 2 && rectangles[0] == "A1:B2" && rectangles[1] == "D5:D6", "calculation targets compact into rectangles");
    }
    private static void HostChecks()
    {
        var app = new ConversionProbeApplication(); var cell = app.ActiveSheet.Range["A1"]; cell.Value2 = 7.0; cell.NumberFormat = "0000";
        var result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.ResultAsText), recovery);
        Check(!result.Failed && (string)cell.Value2 == "0007" && cell.NumberFormat == "@" && !cell.HasFormula, "production COM adapter writes literal formatted result: " + result.Details + "; value=" + cell.Value2 + "; format=" + cell.NumberFormat);
        cell.NumberFormat = "0.00"; cell.Formula2 = "=H23+H24+H26"; cell.Result = 100.0;
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.FormulaAsText), recovery);
        Check(!result.Failed && (string)cell.Value2 == "=H23+H24+H26" && cell.NumberFormat == "0.00" && cell.LastFormulaWrite == "'=H23+H24+H26", "hidden quote-prefix adapter write");
        app.Calculation = -4135; app.ActiveSheet.AfterCalculate = delegate { cell.Result = 100.0; };
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.RestoreTextFormula), recovery);
        Check(!result.Failed && cell.HasFormula && (double)cell.Value2 == 100.0 && cell.NumberFormat == "0.00" && app.ActiveSheet.Calculations == 1 && app.Calculation == -4135, "production adapter freeze/restore invokes local calculation while preserving manual mode");
        cell.Value2 = "=1+"; cell.NumberFormat = "@";
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.RestoreTextFormula), recovery);
        Check(result.Failed && !result.RollbackIncomplete && !cell.HasFormula && (string)cell.Value2 == "=1+" && cell.NumberFormat == "@", "host syntax rejection rolls back original formula text");
        cell.Value2 = "=C1*2"; var second = app.ActiveSheet.Range["C1"]; second.Value2 = "=25+25";
        app.ActiveSheet.AfterCalculate = delegate { Check(cell.HasFormula && second.HasFormula, "production union calculates after every restored write"); second.Result = 50.0; cell.Result = 100.0; };
        using (var host = new ComConversionHost(app)) { host.SetAddress("A1,C1"); result = CellConversionEngine.Run(host, Request(CellConversionKind.RestoreTextFormula), recovery); }
        Check(!result.Failed && result.Changed == 2 && app.ActiveSheet.CalculatedCount == 2 && (double)cell.Value2 == 100.0 && cell.NumberFormat == "General" && app.Calculation == -4135, "production adapter calculates non-contiguous targets as one union; " + result.Details + "; targets=" + app.ActiveSheet.CalculatedCount + "; value=" + cell.Value2 + "; format=" + cell.NumberFormat);
        cell.Value2 = "=1+2"; cell.NumberFormat = "@"; app.ActiveSheet.AfterCalculate = delegate { throw new InvalidOperationException("fixture local calculation failed"); };
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.RestoreTextFormula), recovery);
        Check(result.Failed && !result.RollbackIncomplete && !cell.HasFormula && (string)cell.Value2 == "=1+2" && cell.NumberFormat == "@", "production calculation error restores formula text and number format");
        app.ActiveSheet.AfterCalculate = null; cell.HasArray = true;
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.RestoreTextFormula), recovery);
        Check(result.Failed && !cell.HasFormula && (string)cell.Value2 == "=1+2", "restore rejects array targets before writing"); cell.HasArray = false;
        cell.NumberFormat = "0.00";
        cell.Formula2 = "=$A$2"; cell.Result = 5.0; cell.MissingFormula2 = true;
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.ReferenceStyle), recovery);
        Check(!result.Failed && cell.Formula == "=A2", "Formula2 falls back only when missing"); cell.MissingFormula2 = false;
        cell.Value2 = "abc"; cell.HasArray = true;
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.ResultAsText), recovery);
        Check(result.Failed && (string)cell.Value2 == "abc" && cell.NumberFormat == "0.00", "array region blocked through adapter"); cell.HasArray = false;
        cell.HasSpill = true; result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.ResultAsText), recovery); Check(result.Failed, "spill blocked"); cell.HasSpill = false;
        app.ActiveSheet.ProtectContents = true; cell.Locked = true;
        result = CellConversionEngine.Run(new ComConversionHost(app), Request(CellConversionKind.ResultAsText), recovery); Check(result.Failed, "protected locked adapter cell blocked"); app.ActiveSheet.ProtectContents = false;
        using (var host = new ComConversionHost(app))
        {
            app.ActiveSheet.Range["G11"].Formula2 = "=H23+H24+H26";
            var preview = host.ReadPreview("G11:G14,G16:G18");
            Check(preview.CellAddress == "G11" && preview.IsFormula && preview.Formula == "=H23+H24+H26", "production preview reads actual first cell formula");
            Check(host.PickRange() == null, "native picker cancellation shape");
            app.Picked = app.ActiveSheet.Range["B1"]; Equal(host.PickRange(), "B1", "native picker accepted coordinates");
            var other = new ConversionProbeSheet(new ConversionProbeBook()); app.Picked = other.Range["A1"];
            bool refused = false; try { host.PickRange(); } catch (InvalidOperationException) { refused = true; } Check(refused, "native picker rejects other sheet identity");
            app.ActiveSheet.Parent.FullName = "renamed.xlsx"; refused = false; try { host.RecheckScope(); } catch (InvalidOperationException) { refused = true; } Check(refused, "renamed workbook invalidates original target");
        }
    }
    [STAThread] private static int Main()
    {
        recovery = Path.Combine(Path.GetTempPath(), "CellConversionProbe-" + Guid.NewGuid().ToString("N"));
        try { TextChecks(); ReferenceChecks(); ConversionChecks(); TransactionChecks(); RestoreChecks(); HostChecks(); Console.WriteLine("PASS CellConversionProbe: " + checks + " assertions; pure algorithms, transactions and production adapter; no host connection"); if (Directory.Exists(recovery)) { foreach (string file in Directory.GetFiles(recovery, "ConversionRecovery-*.txt")) File.Delete(file); Directory.Delete(recovery); } return 0; }
        catch (Exception error) { Console.WriteLine("FAIL CellConversionProbe: " + error); Console.WriteLine("Retained recovery fixtures: " + recovery); return 1; }
    }
}

// Public late-bound fixture surface. No real COM application is created or attached.
public sealed class ConversionProbeBook { public string FullName = "conversion-sample.xlsx"; }
public sealed class ConversionProbeCount { public int Count = 1; }
public sealed class ConversionProbeAreas
{
    private readonly ConversionProbeRange range;
    public ConversionProbeAreas(ConversionProbeRange value) { range = value; }
    public int Count { get { return 1; } }
    public object Item(int index) { return range; }
}
public sealed class ConversionProbeRange
{
    public ConversionProbeSheet Worksheet { get; private set; }
    public int Row = 1, Column;
    public ConversionProbeCount Rows = new ConversionProbeCount(), Columns = new ConversionProbeCount();
    public ConversionProbeAreas Areas { get { return new ConversionProbeAreas(this); } }
    public bool HasArray, HasSpill, Locked, MissingFormula2;
    public bool MergeCells { get { return false; } }
    public string NumberFormat = "General", LastFormulaWrite;
    private string formula;
    public object Result;
    public IList<ConversionProbeRange> Members;
    public bool HasFormula { get { return formula != null; } }
    public ConversionProbeRange(ConversionProbeSheet sheet, int column) { Worksheet = sheet; Column = column; }
    public object Value2 { get { return Result; } set { Result = value; formula = null; } }
    public string Formula
    {
        get { return formula; }
        set { LastFormulaWrite = value; if (value.StartsWith("'", StringComparison.Ordinal)) { formula = null; Result = value.Substring(1); } else { formula = value; if (value == "=1+") throw new InvalidOperationException("invalid expression"); } }
    }
    public string Formula2 { get { if (MissingFormula2) throw new MissingMemberException("Formula2"); return Formula; } set { if (MissingFormula2) throw new MissingMemberException("Formula2"); Formula = value; } }
    public string Text { get { return Result is double && NumberFormat == "0000" ? ((double)Result).ToString("0000", CultureInfo.InvariantCulture) : System.Convert.ToString(Result, CultureInfo.InvariantCulture); } }
    public void ClearContents() { formula = null; Result = null; }
    public void Calculate() { Worksheet.Calculations++; Worksheet.CalculatedCount = Members == null ? 1 : Members.Count; if (Worksheet.AfterCalculate != null) Worksheet.AfterCalculate(); }
}
public sealed class ConversionProbeRanges
{
    private readonly ConversionProbeSheet sheet;
    private readonly Dictionary<string, ConversionProbeRange> cells = new Dictionary<string, ConversionProbeRange>();
    public ConversionProbeRanges(ConversionProbeSheet value) { sheet = value; }
    public ConversionProbeRange this[string address]
    {
        get
        {
            if (address.Contains(":")) { var group = new ConversionProbeRange(sheet, 1) { Members = new List<ConversionProbeRange>() }; foreach (string name in ConversionRangeAddress.Cells(address)) group.Members.Add(this[name]); return group; }
            ConversionProbeRange cell; if (!cells.TryGetValue(address, out cell)) cells[address] = cell = new ConversionProbeRange(sheet, address[0] - 'A' + 1); return cell;
        }
    }
}
public sealed class ConversionProbeSheet
{
    public ConversionProbeBook Parent;
    public string Name = "数据";
    public bool ProtectContents;
    public ConversionProbeRanges Range;
    public int Calculations, CalculatedCount;
    public Action AfterCalculate;
    public ConversionProbeSheet(ConversionProbeBook book) { Parent = book; Range = new ConversionProbeRanges(this); }
}
public sealed class ConversionProbeWindow { public ConversionProbeCount SelectedSheets = new ConversionProbeCount(); }
public sealed class ConversionProbeApplication
{
    public ConversionProbeSheet ActiveSheet = new ConversionProbeSheet(new ConversionProbeBook());
    public ConversionProbeWindow ActiveWindow = new ConversionProbeWindow();
    public object Selection { get { return ActiveSheet.Range["A1"]; } }
    public bool EnableEvents = true, ScreenUpdating = true;
    public int Calculation;
    public object Picked = false;
    public object InputBox(object prompt, object title, object initial, object left, object top, object help, object context, object type) { return Picked; }
    public object Union(object first, object second) { var a = (ConversionProbeRange)first; var b = (ConversionProbeRange)second; var result = new ConversionProbeRange(ActiveSheet, 1) { Members = new List<ConversionProbeRange>() }; if (a.Members == null) result.Members.Add(a); else foreach (var item in a.Members) result.Members.Add(item); if (b.Members == null) result.Members.Add(b); else foreach (var item in b.Members) result.Members.Add(item); return result; }
}
