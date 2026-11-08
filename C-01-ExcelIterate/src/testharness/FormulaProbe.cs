using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel_Iterate;

internal static class FormulaProbe
{
    private static int passed, failed;
    private static readonly string Recovery = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "formula-probe-recovery-" + Guid.NewGuid().ToString("N"));
    private static void Check(string name, bool condition)
    { Console.WriteLine((condition ? "PASS " : "FAIL ") + name); if (condition) passed++; else failed++; }
    [STAThread]
    public static int Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        // Core mode never creates or shows a Form and never connects to WPS.
        bool core = Array.IndexOf(args, "--core") >= 0;
        try { Strings(); Parameters(); CustomTemplates(); Batches(); ComAdapter(); ParameterSelection(); IdentityRegression(); if (!core) { ParameterDialogs(); Arithmetic(); Columns(); Status(); } }
        catch (Exception ex) { Check(ex.ToString(), false); }
        string target=Path.GetFullPath(Recovery),expected=Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar);
        if(Path.GetDirectoryName(target)!=expected||!Path.GetFileName(target).StartsWith("formula-probe-recovery-"))throw new InvalidOperationException("Unsafe fixture cleanup");
        if(Directory.Exists(target))Directory.Delete(target,true);
        Console.WriteLine("FormulaProbe: " + passed + " PASS, " + failed + " FAIL");
        return failed == 0 ? 0 : 1;
    }
    private static void Equal(string source, FormulaOperation operation, string expected)
    { Check(operation + " " + source, FormulaWrapper.Transform(source, operation) == expected); }
    private static void Strings()
    {
        Equal("=A1+V2+3", FormulaOperation.IFError, "=IFERROR(A1+V2+3,\"\")");
        Equal("=A1+V2+3", FormulaOperation.Abs, "=ABS(A1+V2+3)");
        Equal("=A1+V2+3", FormulaOperation.Round, "=ROUND(A1+V2+3,2)");
        foreach (string source in new[] { "=ABS(A1+V2+3,\"\")", "=IFERROR(A1+V2+3,\"\")", "=ROUND(A1+V2+3,2)", "=sum(A1+V2+3,2)" })
            Equal(source, FormulaOperation.Unwrap, "=A1+V2+3");
        foreach (string source in new[] { "=SUM(A1,B1)+2", "=A1+V2+3", "=NOW()", "=RAND( )", "=(SUM(A1,B1))", "=SUM(A1)+SUM(B1)", "=+ABS(A1)", "=SUM(,2)", "=IFERROR(A1,2)%", "=Sheet1!A1" })
            Equal(source, FormulaOperation.Unwrap, source);
        Equal("=IFERROR(SUM(A1,B1),\"\")", FormulaOperation.Unwrap, "=SUM(A1,B1)");
        Equal("=IFERROR(\"a,b(\"\"x\"\")\",0)", FormulaOperation.Unwrap, "=\"a,b(\"\"x\"\")\"");
        Equal("=SUM('O''Brien,(x)'!A1,2)", FormulaOperation.Unwrap, "='O''Brien,(x)'!A1");
        Equal("=IFERROR(Table1[[#Data],[cost,part]],0)", FormulaOperation.Unwrap, "=Table1[[#Data],[cost,part]]");
        Equal("=SUM(Table1[O'Brien],2)", FormulaOperation.Unwrap, "=Table1[O'Brien]");
        Equal("=SUM(Table1[code']value],2)", FormulaOperation.Unwrap, "=Table1[code']value]");
        Equal("=SUM({1,2;3,4},2)", FormulaOperation.Unwrap, "={1,2;3,4}");
        Equal("=SUM({\"a,b\",\"x\"\"y\";1,2};3)", FormulaOperation.Unwrap, "={\"a,b\",\"x\"\"y\";1,2}");
        Equal("= _xlfn.FOO(  SUM(A1,B1)  ,2)  ", FormulaOperation.Unwrap, "=  SUM(A1,B1)  ");
        Equal("=求和(A1;B1)", FormulaOperation.Unwrap, "=A1");
        foreach (string bad in new[] { "=SUM(A1", "=SUM(A1]", "=SUM(\"abc,2)", "=SUM('Sheet!A1,2)", "=Table1[[a]", "=SUM({1,2),3)", "not formula" })
        {
            bool rejected = false;
            try { FormulaWrapper.Transform(bad, FormulaOperation.Abs); } catch (FormatException) { rejected = true; }
            Check("malformed rejected " + bad, rejected);
        }
        string wrapped = "=A1";
        foreach (FormulaOperation op in new[] { FormulaOperation.Abs, FormulaOperation.Round, FormulaOperation.IFError }) wrapped = FormulaWrapper.Transform(wrapped, op);
        Check("three layers", wrapped == "=IFERROR(ROUND(ABS(A1),2),\"\")");
        Equal(wrapped, FormulaOperation.Unwrap, "=ROUND(ABS(A1),2)");
        for (int i = 0; i < 3; i++) wrapped = FormulaWrapper.Transform(wrapped, FormulaOperation.Unwrap);
        Check("one layer per removal", wrapped == "=A1");
        Check("column encoding", CellAddress.Name(1048576, 16384) == "XFD1048576");
    }
    private static void Parameters()
    {
        string original = "=A1+V2+3";
        Check("custom IFERROR text", FormulaWrapper.Transform(original, FormulaRequest.IFErrorText("错误")) == "=IFERROR(A1+V2+3,\"错误\")");
        Check("custom ROUND ten", FormulaWrapper.Transform(original, FormulaRequest.Round("10")) == "=ROUND(A1+V2+3,10)");
        Check("custom text empty", FormulaWrapper.Transform(original, FormulaRequest.IFErrorText("")) == "=IFERROR(A1+V2+3,\"\")");
        Check("custom text quotes escaped", FormulaRequest.IFErrorText("错误\"详情").Argument == "\"错误\"\"详情\"");
        Check("text stays literal", FormulaRequest.IFErrorText("A1),SUM(B1),\"x").Argument == "\"A1),SUM(B1),\"\"x\"");
        foreach (string number in new[] { "0", "-1", "1.5", "1E+5" })
            Check("numeric fallback " + number, FormulaWrapper.Transform(original, FormulaRequest.IFErrorNumber(number)).StartsWith("=IFERROR(A1+V2+3,"));
        foreach (string expr in new[] { "A1", "SUM(B1:B3)", "SUM(B1,C1)", "'O''Brien,(x)'!A1", "Table1[[#Data],[price,cost]]", "{1,2;3,4}", "\"a,b\"", "A1+IFERROR(B1,0)" })
            Check("formula fallback is one argument " + expr, FormulaRequest.IFErrorFormula(expr).Argument == expr);
        Check("optional leading equals", FormulaRequest.IFErrorFormula(" =SUM(B1:B3) ").Argument == "SUM(B1:B3)");
        foreach (string digits in new[] { "10", "0", "-2", "2147483647", "-2147483648" })
            Check("integer digits " + digits, FormulaRequest.Round(digits).Argument == digits);
        Check("default wrappers still unchanged", FormulaRequest.Default(FormulaOperation.IFError).Argument == "\"\"" && FormulaRequest.Default(FormulaOperation.Round).Argument == "2");
        foreach (string bad in new[] { "", "NaN", "Infinity", "1e999", "1,2", "0);SUM(A1)", "=A1" })
        {
            bool rejected = false; try { FormulaRequest.IFErrorNumber(bad); } catch (FormatException) { rejected = true; }
            Check("invalid number rejected " + bad, rejected);
        }
        foreach (string bad in new[] { "", "=", "==A1", "A1,B1", "A1;B1", "A1),SUM(B1", "SUM(A1", "\"unfinished" })
        {
            bool rejected = false; try { FormulaRequest.IFErrorFormula(bad); } catch (FormatException) { rejected = true; }
            Check("invalid formula argument rejected " + bad, rejected);
        }
        foreach (string bad in new[] { "", "2.5", "10,A1", "SUM(A1)", "2147483648" })
        {
            bool rejected = false; try { FormulaRequest.Round(bad); } catch (FormatException) { rejected = true; }
            Check("invalid digits rejected " + bad, rejected);
        }
        string doubleWrap = FormulaWrapper.Transform(FormulaWrapper.Transform(original, FormulaRequest.IFErrorText("错误")), FormulaRequest.Round("10"));
        Check("custom wrapping still adds one layer", doubleWrap == "=ROUND(IFERROR(A1+V2+3,\"错误\"),10)");
        Check("custom unwrap preserves first argument", FormulaWrapper.Transform(doubleWrap, FormulaOperation.Unwrap) == "=IFERROR(A1+V2+3,\"错误\")");
    }
    private static void CustomTemplates()
    {
        // WPS upper-cases recognized names when it stores a formula, so the template
        // has to accept any casing. The parser only checks characters and the write-back
        // comparison ignores case, which is what formula semantics already imply.
        foreach (string name in new[] { "TEXTSPLIT", "textsplit", "TextSplit" })
        {
            bool accepted = false;
            try { accepted = FormulaRequest.Custom("=" + name + "(,\"-\",1)").Argument.Contains(FormulaRequest.Placeholder); }
            catch (FormatException) { accepted = false; }
            Check("custom template accepts " + name, accepted);
        }
        // Three empty arguments here; the slot decides which one receives the range.
        string multi = "TEXTSPLIT(,\"-\",,)";
        Check("three empty arguments counted", FormulaWrapper.CountEmptyArguments(multi) == 3);
        Check("first empty argument by default", FormulaWrapper.PrepareCustomTemplate(multi) == "TEXTSPLIT(" + FormulaRequest.Placeholder + ",\"-\",,)");
        Check("second empty argument selected", FormulaWrapper.PrepareCustomTemplate(multi, 1) == "TEXTSPLIT(,\"-\"," + FormulaRequest.Placeholder + ",)");
        Check("third empty argument selected", FormulaWrapper.PrepareCustomTemplate(multi, 2) == "TEXTSPLIT(,\"-\",," + FormulaRequest.Placeholder + ")");
        Check("slot travels on the request", FormulaRequest.Custom("=" + multi, 2).Slot == 2);
        Check("single empty argument counted", FormulaWrapper.CountEmptyArguments("ISBLANK()") == 1);
        Check("quoted comma is not a slot", FormulaWrapper.CountEmptyArguments("TEXTSPLIT(\",\",1)") == 0);
        Check("nested empty argument not counted", FormulaWrapper.CountEmptyArguments("IFERROR(TEXTSPLIT(,\"-\"),\"\")") == 0);
        Check("malformed template counts zero", FormulaWrapper.CountEmptyArguments("SUM(A1") == 0);
        bool outOfRange = false;
        try { FormulaWrapper.PrepareCustomTemplate(multi, 3); } catch (FormatException) { outOfRange = true; }
        Check("slot beyond the last empty argument rejected", outOfRange);
    }

    private static void ParameterSelection()
    {
        FakeApp app = One(); app.Sheet.Set(1, 2, "=C1", true);
        ComFormulaHost host = new ComFormulaHost(app); var snapshot = host.Collect();
        app.Selection = new FakeRange(app.Sheet, 1, 2, 1, 1);
        Check("parameter selection is cached", ReferenceEquals(snapshot, host.Collect()));
        FormulaCommandResult r = FormulaCommandEngine.Run(host, FormulaRequest.Round("10"), delegate { }, Recovery);
        Check("confirmation targets leaf-click range", !r.Failed && (string)app.Sheet.Cell(1, 1).Value == "=ROUND(B1,10)" && app.Sheet.Cell(1, 2).Writes == 0);
        app = One(); host = new ComFormulaHost(app); host.Collect(); app.Sheet.Cell(1, 1).Value = "=C1";
        r = FormulaCommandEngine.Run(host, FormulaRequest.IFErrorText("错误"), delegate { }, Recovery);
        Check("formula edited while parameter editor open rejected", r.Failed && app.Sheet.Cell(1, 1).Writes == 0 && app.EnableEvents && app.ScreenUpdating);
        app = One(); host = new ComFormulaHost(app); host.Collect(); app.Sheet.ProtectContents = true; app.Sheet.Cell(1, 1).Locked = true;
        r = FormulaCommandEngine.Run(host, FormulaRequest.Round("10"), delegate { }, Recovery);
        Check("protection re-read after parameter editor", r.Failed && r.Details.Contains("受保护") && app.Sheet.Cell(1, 1).Writes == 0);
        app = One(); host = new ComFormulaHost(app); host.Collect(); app.Sheet.Cell(1, 1).IsFormula = false;
        r = FormulaCommandEngine.Run(host, FormulaRequest.IFErrorText("错误"), delegate { }, Recovery);
        Check("same text converted to literal while editor open rejected", r.Failed && app.Sheet.Cell(1, 1).Writes == 0);
        app = One(); host = new ComFormulaHost(app); host.Collect(); host.Dispose();
        Check("cancel disposes without writes or flag changes", app.Sheet.Cell(1, 1).Writes == 0 && app.EnableEvents && app.ScreenUpdating);
        app = One(); app.Sheet.Set(2, 1, "=B2", true); app.Selection = new FakeRange(app.Sheet, 1, 1, 2, 1);
        r = FormulaCommandEngine.Run(new ComFormulaHost(app), FormulaRequest.IFErrorFormula("SUM(C1:C3)"), delegate { }, Recovery);
        Check("same custom argument applied to entire batch", !r.Failed && r.Changed == 2 && (string)app.Sheet.Cell(1, 1).Value == "=IFERROR(B1,SUM(C1:C3))" && (string)app.Sheet.Cell(2, 1).Value == "=IFERROR(B2,SUM(C1:C3))");
        Host failing = new Host(); Cell cell = new Cell("A1", "=B1") { ThrowAfterMutation = true }; failing.Cells.Add(cell);
        r = FormulaCommandEngine.Run(failing, FormulaRequest.Round("10"), delegate { }, Recovery);
        Check("custom command keeps verified rollback", r.Failed && !r.RollbackIncomplete && cell.Current == cell.Snapshot && failing.Events && failing.Screen);
    }

    private static void ParameterDialogs()
    {
        using (FormulaOptionsForm form = new FormulaOptionsForm(FormulaOperation.IFError))
        {
            Check("parameter text defaults to empty", form.ValueType.SelectedIndex == 0 && form.Apply.Enabled && form.Request == null);
            form.ValueInput.Text = "错误";
            Check("parameter preview shows requested text", form.Preview.Text == "=IFERROR(A1+V2+3,\"错误\")");
            form.ValueType.SelectedIndex = 1; form.ValueInput.Text = "0";
            Check("numeric UI preview", form.Preview.Text == "=IFERROR(A1+V2+3,0)");
            form.ValueType.SelectedIndex = 2; form.ValueInput.Text = "SUM(B1:B3)";
            Check("formula UI preview", form.Preview.Text == "=IFERROR(A1+V2+3,SUM(B1:B3))");
            form.ValueInput.Text = "A1,B1";
            Check("invalid parameter disables Apply", !form.Apply.Enabled && form.Validation.Text.Length > 0 && form.Request == null);
            Check("Enter cannot silently commit", form.AcceptButton == null);
            form.GetType().GetMethod("OnKeyDown", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(form, new object[] { new System.Windows.Forms.KeyEventArgs(System.Windows.Forms.Keys.Escape) });
            Check("parameter Esc cancels", form.DialogResult == System.Windows.Forms.DialogResult.Cancel && form.Request == null);
        }
        using (FormulaOptionsForm form = new FormulaOptionsForm(FormulaOperation.Round))
        {
            Check("ROUND UI keeps default two", form.ValueInput.Text == "2" && form.Apply.Enabled);
            form.ValueInput.Text = "2.5"; Check("ROUND rejects fractional digits in UI", !form.Apply.Enabled);
            form.ValueInput.Text = "10"; Check("ROUND preview ten", form.Preview.Text == "=ROUND(A1+V2+3,10)");
            form.GetType().GetMethod("OnFormClosing", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).Invoke(form, new object[] { new System.Windows.Forms.FormClosingEventArgs(System.Windows.Forms.CloseReason.UserClosing, false) });
            Check("parameter caption close cancels", form.DialogResult == System.Windows.Forms.DialogResult.Cancel && form.Request == null);
        }
        // Exercise the actual click handler with an isolated form; no host is present.
        using (FormulaOptionsForm form = new FormulaOptionsForm(FormulaOperation.Round))
        {
            form.ValueInput.Text = "10"; form.Show(); form.Apply.PerformClick();
            Check("explicit Apply creates validated request", form.Request != null && form.Request.Argument == "10" && form.DialogResult == System.Windows.Forms.DialogResult.OK);
        }
    }

    private sealed class Cell : IFormulaCell
    {
        public string Address { get; set; }
        public string Snapshot { get; set; }
        internal string Current, Reject;
        internal bool ThrowAfterMutation, FailRestore, BadRestore;
        internal int Writes, Reads;
        internal Action<Cell> BeforeRead;
        internal Cell(string address, string formula) { Address = address; Snapshot = Current = formula; }
        public void CheckWritable() { if (Reject != null) throw new Exception(Reject); }
        public string Read() { Reads++; if (BeforeRead != null) BeforeRead(this); return Current; }
        public void Write(string formula)
        {
            Writes++;
            if (formula == Snapshot && FailRestore) throw new Exception("恢复写入失败");
            if (formula == Snapshot && BadRestore) { Current = "=999"; return; }
            Current = formula;
            if (formula != Snapshot && ThrowAfterMutation) throw new Exception("写入后抛出");
        }
    }
    private sealed class Host : IFormulaHost
    {
        internal List<IFormulaCell> Cells = new List<IFormulaCell>();
        internal bool Events = true, Screen = true, Disposed, IgnoreDisable, ThrowDisable, ThrowRestore, ThrowDispose;
        internal string CollectError;
        public IList<IFormulaCell> Collect() { if (CollectError != null) throw new Exception(CollectError); return Cells; }
        public bool EnableEvents { get { return Events; } set { if (!value && IgnoreDisable) return; Events = value; if (!value && ThrowDisable) throw new Exception("事件 setter 抛出"); } }
        public bool ScreenUpdating { get { return Screen; } set { Screen = value; if (value && ThrowRestore) throw new Exception("刷新恢复异常"); } }
        public void Dispose() { Disposed = true; if (ThrowDispose) throw new Exception("释放异常"); }
    }
    private static FormulaCommandResult Run(Host host, FormulaOperation op)
    { return FormulaCommandEngine.Run(host, op, delegate { }, Recovery); }
    private static void Batches()
    {
        Host h = new Host(); Cell a = new Cell("A1", "=A2"), b = new Cell("C2", "=B2"); h.Cells.AddRange(new IFormulaCell[] { a, b, a });
        FormulaCommandResult r = Run(h, FormulaOperation.Abs);
        Check("overlap dedup / multi-area", !r.Failed && r.Changed == 2 && a.Writes == 1 && b.Writes == 1);
        Check("success state restored / disposed", h.Events && h.Screen && h.Disposed);
        foreach (string reason in new[] { "数组公式", "动态溢出", "受保护且锁定", "必要属性读取失败" })
        {
            h = new Host(); a = new Cell("A1", "=A2"); b = new Cell("B1", "=B2") { Reject = reason }; h.Cells.AddRange(new[] { a, b });
            r = Run(h, FormulaOperation.Abs);
            Check("preflight zero writes " + reason, r.Failed && a.Writes == 0 && b.Writes == 0 && h.Disposed && h.Events && h.Screen);
        }
        h = new Host { Events = false, Screen = false }; a = new Cell("A1", "=B1"); h.Cells.Add(a); r = Run(h, FormulaOperation.Round);
        Check("original false flags preserved", r.Changed == 1 && !h.Events && !h.Screen);
        h = new Host(); a = new Cell("A1", "=A2"); b = new Cell("B1", "=B2") { ThrowAfterMutation = true }; h.Cells.AddRange(new[] { a, b });
        r = Run(h, FormulaOperation.Abs);
        Check("late setter failure restores attempted cell", r.Failed && !r.RollbackIncomplete && a.Current == a.Snapshot && b.Current == b.Snapshot && a.Writes == 2 && b.Writes == 2);
        Check("rollback restores host state", h.Events && h.Screen && h.Disposed);
        foreach (bool badReadback in new[] { false, true })
        {
            h = new Host(); a = new Cell("A1", "=A2") { ThrowAfterMutation = true, FailRestore = !badReadback, BadRestore = badReadback }; h.Cells.Add(a);
            r = Run(h, FormulaOperation.Abs);
            Check("incomplete recovery reported " + badReadback, r.Failed && r.RollbackIncomplete && r.Details.Contains("A1") && r.Details.Contains("=A2") && WpsShared.SourcePaths.Exists(r.RecoveryLog));
            Check("saved original formula " + badReadback, WpsShared.SourcePaths.ReadAllText(r.RecoveryLog).Contains("原公式：=A2"));
        }
        h = new Host(); a = new Cell("A1", "=A2"); b = new Cell("B1", "=B2"); b.BeforeRead = delegate(Cell c) { c.Current = "=999"; }; h.Cells.AddRange(new[] { a, b }); r = Run(h, FormulaOperation.Abs);
        Check("stale snapshot before first write", r.Failed && a.Writes == 0 && b.Writes == 0 && h.Events && h.Screen);
        h = new Host(); a = new Cell("A1", "=A2"); b = new Cell("B1", "=B2"); b.BeforeRead = delegate(Cell c) { if (c.Reads == 2) c.Current = "=999"; }; h.Cells.AddRange(new[] { a, b }); r = Run(h, FormulaOperation.Abs);
        Check("stale later cell rolls back earlier", r.Failed && a.Current == a.Snapshot && a.Writes == 2 && b.Writes == 0);
        h = new Host(); h.Cells.Add(new Cell("A1", "=SUM(A2,B2)+2")); r = Run(h, FormulaOperation.Unwrap);
        Check("no-op / empty feedback", !r.Failed && r.Changed == 0 && r.Details.Contains("没有") && h.Disposed);
        foreach (bool ignore in new[] { false, true })
        {
            h = new Host { IgnoreDisable = ignore, ThrowDisable = !ignore }; a = new Cell("A1", "=A2"); h.Cells.Add(a); r = Run(h, FormulaOperation.Abs);
            Check("disable failure zero writes and restored " + ignore, r.Failed && a.Writes == 0 && h.Events && h.Screen && h.Disposed);
        }
        h = new Host { ThrowRestore = true, ThrowDispose = true }; a = new Cell("A1", "=A2"); h.Cells.Add(a); r = Run(h, FormulaOperation.Abs);
        Check("cleanup failures explicitly reported", r.Failed && r.Changed == 1 && r.Details.Contains("屏幕刷新恢复失败") && r.Details.Contains("临时引用释放失败") && h.Events && h.Disposed);
        h = new Host { CollectError = "非单元格选区" }; r = Run(h, FormulaOperation.Abs);
        Check("collection failure disposed", r.Failed && h.Disposed && h.Events && h.Screen);
    }
    private static FormulaCommandResult Execute(FakeApp app, FormulaOperation op)
    { var result = FormulaCommandEngine.Run(new ComFormulaHost(app), op, delegate { }, Recovery); if (result.Failed) Console.WriteLine("COM result: " + result.Details); return result; }
    private static void ComAdapter()
    {
        FakeApp app = new FakeApp(); app.Sheet.Set(1, 1, "=B1", true); app.Sheet.Set(2, 1, "=not-a-formula", false);
        app.Sheet.Set(1, 3, "=B2", true); app.Sheet.Set(3, 1, 42, false);
        app.Selection = new FakeRange(app.Sheet, 1, 1, 1048576, 1);
        app.Selection.CustomAreas = new[] { app.Selection, new FakeRange(app.Sheet, 1, 1, 1, 3) };
        FormulaCommandResult r = Execute(app, FormulaOperation.IFError);
        Check("COM multi-area overlap and constants/text skipped", !r.Failed && r.Changed == 2 && app.Sheet.Cell(1, 1).Writes == 1 && app.Sheet.Cell(1, 3).Writes == 1 && app.Sheet.Cell(2, 1).Writes == 0);
        Check("full column bounded to UsedRange", app.Sheet.Requests < 20);
        Check("Formula2 preferred", app.Sheet.FormulaWrites == 0 && app.Sheet.Formula2Writes == 2);
        app = One(); app.Sheet.Cell(1, 1).Merge = new FakeRange(app.Sheet, 1, 1, 2, 1); app.Sheet.Set(2, 1, "=B1", true); app.Sheet.Cell(2, 1).Merge = app.Sheet.Cell(1, 1).Merge;
        app.Selection = new FakeRange(app.Sheet, 1, 1, 2, 1); r = Execute(app, FormulaOperation.Abs);
        Check("merged anchor once", !r.Failed && r.Changed == 1 && app.Sheet.Cell(2, 1).Writes == 0);
        app = One(); app.Sheet.Set(2, 1, "=B1", true); app.Sheet.Cell(2, 1).Merge = new FakeRange(app.Sheet, 1, 1, 2, 1); app.Selection = new FakeRange(app.Sheet, 2, 1, 1, 1); r = Execute(app, FormulaOperation.Abs);
        Check("merged anchor outside selection skipped", r.Changed == 0 && app.Sheet.Cell(1, 1).Writes == 0);
        foreach (string rejection in new[] { "array", "spill", "locked", "unreadable", "Formula2error", "spillMissing", "grouped" })
        {
            app = One(); FakeCell cell = app.Sheet.Cell(1, 1);
            if (rejection == "array") cell.Array = true;
            if (rejection == "spill") cell.Spill = true;
            if (rejection == "locked") { app.Sheet.ProtectContents = true; cell.Locked = true; }
            if (rejection == "unreadable") cell.PropertyError = true;
            if (rejection == "Formula2error") app.Sheet.Formula2Error = true;
            if (rejection == "spillMissing") cell.SpillMissing = true;
            if (rejection == "grouped") app.ActiveWindow.SelectedSheets.Count = 2;
            r = Execute(app, FormulaOperation.Abs);
            Check("COM preflight rejects " + rejection, r.Failed && cell.Writes == 0 && app.EnableEvents && app.ScreenUpdating);
        }
        app = One(); app.Sheet.ProtectContents = true; app.Sheet.Cell(1, 1).Locked = false; r = Execute(app, FormulaOperation.Round);
        Check("protected unlocked allowed", !r.Failed && r.Changed == 1);
        app = One(); app.Sheet.NoFormula2 = true; r = Execute(app, FormulaOperation.Round);
        Check("ordinary Formula fallback only missing property", !r.Failed && app.Sheet.FormulaWrites == 1 && app.Sheet.Formula2Writes == 0);
        app = One(); app.Sheet.NoFormula2 = true; app.Sheet.Cell(1, 1).SpillMissing = true; r = Execute(app, FormulaOperation.Abs);
        Check("fallback cannot confirm ordinary zero writes", r.Failed && app.Sheet.Cell(1, 1).Writes == 0);
        app = One(); app.Sheet.Set(2, 1, "=B2", true); app.Selection = new FakeRange(app.Sheet, 1, 1, 2, 1); app.Sheet.Cell(2, 1).FailWrite = true;
        r = Execute(app, FormulaOperation.Abs);
        Check("COM late write rollback", r.Failed && r.Details.Contains("late write failure") && !r.RollbackIncomplete && (string)app.Sheet.Cell(1, 1).Value == "=B1" && (string)app.Sheet.Cell(2, 1).Value == "=B2");
        app = One(); app.Sheet.Cell(1, 1).FailWrite = app.Sheet.Cell(1, 1).FailRestore = true; r = Execute(app, FormulaOperation.Abs);
        Check("COM recovery failure carries workbook and sheet", r.RollbackIncomplete && r.Details.Contains("测试工作簿.xlsx") && r.Details.Contains("工作表：测试表") && r.Details.Contains("原公式：=B1") && WpsShared.SourcePaths.Exists(r.RecoveryLog));
        app = new FakeApp(); for (int row = 1; row <= 34; row++) for (int col = 1; col <= 260; col++) app.Sheet.Set(row, col, row == 34 && col == 260 ? (object)"=B1" : (object)1, row == 34 && col == 260);
        app.Selection = new FakeRange(app.Sheet, 1, 1, 1048576, 16384); r = Execute(app, FormulaOperation.Abs);
        Check("large UsedRange chunking", !r.Failed && r.Changed == 1 && app.Sheet.MaxBlock <= 4096 && app.Sheet.BlockReads >= 2);
    }
    private static FakeApp One()
    { FakeApp app = new FakeApp(); app.Sheet.Set(1, 1, "=B1", true); app.Selection = new FakeRange(app.Sheet, 1, 1, 1, 1); return app; }

    private static FormulaCommandResult Calculate(FakeApp app, FormulaRequest request)
    { return FormulaCommandEngine.Run(new ComFormulaHost(app, true), request, delegate { }, Recovery); }
    private static void Arithmetic()
    {
        string[] expected = { "=20+100", "=20-100", "=20*100", "=20/10000" };
        FormulaOperation[] ops = { FormulaOperation.Add, FormulaOperation.Subtract, FormulaOperation.Multiply, FormulaOperation.Divide };
        for (int i=0;i<4;i++)
        {
            FakeApp app=One(); app.Sheet.Set(1,1,20.0,false);
            FormulaCommandResult result=Calculate(app,FormulaRequest.Default(ops[i]));
            Check("numeric to formula " + ops[i], !result.Failed && (string)app.Sheet.Cell(1,1).Value == expected[i] && app.Sheet.Cell(1,1).IsFormula && app.EnableEvents && app.ScreenUpdating);
            app=One(); app.Sheet.Set(1,1,"=A1+B1",true);
            result=Calculate(app,FormulaRequest.Default(ops[i]));
            Check("formula priority " + ops[i], !result.Failed && (string)app.Sheet.Cell(1,1).Value == "=(A1+B1)"+FormulaRequest.Default(ops[i]).Operator+FormulaRequest.Default(ops[i]).Argument);
        }
        Check("custom negative decimal", FormulaWrapper.Transform("=A1+B1",FormulaRequest.Arithmetic(FormulaOperation.Multiply,"-1.5",false)) == "=(A1+B1)*(-1.5)");
        Check("reference stays intact", FormulaWrapper.Transform("=B2",FormulaRequest.Arithmetic(FormulaOperation.Add,"SUM(B1:B3)",true)) == "=(B2)+(SUM(B1:B3))");
        bool zero=false; try { FormulaRequest.Arithmetic(FormulaOperation.Divide,"-0.0",false); } catch(FormatException) { zero=true; }
        Check("literal zero divisor rejected",zero);
        Check("zero expression left to host",FormulaRequest.Arithmetic(FormulaOperation.Divide,"A1-A1",true).Argument=="(A1-A1)");
        foreach(object skipped in new object[]{"20",true,new ErrorWrapper(2042),unchecked((int)0x800A07FA),null})
        {
            FakeApp app=One(); app.Sheet.Set(1,1,skipped,false); var result=Calculate(app,FormulaRequest.Default(FormulaOperation.Add));
            Check("non numeric skipped " + Convert.ToString(skipped),!result.Failed && result.Changed==0 && app.Sheet.Cell(1,1).Writes==0);
        }
        foreach(string guard in new string[]{"array","spill","protected","unreadable"})
        {
            FakeApp app=One(); app.Sheet.Set(1,1,20.0,false); app.Sheet.Set(1,2,30.0,false); app.Selection=new FakeRange(app.Sheet,1,1,1,2);
            FakeCell bad=app.Sheet.Cell(1,2); bad.Array=guard=="array"; bad.Spill=guard=="spill"; bad.PropertyError=guard=="unreadable";
            if(guard=="protected") { app.Sheet.ProtectContents=true; bad.Locked=true; }
            var result=Calculate(app,FormulaRequest.Default(FormulaOperation.Add));
            Check("arithmetic batch preflight " + guard,result.Failed && app.Sheet.Cell(1,1).Writes==0 && bad.Writes==0 && app.EnableEvents && app.ScreenUpdating);
        }
        foreach(bool incomplete in new bool[]{false,true})
        {
            FakeApp app=One(); app.Sheet.Set(1,1,20.0,false); app.Sheet.Set(1,2,"=C1",true); app.Selection=new FakeRange(app.Sheet,1,1,1,2);
            app.Sheet.Cell(1,2).FailArithmetic=true; app.Sheet.Cell(1,1).FailRestore=incomplete;
            var result=Calculate(app,FormulaRequest.Default(FormulaOperation.Multiply));
            Check("mixed rollback " + incomplete,result.Failed && result.RollbackIncomplete==incomplete && app.EnableEvents && app.ScreenUpdating);
            if(!incomplete) Check("numeric restored via Value2",app.Sheet.Cell(1,1).Value is double && (double)app.Sheet.Cell(1,1).Value==20 && !app.Sheet.Cell(1,1).IsFormula && app.Sheet.ValueWrites==1 && (string)app.Sheet.Cell(1,2).Value=="=C1");
            else Check("numeric recovery log includes type and value",WpsShared.SourcePaths.ReadAllText(result.RecoveryLog).Contains("原数值（System.Double）：20"));
        }
        FakeApp changed=One(); changed.Sheet.Set(1,1,20.0,false); ComFormulaHost host=new ComFormulaHost(changed,true); host.Collect(); changed.Sheet.Cell(1,1).Value="20";
        var r=FormulaCommandEngine.Run(host,FormulaRequest.Default(FormulaOperation.Add),delegate{},Recovery);
        Check("type changed during editor zero writes",r.Failed && changed.Sheet.Cell(1,1).Writes==0);
        foreach(string change in new[]{"group","identity"})
        {
            FakeApp scoped=One(); scoped.Sheet.Set(1,1,20.0,false); ComFormulaHost bound=new ComFormulaHost(scoped,true); bound.Collect();
            if(change=="group")scoped.ActiveWindow.SelectedSheets.Count=2; else scoped.Sheet.SheetName="Other";
            var result=FormulaCommandEngine.Run(bound,FormulaRequest.Default(FormulaOperation.Add),delegate{},Recovery);
            Check("arithmetic scope recheck " + change,result.Failed && scoped.Sheet.Cell(1,1).Writes==0 && scoped.EnableEvents && scoped.ScreenUpdating);
        }
        FakeApp typeRestore=One(); typeRestore.Sheet.Set(1,1,20.0,false); typeRestore.Sheet.Set(1,2,"=C1",true); typeRestore.Selection=new FakeRange(typeRestore.Sheet,1,1,1,2);
        typeRestore.Sheet.Cell(1,1).BadNumericRestore=true; typeRestore.Sheet.Cell(1,2).FailArithmetic=true;
        var typeResult=Calculate(typeRestore,FormulaRequest.Default(FormulaOperation.Add));
        Check("same text numeric restore cannot pass type check",typeResult.RollbackIncomplete && typeRestore.Sheet.Cell(1,1).Value is string);
        FakeApp multi=One(); multi.Sheet.Set(1,1,20.0,false); multi.Sheet.Set(1,2,30.0,false); multi.Selection.CustomAreas=new[]{new FakeRange(multi.Sheet,1,1,1,2),new FakeRange(multi.Sheet,1,2,1,1)};
        r=Calculate(multi,FormulaRequest.Default(FormulaOperation.Add)); Check("arithmetic overlap dedup",!r.Failed && r.Changed==2 && multi.Sheet.Cell(1,2).Writes==1);
        FakeApp merge=One(); merge.Sheet.Set(1,1,20.0,false); merge.Sheet.Set(1,2,20.0,false); merge.Sheet.Cell(1,1).Merge=merge.Sheet.Cell(1,2).Merge=new FakeRange(merge.Sheet,1,1,1,2); merge.Selection=new FakeRange(merge.Sheet,1,1,1,2);
        r=Calculate(merge,FormulaRequest.Default(FormulaOperation.Add)); Check("arithmetic merged anchor once",!r.Failed && r.Changed==1 && merge.Sheet.Cell(1,2).Writes==0);
        foreach(FormulaOperation op in ops) using(FormulaOptionsForm form=new FormulaOptionsForm(op))
        {
            Check("two arithmetic parameter types " + op,form.ValueType.Items.Count==2);
            form.ValueType.SelectedIndex=1; form.ValueInput.Text="A1"; FormulaRequest request;
            Check("custom reference validates " + op,form.TryRequest(out request) && request.Argument=="(A1)");
            form.Show(); form.Close(); Check("arithmetic close cancels " + op,form.Request==null && form.DialogResult==System.Windows.Forms.DialogResult.Cancel);
        }
    }
    private static FakeApp Headers()
    {
        FakeApp app=new FakeApp(); string[] titles={"日期","数据","编制人","审核人"};
        for(int i=0;i<4;i++) { app.Sheet.Set(1,i+1,titles[i],false); app.Sheet.Set(2,i+1,1,false); }
        app.Selection=new FakeRange(app.Sheet,1,1,1,16384); return app;
    }
    private static void Columns()
    {
        TableColumns();
        FakeApp app=Headers();
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        {
            var choices=host.Collect(); Check("whole selected row covers UsedRange only",choices.Count==4 && choices[0].Header=="日期" && choices[3].Letter=="D");
            using(ColumnFilterForm form=new ColumnFilterForm(choices))
            {
                form.Grid.Rows[2].Cells[0].Value=false; form.Grid.Rows[3].Cells[0].Value=false;
                Check("checks are pending only",app.Sheet.HiddenWrites==0);
                var result=ColumnVisibilityEngine.Apply(host,form.Pending(),Recovery);
                Check("A B visible C D hidden",!result.Failed && result.Changed==2 && !host.ReadHidden(1) && !host.ReadHidden(2) && app.Sheet.Hidden[3] && app.Sheet.Hidden[4]);
                Check("outside used region unchanged",!app.Sheet.Hidden.ContainsKey(5) && app.EnableEvents && app.ScreenUpdating);
            }
        }
        app=Headers(); app.Sheet.Hidden[2]=true;
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        using(ColumnFilterForm form=new ColumnFilterForm(host.Collect()))
        {
            form.Show(); form.Grid.Sort(form.Grid.Columns[1],System.ComponentModel.ListSortDirection.Descending);
            foreach(System.Windows.Forms.DataGridViewRow gridRow in form.Grid.Rows) gridRow.Cells[0].Value=Convert.ToString(gridRow.Cells[1].Value)=="A" || Convert.ToString(gridRow.Cells[1].Value)=="B";
            var pending=form.Pending(); Check("sorted grid checks follow column identity",pending[0].Visible && pending[1].Visible && !pending[2].Visible && !pending[3].Visible);
            form.Reset.PerformClick(); pending=form.Pending(); Check("sorted grid restores original checks by identity",pending[0].Visible && !pending[1].Visible && pending[2].Visible && pending[3].Visible);
            form.Close(); Check("sorted cancellation zero writes",app.Sheet.HiddenWrites==0);
        }
        app=Headers(); app.Sheet.Hidden[2]=true; app.Sheet.Set(1,3,"",false); app.Sheet.Set(1,4,"数据",false);
        app.Sheet.Cell(1,1).Merge=app.Sheet.Cell(1,2).Merge=new FakeRange(app.Sheet,1,1,1,2);
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        {
            var list=host.Collect(); Check("hidden blank duplicate merge titles",!list[1].Visible && list[2].Header=="（空标题）" && list[1].Header.Contains("合并") && list[1].Header.Contains("日期"));
            using(ColumnFilterForm form=new ColumnFilterForm(list)) { form.Show(); form.Close(); Check("close without apply zero writes",app.Sheet.HiddenWrites==0); }
        }
        foreach(string invalid in new[]{"group","protected","unreadable"})
        {
            app=Headers(); if(invalid=="multirow")app.Selection=new FakeRange(app.Sheet,1,1,2,1);
            if(invalid=="multi-area-row")app.Selection.CustomAreas=new[]{new FakeRange(app.Sheet,1,1,1,1),new FakeRange(app.Sheet,2,1,1,1)};
            if(invalid=="group")app.ActiveWindow.SelectedSheets.Count=2;
            if(invalid=="protected")app.Sheet.ProtectContents=true;
            if(invalid=="unreadable")app.Sheet.HiddenUnreadable=true;
            bool rejected=false; using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))try{host.Collect();}catch(Exception){rejected=true;}
            Check("invalid column selection " + invalid,rejected && app.Sheet.HiddenWrites==0 && app.EnableEvents && app.ScreenUpdating);
        }
        foreach(string changed in new[]{"header","hidden","bounds","protection","identity"})
        {
            app=Headers(); using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
            {
                var list=host.Collect(); list[3].Visible=false;
                if(changed=="header")app.Sheet.Cell(1,1).Value="新标题";
                if(changed=="hidden")app.Sheet.Hidden[1]=true;
                if(changed=="bounds")app.Sheet.Set(1,5,"新增",false);
                if(changed=="protection")app.Sheet.ProtectContents=true;
                if(changed=="identity")app.Sheet.SheetName="Other";
                var result=ColumnVisibilityEngine.Apply(host,list,Recovery); Check("column stale " + changed,result.Failed && app.Sheet.HiddenWrites==0 && app.EnableEvents && app.ScreenUpdating);
            }
        }
        foreach(bool incomplete in new[]{false,true})
        {
            app=Headers(); using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
            {
                var list=host.Collect(); list[2].Visible=list[3].Visible=false; app.Sheet.HiddenFail=4; app.Sheet.HiddenRestoreFail=incomplete?3:0;
                var result=ColumnVisibilityEngine.Apply(host,list,Recovery);
                Check("column verified rollback " + incomplete,result.Failed && result.RollbackIncomplete==incomplete && app.EnableEvents && app.ScreenUpdating);
                if(incomplete)Check("column recovery log",WpsShared.SourcePaths.Exists(result.RecoveryLog) && result.Details.Contains("C 列"));
                else Check("both attempted columns restored",!app.Sheet.Hidden[3] && !app.Sheet.Hidden[4]);
            }
        }
        app=Headers(); app.Sheet.UsedStart=3; app.Sheet.Cell(1,1).Value=app.Sheet.Cell(1,2).Value=null; app.Sheet.Cell(2,1).Value=app.Sheet.Cell(2,2).Value=null; app.Selection=new FakeRange(app.Sheet,2,3,1,1);
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        {
            var list=host.Collect(); Check("UsedRange begins at C",list.Count==2 && list[0].Letter=="C");
            list[1].Visible=false; var r=ColumnVisibilityEngine.Apply(host,list,Recovery); Check("outside A B untouched",!r.Failed && !app.Sheet.Hidden.ContainsKey(1) && !app.Sheet.Hidden.ContainsKey(2));
        }
        app=Headers(); using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        { var list=host.Collect(); foreach(var item in list)item.Visible=false; var r=ColumnVisibilityEngine.Apply(host,list,Recovery); Check("at least one visible",r.Failed && app.Sheet.HiddenWrites==0); }
        app=Headers(); app.Sheet.ProtectContents=true; app.Sheet.AllowColumnFormatting=true; app.EnableEvents=app.ScreenUpdating=false;
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
        { var list=host.Collect(); list[3].Visible=false; var result=ColumnVisibilityEngine.Apply(host,list,Recovery); Check("protected allowed column formatting and false flags preserved",!result.Failed && !app.EnableEvents && !app.ScreenUpdating); }
        app=Headers(); app.Selection.CustomAreas=new[]{new FakeRange(app.Sheet,1,1,1,1),new FakeRange(app.Sheet,1,3,1,1)};
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app)) Check("same row multi-area accepted",host.Collect().Count==4);
        app=Headers(); app.Selection=new FakeRange(app.Sheet,2,1,1,1);
        using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app)) Check("header row outside UsedRange accepted",host.Collect()[0].Header=="日期");
        bool nonCell=false; using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(new InvalidSelectionApp()))try{host.Collect();}catch(InvalidOperationException){nonCell=true;}
        Check("non-cell selection rejected",nonCell);
    }

    private static void TableColumns()
    {
        FakeApp app=new FakeApp();
        for(int r=1;r<=10;r++)for(int c=1;c<=12;c++)app.Sheet.Set(r,c,r==1?(object)("标题"+c):r,false);
        app.Selection=new FakeRange(app.Sheet,5,2,1,1);
        using(var host=new ComColumnVisibilityHost(app)) {var list=host.Collect();Check("B5 detects A1 L10 header",list.Count==12&&list[0].Header=="标题1"&&list[11].Header=="标题12"&&host.RangeDescription.Contains("A1:L10"));}
        app=Headers();app.Selection=new FakeRange(app.Sheet,2,2,1,1);app.Sheet.Set(1,11,null,false);
        using(var host=new ComColumnVisibilityHost(app))Check("format-only empty columns excluded",host.Collect().Count==4);
        app.Sheet.AutoFilter=new FakeFilter{Range=new FakeRange(app.Sheet,1,1,10,5)};
        using(var host=new ComColumnVisibilityHost(app))Check("containing AutoFilter includes internal blank header",host.Collect().Count==5&&host.RangeDescription.Contains("A1:E10"));
        app.Sheet.ListObjects.Items.Add(new FakeTable{Range=new FakeRange(app.Sheet,1,1,2,4)});
        using(var host=new ComColumnVisibilityHost(app))Check("formal table wins over filter",host.Collect().Count==4&&host.RangeDescription.StartsWith("正式表格"));
        app.Sheet.Cell(1,2).Value=null;
        using(var host=new ComColumnVisibilityHost(app))Check("formal blank header retained",host.Collect()[1].Header=="（空标题）");
        app.Sheet.ListObjects.Items.Add(new FakeTable{Name="Other",Range=new FakeRange(app.Sheet,20,8,4,3)});app.Sheet.Set(20,8,"别表",false);app.Selection=new FakeRange(app.Sheet,22,9,1,1);
        using(var host=new ComColumnVisibilityHost(app)) {var list=host.Collect();Check("multiple tables non-A origin",list.Count==3&&list[0].Letter=="H"&&list[0].Header=="别表");app.Sheet.ListObjects.Items[1].Range=new FakeRange(app.Sheet,20,8,5,3);bool refused=false;try{host.Recheck();}catch{refused=true;}Check("table row-range mutation rejected",refused);}
        app=Headers();using(var host=new ComColumnVisibilityHost(app))using(var form=new ColumnFilterForm(host.Collect()))
        {
            form.Show();form.SelectAll.PerformClick();Check("first all sets every candidate",form.Pending().All(c=>c.Visible)&&form.SelectAll.Text=="取消全选");form.SelectAll.PerformClick();Check("second all clears every candidate",form.Pending().All(c=>!c.Visible));
            form.Search.Text="数据";form.SelectAll.PerformClick();Check("search all touches only matches",form.Pending().Count(c=>c.Visible)==1&&form.Pending()[1].Visible);
            form.Search.Text="编制人";Check("search change resets select action",form.SelectAll.Text=="全选");form.SelectAll.PerformClick();form.SelectAll.PerformClick();Check("searched toggle leaves other checks",form.Pending()[1].Visible&&!form.Pending()[2].Visible);
            form.Reset.PerformClick();Check("restore resets select action",form.SelectAll.Text=="全选");form.Close();
        }
        app=new FakeApp();app.Selection=new FakeRange(app.Sheet,99,12,1,1);using(var host=new ComColumnVisibilityHost(app)){bool refused=false;try{host.Collect();}catch{refused=true;}Check("isolated blank selection refused",refused);}
    }
    private static void IdentityRegression()
    {
        foreach(string code in new[]{"",null,"missing"})
        {
            FakeApp app=One(); app.Sheet.Code=code; app.Sheet.CodeMissing=code=="missing";
            var result=Execute(app,FormulaOperation.IFError);
            Check("formula allows optional VBA CodeName " + code,!result.Failed && result.Changed==1 && app.EnableEvents && app.ScreenUpdating);
            app=Headers(); app.Sheet.Code=code; app.Sheet.CodeMissing=code=="missing";
            using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
            {
                var choices=host.Collect(); choices[3].Visible=false;
                var applied=ColumnVisibilityEngine.Apply(host,choices,Recovery);
                Check("columns allow optional VBA CodeName " + code,!applied.Failed && applied.Changed==1 && app.Sheet.Hidden[4]);
            }
        }
        foreach(FormulaOperation op in new[]{FormulaOperation.Add,FormulaOperation.Subtract,FormulaOperation.Multiply,FormulaOperation.Divide})
        {
            FakeApp app=new FakeApp(); app.Sheet.Code=""; app.Sheet.SheetName="Sheet1";
            for(int r=2;r<=22;r++)for(int c=1;c<=11;c++)app.Sheet.Set(r,c,(double)(r*100+c),false);
            app.Selection=new FakeRange(app.Sheet,2,1,21,11);
            var result=Calculate(app,FormulaRequest.Default(op));
            Check("A2 K22 231 numeric cells " + op,!result.Failed && result.Changed==231 && app.Sheet.Cell(2,1).Writes==1 && app.Sheet.Cell(22,11).Writes==1 && app.EnableEvents && app.ScreenUpdating);
        }
        FakeApp numbers=One(); numbers.Sheet.Code=""; numbers.Sheet.Set(1,1,20.0,false);
        var noFormulas=Execute(numbers,FormulaOperation.IFError);
        Check("numeric wrapping explains arithmetic path",!noFormulas.Failed && noFormulas.Changed==0 && noFormulas.Details.Contains("一键公式 → 数值计算") && numbers.Sheet.Cell(1,1).Writes==0);
        foreach(bool anotherBook in new[]{false,true})
        {
            FakeApp app=One(); FakeSheet other=new FakeSheet(); other.Code=app.Sheet.Code; other.SheetName=app.Sheet.SheetName;
            if(!anotherBook)other.Parent=app.Sheet.Parent;
            app.Selection=new FakeRange(other,1,1,1,1);
            var denied=Execute(app,FormulaOperation.Abs);
            Check("same name selection cannot substitute sheet/book " + anotherBook,denied.Failed && app.Sheet.Cell(1,1).Writes==0);
            using(ComColumnVisibilityHost host=new ComColumnVisibilityHost(app))
            {
                bool rejected=false; try{host.Collect();}catch(InvalidOperationException){rejected=true;}
                Check("column selection identity mismatch " + anotherBook,rejected && app.Sheet.HiddenWrites==0);
            }
        }
        foreach(string change in new[]{"rename","saveas","replace-sheet","replace-book","empty-name","empty-path"})
        {
            FakeApp app=One(); FakeSheet original=app.Sheet;
            ComFormulaHost host=new ComFormulaHost(app); host.Collect(); MutateIdentity(app,change);
            var rejected=FormulaCommandEngine.Run(host,FormulaOperation.Abs,delegate{},Recovery);
            Check("formula identity changed zero writes " + change,rejected.Failed && original.Cell(1,1).Writes==0 && app.EnableEvents && app.ScreenUpdating);
            app=Headers(); original=app.Sheet;
            using(ComColumnVisibilityHost columns=new ComColumnVisibilityHost(app))
            {
                var choices=columns.Collect(); choices[3].Visible=false; MutateIdentity(app,change);
                var denied=ColumnVisibilityEngine.Apply(columns,choices,Recovery);
                Check("column identity changed zero writes " + change,denied.Failed && original.HiddenWrites==0 && app.EnableEvents && app.ScreenUpdating);
            }
        }
        foreach(string field in new[]{"name","path"})
        {
            FakeApp app=One(); if(field=="name")app.Sheet.SheetName="";else app.Sheet.Parent.FullName="";
            var result=Execute(app,FormulaOperation.Abs);
            Check("missing necessary identity field refused " + field,result.Failed && app.Sheet.Cell(1,1).Writes==0);
        }
    }
    private static void MutateIdentity(FakeApp app,string change)
    {
        if(change=="rename")app.Sheet.SheetName="renamed";
        if(change=="saveas")app.Sheet.Parent.FullName="C:\\Fake\\Other.xlsx";
        if(change=="replace-sheet")app.Sheet=new FakeSheet { Parent=app.Sheet.Parent, SheetName=app.Sheet.SheetName };
        if(change=="replace-book")app.Sheet.Parent=new FakeWorkbook();
        if(change=="empty-name")app.Sheet.SheetName="";
        if(change=="empty-path")app.Sheet.Parent.FullName="";
    }
    private static void Status()
    {
        FakeApp app = One(); Connect.HostApplication = app;
        var show = typeof(Connect).GetMethod("ShowFormulaStatus", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
        var restore = typeof(Connect).GetMethod("RestoreFormulaStatus", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
        app.StatusBar = false; show.Invoke(null, new object[] { 3 });
        Check("success count in status bar", Convert.ToString(app.StatusBar).Contains("3"));
        restore.Invoke(null, null); Check("restore previous status bar", app.StatusBar is bool && !(bool)app.StatusBar);
        show.Invoke(null, new object[] { 4 });
        object timer = typeof(Connect).GetField("_formulaStatusTimer", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic).GetValue(null);
        Check("status timer is five seconds", (int)timer.GetType().GetProperty("Interval").GetValue(timer, null) == 5000);
        var doEvents = timer.GetType().Assembly.GetType("System.Windows.Forms.Application").GetMethod("DoEvents");
        DateTime deadline = DateTime.UtcNow.AddSeconds(6);
        while (app.StatusBar is string && DateTime.UtcNow < deadline) { doEvents.Invoke(null, null); System.Threading.Thread.Sleep(15); }
        Check("five-second status automatically restored", app.StatusBar is bool && !(bool)app.StatusBar);
        show.Invoke(null, new object[] { 2 }); app.StatusBar = "host changed"; restore.Invoke(null, null);
        Check("newer host status retained", (string)app.StatusBar == "host changed");
        app.StatusBar = "previous"; show.Invoke(null, new object[] { 2 }); Connect.HostApplication = null;
        Check("shutdown restores and cancels timer", (string)app.StatusBar == "previous");
    }
}

// Public mock classes let the same late-binding adapter run against fake objects.
public sealed class FakeCell
{
    public object Value, Original; public bool IsFormula, BadNumericRestore, FailArithmetic, Array, Spill, Locked, PropertyError, SpillMissing, FailWrite, FailRestore;
    public int Writes; public FakeRange Merge;
}
public sealed class FakeApp
{
    public FakeSheet Sheet = new FakeSheet(); public FakeRange Selection;
    public FakeRange ActiveCell { get { return new FakeRange(Selection.Worksheet,Selection.Row,Selection.Column,1,1); } }
    public FakeSheet ActiveSheet { get { return Sheet; } }
    public FakeWindow ActiveWindow { get; set; }
    public bool EnableEvents { get; set; } public bool ScreenUpdating { get; set; } public object StatusBar { get; set; }
    public FakeApp() { EnableEvents = ScreenUpdating = true; ActiveWindow = new FakeWindow(); }
    public FakeRange Intersect(FakeRange a, FakeRange b)
    {
        int r = Math.Max(a.Row, b.Row), c = Math.Max(a.Column, b.Column), endR = Math.Min(a.Row + a.Rows.Count, b.Row + b.Rows.Count), endC = Math.Min(a.Column + a.Columns.Count, b.Column + b.Columns.Count);
        return endR <= r || endC <= c ? null : new FakeRange(Sheet, r, c, endR - r, endC - c);
    }
}
public sealed class FakeWorkbook { public string FullName { get; set; } public FakeWorkbook() { FullName="C:\\Fake\\测试工作簿.xlsx"; } }
public sealed class FakeWindow { public FakeCount SelectedSheets = new FakeCount { Count = 1 }; }
public sealed class FakeCount { public int Count { get; set; } }
public sealed class FakeSheet
{
    private readonly Dictionary<string, FakeCell> cells = new Dictionary<string, FakeCell>();
    public string Code="Sheet1"; public bool CodeMissing;
    public string CodeName { get { if(CodeMissing) throw new COMException("missing", unchecked((int)0x80020003)); return Code; } }
    public string SheetName="测试表"; public string Name { get { return SheetName; } }
    public FakeWorkbook Parent { get; set; }
    public FakeSheet() { Parent=new FakeWorkbook(); }
    public int UsedStart=1, ValueWrites, HiddenWrites, HiddenFail, HiddenRestoreFail; public bool HiddenUnreadable; public Dictionary<int,bool> Hidden=new Dictionary<int,bool>();
    public bool AllowColumnFormatting;
    public FakeProtection Protection { get { return new FakeProtection { AllowFormattingColumns=AllowColumnFormatting }; } }
    public bool ProtectContents, NoFormula2, Formula2Error; public int Requests, MaxBlock, BlockReads, FormulaWrites, Formula2Writes;
    private int maxRow = 1, maxCol = 1;
    public FakeSheet Range { get { return this; } }
    public FakeRange this[string address]
    {
        get
        {
            Requests++; string[] parts = address.Split(':'); int r1, c1, r2, c2;
            Parse(parts[0], out r1, out c1); Parse(parts[parts.Length - 1], out r2, out c2);
            return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
        }
    }
    private static void Parse(string address, out int row, out int col)
    { int i = 0; col = 0; while (i < address.Length && char.IsLetter(address[i])) col = col * 26 + address[i++] - 'A' + 1; row = int.Parse(address.Substring(i)); }
    public void Set(int row, int col, object value, bool formula)
    { Cell(row, col).Original = Cell(row, col).Value = value; Cell(row, col).IsFormula = formula; maxRow = Math.Max(maxRow, row); maxCol = Math.Max(maxCol, col); }
    public FakeCell Cell(int row, int col)
    { string key = row + ":" + col; if (!cells.ContainsKey(key)) cells[key] = new FakeCell(); return cells[key]; }
    public FakeTables ListObjects = new FakeTables(); public FakeFilter AutoFilter;
    public FakeRange Region(int row,int col)
    {
        int top=row,bottom=row,left=col,right=col; bool changed;
        do { changed=false;
            if(top>1 && Enumerable.Range(left,right-left+1).Any(c=>Cell(top-1,c).Value!=null)){top--;changed=true;}
            if(bottom<maxRow && Enumerable.Range(left,right-left+1).Any(c=>Cell(bottom+1,c).Value!=null)){bottom++;changed=true;}
            if(left>1 && Enumerable.Range(top,bottom-top+1).Any(r=>Cell(r,left-1).Value!=null)){left--;changed=true;}
            if(right<maxCol && Enumerable.Range(top,bottom-top+1).Any(r=>Cell(r,right+1).Value!=null)){right++;changed=true;}
        } while(changed);
        return new FakeRange(this,top,left,bottom-top+1,right-left+1);
    }
    public FakeRange UsedRange { get { return new FakeRange(this, 1, UsedStart, maxRow, maxCol-UsedStart+1); } }
}
public sealed class FakeAreas
{
    private readonly FakeRange[] areas;
    public FakeAreas(FakeRange[] ranges) { areas = ranges; }
    public int Count { get { return areas.Length; } }
    public FakeRange Item(int index) { return areas[index - 1]; }
}
public sealed class FakeRange
{
    public FakeSheet Worksheet { get; private set; }
    public int Row { get; private set; } public int Column { get; private set; }
    public FakeCount Rows { get; private set; } public FakeCount Columns { get; private set; }
    public FakeRange[] CustomAreas;
    public FakeAreas Areas { get { return new FakeAreas(CustomAreas ?? new[] { this }); } }
    public FakeRange(FakeSheet sheet, int row, int col, int height, int width)
    { Worksheet = sheet; Row = row; Column = col; Rows = new FakeCount { Count = height }; Columns = new FakeCount { Count = width }; }
    private FakeCell Cell { get { return Worksheet.Cell(Row, Column); } }
    public object HasFormula
    {
        get { bool yes = false, no = false; for (int r = 0; r < Rows.Count; r++) for (int c = 0; c < Columns.Count; c++) { if (Worksheet.Cell(Row + r, Column + c).IsFormula) yes = true; else no = true; } return yes && no ? null : (object)yes; }
    }
    public bool MergeCells { get { return Cell.Merge != null; } }
    public FakeRange MergeArea { get { return Cell.Merge; } }
    public bool HasArray { get { if (Cell.PropertyError) throw new Exception("属性失败"); return Cell.Array; } }
    public bool HasSpill { get { if (Cell.SpillMissing) throw new COMException("missing", unchecked((int)0x80020003)); return Cell.Spill; } }
    public bool Locked { get { return Cell.Locked; } }
    public object Value2 { get { return Read(); } set { Worksheet.ValueWrites++; Write(value); Cell.IsFormula=false; if(Cell.BadNumericRestore)Cell.Value=Convert.ToString(value); } }
    public string Text { get { return Convert.ToString(Cell.Value); } }
    public string Address { get { return "$"+CellAddress.Name(Row,Column)+":$"+CellAddress.Name(Row+Rows.Count-1,Column+Columns.Count-1); } }
    public FakeRange CurrentRegion { get { return Worksheet.Region(Row,Column); } }
    public FakeRange EntireColumn { get { return this; } }
    public object Hidden
    {
        get { if(Worksheet.HiddenUnreadable)return null; bool hidden; return Worksheet.Hidden.TryGetValue(Column,out hidden) && hidden; }
        set { Worksheet.HiddenWrites++; bool hidden=(bool)value; if(!hidden && Worksheet.HiddenRestoreFail==Column)throw new Exception("restore Hidden failed"); Worksheet.Hidden[Column]=hidden; if(hidden && Worksheet.HiddenFail==Column)throw new Exception("Hidden setter mutated then failed"); }
    }
    public object Formula2
    {
        get { if (Worksheet.NoFormula2) throw new COMException("missing", unchecked((int)0x80020003)); if (Worksheet.Formula2Error) throw new COMException("generic error", unchecked((int)0x80004005)); return Read(); }
        set { Worksheet.Formula2Writes++; Write(value); }
    }
    public object Formula { get { return Read(); } set { Worksheet.FormulaWrites++; Write(value); } }
    private object Read()
    {
        Worksheet.MaxBlock = Math.Max(Worksheet.MaxBlock, Rows.Count * Columns.Count); Worksheet.BlockReads++;
        if (Rows.Count * Columns.Count == 1) return Cell.Value;
        Array a = Array.CreateInstance(typeof(object), new[] { Rows.Count, Columns.Count }, new[] { 1, 1 });
        for (int r = 0; r < Rows.Count; r++) for (int c = 0; c < Columns.Count; c++) a.SetValue(Worksheet.Cell(Row + r, Column + c).Value, r + 1, c + 1);
        return a;
    }
    private void Write(object value)
    { Cell.Writes++; if (Cell.FailRestore && Equals(value, Cell.Original)) throw new Exception("fake restore failure"); Cell.Value = value; Cell.IsFormula = value is string && ((string)value).StartsWith("="); if(Cell.FailArithmetic && Cell.IsFormula && !Equals(value,Cell.Original)) throw new Exception("arithmetic setter mutated then failed"); if (Cell.FailWrite && Convert.ToString(value).StartsWith("=ABS(")) throw new Exception("late write failure"); }
}

public sealed class FakeProtection { public bool AllowFormattingColumns { get; set; } }
public sealed class InvalidSelectionApp
{
    public FakeWindow ActiveWindow { get { return new FakeWindow(); } }
    public FakeSheet ActiveSheet { get { return new FakeSheet(); } }
    public object Selection { get { return new object(); } }
    public object ActiveCell { get { return new object(); } }
}

public sealed class FakeTables { public List<FakeTable> Items=new List<FakeTable>(); public int Count{get{return Items.Count;}} public FakeTable Item(int i){return Items[i-1];} }
public sealed class FakeTable { public string Name="Table1"; public FakeRange Range; public FakeRange HeaderRowRange{get{return new FakeRange(Range.Worksheet,Range.Row,Range.Column,1,Range.Columns.Count);}} }
public sealed class FakeFilter { public FakeRange Range; }
