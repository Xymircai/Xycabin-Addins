using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel_Iterate;

internal static class FormulaInteractionProbe
{
    private static int checks, hostThread;
    private static string sandbox;
    private static volatile FormulaOptionsForm editor;
    private static int editorThread, editorSta, editorModal;
    private static FormulaCommandResult result;
    private static long ticket;
    [DllImport("user32.dll")] private static extern IntPtr GetFocus();
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
    [DllImport("user32.dll")] private static extern IntPtr GetShellWindow();
    private static void Check(bool condition, string name)
    { if (!condition) throw new InvalidOperationException(name); checks++; Console.WriteLine("PASS " + name); }
    private static void PumpUntil(Func<bool> condition)
    {
        DateTime until = DateTime.UtcNow.AddSeconds(12);
        while (!condition() && DateTime.UtcNow < until) { Application.DoEvents(); Thread.Sleep(2); }
        Check(condition(), "asynchronous operation finishes");
    }
    private static void Created(FormulaOptionsForm form)
    {
        editorThread = Thread.CurrentThread.ManagedThreadId;
        editorSta = Thread.CurrentThread.GetApartmentState() == ApartmentState.STA ? 1 : 0;
        editorModal = form.Modal ? 1 : 0;
        Check(!IsWindowVisible(form.Handle), "fixture has no visible native window");
        editor = form;
    }
    private static Host Open(FormulaOperation operation)
    {
        editor = null; result = null; var host = new Host();
        SynchronizationContext.SetSynchronizationContext(null);
        ticket = FormulaInteraction.Open(host, operation, IntPtr.Zero);
        Check(SynchronizationContext.Current == null, "host synchronization context is preserved");
        PumpUntil(() => editor != null);
        Check(editorThread != hostThread && editorSta == 1 && editorModal == 0, "parameters use a separate non-modal STA");
        Check(host.Events && host.Screen && host.Disposals == 0, "waiting leaves host enabled and target alive");
        Check(Connect.TryBeginFormulaExecution(), "waiting does not occupy global busy flag"); Connect.EndFormulaExecution();
        return host;
    }
    private static void CloseEditor(string action)
    {
        FormulaOptionsForm window = editor;
        if (action == "apply")
        {
            window.BeginInvoke((MethodInvoker)delegate { window.Apply.PerformClick(); });
            PumpUntil(() => result != null);
            int inspected = 0;
            window.BeginInvoke((MethodInvoker)delegate
            {
                Check(window.Enabled && !window.IsDisposed, "apply returns control to the same formula window");
                Interlocked.Exchange(ref inspected, 1);
            });
            PumpUntil(() => Volatile.Read(ref inspected) == 1);
            Check(FormulaInteraction.HasPending && FormulaWindowThread.Count == 1, "the formula session stays open after apply");
            action = "close";
        }
        window.BeginInvoke((MethodInvoker)delegate
        {
            if (action == "cancel") window.Cancel.PerformClick();
            else if (action == "escape") typeof(FormulaOptionsForm).GetMethod("OnKeyDown", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .Invoke(window, new object[] { new KeyEventArgs(Keys.Escape) });
            else window.Close();
        });
        PumpUntil(() => !FormulaInteraction.HasPending && FormulaWindowThread.Count == 0);
    }
    private static void Cases()
    {
        CheckArithmeticSelector();
        foreach (var operation in new[] { FormulaOperation.IFError, FormulaOperation.Round, FormulaOperation.Abs, FormulaOperation.Add, FormulaOperation.Subtract, FormulaOperation.Multiply, FormulaOperation.Divide })
        foreach (string action in new[] { "apply", "cancel", "escape", "close" })
        {
            IntPtr focus = GetFocus(); Host host = Open(operation);
            Check(FormulaInteraction.ActivatePending() && FormulaWindowThread.Count == 1, "repeat command retains one pending editor");
            CloseEditor(action);
            Check(host.Disposals == 1 && host.Events && host.Screen, "target releases exactly once and host state restores " + operation + " " + action);
            Check(action == "apply" ? result != null && !result.Failed && result.Changed == 1 && host.Cell.Writes == 1 : result == null && host.Cell.Writes == 0,
                "apply or zero-write cancellation " + operation + " " + action);
            Check(GetFocus() == focus, "hidden regression never changes desktop focus");
        }
        Host changed = Open(FormulaOperation.Add); changed.Cell.Current = "=private_changed"; CloseEditor("apply");
        Check(result.Failed && changed.Cell.Writes == 0, "changed original content refuses submission");
        foreach (string reason in new[] { "closed", "other-sheet", "protected" })
        {
            Host host = Open(FormulaOperation.Round); host.ScopeFailure = reason; CloseEditor("apply");
            Check(result.Failed && host.Cell.Writes == 0 && host.Disposals == 1, "changed scope refuses " + reason);
        }
        Host late = Open(FormulaOperation.Add); late.Cell.FailWrite = true; CloseEditor("apply");
        Check(result.Failed && !result.RollbackIncomplete && late.Cell.Current == late.Cell.Snapshot && late.Events && late.Screen, "late setter restores original and flags");
        long completedTicket = ticket; int writes = late.Cell.Writes;
        FormulaInteraction.Receive(completedTicket, FormulaRequest.Default(FormulaOperation.Add)); Application.DoEvents();
        Check(late.Cell.Writes == writes && late.Disposals == 1, "duplicate completion cannot write twice");
        Host newer = Open(FormulaOperation.Add);
        FormulaInteraction.Receive(completedTicket, FormulaRequest.Default(FormulaOperation.Round)); Application.DoEvents();
        Check(FormulaInteraction.HasPending && newer.Cell.Writes == 0, "old ticket cannot submit a newer session"); CloseEditor("cancel");
        FormulaWindowThread.NoticeForProbe = delegate { FormulaWindowThread.CloseAll(); };
        Host busy = Open(FormulaOperation.Add); Check(Connect.TryBeginFormulaExecution(), "fixture owns an unrelated busy operation"); CloseEditor("apply");
        Check(busy.Cell.Writes == 0 && busy.Disposals == 1 && !Connect.TryBeginFormulaExecution(), "busy rejection is zero-write and does not release another operation's guard");
        Connect.EndFormulaExecution(); FormulaWindowThread.NoticeForProbe = null;
        Host shutdown = Open(FormulaOperation.Multiply); FormulaInteraction.Shutdown();
        PumpUntil(() => FormulaWindowThread.Count == 0);
        FormulaInteraction.Receive(1, FormulaRequest.Default(FormulaOperation.Add)); Application.DoEvents();
        Check(shutdown.Disposals == 1 && shutdown.Cell.Writes == 0, "shutdown releases bound target and ignores late delivery");
        for (int i = 0; i < 4; i++)
        {
            var early = new Host(); FormulaInteraction.Open(early, FormulaOperation.Round, IntPtr.Zero); FormulaInteraction.Shutdown();
            PumpUntil(() => FormulaWindowThread.Count == 0); Check(early.Disposals == 1 && early.Cell.Writes == 0, "shutdown during editor startup");
        }
        var capture = new Host { FailCapture = true }; bool rejected = false;
        try { FormulaInteraction.Open(capture, FormulaOperation.Round, IntPtr.Zero); } catch (InvalidOperationException) { rejected = true; }
        Check(rejected && capture.Disposals == 1 && !FormulaInteraction.HasPending, "failed capture releases target without editor");
        FormulaWindowThread.CreatedForProbe = delegate { throw new InvalidOperationException("private-input"); };
        var startup = new Host(); FormulaInteraction.Open(startup, FormulaOperation.Round, IntPtr.Zero);
        PumpUntil(() => !FormulaInteraction.HasPending && FormulaWindowThread.Count == 0);
        Check(startup.Disposals == 1 && startup.Cell.Writes == 0, "editor construction failure cancels on host thread");
        FormulaWindowThread.CreatedForProbe = Created;
        Host resumed = Open(FormulaOperation.Add); CloseEditor("apply"); Check(result.Changed == 1, "new session works after shutdown and startup failure");
        Check(FormulaWindowState.ValidateOwner(IntPtr.Zero) == IntPtr.Zero && FormulaWindowState.ValidateOwner(new IntPtr(123)) == IntPtr.Zero && FormulaWindowState.ValidateOwner(GetShellWindow()) == IntPtr.Zero,
            "invalid and foreign owners are rejected");
        int noticeThread = 0;
        FormulaWindowThread.NoticeForProbe = delegate { noticeThread = Thread.CurrentThread.ManagedThreadId; FormulaWindowThread.CloseAll(); };
        FormulaWindowThread.Notice("测试提示", IntPtr.Zero, true);
        PumpUntil(() => FormulaWindowThread.Count == 0);
        Check(noticeThread != 0 && noticeThread != hostThread, "formula result notice also uses an independent loop");
        Check(UiDiagnostics.Text.Contains("native-thread=") && UiDiagnostics.Text.Contains("assembly=") && !UiDiagnostics.Text.Contains("private-input") && !UiDiagnostics.Text.Contains("private_changed"),
            "diagnostics contain state without input or formulas");
    }

    private static void CheckArithmeticSelector()
    {
        var host = Open(FormulaOperation.Add);
        var window = editor;
        int inspected = 0;
        window.BeginInvoke((MethodInvoker)delegate
        {
            var picker = window.OperatorChoice;
            Check(picker != null && picker.Items.Count == 4 && picker.SelectedIndex == 0, "numeric dialog starts at addition and lists four operators");
            for (int index = 0; index < picker.Items.Count; index++)
            {
                picker.SelectedIndex = index;
                FormulaRequest request;
                var expected = (FormulaOperation)((int)FormulaOperation.Add + index);
                Check(window.TryRequest(out request) && request.Operation == expected &&
                      request.Argument == (index == 3 ? "10000" : "100"), "operator switch loads the corresponding default " + expected);
            }
            window.ValueInput.Text = "0";
            Check(!window.Apply.Enabled, "division validation blocks a zero divisor");
            window.ValueInput.Text = "2.5";
            FormulaRequest custom;
            Check(window.TryRequest(out custom) && custom.Operation == FormulaOperation.Divide && custom.Argument == "2.5", "selected operator accepts a custom value");
            Check(!window.Undo.Enabled, "no formula undo enables without an owned batch");
            Interlocked.Exchange(ref inspected, 1);
        });
        PumpUntil(() => Volatile.Read(ref inspected) == 1);
        CloseEditor("cancel");
        Check(host.Cell.Writes == 0 && host.Disposals == 1, "editing arithmetic settings does not write to the workbook");
    }
    private sealed class Host : IFormulaHost, IFormulaScopeHost
    {
        internal readonly Cell Cell = new Cell(); internal bool Events = true, Screen = true, FailCapture; internal int Disposals; internal string ScopeFailure;
        private void OnHost() { if (Thread.CurrentThread.ManagedThreadId != hostThread) throw new InvalidOperationException("COM adapter crossed threads"); }
        public IList<IFormulaCell> Collect() { OnHost(); if (FailCapture) throw new InvalidOperationException("capture failure"); return new IFormulaCell[] { Cell }; }
        public void RecheckScope() { OnHost(); if (ScopeFailure != null) throw new InvalidOperationException(ScopeFailure); }
        public bool EnableEvents { get { OnHost(); return Events; } set { OnHost(); Events = value; } }
        public bool ScreenUpdating { get { OnHost(); return Screen; } set { OnHost(); Screen = value; } }
        public void Dispose() { OnHost(); Disposals++; }
    }
    private sealed class Cell : IFormulaCell
    {
        internal string Current = "=B1"; internal int Writes; internal bool FailWrite;
        public string Address { get { return "A1"; } }
        public string Snapshot { get { return "=B1"; } }
        private void OnHost() { if (Thread.CurrentThread.ManagedThreadId != hostThread) throw new InvalidOperationException("cell crossed threads"); }
        public void CheckWritable() { OnHost(); if (Current != Snapshot) throw new InvalidOperationException("changed"); }
        public string Read() { OnHost(); return Current; }
        public void Write(string text) { OnHost(); Writes++; Current = text; if (FailWrite && text != Snapshot) throw new InvalidOperationException("late write"); }
    }
    [STAThread] public static int Main()
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8; hostThread = Thread.CurrentThread.ManagedThreadId;
        sandbox = Path.Combine(Path.GetTempPath(), "FormulaInteractionProbe-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(sandbox);
        UiDiagnostics.RootOverride = sandbox; DialogSizeStore.RootOverride = sandbox; FormatSession.SourceOverride = Path.Combine(sandbox, "format.xml");
        FormulaInteraction.RecoveryForProbe = sandbox; FormulaInteraction.ResultForProbe = value => result = value;
        FormulaWindowThread.HiddenForProbe = true; FormulaWindowThread.CreatedForProbe = Created; Control.CheckForIllegalCrossThreadCalls = true;
        bool success = false;
        try { Cases(); success = true; Console.WriteLine("FormulaInteractionProbe: " + checks + " PASS"); return 0; }
        catch (Exception error) { Console.WriteLine("FAIL " + error); return 1; }
        finally
        {
            FormulaInteraction.Shutdown();
            DateTime until = DateTime.UtcNow.AddSeconds(12); while (FormulaWindowThread.Count != 0 && DateTime.UtcNow < until) { Application.DoEvents(); Thread.Sleep(2); }
            if (success && FormulaWindowThread.Count == 0 && Path.GetDirectoryName(Path.GetFullPath(sandbox)) == Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar)) Directory.Delete(sandbox, true);
        }
    }
}
