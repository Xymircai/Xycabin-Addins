using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal interface IFormulaOperationDialog
    {
        WorksheetUndoButton Undo { get; }
        event Action<FormulaRequest> Submitted;
        void SetTarget(string target, bool ready);
        void Resume(string message);
    }

    // Deliberately contains no Application, Workbook, Worksheet, Range or host
    // callback closure. Completion carries only a ticket and a pure request.
    internal sealed class FormulaWindowThread
    {
        private static readonly object registryLock = new object();
        private static readonly List<FormulaWindowThread> windows = new List<FormulaWindowThread>();
        private readonly object sync = new object();
        private readonly long ticket;
        private readonly FormulaOperation operation;
        private readonly IntPtr owner;
        private readonly string notice;
        private readonly bool warning;
        private Form form;
        private Thread thread;
        private bool closing;
        private WorksheetUndoState undoState;
        private readonly string target;
        private readonly bool canApply;
        internal static bool HiddenForProbe;
        internal static Action<FormulaOptionsForm> CreatedForProbe;
        internal static Action NoticeForProbe;
        internal static int Count { get { lock (registryLock) return windows.Count; } }
        internal FormulaWindowThread(long id, FormulaOperation selected, IntPtr ownerHandle, string description, bool ready)
        { ticket = id; operation = selected; owner = ownerHandle; target = description; canApply = ready; }
        private FormulaWindowThread(string text, IntPtr ownerHandle, bool isWarning)
        { notice = text; owner = ownerHandle; warning = isWarning; }
        private sealed class Owner : IWin32Window
        {
            private readonly IntPtr value;
            internal Owner(IntPtr handle) { value = handle; }
            public IntPtr Handle { get { return value; } }
        }
        internal void Start()
        {
            lock (registryLock) windows.Add(this);
            try
            {
                thread = new Thread(Run) { IsBackground = true, Name = "Formula parameter UI" };
                thread.SetApartmentState(ApartmentState.STA); thread.Start();
            }
            catch { lock (registryLock) windows.Remove(this); throw; }
        }
        private void Run()
        {
            Form window = null; Panel parking = null; bool sent = false, failed = false;
            ThreadExceptionEventHandler onError = delegate(object sender, ThreadExceptionEventArgs args)
            { failed = true; UiDiagnostics.Record("formula-editor-event", args.Exception); Close(); };
            Application.ThreadException += onError;
            try
            {
                if (notice == null) window = operation == FormulaOperation.Custom ? (Form)new CustomFunctionForm() { SuppressActivation = HiddenForProbe } : new FormulaOptionsForm(operation) { SuppressActivation = HiddenForProbe };
                else
                {
                    var message = new UiMessageForm(notice, "一键公式", MessageBoxButtons.OK,
                        warning ? MessageBoxIcon.Warning : MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    // DialogResult alone closes only modal forms.
                    message.Positive.Click += delegate { message.Close(); }; window = message;
                }
                var dialog = window as IFormulaOperationDialog;
                if (dialog != null)
                {
                    dialog.SetTarget(target, canApply);
                    dialog.Submitted += delegate(FormulaRequest request) { FormulaInteraction.Receive(ticket, request); };
                    dialog.Undo.Click += delegate
                    {
                        var state = dialog.Undo.State;
                        if (state == null || !window.Enabled) return;
                        window.Enabled = false;
                        FormulaInteraction.ReceiveUndo(ticket, state.Id);
                    };
                }
                lock (sync) { form = window; if (dialog != null) dialog.Undo.SetState(undoState); }
                window.Load += delegate
                {
                    lock (sync) { if (dialog != null) dialog.Undo.SetState(undoState); if (closing) QueueClose(); }
                };
                window.FormClosed += delegate
                {
                    if (ticket == 0 || sent) return;
                    sent = true;
                    FormulaInteraction.Receive(ticket, null);
                };
                // A hidden, non-top-level fixture exercises the real loop without
                // showing, disabling or focusing a desktop window.
                if (HiddenForProbe)
                {
                    parking = new Panel { Size = window.Size };
                    window.TopLevel = false; window.FormBorderStyle = FormBorderStyle.None; parking.Controls.Add(window);
                }
                using (var context = new ApplicationContext(window))
                {
                    if (HiddenForProbe || owner == IntPtr.Zero) window.Show(); else window.Show(new Owner(owner));
                    if (notice == null && CreatedForProbe != null && window is FormulaOptionsForm) CreatedForProbe((FormulaOptionsForm)window);
                    if (notice != null && NoticeForProbe != null) NoticeForProbe();
                    Application.Run(context);
                }
            }
            catch (Exception error) { failed = true; UiDiagnostics.Record("formula-editor-thread", error); }
            finally
            {
                if (ticket != 0 && !sent) FormulaInteraction.Receive(ticket, null);
                try { if (window != null) window.Dispose(); if (parking != null) parking.Dispose(); }
                catch (Exception error) { UiDiagnostics.Record("formula-editor-dispose", error); }
                finally
                {
                    Application.ThreadException -= onError;
                    lock (sync) { form = null; thread = null; }
                    lock (registryLock) windows.Remove(this);
                }
            }
            if (failed && notice == null && !closing && !HiddenForProbe)
                Notice("参数窗口无法正常运行，本次没有提交公式。请重新执行；详细信息见本机 UiDiagnostics.log。", owner, true);
        }
        internal void Activate()
        {
            lock (sync)
            {
                var active = form;
                if (closing || active == null || !active.IsHandleCreated || active.IsDisposed || HiddenForProbe) return;
                try { active.BeginInvoke((MethodInvoker)delegate { if (!active.IsDisposed) { if (active.WindowState == FormWindowState.Minimized) active.WindowState = FormWindowState.Normal; active.Activate(); } }); }
                catch (InvalidOperationException) { }
            }
        }
        internal void Resume(string message, string description = null)
        {
            Queue(delegate(Form active)
            {
                var dialog = (IFormulaOperationDialog)active;
                if (description != null) dialog.SetTarget(description, true);
                dialog.Resume(message);
                active.Visible = true;
                if (!HiddenForProbe)
                {
                    if (active.WindowState == FormWindowState.Minimized) active.WindowState = FormWindowState.Normal;
                    active.Activate();
                }
            });
        }
        internal void SetUndo(WorksheetUndoState state)
        {
            lock (sync) { undoState = state; Queue(active => ((IFormulaOperationDialog)active).Undo.SetState(state)); }
        }
        private void Queue(Action<Form> action)
        {
            lock (sync)
            {
                var active = form;
                if (closing || active == null || !active.IsHandleCreated || active.IsDisposed) return;
                try { active.BeginInvoke((MethodInvoker)delegate { if (!active.IsDisposed) action(active); }); }
                catch (InvalidOperationException) { }
            }
        }
        private void QueueClose()
        {
            var active = form;
            if (active == null || !active.IsHandleCreated || active.IsDisposed) return;
            try { active.BeginInvoke((MethodInvoker)delegate { if (!active.IsDisposed) active.Close(); }); }
            catch (InvalidOperationException) { }
        }
        internal void Close() { lock (sync) { closing = true; QueueClose(); } }
        internal static void CloseAll()
        {
            FormulaWindowThread[] pending; lock (registryLock) pending = windows.ToArray();
            foreach (var window in pending) window.Close();
        }
        internal static void Notice(string text, IntPtr ownerHandle, bool isWarning)
        { new FormulaWindowThread(text, ownerHandle, isWarning).Start(); }
    }
}
