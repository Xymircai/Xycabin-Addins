using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Independent STA: only strings, enums, HWND and ticket cross this boundary.
    internal sealed class ConversionWindowThread
    {
        private readonly long ticket;
        private readonly CellConversionKind kind;
        private readonly string description, address;
        private readonly ConversionPreview initialPreview;
        private readonly string initialError;
        private readonly IntPtr owner;
        private readonly object sync = new object();
        private CellConversionForm form;
        private bool closing;
        private static readonly object registry = new object();
        private static readonly List<ConversionWindowThread> windows = new List<ConversionWindowThread>();
        internal static bool HiddenForProbe;
        internal static Action<CellConversionForm> CreatedForProbe;
        internal static int Count { get { lock (registry) return windows.Count; } }
        internal bool IsAlive { get { lock (sync) return !closing && form != null && !form.IsDisposed; } }
        internal ConversionWindowThread(long id, CellConversionKind selected, IntPtr handle, string scope, string range, ConversionPreview preview, string error)
        { ticket = id; kind = selected; owner = handle; description = scope; address = range; initialPreview = preview; initialError = error; }
        internal void Start()
        {
            lock (registry) windows.Add(this);
            try { var thread = new Thread(Run) { IsBackground = true, Name = "Cell conversion parameters" }; thread.SetApartmentState(ApartmentState.STA); thread.Start(); }
            catch { lock (registry) windows.Remove(this); throw; }
        }
        private sealed class Owner : IWin32Window
        { private readonly IntPtr handle; internal Owner(IntPtr value) { handle = value; } public IntPtr Handle { get { return handle; } } }
        private void Run()
        {
            CellConversionForm window = null; Panel parking = null;
            ThreadExceptionEventHandler onError = delegate(object sender, ThreadExceptionEventArgs args) { UiDiagnostics.Record("conversion-ui-event", args.Exception); Close(); };
            Application.ThreadException += onError;
            try
            {
                window = new CellConversionForm(kind, description, address, initialPreview, initialError) { SuppressActivation = HiddenForProbe };
                lock (sync) form = window;
                window.Submitted += delegate(bool pick, ConversionRequest request) { ConversionInteraction.Receive(ticket, pick, request); };
                window.PreviewRequested += delegate(long version, string range) { ConversionInteraction.ReadPreview(ticket, version, range); };
                window.FormClosed += delegate { ConversionInteraction.Receive(ticket, false, null); };
                window.Load += delegate { lock (sync) { if (closing) Queue(delegate(CellConversionForm f) { f.Close(); }); } };
                if (HiddenForProbe) { parking = new Panel { Size = window.Size }; window.TopLevel = false; window.FormBorderStyle = FormBorderStyle.None; parking.Controls.Add(window); }
                using (var context = new ApplicationContext(window))
                {
                    if (HiddenForProbe || owner == IntPtr.Zero) window.Show(); else window.Show(new Owner(owner));
                    if (CreatedForProbe != null) CreatedForProbe(window);
                    Application.Run(context);
                }
            }
            catch (Exception error) { UiDiagnostics.Record("conversion-ui-thread", error); }
            finally
            {
                ConversionInteraction.Receive(ticket, false, null);
                try { if (window != null) window.Dispose(); if (parking != null) parking.Dispose(); }
                finally { Application.ThreadException -= onError; lock (sync) form = null; lock (registry) windows.Remove(this); }
            }
        }
        private void Queue(Action<CellConversionForm> action)
        {
            lock (sync)
            {
                var active = form; if (active == null || !active.IsHandleCreated || active.IsDisposed) return;
                try { active.BeginInvoke((MethodInvoker)delegate { if (!active.IsDisposed) action(active); }); } catch (InvalidOperationException) { }
            }
        }
        internal void Resume(string address, string error) { Queue(delegate(CellConversionForm f) { f.Resume(address, error); }); }
        internal void Preview(long version, string address, ConversionPreview value, string error)
        { Queue(delegate(CellConversionForm f) { f.ReceivePreview(version, address, value, error); }); }
        internal void Activate() { Queue(delegate(CellConversionForm f) { if (!HiddenForProbe && !closing) { if (f.WindowState == FormWindowState.Minimized) f.WindowState = FormWindowState.Normal; f.Activate(); } }); }
        internal void Close() { lock (sync) closing = true; Queue(delegate(CellConversionForm f) { f.Close(); }); }
        internal static void CloseAll() { ConversionWindowThread[] copy; lock (registry) copy = windows.ToArray(); foreach (var window in copy) window.Close(); }
    }
}
