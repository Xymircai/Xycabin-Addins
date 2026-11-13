using System;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal static class ConversionInteraction
    {
        private sealed class Session { internal long Ticket; internal IConversionSessionHost Host; internal ConversionWindowThread Window; }
        private static Control dispatcher;
        private static int thread;
        private static long sequence;
        private static Session current;
        internal static Action<CellConversionResult> ResultForProbe;
        internal static string RecoveryForProbe;
        internal static bool HasPending { get { return current != null; } }
        private static void EnsureDispatcher()
        {
            if (dispatcher != null) { AssertThread(); return; }
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("转换命令必须在宿主 STA 打开。");
            var previous = SynchronizationContext.Current;
            try { dispatcher = new Control(); var handle = dispatcher.Handle; thread = Thread.CurrentThread.ManagedThreadId; }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        private static void AssertThread() { if (thread != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("转换目标只能在原宿主线程访问。"); }
        internal static bool ActivatePending() { if (current == null) return false; AssertThread(); current.Window.Activate(); return true; }
        internal static long Open(IConversionSessionHost host, CellConversionKind kind, IntPtr owner)
        {
            bool transferred = false;
            try
            {
                EnsureDispatcher(); if (ActivatePending()) return current.Ticket;
                ConversionPreview preview = null; string previewError = null;
                try { preview = host.ReadPreview(host.Address); } catch (Exception error) { previewError = error.Message; }
                var session = new Session { Ticket = ++sequence, Host = host };
                session.Window = new ConversionWindowThread(session.Ticket, kind, owner, host.Description, host.Address, preview, previewError);
                current = session; transferred = true;
                try { session.Window.Start(); } catch { current = null; transferred = false; throw; }
                return session.Ticket;
            }
            finally { if (!transferred) host.Dispose(); }
        }
        internal static void Receive(long ticket, bool pick, ConversionRequest request)
        {
            var target = dispatcher; if (target == null || target.IsDisposed) return;
            try { target.BeginInvoke((MethodInvoker)delegate { Complete(ticket, pick, request); }); } catch (InvalidOperationException) { }
        }
        internal static void ReadPreview(long ticket, long version, string address)
        {
            var target = dispatcher; if (target == null || target.IsDisposed) return;
            try
            {
                target.BeginInvoke((MethodInvoker)delegate
                {
                    AssertThread(); var session = current; if (session == null || session.Ticket != ticket) return;
                    if (!Connect.TryBeginConversion()) { session.Window.Preview(version, address, null, "其他操作正在执行，请稍后重新选择区域。"); return; }
                    try { session.Window.Preview(version, address, session.Host.ReadPreview(address), null); }
                    catch (Exception error) { session.Window.Preview(version, address, null, error.Message); }
                    finally { Connect.EndConversion(); }
                });
            }
            catch (InvalidOperationException) { }
        }
        private static void Complete(long ticket, bool pick, ConversionRequest request)
        {
            AssertThread(); var session = current; if (session == null || session.Ticket != ticket) return;
            if (!pick && request == null)
            {
                current = null;
                try { session.Host.Dispose(); } catch (Exception error) { UiDiagnostics.Record("conversion-cancel-release", error); }
                finally { session.Window.Close(); }
                return;
            }
            if (!Connect.TryBeginConversion()) { session.Window.Resume(null, "其他操作正在执行，请稍后重试。"); return; }
            bool consumed = false;
            try
            {
                if (pick)
                {
                    if (request != null) session.Host.SetAddress(request.Address);
                    string address = session.Host.PickRange();
                    if (address != null) session.Host.SetAddress(address);
                    session.Window.Resume(address, null); return;
                }
                session.Host.SetAddress(request.Address); current = null; consumed = true;
                var result = CellConversionEngine.Run(session.Host, request, RecoveryForProbe ?? Connect.ConversionRecoveryDirectory,
                    delegate(BatchSyntaxReport report)
                    {
                        if (dispatcher == null || dispatcher.IsDisposed || !session.Window.IsAlive) return false;
                        return BatchTextEditConfirmation.Show(report) && dispatcher != null && !dispatcher.IsDisposed && session.Window.IsAlive;
                    });
                if (ResultForProbe != null) ResultForProbe(result); else Connect.PresentConversionResult(result);
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("conversion-session", error);
                if (!consumed) session.Window.Resume(null, error.Message);
                else if (ResultForProbe != null) ResultForProbe(new CellConversionResult { Failed = true, Details = error.Message });
                else UiDialog.Show(error.Message, "转换未完成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { Connect.EndConversion(); if (consumed) session.Window.Close(); }
        }
        internal static void Shutdown()
        {
            if (dispatcher != null) AssertThread();
            var session = current; current = null; ConversionWindowThread.CloseAll();
            try { if (session != null) session.Host.Dispose(); }
            finally { var old = dispatcher; dispatcher = null; thread = 0; if (old != null) old.Dispose(); }
        }
    }
}
