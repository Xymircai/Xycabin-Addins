using System;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal static class ConversionInteraction
    {
        private sealed class Session
        {
            internal long Ticket;
            internal CellConversionKind Kind;
            internal IConversionSessionHost Host;
            internal object Application;
            internal ConversionWindowThread Window;
        }
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
            try
            {
                dispatcher = new Control(); var handle = dispatcher.Handle; thread = Thread.CurrentThread.ManagedThreadId;
                Connect.WorksheetUndoChanged += RefreshUndo;
            }
            finally { SynchronizationContext.SetSynchronizationContext(previous); }
        }
        private static void AssertThread() { if (thread != Thread.CurrentThread.ManagedThreadId) throw new InvalidOperationException("转换目标只能在原宿主线程访问。"); }
        internal static string UndoSource(CellConversionKind kind) { return "conversion/" + kind; }
        private static void RefreshUndo()
        {
            var session = current;
            if (session != null) session.Window.SetUndo(Connect.UndoState(UndoSource(session.Kind)));
        }
        internal static bool ActivatePending(CellConversionKind kind)
        { if (current == null || current.Kind != kind) return false; AssertThread(); current.Window.Activate(); return true; }
        internal static long Open(IConversionSessionHost host, CellConversionKind kind, IntPtr owner, object application = null, string contextError = null)
        {
            bool transferred = false;
            try
            {
                EnsureDispatcher(); if (ActivatePending(kind)) return current.Ticket;
                CloseSession();
                ConversionPreview preview = null; string previewError = contextError;
                if (host != null)
                    try { preview = host.ReadPreview(host.Address); } catch (Exception error) { previewError = error.Message; }
                var session = new Session { Ticket = ++sequence, Kind = kind, Host = host, Application = application };
                session.Window = new ConversionWindowThread(session.Ticket, kind, owner,
                    host == null ? "当前没有可用的单元格选区。" : host.Description, host == null ? null : host.Address, preview, previewError);
                current = session; transferred = true;
                RefreshUndo();
                try { session.Window.Start(); } catch { current = null; transferred = false; throw; }
                return session.Ticket;
            }
            finally { if (!transferred && host != null) host.Dispose(); }
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
                    AssertThread(); var session = current; if (session == null || session.Ticket != ticket || session.Host == null) return;
                    if (!Connect.TryBeginConversion()) { session.Window.Preview(version, address, null, "其他操作正在执行，请稍后重新选择区域。"); return; }
                    try { session.Window.Preview(version, address, session.Host.ReadPreview(address), null); }
                    catch (Exception error) { session.Window.Preview(version, address, null, error.Message); }
                    finally { Connect.EndConversion(); }
                });
            }
            catch (InvalidOperationException) { }
        }
        internal static void ReceiveUndo(long ticket, long id)
        {
            var target = dispatcher; if (target == null || target.IsDisposed) return;
            try
            {
                target.BeginInvoke((MethodInvoker)delegate
                {
                    AssertThread(); var session = current;
                    if (session == null || session.Ticket != ticket) return;
                    string message;
                    try
                    {
                        message = Connect.UndoWorksheetEdit(UndoSource(session.Kind), id, session.Application, Connect.ConversionRecoveryDirectory);
                        ClearSnapshot(session);
                    }
                    catch (Exception error) { message = error.Message; }
                    session.Window.Resume(null, null, message, true);
                });
            }
            catch (InvalidOperationException) { }
        }

        private static void ClearSnapshot(Session session)
        {
            var host = session.Host as ComConversionHost;
            if (host != null) host.ClearCells();
        }
        private static void Complete(long ticket, bool pick, ConversionRequest request)
        {
            AssertThread(); var session = current; if (session == null || session.Ticket != ticket) return;
            if (!pick && request == null)
            {
                CloseSession();
                return;
            }
            if (session.Host == null || (request != null && request.Kind != session.Kind)) return;
            if (!Connect.TryBeginConversion()) { session.Window.Resume(null, "其他操作正在执行，请稍后重试。"); return; }
            try
            {
                if (pick)
                {
                    if (request != null) session.Host.SetAddress(request.Address);
                    string address = session.Host.PickRange();
                    if (address != null) session.Host.SetAddress(address);
                    session.Window.Resume(address, null); return;
                }
                session.Host.SetAddress(request.Address);
                ClearSnapshot(session);
                var result = CellConversionEngine.Run(session.Host, request, RecoveryForProbe ?? Connect.ConversionRecoveryDirectory,
                    delegate(BatchSyntaxReport report)
                    {
                        if (dispatcher == null || dispatcher.IsDisposed || !session.Window.IsAlive) return false;
                        return BatchTextEditConfirmation.Show(report) && dispatcher != null && !dispatcher.IsDisposed && session.Window.IsAlive;
                    }, recordUndo: ResultForProbe == null, ownsHost: false);
                if (ResultForProbe != null) ResultForProbe(result);
                else Connect.RememberWorksheetEdit(UndoSource(session.Kind), result.Undo);
                session.Window.Resume(null, null, result.Details, true);
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("conversion-session", error);
                session.Window.Resume(null, error.Message);
                if (!pick && ResultForProbe != null) ResultForProbe(new CellConversionResult { Failed = true, Details = error.Message });
            }
            finally { Connect.EndConversion(); }
        }
        private static void CloseSession()
        {
            var session = current; current = null;
            if (session == null) return;
            try { if (session.Host != null) session.Host.Dispose(); }
            catch (Exception error) { UiDiagnostics.Record("conversion-cancel-release", error); }
            finally { session.Window.Close(); }
        }
        internal static void Shutdown()
        {
            if (dispatcher != null) AssertThread();
            Connect.WorksheetUndoChanged -= RefreshUndo;
            ConversionWindowThread.CloseAll();
            try { CloseSession(); }
            finally { var old = dispatcher; dispatcher = null; thread = 0; if (old != null) old.Dispose(); }
        }
    }
}
