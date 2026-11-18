using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;

namespace Excel_Iterate
{
    internal static class MigrationComBridge
    {
        internal const string DesktopPrefix = "ExcelIterate-Migration-";
        internal const int StartupMilliseconds = 20000;
        private const int ProtocolVersion = 1, MaxPacket = 65536;
        private static readonly Guid DispatchId = new Guid("00020400-0000-0000-C000-000000000046");

        internal static string PrivateAutomationDesktop()
        {
            string desktop = DesktopName(GetCurrentThreadId());
            Guid id;
            return desktop.StartsWith(DesktopPrefix, StringComparison.Ordinal) &&
                Guid.TryParseExact(desktop.Substring(DesktopPrefix.Length), "N", out id) &&
                InPrivateJob(GetCurrentProcessId(), desktop) ? desktop : null;
        }

        internal static string PipeName(string desktop) { return desktop + "-COM"; }

        internal static NamedPipeServerStream CreatePipe(string name)
        {
            var security = new PipeSecurity();
            using (var identity = WindowsIdentity.GetCurrent())
                security.AddAccessRule(new PipeAccessRule(identity.User, PipeAccessRights.FullControl, AccessControlType.Allow));
            return new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous, 4096, 4096, security);
        }

        internal static object Receive(NamedPipeServerStream pipe, IAsyncResult connection, Func<uint, bool> ownsProcess,
            CancellationToken cancellation, int timeout = StartupMilliseconds)
        {
            using (var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellation))
            using (deadline.Token.Register(pipe.Dispose))
            {
                deadline.CancelAfter(timeout);
                try
                {
                    pipe.EndWaitForConnection(connection);
                    deadline.Token.ThrowIfCancellationRequested();
                    uint client;
                    if (!GetNamedPipeClientProcessId(pipe.SafePipeHandle, out client) || !ownsProcess(client))
                        throw new IOException("后台 WPS 的管道进程不属于本次会话。");
                    using (var reader = new BinaryReader(pipe, Encoding.UTF8, true))
                    {
                        if (reader.ReadInt32() != ProtocolVersion)
                            throw new InvalidDataException("后台 WPS 加载项的连接协议与当前版本不一致。");
                        int length = reader.ReadInt32();
                        if (length <= 0 || length > MaxPacket)
                            throw new InvalidDataException("后台 WPS 接口长度无效。");
                        byte[] bytes = reader.ReadBytes(length);
                        if (bytes.Length != length) throw new EndOfStreamException("后台 WPS 接口传输不完整。");
                        deadline.Token.ThrowIfCancellationRequested();
                        object app = Unmarshal(bytes);
                        try { pipe.WriteByte(1); pipe.Flush(); }
                        catch { Marshal.ReleaseComObject(app); throw; }
                        Connect.Log("复制后台 WPS 接口已交接，PID=" + client);
                        return app;
                    }
                }
                catch (Exception error)
                {
                    if (cancellation.IsCancellationRequested)
                        throw new OperationCanceledException("已取消后台 WPS 连接。", error, cancellation);
                    if (deadline.IsCancellationRequested)
                        throw new TimeoutException("后台 WPS 未能及时从自身加载项回传自动化接口。", error);
                    throw;
                }
            }
        }

        // Called by the VSTO instance in the private WPS process. Only the
        // host STA touches the Application and the COM marshal lease.
        internal sealed class HostedBinding : IDisposable
        {
            private readonly string desktop;
            private readonly Func<object> readApplication;
            private readonly Stopwatch clock = Stopwatch.StartNew();
            private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 100 };
            private NamedPipeClientStream pipe;
            private Task transmission;
            private IStream lease;
            private bool disposed;
            private string lastComError;

            internal HostedBinding(string desktop, Func<object> readApplication)
            {
                this.desktop = desktop;
                this.readApplication = readApplication;
                timer.Tick += Tick;
                timer.Start();
            }

            private void Tick(object sender, EventArgs args)
            {
                timer.Stop();
                try
                {
                    if (transmission != null && transmission.IsCompleted) { Dispose(); return; }
                    if (clock.ElapsedMilliseconds >= StartupMilliseconds)
                    {
                        Connect.Log("后台 WPS 等待接口交接超时，桌面=" + desktop + "，最近 COM 错误=" + lastComError);
                        Dispose();
                        return;
                    }
                    if (transmission != null) return;
                    object app = readApplication();
                    if (app == null) return;
                    try
                    {
                        var window = new IntPtr(Convert.ToInt64(((dynamic)app).Hwnd));
                        if (window == IntPtr.Zero) return;
                        uint process;
                        uint thread = GetWindowThreadProcessId(window, out process);
                        if (!string.Equals(DesktopName(thread), desktop, StringComparison.OrdinalIgnoreCase) || !InPrivateJob(process, desktop))
                            throw new InvalidOperationException("后台 WPS 加载项取得的 Application 属于其他窗口。");

                        byte[] packet = MarshalApplication(app, out lease);
                        int remaining = Math.Max(1, StartupMilliseconds - (int)clock.ElapsedMilliseconds);
                        pipe = new NamedPipeClientStream(".", PipeName(desktop), PipeDirection.InOut, PipeOptions.Asynchronous, TokenImpersonationLevel.Identification);
                        NamedPipeClientStream connection = pipe;
                        transmission = Task.Run(() => Send(connection, packet, remaining));
                    }
                    finally { if (Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); }
                }
                catch (COMException error) { lastComError = error.HResult.ToString("X8"); }
                catch (Exception error)
                {
                    Connect.Log("后台 WPS 获取自身接口失败，桌面=" + desktop + "，" + error);
                    Dispose();
                }
                finally { if (!disposed) timer.Start(); }
            }

            private void Send(NamedPipeClientStream connection, byte[] packet, int timeout)
            {
                try
                {
                    using (connection)
                    using (var writer = new BinaryWriter(connection, Encoding.UTF8, true))
                    {
                        connection.Connect(timeout);
                        writer.Write(ProtocolVersion);
                        writer.Write(packet.Length);
                        writer.Write(packet);
                        writer.Flush();
                        if (connection.ReadByte() != 1) throw new IOException("未收到后台 WPS 接口确认。");
                    }
                }
                catch (Exception error)
                {
                    Connect.Log("后台 WPS 回传接口失败，桌面=" + desktop + "，" + error);
                }
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                timer.Stop();
                timer.Dispose();
                if (pipe != null) pipe.Dispose();
                if (lease == null) return;
                try { ReleaseMarshalLease(lease); }
                catch (Exception error) { Connect.Log("释放后台 WPS 接口封送数据：" + error); }
                finally { lease = null; }
            }
        }

        private static byte[] MarshalApplication(object application, out IStream lease)
        {
            lease = null;
            IStream stream;
            Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
            IntPtr unknown = IntPtr.Zero;
            bool registered = false;
            try
            {
                unknown = Marshal.GetIUnknownForObject(application);
                var iid = DispatchId;
                // A table-strong packet is released by this STA after the
                // receiver acknowledges its own proxy or the session ends.
                Marshal.ThrowExceptionForHR(CoMarshalInterface(stream, ref iid, unknown, 0, IntPtr.Zero, 1));
                registered = true;
                System.Runtime.InteropServices.ComTypes.STATSTG stat;
                stream.Stat(out stat, 1);
                if (stat.cbSize <= 0 || stat.cbSize > MaxPacket) throw new IOException("后台 WPS 接口长度无效。");
                var bytes = new byte[(int)stat.cbSize];
                stream.Seek(0, 0, IntPtr.Zero);
                stream.Read(bytes, bytes.Length, IntPtr.Zero);
                lease = stream;
                return bytes;
            }
            finally
            {
                if (unknown != IntPtr.Zero) Marshal.Release(unknown);
                if (lease == null)
                {
                    if (registered) ReleaseMarshalLease(stream);
                    else Marshal.ReleaseComObject(stream);
                }
            }
        }

        private static void ReleaseMarshalLease(IStream stream)
        {
            try { stream.Seek(0, 0, IntPtr.Zero); Marshal.ThrowExceptionForHR(CoReleaseMarshalData(stream)); }
            finally { Marshal.ReleaseComObject(stream); }
        }

        private static object Unmarshal(byte[] bytes)
        {
            IStream stream;
            Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
            try
            {
                stream.Write(bytes, bytes.Length, IntPtr.Zero);
                stream.Seek(0, 0, IntPtr.Zero);
                var iid = DispatchId;
                IntPtr pointer;
                Marshal.ThrowExceptionForHR(CoUnmarshalInterface(stream, ref iid, out pointer));
                try { return Marshal.GetObjectForIUnknown(pointer); }
                finally { Marshal.Release(pointer); }
            }
            finally { Marshal.ReleaseComObject(stream); }
        }

        private static bool InPrivateJob(uint pid, string desktop)
        {
            IntPtr job = OpenJobObject(4, false, @"Local\" + desktop);
            if (job == IntPtr.Zero) return false;
            try
            {
                IntPtr process = OpenProcess(0x1000, false, pid);
                if (process == IntPtr.Zero) return false;
                try { bool owned; return IsProcessInJob(process, job, out owned) && owned; }
                finally { CloseHandle(process); }
            }
            finally { CloseHandle(job); }
        }

        internal static string DesktopName(uint thread)
        {
            var name = new StringBuilder(128);
            int required;
            if (thread == 0 || !GetUserObjectInformation(GetThreadDesktop(thread), 2, name, name.Capacity * 2, out required)) return "";
            return name.ToString();
        }

        [DllImport("kernel32.dll")] private static extern bool GetNamedPipeClientProcessId(SafePipeHandle pipe, out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr OpenJobObject(int access, bool inherit, string name);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(int access, bool inherit, uint pid);
        [DllImport("kernel32.dll")] private static extern bool IsProcessInJob(IntPtr process, IntPtr job, out bool result);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll")] private static extern IntPtr GetThreadDesktop(uint thread);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool GetUserObjectInformation(IntPtr handle, int index, StringBuilder value, int size, out int required);
        [DllImport("ole32.dll")] private static extern int CreateStreamOnHGlobal(IntPtr memory, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease, out IStream stream);
        [DllImport("ole32.dll")] private static extern int CoMarshalInterface(IStream stream, ref Guid iid, IntPtr unknown, int context, IntPtr target, int flags);
        [DllImport("ole32.dll")] private static extern int CoReleaseMarshalData(IStream stream);
        [DllImport("ole32.dll")] private static extern int CoUnmarshalInterface(IStream stream, ref Guid iid, out IntPtr pointer);
    }
}
