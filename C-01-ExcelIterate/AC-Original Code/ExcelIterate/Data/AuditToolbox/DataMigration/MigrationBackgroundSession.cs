using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Excel_Iterate
{
    internal sealed class MigrationBackgroundUnavailableException : InvalidOperationException
    {
        internal MigrationBackgroundUnavailableException(Exception inner)
            : base("无法启动独立的 WPS 后台会话：" + inner.Message, inner) { }
    }

    internal sealed class MigrationBackgroundSession
    {
        private const int StartupMilliseconds = MigrationComBridge.StartupMilliseconds, CancelGraceMilliseconds = 10000;
        private readonly string desktopName = MigrationComBridge.DesktopPrefix + Guid.NewGuid().ToString("N");
        private IntPtr desktop, job;
        private volatile bool ready;
        private volatile string startupStep = "查找 WPS 表格自动化程序";
        private volatile MigrationProgress progress = new MigrationProgress("正在启动独立的 WPS 后台会话…", true);

        // The WinForms timer consumes only progress; the same iterator can still
        // be pumped on the host STA for an explicitly confirmed foreground run.
        internal static IEnumerable<MigrationProgress> Run(Func<object, IEnumerable<MigrationProgress>> operation, CancellationToken cancellation)
        {
            var session = new MigrationBackgroundSession();
            var task = Task.Factory.StartNew(delegate { session.Work(operation, cancellation); },
                CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
            while (!task.IsCompleted) yield return session.progress;
            task.GetAwaiter().GetResult();
        }

        private void Work(Func<object, IEnumerable<MigrationProgress>> operation, CancellationToken cancellation)
        {
            // This task owns the handles, including when the UI stops observing.
            try
            {
                cancellation.ThrowIfCancellationRequested();
                var server = RegisteredServer();
                startupStep = "创建后台会话";
                desktop = CreateDesktop(desktopName, IntPtr.Zero, IntPtr.Zero, 0, 0x000F01FF, IntPtr.Zero);
                Require(desktop != IntPtr.Zero, "创建私有桌面");
                job = CreateJobObject(IntPtr.Zero, @"Local\" + desktopName);
                Require(job != IntPtr.Zero, "创建后台进程作业");
                var limits = new ExtendedLimits();
                limits.Basic.Flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
                Require(SetInformationJobObject(job, 9, ref limits, Marshal.SizeOf(typeof(ExtendedLimits))), "设置后台进程退出策略");
                startupStep = "创建后台连接管道";
                using (var pipe = MigrationComBridge.CreatePipe(MigrationComBridge.PipeName(desktopName)))
                {
                    var connection = pipe.BeginWaitForConnection(null, null);
                    cancellation.ThrowIfCancellationRequested();
                    startupStep = "启动后台 WPS";
                    int serverPid = StartProcess(server.Executable, server.Command + " -Embedding");
                    using (var caller = Process.GetCurrentProcess())
                        Connect.Log("复制后台 WPS 已启动，PID=" + serverPid + "，桌面=" + desktopName
                            + "，调用方=" + caller.ProcessName + "/" + caller.Id);
                    startupStep = "创建后台自动化线程";
                    RunSta(() => MigrationComBridge.Receive(pipe, connection, pid => OwnsRegisteredProcess(pid, server.Executable), cancellation),
                        operation, cancellation);
                }
            }
            catch (Exception error)
            {
                if (cancellation.IsCancellationRequested) throw new OperationCanceledException(cancellation);
                if (!ready) throw new MigrationBackgroundUnavailableException(new InvalidOperationException(startupStep + "：" + error.Message, error));
                throw;
            }
            finally { Dispose(); }
        }

        private int StartProcess(string executable, string command)
        {
            var startup = new Startup { Size = Marshal.SizeOf(typeof(Startup)), Desktop = desktopName, Flags = 0x80 };
            ProcessInfo child;
            // Join the job before the child can run. WPS starts directly
            // on the private desktop because SCM activation can reuse a live host.
            Require(CreateProcess(executable, new StringBuilder(command), IntPtr.Zero, IntPtr.Zero,
                false, 0x08000004, IntPtr.Zero, Path.GetDirectoryName(executable), ref startup, out child), "启动后台进程");
            try
            {
                if (!AssignProcessToJobObject(job, child.Process))
                {
                    var error = NativeFailure("将进程加入后台作业");
                    TerminateProcess(child.Process, 1); // Still suspended and known to be ours.
                    throw error;
                }
                Require(ResumeThread(child.Thread) != uint.MaxValue, "启动后台进程线程");
                return child.Pid;
            }
            finally { CloseHandle(child.Thread); CloseHandle(child.Process); }
        }

        private void RunSta(Func<object> bind, Func<object, IEnumerable<MigrationProgress>> operation, CancellationToken cancellation)
        {
            Exception failure = null;
            // A managed STA Thread may already own COM windows when its
            // delegate starts. Bind this thread's desktop before COM init.
            ThreadStart body = delegate(IntPtr unused)
            {
                try
                {
                    Require(SetThreadDesktop(desktop), "绑定后台线程的私有桌面");
                    Marshal.ThrowExceptionForHR(CoInitializeEx(IntPtr.Zero, 2));
                    try
                    {
                        IOleMessageFilter previous;
                        Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(new BusyMessageFilter(cancellation), out previous));
                        try { Execute(bind, operation, cancellation); }
                        finally
                        {
                            IOleMessageFilter removed;
                            CoRegisterMessageFilter(previous, out removed);
                            if (previous != null && Marshal.IsComObject(previous)) Marshal.ReleaseComObject(previous);
                        }
                    }
                    finally { CoUninitialize(); }
                }
                catch (Exception error) { failure = error; }
                return 0;
            };
            IntPtr thread = CreateThread(IntPtr.Zero, UIntPtr.Zero, body, IntPtr.Zero, 0, IntPtr.Zero);
            Require(thread != IntPtr.Zero, "创建后台 STA 线程");
            var clock = Stopwatch.StartNew();
            long cancelledAt = -1;
            bool stopped = false, timedOut = false;
            try
            {
                uint wait;
                while ((wait = WaitForSingleObject(thread, 100)) == 258)
                {
                    if (cancellation.IsCancellationRequested && cancelledAt < 0) cancelledAt = clock.ElapsedMilliseconds;
                    if (stopped) continue;
                    timedOut = !ready && clock.ElapsedMilliseconds >= StartupMilliseconds;
                    if (timedOut || (cancelledAt >= 0 && clock.ElapsedMilliseconds - cancelledAt >= CancelGraceMilliseconds))
                    {
                        // Target files are committed by the managed worker only
                        // after saving and verifying the closed stage. Stop just
                        // the WPS server to unblock a stuck COM call.
                        stopped = true;
                        TerminateJobObject(job, 1);
                    }
                }
                Require(wait == 0, "等待后台 STA 线程");
                cancellation.ThrowIfCancellationRequested();
                if (timedOut) throw new TimeoutException("后台 WPS 未在 20 秒内就绪，阶段：" + startupStep + "。");
                if (failure != null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
            }
            finally { GC.KeepAlive(body); CloseHandle(thread); }
        }

        private void Execute(Func<object> bind, Func<object, IEnumerable<MigrationProgress>> operation, CancellationToken cancellation)
        {
            using (var references = new ComBag())
            {
                startupStep = "等待后台 WPS 加载项回传自动化接口";
                cancellation.ThrowIfCancellationRequested();
                object application = references.Own(bind());
                dynamic app = application;
                if (!OwnsWindow(new IntPtr(Convert.ToInt64(app.Hwnd))))
                    throw new InvalidOperationException("自动化实例的窗口和进程不属于本次后台会话。");
                startupStep = "检查后台工作簿";
                dynamic books = references.Own(app.Workbooks);
                int count = Convert.ToInt32(books.Count);
                if (count > 1) throw new InvalidOperationException("后台 WPS 启动时打开了其他工作簿。");
                dynamic parking = count == 1 ? references.Own(books.Item(1)) : null;
                if (parking != null && !string.IsNullOrEmpty(Convert.ToString(parking.Path)))
                    throw new InvalidOperationException("后台 WPS 启动时打开了磁盘上的工作簿。");
                try
                {
                    // WPS needs one visible view even for automation. This
                    // unsaved book only exists on the inactive desktop.
                    app.DisplayAlerts = false;
                    app.EnableEvents = false;
                    app.ScreenUpdating = false;
                    app.Visible = true;
                    if (parking == null) parking = references.Own(books.Add());
                    startupStep = "核验后台可见工作簿";
                    dynamic windows = references.Own(parking.Windows), window = references.Own(windows.Item(1));
                    if (!OwnsWindow(new IntPtr(Convert.ToInt64(window.Hwnd))) || !ComFormulaHost.Boolean(window.Visible, "Visible"))
                        throw new InvalidOperationException("后台工作簿窗口归属或可见状态无法确认。");
                    cancellation.ThrowIfCancellationRequested();
                    ready = true;
                    Connect.Log("复制后台 WPS 已就绪，私有桌面=" + desktopName);
                    var pump = Stopwatch.StartNew();
                    var report = Stopwatch.StartNew();
                    foreach (var step in operation(application))
                    {
                        if (step.Pause || report.ElapsedMilliseconds >= 200)
                        { progress = new MigrationProgress(step.Message, true); report.Restart(); }
                        if (step.Pause || pump.ElapsedMilliseconds >= 40)
                        {
                            PumpAutomationMessages();
                            pump.Restart();
                        }
                    }
                }
                finally
                {
                    try { if (parking != null) parking.Close(false); }
                    catch (Exception error) { Connect.Log("关闭后台占位工作簿：" + error.HResult.ToString("X8")); }
                    try { app.Quit(); }
                    catch (Exception error) { Connect.Log("后台 WPS 正常退出未响应：" + error.HResult.ToString("X8")); }
                }
            }
        }

        private bool OwnsWindow(IntPtr window)
        {
            if (window == IntPtr.Zero) return false;
            IntPtr root = GetAncestor(window, 2), owner = GetAncestor(window, 3);
            return OwnsWindowThread(window) && (root == window || OwnsWindowThread(root))
                && (owner == root || owner == window || OwnsWindowThread(owner));
        }

        private bool OwnsWindowThread(IntPtr window)
        {
            uint pid; uint thread = GetWindowThreadProcessId(window, out pid);
            if (thread == 0 || pid == 0 || !string.Equals(MigrationComBridge.DesktopName(thread), desktopName, StringComparison.OrdinalIgnoreCase))
                return false;
            IntPtr process = OpenProcess(0x1000, false, pid);
            if (process == IntPtr.Zero) return false;
            try { bool owned; return IsProcessInJob(process, job, out owned) && owned; }
            finally { CloseHandle(process); }
        }

        private bool OwnsRegisteredProcess(uint pid, string executable)
        {
            IntPtr process = OpenProcess(0x1000, false, pid);
            if (process == IntPtr.Zero) return false;
            try
            {
                bool owned;
                if (!IsProcessInJob(process, job, out owned) || !owned) return false;
                var path = new StringBuilder(32768); int length = path.Capacity;
                if (!QueryFullProcessImageName(process, 0, path, ref length)) return false;
                string image = path.ToString();
                // The registered wps.exe launcher may run the spreadsheet
                // host as et.exe under the same installation and job.
                return string.Equals(image, executable, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(image, Path.Combine(Path.GetDirectoryName(executable), "et.exe"), StringComparison.OrdinalIgnoreCase);
            }
            finally { CloseHandle(process); }
        }

        private sealed class Server
        {
            internal string Command, Executable;
        }

        private static Server RegisteredServer()
        {
            foreach (string id in new[] { "KET.Application", "et.Application" })
            {
                var type = Type.GetTypeFromProgID(id, false);
                if (type == null) continue;
                string command;
                using (var key = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + type.GUID.ToString("B") + @"\LocalServer32"))
                    command = key == null ? null : key.GetValue(null) as string;
                if (string.IsNullOrWhiteSpace(command)) continue;
                command = command.Trim();
                int end = command[0] == '"' ? command.IndexOf('"', 1) : command.IndexOf(' ');
                string executable = (end > 0 ? command.Substring(0, end) : command).Trim('"');
                if (Path.IsPathRooted(executable) && File.Exists(executable))
                    return new Server { Command = command, Executable = executable };
            }
            throw new InvalidOperationException("没有找到已注册的 WPS 表格自动化程序。");
        }

        private static void Require(bool success, string operation)
        {
            if (!success) throw NativeFailure(operation);
        }

        private static Win32Exception NativeFailure(string operation)
        {
            int code = Marshal.GetLastWin32Error();
            return new Win32Exception(code, operation + "失败（Win32 " + code + "） " + new Win32Exception(code).Message);
        }

        private static void PumpAutomationMessages()
        {
            // Dispatch only this private STA's pending native messages. Limit
            // each pass so queued messages cannot starve the workbook iterator.
            Message message;
            for (int i = 0; i < 64 && PeekMessage(out message, IntPtr.Zero, 0, 0, 1); i++)
            {
                TranslateMessage(ref message);
                DispatchMessage(ref message);
            }
        }

        private void Dispose()
        {
            if (job != IntPtr.Zero) { CloseHandle(job); job = IntPtr.Zero; }
            if (desktop != IntPtr.Zero) { CloseDesktop(desktop); desktop = IntPtr.Zero; }
        }

        private delegate uint ThreadStart(IntPtr data);
        [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleMessageFilter
        {
            [PreserveSig] int HandleInComingCall(int kind, IntPtr caller, int elapsed, IntPtr info);
            [PreserveSig] int RetryRejectedCall(IntPtr callee, int elapsed, int reason);
            [PreserveSig] int MessagePending(IntPtr callee, int elapsed, int kind);
        }
        private sealed class BusyMessageFilter : IOleMessageFilter
        {
            private readonly CancellationToken cancellation;
            internal BusyMessageFilter(CancellationToken cancellation) { this.cancellation = cancellation; }
            public int HandleInComingCall(int kind, IntPtr caller, int elapsed, IntPtr info) { return 0; }
            public int RetryRejectedCall(IntPtr callee, int elapsed, int reason)
            { return !cancellation.IsCancellationRequested && elapsed < 3000 ? 100 : -1; }
            public int MessagePending(IntPtr callee, int elapsed, int kind) { return 2; }
        }
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr mode, int flags, int access, IntPtr attributes);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetThreadDesktop(IntPtr desktop);
        [DllImport("user32.dll")] private static extern bool CloseDesktop(IntPtr desktop);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint kind);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool PeekMessage(out Message message, IntPtr window, uint first, uint last, uint remove);
        [DllImport("user32.dll")] private static extern bool TranslateMessage(ref Message message);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr DispatchMessage(ref Message message);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern IntPtr CreateJobObject(IntPtr options, string name);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool SetInformationJobObject(IntPtr job, int type, ref ExtendedLimits limits, int size);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
        [DllImport("kernel32.dll")] private static extern bool IsProcessInJob(IntPtr process, IntPtr job, out bool result);
        [DllImport("kernel32.dll")] private static extern bool TerminateJobObject(IntPtr job, uint code);
        [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(int access, bool inherit, uint pid);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern bool QueryFullProcessImageName(IntPtr process, int flags, StringBuilder path, ref int size);
        [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint ResumeThread(IntPtr thread);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr CreateThread(IntPtr attributes, UIntPtr stackSize, ThreadStart start, IntPtr data, int flags, IntPtr threadId);
        [DllImport("kernel32.dll")] private static extern bool TerminateProcess(IntPtr process, int code);
        [DllImport("kernel32.dll", SetLastError = true)] private static extern uint WaitForSingleObject(IntPtr handle, uint milliseconds);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CreateProcess(string app, StringBuilder command, IntPtr processAttributes, IntPtr threadAttributes, bool inherit, int flags, IntPtr environment, string directory, ref Startup startup, out ProcessInfo process);
        [DllImport("ole32.dll")] private static extern int CoInitializeEx(IntPtr reserved, uint options);
        [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter filter, out IOleMessageFilter previous);
        [DllImport("ole32.dll")] private static extern void CoUninitialize();

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] private struct Startup
        {
            public int Size;
            public string Reserved, Desktop, Title;
            public int X, Y, Width, Height, CountX, CountY, Fill, Flags;
            public short Show, ReservedSize;
            public IntPtr ReservedData, Input, Output, Error;
        }
        [StructLayout(LayoutKind.Sequential)] private struct ProcessInfo { public IntPtr Process, Thread; public int Pid, Tid; }
        [StructLayout(LayoutKind.Sequential)] private struct Message
        {
            public IntPtr Window;
            public uint Id;
            public UIntPtr WParam;
            public IntPtr LParam;
            public uint Time;
            public int X, Y;
            public uint Private;
        }
        [StructLayout(LayoutKind.Sequential)] private struct Limits
        {
            public long ProcessTime, JobTime;
            public uint Flags;
            public UIntPtr MinWorkingSet, MaxWorkingSet;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint Priority, Scheduling;
        }
        [StructLayout(LayoutKind.Sequential)] private struct ExtendedLimits
        {
            public Limits Basic;
            public ulong ReadOperations, WriteOperations, OtherOperations, ReadBytes, WriteBytes, OtherBytes;
            public UIntPtr ProcessMemory, JobMemory, PeakProcessMemory, PeakJobMemory;
        }
    }
}
