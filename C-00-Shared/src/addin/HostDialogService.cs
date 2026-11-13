using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Synchronous confirmations still return a result to the original business
    // call. Only their pure UI runs on another STA. No WinForms modal loop runs
    // on the native WPS thread, and no COM object enters the UI callback.
    internal static class HostDialogService
    {
        [ThreadStatic] private static bool registered;
        [ThreadStatic] private static int waiting;
        [DllImport("ole32.dll")] private static extern int CoWaitForMultipleHandles(uint flags, uint milliseconds, uint count, IntPtr[] handles, out uint index);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr handle);
        [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr handle, bool enabled);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr handle, uint flags);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr handle, out uint process);
        [DllImport("user32.dll")] private static extern IntPtr GetActiveWindow();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        private delegate bool WindowVisitor(IntPtr handle, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool EnumThreadWindows(uint thread, WindowVisitor visitor, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr handle, uint message, IntPtr wParam, IntPtr lParam);
        private static int workers;
        internal static int ActiveWorkers { get { return Volatile.Read(ref workers); } }
        internal static Func<int> WaitForProbe;

        /// <summary>宿主根窗口当前是否可交互；取不到根窗口时视为正常（true）。</summary>
        internal static bool HostEnabled()
        {
            try
            {
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                if (handle == IntPtr.Zero) return true;
                IntPtr root = ValidateOwner(handle, false);
                return root == IntPtr.Zero || IsWindowEnabled(root);
            }
            catch { return true; }
        }

        /// <summary>
        /// 防御性恢复宿主根窗口：确认对话框走 Run 会临时禁用宿主，finally 恢复；
        /// 若因异常路径或跨线程竞态漏恢复（表现为"关窗后工作表/功能区锁死"），
        /// 在功能窗关闭时兜底。仅在没有 worker 运行时才恢复——有 worker 说明
        /// 正处于正常 modal 等待，不能破坏其禁用语义。
        /// </summary>
        internal static void EnsureHostEnabled()
        {
            if (Volatile.Read(ref workers) != 0) return;
            try
            {
                IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;
                if (handle == IntPtr.Zero) return;
                IntPtr root = ValidateOwner(handle, false);
                if (root != IntPtr.Zero && IsWindow(root) && !IsWindowEnabled(root)) EnableWindow(root, true);
            }
            catch { }
        }
        internal static void Register()
        { registered = true; Ui.HostDialogPresenter = Show; }
        // A modeless pure-UI window owns a separate STA message loop. Callers
        // exchange only thread-safe state; Excel COM remains on the host STA.
        internal sealed class ModelessWindow
        {
            private int closing;
            internal void Close() { Interlocked.Exchange(ref closing, 1); }
            internal bool Closing { get { return Volatile.Read(ref closing) != 0; } }
        }
        internal static ModelessWindow OpenModeless(Func<Form> create)
        {
            var lifetime = new ModelessWindow();
            var thread = new Thread(delegate()
            {
                int workerId = Thread.CurrentThread.ManagedThreadId;
                ThreadExceptionEventHandler onError = delegate(object sender, ThreadExceptionEventArgs args)
                { if (Thread.CurrentThread.ManagedThreadId == workerId) { UiDiagnostics.Record("modeless-progress-event", args.Exception); Application.ExitThread(); } };
                Application.ThreadException += onError;
                try
                {
                    using (var form = create())
                    using (var timer = new System.Windows.Forms.Timer { Interval = 100 })
                    {
                        timer.Tick += delegate { if (lifetime.Closing) form.Close(); };
                        form.Load += delegate { if (lifetime.Closing) form.BeginInvoke((MethodInvoker)delegate { form.Close(); }); };
                        timer.Start();
                        Application.Run(form);
                    }
                }
                catch (Exception error) { UiDiagnostics.Record("modeless-progress", error); }
                finally { Application.ThreadException -= onError; }
            }) { IsBackground = true, Name = "Addin modeless progress UI" };
            thread.SetApartmentState(ApartmentState.STA); thread.Start();
            return lifetime;
        }
        private sealed class Owner : IWin32Window
        {
            private readonly IntPtr handle; internal Owner(IntPtr value) { handle = value; }
            public IntPtr Handle { get { return handle; } }
        }
        private static IntPtr OwnerHandle(IWin32Window owner)
        {
            IntPtr handle = owner == null ? GetActiveWindow() : owner.Handle;
            if (handle == IntPtr.Zero) handle = Process.GetCurrentProcess().MainWindowHandle;
            return ValidateOwner(handle, false);
        }
        // Isolated modal display for on-demand chooser windows built inside feature forms.
        // The dialog is created and shown on the worker STA; the owner is managed by Run.
        internal static void ShowIsolated(IWin32Window owner, Func<Form> create)
        {
            Run<bool>(owner, delegate(IWin32Window uiOwner) { using (var dialog = create()) dialog.ShowDialog(uiOwner); return true; }, false);
        }
        internal static DialogResult ShowIsolatedResult(IWin32Window owner, Func<Form> create)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var dialog = create()) return dialog.ShowDialog(uiOwner);
            }, DialogResult.Cancel);
        }

        // Creation, result extraction and disposal all belong to the UI STA.
        // Callers return only values (for example JSON or paths), never controls
        // or COM objects, and apply the result on their original host thread.
        internal static T ShowIsolatedResult<T>(IWin32Window owner, Func<Form> create, Func<Form,T> read, T cancelled)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var dialog = create())
                    return dialog.ShowDialog(uiOwner) == DialogResult.OK ? read(dialog) : cancelled;
            }, cancelled);
        }

        internal static IntPtr ValidateOwner(IntPtr handle, bool requireVisible = true)
        {
            if (handle == IntPtr.Zero || !IsWindow(handle)) return IntPtr.Zero;
            handle = GetAncestor(handle, 2); uint pid;
            uint thread = GetWindowThreadProcessId(handle, out pid);
            return pid == (uint)Process.GetCurrentProcess().Id && thread == GetCurrentThreadId() &&
                (!requireVisible || IsWindowVisible(handle)) ? handle : IntPtr.Zero;
        }
        internal static T Run<T>(IWin32Window owner, Func<IWin32Window,T> ui, T cancelled)
        {
            if (!registered) return ui(owner);
            if (waiting != 0) return cancelled; // Do not nest confirmations through COM reentrancy.
            IntPtr handle = OwnerHandle(owner);
            bool disabled = handle != IntPtr.Zero && IsWindowEnabled(handle);
            T result = cancelled; Exception failure = null;
            var completed = new ManualResetEvent(false); int eventUsers = 2, aborted = 0;
            var thread = new Thread(delegate()
            {
                Interlocked.Increment(ref workers);
                ThreadExceptionEventHandler onError = delegate(object sender,ThreadExceptionEventArgs args)
                { failure = args.Exception; Interlocked.Exchange(ref aborted,1); CloseCurrentWindows(); };
                Application.ThreadException += onError;
                try
                {
                    using (var cancellation = new System.Windows.Forms.Timer { Interval = 100 })
                    {
                        cancellation.Tick += delegate { if (Volatile.Read(ref aborted) != 0) CloseCurrentWindows(); };
                        cancellation.Start();
                        if (Volatile.Read(ref aborted) == 0) result = ui(handle == IntPtr.Zero ? null : new Owner(handle));
                    }
                }
                catch (Exception error) { failure = error; }
                finally
                {
                    Application.ThreadException -= onError;
                    completed.Set(); if (Interlocked.Decrement(ref eventUsers) == 0) completed.Dispose();
                    Interlocked.Decrement(ref workers);
                }
            }) { IsBackground = true, Name = "Addin isolated confirmation" };
            thread.SetApartmentState(ApartmentState.STA);
            waiting++;
            bool started = false;
            SynchronizationContext previous = SynchronizationContext.Current;
            try
            {
                if (disabled) EnableWindow(handle, false);
                thread.Start(); started = true;
                var handles = new[] { completed.SafeWaitHandle.DangerousGetHandle() };
                while (!completed.WaitOne(0))
                {
                    uint index;
                    // COM-aware native wait dispatches host calls/window messages.
                    // It does not call Application.Run/DoEvents on the WPS STA.
                    int code = WaitForProbe != null ? WaitForProbe() : CoWaitForMultipleHandles(0x1C, 100, 1, handles, out index);
                    if (code < 0 && code != unchecked((int)0x80010115)) Marshal.ThrowExceptionForHR(code);
                }
                if (failure != null) { UiDiagnostics.Record("isolated-dialog", failure); return cancelled; }
                return result;
            }
            catch (Exception error)
            {
                // A broken native wait must cancel the UI task rather than leave
                // an unobserved confirmation accepting a destructive operation.
                Interlocked.Exchange(ref aborted,1); UiDiagnostics.Record("isolated-dialog-wait",error); return cancelled;
            }
            finally
            {
                // Restore only the same HWND we disabled, and never enable an
                // owner which was already disabled before entering this service.
                if (disabled && IsWindow(handle)) EnableWindow(handle, true);
                SynchronizationContext.SetSynchronizationContext(previous); waiting--;
                if (!started) Interlocked.Decrement(ref eventUsers);
                if (Interlocked.Decrement(ref eventUsers) == 0) completed.Dispose();
            }
        }
        private static void CloseCurrentWindows()
        {
            // Only this worker's own transient UI; never enumerate/close WPS.
            EnumThreadWindows(GetCurrentThreadId(), delegate(IntPtr window,IntPtr unused)
            { PostMessage(window,0x10,IntPtr.Zero,IntPtr.Zero); return true; }, IntPtr.Zero);
            Application.ExitThread();
        }
        private static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
        {
            return Run(owner, delegate(IWin32Window uiOwner) { return UiDialog.ShowDirect(uiOwner, text, caption, buttons, icon, defaultButton); }, DialogResult.Cancel);
        }
        internal static string Folder(IWin32Window owner, string description, bool newFolder)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var picker = new FolderBrowserDialog { Description = description, ShowNewFolderButton = newFolder })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? picker.SelectedPath : null;
            }, (string)null);
        }
        internal static string[] OpenFiles(IWin32Window owner)
        { return OpenFiles(owner, "Excel 工作簿 (*.xlsx)|*.xlsx"); }
        internal static string[] OpenFiles(IWin32Window owner, string filter)
        { return OpenFiles(owner, filter, "", true, ""); }
        internal static string[] OpenFiles(IWin32Window owner, string filter, string title, bool multiple, string filename)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var picker = new OpenFileDialog { Filter = filter, Title = title, Multiselect = multiple, CheckFileExists = true, FileName = filename })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? picker.FileNames : new string[0];
            }, new string[0]);
        }
        internal static string SaveFile(IWin32Window owner, string title, string filter, string filename)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var picker = new SaveFileDialog { Title = title, Filter = filter, FileName = filename, OverwritePrompt = true })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? picker.FileName : null;
            }, (string)null);
        }
        internal static int? PdfColor(IWin32Window owner, int color)
        {
            return Run<int?>(owner, delegate(IWin32Window uiOwner)
            {
                using (var picker = new ColorDialog { Color = System.Drawing.Color.FromArgb(255, (color >> 16) & 255, (color >> 8) & 255, color & 255), FullOpen = true })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? (int?)(picker.Color.ToArgb() & 0xFFFFFF) : null;
            }, null);
        }
        internal static string SaveWorkbook(string filename)
        {
            return Run(null, delegate(IWin32Window uiOwner)
            {
                using (var picker = new SaveFileDialog { Title = "保存本次清单工作簿", Filter = "Excel 工作簿 (*.xlsx)|*.xlsx", DefaultExt = "xlsx", AddExtension = true, OverwritePrompt = true, FileName = filename })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? picker.FileName : null;
            }, (string)null);
        }
        internal static string SaveSettings(IWin32Window owner, string filename)
        {
            return Run(owner, delegate(IWin32Window uiOwner)
            {
                using (var picker = new SaveFileDialog { Title = "导出单体功能预设", Filter = "单体功能预设 (*.xybatch.json)|*.xybatch.json", DefaultExt = "xybatch.json", AddExtension = true, OverwritePrompt = true, FileName = filename })
                    return picker.ShowDialog(uiOwner) == DialogResult.OK ? picker.FileName : null;
            }, (string)null);
        }
    }
}

