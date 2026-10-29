using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

// Native-host regression: no WinForms SynchronizationContext when the scan starts.
// Uses only its own files; never connects to Office/WPS or calls PDF export.
internal static class PdfAsyncUiProbe
{
    private const BindingFlags Any = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    private static void Require(bool value, string message) { if (!value) throw new Exception(message); }
    private static void Pump(Func<bool> done)
    {
        var timer = Stopwatch.StartNew();
        while (!done() && timer.ElapsedMilliseconds < 10000) { Application.DoEvents(); Thread.Sleep(10); }
        Require(done(), "Completion timed out");
    }
    [STAThread]
    private static int Main(string[] args)
    {
        string folder = Path.Combine(Path.GetFullPath(args[1]), "async-ui-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            Control.CheckForIllegalCrossThreadCalls = true;
            SynchronizationContext.SetSynchronizationContext(null);
            WindowsFormsSynchronizationContext.AutoInstall = false;
            var assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            string ns = assembly.GetName().Name;
            var type = assembly.GetType(ns + ".PdfToolsForm", true);
            string source = Path.Combine(folder, "source.xlsx"); File.WriteAllText(source, "scan only");
            using (var form = (Form)Activator.CreateInstance(type, Any, null, new object[] { "convert" }, null))
            {
                IntPtr handle = form.Handle;
                var selection = type.GetField("selection", Any).GetValue(form);
                selection.GetType().GetMethod("AddInput", Any).Invoke(selection, new object[] { source });
                WindowsFormsSynchronizationContext.AutoInstall = false;
                SynchronizationContext.SetSynchronizationContext(null);
                Require(SynchronizationContext.Current == null, "Fixture installed a synchronization context");
                type.GetMethod("Scan", Any).Invoke(form, null);
                Pump(() => !(bool)type.GetField("working", Any).GetValue(form));
                var bar = (Control)type.GetField("extensionsBar", Any).GetValue(form);
                Require(bar.Controls.Count == 1, "Scan did not build the extension control");
                Require(!bar.Controls[0].InvokeRequired, "Extension control belongs to another thread");
                var files = (ListBox)type.GetField("files", Any).GetValue(form);
                Require(files.Items.Count == 1, "Scanned file did not reach the UI");
                var log = (TextBoxBase)type.GetField("log", Any).GetValue(form);
                Require(!log.Text.Contains("失败"), "Completion reported an error: " + log.Text);
                log.Text = "OLD BATCH ERROR";
                IntPtr tabHandle = ((TabControl)type.GetField("tabs", Any).GetValue(form)).Handle;
                type.GetMethod("SelectMode", Any).Invoke(form, new object[] { "current" });
                Require(!log.Text.Contains("OLD BATCH ERROR"), "Current-document tab retained the old batch error");
                Require((string)type.GetProperty("Mode", Any).GetValue(form, null) == "current", "Current-document tab did not select current mode");
                int ownerThread = Thread.CurrentThread.ManagedThreadId, callbackThread = 0;
                var callbackPosted = new ManualResetEvent(false);
                ThreadPool.QueueUserWorkItem(delegate {
                    type.GetMethod("CompleteOnUi", Any).Invoke(form, new object[] { new Action(() => callbackThread = Thread.CurrentThread.ManagedThreadId) });
                    callbackPosted.Set();
                });
                Pump(() => callbackThread != 0);
                Require(callbackThread == ownerThread, "Worker completion did not return to owner thread");
                Require(callbackPosted.WaitOne(1000), "Worker did not finish posting"); callbackPosted.Dispose();
                // A completion arriving after disposal must be discarded without creating a handle.
                form.Dispose();
                type.GetMethod("CompleteOnUi", Any).Invoke(form, new object[] { new Action(() => { throw new Exception("Disposed callback ran"); }) });
                Require(!form.IsHandleCreated, "Late completion recreated the disposed handle");
            }
            Console.WriteLine("PASS: scan without synchronization context; owner-thread controls and completion; late disposal.");
            return 0;
        }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
}
