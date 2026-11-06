using System;
using System.Threading;
using System.Windows.Forms;
using Excel_Iterate;

internal static class BatchProgressChecks
{
    internal static void Run(Action<bool, string> check)
    {
        var state = new BatchProgressState("离线进度检查");
        state.Update(3, 1, "处理中", "保存副本");
        check(state.Snapshot().Completed == 1 && !state.Snapshot().Done, "progress counts only explicit completed work");
        state.Log("step one"); state.Log("step two"); string first, second;
        check(state.ReadLog(out first) && state.ReadLog(out second) && first.Contains("step one") && second.Contains("step two"), "streamed progress messages retain ordering");
        state.Cancel(); check(state.Cancelled && state.Snapshot().Completed == 1, "cancellation flag does not advance completion");
        state.Finish("cancelled"); state.Finish("wrong terminal message");
        check(state.Snapshot().Done && state.Snapshot().Status == "cancelled" && state.Snapshot().Completed == 1, "cancelled progress keeps actual completion and first terminal status");
        var closingMethod = typeof(Form).GetMethod("OnFormClosing", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
        using (var preview = DialogPreviewFactory.Create("audit.batch.progress", new System.Drawing.Size(840, 600)))
        {
            var args = new FormClosingEventArgs(CloseReason.UserClosing, false); closingMethod.Invoke(preview, new object[] { args });
            check(!args.Cancel, "inert progress size preview can close without a running task");
        }
        var pending = new BatchProgressState("pending");
        using (var running = new BatchProgressForm(pending))
        {
            var args = new FormClosingEventArgs(CloseReason.UserClosing, false); closingMethod.Invoke(running, new object[] { args });
            check(args.Cancel && pending.Cancelled, "closing a running progress requests safe cancellation");
            pending.Finish("stopped"); args = new FormClosingEventArgs(CloseReason.UserClosing, false); closingMethod.Invoke(running, new object[] { args });
            check(!args.Cancel, "completed progress can be closed");
        }
        var live = new BatchProgressState("独立UI检查");
        using (var ready = new ManualResetEvent(false))
        using (var closed = new ManualResetEvent(false))
        {
            int hostThread = Thread.CurrentThread.ManagedThreadId, uiThread = 0, ticks = 0, rendered = 0;
            var window = HostDialogService.OpenModeless(delegate
            {
                var form = new BatchProgressForm(live) { TopLevel = false };
                var parking = new Panel(); parking.Controls.Add(form);
                var timer = new System.Windows.Forms.Timer { Interval = 50 };
                timer.Tick += delegate
                {
                    Interlocked.Increment(ref ticks);
                    if (live.Snapshot().Completed == 1)
                    {
                        form.Render();
                        if (form.Controls.Find("batchProgressLog", true)[0].Text.Contains("live message")) Interlocked.Exchange(ref rendered, 1);
                    }
                };
                form.Load += delegate { uiThread = Thread.CurrentThread.ManagedThreadId; timer.Start(); ready.Set(); };
                form.Disposed += delegate { timer.Dispose(); parking.Dispose(); closed.Set(); };
                return form;
            });
            try
            {
                check(ready.WaitOne(5000), "isolated progress UI initializes without Excel");
                live.Update(2, 1, "saving", "long host operation"); live.Log("live message");
                // Deliberately do not pump the host thread while the UI keeps painting.
                Thread.Sleep(350);
                check(uiThread != hostThread && Volatile.Read(ref ticks) >= 2 && Volatile.Read(ref rendered) == 1, "progress UI updates while host STA is blocked");
            }
            finally { live.Finish("done"); window.Close(); check(closed.WaitOne(5000), "progress thread closes and disposes its own UI"); }
        }
    }
}
