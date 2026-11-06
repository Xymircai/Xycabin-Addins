using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class BatchProgressState
    {
        private readonly object sync = new object();
        private readonly ConcurrentQueue<string> lines = new ConcurrentQueue<string>();
        private readonly Stopwatch elapsed = Stopwatch.StartNew();
        private int cancel;
        private int total, completed;
        private bool done;
        private string status = "准备任务", current = "正在准备，请稍候…";
        internal readonly string Title;
        internal BatchProgressState(string title) { Title = title; }
        internal bool Cancelled { get { return Volatile.Read(ref cancel) != 0; } }
        internal void Cancel() { Interlocked.Exchange(ref cancel, 1); }
        internal void Update(int count, int finished, string message, string task)
        {
            lock (sync) { total = Math.Max(0, count); completed = Math.Max(0, Math.Min(finished, total)); status = message ?? ""; current = task ?? ""; }
        }
        internal void Log(string message) { lines.Enqueue(DateTime.Now.ToString("HH:mm:ss") + "  " + message); }
        internal void Finish(string message)
        {
            lock (sync) { if (done) return; status = message; done = true; elapsed.Stop(); }
            Log(message);
        }
        internal BatchProgressSnapshot Snapshot()
        {
            lock (sync) return new BatchProgressSnapshot { Total = total, Completed = completed, Done = done, Status = status, Current = current, Seconds = elapsed.Elapsed.TotalSeconds };
        }
        internal bool ReadLog(out string line) { return lines.TryDequeue(out line); }
    }
    internal sealed class BatchProgressSnapshot
    {
        internal int Total, Completed;
        internal bool Done;
        internal double Seconds;
        internal string Status, Current;
    }
    internal sealed class BatchProgressForm : Form
    {
        private readonly BatchProgressState state;
        private readonly RichTextBox logs = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White, BorderStyle = BorderStyle.None, WordWrap = true };
        private readonly Label phase = Ui.Label("准备任务", 13, true), task = Ui.Label("", 10, false), counts = Ui.Label("", 10, true), timing = Ui.Label("", 9, false);
        private readonly ProgressBar overall = new ProgressBar { Dock = DockStyle.Top, Height = 18, Style = ProgressBarStyle.Continuous };
        private readonly ProgressBar activity = new ProgressBar { Dock = DockStyle.Top, Height = 5, Style = ProgressBarStyle.Marquee, MarqueeAnimationSpeed = 30 };
        private readonly RoundedButton cancel = Ui.Button("取消任务", 110), close = Ui.Button("关闭", 90);
        private readonly System.Windows.Forms.Timer refresh = new System.Windows.Forms.Timer { Interval = 100 };
        internal BatchProgressForm(BatchProgressState value)
        {
            state = value; Text = value.Title + " · 任务进度";
            DialogSizing.Prepare(this, "audit.batch.progress", "table", Ui.Blue);
            foreach (var label in new[] { phase, task, counts, timing })
            { label.AutoSize = true; label.Dock = DockStyle.Top; label.AutoEllipsis = false; label.Margin = new Padding(0, 3, 0, 5); }
            phase.ForeColor = Ui.Blue;
            logs.Name = "batchProgressLog";
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(18) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var heading = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 1, Padding = new Padding(0, 0, 0, 12) };
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++) heading.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            heading.Controls.Add(phase, 0, 0); heading.Controls.Add(task, 0, 1); heading.Controls.Add(activity, 0, 2);
            var logCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12) }; logCard.Controls.Add(logs);
            var footer = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 1, Padding = new Padding(0, 10, 0, 8) };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++) footer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            footer.Controls.Add(counts, 0, 0); footer.Controls.Add(overall, 0, 1); footer.Controls.Add(timing, 0, 2);
            var actions = Ui.ActionBar(); actions.Controls.Add(cancel); actions.Controls.Add(close);
            root.Controls.Add(heading, 0, 0); root.Controls.Add(logCard, 0, 1); root.Controls.Add(footer, 0, 2); root.Controls.Add(actions, 0, 3); Controls.Add(root);
            close.Enabled = false;
            cancel.Click += delegate { RequestCancel(); };
            close.Click += delegate { Close(); };
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (!DialogPreviewFactory.IsPreview(this) && !Ui.IsInertPreview(this) && !state.Snapshot().Done) { RequestCancel(); e.Cancel = true; } };
            refresh.Tick += delegate { Render(); }; refresh.Start(); Render();
            SizeChanged += delegate { task.MaximumSize = new Size(Math.Max(100, ClientSize.Width - 50), 0); };
            Disposed += delegate { refresh.Stop(); refresh.Dispose(); };
        }
        private void RequestCancel()
        {
            if (state.Cancelled || state.Snapshot().Done) return;
            state.Cancel(); cancel.Enabled = false; state.Log("已请求取消，等待当前安全步骤收尾。");
        }
        internal void Render()
        {
            var snapshot = state.Snapshot();
            task.MaximumSize = phase.MaximumSize = new Size(Math.Max(100, ClientSize.Width - 50), 0);
            phase.Text = snapshot.Status; task.Text = snapshot.Current;
            overall.Maximum = Math.Max(1, snapshot.Total); overall.Value = snapshot.Completed;
            counts.Text = snapshot.Total == 0 ? (snapshot.Done ? "尚未生成文件" : "正在预检，尚未确定任务数") : "已完成 " + snapshot.Completed + " / " + snapshot.Total + " 项（" + Math.Floor(1000.0 * snapshot.Completed / snapshot.Total) / 10 + "%）";
            timing.Text = "已用 " + TimeSpan.FromSeconds(snapshot.Seconds).ToString(@"hh\:mm\:ss") + (snapshot.Done ? "" : " · 指示条表示程序仍在处理，完成数量只按实际任务更新");
            activity.Visible = !snapshot.Done; cancel.Enabled = !snapshot.Done && !state.Cancelled; close.Enabled = snapshot.Done;
            string line; int count = 0;
            while (count < 250 && state.ReadLog(out line)) { logs.AppendText(line + Environment.NewLine); count++; }
            if (logs.TextLength > 500000) { logs.Select(0, 100000); logs.SelectedText = ""; }
            if (count > 0) { logs.SelectionStart = logs.TextLength; logs.ScrollToCaret(); }
        }
    }
    internal sealed class BatchProgressSession
    {
        internal readonly BatchProgressState State;
        private readonly HostDialogService.ModelessWindow window;
        internal BatchProgressSession(string title)
        {
            State = new BatchProgressState(title);
            BatchProgressState model = State;
            window = HostDialogService.OpenModeless(delegate { return new BatchProgressForm(model); });
        }
        internal void Close() { State.Finish("窗口已关闭"); window.Close(); }
    }
}
