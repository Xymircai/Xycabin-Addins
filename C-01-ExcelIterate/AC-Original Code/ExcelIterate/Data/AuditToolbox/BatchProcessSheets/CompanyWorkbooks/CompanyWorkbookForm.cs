using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class CompanyWorkbookForm : Form
    {
        private readonly TextBox template = new TextBox { ReadOnly = true, Width = 520 }, output = new TextBox { ReadOnly = true, Width = 520 };
        private readonly ComboBox source = Choice(), target = Choice(), naming = Choice(), code = Choice(), full = Choice(), targetCode = Choice(), targetName = Choice();
        private readonly NumericUpDown header = Number(1), first = Number(0), last = Number(0), leadHeader = Number(1), leadRow = Number(0);
        private readonly CheckBox update = new CheckBox { Text = "同时更新 Lead 的指定行", Checked = true, AutoSize = true };
        private readonly CheckBox confirmed = new CheckBox { Text = "已确认上述来源范围及 Lead 唯一写入行；仅更新指定的两个字段", AutoSize = true };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Top };
        private readonly FlowLayoutPanel inputs = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        private readonly RoundedButton start = Ui.Button("预检并开始生成", 155), cancel = Ui.Button("取消任务", 100);
        private readonly Timer timer = new Timer { Interval = 30 };
        private CompanyWorkbookHost host;
        private CompanyWorkbookRequest request;
        private List<CompanyWorkbookRecord> records;
        private int index, success, failed;
        private bool ticking, cancelled, closePending, loading;
        internal event Action<bool> BusyChanged;
        internal Func<string, CompanyWorkbookHost> OpenTemplate;
        internal Func<bool> CanRun;
        internal bool Running { get { return records != null; } }
        private static ComboBox Choice() { return new ComboBox { Width = 180, DropDownWidth = 300, DropDownStyle = ComboBoxStyle.DropDownList }; }
        private static NumericUpDown Number(int value) { return new NumericUpDown { Minimum = 0, Maximum = 1048576, Value = value, Width = 85 }; }
        internal CompanyWorkbookForm()
        {
            Text = "按公司生成工作簿"; DialogSizing.Prepare(this, "company.workbooks", "table", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 38)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var browse = Ui.Button("选择模板…", 115); browse.Click += delegate { PickTemplate(); };
            AddRow("模板文件", template, browse);
            AddRow("来源工作表", source, Label("表头行"), header);
            AddRow("来源数据行", Label("起始（必填）"), first, Label("结束（必填，含此行）"), last);
            AddRow("文件命名字段", naming);
            AddRow("生成方式", update);
            AddRow("公司来源字段", Label("编码"), code, Label("全称"), full);
            AddRow("目标工作表", target, Label("表头行"), leadHeader);
            AddRow("目标字段", targetCode, targetName);
            AddRow("Lead 写入行", leadRow, Label("必填：仅替换该行两个字段，其他行保留"));
            var chooseOutput = Ui.Button("选择文件夹…", 115); chooseOutput.Click += delegate
            { string selected = HostDialogService.Folder(this, "选择过程表输出文件夹", true); if (!string.IsNullOrEmpty(selected)) output.Text = selected; };
            AddRow("输出文件夹", output, chooseOutput);
            AddRow("核对", confirmed);
            AddRow("说明", Label("读取磁盘模板快照；空行提示跳过，其他异常阻断；已有文件绝不覆盖。\n只复制模式保留原文件字节；更新 Lead 模式会经宿主保存，请先用小样本核验。"));
            root.Controls.Add(inputs, 0, 0); root.Controls.Add(progress, 0, 1); root.Controls.Add(log, 0, 2);
            var actions = Ui.ActionBar(); var close = Ui.Button("关闭", 90);
            actions.Controls.Add(start); actions.Controls.Add(cancel); actions.Controls.Add(close); root.Controls.Add(actions, 0, 3); Controls.Add(root);
            start.Click += delegate { Start(); }; cancel.Click += delegate { cancelled = true; Append("已请求取消，当前文件完成后停止。"); }; close.Click += delegate { Close(); };
            cancel.Enabled = false;
            source.SelectedIndexChanged += delegate { RefreshFields(false); };
            target.SelectedIndexChanged += delegate { RefreshFields(true); };
            header.ValueChanged += delegate { RefreshFields(false); }; leadHeader.ValueChanged += delegate { RefreshFields(true); };
            update.CheckedChanged += delegate { SetLeadEnabled(); };
            timer.Tick += delegate { Tick(); };
            FormClosing += delegate(object sender, FormClosingEventArgs e)
            { if (Running || ticking) { e.Cancel = true; closePending = true; cancelled = true; } };
            Disposed += delegate { timer.Stop(); timer.Dispose(); ReleaseHost(); };
        }
        private static Label Label(string text) { return new Label { Text = text, AutoSize = true, Padding = new Padding(0, 6, 0, 0) }; }
        private void AddRow(string title, params Control[] controls)
        {
            var row = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Width = 1000, Margin = new Padding(0, 3, 0, 3) };
            row.Controls.Add(new Label { Text = title, Width = 105, Padding = new Padding(0, 6, 0, 0) });
            row.Controls.AddRange(controls); inputs.Controls.Add(row);
        }
        internal void Append(string message) { if (!log.IsDisposed) log.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + message + Environment.NewLine); }
        internal void Shutdown()
        {
            cancelled = true; closePending = true;
            if (!ticking)
            {
                timer.Stop(); if (Running) Finish(); Close();
            }
        }
        private void ReleaseHost()
        {
            var old = host; host = null;
            if (old != null) try { old.Dispose(); } catch (Exception error) { Append("读取副本清理失败：" + error.Message); }
        }
        private void PickTemplate()
        {
            if (CanRun != null && !CanRun()) { Append("宿主正在执行其他任务，请稍后再试。"); return; }
            string[] paths = HostDialogService.OpenFiles(this, "Excel 模板 (*.xlsx;*.xls)|*.xlsx;*.xls");
            if (paths.Length == 0) return;
            if (paths.Length != 1) { Append("请每次只选择一个原始模板。"); return; }
            {
                try
                {
                    SetBusy(true);
                    ReleaseHost(); template.Clear();
                    loading = true; source.Items.Clear(); target.Items.Clear(); naming.Items.Clear(); code.Items.Clear(); full.Items.Clear(); targetCode.Items.Clear(); targetName.Items.Clear();
                    if (OpenTemplate == null) throw new InvalidOperationException("模板读取服务尚未连接。");
                    host = OpenTemplate(paths[0]); template.Text = paths[0];
                    source.Items.AddRange(host.Sheets.ToArray()); target.Items.AddRange(host.Sheets.ToArray());
                    Select(source, "公司主题汇总"); Select(target, "lead");
                    confirmed.Checked = false;
                    Append("已读取磁盘模板快照，共 " + host.Sheets.Count + " 张表；如修改原模板，请重新选择以刷新快照。");
                    if (source.SelectedIndex < 0 || target.SelectedIndex < 0) Append("默认来源表或 lead 不存在，请从实际工作表中选择。");
                }
                catch (Exception error) { Append("读取模板失败 [" + paths[0] + "]：" + error); }
                finally { loading = false; SetBusy(false); RefreshFields(false); RefreshFields(true); }
            }
        }
        private static void Select(ComboBox box, string value)
        {
            box.SelectedIndex = -1;
            for (int i = 0; i < box.Items.Count; i++) if (string.Equals(Convert.ToString(box.Items[i]), value, StringComparison.OrdinalIgnoreCase)) { box.SelectedIndex = i; break; }
        }
        private void RefreshFields(bool isTarget)
        {
            if (loading || Running) return;
            var boxes = isTarget ? new[] { targetCode, targetName } : new[] { naming, code, full };
            string[] previous = Array.ConvertAll(boxes, b => b.Text);
            foreach (var box in boxes) box.Items.Clear();
            confirmed.Checked = false;
            if (host == null || (isTarget ? target : source).SelectedIndex < 0) return;
            try
            {
                var headers = host.Headers((isTarget ? target : source).Text, (int)(isTarget ? leadHeader : header).Value);
                string[] values = new List<string>(headers.Keys).ToArray();
                for (int i = 0; i < boxes.Length; i++)
                {
                    boxes[i].Items.AddRange(values);
                    Select(boxes[i], previous[i]);
                    if (boxes[i].SelectedIndex < 0) Select(boxes[i], isTarget ? (i == 0 ? "编码&简称" : "全称") : (i == 0 ? "公司名称" : i == 1 ? "编码&简称" : "全称"));
                }
            }
            catch (Exception error) { Append("读取字段失败：" + error.Message); }
        }
        private void SetLeadEnabled() { foreach (Control c in new Control[] { target, leadHeader, leadRow, targetCode, targetName, code, full }) c.Enabled = update.Checked; }
        private void SetBusy(bool busy)
        {
            inputs.Enabled = !busy; start.Enabled = !busy; cancel.Enabled = busy;
            if (BusyChanged != null) BusyChanged(busy);
        }
        private void Start()
        {
            if (Running || ticking) return;
            if (CanRun != null && !CanRun()) { Append("宿主正在执行其他任务，请稍后再试。"); return; }
            try
            {
                if (host == null) throw new InvalidOperationException("请先选择模板。");
                if (!confirmed.Checked) throw new InvalidOperationException("请核对并勾选确认来源范围及目标写入行。");
                SetBusy(true);
                request = new CompanyWorkbookRequest { SourceSheet = source.Text, TargetSheet = target.Text, NameField = naming.Text, CodeField = code.Text, FullNameField = full.Text,
                    TargetCodeField = targetCode.Text, TargetNameField = targetName.Text, Output = output.Text, HeaderRow = (int)header.Value,
                    FirstRow = (int)first.Value, LastRow = (int)last.Value, TargetHeaderRow = (int)leadHeader.Value, TargetRow = (int)leadRow.Value, UpdateLead = update.Checked };
                records = host.Preview(request, Append); index = success = failed = 0; cancelled = false;
                progress.Maximum = records.Count; progress.Value = 0;
                Append("预检通过，共 " + records.Count + " 家；来源「" + request.SourceSheet + "」，目标「" + request.TargetSheet + "」第 " + request.TargetRow + " 行。开始生成。");
                timer.Start();
            }
            catch (Exception error) { records = null; Append("预检失败：" + error); SetBusy(false); }
        }
        private void Tick()
        {
            if (ticking || !Running) return;
            ticking = true; timer.Stop();
            try
            {
                if (cancelled || index >= records.Count) { Finish(); return; }
                var record = records[index];
                try { host.Generate(request, record, CompanyWorkbookHost.UpdateLead); success++; Append("成功：来源第 " + record.Row + " 行 → " + record.OutputPath); }
                catch (Exception error)
                {
                    failed++;
                    if (error is CompanyWorkbookStateException) cancelled = true;
                    Append("失败：来源第 " + record.Row + " 行 → " + record.OutputPath + "；工作表「" + request.TargetSheet + "」，字段「" + request.TargetCodeField + " / " + request.TargetNameField + "」：" + error);
                }
                index++; progress.Value = index;
            }
            finally
            {
                ticking = false;
                if (Running) timer.Start();
                else if (closePending) Close();
            }
        }
        private void Finish()
        {
            Append((cancelled ? "已取消" : "已结束") + "：成功 " + success + "，失败 " + failed + "，未处理 " + (records.Count - index) + "。已成功文件保留；日志可选中复制。");
            records = null; SetBusy(false);
        }
    }
}
