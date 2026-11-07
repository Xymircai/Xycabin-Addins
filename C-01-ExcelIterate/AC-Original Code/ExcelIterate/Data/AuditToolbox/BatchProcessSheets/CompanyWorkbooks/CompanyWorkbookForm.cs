using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class CompanyWorkbookForm : Form
    {
        private readonly TextBox template = new TextBox { ReadOnly = true, Width = 520 }, output = new TextBox { ReadOnly = true, Width = 520 };
        private readonly ComboBox source = Choice(), target = Choice(), naming = Choice(), code = Choice(), full = Choice(), targetCode = Choice(), targetName = Choice();
        private readonly NumericUpDown header = Number(1), first = Number(0), last = Number(0), leadHeader = Number(1), leadRow = Number(0);
        private readonly CheckBox update = new CheckBox { Text = "同时更新 Lead 的指定行", Checked = true, AutoSize = true };
        internal readonly CheckBox StripExcelSuffix = new CheckBox { Text = "自动去除末尾 Excel 后缀", Checked = true, AutoSize = true, Name = "companyStripExcelSuffix" };
        internal readonly NumericUpDown RemoveTailCharacters = new NumericUpDown { Minimum = 0, Maximum = 32767, Width = 85, Name = "companyRemoveTailCharacters" };
        internal readonly RichTextBox NamePreview = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White, WordWrap = false, Name = "companyNamePreview" };
        private readonly TabControl details = new TabControl { Dock = DockStyle.Fill };
        private readonly TabPage namesTab = new TabPage("文件名预览"), logTab = new TabPage("运行日志");
        private readonly CheckBox confirmed = new CheckBox { Text = "已确认上述来源范围及 Lead 唯一写入行；仅更新指定的两个字段", AutoSize = true };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White };
        internal readonly ProgressBar Progress = new ProgressBar { Dock = DockStyle.Top, Style = ProgressBarStyle.Continuous, Height = 18, Name = "companyTaskProgress" };
        internal readonly Label ProgressCounts = new Label { AutoSize = true, Dock = DockStyle.Top, Name = "companyProgressCounts" };
        internal readonly Label ProgressTiming = new Label { AutoSize = true, Dock = DockStyle.Top, Name = "companyProgressTiming" };
        internal readonly Label CurrentTask = new Label { Height = 24, Dock = DockStyle.Top, Name = "companyCurrentTask", AutoEllipsis = true };
        private readonly TableLayoutPanel progressArea = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 4, Padding = new Padding(0, 6, 0, 6) };
        private readonly CompanyWorkbookProgress taskProgress = new CompanyWorkbookProgress();
        private readonly Stopwatch elapsed = new Stopwatch();
        private readonly FlowLayoutPanel inputs = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        private readonly RoundedButton start = Ui.Button("开始生成", 115), cancel = Ui.Button("取消任务", 100);
        internal readonly RoundedButton PreviewNames = Ui.Button("预览文件名", 120);
        private readonly Dictionary<string, Control> selectionButtons = new Dictionary<string, Control>();
        private readonly Label selectionInfo = Label("在表中选择会打开临时只读模板副本；取消保留原参数，原模板修改后须重新选择模板。");
        private readonly Timer timer = new Timer { Interval = 30 };
        private BatchProgressSession progressView;
        private CompanyWorkbookHost host;
        private CompanyWorkbookRequest request;
        private List<CompanyWorkbookRecord> records;
        private bool ticking, cancelled, closePending, loading, preparing;
        internal event Action<bool> BusyChanged;
        internal Func<string, CompanyWorkbookHost> OpenTemplate;
        internal Func<bool> CanRun;
        internal bool Running { get { return records != null; } }
        private static ComboBox Choice() { return new ComboBox { Width = 180, DropDownWidth = 300, DropDownStyle = ComboBoxStyle.DropDownList }; }
        private static void SetNumber(NumericUpDown control, int value)
        { control.Value = value; control.Text = value.ToString(System.Globalization.CultureInfo.InvariantCulture); }
        private static NumericUpDown Number(int value) { return new NumericUpDown { Minimum = 0, Maximum = 1048576, Value = value, Width = 85 }; }
        internal CompanyWorkbookForm()
        {
            Text = "单体模板生成"; DialogSizing.Prepare(this, "company.workbooks", "table", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 62)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 38)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var browse = Ui.Button("选择模板…", 115); browse.Click += delegate { PickTemplate(); };
            AddRow("模板文件", template, browse);
            AddRow("设置预设", new BatchPresetBar("company", CapturePreset, ApplyPreset,
                delegate { return !Running && !ticking && !preparing && (CanRun == null || CanRun()); }, Append));
            AddRow("来源工作表", source, Label("表头行"), header, PickButton("sourceHeader"));
            AddRow("来源数据行", Label("起始（必填）"), first, Label("结束（必填，含此行）"), last, PickButton("data"));
            AddRow("命名名称", naming, PickButton("naming"));
            AddRow("名称处理", Label("从原文末尾去除"), RemoveTailCharacters, Label("个字符，再"), StripExcelSuffix);
            AddRow("命名顺序", Label("先去除指定字符，再去掉 Excel 后缀，最后统一添加模板后缀；默认去除 0 个字符。"));
            AddRow("生成方式", update);
            AddRow("目标表中替换字段", Label("编码"), code, PickButton("code"), Label("全称"), full, PickButton("full"));
            AddRow("目标工作表", target, Label("表头行"), leadHeader, PickButton("targetHeader"));
            AddRow("目标字段", Label("编码"), targetCode, PickButton("targetCode"), Label("全称"), targetName, PickButton("targetName"));
            AddRow("Lead 写入行", leadRow, PickButton("leadRow"), Label("必填：仅替换该行两个字段，其他行保留"));
            var chooseOutput = Ui.Button("选择文件夹…", 115); chooseOutput.Click += delegate
            { string selected = HostDialogService.Folder(this, "选择过程表输出文件夹", true); if (!string.IsNullOrEmpty(selected)) output.Text = selected; };
            AddRow("输出文件夹", output, chooseOutput);
            AddRow("选区结果", selectionInfo);
            AddRow("核对", confirmed);
            AddRow("来源说明", Label("每个公司数据行生成一个工作簿。数据范围不含表头；选中哪些列不限制实际读取字段。"));
            AddRow("字段说明", Label("字段请选择实际表头单元格；新选区会替换原工作表、表头行与字段。程序按字段读取编码、全称及文件名。下拉与手填行号仍可使用。"));
            AddRow("目标说明", Label("仅更新生成副本的指定数据行中编码、全称两个字段。原始模板不会写入或覆盖。"));
            AddRow("说明", Label("读取磁盘模板快照；空行提示跳过，其他异常阻断；已有文件绝不覆盖。\n只复制模式保留原文件字节；更新 Lead 模式会经宿主保存，请先用小样本核验。"));
            progressArea.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 4; i++) progressArea.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            progressArea.Controls.Add(ProgressCounts, 0, 0); progressArea.Controls.Add(Progress, 0, 1);
            progressArea.Controls.Add(ProgressTiming, 0, 2); progressArea.Controls.Add(CurrentTask, 0, 3);
            namesTab.Controls.Add(NamePreview); logTab.Controls.Add(log); details.TabPages.Add(namesTab); details.TabPages.Add(logTab);
            root.Controls.Add(inputs, 0, 0); root.Controls.Add(progressArea, 0, 1); root.Controls.Add(details, 0, 2);
            var actions = Ui.ActionBar(); var close = Ui.Button("关闭", 90);
            actions.Controls.Add(PreviewNames); actions.Controls.Add(start); actions.Controls.Add(cancel); actions.Controls.Add(close); root.Controls.Add(actions, 0, 3); Controls.Add(root);
            PreviewNames.Click += delegate { PreviewFileNames(); };
            start.Click += delegate { Start(); }; cancel.Click += delegate { RequestCancel(); }; close.Click += delegate { Close(); };
            cancel.Enabled = false;
            source.SelectedIndexChanged += delegate { RefreshFields(false); };
            target.SelectedIndexChanged += delegate { RefreshFields(true); };
            header.ValueChanged += delegate { RefreshFields(false); }; leadHeader.ValueChanged += delegate { RefreshFields(true); };
            update.CheckedChanged += delegate { SetLeadEnabled(); };
            StripExcelSuffix.CheckedChanged += delegate { InvalidateNames(); };
            RemoveTailCharacters.ValueChanged += delegate { InvalidateNames(); };
            naming.SelectedIndexChanged += delegate { InvalidateNames(); };
            first.ValueChanged += delegate { InvalidateNames(); }; last.ValueChanged += delegate { InvalidateNames(); };
            output.TextChanged += delegate { InvalidateNames(); };
            foreach (var field in new[] { code, full, targetCode, targetName }) field.SelectedIndexChanged += delegate { InvalidateNames(); };
            leadRow.ValueChanged += delegate { InvalidateNames(); };
            timer.Tick += delegate { Tick(); };
            FormClosing += delegate(object sender, FormClosingEventArgs e)
            { if (Running || ticking || preparing) { e.Cancel = true; closePending = true; RequestCancel(); } };
            Disposed += delegate { timer.Stop(); timer.Dispose(); if (progressView != null) progressView.Close(); ReleaseHost(); };
            inputs.ClientSizeChanged += delegate { FitInputRows(); };
            Shown += delegate { FitInputRows(); };
            ResetProgress("等待开始");
        }
        internal void ResetProgress(string state)
        {
            elapsed.Reset(); taskProgress.Reset(0); CurrentTask.Text = "";
            RenderProgress(state, false);
        }
        private Dictionary<string, Control> PresetFields()
        {
            return new Dictionary<string, Control> {
                { "source", source }, { "header", header }, { "first", first }, { "last", last },
                { "naming", naming }, { "code", code }, { "full", full }, { "update", update },
                { "target", target }, { "targetHeader", leadHeader }, { "targetRow", leadRow },
                { "targetCode", targetCode }, { "targetName", targetName }, { "output", output },
                { "stripSuffix", StripExcelSuffix }, { "removeTail", RemoveTailCharacters }
            };
        }
        internal BatchPreset CapturePreset()
        {
            ValidateChildren();
            return new BatchPreset { Feature = "company", Context = template.Text, Values = BatchPresetStore.Capture(PresetFields()) };
        }
        internal void ApplyPreset(BatchPreset preset)
        {
            if (Running || ticking || preparing || (CanRun != null && !CanRun())) throw new InvalidOperationException("任务执行期间不能应用预设。");
            BatchPresetStore.Validate(preset, "company");
            var fields = PresetFields(); var values = preset.Values;
            BatchPresetStore.ValidateFields(fields, values);
            if (host == null) throw new InvalidOperationException("请先选择模板；原模板位置：" + preset.Context);
            bool updating = bool.Parse(values["update"]);
            if (!host.Sheets.Contains(values["source"]) || (updating && !host.Sheets.Contains(values["target"]))) throw new InvalidOperationException("当前模板缺少预设工作表，请选择对应模板或手动配置。");
            var sourceHeaders = host.Headers(values["source"], int.Parse(values["header"]));
            var targetHeaders = updating ? host.Headers(values["target"], int.Parse(values["targetHeader"])) : new Dictionary<string, int>();
            foreach (string key in new[] { "naming", "code", "full" })
                if (values[key].Length > 0 && !sourceHeaders.ContainsKey(values[key])) throw new InvalidOperationException("当前来源表缺少预设字段：" + values[key]);
            foreach (string key in new[] { "targetCode", "targetName" })
                if (updating && values[key].Length > 0 && !targetHeaders.ContainsKey(values[key])) throw new InvalidOperationException("当前目标表缺少预设字段：" + values[key]);
            loading = true;
            try
            {
                foreach (var box in new[] { naming, code, full }) { box.Items.Clear(); box.Items.AddRange(new List<string>(sourceHeaders.Keys).ToArray()); }
                foreach (var box in new[] { targetCode, targetName }) { box.Items.Clear(); box.Items.AddRange(new List<string>(targetHeaders.Keys).ToArray()); }
                BatchPresetStore.Apply(fields, values);
            }
            finally { loading = false; }
            confirmed.Checked = false; InvalidateNames(); SetLeadEnabled(); ResetProgress("预设已应用，请重新预览");
            selectionInfo.Text = "预设仅应用参数，当前模板：" + template.Text + "；保存时模板：" + preset.Context;
        }
        private void RenderProgress(string state, bool running)
        {
            Progress.Maximum = Math.Max(1, taskProgress.Total); Progress.Value = taskProgress.Completed;
            ProgressCounts.Text = taskProgress.Counts(state);
            ProgressTiming.Text = taskProgress.Timing(elapsed.Elapsed, running);
            if (progressView != null) progressView.State.Update(taskProgress.Total, taskProgress.Completed, state, CurrentTask.Text);
            // Paint before the next STA COM operation without pumping other commands.
            progressArea.Refresh();
        }
        private void RequestCancel()
        {
            if (cancelled) return;
            cancelled = true; cancel.Enabled = false;
            if (progressView != null) progressView.State.Cancel();
            RenderProgress("正在取消，等待当前文件收尾", true);
            Append("已请求取消，当前文件完成后停止。");
        }
        private static Label Label(string text) { return new Label { Text = text, AutoSize = true, Margin = new Padding(3, 6, 3, 3) }; }
        private void InvalidateNames()
        {
            if (Running || preparing) return;
            NamePreview.Clear(); confirmed.Checked = false;
        }
        private CompanyWorkbookRequest BuildRequest(bool namesOnly)
        {
            return new CompanyWorkbookRequest { SourceSheet = source.Text, TargetSheet = target.Text, NameField = naming.Text, CodeField = code.Text, FullNameField = full.Text,
                TargetCodeField = targetCode.Text, TargetNameField = targetName.Text, Output = output.Text, HeaderRow = (int)header.Value,
                FirstRow = (int)first.Value, LastRow = (int)last.Value, TargetHeaderRow = (int)leadHeader.Value, TargetRow = (int)leadRow.Value, UpdateLead = update.Checked && !namesOnly,
                StripExcelSuffix = StripExcelSuffix.Checked, RemoveTailCharacters = (int)RemoveTailCharacters.Value };
        }
        internal void ShowNamePreview(IList<CompanyWorkbookRecord> rows)
        {
            var text = new System.Text.StringBuilder("共 " + rows.Count + " 个文件；后缀与模板一致。\r\n来源行｜来源名称 → 最终文件名\r\n");
            foreach (var row in rows)
                text.Append(row.Row).Append("｜").Append((row.Name ?? "").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t")).Append(" → ").Append(System.IO.Path.GetFileName(row.OutputPath)).AppendLine();
            NamePreview.Text = text.ToString();
        }
        private void PreviewFileNames()
        {
            if (Running || ticking || preparing) return;
            if (CanRun != null && !CanRun()) { Append("宿主正在执行其他任务，请稍后再试。"); return; }
            try
            {
                if (host == null) throw new InvalidOperationException("请先选择模板。");
                ValidateChildren();
                var previewRequest = BuildRequest(true);
                preparing = true; SetBusy(true); cancel.Enabled = false; NamePreview.Clear(); confirmed.Checked = false;
                var preview = host.Preview(previewRequest, Append);
                ShowNamePreview(preview); details.SelectedTab = namesTab;
                elapsed.Reset(); taskProgress.Reset(preview.Count); CurrentTask.Text = "尚未开始生成。"; RenderProgress("文件名已预览，等待生成", false);
                Append("文件名预览完成，共 " + preview.Count + " 项；未生成工作簿。生成前还会校验 Lead 配置和数据。");
            }
            catch (Exception error) { NamePreview.Text = "预览失败，请查看运行日志。"; details.SelectedTab = logTab; Append("预览文件名失败：" + error.Message); }
            finally { preparing = false; SetBusy(false); if (closePending) Close(); }
        }
        private Control PickButton(string kind)
        {
            var button = Ui.Button("在表中选择…", 116); button.Name = "companyPick_" + kind;
            button.Click += delegate { PickParameter(kind); };
            selectionButtons.Add(kind, button); return button;
        }
        private void PickParameter(string kind)
        {
            if (Running || ticking || preparing || (CanRun != null && !CanRun())) return;
            bool hidden = false;
            BatchProgressSession selectionProgress = null;
            try
            {
                if (host == null) throw new InvalidOperationException("请先选择磁盘模板。");
                ValidateChildren();
                var currentRequest = BuildRequest(false);
                preparing = true; SetBusy(true); cancel.Enabled = false;
                selectionProgress = new BatchProgressSession("正在准备表中选区");
                Hide(); hidden = true;
                var selected = host.Pick(kind, currentRequest, null, delegate(string stage)
                {
                    if (selectionProgress.State.Cancelled) throw new OperationCanceledException("已取消准备选区。");
                    if (stage == null) selectionProgress.Close();
                    else { selectionProgress.State.Update(0, 0, stage, "只读副本仅用于读取坐标，不运行外部脚本。"); selectionProgress.State.Log(stage); }
                });
                if (selected != null) ApplyPicked(kind, selected);
            }
            catch (Exception error) { Append("选区未应用：" + error.Message + "；可重新选择或手动配置。"); }
            finally
            {
                if (selectionProgress != null) selectionProgress.Close();
                preparing = false;
                if (!IsDisposed)
                {
                    SetBusy(false);
                    if (hidden) { Show(); Activate(); }
                    if (closePending) Close();
                }
            }
        }
        internal void ApplyPicked(string kind, CompanyWorkbookSelection selected)
        {
            bool fieldPicker = kind == "naming" || kind == "code" || kind == "full" || kind == "targetCode" || kind == "targetName";
            if (fieldPicker)
            {
                bool isTarget = kind == "targetCode" || kind == "targetName";
                var sheetBox = isTarget ? target : source;
                var headerBox = isTarget ? leadHeader : header;
                bool changed = sheetBox.Text != selected.Sheet || headerBox.Value != selected.Row;
                loading = true;
                try { Select(sheetBox, selected.Sheet); SetNumber(headerBox, selected.Row); }
                finally { loading = false; }
                RefreshFields(isTarget);
                if (changed) { if (isTarget) SetNumber(leadRow, 0); else { SetNumber(first, 0); SetNumber(last, 0); } }
            }
            if (kind == "sourceHeader" || kind == "targetHeader")
            {
                bool isTarget = kind == "targetHeader";
                bool changed = (isTarget ? target : source).Text != selected.Sheet || (isTarget ? leadHeader : header).Value != selected.Row;
                loading = true;
                try { Select(isTarget ? target : source, selected.Sheet); SetNumber(isTarget ? leadHeader : header, selected.Row); }
                finally { loading = false; }
                RefreshFields(isTarget);
                if (changed) { if (isTarget) SetNumber(leadRow, 0); else { SetNumber(first, 0); SetNumber(last, 0); } }
            }
            else if (kind == "data") { SetNumber(first, selected.Row); SetNumber(last, selected.LastRow); }
            else if (kind == "leadRow") SetNumber(leadRow, selected.Row);
            else
            {
                ComboBox box = kind == "naming" ? naming : kind == "code" ? code : kind == "full" ? full : kind == "targetCode" ? targetCode : targetName;
                Select(box, selected.Field);
            }
            NamePreview.Clear(); confirmed.Checked = false;
            selectionInfo.Text = "已选择：" + selected.Sheet + " / 第 " + selected.Row + (selected.LastRow > selected.Row ? "—" + selected.LastRow : "") + " 行，列 " + SummaryColumns.Letter(selected.Column)
                + (string.IsNullOrEmpty(selected.Field) ? "" : "，字段「" + selected.Field + "」") + "。请重新预览并核对。";
            Append(selectionInfo.Text);
        }
        private void AddRow(string title, params Control[] controls)
        {
            var row = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 3, 0, 3) };
            row.Controls.Add(new Label { Text = title, AutoSize = true, UseMnemonic = false, Tag = "companyRowTitle", MinimumSize = new Size(105, 0), Margin = new Padding(3, 6, 3, 3) });
            row.Controls.AddRange(controls); inputs.Controls.Add(row);
        }
        private void FitInputRows()
        {
            int width = Math.Max(1, inputs.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - inputs.Padding.Horizontal);
            int titleWidth = 105;
            foreach (Control row in inputs.Controls)
                foreach (Control control in row.Controls)
                    if (Equals(control.Tag, "companyRowTitle"))
                        titleWidth = Math.Max(titleWidth, TextRenderer.MeasureText(control.Text, control.Font, Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width + 4);
            foreach (Control row in inputs.Controls)
            {
                row.MaximumSize = new Size(width, 0);
                row.Width = width;
                foreach (Control control in row.Controls)
                {
                    if (control is BatchPresetBar) control.MaximumSize = new Size(Math.Max(240, width - titleWidth - 20), 0);
                    if (Equals(control.Tag, "companyRowTitle")) control.MinimumSize = control.MaximumSize = new Size(titleWidth, 0);
                    else if (control is Label) control.MaximumSize = new Size(Math.Min(control.MinimumSize.Width > 0 ? control.MinimumSize.Width : width - 12, width), 0);
                }
            }
        }
        internal void Append(string message) { if (!log.IsDisposed) { log.AppendText(DateTime.Now.ToString("HH:mm:ss") + " " + message + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); } if (progressView != null) progressView.State.Log(message); }
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
                    ResetProgress("等待配置");
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
            InvalidateNames();
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
        private void SetLeadEnabled() { foreach (Control c in new Control[] { target, leadHeader, leadRow, targetCode, targetName, code, full }) c.Enabled = update.Checked;
            foreach (string key in new[] { "targetHeader", "targetCode", "targetName", "leadRow", "code", "full" }) selectionButtons[key].Enabled = update.Checked; }
        private void SetBusy(bool busy)
        {
            inputs.Enabled = !busy; start.Enabled = PreviewNames.Enabled = !busy; cancel.Enabled = busy;
            if (!busy) SetLeadEnabled();
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
                ValidateChildren();
                request = BuildRequest(false);
                if (progressView != null) progressView.Close();
                progressView = new BatchProgressSession("单体模板生成");
                preparing = true; cancelled = false;
                SetBusy(true);
                cancel.Enabled = false;
                ResetProgress("正在预检，尚未开始生成");
                NamePreview.Clear();
                Append("正在预检模板、来源行、字段与输出目录…");
                records = host.Preview(request, Append);
                ShowNamePreview(records); details.SelectedTab = logTab;
                taskProgress.Reset(records.Count); elapsed.Restart();
                cancel.Enabled = !cancelled;
                RenderProgress(cancelled ? "正在取消" : "开始生成", true);
                Append("预检通过，共 " + records.Count + " 家；来源「" + request.SourceSheet + "」，目标「" + request.TargetSheet + "」第 " + request.TargetRow + " 行。开始生成。");
                if (progressView.State.Cancelled) RequestCancel();
                timer.Start();
            }
            catch (Exception error) { records = null; NamePreview.Clear(); details.SelectedTab = logTab; ResetProgress("预检失败，未开始生成"); Append("预检失败：" + error); if (progressView != null) progressView.State.Finish("预检失败，未生成文件"); SetBusy(false); }
            finally { preparing = false; if (closePending && !Running) Close(); }
        }
        private void Tick()
        {
            if (ticking || !Running) return;
            ticking = true; timer.Stop();
            try
            {
                if (progressView != null && progressView.State.Cancelled && !cancelled) RequestCancel();
                if (cancelled || taskProgress.Remaining == 0) { Finish(); return; }
                var record = records[taskProgress.Completed];
                CurrentTask.Text = "正在处理第 " + (taskProgress.Completed + 1) + " / " + taskProgress.Total + " 项：" + System.IO.Path.GetFileName(record.OutputPath);
                RenderProgress("正在生成", true);
                bool succeeded = false;
                try { host.Generate(request, record, CompanyWorkbookHost.UpdateLead, Append); succeeded = true; Append("成功：来源第 " + record.Row + " 行 → " + record.OutputPath); }
                catch (Exception error)
                {
                    if (error is CompanyWorkbookStateException) cancelled = true;
                    Append("失败：来源第 " + record.Row + " 行 → " + record.OutputPath + "；工作表「" + request.TargetSheet + "」，字段「" + request.TargetCodeField + " / " + request.TargetNameField + "」：" + error);
                }
                taskProgress.Complete(succeeded);
                RenderProgress(cancelled ? "正在取消" : "正在生成", true);
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
            elapsed.Stop();
            string state = taskProgress.Remaining == 0 ? (taskProgress.Failed > 0 ? "已结束（存在失败）" : "全部完成") : "已取消";
            CurrentTask.Text = taskProgress.Remaining == 0 ? "本批任务已全部处理。" : "已停止，已成功文件保留。";
            RenderProgress(state, false);
            Append(state + "：成功 " + taskProgress.Succeeded + "，失败 " + taskProgress.Failed + "，未处理 " + taskProgress.Remaining + "。已成功文件保留；日志可选中复制。");
            if (progressView != null) progressView.State.Finish(state);
            records = null; SetBusy(false);
        }
    }
}
