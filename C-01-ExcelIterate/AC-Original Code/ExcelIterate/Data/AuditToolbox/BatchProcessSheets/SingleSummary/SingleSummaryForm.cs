using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 单体汇总表生成 —— 界面
    //
    // 只做参数收集与结果呈现；COM 一律经事件抛回 Connect，在 Ribbon 宿主线程执行。
    // 沿用「执行 → 取消操作 / 确认操作」二段式：执行后保留备份表，可整段还原。
    //
    // 交互要点：
    //   · Lead 工作表可下拉改选，默认 lead
    //   · 模板列填列字母（默认取点按钮时的活动列）
    //   · 名单里勾选=生成该列；勾选状态执行时回写 Lead 的「选用」列
    // =====================================================================

    internal sealed class SingleSummaryForm : Form, IDialogDefaultLayout
    {
        private readonly Label workbookLabel = Ui.Label("", 9, false);
        private readonly ComboBox leadSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 180, Name = "leadSheetBox" };
        private readonly TextBox headerRow = new TextBox { Width = 56, Text = "1", Name = "headerRowBox" };
        private readonly TextBox templateColumn = new TextBox { Width = 56, Text = "D", Name = "templateColumnBox" };
        private readonly TextBox placeholderBox = new TextBox { Width = 90, Text = SummaryHeaders.Placeholder, Name = "placeholderBox" };
        private readonly TextBox sourceFile = new TextBox { Width = 280, Name = "sourceFileBox" };
        private readonly Label targetLabel = Ui.Label("", 9, false);
        private readonly Label hostLabel = Ui.Label("", 9, false);
        private readonly Label hint = Ui.Label("", 9, true);
        private readonly RichTextBox logView = new RichTextBox
        { ReadOnly = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White };
        private readonly FlowLayoutPanel actions = Ui.ActionBar();
        private readonly TableLayoutPanel root;
        internal readonly ResultGrid Companies = new ResultGrid();
        internal readonly RoundedButton Refresh, Preview, Execute, Undo, Confirm, CopyLog, CloseButton;

        private bool busy;
        private bool canUndo;
        private bool previewOk;
        private bool suspendInvalidate;
        private readonly List<string> pendingLog = new List<string>();

        internal event Action SheetsRequested;
        internal event Action PreviewRequested;
        internal event Action ExecuteRequested;
        internal event Action UndoRequested;
        internal event Action ConfirmRequested;

        internal SingleSummaryForm()
        {
            Text = "单体汇总表生成";
            DialogSizing.Prepare(this, "summary.single", "table", Ui.Blue);
            KeyPreview = true;

            root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 0 参数区
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 1 提示
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 52));   // 2 名单
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 48));   // 3 日志
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 4 按钮

            // ---- 参数区 ----
            TableLayoutPanel row1 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 5, Name = "sheetRow" };
            row1.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            for (int i = 0; i < 5; i++) row1.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row1.Controls.Add(Ui.Grip(leadSheet, "Lead 工作表"), 0, 0);
            row1.Controls.Add(Ui.Grip(headerRow, "表头行"), 1, 0);
            row1.Controls.Add(Ui.Grip(templateColumn, "模板列"), 2, 0);
            row1.Controls.Add(Ui.Grip(placeholderBox, "占位名"), 3, 0);
            row1.Controls.Add(Ui.Grip(sourceFile, "模板公司全称"), 4, 0);

            TableLayoutPanel row2 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Name = "targetRow" };
            row2.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row2.Controls.Add(workbookLabel, 0, 0);
            row2.Controls.Add(targetLabel, 0, 1);
            row2.Controls.Add(hostLabel, 0, 2);

            FlowLayoutPanel paramsPanel = new FlowLayoutPanel
            { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Name = "paramPanel" };
            paramsPanel.Controls.Add(row2);
            paramsPanel.Controls.Add(row1);
            // 参数区按内容高度撑开：Dock=Fill 放在 AutoSize 行里会把这一行压扁，
            // 表现就是参数行被裁掉半截（渲染截图实测）。改为 Dock=Top + AutoSize。
            UiCard paramCard = new UiCard
            {
                Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12), Margin = new Padding(0, 0, 0, 8)
            };
            paramCard.Controls.Add(paramsPanel);

            hint.Margin = new Padding(0, 4, 0, 8);
            hint.MaximumSize = new Size(4096, TextRenderer.MeasureText("测", hint.Font, Size.Empty,
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding).Height + 2);
            hint.Text = "Lead 表两列都非空的行是候选公司，默认全勾；勾选状态会回写「选用」列。模板列自身就是第 1 家公司，向右按勾选顺序各占一列。";

            // ---- 名单 ----
            BuildGrid();
            UiCard gridCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 8) };
            gridCard.Controls.Add(Companies);

            // ---- 日志 ----
            UiCard logCard = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = new Padding(0, 0, 0, 8) };
            logCard.Controls.Add(logView);

            // ---- 按钮 ----
            Refresh = Ui.Button("刷新工作表", 112);
            Preview = Ui.Button("读取 Lead 并预检", 150); Ui.StylePrimary(Preview);
            Execute = Ui.Button("执行生成", 108);
            Undo = Ui.Button("取消操作", 104);
            Confirm = Ui.Button("确认操作", 104);
            CopyLog = Ui.Button("复制日志", 100);
            CloseButton = Ui.Button("关闭", 88);
            actions.Controls.AddRange(new Control[] { Refresh, Preview, Execute, Undo, Confirm, CopyLog, CloseButton });
            ToolTip tip = new ToolTip();
            tip.SetToolTip(Refresh, "重新读取当前工作簿的工作表清单（换了工作簿或刚打开文件时用）");
            tip.SetToolTip(Preview, "读 Lead 名单与模板列体检，不写入任何单元格（Ctrl+Shift+P）");
            tip.SetToolTip(Execute, "按勾选铺列并逐列换链；执行前自动备份工作表（Ctrl+Shift+W）");
            tip.SetToolTip(Undo, "按备份表还原本次写入的目标区间");
            tip.SetToolTip(Confirm, "确认本次结果，结束可还原状态（备份表保留）");

            Refresh.Click += delegate { if (!busy && SheetsRequested != null) SheetsRequested(); };
            Preview.Click += delegate { if (!busy && PreviewRequested != null) PreviewRequested(); };
            Execute.Click += delegate { if (!busy && ExecuteRequested != null) ExecuteRequested(); };
            Undo.Click += delegate { if (!busy && UndoRequested != null) UndoRequested(); };
            Confirm.Click += delegate
            {
                if (busy) return;
                if (UiDialog.Show(this, "确认后本批不再提供还原；隐藏的备份工作表保留在工作簿里，可自行删除。",
                    "确认操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
                canUndo = false;
                if (ConfirmRequested != null) ConfirmRequested();
                AppendLog("已确认；本批不再提供还原。");
                RefreshActions();
            };
            CopyLog.Click += delegate
            { try { Clipboard.SetDataObject(logView.Text, true); } catch (Exception ex) { AppendLog("复制失败：" + ex.Message); } };
            CloseButton.Click += delegate { Close(); };

            root.Controls.Add(paramCard, 0, 0);
            root.Controls.Add(hint, 0, 1);
            root.Controls.Add(gridCard, 0, 2);
            root.Controls.Add(logCard, 0, 3);
            root.Controls.Add(actions, 0, 4);
            Controls.Add(root);

            // 预检后改任何参数都要重新预检，避免拿旧体检结果去写新列
            leadSheet.TextChanged += InvalidatePreview;
            headerRow.TextChanged += InvalidatePreview;
            templateColumn.TextChanged += InvalidatePreview;
            placeholderBox.TextChanged += InvalidatePreview;
            sourceFile.TextChanged += InvalidatePreview;

            RefreshActions();
        }

        private void BuildGrid()
        {
            Companies.Dock = DockStyle.Fill;
            Companies.AllowUserToAddRows = false;
            Companies.AllowUserToDeleteRows = false;
            Companies.AllowUserToResizeRows = false;
            Companies.RowHeadersVisible = false;
            Companies.MultiSelect = false;
            Companies.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            Companies.AutoGenerateColumns = false;
            Companies.ScrollBars = ScrollBars.Both;
            Companies.EnableHeadersVisualStyles = false;
            Companies.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            Companies.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            Companies.ShowCellToolTips = true;

            DataGridViewCheckBoxColumn pick = new DataGridViewCheckBoxColumn
            { Name = "pick", HeaderText = "选用", Width = 60, SortMode = DataGridViewColumnSortMode.NotSortable };
            Companies.Columns.Add(pick);
            MakeColumn("code", "编码&简称", 160, 320);
            MakeColumn("name", "全称", 260, 420);
            MakeColumn("leadRow", "Lead 行", 80, 120);
            MakeColumn("prior", "原选用值", 100, 0);
            Companies.Columns[Companies.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            Companies.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            Ui.StyleGrid(Companies);
        }

        private void MakeColumn(string name, string header, int width, int cap)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.Name = name;
            column.HeaderText = header;
            column.ReadOnly = true;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.MinimumWidth = 50;
            column.SortMode = DataGridViewColumnSortMode.NotSortable;
            Font headerFont = Companies.ColumnHeadersDefaultCellStyle.Font ?? Companies.Font;
            int needed = TextRenderer.MeasureText(header, headerFont).Width + 32;
            column.Width = Math.Max(width, needed);
            if (cap > 0) column.Tag = Math.Max(cap, needed);
            Companies.Columns.Add(column);
        }

        // ---------------------------- 数据 ----------------------------

        /// <summary>Connect 打开窗口时回填当前工作簿、工作表清单、模板列。</summary>
        internal void SetContext(string workbook, List<string> sheets, string lead, string targetSheet, string columnLetter)
        {
            if (IsDisposed) return;
            workbookLabel.Text = "工作簿：" + (string.IsNullOrEmpty(workbook) ? "（当前工作簿）" : workbook);
            targetLabel.Text = "汇总工作表：" + (string.IsNullOrEmpty(targetSheet) ? "（未取到）" : targetSheet);
            leadSheet.Items.Clear();
            if (sheets != null) foreach (string name in sheets) leadSheet.Items.Add(name);
            if (sheets == null || sheets.Count == 0)
            {
                hint.Text = "⚠ 没取到活动工作簿的工作表清单：先在 WPS 表格里打开要处理的工作簿，让它成为当前窗口，再点「刷新工作表」。";
                return;
            }
            if (!string.IsNullOrEmpty(lead))
            {
                int index = leadSheet.Items.IndexOf(lead);
                leadSheet.SelectedIndex = index >= 0 ? index : (leadSheet.Items.Count > 0 ? 0 : -1);
            }
            else if (leadSheet.Items.Count > 0) leadSheet.SelectedIndex = 0;
            if (!string.IsNullOrEmpty(columnLetter)) templateColumn.Text = columnLetter;
        }

        /// <summary>从界面收集请求（不含 COM）。</summary>
        internal SummaryRequest BuildRequest()
        {
            SummaryRequest request = new SummaryRequest();
            request.LeadSheet = Convert.ToString(leadSheet.Text).Trim();
            request.TargetSheet = TargetSheetName;
            request.WorkbookName = workbookLabel.Text;
            int row;
            if (!int.TryParse(headerRow.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out row)) row = 1;
            request.HeaderRow = row;
            int column;
            if (!SummaryColumns.TryParse(templateColumn.Text, out column)) column = 0;
            request.TemplateColumn = column;
            request.Placeholder = placeholderBox.Text.Trim();
            string source = sourceFile.Text.Trim();
            if (source.Length > 0 && !source.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) source += ".xlsx";
            request.SourceFile = source;
            foreach (DataGridViewRow gridRow in Companies.Rows)
            {
                SummaryCompany company = gridRow.Tag as SummaryCompany;
                if (company == null) continue;
                bool picked = gridRow.Cells["pick"].Value is bool && (bool)gridRow.Cells["pick"].Value;
                company.Selected = picked;
                request.Companies.Add(company);
            }
            return request;
        }

        internal string TargetSheetName = "";

        internal void SetCompanies(List<SummaryCompany> list)
        {
            Companies.Rows.Clear();
            if (list == null) { RefreshActions(); return; }
            foreach (SummaryCompany company in list)
            {
                int index = Companies.Rows.Add(company.Selected, company.Code, company.Name,
                    company.Row.ToString(CultureInfo.InvariantCulture),
                    company.PriorSelection.Length == 0 ? "（空）" : company.PriorSelection);
                Companies.Rows[index].Tag = company;
            }
            Ui.AutoFitColumns(Companies);
            RefreshActions();
        }

        /// <summary>预检结果：填模板公司全称默认值 + 决定能不能执行。</summary>
        internal void SetPreview(string suggestedSource, bool ok, string text)
        {
            if (IsDisposed) return;
            suspendInvalidate = true;
            try { if (suggestedSource != null && suggestedSource.Length > 0) sourceFile.Text = suggestedSource; }
            finally { suspendInvalidate = false; }
            previewOk = ok;
            SetLog(text);
            RefreshActions();
        }

        /// <summary>参数改动后旧预检结果作废，必须重新体检。</summary>
        private void InvalidatePreview(object sender, EventArgs e)
        {
            if (suspendInvalidate || !previewOk) return;
            previewOk = false;
            RefreshActions();
            AppendLog("参数已改动，请重新点「读取 Lead 并预检」。");
        }

        /// <summary>把占位名换掉（预检发现模板列实际表头与默认值不一致时用）。</summary>
        internal void SetPlaceholder(string text)
        {
            if (IsDisposed) return;
            suspendInvalidate = true;
            try { placeholderBox.Text = text ?? ""; }
            finally { suspendInvalidate = false; }
        }

        /// <summary>宿主信息（版本等），空工作簿时用来告诉用户到底连到了谁。</summary>
        internal void SetHint(string text) { if (!IsDisposed) hostLabel.Text = text ?? ""; }

        internal void SetLog(string text) { logView.Clear(); logView.AppendText(text ?? ""); }

        internal void AppendLog(string text)
        {
            if (IsDisposed) return;
            // 句柄没建时 BeginInvoke 排不上队，日志会被静默丢掉——先缓存，OnHandleCreated 再刷出
            if (!IsHandleCreated) { pendingLog.Add(text); return; }
            BeginInvoke((MethodInvoker)delegate
            { logView.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + text + "\r\n"); });
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (pendingLog.Count == 0) return;
            System.Text.StringBuilder text = new System.Text.StringBuilder(logView.Text);
            foreach (string line in pendingLog) text.AppendLine(DateTime.Now.ToString("HH:mm:ss") + "  " + line);
            pendingLog.Clear();
            logView.Text = text.ToString();
            logView.SelectionStart = logView.TextLength;
            logView.ScrollToCaret();
        }

        internal void SetBusy(bool value) { busy = value; RefreshActions(); }
        internal void SetUndoAvailable(bool value) { canUndo = value; RefreshActions(); }
        internal void MarkUndone() { canUndo = false; RefreshActions(); }
        internal bool IsBusy { get { return busy; } }

        private void RefreshActions()
        {
            bool idle = !busy && !canUndo;
            leadSheet.Enabled = headerRow.Enabled = templateColumn.Enabled = placeholderBox.Enabled = sourceFile.Enabled = idle;
            Companies.ReadOnly = !idle;
            Refresh.Enabled = idle;
            Preview.Enabled = idle;
            Execute.Enabled = idle && previewOk && Companies.Rows.Count > 0;
            Undo.Enabled = canUndo;
            Confirm.Enabled = !busy && canUndo;
            CopyLog.Enabled = CloseButton.Enabled = true;
        }

        public Size MeasureDefaultClient()
        {
            root.PerformLayout();
            return new Size(DialogSizing.LogicalClient(this).Width, 640);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { e.Handled = true; e.SuppressKeyPress = true; Close(); }
            if (e.Control && e.Shift && e.KeyCode == Keys.P && Preview.Enabled && PreviewRequested != null) { e.Handled = true; PreviewRequested(); }
            if (e.Control && e.Shift && e.KeyCode == Keys.W && Execute.Enabled && ExecuteRequested != null) { e.Handled = true; ExecuteRequested(); }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (busy)
            {
                e.Cancel = true;
                AppendLog("正在处理，请等待当前步骤结束再关闭。");
                return;
            }
            if (canUndo) AppendLog("本批仍可还原；如需保留现场，请先点「取消操作」或「确认操作」。");
            base.OnFormClosing(e);
        }
    }
}
