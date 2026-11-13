using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // =====================================================================
    // 单体汇总批量链接 —— 界面
    //
    // 只做参数收集与结果呈现；COM 一律经事件抛回 Connect，在 Ribbon 宿主线程执行。
    // 沿用「执行 → 取消操作 / 确认操作」二段式：执行后保留备份表，可整段还原。
    //
    // 交互要点：
    //   · Lead 工作表可下拉改选，默认 lead
    //   · 工作簿初始留空，模板列填列字母，不从活动单元格推断
    //   · 名单里勾选=生成该列；勾选状态执行时回写 Lead 的「选用」列
    // =====================================================================

    internal sealed class SingleSummaryForm : Form, IDialogDefaultLayout
    {
        private readonly Label workbookLabel = Ui.Label("", 9, false);
        internal readonly ComboBox Workbooks = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 410, DropDownWidth = 800, Name = "summaryWorkbook" };
        internal readonly ComboBox TargetSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180, Name = "summaryTargetSheet" };
        internal readonly RoundedButton BrowseWorkbook, CurrentWorkbook, RefreshWorkbooks;
        private string workbookPath = "";
        private bool settingWorkbooks;
        private readonly ComboBox leadSheet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 180, Name = "leadSheetBox" };
        private readonly TextBox headerRow = new TextBox { Width = 56, Text = "1", Name = "headerRowBox" };
        private readonly TextBox sheetHeaderRow = new TextBox { Width = 56, Text = "1", Name = "sheetHeaderRowBox" };
        private readonly TextBox templateColumn = new TextBox { Width = 56, Text = "D", Name = "templateColumnBox" };
        private readonly TextBox placeholderBox = new TextBox { Width = 90, Text = SummaryHeaders.Placeholder, Name = "placeholderBox" };
        private readonly ComboBox sourceFile = new ComboBox { Width = 280, DropDownWidth = 600, DropDownStyle = ComboBoxStyle.DropDown, Name = "sourceFileBox" };
        internal readonly CheckBox AnnualMode = new CheckBox { Text = "按年度展开公司列", AutoSize = true, Name = "summaryAnnualMode" };
        internal readonly NumericUpDown FirstYear = new NumericUpDown { Minimum = 1900, Maximum = 9999, Value = DateTime.Now.Year, Width = 80, Name = "summaryFirstYear" };
        internal readonly NumericUpDown LastYear = new NumericUpDown { Minimum = 1900, Maximum = 9999, Value = DateTime.Now.Year, Width = 80, Name = "summaryLastYear" };
        internal readonly NumericUpDown YearRow = new NumericUpDown { Minimum = 0, Maximum = 1048576, Value = 0, Width = 80, Name = "summaryYearRow" };
        internal readonly CheckBox DescendingYears = new CheckBox { Text = "年度倒序", AutoSize = true, Checked = true, Name = "summaryYearDescending" };
        internal readonly CheckBox CheckYearFiles = new CheckBox { Text = "检查年度来源文件是否存在", AutoSize = true, Checked = false, Name = "summaryCheckYearFiles" };
        internal readonly TextBox YearPathPattern = new TextBox { Width = 310, Name = "summaryYearPathPattern" };
        private readonly TableLayoutPanel annualOptions = new TableLayoutPanel { AutoSize = true, ColumnCount = 1, RowCount = 2, Dock = DockStyle.Top, Visible = false };
        private readonly Label hostLabel = Ui.Label("", 9, false);
        private readonly Label hint = Ui.Label("", 9, true);
        private readonly RichTextBox logView = new RichTextBox
        { ReadOnly = true, WordWrap = true, ScrollBars = RichTextBoxScrollBars.Vertical, Dock = DockStyle.Fill, BorderStyle = BorderStyle.None, BackColor = Color.White };
        private readonly FlowLayoutPanel actions = Ui.ActionBar();
        private readonly TableLayoutPanel root;
        internal readonly ResultGrid Companies = new ResultGrid();
        internal readonly RoundedButton Refresh, Preview, Execute, Undo, Confirm, CopyLog, CloseButton;

        internal BatchProgressSession ProgressSession;
        internal BatchProgressSession OpenProgress()
        { if (ProgressSession != null) ProgressSession.Close(); ProgressSession = new BatchProgressSession("单体汇总批量链接"); return ProgressSession; }
        private bool busy;
        private BatchPresetActions presets;
        private bool canUndo;
        private bool previewOk;
        private bool suspendInvalidate;
        private readonly List<string> pendingLog = new List<string>();

        internal event Action SheetsRequested;
        internal event Action<string> SelectionRequested;
        internal event Action<string> RangeRequested;
        private readonly List<Control> selectionButtons = new List<Control>();
        internal event Action WorkbooksRequested, BrowseRequested, CurrentRequested;
        internal event Action<string> WorkbookRequested;
        internal event Action<BatchPreset> PresetRequested;
        internal event Action PreviewRequested;
        internal event Action ExecuteRequested;
        internal event Action UndoRequested;
        internal event Action ConfirmRequested;

        internal SingleSummaryForm()
        {
            Text = "单体汇总批量链接";
            Disposed += delegate { if (ProgressSession != null) ProgressSession.Close(); };
            DialogSizing.Prepare(this, "summary.single", "table", Ui.Blue);
            KeyPreview = true;

            root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 50));    // 0 可滚动参数区
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 1 提示
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 26));   // 2 名单
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 24));   // 3 日志
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 4 按钮

            // ---- 参数区 ----
            BrowseWorkbook = Ui.Button("浏览文件…", 108);
            CurrentWorkbook = Ui.Button("使用当前工作簿", 150);
            RefreshWorkbooks = Ui.Button("刷新列表", 96);
            var picker = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true, Name = "workbookPicker" };
            picker.Controls.Add(Ui.Grip(Workbooks, "工作簿"));
            picker.Controls.Add(RefreshWorkbooks); picker.Controls.Add(BrowseWorkbook); picker.Controls.Add(CurrentWorkbook);
            BrowseWorkbook.Click += delegate { if (!busy && !canUndo && BrowseRequested != null) BrowseRequested(); };
            CurrentWorkbook.Click += delegate { if (!busy && !canUndo && CurrentRequested != null) CurrentRequested(); };
            RefreshWorkbooks.Click += delegate { if (!busy && !canUndo && WorkbooksRequested != null) WorkbooksRequested(); };
            Workbooks.SelectedIndexChanged += delegate
            {
                var choice = Workbooks.SelectedItem as SummaryWorkbookChoice;
                if (!settingWorkbooks && !busy && !canUndo && choice != null && WorkbookRequested != null) WorkbookRequested(choice.Path);
            };
            // 两行放参数：Lead 的三项一行，汇总表的三项一行（1180 宽下不至于挤成一条）
            var rowA = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Name = "leadRow" };
            rowA.Controls.Add(Ui.Grip(leadSheet, "Lead 工作表"));
            rowA.Controls.Add(SelectionParameter(headerRow, "Lead 表头行", "lead"));
            rowA.Controls.Add(Ui.Grip(TargetSheet, "汇总工作表"));
            rowA.Controls.Add(SelectionParameter(sheetHeaderRow, "汇总表表头行", "header"));
            var rowB = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Name = "templateRow" };
            rowB.Controls.Add(SelectionParameter(templateColumn, "模板列", "column"));
            rowB.Controls.Add(Ui.Grip(placeholderBox, "占位名（可读取）"));
            rowB.Controls.Add(Ui.Grip(sourceFile, "模板公司文件名（可识别或选择）"));

            TableLayoutPanel row2 = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Name = "targetRow" };
            row2.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row2.Controls.Add(workbookLabel, 0, 0);
            row2.Controls.Add(hostLabel, 0, 1);
            Action showContext = delegate { workbookLabel.Visible = workbookLabel.Text.Length > 0; hostLabel.Visible = hostLabel.Text.Length > 0; row2.Visible = workbookLabel.Text.Length > 0 || hostLabel.Text.Length > 0; };
            workbookLabel.TextChanged += delegate { showContext(); }; hostLabel.TextChanged += delegate { showContext(); }; showContext();

            TableLayoutPanel paramsPanel = new TableLayoutPanel
            { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, RowCount = 7, Name = "paramPanel" };
            paramsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 7; i++) paramsPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            presets = new BatchPresetActions(this, "summary", CapturePreset,
                delegate(BatchPreset value) { if (PresetRequested != null) PresetRequested(value); else ApplyPreset(value); },
                delegate { return !busy && !canUndo; }, AppendLog);
            paramsPanel.Controls.Add(presets.Header(picker), 0, 0);
            paramsPanel.Controls.Add(row2, 0, 1);
            paramsPanel.Controls.Add(rowA, 0, 2);
            paramsPanel.Controls.Add(rowB, 0, 3);
            paramsPanel.Controls.Add(AnnualMode, 0, 4);
            var years = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Dock = DockStyle.Top };
            years.Controls.Add(Ui.Grip(FirstYear, "开始年度")); years.Controls.Add(Ui.Grip(LastYear, "结束年度"));
            years.Controls.Add(SelectionParameter(YearRow, "年度表头行（必填）", "year")); years.Controls.Add(DescendingYears);
            var paths = new FlowLayoutPanel { AutoSize = true, WrapContents = true, Dock = DockStyle.Top };
            paths.Controls.Add(Ui.Grip(YearPathPattern, "年度目录段（留空自动识别，手填含 {yyyy}）")); paths.Controls.Add(CheckYearFiles);
            annualOptions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            annualOptions.RowStyles.Add(new RowStyle(SizeType.AutoSize)); annualOptions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            annualOptions.Controls.Add(years, 0, 0); annualOptions.Controls.Add(paths, 0, 1); paramsPanel.Controls.Add(annualOptions, 0, 5);
            // 参数区按内容高度撑开：Dock=Fill 放在 AutoSize 行里会把这一行压扁，
            // 表现就是参数行被裁掉半截（渲染截图实测）。改为 Dock=Top + AutoSize。
            UiCard paramCard = new UiCard
            {
                Dock = DockStyle.Fill, AutoSize = false, AutoScroll = true,
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
            Execute = Ui.Button("执行链接", 108);
            Undo = Ui.Button("取消操作", 104);
            Confirm = Ui.Button("确认操作", 104);
            CopyLog = Ui.Button("复制日志", 100);
            CloseButton = Ui.Button("关闭", 88);
            actions.Controls.AddRange(new Control[] { Refresh, Preview, Execute, Undo, Confirm, CopyLog, CloseButton });
            ToolTip tip = new ToolTip();
            tip.SetToolTip(Refresh, "重新读取已选择工作簿的工作表清单，不跟随活动工作簿切换");
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
            TargetSheet.SelectedIndexChanged += InvalidatePreview;
            headerRow.TextChanged += InvalidatePreview;
            sheetHeaderRow.TextChanged += InvalidatePreview;
            templateColumn.TextChanged += InvalidatePreview;
            placeholderBox.TextChanged += InvalidatePreview;
            sourceFile.TextChanged += InvalidatePreview;
            AnnualMode.CheckedChanged += delegate { annualOptions.Visible = AnnualMode.Checked; InvalidatePreview(AnnualMode, EventArgs.Empty); };
            FirstYear.ValueChanged += InvalidatePreview; LastYear.ValueChanged += InvalidatePreview; YearRow.ValueChanged += InvalidatePreview;
            DescendingYears.CheckedChanged += InvalidatePreview; CheckYearFiles.CheckedChanged += InvalidatePreview; YearPathPattern.TextChanged += InvalidatePreview;
            Companies.CurrentCellDirtyStateChanged += delegate { if (Companies.IsCurrentCellDirty) Companies.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            Companies.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            { if (!busy && AnnualMode.Checked && e.ColumnIndex == 0) InvalidatePreview(Companies, EventArgs.Empty); };

            RefreshActions();
        }

        private Control SelectionParameter(Control field, string title, string kind)
        {
            var panel = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
            var button = Ui.Button("读取当前选区", 120); button.Name = "summaryPick_" + kind;
            button.Click += delegate { if (!busy && !canUndo && SelectionRequested != null) SelectionRequested(kind); };
            selectionButtons.Add(button);
            panel.Controls.Add(Ui.Grip(field, title)); panel.Controls.Add(button);
            var pick = Ui.Button("在表中选择…", 116); pick.Name = "summaryRange_" + kind;
            pick.Click += delegate { if (!busy && !canUndo && RangeRequested != null) RangeRequested(kind); };
            selectionButtons.Add(pick); panel.Controls.Add(pick);
            return panel;
        }
        internal void ApplySelection(string kind, int row, int column, string description, string title)
        {
            if (kind == "lead") headerRow.Text = row.ToString(CultureInfo.InvariantCulture);
            else if (kind == "header") sheetHeaderRow.Text = row.ToString(CultureInfo.InvariantCulture);
            else if (kind == "year") YearRow.Value = row;
            else if (kind == "column") templateColumn.Text = SummaryColumns.Letter(column);
            if ((kind == "column" || kind == "header") && !string.IsNullOrWhiteSpace(title)) SetPlaceholder(title.Trim());
            InvalidatePreview(this, EventArgs.Empty);
            hint.Text = description;
            AppendLog(description + "；请重新预检。手动输入仍可使用。");
        }
        internal void SetSourceCandidates(IEnumerable<string> candidates)
        {
            string prior = sourceFile.Text;
            sourceFile.Items.Clear();
            foreach (string candidate in candidates) sourceFile.Items.Add(candidate);
            sourceFile.Text = prior;
        }
        internal void SetYearPattern(string pattern)
        {
            suspendInvalidate = true;
            try { YearPathPattern.Text = pattern ?? ""; }
            finally { suspendInvalidate = false; }
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
            MakeColumn("sourceFileName", "实际来源文件名", 250, 450);
            MakeColumn("sourceMapping", "映射依据", 140, 200);
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

        internal void SetWorkbooks(List<SummaryWorkbookChoice> choices, string selectedPath)
        {
            settingWorkbooks = true;
            try
            {
                Workbooks.Items.Clear();
                foreach (var choice in choices)
                {
                    int index = Workbooks.Items.Add(choice);
                    if (!string.IsNullOrEmpty(selectedPath) && string.Equals(choice.Path, selectedPath, StringComparison.OrdinalIgnoreCase)) Workbooks.SelectedIndex = index;
                }
            }
            finally { settingWorkbooks = false; }
        }
        internal void ClearWorkbookContext()
        {
            SetContext("", new List<string>(), "", "", "");
            sourceFile.Items.Clear(); sourceFile.Text = ""; Companies.Rows.Clear();
            YearPathPattern.Clear();
        }
        /// <summary>Only the explicitly selected workbook supplies this context.</summary>
        internal void SetContext(string workbook, List<string> sheets, string lead, string targetSheet, string columnLetter)
        {
            if (IsDisposed) return;
            previewOk = false;
            if (!string.Equals(workbookPath, workbook ?? "", StringComparison.OrdinalIgnoreCase))
            { Companies.Rows.Clear(); sourceFile.Items.Clear(); sourceFile.Text = ""; }
            workbookPath = workbook ?? "";
            workbookLabel.Text = "工作簿：" + (string.IsNullOrEmpty(workbook) ? "（尚未选择）" : workbook);
            leadSheet.Items.Clear();
            leadSheet.SelectedIndex = -1; leadSheet.Text = "";
            TargetSheet.Items.Clear(); TargetSheet.SelectedIndex = -1; TargetSheet.Text = "";
            if (sheets != null) foreach (string name in sheets) leadSheet.Items.Add(name);
            if (sheets != null) foreach (string name in sheets) TargetSheet.Items.Add(name);
            if (sheets == null || sheets.Count == 0)
            {
                hint.Text = "请选择已打开的工作簿、浏览文件，或主动点击“使用当前工作簿”，然后选择 Lead 和汇总工作表。";
                RefreshActions();
                return;
            }
            if (!string.IsNullOrEmpty(lead))
            {
                int index = leadSheet.Items.IndexOf(lead);
                leadSheet.SelectedIndex = index;
            }
            TargetSheet.SelectedIndex = TargetSheet.Items.IndexOf(targetSheet ?? "");
            if (!string.IsNullOrEmpty(columnLetter)) templateColumn.Text = columnLetter;
            hint.Text = "请明确选择 Lead 和汇总工作表后预检；切换其他活动工作簿不会改变本窗口的处理对象。";
            RefreshActions();
        }

        /// <summary>从界面收集请求（不含 COM）。</summary>
        internal SummaryRequest BuildRequest()
        {
            Companies.EndEdit();
            SummaryRequest request = new SummaryRequest();
            request.LeadSheet = Convert.ToString(leadSheet.Text).Trim();
            request.TargetSheet = TargetSheetName;
            request.WorkbookName = workbookPath;
            int row;
            if (!int.TryParse(headerRow.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out row)) row = 1;
            request.HeaderRow = row;
            int sheetRow;
            if (!int.TryParse(sheetHeaderRow.Text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out sheetRow)) sheetRow = 1;
            request.SheetHeaderRow = sheetRow;
            int column;
            if (!SummaryColumns.TryParse(templateColumn.Text, out column)) column = 0;
            request.TemplateColumn = column;
            request.Placeholder = placeholderBox.Text.Trim();
            string source = sourceFile.Text.Trim();
            if (source.Length > 0 && !source.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) source += ".xlsx";
            request.SourceFile = source;
            request.Annual = AnnualMode.Checked; request.FirstYear = (int)FirstYear.Value; request.LastYear = (int)LastYear.Value;
            request.YearHeaderRow = (int)YearRow.Value; request.DescendingYears = DescendingYears.Checked;
            request.YearPathPattern = YearPathPattern.Text.Trim(); request.CheckYearFiles = CheckYearFiles.Checked;
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

        internal string TargetSheetName { get { return Convert.ToString(TargetSheet.SelectedItem) ?? ""; } set { TargetSheet.SelectedIndex = TargetSheet.Items.IndexOf(value ?? ""); } }

        internal void SetCompanies(List<SummaryCompany> list)
        {
            Companies.Rows.Clear();
            if (list == null) { RefreshActions(); return; }
            foreach (SummaryCompany company in list)
            {
                int index = Companies.Rows.Add(company.Selected, company.Code, company.Name,
                    company.FileName, string.IsNullOrWhiteSpace(company.SourceFileName) ? "全称.xlsx（回退）" : "Lead 来源文件名",
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
            if (suspendInvalidate) return;
            bool hadPreview = previewOk;
            previewOk = false;
            RefreshActions();
            if (hadPreview) AppendLog("参数已改动，请重新点「读取 Lead 并预检」。");
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

        internal void SetLog(string text) { logView.Clear(); logView.AppendText(text ?? ""); if (ProgressSession != null) ProgressSession.State.Log(text ?? ""); }

        internal void AppendLog(string text)
        {
            if (IsDisposed) return;
            if (ProgressSession != null) ProgressSession.State.Log(text);
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
        private Dictionary<string, Control> PresetFields()
        {
            return new Dictionary<string, Control> {
                { "lead", leadSheet }, { "header", headerRow }, { "target", TargetSheet },
                { "targetHeader", sheetHeaderRow }, { "column", templateColumn }, { "placeholder", placeholderBox },
                { "sourceFile", sourceFile }, { "annual", AnnualMode }, { "firstYear", FirstYear },
                { "lastYear", LastYear }, { "yearRow", YearRow }, { "descending", DescendingYears },
                { "checkFiles", CheckYearFiles }, { "yearPattern", YearPathPattern }
            };
        }
        internal BatchPreset CapturePreset()
        { ValidateChildren(); return new BatchPreset { Feature = "summary", Context = workbookPath, Values = BatchPresetStore.Capture(PresetFields()) }; }
        internal void ApplyPreset(BatchPreset preset)
        {
            if (busy || canUndo) throw new InvalidOperationException("执行期间或有待还原结果时不能应用预设。");
            if (workbookPath.Length == 0) throw new InvalidOperationException("请先选择工作簿；原文件位置：" + preset.Context);
            ValidatePreset(preset, leadSheet.Items.Cast<string>().ToList());
            ApplyPresetValues(preset);
        }
        internal void ApplyPreparedPreset(BatchPreset preset, string path, List<string> sheets)
        {
            if (canUndo) throw new InvalidOperationException("有待还原结果时不能应用预设。");
            ValidatePreset(preset, sheets);
            SetContext(path, sheets, preset.Values["lead"], preset.Values["target"], "");
            ApplyPresetValues(preset);
        }
        private void ApplyPresetValues(BatchPreset preset)
        {
            var fields = PresetFields(); var values = preset.Values;
            suspendInvalidate = true;
            try { BatchPresetStore.Apply(fields, values); }
            finally { suspendInvalidate = false; }
            annualOptions.Visible = AnnualMode.Checked; Companies.Rows.Clear();
            InvalidatePreview(this, EventArgs.Empty);
            AppendLog("预设未携带公司名单；请刷新公司列表并重新预检。当前文件：" + workbookPath + "；保存时文件：" + preset.Context);
        }
        internal void ValidatePreset(BatchPreset preset, IList<string> sheets)
        {
            BatchPresetStore.Validate(preset, "summary");
            var fields = PresetFields(); var values = preset.Values; BatchPresetStore.ValidateFields(fields, values);
            if (!sheets.Contains(values["lead"]) || !sheets.Contains(values["target"])) throw new InvalidOperationException("工作簿缺少预设中的 Lead 或汇总工作表。");
            int n;
            foreach (string key in new[] { "header", "targetHeader" })
                if (!int.TryParse(values[key], out n) || n < 1 || n > 1048576) throw new InvalidOperationException("预设表头行无效：" + key);
            if (!SummaryColumns.TryParse(values["column"], out n)) throw new InvalidOperationException("预设模板列无效。");
        }
        internal void SetUndoAvailable(bool value) { canUndo = value; RefreshActions(); }
        internal void MarkUndone() { canUndo = false; RefreshActions(); }
        internal bool IsBusy { get { return busy; } }

        private void RefreshActions()
        {
            bool idle = !busy && !canUndo;
            if (presets != null) presets.Gear.Enabled = idle;
            Workbooks.Enabled = BrowseWorkbook.Enabled = CurrentWorkbook.Enabled = RefreshWorkbooks.Enabled = idle;
            TargetSheet.Enabled = idle && workbookPath.Length > 0;
            AnnualMode.Enabled = annualOptions.Enabled = idle;
            foreach (Control button in selectionButtons) button.Enabled = idle && workbookPath.Length > 0;
            leadSheet.Enabled = headerRow.Enabled = sheetHeaderRow.Enabled = templateColumn.Enabled = placeholderBox.Enabled = sourceFile.Enabled = idle;
            Companies.ReadOnly = !idle;
            Refresh.Enabled = idle && workbookPath.Length > 0;
            Preview.Enabled = idle && workbookPath.Length > 0 && leadSheet.Text.Length > 0 && TargetSheet.SelectedIndex >= 0;
            Execute.Enabled = idle && previewOk && Companies.Rows.Count > 0;
            Undo.Enabled = canUndo;
            Confirm.Enabled = !busy && canUndo;
            CopyLog.Enabled = CloseButton.Enabled = true;
        }

        public Size MeasureDefaultClient()
        {
            root.PerformLayout();
            return new Size(DialogSizing.LogicalClient(this).Width, 760);
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
