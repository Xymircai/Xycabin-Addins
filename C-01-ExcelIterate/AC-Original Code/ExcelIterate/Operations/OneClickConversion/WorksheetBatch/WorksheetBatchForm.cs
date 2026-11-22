using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class WorksheetBatchForm : Form, IDialogDefaultLayout
    {
        private sealed class FormatViewport : Panel
        {
            internal void ScrollFromChoice(object sender, MouseEventArgs e)
            {
                if (!((ComboBox)sender).DroppedDown) OnMouseWheel(e);
            }
        }

        private static readonly string[] NumberFormats = { null, "General", "#,##0", "#,##0.00", "0.00%", "yyyy-mm-dd", "@" };
        private static readonly WorksheetBatchHorizontal?[] HorizontalValues = { null, WorksheetBatchHorizontal.General,
            WorksheetBatchHorizontal.Left, WorksheetBatchHorizontal.Center, WorksheetBatchHorizontal.Right,
            WorksheetBatchHorizontal.Justify, WorksheetBatchHorizontal.Distributed };
        private static readonly WorksheetBatchVertical?[] VerticalValues =
            { null, WorksheetBatchVertical.Top, WorksheetBatchVertical.Center, WorksheetBatchVertical.Bottom };
        private readonly TextBox sheet = Input("batchSheetName", "工作表名称", 31);
        private readonly TextBox area = Input("batchSheetRange", "单元格区域", 32767);
        private readonly TextBox find = Input("batchFind", "查找文字", 32767);
        private readonly TextBox replacement = Input("batchReplacement", "替换为", 32767);
        private readonly ComboBox source = Choice("batchSource", "工作簿来源", "已打开的工作簿", "指定文件夹", "指定多个文件");
        private readonly ComboBox textStyle = Choice("batchTextStyle", "文本格式", "保持原样", "加粗", "斜体", "下划线", "删除线");
        private readonly ComboBox fontFamily = InstalledFontChoice();
        private readonly ComboBox horizontal = Choice("batchHorizontal", "水平对齐", "保持原样", "常规", "左对齐", "居中", "右对齐", "两端对齐", "分散对齐");
        private readonly ComboBox vertical = Choice("batchVertical", "垂直对齐", "保持原样", "顶端", "居中", "底端");
        private readonly ComboBox number = Choice("batchNumberFormat", "数字格式", "保持原样", "常规", "千分位 · 整数", "千分位 · 两位小数", "百分比 · 两位小数", "日期 · 年-月-日", "文本");
        private readonly ComboBox font = Choice("batchFontColor", "字体颜色", "保持原样", "设置颜色");
        private readonly ComboBox fill = Choice("batchFillColor", "填充颜色", "保持原样", "设置颜色", "无填充");
        private readonly CheckBox replaceText = new CheckBox { Name = "batchReplaceText", Text = "替换文字", AutoSize = true };
        private readonly CheckBox matchCase = new CheckBox { Name = "batchMatchCase", Text = "区分大小写", AutoSize = true };
        private readonly CheckBox wholeCell = new CheckBox { Name = "batchWholeCell", Text = "匹配整个单元格", AutoSize = true };
        private readonly CheckBox clearFormatting = new CheckBox { Name = "batchClearFormatting", Text = "清除格式", AutoSize = true,
            AccessibleDescription = "重置字体名称、文本样式、字体颜色、填充、数字格式和对齐。" };
        private readonly Label clearHint = ConversionUi.Hint("重置常用格式，保留字号、边框、条件格式和保护设置。");
        private readonly FormatViewport formatViewport = new FormatViewport
        {
            Name = "batchFormatViewport", AccessibleName = "可滚动的单元格格式选项",
            AutoScroll = true, Dock = DockStyle.Top, Height = 154, Margin = new Padding(0, 0, 0, 4),
            Padding = new Padding(0, 0, 4, 0), TabStop = false
        };
        private readonly Label sourceDetail = ConversionUi.Hint("");
        private readonly Label status = ConversionUi.Hint("设置目标和操作后，先预览。");
        private readonly ListView review = new ListView { Name = "batchReview", AccessibleName = "工作簿预览", View = View.Details,
            FullRowSelect = true, ShowItemToolTips = true, UseCompatibleStateImageBehavior = false, MultiSelect = false, OwnerDraw = true,
            BorderStyle = BorderStyle.None, Dock = DockStyle.Fill };
        private readonly RoundedButton selectSource = ConversionUi.Button("选择文件…", 116);
        private readonly RoundedButton pickRange = ConversionUi.Button("在表中选择…", 126);
        private readonly RoundedButton fontSwatch = ConversionUi.Button("", 96);
        private readonly RoundedButton fillSwatch = ConversionUi.Button("", 96);
        private readonly RoundedButton check = ConversionUi.Button("预览", 96);
        private readonly RoundedButton apply = ConversionUi.Button("应用批量操作", 136, true);
        private readonly RoundedButton close = ConversionUi.Button("关闭", 80);
        internal readonly WorksheetUndoButton Undo = new WorksheetUndoButton();
        private readonly TableLayoutPanel root, parameters;
        private readonly UiCard results;
        private readonly TableLayoutPanel replaceFields, formatFields;
        private Color fontColor = Color.FromArgb(29, 78, 216), fillColor = Color.FromArgb(254, 243, 199);
        private WorksheetBatchReport preview;
        private string folder;
        private string[] files;
        private bool busy;
        internal event Action PickRange;
        internal event Func<WorksheetBatchRequest, WorksheetBatchReport, WorksheetBatchReport> Run;

        internal WorksheetBatchForm(string initialSheet = null, string initialRange = null, bool sample = false)
        {
            Text = "工作表批量操作";
            DialogSizing.Prepare(this, "conversion.worksheetBatch", "table", Ui.Blue);
            AutoScroll = true;
            root = Table(1);
            root.Dock = DockStyle.Fill; root.AutoSize = false; root.Padding = new Padding(22, 16, 22, 12);
            Controls.Add(root);

            parameters = Table(1);
            var target = Table(3);
            target.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            target.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            target.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132));
            AddFields(target, "来源", source, selectSource);
            AddRow(target, sourceDetail);
            AddFields(target, "工作表名称", sheet);
            AddFields(target, "单元格区域", area, pickRange);
            var hint = ConversionUi.Hint("可输入 A1:D20，或用逗号分隔多个区域。相同地址将用于各工作簿。");
            AddRow(target, hint);
            AddRow(parameters, Card("目标工作簿与区域", target));

            replaceFields = FieldColumn();
            AddFields(replaceFields, "查找", find);
            AddFields(replaceFields, "替换为", replacement);
            var matches = Ui.ActionBar();
            matches.BackColor = Color.White;
            matches.Controls.AddRange(new Control[] { matchCase, wholeCell });
            AddRow(replaceFields, matches);
            AddRow(replaceFields, ConversionUi.Hint("数值按底层值查找。留空可清空全部匹配内容，公式与日期时间保留。"));
            var textOptions = Table(1);
            AddRow(textOptions, replaceText); AddRow(textOptions, replaceFields);

            formatFields = FieldColumn(hasAction: true);
            AddFields(formatFields, "文本格式", textStyle);
            AddFields(formatFields, "数字格式", number);
            AddFields(formatFields, "字体颜色", font, fontSwatch);
            AddFields(formatFields, "填充颜色", fill, fillSwatch);
            AddFields(formatFields, "字体", fontFamily);
            AddFields(formatFields, "水平对齐", horizontal);
            AddFields(formatFields, "垂直对齐", vertical);
            foreach (ComboBox choice in new[] { textStyle, number, font, fill, fontFamily, horizontal, vertical })
                choice.MouseWheel += formatViewport.ScrollFromChoice;
            formatViewport.Controls.Add(formatFields);
            var formatting = Table(1);
            AddRow(formatting, formatViewport);
            AddRow(formatting, clearFormatting);
            AddRow(formatting, clearHint);

            var operations = TwoColumns();
            operations.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var textCard = Card("文字替换", textOptions);
            textCard.Name = "batchTextCard";
            textCard.Margin = new Padding(0, 0, 6, 7);
            textCard.Dock = DockStyle.Fill;
            var formatCard = Card("单元格格式", formatting);
            formatCard.Name = "batchFormatCard";
            formatCard.Margin = new Padding(6, 0, 0, 7);
            formatCard.Dock = DockStyle.Fill;
            operations.Controls.Add(textCard, 0, 0);
            operations.Controls.Add(formatCard, 1, 0);
            AddRow(parameters, operations);
            AddRow(root, parameters);

            review.Columns.Add("工作簿", 235);
            review.Columns.Add("改动格数", 90, HorizontalAlignment.Right);
            review.Columns.Add("预览结果", 380);
            review.SizeChanged += delegate { ResizeColumns(); };
            review.DrawColumnHeader += DrawReviewHeader;
            review.DrawSubItem += delegate(object sender, DrawListViewSubItemEventArgs e) { e.DrawDefault = true; };
            review.ColumnWidthChanging += delegate(object sender, ColumnWidthChangingEventArgs e)
            {
                int minimum = MinimumColumn(e.ColumnIndex);
                if (e.NewWidth < minimum) { e.Cancel = true; e.NewWidth = minimum; }
            };
            results = Ui.Card(review, "批量预览");
            results.Margin = new Padding(0, 0, 0, 6);
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(results, 0, root.RowCount++);
            Resize += delegate { ResizeColumns(); };

            status.Name = "batchStatus";
            AddRow(root, status);
            var actions = Ui.ActionBar();
            actions.FlowDirection = FlowDirection.RightToLeft;
            apply.Name = "batchApply"; check.Name = "batchCheck"; close.Name = "batchClose";
            actions.Controls.AddRange(new Control[] { apply, Undo, check, close });
            AddRow(root, actions);
            CancelButton = close;
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (busy && e.CloseReason == CloseReason.UserClosing) e.Cancel = true; };

            foreach (ComboBox choice in new[] { source, textStyle, fontFamily, horizontal, vertical, number, font, fill })
                choice.SelectedIndexChanged += delegate { InvalidatePreview(); };
            foreach (TextBox input in new[] { sheet, area, find, replacement }) input.TextChanged += delegate { InvalidatePreview(); };
            foreach (CheckBox choice in new[] { replaceText, matchCase, wholeCell, clearFormatting }) choice.CheckedChanged += delegate { InvalidatePreview(); };
            selectSource.Click += delegate { SelectSource(); };
            pickRange.Click += delegate { if (PickRange != null) PickRange(); };
            fontSwatch.Click += delegate { ChooseColor(false); };
            fillSwatch.Click += delegate { ChooseColor(true); };
            check.Click += delegate { RunBatch(false); };
            apply.Click += delegate { RunBatch(true); };
            close.Click += delegate { Close(); };
            Shown += delegate { UpdateScrollExtent(); };
            ClientSizeChanged += delegate { UpdateScrollExtent(); };
            sheet.Text = initialSheet ?? "";
            area.Text = initialRange ?? "A1";
            UpdateSwatch(fontSwatch, fontColor, "字体颜色");
            UpdateSwatch(fillSwatch, fillColor, "填充颜色");
            if (sample)
            {
                replaceText.Checked = true; find.Text = "暂估"; replacement.Text = "已入账";
                textStyle.SelectedIndex = (int)WorksheetBatchTextStyle.Italic;
                fontFamily.SelectedItem = fontFamily.Items.Contains("Times New Roman") ? "Times New Roman" : "保持原样";
                horizontal.SelectedIndex = Array.IndexOf(HorizontalValues, WorksheetBatchHorizontal.Justify);
                number.SelectedIndex = 3; fill.SelectedIndex = 1;
                var sampleReport = new WorksheetBatchReport();
                sampleReport.Rows.Add(new WorksheetBatchRow { Workbook = "华东-底稿.xlsx", Changed = 18, Detail = "可处理" });
                sampleReport.Rows.Add(new WorksheetBatchRow { Workbook = "华南-底稿.xlsx", Changed = 12, Detail = "可处理" });
                sampleReport.Rows.Add(new WorksheetBatchRow { Workbook = "华北-底稿.xlsx", Detail = "缺少工作表「Notes」。" });
                ShowReport(sampleReport, false);
            }
            RefreshControls();
        }

        private static TextBox Input(string name, string label, int limit)
        { return new TextBox { Name = name, AccessibleName = label, MaxLength = limit, Width = 120, Margin = Padding.Empty }; }

        private static ComboBox Choice(string name, string label, params string[] choices)
        {
            var box = new ComboBox { Name = name, AccessibleName = label, DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = Padding.Empty, Width = 120 };
            box.Items.AddRange(choices); box.SelectedIndex = 0;
            return box;
        }

        private static ComboBox InstalledFontChoice()
        {
            var box = Choice("batchFontFamily", "字体", "保持原样");
            var installed = new SortedDictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (FontFamily family in FontFamily.Families)
                using (family)
                {
                    string name = family.Name;
                    string canonical = WorksheetBatchFormat.CanonicalFontName(name);
                    if (!installed.ContainsKey(canonical)) installed.Add(canonical, name);
                }
            foreach (var preferred in new[] { "SimSun", "KaiTi", "Times New Roman" })
            {
                if (!installed.Remove(preferred)) continue;
                box.Items.Add(preferred == "SimSun" ? "宋体" : preferred == "KaiTi" ? "楷体" : preferred);
            }
            foreach (string name in installed.Values) box.Items.Add(name);
            box.AutoCompleteSource = AutoCompleteSource.ListItems;
            box.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            box.MaxDropDownItems = 12;
            return box;
        }

        private static TableLayoutPanel Table(int columns)
        {
            return new TableLayoutPanel { ColumnCount = columns, Dock = DockStyle.Top, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = Padding.Empty };
        }

        private static TableLayoutPanel TwoColumns()
        {
            var table = Table(2);
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            return table;
        }

        private static TableLayoutPanel FieldColumn(bool hasAction = false)
        {
            var table = Table(hasAction ? 3 : 2);
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 83));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            if (hasAction) table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            return table;
        }

        private static void AddRow(TableLayoutPanel table, Control control)
        {
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(control, 0, table.RowCount++);
            if (table.ColumnCount > 1) table.SetColumnSpan(control, table.ColumnCount);
        }

        private static void AddFields(TableLayoutPanel table, string label, Control editor, Control action = null)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var caption = ConversionUi.Label(label);
            caption.Anchor = AnchorStyles.Left;
            table.Controls.Add(caption, 0, row);
            table.Controls.Add(new RoundedEditor(editor, 33) { Dock = DockStyle.Fill, Margin = new Padding(0, 1, 5, 3) }, 1, row);
            if (action != null)
            {
                action.Anchor = AnchorStyles.Left;
                action.Margin = new Padding(2, 0, 0, 2);
                table.Controls.Add(action, 2, row);
            }
            else if (table.ColumnCount == 3) table.SetColumnSpan(table.GetControlFromPosition(1, row), 2);
        }

        private static UiCard Card(string title, Control content)
        {
            var card = new UiCard { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top,
                Margin = new Padding(0, 0, 0, 7), Padding = new Padding(13, 7, 13, 9) };
            var table = Table(1);
            AddRow(table, ConversionUi.Label(title, 10, true));
            AddRow(table, content);
            card.Controls.Add(table);
            return card;
        }

        private void ResizeColumns()
        {
            int width = review.ClientSize.Width - SystemInformation.VerticalScrollBarWidth;
            if (width < 300) return;
            review.Columns[0].Width = Math.Max(MinimumColumn(0), (int)(width * .38));
            review.Columns[1].Width = Math.Max(90, MinimumColumn(1));
            review.Columns[2].Width = Math.Max(MinimumColumn(2), width - review.Columns[0].Width - review.Columns[1].Width);
        }

        private int MinimumColumn(int column)
        {
            using (var title = new Font(review.Font, FontStyle.Bold))
                return TextRenderer.MeasureText(review.Columns[column].Text, title).Width + Ui.Px(this, 14);
        }

        private void DrawReviewHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            Color background = SystemInformation.HighContrast ? SystemColors.Control : Ui.GridHead;
            Color ink = SystemInformation.HighContrast ? SystemColors.ControlText : Ui.GridHeadInk;
            using (var fill = new SolidBrush(background)) e.Graphics.FillRectangle(fill, e.Bounds);
            using (var stroke = new Pen(Ui.GridFrame))
                e.Graphics.DrawLine(stroke, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            using (var title = new Font(review.Font, FontStyle.Bold))
                TextRenderer.DrawText(e.Graphics, e.Header.Text, title, e.Bounds, ink,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
        }

        public Size MeasureDefaultClient()
        {
            float scale = DialogSizing.Scale(this);
            return new Size(900, (int)Math.Ceiling(FixedContentHeight() / scale) + 216);
        }

        private int FixedContentHeight()
        {
            int content = root.Padding.Vertical;
            foreach (Control child in root.Controls)
                if (child != results) content += child.GetPreferredSize(new Size(Math.Max(1, ClientSize.Width - root.Padding.Horizontal), 0)).Height + child.Margin.Vertical;
            return content;
        }

        private void UpdateScrollExtent()
        {
            if (root == null) return;
            int cardWidth = (ClientSize.Width - root.Padding.Horizontal) / 2;
            clearHint.MaximumSize = new Size(Math.Max(100, cardWidth - Ui.Px(this, 48)), 0);
            int height = FixedContentHeight() + Ui.Px(this, 148);
            if (height != AutoScrollMinSize.Height) AutoScrollMinSize = new Size(0, height);
        }

        private WorksheetBatchRequest Request()
        {
            var result = new WorksheetBatchRequest
            {
                Source = (WorksheetBatchSource)source.SelectedIndex, Folder = folder, Files = files,
                SheetName = sheet.Text, Address = area.Text,
                Find = replaceText.Checked ? find.Text : null, Replacement = replacement.Text,
                MatchCase = matchCase.Checked, WholeCell = wholeCell.Checked, ClearFormatting = clearFormatting.Checked
            };
            if (result.ClearFormatting) return result;
            result.TextStyle = (WorksheetBatchTextStyle)textStyle.SelectedIndex;
            result.FontName = fontFamily.SelectedIndex > 0 ? WorksheetBatchFormat.CanonicalFontName((string)fontFamily.SelectedItem) : null;
            result.Horizontal = HorizontalValues[horizontal.SelectedIndex];
            result.Vertical = VerticalValues[vertical.SelectedIndex];
            result.NumberFormat = NumberFormats[number.SelectedIndex];
            result.FontColor = font.SelectedIndex == 1 ? ColorTranslator.ToOle(fontColor) : (int?)null;
            result.FillColor = fill.SelectedIndex == 1 ? ColorTranslator.ToOle(fillColor) : (int?)null;
            result.ClearFill = fill.SelectedIndex == 2;
            return result;
        }

        private void RefreshControls()
        {
            bool inFolder = source.SelectedIndex == (int)WorksheetBatchSource.Folder;
            bool inFiles = source.SelectedIndex == (int)WorksheetBatchSource.Files;
            selectSource.Visible = inFolder || inFiles;
            selectSource.Text = inFolder ? "选择文件夹…" : "选择文件…";
            sourceDetail.Text = inFolder ? folder ?? "扫描所选文件夹的本层文件。" :
                inFiles ? (files == null ? "可在文件选择器中多选。" : "已选择 " + files.Length + " 个文件") :
                "仅处理当前 WPS 实例中已打开的工作簿。";
            sourceDetail.AutoEllipsis = true;
            sourceDetail.MaximumSize = new Size(Math.Max(200, ClientSize.Width - 84), 0);
            replaceFields.Enabled = replaceText.Checked;
            formatFields.Enabled = !clearFormatting.Checked;
            fontSwatch.Enabled = font.SelectedIndex == 1;
            fillSwatch.Enabled = fill.SelectedIndex == 1;
            parameters.Enabled = !busy;
            check.Enabled = !busy;
            Undo.Enabled = !busy && Undo.State != null;
            apply.Enabled = !busy && preview != null && preview.Cells > 0;
            close.Enabled = !busy;
            if (Visible) UpdateScrollExtent();
        }

        private void InvalidatePreview()
        {
            preview = null;
            review.Items.Clear();
            status.ForeColor = Ui.InkDim;
            status.Text = "设置目标和操作后，先预览。";
            RefreshControls();
        }

        internal void SetRange(string sheetName, string address)
        { sheet.Text = sheetName; area.Text = address; }

        internal void ShowError(string message)
        {
            preview = null;
            status.ForeColor = Ui.Danger;
            status.Text = message;
            RefreshControls();
        }

        internal void ShowUndoResult(string message)
        {
            InvalidatePreview();
            status.Text = message;
            RefreshControls();
        }

        private void SelectSource()
        {
            if (source.SelectedIndex == (int)WorksheetBatchSource.Folder)
            {
                string chosen = HostDialogService.Folder(this, "选择工作簿文件夹", false);
                if (chosen == null) return;
                folder = chosen;
            }
            else
            {
                string[] chosen = HostDialogService.OpenFiles(this, WorksheetBatchRequest.FileFilter, "选择要批量处理的工作簿", true, "");
                if (chosen == null || chosen.Length == 0) return;
                files = chosen;
            }
            InvalidatePreview();
        }

        private void ChooseColor(bool background)
        {
            var current = background ? fillColor : fontColor;
            int? picked = HostDialogService.PickColor(this, current.ToArgb() & 0xFFFFFF);
            if (!picked.HasValue) return;
            Color selected = Color.FromArgb(unchecked((int)0xFF000000) | picked.Value);
            if (background) fillColor = selected; else fontColor = selected;
            UpdateSwatch(background ? fillSwatch : fontSwatch, background ? fillColor : fontColor,
                background ? "填充颜色" : "字体颜色");
            InvalidatePreview();
        }

        private static void UpdateSwatch(RoundedButton button, Color color, string label)
        {
            button.Text = string.Format("#{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B);
            button.AccessibleName = label + " " + button.Text;
            button.BaseColor = button.HoverColor = button.DownColor = color;
            button.ForeColor = (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) > 135 ? Color.Black : Color.White;
            button.Invalidate();
        }

        private void RunBatch(bool applyChanges)
        {
            var run = Run;
            if (busy || run == null || (applyChanges && preview == null)) return;
            busy = true; UseWaitCursor = true;
            status.Text = applyChanges ? "正在复核并应用批量操作…" : "正在检查工作簿…";
            status.ForeColor = Ui.InkDim;
            RefreshControls(); Update();
            try { ShowReport(run(Request(), applyChanges ? preview : null), applyChanges); }
            catch (Exception error) { ShowError(error.Message); }
            finally { busy = false; UseWaitCursor = false; RefreshControls(); }
        }

        private void ShowReport(WorksheetBatchReport report, bool applied)
        {
            preview = applied ? null : report;
            review.BeginUpdate();
            try
            {
                review.Items.Clear();
                foreach (WorksheetBatchRow row in report.Rows)
                {
                    var item = new ListViewItem(System.IO.Path.GetFileName(row.Workbook));
                    item.SubItems.Add(row.Changed > 0 ? row.Changed.ToString() : "—");
                    string state = applied && row.Changed > 0 ? row.Detail.Replace("可处理", "已应用") : row.Detail;
                    item.SubItems.Add(state);
                    item.ToolTipText = row.Workbook + "\r\n" + state;
                    review.Items.Add(item);
                }
            }
            finally { review.EndUpdate(); }
            status.ForeColor = Ui.InkDim;
            status.Text = applied ? "已修改 " + report.Workbooks + " 个工作簿、" + report.Cells + " 个单元格。请按需保存工作簿。"
                : report.Cells > 0 ? "预计修改 " + report.Workbooks + " 个工作簿、" + report.Cells + " 个单元格。应用后保留工作簿供你核对和保存。"
                : "没有可应用的改动。请查看各工作簿的预览结果。";
            if (report.CleanupWarning != null) { status.ForeColor = Ui.Danger; status.Text += "\r\n收尾时出现错误：" + report.CleanupWarning; }
        }
    }
}
