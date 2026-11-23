using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class WorksheetBatchForm : Form
    {
        // The native ComboBox still owns the list, selection, focus and keyboard input.
        private sealed class BatchChoice : ComboBox, IDialogPreviewSelfPainted
        {
            [DllImport("user32.dll", EntryPoint = "SendMessageW")]
            private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
            private const int SetItemHeight = 0x0153;
            private bool hovered;
            internal Action<int> ScrollClosed;
            internal Action PrepareItems;

            internal BatchChoice()
            {
                DropDownStyle = ComboBoxStyle.DropDownList;
                DrawMode = DrawMode.OwnerDrawVariable;
                IntegralHeight = false;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }

            private void MeasureList()
            {
                using (var graphics = CreateGraphics())
                {
                    int rowHeight = ListRowHeight(graphics);
                    for (int i = 0; i < Items.Count; i++)
                        if (GetItemHeight(i) != rowHeight)
                            SendMessage(Handle, SetItemHeight, new IntPtr(i), new IntPtr(rowHeight));
                    int rows = Math.Min(Items.Count, MaxDropDownItems);
                    int limit = Math.Max(rowHeight, Screen.FromControl(this).WorkingArea.Height / 2 - 2);
                    int listHeight = Math.Max(1, Math.Min(rows, limit / rowHeight)) * rowHeight + 2;
                    if (DropDownHeight != listHeight) DropDownHeight = listHeight;
                }
            }

            private int ListRowHeight(Graphics graphics)
            { return (int)Math.Ceiling(Font.GetHeight(graphics) + 16 * graphics.DpiY / 96f); }

            protected override void OnMeasureItem(MeasureItemEventArgs e)
            {
                if (e.Index >= 0) e.ItemHeight = ListRowHeight(e.Graphics);
                base.OnMeasureItem(e);
            }

            protected override void OnDrawItem(DrawItemEventArgs e)
            {
                var graphics = e.Graphics;
                var bounds = e.Bounds;
                bool contrast = SystemInformation.HighContrast;
                bool inList = (e.State & DrawItemState.ComboBoxEdit) == 0;
                bool current = inList && e.Index >= 0 && e.Index == SelectedIndex;
                bool highlighted = inList && (e.State & DrawItemState.Selected) != 0;
                Color background = contrast ? SystemColors.Window : BackColor;
                if (contrast && highlighted) background = SystemColors.Highlight;
                Color ink = contrast ? (highlighted ? SystemColors.HighlightText : SystemColors.WindowText) : current ? Ui.Blue : ForeColor;
                using (var brush = new SolidBrush(background)) graphics.FillRectangle(brush, bounds);
                float scale = graphics.DpiX / 96f;
                if (!contrast && (current || highlighted) && bounds.Width > 8 * scale && bounds.Height > 4 * scale)
                {
                    var pill = new RectangleF(bounds.X + 4 * scale, bounds.Y + 2 * scale,
                        bounds.Width - 8 * scale, bounds.Height - 4 * scale);
                    var smoothing = graphics.SmoothingMode;
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var path = Ui.RoundRect(pill, 6 * scale))
                    using (var brush = new SolidBrush(current ? Ui.SelectBack : Ui.GridHead)) graphics.FillPath(brush, path);
                    graphics.SmoothingMode = smoothing;
                }
                int inset = (int)Math.Round((inList ? 16 : 3) * scale);
                if (e.Index >= 0 && e.Index < Items.Count && bounds.Width > inset * 2)
                    TextRenderer.DrawText(graphics, GetItemText(Items[e.Index]), Font, Rectangle.Inflate(bounds, -inset, 0), ink,
                        TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis |
                        TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
                if (contrast && inList) e.DrawFocusRectangle();
            }

            protected override void WndProc(ref Message message)
            {
                const int LeftButtonDown = 0x0201, KeyDown = 0x0100, SystemKeyDown = 0x0104, ShowDropDown = 0x014f;
                if (message.Msg == LeftButtonDown || message.Msg == KeyDown || message.Msg == SystemKeyDown ||
                    (message.Msg == ShowDropDown && message.WParam != IntPtr.Zero))
                    EnsureItems();
                const int MouseWheel = 0x020a;
                if (message.Msg == MouseWheel && ScrollClosed != null && !DroppedDown)
                {
                    ScrollClosed(unchecked((short)(message.WParam.ToInt64() >> 16)));
                    message.Result = IntPtr.Zero;
                    return;
                }
                base.WndProc(ref message);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var graphics = e.Graphics;
                bool contrast = SystemInformation.HighContrast;
                Color background = !Enabled ? SystemColors.Control : contrast ? SystemColors.Window : BackColor;
                Color ink = !Enabled ? SystemColors.GrayText : contrast ? SystemColors.WindowText : ForeColor;
                graphics.Clear(background);
                float scale = graphics.DpiX / 96f;
                int arrowWidth = Math.Min(ClientSize.Width, (int)Math.Round(24 * scale));
                int left = ClientSize.Width - arrowWidth;
                TextRenderer.DrawText(graphics, GetItemText(SelectedItem), Font,
                    new Rectangle(0, 0, Math.Max(0, left - 4), ClientSize.Height), ink,
                    TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis |
                    TextFormatFlags.NoPrefix | TextFormatFlags.PreserveGraphicsClipping);
                if (arrowWidth < 4 || ClientSize.Height < 4) return;

                Color arrowBack = contrast ? SystemColors.Control :
                    !Enabled ? Ui.GridHead : DroppedDown ? Ui.Blue : hovered ? Ui.Btn2Down : Ui.SelectBack;
                Color arrowInk = !Enabled ? SystemColors.GrayText :
                    contrast ? SystemColors.ControlText : DroppedDown ? Color.White : Ui.Blue;
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Ui.RoundRect(new RectangleF(left, 1, arrowWidth - 1, ClientSize.Height - 2), 5 * scale))
                using (var brush = new SolidBrush(arrowBack)) graphics.FillPath(brush, path);
                float x = left + arrowWidth / 2f, y = ClientSize.Height / 2f;
                float half = Math.Min(4 * scale, (arrowWidth - 2) / 2f);
                using (var stroke = new Pen(arrowInk, Math.Max(1, 1.5f * scale)))
                {
                    stroke.StartCap = stroke.EndCap = LineCap.Round;
                    graphics.DrawLines(stroke, new[] {
                        new PointF(x - half, y - half / 3), new PointF(x, y + half / 2),
                        new PointF(x + half, y - half / 3)
                    });
                }
            }

            protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); hovered = true; Invalidate(); }
            protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hovered = false; Invalidate(); }
            protected override void OnDropDown(EventArgs e) { EnsureItems(); MeasureList(); base.OnDropDown(e); Invalidate(); }
            protected override void OnDropDownClosed(EventArgs e) { base.OnDropDownClosed(e); Invalidate(); }
            protected override void OnSelectedIndexChanged(EventArgs e) { base.OnSelectedIndexChanged(e); Invalidate(); }

            internal void EnsureItems()
            {
                if (PrepareItems == null) return;
                PrepareItems();
                PrepareItems = null;
            }
        }

        private sealed class FormatViewport : Panel
        {
            private readonly Panel clip = new Panel { TabStop = false, Margin = Padding.Empty };
            private readonly FormatScrollBar scroll = new FormatScrollBar();
            private readonly TableLayoutPanel content;
            private int scrollWidth, scrollGap;
            private bool arranging;

            internal FormatViewport(TableLayoutPanel fields)
            {
                Name = "batchFormatViewport";
                AccessibleName = "可滚动的单元格格式选项";
                Dock = DockStyle.Top;
                Margin = new Padding(0, 0, 0, 4);
                TabStop = false;
                content = fields;
                content.Dock = DockStyle.None;
                content.AutoSize = false;
                clip.Controls.Add(content);
                Controls.Add(clip);
                Controls.Add(scroll);
                scroll.OffsetChanged += delegate { content.Top = -scroll.Offset; };
                clip.MouseWheel += ForwardWheel;
                scroll.MouseWheel += ForwardWheel;
                content.MouseWheel += ForwardWheel;
                foreach (Control item in content.Controls)
                {
                    item.MouseWheel += ForwardWheel;
                    Control field = item;
                    field.Enter += delegate { Reveal(field); };
                }
            }

            internal void SetMetrics(int fullHeight, int visibleHeight, int width, int gap, int lineHeight)
            {
                if (content.Height == fullHeight && Height == visibleHeight && scrollWidth == width && scrollGap == gap && scroll.LineHeight == lineHeight) return;
                scrollWidth = width;
                scrollGap = gap;
                scroll.LineHeight = lineHeight;
                content.Height = fullHeight;
                Height = visibleHeight;
                PerformLayout();
            }

            protected override void OnLayout(LayoutEventArgs e)
            {
                base.OnLayout(e);
                if (content == null || arranging) return;
                arranging = true;
                try
                {
                    clip.SetBounds(0, 0, Math.Max(0, ClientSize.Width - scrollWidth - scrollGap), ClientSize.Height);
                    scroll.SetBounds(clip.Right + scrollGap, 0, scrollWidth, ClientSize.Height);
                    scroll.SetExtent(content.Height, clip.Height);
                    content.SetBounds(0, -scroll.Offset, Math.Max(1, clip.Width), content.Height);
                }
                finally { arranging = false; }
            }

            private void Reveal(Control field)
            {
                if (clip.Height <= 0) return;
                if (field.Top < scroll.Offset) scroll.Offset = field.Top;
                else if (field.Bottom > scroll.Offset + clip.Height) scroll.Offset = field.Bottom - clip.Height;
            }

            protected override void OnMouseWheel(MouseEventArgs e)
            {
                ForwardWheel(this, e);
                base.OnMouseWheel(e);
            }

            private void ForwardWheel(object sender, MouseEventArgs e)
            {
                ScrollByWheel(e.Delta);
                var handled = e as HandledMouseEventArgs;
                if (handled != null) handled.Handled = true;
            }

            internal void ScrollByWheel(int delta)
            {
                int lines = SystemInformation.MouseWheelScrollLines;
                int distance = lines < 0 ? clip.Height : lines * scroll.LineHeight;
                scroll.Offset -= (int)Math.Round(delta * distance / 120d);
            }
        }

        private sealed class FormatScrollBar : Control
        {
            private int extent, viewportHeight, offset;
            private int grabOffset = -1;
            internal int LineHeight { get; set; }
            internal event Action OffsetChanged;

            internal FormatScrollBar()
            {
                Name = "batchFormatScroll";
                AccessibleName = "单元格格式滚动条";
                AccessibleRole = AccessibleRole.ScrollBar;
                TabStop = true;
                SetStyle(ControlStyles.Selectable | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }

            private int Limit { get { return Math.Max(0, extent - viewportHeight); } }
            internal int Offset
            {
                get { return offset; }
                set
                {
                    int next = Math.Max(0, Math.Min(Limit, value));
                    if (offset == next) return;
                    offset = next;
                    Invalidate();
                    if (OffsetChanged != null) OffsetChanged();
                }
            }

            internal void SetExtent(int fullHeight, int visibleHeight)
            {
                if (extent == fullHeight && viewportHeight == visibleHeight) return;
                extent = fullHeight;
                viewportHeight = visibleHeight;
                Offset = offset;
                Invalidate();
            }

            private Rectangle Track
            {
                get
                {
                    int inset = Math.Max(1, Width / 5);
                    return Rectangle.Inflate(ClientRectangle, -inset, -inset);
                }
            }

            private Rectangle Thumb(Rectangle track)
            {
                int height = Math.Min(track.Height, Math.Max(Width, (int)((long)track.Height * viewportHeight / Math.Max(1, extent))));
                int travel = track.Height - height;
                int top = track.Top + (int)((long)Offset * travel / Math.Max(1, Limit));
                return new Rectangle(track.Left, top, track.Width, height);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics;
                g.Clear(BackColor);
                Rectangle track = Track;
                if (track.Width <= 0 || track.Height <= 0) return;
                bool contrast = SystemInformation.HighContrast;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Ui.RoundRect(track, track.Width / 2f))
                using (var brush = new SolidBrush(contrast ? SystemColors.Control : Ui.SelectBack)) g.FillPath(brush, path);
                Rectangle thumb = Thumb(track);
                if (Limit > 0 && thumb.Height > 0)
                {
                    using (var path = Ui.RoundRect(thumb, Math.Min(thumb.Width, thumb.Height) / 2f))
                    using (var brush = new SolidBrush(contrast ? SystemColors.Highlight : Enabled ? Ui.Blue : SystemColors.GrayText)) g.FillPath(brush, path);
                }
                if (Focused) ControlPaint.DrawFocusRectangle(g, ClientRectangle);
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                if (e.Button != MouseButtons.Left || Limit == 0) return;
                Focus();
                Rectangle thumb = Thumb(Track);
                if (thumb.Contains(e.Location)) { grabOffset = e.Y - thumb.Top; Capture = true; }
                else Offset += e.Y < thumb.Top ? -viewportHeight : viewportHeight;
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                if (!Capture || grabOffset < 0) return;
                Rectangle track = Track;
                int travel = track.Height - Thumb(track).Height;
                if (travel > 0) Offset = (int)Math.Round((e.Y - track.Top - grabOffset) * (double)Limit / travel);
            }

            protected override void OnMouseUp(MouseEventArgs e)
            {
                base.OnMouseUp(e);
                if (e.Button == MouseButtons.Left) Capture = false;
            }

            protected override void OnMouseCaptureChanged(EventArgs e)
            {
                if (!Capture) grabOffset = -1;
                base.OnMouseCaptureChanged(e);
            }

            protected override bool IsInputKey(Keys keyData)
            {
                switch (keyData & Keys.KeyCode)
                {
                    case Keys.Up: case Keys.Down: case Keys.PageUp: case Keys.PageDown: case Keys.Home: case Keys.End: return true;
                    default: return base.IsInputKey(keyData);
                }
            }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                base.OnKeyDown(e);
                switch (e.KeyCode)
                {
                    case Keys.Up: Offset -= LineHeight; break;
                    case Keys.Down: Offset += LineHeight; break;
                    case Keys.PageUp: Offset -= viewportHeight; break;
                    case Keys.PageDown: Offset += viewportHeight; break;
                    case Keys.Home: Offset = 0; break;
                    case Keys.End: Offset = Limit; break;
                    default: return;
                }
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        private static readonly string[] NumberFormats = { null, "General", "#,##0", "#,##0.00", "0.00%", "yyyy-mm-dd", "@" };
        private static readonly WorksheetBatchHorizontal?[] HorizontalValues = { null, WorksheetBatchHorizontal.General,
            WorksheetBatchHorizontal.Left, WorksheetBatchHorizontal.Center, WorksheetBatchHorizontal.Right,
            WorksheetBatchHorizontal.Justify, WorksheetBatchHorizontal.Distributed };
        private static readonly WorksheetBatchVertical?[] VerticalValues =
            { null, WorksheetBatchVertical.Top, WorksheetBatchVertical.Center, WorksheetBatchVertical.Bottom };
        private static readonly Color ReviewHeaderBack = Color.FromArgb(253, 232, 235);
        private const TextFormatFlags ReviewText = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
            TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
        private readonly TextBox sheet = Input("batchSheetName", "工作表名称", 31);
        private readonly TextBox area = Input("batchSheetRange", "单元格区域", 32767);
        private readonly TextBox find = Input("batchFind", "查找文字", 32767);
        private readonly TextBox replacement = Input("batchReplacement", "替换为", 32767);
        private readonly BatchChoice source = Choice("batchSource", "目标工作簿", "已打开的工作簿", "指定文件夹", "指定多个文件");
        private readonly BatchChoice textStyle = Choice("batchTextStyle", "文本格式", "保持原样", "加粗", "斜体", "下划线", "删除线");
        private readonly BatchChoice fontFamily = InstalledFontChoice();
        private readonly BatchChoice horizontal = Choice("batchHorizontal", "水平对齐", "保持原样", "常规", "左对齐", "居中", "右对齐", "两端对齐", "分散对齐");
        private readonly BatchChoice vertical = Choice("batchVertical", "垂直对齐", "保持原样", "顶端", "居中", "底端");
        private readonly BatchChoice number = Choice("batchNumberFormat", "数字格式", "保持原样", "常规", "千分位 · 整数", "千分位 · 两位小数", "百分比 · 两位小数", "日期 · 年-月-日", "文本");
        private readonly BatchChoice font = Choice("batchFontColor", "字体颜色", "保持原样", "设置颜色");
        private readonly BatchChoice fill = Choice("batchFillColor", "填充颜色", "保持原样", "设置颜色", "无填充");
        private readonly CheckBox replaceText = new CheckBox { Name = "batchReplaceText", Text = "替换文字", AutoSize = true };
        private readonly CheckBox matchCase = new CheckBox { Name = "batchMatchCase", Text = "区分大小写", AutoSize = true };
        private readonly CheckBox wholeCell = new CheckBox { Name = "batchWholeCell", Text = "匹配整个单元格", AutoSize = true };
        private readonly CheckBox clearFormatting = new CheckBox { Name = "batchClearFormatting", Text = "清除格式", AutoSize = true,
            AccessibleDescription = "重置字体名称、文本样式、字体颜色、填充、数字格式和对齐。" };
        private readonly Label clearHint = ConversionUi.Hint("重置常用格式，保留字号、边框、条件格式和保护设置。");
        private readonly FormatViewport formatViewport;
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
        private readonly UiCard targetCard, textCard, formatCard, results;
        private readonly TableLayoutPanel targetFields, replaceFields, formatFields;
        private Color fontColor = Color.FromArgb(29, 78, 216), fillColor = Color.FromArgb(254, 243, 199);
        private WorksheetBatchReport preview;
        private string folder;
        private string[] files;
        private bool busy, measuring, layoutDirty;
        internal event Action PickRange;
        internal event Func<WorksheetBatchRequest, WorksheetBatchReport, WorksheetBatchReport> Run;

        internal WorksheetBatchForm(string initialSheet = null, string initialRange = null, bool sample = false)
        {
            Text = "工作表批量操作";
            SuspendLayout();
            DialogSizing.Prepare(this, "conversion.worksheetBatch", "table", Ui.Blue);
            AutoScroll = true;
            root = Table(1);
            root.Dock = DockStyle.Fill; root.AutoSize = false; root.Padding = new Padding(22, 16, 22, 12);

            parameters = Table(1);
            targetFields = FieldColumn(hasAction: true);
            AddFields(targetFields, "目标工作簿", source, selectSource);
            AddRow(targetFields, sourceDetail);
            AddFields(targetFields, "工作表名称", sheet);
            AddFields(targetFields, "单元格区域", area, pickRange);
            AddRow(targetFields, ConversionUi.Hint("可输入 A1:D20，或用逗号分隔多个区域。相同地址将用于各工作簿。"));
            targetCard = Card("目标与区域", targetFields);
            targetCard.Name = "batchTargetCard";
            AddRow(parameters, targetCard);

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
            formatViewport = new FormatViewport(formatFields);
            foreach (BatchChoice choice in new[] { textStyle, number, font, fill, fontFamily, horizontal, vertical })
                choice.ScrollClosed = formatViewport.ScrollByWheel;
            var formatting = Table(1);
            AddRow(formatting, formatViewport);
            AddRow(formatting, clearFormatting);
            AddRow(formatting, clearHint);

            var operations = TwoColumns();
            operations.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            textCard = Card("文字替换", textOptions);
            textCard.Name = "batchTextCard";
            textCard.Margin = new Padding(0, 0, 6, 7);
            textCard.Dock = DockStyle.Fill;
            formatCard = Card("单元格格式", formatting);
            formatCard.Name = "batchFormatCard";
            formatCard.Margin = new Padding(6, 0, 0, 7);
            formatCard.Dock = DockStyle.Fill;
            operations.Controls.Add(textCard, 0, 0);
            operations.Controls.Add(formatCard, 1, 0);
            AddRow(parameters, operations);
            AddRow(root, parameters);

            review.Columns.Add("工作簿", 235);
            review.Columns.Add("改动格数", 90);
            review.Columns.Add("预览结果", 380);
            review.SizeChanged += delegate { ResizeColumns(); };
            review.DrawColumnHeader += DrawReviewHeader;
            review.DrawSubItem += DrawReviewCell;
            review.ColumnWidthChanging += delegate(object sender, ColumnWidthChangingEventArgs e)
            {
                int minimum = MinimumColumn(e.ColumnIndex);
                if (e.NewWidth < minimum) { e.Cancel = true; e.NewWidth = minimum; }
            };
            results = Ui.Card(review, "批量预览");
            results.Margin = new Padding(0, 0, 0, 6);
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.Controls.Add(results, 0, root.RowCount++);

            status.Name = "batchStatus";
            AddRow(root, status);
            var actions = Ui.ActionBar();
            actions.FlowDirection = FlowDirection.RightToLeft;
            apply.Name = "batchApply"; check.Name = "batchCheck"; close.Name = "batchClose";
            actions.Controls.AddRange(new Control[] { apply, Undo, check, close });
            AddRow(root, actions);
            Controls.Add(root);
            CancelButton = close;
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (busy && e.CloseReason == CloseReason.UserClosing) e.Cancel = true; };
            ClientSizeChanged += delegate { RequestLayout(); };

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
            sheet.Text = initialSheet ?? "";
            area.Text = initialRange ?? "A1";
            UpdateSwatch(fontSwatch, fontColor, "字体颜色");
            UpdateSwatch(fillSwatch, fillColor, "填充颜色");
            if (sample)
            {
                replaceText.Checked = true; find.Text = "暂估"; replacement.Text = "已入账";
                textStyle.SelectedIndex = (int)WorksheetBatchTextStyle.Italic;
                fontFamily.EnsureItems();
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
            ResumeLayout(true);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                // Paint the cards and their native children together on the first frame.
                const int Composited = 0x02000000;
                var parameters = base.CreateParams;
                if (TopLevel) parameters.ExStyle |= Composited;
                return parameters;
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            SuspendLayout();
            try { base.OnLoad(e); RequestLayout(); }
            finally { ResumeLayout(true); }
        }

        private void RequestLayout()
        {
            if (results == null || measuring) return;
            layoutDirty = true;
            PerformLayout();
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            if (layoutDirty && !measuring)
            {
                layoutDirty = false;
                UpdateScrollExtent();
            }
            base.OnLayout(e);
        }

        private static TextBox Input(string name, string label, int limit)
        { return new TextBox { Name = name, AccessibleName = label, MaxLength = limit, Width = 120, Margin = Padding.Empty }; }

        private static BatchChoice Choice(string name, string label, params string[] choices)
        {
            var box = new BatchChoice { Name = name, AccessibleName = label, Margin = Padding.Empty, Width = 120 };
            box.Items.AddRange(choices); box.SelectedIndex = 0;
            return box;
        }

        private static BatchChoice InstalledFontChoice()
        {
            var box = Choice("batchFontFamily", "字体", "保持原样");
            box.MaxDropDownItems = 8;
            box.PrepareItems = delegate
            {
                box.Items.AddRange(InstalledFonts.Value);
                box.AutoCompleteSource = AutoCompleteSource.ListItems;
                box.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            };
            return box;
        }

        private static readonly Lazy<string[]> InstalledFonts = new Lazy<string[]>(GetInstalledFonts);

        private static string[] GetInstalledFonts()
        {
            var installed = new SortedDictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            foreach (FontFamily family in FontFamily.Families)
                using (family)
                {
                    string name = family.Name;
                    string canonical = WorksheetBatchFormat.CanonicalFontName(name);
                    if (!installed.ContainsKey(canonical)) installed.Add(canonical, name);
                }
            var choices = new List<string>();
            foreach (var preferred in new[] { "SimSun", "KaiTi", "Times New Roman" })
            {
                if (!installed.Remove(preferred)) continue;
                choices.Add(preferred == "SimSun" ? "宋体" : preferred == "KaiTi" ? "楷体" : preferred);
            }
            choices.AddRange(installed.Values);
            return choices.ToArray();
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
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            if (hasAction) table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute));
            return table;
        }

        private static void AddRow(TableLayoutPanel table, Control control)
        {
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.Controls.Add(control, 0, table.RowCount++);
            if (table.ColumnCount > 1) table.SetColumnSpan(control, table.ColumnCount);
        }

        private void AddFields(TableLayoutPanel table, string label, Control editor, Control action = null)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 37));
            var caption = ConversionUi.Label(label);
            caption.Dock = DockStyle.None;
            caption.AutoSize = false;
            caption.UseMnemonic = false;
            caption.Anchor = AnchorStyles.Left;
            caption.FontChanged += delegate { RequestLayout(); };
            caption.TextChanged += delegate { RequestLayout(); };
            editor.FontChanged += delegate { RequestLayout(); };
            table.Controls.Add(caption, 0, row);
            table.Controls.Add(new RoundedEditor(editor, 33) { Dock = DockStyle.Fill, Margin = new Padding(0, 1, 5, 3) }, 1, row);
            if (action != null)
            {
                action.Anchor = AnchorStyles.Left;
                action.Margin = new Padding(2, 0, 0, 2);
                action.FontChanged += delegate { RequestLayout(); };
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
            Color background = SystemInformation.HighContrast ? SystemColors.Control : ReviewHeaderBack;
            Color ink = SystemInformation.HighContrast ? SystemColors.ControlText : Ui.GridHeadInk;
            using (var fill = new SolidBrush(background)) e.Graphics.FillRectangle(fill, e.Bounds);
            using (var stroke = new Pen(Ui.GridFrame))
                e.Graphics.DrawLine(stroke, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            using (var title = new Font(review.Font, FontStyle.Bold))
                TextRenderer.DrawText(e.Graphics, e.Header.Text, title, e.Bounds, ink, ReviewText);
        }

        private void DrawReviewCell(object sender, DrawListViewSubItemEventArgs e)
        {
            bool selected = e.Item.Selected && (review.Focused || !review.HideSelection);
            bool contrast = SystemInformation.HighContrast;
            Color background = selected ? (contrast ? SystemColors.Highlight : Ui.SelectBack) : e.SubItem.BackColor;
            Color ink = selected ? (contrast ? SystemColors.HighlightText : Ui.SelectInk) : e.SubItem.ForeColor;
            using (var fill = new SolidBrush(background)) e.Graphics.FillRectangle(fill, e.Bounds);
            int padding = Math.Min((int)Math.Round(4 * e.Graphics.DpiX / 96f), Math.Max(0, (e.Bounds.Width - 1) / 2));
            Rectangle textBounds = Rectangle.Inflate(e.Bounds, -padding, 0);
            TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, textBounds, ink, ReviewText);
            if (review.Focused && e.Item.Focused && e.ColumnIndex == review.Columns.Count - 1)
                ControlPaint.DrawFocusRectangle(e.Graphics, e.Item.Bounds, ink, background);
        }

        private int FixedContentHeight(int width)
        {
            int content = root.Padding.Vertical;
            var available = new Size(Math.Max(1, width - root.Padding.Horizontal), 0);
            foreach (Control child in root.Controls)
                if (child != results) content += child.GetPreferredSize(available).Height + child.Margin.Vertical;
            return content;
        }

        private static int MeasureFields(TableLayoutPanel fields, Size minimum, int gap)
        {
            int captions = 0, editors = 0, buttons = 0;
            for (int row = 0; row < fields.RowCount; row++)
            {
                var frame = fields.GetControlFromPosition(1, row) as RoundedEditor;
                if (frame == null) continue;
                var caption = (Label)fields.GetControlFromPosition(0, row);
                Size text = TextRenderer.MeasureText(caption.Text, caption.Font, Size.Empty,
                    TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix) + caption.Padding.Size;
                Size saved = DialogControlSizing.Resolve(caption, text);
                Size captionSize = new Size(Math.Max(text.Width, saved.Width), Math.Max(text.Height, saved.Height));
                if (caption.Size != captionSize) caption.Size = captionSize;
                captions = Math.Max(captions, captionSize.Width + caption.Margin.Horizontal + gap);
                Size field = frame.MeasureField(minimum) + frame.Margin.Size;
                editors = Math.Max(editors, field.Width);
                int rowHeight = Math.Max(field.Height, captionSize.Height + caption.Margin.Vertical);
                var button = fields.ColumnCount == 3 ? fields.GetControlFromPosition(2, row) as RoundedButton : null;
                if (button != null)
                {
                    Size preferred = TextRenderer.MeasureText(button.Text, button.Font, Size.Empty,
                        TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix) + new Size(gap * 2, gap);
                    Size buttonSize = DialogControlSizing.Resolve(button,
                        new Size(Math.Max(button.Width, preferred.Width), Math.Max(button.Height, preferred.Height)));
                    if (button.Size != buttonSize) button.Size = buttonSize;
                    rowHeight = Math.Max(rowHeight, buttonSize.Height + button.Margin.Vertical);
                    buttons = Math.Max(buttons, buttonSize.Width + button.Margin.Horizontal);
                }
                var style = fields.RowStyles[row];
                if (style.SizeType != SizeType.Absolute) style.SizeType = SizeType.Absolute;
                if (style.Height != rowHeight) style.Height = rowHeight;
            }
            if (fields.ColumnStyles[0].Width != captions) fields.ColumnStyles[0].Width = captions;
            if (fields.ColumnCount == 3 && fields.ColumnStyles[2].Width != buttons) fields.ColumnStyles[2].Width = buttons;
            return captions + editors + buttons;
        }

        private static void WrapDetails(TableLayoutPanel fields, int width)
        {
            foreach (Control item in fields.Controls)
                if (item is Label && fields.GetColumnSpan(item) > 1)
                    item.MaximumSize = new Size(Math.Max(1, width - item.Margin.Horizontal), 0);
        }

        private void UpdateScrollExtent()
        {
            if (measuring || results == null) return;
            measuring = true;
            SuspendLayout();
            root.SuspendLayout();
            targetFields.SuspendLayout();
            replaceFields.SuspendLayout();
            formatFields.SuspendLayout();
            try
            {
                float scale = DialogSizing.Scale(this);
                var minimum = new Size((int)Math.Round(96 * scale), (int)Math.Round(33 * scale));
                int gap = (int)Math.Round(8 * scale), scrollWidth = (int)Math.Round(16 * scale), scrollGap = (int)Math.Round(4 * scale);
                int targetWidth = MeasureFields(targetFields, minimum, gap) + targetCard.Padding.Horizontal;
                int textWidth = MeasureFields(replaceFields, minimum, gap) + textCard.Padding.Horizontal + textCard.Margin.Horizontal;
                int formatWidth = MeasureFields(formatFields, minimum, gap) + formatCard.Padding.Horizontal +
                    formatCard.Margin.Horizontal + scrollWidth + scrollGap;
                int minimumWidth = root.Padding.Horizontal + Math.Max(targetWidth, 2 * Math.Max(textWidth, formatWidth));
                int width = Math.Max(ClientSize.Width, minimumWidth);
                int contentWidth = width - root.Padding.Horizontal;
                int cardWidth = contentWidth / 2;
                WrapDetails(targetFields, contentWidth - targetCard.Padding.Horizontal);
                WrapDetails(replaceFields, cardWidth - textCard.Padding.Horizontal - textCard.Margin.Horizontal);
                clearHint.MaximumSize = new Size(Math.Max(1, cardWidth - formatCard.Padding.Horizontal - formatCard.Margin.Horizontal), 0);

                int fullHeight = 0, visibleHeight = 0;
                for (int row = 0; row < formatFields.RowCount; row++)
                {
                    fullHeight += (int)formatFields.RowStyles[row].Height;
                    if (row < 4) visibleHeight = fullHeight;
                }
                formatViewport.SetMetrics(fullHeight, visibleHeight, scrollWidth, scrollGap, (int)Math.Round(14 * scale));
                var extent = new Size(minimumWidth, FixedContentHeight(width) + (int)Math.Round(148 * scale));
                if (AutoScrollMinSize != extent) AutoScrollMinSize = extent;
            }
            finally
            {
                formatFields.ResumeLayout(true);
                replaceFields.ResumeLayout(true);
                targetFields.ResumeLayout(true);
                root.ResumeLayout(true);
                ResumeLayout(true);
                measuring = false;
            }
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
            replaceFields.Enabled = replaceText.Checked;
            formatFields.Enabled = !clearFormatting.Checked;
            fontSwatch.Enabled = font.SelectedIndex == 1;
            fillSwatch.Enabled = fill.SelectedIndex == 1;
            parameters.Enabled = !busy;
            check.Enabled = !busy;
            Undo.Enabled = !busy && Undo.State != null;
            apply.Enabled = !busy && preview != null && preview.Cells > 0;
            close.Enabled = !busy;
            RequestLayout();
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
