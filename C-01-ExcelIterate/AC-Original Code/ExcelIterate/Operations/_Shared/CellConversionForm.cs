using System;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Two registered parameter windows share the same range controls and submission boundary.
    internal sealed class CellConversionForm : Form, IDialogDefaultLayout
    {
        private readonly CellConversionKind kind;
        private readonly TableLayoutPanel root;
        private readonly RadioButton extract, filter;
        private readonly CheckBox[] kinds;
        private readonly RadioButton[] styles;
        private readonly Label beforePreview, preview;
        private readonly FlowLayoutPanel actions;
        private readonly Timer previewTimer;
        private ConversionPreview sample;
        private string readError;
        private long revision;
        private bool arranging, shown, userSized, resizing;
        private Size resizeStart;
        internal readonly TextBox RangeInput;
        internal readonly Label Validation;
        internal readonly RoundedButton Pick, Apply, Cancel;
        internal event Action<bool, ConversionRequest> Submitted;
        internal event Action<long, string> PreviewRequested;
        private bool pending;
        internal bool SuppressActivation;

        internal CellConversionForm(CellConversionKind selectedKind, string description = null, string address = null, ConversionPreview initialPreview = null, string initialError = null)
        {
            if (selectedKind != CellConversionKind.TextFilter && selectedKind != CellConversionKind.ReferenceStyle) throw new ArgumentException("此命令无需参数窗口。");
            kind = selectedKind; Text = kind == CellConversionKind.TextFilter ? "文本提取与过滤" : "相对-绝对转换";
            DialogSizing.Prepare(this, kind == CellConversionKind.TextFilter ? "conversion.text" : "formula.references", "table", Ui.Blue);
            KeyPreview = true; AutoScroll = true;
            root = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(16, 16, 16, 12), BackColor = Ui.Canvas };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); Controls.Add(root);
            Add(Label(Text, 16, true));
            Add(Label(description ?? "工作簿：示例.xlsx\r\n工作表：数据", 9, false));
            var range = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 3, Margin = new Padding(0, 8, 0, 8) };
            range.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); range.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); range.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var rangeLabel = Label("单元格区域", 10, true); rangeLabel.Anchor = AnchorStyles.Left;
            RangeInput = new TextBox { Name = "conversionRange", Text = address ?? "A1:B10", Dock = DockStyle.Fill, Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(8, 6, 8, 6), MaxLength = 32767 };
            Pick = Ui.Button("在表中选择…", 128); Pick.Name = "conversionPick";
            range.Controls.Add(rangeLabel, 0, 0); range.Controls.Add(RangeInput, 1, 0); range.Controls.Add(Pick, 2, 0); Add(range);
            Add(Label("可输入 A1:B10 或 A1:A5,C1:C5；每批最多十万格，包含隐藏行列。", 9, false));
            if (kind == CellConversionKind.TextFilter)
            {
                var modes = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 8, 0, 4) };
                extract = new RadioButton { Name = "conversionExtract", Text = "提取所选类型", AutoSize = true, Checked = true, Margin = new Padding(0, 4, 24, 4) };
                filter = new RadioButton { Name = "conversionFilter", Text = "过滤所选类型", AutoSize = true, Margin = new Padding(0, 4, 0, 4) };
                modes.Controls.AddRange(new Control[] { extract, filter }); Add(modes);
                var choices = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Margin = new Padding(0, 4, 0, 8) };
                string[] labels = { "空格", "英文", "中文", "数字", "其他" }; kinds = new CheckBox[labels.Length];
                for (int i = 0; i < labels.Length; i++)
                {
                    kinds[i] = new CheckBox { Name = "conversionKind" + i, Text = labels[i], AutoSize = true, Checked = i == 2, Margin = new Padding(0, 4, 20, 4) };
                    kinds[i].CheckedChanged += delegate { RefreshPreview(); }; choices.Controls.Add(kinds[i]);
                }
                extract.CheckedChanged += delegate { RefreshPreview(); }; Add(choices);
                Add(Label("文本、数字、日期及公式结果按显示文字筛选，结果以文本覆盖原区域。空格包括 Tab 和换行。", 9, false));
            }
            else
            {
                var choices = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Margin = new Padding(0, 8, 0, 8) };
                string[] labels = { "行列均绝对", "列相对行绝对", "行相对列绝对", "行列均相对" }; styles = new RadioButton[labels.Length];
                for (int i = 0; i < labels.Length; i++)
                {
                    styles[i] = new RadioButton { Name = "conversionReference" + i, Text = labels[i], AutoSize = true, Checked = i == 3, Margin = new Padding(0, 3, 0, 3) };
                    styles[i].CheckedChanged += delegate { RefreshPreview(); }; choices.Controls.Add(styles[i]);
                }
                Add(choices); Add(Label("统一修改真实单元格引用的 $ 锁定标记；普通文本和非公式内容保持原样。", 9, false));
            }
            beforePreview = Label("", 10, true); beforePreview.Name = "conversionBefore"; beforePreview.BackColor = Ui.GridHead; beforePreview.Padding = new Padding(8); Add(beforePreview);
            preview = Label("", 10, true); preview.Name = "conversionPreview"; preview.BackColor = Ui.GridHead; preview.Padding = new Padding(8); Add(preview);
            Validation = Label("", 9, false); Validation.Name = "conversionValidation"; Validation.ForeColor = Ui.Danger; Add(Validation);
            actions = Ui.ActionBar(); actions.Dock = DockStyle.Top;
            Apply = Ui.Button("执行", 100); Apply.Name = "conversionApply"; Ui.StylePrimary(Apply);
            Cancel = Ui.Button("取消", 96); Cancel.Name = "conversionCancel";
            Apply.Margin = new Padding(0, 0, 12, 0); Cancel.Margin = Padding.Empty;
            actions.Controls.AddRange(new Control[] { Apply, Cancel }); Add(actions);
            AcceptButton = null; CancelButton = Cancel;
            Pick.Click += delegate { Send(true); }; Apply.Click += delegate { Send(false); }; Cancel.Click += delegate { Close(); };
            previewTimer = new Timer { Interval = 250 };
            previewTimer.Tick += delegate { previewTimer.Stop(); RequestPreview(); };
            RangeInput.TextChanged += delegate
            {
                revision++; sample = null; readError = null; previewTimer.Stop(); RefreshPreview();
                if (!pending) previewTimer.Start();
            };
            SizeChanged += delegate { Arrange(false); };
            ResizeBegin += delegate { resizeStart = ClientSize; resizing = true; };
            ResizeEnd += delegate { if (ClientSize != resizeStart) userSized = true; resizing = false; };
            Shown += delegate { shown = true; Arrange(true); if (!SuppressActivation && !DialogPreviewFactory.IsFormatRendering(this)) RangeInput.Focus(); };
            sample = initialPreview; readError = initialError;
            RefreshPreview();
        }
        private static Label Label(string text, float size, bool bold)
        { var label = Ui.Label(text, size, bold); label.AutoSize = true; label.Dock = DockStyle.Top; label.AutoEllipsis = false; label.Margin = new Padding(0, 4, 0, 4); return label; }
        private void Add(Control control) { root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.Controls.Add(control, 0, root.RowCount++); }
        private void WrapLabels()
        { if (root == null) return; foreach (Control control in root.Controls) if (control is Label) control.MaximumSize = new Size(Math.Max(1, ClientSize.Width - root.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth), 0); }
        private void Arrange(bool fit)
        {
            if (arranging || actions == null) return; arranging = true;
            try
            {
                WrapLabels(); Validation.Visible = Validation.Text.Length > 0; root.PerformLayout();
                int height = root.GetPreferredSize(new Size(ClientSize.Width, 0)).Height;
                AutoScrollMinSize = new Size(0, height);
                if (fit && shown && !resizing && !userSized && !DialogSizing.HasOverride(this) && WindowState == FormWindowState.Normal)
                {
                    Rectangle area = Screen.FromControl(this).WorkingArea;
                    int frame = Height - ClientSize.Height, margin = (int)Math.Round(24 * DialogSizing.Scale(this));
                    // A scrollbar changes ClientSize.Width without changing the outer width.
                    // Adjust only height so a long sample cannot progressively narrow the window.
                    Height = frame + Math.Min(Math.Max((int)Math.Round(320 * DialogSizing.Scale(this)), height), Math.Max(1, area.Height - frame - margin));
                    if (TopLevel && Bottom > area.Bottom) Top = Math.Max(area.Top, area.Bottom - Height);
                }
            }
            finally { arranging = false; }
        }
        internal ConversionRequest Request()
        {
            var request = new ConversionRequest { Kind = kind, Address = ConversionRangeAddress.Normalize(RangeInput.Text) };
            if (kind == CellConversionKind.TextFilter)
            {
                request.Extract = extract.Checked; request.Kinds = CellCharacterKinds.None;
                for (int i = 0; i < kinds.Length; i++) if (kinds[i].Checked) request.Kinds |= (CellCharacterKinds)(1 << i);
            }
            else for (int i = 0; i < styles.Length; i++) if (styles[i].Checked) request.Style = (ReferenceStyle)i;
            request.Validate(); return request;
        }
        private void RefreshPreview()
        {
            if (preview == null || Apply == null) return;
            try
            {
                beforePreview.Text = "修改前示例："; preview.Text = "修改后示例：";
                var request = Request(); Apply.Enabled = false; Validation.Text = "";
                if (readError != null) throw new InvalidOperationException(readError);
                if (sample == null || sample.Address != request.Address) { Validation.Text = "正在读取所选区域首格…"; return; }
                string before = kind == CellConversionKind.ReferenceStyle && sample.IsFormula ? sample.Formula : sample.Display;
                beforePreview.Text += string.IsNullOrEmpty(before) ? "（空白）" : before;
                string after;
                if (kind == CellConversionKind.TextFilter)
                {
                    if (sample.IsNumeric && before.Trim().Length > 0 && before.Trim().Trim('#').Length == 0) throw new InvalidOperationException("数值显示为 ####，请先调宽该列后重试。");
                    after = CellCharacterClassifier.Filter(before, request.Kinds, request.Extract);
                }
                else
                {
                    after = sample.IsFormula ? ReferenceStyleConverter.Convert(before, request.Style) : before;
                    if (!sample.IsFormula) Validation.Text = "首格不是公式；执行时仍会处理区域中的其他公式。";
                }
                preview.Text += string.IsNullOrEmpty(after) ? "（空白）" : after;
                Apply.Enabled = !pending;
            }
            catch (Exception error) { Apply.Enabled = false; Validation.Text = error.Message; }
            finally { Arrange(true); }
        }
        private void RequestPreview()
        {
            if (pending || IsDisposed) return;
            try
            {
                string address = ConversionRangeAddress.Normalize(RangeInput.Text);
                var read = PreviewRequested; if (read != null) read(revision, address);
            }
            catch (Exception error) { readError = error.Message; sample = null; RefreshPreview(); }
        }
        internal void ReceivePreview(long version, string address, ConversionPreview value, string error)
        {
            if (pending || version != revision || IsDisposed) return;
            string current; try { current = ConversionRangeAddress.Normalize(RangeInput.Text); } catch (FormatException) { return; }
            if (address != current) return;
            sample = value; readError = error; RefreshPreview();
        }
        private void Send(bool pick)
        {
            if (pending || DialogPreviewFactory.IsConstructing || DialogPreviewFactory.IsPreview(this) || DialogPreviewFactory.IsFormatRendering(this)) return;
            if (!pick && !Apply.Enabled) return;
            try
            {
                ConversionRequest request = null;
                if (pick)
                {
                    try { request = new ConversionRequest { Kind = kind, Address = ConversionRangeAddress.Normalize(RangeInput.Text) }; }
                    catch (FormatException) { } // An invalid typed address must not prevent choosing a valid range.
                }
                else request = Request();
                var submit = Submitted; if (submit == null) return;
                previewTimer.Stop(); pending = true; Enabled = false; Hide(); submit(pick, request);
            }
            catch (Exception error) { Resume(null, error.Message); }
        }
        internal void Resume(string address, string error)
        {
            if (address != null) { RangeInput.Text = address; revision++; sample = null; readError = null; }
            pending = false; Enabled = true;
            if (address != null || sample == null) RequestPreview();
            RefreshPreview(); if (!string.IsNullOrEmpty(error)) Validation.Text = error; Arrange(true);
            Visible = true; if (!SuppressActivation) { if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal; Activate(); }
        }
        public Size MeasureDefaultClient()
        { Arrange(false); return new Size(700, (int)Math.Ceiling(root.GetPreferredSize(new Size(ClientSize.Width, 0)).Height / DialogSizing.Scale(this))); }
        protected override void Dispose(bool disposing)
        { if (disposing && previewTimer != null) previewTimer.Dispose(); base.Dispose(disposing); }
    }
}
