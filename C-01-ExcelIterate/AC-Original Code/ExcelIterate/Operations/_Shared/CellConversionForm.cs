using System;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Registered parameter windows share the same range controls and submission boundary.
    internal sealed class CellConversionForm : Form, IDialogDefaultLayout
    {
        private readonly CellConversionKind kind;
        private readonly TableLayoutPanel root;
        private readonly RadioButton extract, filter;
        private readonly CheckBox[] kinds;
        private readonly RadioButton[] styles;
        private readonly BatchInsertTrimDeletePanel textEdit;
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
        private readonly Label outcome;
        internal readonly RoundedButton Pick, Apply, Cancel;
        internal readonly WorksheetUndoButton Undo = new WorksheetUndoButton();
        internal event Action<bool, ConversionRequest> Submitted;
        internal event Action<long, string> PreviewRequested;
        internal event Action<long> UndoRequested;
        private bool pending;
        internal bool SuppressActivation;

        internal static string DialogId(CellConversionKind kind)
        {
            switch (kind)
            {
                case CellConversionKind.TextFilter: return "conversion.text";
                case CellConversionKind.BatchInsertTrimDelete: return "conversion.batchInsertTrimDelete";
                case CellConversionKind.ReferenceStyle: return "formula.references";
                case CellConversionKind.FormulaAsText: return "formula.formulaAsText";
                case CellConversionKind.RestoreTextFormula: return "formula.restoreTextFormula";
                case CellConversionKind.ResultAsText: return "formula.resultAsText";
                case CellConversionKind.ValuesOnly: return "formula.valuesOnly";
                default: throw new ArgumentOutOfRangeException("kind");
            }
        }

        private static string Caption(CellConversionKind kind)
        {
            switch (kind)
            {
                case CellConversionKind.TextFilter: return "文本提取与过滤";
                case CellConversionKind.BatchInsertTrimDelete: return "批量插入与删截";
                case CellConversionKind.ReferenceStyle: return "相对-绝对转换";
                case CellConversionKind.FormulaAsText: return "转为公式文本";
                case CellConversionKind.RestoreTextFormula: return "还原文本公式";
                case CellConversionKind.ResultAsText: return "转文本格式";
                default: return "仅保留数值";
            }
        }

        private static string Explanation(CellConversionKind kind)
        {
            switch (kind)
            {
                case CellConversionKind.FormulaAsText: return "将选区中的公式按原文保存为文本，保留原数字格式。";
                case CellConversionKind.RestoreTextFormula: return "将选区中以等号开头的文本还原为公式并计算。";
                case CellConversionKind.ResultAsText: return "将选区中的计算结果及常量按显示文字保存为文本。";
                default: return "将选区中的公式替换为计算结果，保留值类型及原数字格式。";
            }
        }

        internal CellConversionForm(CellConversionKind selectedKind, string description = null, string address = null, ConversionPreview initialPreview = null, string initialError = null)
        {
            kind = selectedKind; Text = Caption(kind);
            DialogSizing.Prepare(this, DialogId(kind), "table", Ui.Blue);
            KeyPreview = true; AutoScroll = true;
            root = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1, Padding = new Padding(24, 17, 24, 14), BackColor = Ui.Canvas };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); Controls.Add(root);
            if (kind == CellConversionKind.ReferenceStyle) Add(ConversionUi.Label(Text, 18, true));
            Add(ConversionUi.Hint(description ?? "工作簿：示例.xlsx\r\n工作表：数据"));
            var range = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 3, Margin = new Padding(0, 10, 0, 3) };
            range.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); range.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); range.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var rangeLabel = Label("单元格区域", 10, true); rangeLabel.Anchor = AnchorStyles.Left;
            RangeInput = new TextBox { Name = "conversionRange", AccessibleName = "单元格区域", Text = address ?? "A1:B10", Margin = new Padding(10, 0, 10, 0), MaxLength = 32767 };
            Pick = ConversionUi.Button("在表中选择…", 128); Pick.Name = "conversionPick"; Pick.Margin = Padding.Empty;
            var field = new RoundedEditor(RangeInput, 38) { Dock = DockStyle.Fill };
            range.Controls.Add(rangeLabel, 0, 0); range.Controls.Add(field, 1, 0); range.Controls.Add(Pick, 2, 0); Add(range);
            if (kind != CellConversionKind.BatchInsertTrimDelete) Add(ConversionUi.Hint("可输入 A1:B10 或 A1:A5,C1:C5；每批最多十万格，包含隐藏行列。"));
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
                Add(ConversionUi.Hint("按单元格显示的文字处理，结果以文本写回。空格类包含 Tab 和换行。"));
            }
            else if (kind == CellConversionKind.BatchInsertTrimDelete)
            {
                textEdit = new BatchInsertTrimDeletePanel(); textEdit.OptionsChanged += RefreshPreview; Add(textEdit);
                Add(ConversionUi.Hint("公式开头的 = 参与计数。移除后按普通内容写回。语法无效的公式需确认后才转为文本。"));
            }
            else if (kind == CellConversionKind.ReferenceStyle)
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
            else Add(ConversionUi.Hint(Explanation(kind)));
            beforePreview = Label("", 10, false); beforePreview.Name = "conversionBefore"; beforePreview.BackColor = Color.White; beforePreview.Padding = new Padding(10, 6, 10, 6); Add(beforePreview);
            preview = Label("", 10, true); preview.Name = "conversionPreview"; preview.BackColor = Ui.SelectBack; preview.ForeColor = Ui.SelectInk; preview.Padding = new Padding(10, 6, 10, 6); Add(preview);
            Validation = Label("", 9, false); Validation.Name = "conversionValidation"; Validation.ForeColor = Ui.Danger; Add(Validation);
            outcome = ConversionUi.Hint(""); outcome.Name = "conversionOutcome"; Add(outcome);
            actions = Ui.ActionBar(); actions.Dock = DockStyle.Top;
            Apply = ConversionUi.Button("应用转换", 108, true); Apply.Name = "conversionApply";
            Cancel = ConversionUi.Button("关闭", 96); Cancel.Name = "conversionCancel";
            Apply.Margin = new Padding(0, 0, 12, 0); Cancel.Margin = Padding.Empty;
            actions.Controls.AddRange(new Control[] { Apply, Undo, Cancel }); Add(actions);
            AcceptButton = null; CancelButton = Cancel;
            Pick.Click += delegate { Send(true); }; Apply.Click += delegate { Send(false); }; Cancel.Click += delegate { Close(); };
            Undo.Click += delegate
            {
                if (pending || Undo.State == null || UndoRequested == null) return;
                pending = true; Enabled = false;
                UndoRequested(Undo.State.Id);
            };
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
            if (description != null && address == null) { RangeInput.Enabled = Pick.Enabled = false; }
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
                WrapLabels(); Validation.Visible = Validation.Text.Length > 0; outcome.Visible = outcome.Text.Length > 0; root.PerformLayout();
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
            else if (kind == CellConversionKind.BatchInsertTrimDelete) request.TextEdit = textEdit.Options();
            else if (styles != null) for (int i = 0; i < styles.Length; i++) if (styles[i].Checked) request.Style = (ReferenceStyle)i;
            request.Validate(); return request;
        }
        private void RefreshPreview()
        {
            if (preview == null || Apply == null) return;
            try
            {
                beforePreview.Text = "修改前示例："; preview.Text = "修改后示例：";
                string before = sample == null ? "" : kind != CellConversionKind.TextFilter && sample.IsFormula ? sample.Formula : sample.Display;
                if (sample != null) beforePreview.Text += string.IsNullOrEmpty(before) ? "（空白）" : before;
                var request = Request(); Apply.Enabled = false; Validation.Text = "";
                if (readError != null) throw new InvalidOperationException(readError);
                if (sample == null || sample.Address != request.Address) { Validation.Text = "正在读取所选区域首格…"; return; }
                string after;
                if (kind == CellConversionKind.TextFilter)
                {
                    if (sample.IsNumeric && before.Trim().Length > 0 && before.Trim().Trim('#').Length == 0) throw new InvalidOperationException("数值显示为 ####，请先调宽该列后重试。");
                    after = CellCharacterClassifier.Filter(before, request.Kinds, request.Extract);
                }
                else if (kind == CellConversionKind.BatchInsertTrimDelete)
                {
                    if (sample.IsBlank) { after = ""; Validation.Text = "首格为空白，将跳过；执行时仍处理区域中的其他非空单元格。"; }
                    else
                    {
                        if (!sample.IsFormula && sample.IsNumeric && before.Trim().Length > 0 && before.Trim().Trim('#').Length == 0) throw new InvalidOperationException("数值显示为 ####，请先调宽该列后重试。");
                        after = BatchTextEdit.Transform(before, request.TextEdit);
                    }
                }
                else if (kind == CellConversionKind.ReferenceStyle)
                {
                    after = sample.IsFormula ? ReferenceStyleConverter.Convert(before, request.Style) : before;
                    if (!sample.IsFormula) Validation.Text = "首格不是公式；执行时仍会处理区域中的其他公式。";
                }
                else
                {
                    preview.Visible = false;
                    Apply.Enabled = !pending;
                    return;
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
        internal void Resume(string address, string error, string result = null, bool refresh = false)
        {
            if (address != null) { RangeInput.Text = address; revision++; sample = null; readError = null; }
            pending = false; Enabled = true;
            if (refresh && RangeInput.Enabled) { revision++; sample = null; readError = null; }
            if (result != null) outcome.Text = result;
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
