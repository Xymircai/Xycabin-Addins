using System;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Lightweight editor for the custom-function wrap. The user types a single
    // function template with one empty argument reserved for each cell's content.
    internal sealed class CustomFunctionForm : Form, IDialogDefaultLayout, IFormulaOperationDialog
    {
        internal readonly TextBox Template, Preview;
        internal readonly Label Hint, Validation;
        internal readonly RoundedButton Apply, Cancel;
        private readonly Label intro, inputLabel, previewLabel, slotLabel;
        private readonly ComboBox Slot;
        private readonly UiCard card;
        private readonly FlowLayoutPanel actions;
        private bool arranging, refreshingSlots;
        private readonly WidthGrip templateGrip, slotGrip;
        internal bool SuppressActivation;
        internal FormulaRequest Request { get; private set; }
        // Raised after the template validates and the window hides itself, so the
        // sheet is visible while the host STA runs the native range picker.
        public event Action<FormulaRequest> Submitted;
        public WorksheetUndoButton Undo { get; private set; }
        private bool canApply = true;

        internal CustomFunctionForm()
        {
            Text = "自定义函数 · 自定义封装";
            DialogSizing.Prepare(this, "formula.custom", "table", Ui.Blue);
            ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = false; KeyPreview = true;
            AutoScroll = true;
            intro = TextLabel("输入模板后点「应用到选区」，本窗口会暂时隐藏，请在表格中框选要处理的区域。原内容变化时停止。", 9, false);
            card = new UiCard();
            inputLabel = TextLabel("自定义函数模板", 10, true);
            Template = new TextBox { BorderStyle = BorderStyle.FixedSingle, TabIndex = 0, MaxLength = 32767 };
            Hint = TextLabel("", 9, false); Hint.ForeColor = Ui.InkDim;
            slotLabel = TextLabel("多参数选区填入", 10, true);
            Slot = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 1 };
            previewLabel = TextLabel("示例预览（文本 a-b-c 与公式 =A1+V2+3）", 9, true);
            // Word wrapping and no scroll bars keep the preview inside the default
            // size, so the window never asks the user to scroll.
            Preview = new TextBox { Multiline = true, ReadOnly = true, WordWrap = true, ScrollBars = ScrollBars.None,
                BorderStyle = BorderStyle.None, BackColor = Ui.GridHead, Font = Ui.MonoFont(10, Font), TabIndex = 2 };
            card.Controls.AddRange(new Control[] { inputLabel, Template, Hint, slotLabel, Slot, previewLabel, Preview });
            templateGrip=new WidthGrip(Template);
            slotGrip=new WidthGrip(Slot);
            Validation = TextLabel("", 9, false); Validation.ForeColor = Ui.Danger;
            actions = Ui.ActionBar(); actions.Dock = DockStyle.None; actions.AutoSize = false; actions.Padding = Padding.Empty;
            Apply = Ui.Button("应用到选区", 140); Ui.StylePrimary(Apply); Apply.TabIndex = 3;
            Undo = new WorksheetUndoButton();
            Cancel = Ui.Button("关闭", 96); Cancel.TabIndex = 5; Cancel.DialogResult = DialogResult.Cancel;
            actions.Controls.AddRange(new Control[] { Apply, Undo, Cancel });
            Controls.AddRange(new Control[] { intro, card, Validation, actions });
            AcceptButton = null; CancelButton = Cancel;
            Apply.Click += delegate { Submit(); };
            Cancel.Click += delegate { Request = null; DialogResult = DialogResult.Cancel; Close(); };
            Template.TextChanged += delegate { RefreshSlots(); UpdatePreview(); };
            Slot.SelectedIndexChanged += delegate { if (!refreshingSlots) UpdatePreview(); };
            Shown += delegate { Arrange(false); if (!SuppressActivation && !DialogPreviewFactory.IsFormatRendering(this)) { Template.Focus(); Template.SelectAll(); } };
            // Seed the picker so its disabled state is correct before the first keystroke.
            RefreshSlots();
            UpdatePreview();
        }

        public void SetTarget(string target, bool ready)
        {
            canApply = ready;
            if (!string.IsNullOrEmpty(target)) intro.Text = target;
            UpdatePreview();
        }

        public void Resume(string message)
        {
            Enabled = true;
            UpdatePreview();
            if (!string.IsNullOrEmpty(message)) { Validation.Text = message; Arrange(true); }
        }

        private static Label TextLabel(string text, float size, bool bold)
        { Label label = Ui.Label(text, size, bold); label.Dock = DockStyle.None; label.AutoEllipsis = false; return label; }

        private int Logical(int value)
        { using (Graphics g = CreateGraphics()) return Math.Max(1, (int)Math.Round(value * g.DpiX / 96f * Font.SizeInPoints / 10f)); }
        private static int TextHeight(Label label, int width)
        { return TextRenderer.MeasureText(label.Text, label.Font, new Size(Math.Max(1, width), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height; }

        // The template is written with or without the leading '='; slot counting must
        // see the same shape that FormulaRequest.Custom parses.
        private static string StripEquals(string text)
        {
            string expression = (text ?? "").Trim();
            return expression.StartsWith("=", StringComparison.Ordinal) ? expression.Substring(1).Trim() : expression;
        }
        private int SelectedSlot { get { return Slot.SelectedIndex < 0 ? 0 : Slot.SelectedIndex; } }

        // Rebuilds the slot picker as the template changes. The row is always shown so the
        // dialog does not jump; with a single slot there is nothing to choose, so the box
        // is greyed instead of hidden.
        private void RefreshSlots()
        {
            int count = FormulaWrapper.CountEmptyArguments(StripEquals(Template.Text));
            refreshingSlots = true;
            try
            {
                int keep = Slot.SelectedIndex;
                Slot.Items.Clear();
                for (int i = 0; i < count; i++) Slot.Items.Add("第 " + (i + 1) + " 个空参数位");
                Slot.Enabled = count > 1;
                if (count > 0) Slot.SelectedIndex = keep >= 0 && keep < count ? keep : 0;
            }
            finally { refreshingSlots = false; }
        }

        protected override void OnLayout(LayoutEventArgs e)
        { base.OnLayout(e); Arrange(false); }
        private static void Place(Control control, int x, int y, int width, int height)
        { Size wanted = DialogControlSizing.Resolve(control, new Size(width, height)); control.SetBounds(x, y, wanted.Width, wanted.Height); }
        private void Arrange(bool compact)
        {
            if (arranging || actions == null) return;
            compact = compact && !DialogSizing.HasOverride(this);
            arranging = true;
            try
            {
                int pad = Logical(12), gap = Logical(6), section = Logical(8), inset = Logical(10);
                int width = Math.Max(Logical(400), ClientSize.Width - pad * 2);
                int labelWidth = Math.Max(Logical(116), DialogControlSizing.Resolve(inputLabel, new Size(Logical(116), 1)).Width);
                int requiredEditor = DialogControlSizing.Resolve(Template, Size.Empty).Width;
                width = Math.Max(width, requiredEditor + labelWidth + gap + inset * 2);
                width = Math.Max(width, DialogControlSizing.Resolve(Preview, Size.Empty).Width + inset * 2);
                int y = pad;
                Place(intro, pad, y, width, TextHeight(intro, width)); y = intro.Bottom + section;
                int inner = width - inset * 2, editorLeft = inset + labelWidth + gap;
                int editorWidth = inner - labelWidth - gap, rowY = inset;
                int inputHeight = Template.GetPreferredSize(new Size(editorWidth, 0)).Height;
                Place(Template, editorLeft, rowY, editorWidth, inputHeight);
                Place(inputLabel, inset, rowY, labelWidth, inputHeight);
                rowY = Math.Max(Template.Bottom, inputLabel.Bottom) + gap;
                Place(Hint, editorLeft, rowY, editorWidth, TextHeight(Hint, editorWidth));
                rowY = Hint.Bottom + gap;
                int slotHeight = Math.Max(Slot.GetPreferredSize(new Size(editorWidth, 0)).Height, inputHeight);
                Place(Slot, editorLeft, rowY, editorWidth, slotHeight);
                Place(slotLabel, inset, rowY, labelWidth, slotHeight);
                rowY = Math.Max(Slot.Bottom, slotLabel.Bottom) + section;
                Place(previewLabel, inset, rowY, inner, TextHeight(previewLabel, inner));
                rowY = previewLabel.Bottom + gap;
                int previewHeight = Math.Max(Logical(40),
                    TextRenderer.MeasureText(Preview.Text, Preview.Font, new Size(Math.Max(1, inner), int.MaxValue),
                        TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height + Logical(8));
                Place(Preview, inset, rowY, inner, previewHeight);
                card.SetBounds(pad, y, width, Preview.Bottom + inset); y = card.Bottom + section;
                Validation.Visible = Validation.Text.Length > 0;
                Place(Validation, pad, y, width, Validation.Visible ? TextHeight(Validation, width) : 0);
                if (Validation.Visible) y = Validation.Bottom + section;
                Apply.Size = DialogControlSizing.Resolve(Apply, new Size(Logical(140), Logical(38))); Cancel.Size = DialogControlSizing.Resolve(Cancel, new Size(Logical(96), Logical(38)));
                Undo.Size = DialogControlSizing.Resolve(Undo, new Size(Logical(112), Logical(38)));
                Apply.Margin = new Padding(0, 0, section, 0); Cancel.Margin = Padding.Empty;
                actions.SetBounds(pad, y, width, Math.Max(Apply.Height, Cancel.Height)); actions.PerformLayout(); actions.Height = Math.Max(actions.Height, Math.Max(Undo.Bottom, Math.Max(Apply.Bottom, Cancel.Bottom)));
                int contentHeight = actions.Bottom + pad;
                AutoScrollMinSize = new Size(Math.Max(width + pad * 2, intro.Right + pad), contentHeight);
            }
            finally { arranging = false; }
        }

        public Size MeasureDefaultClient()
        { Arrange(false); return new Size(DialogSizing.LogicalClient(this).Width, (int)Math.Ceiling((actions.Bottom + Logical(12)) / DialogSizing.Scale(this))); }

        // Kept on screen once the template validates. Clearing the line made the dialog
        // look unfinished and left the next step unexplained.
        private const string AcceptedHint = "公式校验正确，下一步请点击「应用到选区」";

        private bool TryRequest(out FormulaRequest request)
        {
            request = null;
            try
            {
                request = FormulaRequest.Custom(Template.Text, SelectedSlot);
                Validation.Text = AcceptedHint;
                Validation.ForeColor = Ui.Green;
                return true;
            }
            catch (FormatException ex)
            {
                Validation.Text = ex.Message;
                Validation.ForeColor = Ui.Danger;
                return false;
            }
        }

        // Hides instead of closing: the host STA picks the range, and the window comes
        // back if the user cancels the picker.
        private void Submit()
        {
            FormulaRequest request;
            if (!canApply || Submitted == null || !TryRequest(out request)) return;
            Request = request;
            Enabled = false;
            Hide();
            Action<FormulaRequest> submit = Submitted;
            if (submit != null) submit(request);
        }

        private void UpdatePreview()
        {
            Hint.Text = "留空参数位表示每个单元格内容，例如 =TEXTSPLIT(,\"-\",1)。公式单元格填入公式体，文本/数值单元格填入其值。";
            FormulaRequest request;
            Apply.Enabled = TryRequest(out request) && canApply;
            Preview.Text = request == null ? "模板尚未通过检查" :
                "文本：=" + request.ApplyCustom(FormulaRequest.SerializeValue("a-b-c")) + "\r\n公式：" + FormulaWrapper.Transform("=A1+V2+3", request);
            Arrange(true);
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape) { Request = null; DialogResult = DialogResult.Cancel; Close(); e.Handled = true; e.SuppressKeyPress = true; }
            base.OnKeyDown(e);
        }
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult != DialogResult.OK) { Request = null; DialogResult = DialogResult.Cancel; }
            base.OnFormClosing(e);
        }
        protected override void Dispose(bool disposing)
        { if (disposing && Preview != null) Preview.Font.Dispose(); base.Dispose(disposing); }
    }
}
