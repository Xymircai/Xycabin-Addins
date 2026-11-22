using System;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class FormulaOptionsForm : Form, IDialogDefaultLayout, IFormulaOperationDialog
    {
        private FormulaOperation operation;
        private readonly bool simple;
        private bool canApply = true;
        private string targetDescription;
        internal readonly ComboBox OperatorChoice;
        internal readonly ComboBox ValueType;
        internal readonly TextBox ValueInput, Preview;
        internal readonly Label Hint, Validation;
        internal readonly RoundedButton Apply, Cancel;
        public WorksheetUndoButton Undo { get; private set; }
        public event Action<FormulaRequest> Submitted;
        private readonly Label intro, typeLabel, inputLabel, previewLabel, operatorLabel;
        private readonly UiCard card;
        private readonly FlowLayoutPanel actions;
        private bool arranging;
        private readonly WidthGrip valueTypeGrip, valueInputGrip;
        internal bool SuppressActivation;
        internal FormulaRequest Request { get; private set; }

        internal FormulaOptionsForm(FormulaOperation selectedOperation)
        {
            operation = selectedOperation;
            simple = operation == FormulaOperation.Abs || operation == FormulaOperation.Unwrap;
            Text = operation >= FormulaOperation.Add ? "数值计算" : operation == FormulaOperation.Abs ? "ABS封装" :
                operation == FormulaOperation.Unwrap ? "移除公式封装" : operation == FormulaOperation.Round ? "ROUND封装" : "IFERROR封装";
            string id = operation >= FormulaOperation.Add ? "formula.add" : "formula." + operation.ToString().ToLowerInvariant();
            DialogSizing.Prepare(this, id, "table", Ui.Blue);
            ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = false; KeyPreview = true;
            AutoScroll = true;
            intro = TextLabel("可继续操作表格；确认只处理打开窗口时的选区。原内容变化时停止，取消不修改。", 9, false);
            card = new UiCard();
            if (operation >= FormulaOperation.Add)
            {
                operatorLabel = TextLabel("运算方式", 10, true);
                OperatorChoice = new ComboBox { Name = "formulaOperator", DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat };
                OperatorChoice.Items.AddRange(new object[] { "加", "减", "乘", "除" });
                OperatorChoice.SelectedIndex = (int)operation - (int)FormulaOperation.Add;
                card.Controls.AddRange(new Control[] { operatorLabel, OperatorChoice });
            }
            typeLabel = TextLabel(operation >= FormulaOperation.Add ? "参数类型" : "替代值类型", 10, true);
            inputLabel = TextLabel(operation >= FormulaOperation.Add ? "计算参数" : operation == FormulaOperation.IFError ? "出错时返回" : "保留位数", 10, true);
            ValueType = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = Color.White, TabIndex = 0 };
            ValueType.Items.AddRange(operation >= FormulaOperation.Add ? new object[] { "数值", "公式 / 引用" } : new object[] { "文字", "数值", "公式 / 引用" });
            ValueType.SelectedIndex = 0;
            ValueType.Visible = typeLabel.Visible = operation != FormulaOperation.Round;
            ValueInput = new TextBox { BorderStyle = BorderStyle.FixedSingle, TabIndex = 1, MaxLength = 32767, Text = operation >= FormulaOperation.Add ? FormulaRequest.DefaultNumber(operation) : operation == FormulaOperation.Round ? "2" : "" };
            Hint = TextLabel("", 9, false); Hint.ForeColor = Ui.InkDim;
            previewLabel = TextLabel(operation >= FormulaOperation.Add ? "示例预览（数值 20 与公式 =A1+V2+3）" : "示例预览（原公式：=A1+V2+3）", 9, true);
            Preview = new TextBox { Multiline = true, ReadOnly = true, WordWrap = false, ScrollBars = ScrollBars.Both,
                BorderStyle = BorderStyle.None, BackColor = Ui.GridHead, Font = Ui.MonoFont(10, Font), TabIndex = 2 };
            card.Controls.AddRange(new Control[] { typeLabel, ValueType, inputLabel, ValueInput, Hint, previewLabel, Preview });
            if (simple) { inputLabel.Visible = ValueInput.Visible = typeLabel.Visible = ValueType.Visible = false; }
            valueTypeGrip=new WidthGrip(ValueType);
            valueInputGrip=new WidthGrip(ValueInput);
            Validation = TextLabel("", 9, false); Validation.ForeColor = Ui.Danger;
            actions = Ui.ActionBar(); actions.Dock = DockStyle.None; actions.AutoSize = false; actions.Padding = Padding.Empty;
            Apply = Ui.Button("应用到选区", 140); Ui.StylePrimary(Apply); Apply.TabIndex = 3;
            Undo = new WorksheetUndoButton();
            Cancel = Ui.Button("关闭", 96); Cancel.TabIndex = 5; Cancel.DialogResult = DialogResult.Cancel;
            actions.Controls.AddRange(new Control[] { Apply, Undo, Cancel });
            Controls.AddRange(new Control[] { intro, card, Validation, actions });
            AcceptButton = null; CancelButton = Cancel;
            Apply.Click += delegate
            {
                FormulaRequest request;
                if (!canApply || Submitted == null || !TryRequest(out request)) return;
                Request = request; Enabled = false;
                Submitted(request);
            };
            Cancel.Click += delegate { Request = null; DialogResult = DialogResult.Cancel; Close(); };
            ValueInput.TextChanged += delegate { UpdatePreview(); };
            ValueType.SelectedIndexChanged += delegate { UpdatePreview(); };
            if (OperatorChoice != null)
                OperatorChoice.SelectedIndexChanged += delegate
                {
                    operation = (FormulaOperation)((int)FormulaOperation.Add + OperatorChoice.SelectedIndex);
                    ValueType.SelectedIndex = 0;
                    ValueInput.Text = FormulaRequest.DefaultNumber(operation);
                    UpdatePreview();
                };
            Shown += delegate { Arrange(false); if(!SuppressActivation&&!DialogPreviewFactory.IsFormatRendering(this)){ValueInput.Focus();ValueInput.SelectAll();} };
            UpdatePreview();
        }

        public void SetTarget(string target, bool ready)
        {
            canApply = ready;
            targetDescription = target;
            if (!string.IsNullOrEmpty(targetDescription)) intro.Text = targetDescription;
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

        // Font and native preferred sizes participate in measurement. Fixed logical
        // gaps are applied once, after DPI/font scaling, rather than fixed row heights.
        private int Logical(int value)
        { using (Graphics g = CreateGraphics()) return Math.Max(1, (int)Math.Round(value * g.DpiX / 96f * Font.SizeInPoints / 10f)); }
        private static int TextHeight(Label label, int width)
        { return TextRenderer.MeasureText(label.Text, label.Font, new Size(Math.Max(1, width), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height; }

        protected override void OnLayout(LayoutEventArgs e)
        { base.OnLayout(e); Arrange(false); }
        private static void Place(Control control,int x,int y,int width,int height)
        {Size wanted=DialogControlSizing.Resolve(control,new Size(width,height));control.SetBounds(x,y,wanted.Width,wanted.Height);}
        private void Arrange(bool compact)
        {
            if (arranging || actions == null) return;
            compact = compact && !DialogSizing.HasOverride(this);
            arranging = true;
            try
            {
                int pad = Logical(16), gap = Logical(8), section = Logical(12), inset = Logical(12);
                int width = Math.Max(Logical(400), ClientSize.Width - pad * 2);
                int labelWidth=Math.Max(Logical(106),Math.Max(DialogControlSizing.Resolve(typeLabel,new Size(Logical(106),1)).Width,DialogControlSizing.Resolve(inputLabel,new Size(Logical(106),1)).Width));
                int requiredEditor=Math.Max(DialogControlSizing.Resolve(ValueInput,Size.Empty).Width,ValueType.Visible?DialogControlSizing.Resolve(ValueType,Size.Empty).Width:0);
                width=Math.Max(width,requiredEditor+labelWidth+gap+inset*2);
                width=Math.Max(width,DialogControlSizing.Resolve(Preview,Size.Empty).Width+inset*2);
                int y = pad;
                Place(intro,pad, y, width, TextHeight(intro, width)); y = intro.Bottom + section;
                int inner = width - inset * 2, editorLeft = inset + labelWidth + gap;
                int editorWidth = inner - labelWidth - gap, rowY = inset;
                if (OperatorChoice != null)
                {
                    Place(OperatorChoice, editorLeft, rowY, editorWidth, OperatorChoice.PreferredHeight);
                    Place(operatorLabel, inset, rowY, labelWidth, OperatorChoice.Height);
                    rowY = Math.Max(OperatorChoice.Bottom, operatorLabel.Bottom) + gap;
                }
                if (!simple && operation != FormulaOperation.Round)
                {
                    int typeHeight = ValueType.PreferredHeight;
                    Place(ValueType,editorLeft, rowY, editorWidth, typeHeight);
                    Place(typeLabel,inset, rowY, labelWidth, typeHeight);
                    rowY = Math.Max(ValueType.Bottom,typeLabel.Bottom) + gap;
                }
                if (!simple)
                {
                    int inputHeight = ValueInput.GetPreferredSize(new Size(editorWidth, 0)).Height;
                    Place(ValueInput,editorLeft, rowY, editorWidth, inputHeight);
                    Place(inputLabel,inset, rowY, labelWidth, inputHeight);
                    rowY = Math.Max(ValueInput.Bottom,inputLabel.Bottom) + gap;
                }
                int hintLeft = simple ? inset : editorLeft, hintWidth = simple ? inner : editorWidth;
                Place(Hint,hintLeft, rowY, hintWidth, TextHeight(Hint, hintWidth));
                rowY = Hint.Bottom + section;
                Place(previewLabel,inset, rowY, inner, TextHeight(previewLabel, inner));
                rowY = previewLabel.Bottom + gap;
                Place(Preview,inset, rowY, inner, Logical(84));
                card.SetBounds(pad, y, width, Preview.Bottom + inset); y = card.Bottom + section;
                Validation.Visible = Validation.Text.Length > 0;
                Place(Validation,pad, y, width, Validation.Visible ? TextHeight(Validation, width) : 0);
                if (Validation.Visible) y = Validation.Bottom + section;
                Apply.Size = DialogControlSizing.Resolve(Apply,new Size(Logical(140), Logical(38))); Cancel.Size = DialogControlSizing.Resolve(Cancel,new Size(Logical(96), Logical(38)));
                Undo.Size = DialogControlSizing.Resolve(Undo, new Size(Logical(112), Logical(38)));
                Apply.Margin = new Padding(0, 0, section, 0); Cancel.Margin = Padding.Empty;
                actions.SetBounds(pad, y, width, Math.Max(Apply.Height,Cancel.Height)); actions.PerformLayout();actions.Height=Math.Max(actions.Height,Math.Max(Undo.Bottom,Math.Max(Apply.Bottom,Cancel.Bottom)));
                int contentHeight = actions.Bottom + pad;
                AutoScrollMinSize = new Size(Math.Max(width + pad * 2, intro.Right + pad), contentHeight);
            }
            finally { arranging = false; }
        }

        public Size MeasureDefaultClient()
        {Arrange(false);return new Size(DialogSizing.LogicalClient(this).Width,(int)Math.Ceiling((actions.Bottom+Logical(16))/DialogSizing.Scale(this)));}

        internal bool TryRequest(out FormulaRequest request)
        {
            request = null;
            try
            {
                request = simple ? FormulaRequest.Default(operation) : operation >= FormulaOperation.Add ? FormulaRequest.Arithmetic(operation, ValueInput.Text, ValueType.SelectedIndex == 1) : operation == FormulaOperation.Round ? FormulaRequest.Round(ValueInput.Text) :
                    ValueType.SelectedIndex == 0 ? FormulaRequest.IFErrorText(ValueInput.Text) :
                    ValueType.SelectedIndex == 1 ? FormulaRequest.IFErrorNumber(ValueInput.Text) : FormulaRequest.IFErrorFormula(ValueInput.Text);
                Validation.Text = ""; return true;
            }
            catch (FormatException ex) { Validation.Text = ex.Message; return false; }
        }
        private void UpdatePreview()
        {
            Hint.Text = simple ? (operation == FormulaOperation.Abs ? "为所选公式加一层 ABS。" : "移除包住整条公式的最外层函数，保留第一个参数。") :
                operation >= FormulaOperation.Add ? (ValueType.SelectedIndex == 0 ? "例如 100、-2 或 1.5；除数不能为零。自定义值只用于本次操作。" : "例如 A1 或 SUM(B1:B3)；引用保持原样，由 WPS 计算，不在窗口内求值。") : operation == FormulaOperation.Round ? "正数：保留小数；0：取整；负数：向小数点左侧舍入。" :
                ValueType.SelectedIndex == 0 ? "直接输入文字，不用加引号；留空返回空文字，双引号会自动转义。" :
                ValueType.SelectedIndex == 1 ? "例如 0、-1 或 1.5；使用英文小数点。" :
                "例如 A1 或 SUM(B1:B3)；函数名使用英文，语义由 WPS 检查。";
            FormulaRequest request;
            Apply.Enabled = TryRequest(out request) && canApply;
            string sample = operation == FormulaOperation.Unwrap ? "=IFERROR(A1+V2+3,0)" : "=A1+V2+3";
            previewLabel.Text = request != null && request.IsArithmetic ? "示例预览（数值 20 与公式 =A1+V2+3）" : "示例预览（原公式：" + sample + "）";
            Preview.Text = request == null ? "参数尚未通过检查" : request.IsArithmetic ? "数值：=20" + request.Operator + request.Argument + "\r\n公式：" + FormulaWrapper.Transform(sample, request) : FormulaWrapper.Transform(sample, request);
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
