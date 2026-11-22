using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class RoundedEditor : Panel
    {
        private readonly Control editor;
        private readonly bool multiline;

        internal RoundedEditor(Control editor, int height = 36)
        {
            this.editor = editor;
            var text = editor as TextBox;
            var number = editor as NumericUpDown;
            var choice = editor as ComboBox;
            if (text != null) { text.BorderStyle = BorderStyle.None; text.AutoSize = false; multiline = text.Multiline; }
            if (number != null) number.BorderStyle = BorderStyle.None;
            if (choice != null) choice.FlatStyle = FlatStyle.Flat;
            editor.AutoSize = false;
            editor.Dock = DockStyle.None;
            editor.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            editor.BackColor = SystemColors.Window;
            editor.ForeColor = Ui.Ink;
            editor.TabIndex = 0;
            Margin = editor.Margin;
            Padding = new Padding(11, 5, 11, 5);
            Size = new Size(editor.Width + Padding.Horizontal, height);
            TabStop = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            Controls.Add(editor);
            editor.Enter += delegate { Invalidate(); };
            editor.Leave += delegate { Invalidate(); };
            editor.FontChanged += delegate { PerformLayout(); };
            editor.EnabledChanged += delegate { Invalidate(); };
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            if (editor == null || editor.Parent != this) return;
            var text = editor as TextBox;
            int naturalHeight = multiline ? Height - Padding.Vertical : text != null ? Math.Max(18, text.PreferredHeight) : editor.PreferredSize.Height;
            Size size = DialogControlSizing.Resolve(editor, new Size(Math.Max(1, Width - Padding.Horizontal), Math.Max(1, naturalHeight)));
            var bounds = new Rectangle(Padding.Left, multiline ? Padding.Top : Math.Max(2, (Height - size.Height) / 2), size.Width, size.Height);
            if (editor.Bounds != bounds) editor.Bounds = bounds;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (editor.CanFocus) editor.Focus();
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent == null ? Ui.Canvas : Parent.BackColor);
            if (Width < 4 || Height < 4) return;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float scale = g.DpiX / 96f;
            Color outline = SystemInformation.HighContrast ? SystemColors.WindowText : ContainsFocus ? Ui.Blue : Color.FromArgb(197, 207, 223);
            using (var path = Ui.RoundRect(new RectangleF(1, 1, Width - 2, Height - 2), 9 * scale))
            using (var brush = new SolidBrush(editor.BackColor))
            using (var pen = new Pen(outline, ContainsFocus ? Math.Max(1, 1.6f * scale) : 1))
            {
                g.FillPath(brush, path);
                g.DrawPath(pen, path);
            }
        }
    }

    internal static class ConversionUi
    {
        internal static Label Label(string text, float size = 10, bool bold = false)
        {
            var label = Ui.Label(text, size, bold);
            label.AutoSize = true;
            label.Dock = DockStyle.Top;
            label.AutoEllipsis = false;
            label.Margin = new Padding(0, 3, 0, 3);
            return label;
        }

        internal static Label Hint(string text)
        {
            var label = Label(text, 9);
            label.ForeColor = SystemInformation.HighContrast ? SystemColors.WindowText : Ui.InkDim;
            return label;
        }

        internal static RoundedButton Button(string text, int width, bool primary = false)
        {
            var button = Ui.Button(text, width);
            button.Radius = 10;
            if (primary) Ui.StylePrimary(button);
            return button;
        }

        internal static void DrawTab(TabControl tabs, DrawItemEventArgs e)
        {
            Rectangle bounds = e.Bounds;
            bounds.Inflate(-2, -2);
            if (bounds.Width < 5 || bounds.Height < 5) return;
            bool selected = e.Index == tabs.SelectedIndex;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color back = selected ? Ui.SelectBack : Color.White;
            Color ink = selected ? Ui.SelectInk : Ui.InkDim;
            if (SystemInformation.HighContrast)
            { back = selected ? SystemColors.Highlight : SystemColors.Window; ink = selected ? SystemColors.HighlightText : SystemColors.WindowText; }
            using (var path = Ui.RoundRect(bounds, 7 * e.Graphics.DpiX / 96f))
            using (var brush = new SolidBrush(back)) e.Graphics.FillPath(brush, path);
            TextRenderer.DrawText(e.Graphics, tabs.TabPages[e.Index].Text, tabs.Font, bounds, ink,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            if (selected && tabs.Focused) ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(bounds, -3, -3), ink, back);
        }
    }
}
