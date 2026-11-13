using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Extracted from MigrationPresetManager. All three features use this
    // same layout and renderer; storage and workbook logic stay in adapters.
        internal enum PresetDetailKind { Title, Meta, Heading, Number, Label, Text, Dim, Tag, Warn, Ok }
        internal sealed class PresetDetailLine
        {
            // 0: flush left; 1: section body; 2: under a mapping number.
            internal int Indent;
            // Leave a little room above (a new section or mapping).
            internal bool Space;
            internal readonly List<KeyValuePair<PresetDetailKind, string>> Parts = new List<KeyValuePair<PresetDetailKind, string>>();
            internal PresetDetailLine Add(PresetDetailKind kind, string text) { if (!string.IsNullOrEmpty(text)) Parts.Add(new KeyValuePair<PresetDetailKind, string>(kind, text)); return this; }
            internal string Text { get { return string.Concat(Parts.Select(p => p.Value).ToArray()); } }
        }

    internal sealed class PresetManagerView
    {
        private readonly ListBox list;
        private readonly RichTextBox detail;
        private readonly TextBox search;
        private readonly Func<object,string> nameOf, infoOf;
        private readonly Font bold, title, meta, gap;
        internal readonly RoundedButton Import, Close;
        internal readonly Label Hint;
        internal PresetManagerView(Form owner, ListBox list, RichTextBox detail, TextBox search,
            RoundedButton apply, RoundedButton rename, RoundedButton remove, RoundedButton export,
            Func<object,string> nameOf, Func<object,string> infoOf)
        {
            this.list=list; this.detail=detail; this.search=search; this.nameOf=nameOf; this.infoOf=infoOf;
            list.DrawMode=DrawMode.OwnerDrawFixed; list.BorderStyle=BorderStyle.None; list.IntegralHeight=false; list.HorizontalScrollbar=false;
            detail.BorderStyle=BorderStyle.None; detail.WordWrap=true; detail.BackColor=Color.White;
            search.Width=220;
            bold=Ui.MakeFont(10f,FontStyle.Bold,owner.Font); title=Ui.MakeFont(12f,FontStyle.Bold,owner.Font);
            meta=Ui.MakeFont(9f,FontStyle.Regular,owner.Font); gap=Ui.MakeFont(5f,FontStyle.Regular,owner.Font);
            var layout=new TableLayoutPanel { Dock=DockStyle.Fill, ColumnCount=1, RowCount=3, Padding=new Padding(16), BackColor=Ui.Canvas };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent,100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // Top: import on the left, search on the right.
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, Margin = new Padding(0, 0, 0, 10), BackColor = Ui.Canvas };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            Import = Ui.Button("导入预设…", 120); var import = Import; Ui.StyleSecondary(import); import.Margin = Padding.Empty;
            Hint = new Label { Text = "预设只保存设置，不含表格数据；应用后请检查并预检。", AutoSize = true, ForeColor = Ui.InkDim, Anchor = AnchorStyles.Left, Margin = new Padding(12, 0, 0, 0) };
            var hint=Hint;
            search.Anchor = AnchorStyles.Right; search.Margin = new Padding(0, 4, 0, 0);
            top.Controls.Add(import, 0, 0); top.Controls.Add(hint, 1, 0); top.Controls.Add(search, 2, 0);
            layout.Controls.Add(top, 0, 0);
            // Middle: list | detail, each in a white framed card.
            var middle = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty, BackColor = Ui.Canvas };
            middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36)); middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 64));
            middle.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            middle.Controls.Add(Card(list, new Padding(0, 0, 6, 0), new Padding(1)), 0, 0);
            middle.Controls.Add(Card(detail, new Padding(6, 0, 0, 0), new Padding(14, 12, 10, 10)), 1, 0);
            layout.Controls.Add(middle, 0, 1);
            // Bottom: item actions on the left, apply/close on the right.
            var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, Margin = new Padding(0, 12, 0, 0), BackColor = Ui.Canvas };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var leftBar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty, BackColor = Ui.Canvas };
            var rightBar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = Padding.Empty, BackColor = Ui.Canvas };
            Close = Ui.Button("关闭", 90); var close = Close; Ui.StyleSecondary(close);
            Ui.StyleSecondary(rename); Ui.StyleSecondary(remove); Ui.StyleSecondary(export); Ui.StylePrimary(apply);
            foreach (var b in new Control[] { rename, remove, export }) { b.Margin = new Padding(0, 0, 8, 0); leftBar.Controls.Add(b); }
            apply.Margin = new Padding(8, 0, 0, 0); close.Margin = new Padding(8, 0, 0, 0);
            rightBar.Controls.Add(close); rightBar.Controls.Add(apply);
            bottom.Controls.Add(leftBar, 0, 0); bottom.Controls.Add(rightBar, 1, 0);
            layout.Controls.Add(bottom, 0, 2);
            owner.Controls.Add(layout); owner.CancelButton = close;

            list.HandleCreated += delegate { list.ItemHeight=Math.Max(Ui.Px(list,52),owner.Font.Height*2+Ui.Px(list,16)); };
            list.DrawItem += DrawPreset;
            search.HandleCreated += delegate { PresetViewNative.Cue(search,"搜索预设名称"); };
            owner.Disposed += delegate { bold.Dispose(); title.Dispose(); meta.Dispose(); gap.Dispose(); };
        }
        private static Control Card(Control content, Padding margin, Padding inner)
        {
            var card = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = inner, Margin = margin };
            card.Paint += delegate(object sender, PaintEventArgs e) { using (var pen = new Pen(Ui.GridFrame)) e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1); };
            card.Resize += delegate { card.Invalidate(); };
            content.Dock = DockStyle.Fill; card.Controls.Add(content); return card;
        }
        internal void ShowLines(List<PresetDetailLine> lines)
        {
            // Formats only stick once the box has a window; until then the
            // first preset came out plain. The handler shows it again.
            if (!detail.IsHandleCreated) { SetDetail(string.Join("\n", lines.Select(l => l.Text).ToArray())); return; }
            detail.SuspendLayout();
            try
            {
                Color warn = Color.FromArgb(180, 83, 9);
                int body = Ui.Px(detail, 16), number = 0;
                // Build the plain text first, then format each range: setting
                // the indent while appending landed on the previous paragraph.
                var text = new System.Text.StringBuilder();
                var runs = new List<Tuple<int, int, Font, Color>>();
                var paragraphs = new List<Tuple<int, int, int>>();
                foreach (var line in lines)
                {
                    if (text.Length > 0) text.Append('\n');
                    if (line.Space) { runs.Add(Tuple.Create(text.Length, 1, gap, Ui.Ink)); text.Append('\n'); }
                    int start = text.Length;
                    foreach (var part in line.Parts)
                    {
                        Font font; Color color;
                        switch (part.Key)
                        {
                            case PresetDetailKind.Title: font = title; color = Ui.Ink; break;
                            case PresetDetailKind.Meta: font = meta; color = Ui.InkDim; break;
                            case PresetDetailKind.Heading: font = bold; color = Ui.Blue; break;
                            case PresetDetailKind.Number:
                                font = bold; color = Ui.Blue;
                                number = TextRenderer.MeasureText(part.Value, bold, Size.Empty, TextFormatFlags.NoPadding).Width;
                                break;
                            case PresetDetailKind.Label: case PresetDetailKind.Dim: font = list.Font; color = Ui.InkDim; break;
                            case PresetDetailKind.Tag: font = list.Font; color = Ui.Blue; break;
                            case PresetDetailKind.Warn: font = list.Font; color = warn; break;
                            case PresetDetailKind.Ok: font = list.Font; color = Ui.Green; break;
                            default: font = list.Font; color = Ui.Ink; break;
                        }
                        runs.Add(Tuple.Create(text.Length, part.Value.Length, font, color)); text.Append(part.Value);
                    }
                    int indent = line.Indent == 0 ? 0 : line.Indent == 1 ? body : body + number;
                    paragraphs.Add(Tuple.Create(start, Math.Max(1, text.Length - start), indent));
                }
                SetDetail(text.ToString());
                foreach (var run in runs) { detail.Select(run.Item1, run.Item2); detail.SelectionFont = run.Item3; detail.SelectionColor = run.Item4; }
                foreach (var paragraph in paragraphs) { detail.Select(paragraph.Item1, paragraph.Item2); detail.SelectionIndent = paragraph.Item3; }
                // Leave the caret plain so the next preset does not inherit a format.
                detail.Select(0, 0); detail.SelectionFont = list.Font; detail.SelectionColor = Ui.Ink; detail.SelectionIndent = 0;
            }
            finally { detail.ResumeLayout(); }
        }
        // A RichTextBox keeps the format at the caret for new text, so after a
        // bold title the next preset would come out all bold. Reset first.
        internal void SetDetail(string text)
        {
            detail.Clear(); detail.SelectionFont = list.Font; detail.SelectionColor = Ui.Ink; detail.SelectionIndent = 0;
            detail.Text = text;
            detail.SelectAll(); detail.SelectionFont = list.Font; detail.SelectionColor = Ui.Ink; detail.Select(0, 0);
        }
        private void DrawPreset(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= list.Items.Count) return;
            var preset = list.Items[e.Index]; var g = e.Graphics; var r = e.Bounds;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var back = new SolidBrush(selected ? Ui.SelectBack : Color.White)) g.FillRectangle(back, r);
            if (selected) using (var bar = new SolidBrush(Ui.Blue)) g.FillRectangle(bar, r.Left, r.Top + 6, 3, r.Height - 12);
            using (var line = new Pen(Ui.GridLine)) g.DrawLine(line, r.Left + 12, r.Bottom - 1, r.Right - 12, r.Bottom - 1);
            var flags = TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
            int pad = Ui.Px(list, 14), half = r.Height / 2;
            TextRenderer.DrawText(g, nameOf(preset), bold, new Rectangle(r.Left + pad, r.Top + 4, r.Width - pad * 2, half - 2), selected ? Ui.SelectInk : Ui.Ink, flags | TextFormatFlags.Bottom);
            string info = infoOf(preset);
            TextRenderer.DrawText(g, info, list.Font, new Rectangle(r.Left + pad, r.Top + half + 1, r.Width - pad * 2, half - 4), Ui.InkDim, flags | TextFormatFlags.Top);
        }

    }
    internal static class PresetNameView
    {
        internal static void Build(Form owner, TextBox input, string prompt, string value, Func<bool> validate)
        {
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(18, 16, 18, 14), BackColor = Color.White };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = prompt, AutoSize = true, ForeColor = Ui.Ink, Margin = new Padding(0, 0, 0, 8) }, 0, 0);
            input.Text = value; input.Dock = DockStyle.Top; input.Margin = new Padding(0, 0, 0, 12); layout.Controls.Add(input, 0, 1);
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, BackColor = Color.White, Margin = Padding.Empty };
            var ok = Ui.Button("确定", 90); Ui.StylePrimary(ok); var cancel = Ui.Button("取消", 90); Ui.StyleSecondary(cancel);
            ok.Margin = new Padding(8, 0, 0, 0); cancel.Margin = new Padding(8, 0, 0, 0);
            ok.Click += delegate { if (input.Text.Trim().Length == 0) { input.Focus(); return; } if (validate()) owner.DialogResult = DialogResult.OK; };
            cancel.Click += delegate { owner.DialogResult = DialogResult.Cancel; };
            bar.Controls.Add(ok); bar.Controls.Add(cancel); layout.Controls.Add(bar, 0, 2);
            owner.Controls.Add(layout); owner.AcceptButton = ok; owner.CancelButton = cancel;
            owner.Shown += delegate { input.Focus(); input.SelectAll(); };
        }
    }
    internal static class PresetViewNative
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet=System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd,int msg,IntPtr wParam,string lParam);
        internal static void Cue(TextBox box,string text) { try { SendMessage(box.Handle,0x1501,(IntPtr)1,text); } catch(Exception) { } }
    }
    // One clickable line of the settings popup: title at TextLeft, the grey
    // note at NoteLeft, light blue rounded hover.
    internal sealed class PresetMenuRow : Control
    {
        internal string Title = "", Note = "";
        internal int TextLeft, NoteLeft;
        private bool hover;
        internal PresetMenuRow()
        {
            Cursor = Cursors.Hand; TabStop = false; BackColor = Color.White;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.Clear(BackColor); float u = g.DpiX / 96f;
            if (hover)
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (var path = Ui.RoundRect(new RectangleF(1, 1, Width - 2, Height - 2), 6 * u)) using (var brush = new SolidBrush(Ui.SelectBack)) g.FillPath(brush, path);
                g.SmoothingMode = SmoothingMode.Default;
            }
            var flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis;
            TextRenderer.DrawText(g, Title, Font, new Rectangle(TextLeft, 0, Math.Max(0, NoteLeft - TextLeft), Height), hover ? Ui.SelectInk : Ui.Ink, flags);
            TextRenderer.DrawText(g, Note, Font, new Rectangle(NoteLeft, 0, Math.Max(0, Width - NoteLeft - TextLeft), Height), Ui.InkDim, flags);
        }
    }

    internal static class PresetMenuView
    {
        internal static TableLayoutPanel Rows(Control owner,ToolStripDropDown drop,Font font,Action save,Action manage,int minimumWidth)
        {
            int pad=Ui.Px(owner,14), gap=Ui.Px(owner,28), inset=Ui.Px(owner,4);
            var entries=new[] { new[] { "保存为预设…","把当前整套设置存进预设库" },new[] { "预设管理…","查看、应用、导入、导出、删除" } };
            int titleWidth=entries.Max(e=>TextRenderer.MeasureText(e[0],font).Width),noteWidth=entries.Max(e=>TextRenderer.MeasureText(e[1],font).Width);
            int width=Math.Max(minimumWidth,pad+titleWidth+gap+noteWidth+pad);
            var panel=new TableLayoutPanel { AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=1,BackColor=Color.White,Font=font,Padding=new Padding(inset,Ui.Px(owner,6),inset,Ui.Px(owner,8)) };
            panel.Controls.Add(new Label { Text="预设",AutoSize=false,Size=new Size(width,Ui.Px(owner,28)),TextAlign=ContentAlignment.BottomLeft,ForeColor=Ui.InkDim,Padding=new Padding(pad,0,0,Ui.Px(owner,3)),Margin=Padding.Empty });
            var actions=new[] { save,manage };
            for(int i=0;i<entries.Length;i++) {
                var action=actions[i];
                var row=new PresetMenuRow { Title=entries[i][0],Note=entries[i][1],TextLeft=pad,NoteLeft=pad+titleWidth+gap,Size=new Size(width,Ui.Px(owner,34)),Font=font,Margin=Padding.Empty };
                row.Click+=delegate { drop.Close(); owner.BeginInvoke(action); }; panel.Controls.Add(row);
            }
            return panel;
        }
        internal static void Frame(Control panel)
        {
            panel.Paint+=delegate(object sender,PaintEventArgs e) { using(var pen=new Pen(Color.FromArgb(160,170,186))) e.Graphics.DrawRectangle(pen,0,0,panel.Width-1,panel.Height-1); };
            panel.Resize+=delegate { panel.Invalidate(); };
        }
    }
}
