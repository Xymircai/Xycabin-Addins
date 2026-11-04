using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using RadioButtonState = System.Windows.Forms.VisualStyles.RadioButtonState;

namespace Excel_Iterate
{
    // The "选择" header carries an Excel-style filter button: a small framed
    // square with a down arrow, which turns into a blue funnel while some
    // targets are excluded, just like an active AutoFilter.
    internal sealed class MigrationSelectionHeader : DataGridViewColumnHeaderCell
    {
        internal bool HasExcludedTargets;
        internal static Rectangle FilterButtonBounds(DataGridView grid, Rectangle cellBounds)
        {
            int size = Math.Max(Ui.Px(grid, 12), Math.Min(Ui.Px(grid, 18), cellBounds.Height - Ui.Px(grid, 8)));
            return new Rectangle(cellBounds.Right - Ui.Px(grid, 5) - size, cellBounds.Top + (cellBounds.Height - size) / 2, size, size);
        }
        protected override void Paint(Graphics graphics, Rectangle clipBounds, Rectangle cellBounds, int rowIndex,
            DataGridViewElementStates state, object value, object formattedValue, string errorText,
            DataGridViewCellStyle style, DataGridViewAdvancedBorderStyle borders, DataGridViewPaintParts parts)
        {
            base.Paint(graphics, clipBounds, cellBounds, rowIndex, state, value, formattedValue, errorText, style, borders,
                parts & ~DataGridViewPaintParts.ContentForeground);
            if ((parts & DataGridViewPaintParts.ContentForeground) == 0) return;
            var button = FilterButtonBounds(DataGridView, cellBounds);
            int gap = Ui.Px(DataGridView, 4);
            var label = new Rectangle(cellBounds.Left + gap, cellBounds.Top, Math.Max(0, button.Left - gap - cellBounds.Left - gap), cellBounds.Height);
            TextRenderer.DrawText(graphics, Convert.ToString(formattedValue), style.Font ?? DataGridView.Font, label, style.ForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            bool enabled = DataGridView.Enabled;
            Color ink = !enabled ? SystemColors.GrayText : HasExcludedTargets ? Ui.Blue : style.ForeColor;
            var smoothing = graphics.SmoothingMode;
            try
            {
                using (var fill = new SolidBrush(HasExcludedTargets && enabled ? Color.FromArgb(232, 240, 254) : Color.White))
                    graphics.FillRectangle(fill, button);
                using (var pen = new Pen(HasExcludedTargets && enabled ? Ui.Blue : Color.FromArgb(190, 199, 214)))
                    graphics.DrawRectangle(pen, button.X, button.Y, button.Width - 1, button.Height - 1);
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                float cx = button.Left + (button.Width - 1) / 2f, cy = button.Top + (button.Height - 1) / 2f, u = button.Width / 16f;
                using (var brush = new SolidBrush(ink))
                {
                    if (HasExcludedTargets)
                        // Funnel: a wide top narrowing into a short stem.
                        graphics.FillPolygon(brush, new[] {
                            new PointF(cx - 5f * u, cy - 4.5f * u), new PointF(cx + 5f * u, cy - 4.5f * u), new PointF(cx + 1.2f * u, cy + 0.2f * u),
                            new PointF(cx + 1.2f * u, cy + 4.8f * u), new PointF(cx - 1.2f * u, cy + 3.4f * u), new PointF(cx - 1.2f * u, cy + 0.2f * u) });
                    else
                        graphics.FillPolygon(brush, new[] { new PointF(cx - 4f * u, cy - 2f * u), new PointF(cx + 4f * u, cy - 2f * u), new PointF(cx, cy + 2.5f * u) });
                }
            }
            finally { graphics.SmoothingMode = smoothing; }
        }
    }

    // Shared chrome for the grid's Excel-style dropdowns (workbooks and
    // "工作表"): framed search box, a framed list area the subclass fills,
    // count on the left, 确定/取消 on the right, and a corner resize grip.
    // It is a lightweight ToolStripDropDown, not a second modeless Form.
    internal abstract class MigrationFilterPopup : ToolStripDropDown
    {
        // 1px framed white panel; the frame turns blue while a child has focus.
        protected sealed class FramePanel : Panel
        {
            private Color border;
            internal Color Border { get { return border; } set { if (border != value) { border = value; Invalidate(); } } }
            internal FramePanel(Color color)
            {
                border = color; BackColor = Color.White; Padding = new Padding(1); Margin = Padding.Empty;
                SetStyle(ControlStyles.ResizeRedraw | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                using (var pen = new Pen(border)) e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
            }
        }

        // Magnifier while the search is empty; a clickable × once text is typed.
        private sealed class SearchGlyph : Control
        {
            private bool clear;
            internal bool ClearMode { get { return clear; } set { if (clear != value) { clear = value; Cursor = value ? Cursors.Hand : Cursors.Default; Invalidate(); } } }
            internal SearchGlyph()
            {
                BackColor = Color.White; TabStop = false;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                var g = e.Graphics; g.Clear(BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
                float s = Math.Min(Width, Height), u = s / 16f, cx = Width / 2f, cy = Height / 2f;
                using (var pen = new Pen(clear ? Ui.Ink : Ui.InkDim, Math.Max(1f, 1.4f * u)))
                {
                    if (clear)
                    {
                        g.DrawLine(pen, cx - 3.5f * u, cy - 3.5f * u, cx + 3.5f * u, cy + 3.5f * u);
                        g.DrawLine(pen, cx + 3.5f * u, cy - 3.5f * u, cx - 3.5f * u, cy + 3.5f * u);
                    }
                    else
                    {
                        float r = 4f * u, ox = cx - 1.2f * u, oy = cy - 1.2f * u;
                        g.DrawEllipse(pen, ox - r, oy - r, r * 2, r * 2);
                        g.DrawLine(pen, ox + r * 0.72f, oy + r * 0.72f, cx + 5.5f * u, cy + 5.5f * u);
                    }
                }
            }
        }

        // Bottom-right drag handle, like the one on Excel's filter menu.
        private sealed class SizeGrip : Control
        {
            private readonly MigrationFilterPopup owner;
            private Point start; private Size origin; private bool dragging;
            internal SizeGrip(MigrationFilterPopup popup)
            {
                owner = popup; Cursor = Cursors.SizeNWSE; BackColor = Color.White; TabStop = false;
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            }
            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.Clear(BackColor);
                int cell = Math.Max(2, Width / 3);
                using (var brush = new SolidBrush(Color.FromArgb(165, 174, 190)))
                    for (int r = 0; r < 3; r++)
                        for (int c = 0; r + c < 3; c++)
                            e.Graphics.FillRectangle(brush, Width - (c + 1) * cell, Height - (r + 1) * cell, cell - 1, cell - 1);
            }
            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left) { dragging = true; start = Cursor.Position; origin = owner.Size; Capture = true; }
                base.OnMouseDown(e);
            }
            protected override void OnMouseMove(MouseEventArgs e)
            {
                if (dragging && Capture)
                {
                    var p = Cursor.Position;
                    owner.ResizeTo(new Size(origin.Width + p.X - start.X, origin.Height + p.Y - start.Y));
                }
                base.OnMouseMove(e);
            }
            protected override void OnMouseUp(MouseEventArgs e) { dragging = false; Capture = false; base.OnMouseUp(e); }
            protected override void OnMouseCaptureChanged(EventArgs e) { dragging = false; base.OnMouseCaptureChanged(e); }
        }

        private static class Native
        {
            internal const int EM_SETCUEBANNER = 0x1501;
            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        }

        private static readonly Color OuterFrame = Color.FromArgb(160, 170, 186);
        protected static readonly Color InnerFrame = Color.FromArgb(205, 212, 224);
        // Last size the user dragged each popup to, in 96-DPI units, per session.
        private static readonly Dictionary<string, Size> rememberedSizes96 = new Dictionary<string, Size>();

        private readonly TextBox search = new TextBox { BorderStyle = BorderStyle.None };
        private readonly SearchGlyph glyph = new SearchGlyph();
        protected readonly Label CountLabel = new Label { AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Ui.InkDim };
        protected readonly Label EmptyNote = new Label { AutoSize = false, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Ui.InkDim, BackColor = Color.White, Visible = false };
        protected readonly ToolTip Tips = new ToolTip();
        // The framed list area: Strip on top, a hairline, then ListHost (Fill).
        protected readonly Panel Strip, ListHost;
        private readonly FramePanel root;
        private readonly ToolStripControlHost host;
        private readonly Size minimum;
        private readonly float scale;
        private readonly string sizeKey;
        internal TextBox Search { get { return search; } }

        protected MigrationFilterPopup(DataGridView grid, string prefix, string cue, string emptyText, int columnWidth, Size defaultSize96)
        {
            Name = prefix + "Popup"; sizeKey = prefix;
            AutoSize = false; Padding = Padding.Empty; Margin = Padding.Empty; BackColor = Color.White; DropShadowEnabled = true;
            Font = grid.Font;
            using (var g = grid.CreateGraphics()) scale = g.DpiX / 96f;
            search.Name = prefix + "Search"; glyph.Name = prefix + "Glyph"; CountLabel.Name = prefix + "Count"; EmptyNote.Name = prefix + "Empty"; EmptyNote.Text = emptyText;
            foreach (Control control in new Control[] { search, CountLabel, EmptyNote }) control.Font = grid.Font;

            minimum = new Size(P(240), P(240));
            Size remembered;
            Size size = rememberedSizes96.TryGetValue(sizeKey, out remembered) ? new Size(P(remembered.Width), P(remembered.Height)) : new Size(Math.Max(P(defaultSize96.Width), columnWidth), P(defaultSize96.Height));
            var area = Screen.FromControl(grid).WorkingArea;
            size = new Size(Math.Max(minimum.Width, Math.Min(size.Width, area.Width - P(24))), Math.Max(minimum.Height, Math.Min(size.Height, area.Height - P(24))));

            root = new FramePanel(OuterFrame) { Name = prefix + "Frame", Size = size, Font = grid.Font };
            var layout = new TableLayoutPanel { Name = prefix + "Panel", Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, BackColor = Color.White, Padding = new Padding(P(8)), Margin = Padding.Empty };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            int searchHeight = Math.Max(P(28), search.PreferredHeight + P(10));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, searchHeight + P(6)));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, P(30) + P(8)));

            // Search box: borderless TextBox + glyph inside a framed strip.
            var searchFrame = new FramePanel(InnerFrame) { Name = prefix + "SearchFrame", Dock = DockStyle.Top, Height = searchHeight, Margin = new Padding(0, 0, 0, P(6)) };
            search.BackColor = Color.White;
            searchFrame.Controls.Add(search); searchFrame.Controls.Add(glyph);
            searchFrame.Layout += delegate
            {
                int inner = searchFrame.Height - 2, glyphWidth = P(24);
                glyph.SetBounds(searchFrame.Width - 1 - glyphWidth, 1, glyphWidth, inner);
                search.SetBounds(1 + P(7), 1 + (inner - search.Height) / 2, Math.Max(P(20), searchFrame.Width - 2 - P(7) - glyphWidth), search.Height);
            };
            layout.Controls.Add(searchFrame, 0, 0);

            // List area: subclass content goes into Strip and ListHost.
            var listFrame = new FramePanel(InnerFrame) { Name = prefix + "ListFrame", Dock = DockStyle.Fill, Margin = Padding.Empty };
            Strip = new Panel { Name = prefix + "Strip", Dock = DockStyle.Top, Height = Math.Max(P(28), Font.Height + P(10)), BackColor = Color.White };
            var rule = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Ui.GridLine };
            ListHost = new Panel { Name = prefix + "ListHost", Dock = DockStyle.Fill, BackColor = Color.White, Padding = new Padding(P(4), P(3), 0, 0) };
            ListHost.Controls.Add(EmptyNote);
            ListHost.Layout += delegate { EmptyNote.SetBounds(ListHost.Padding.Left, ListHost.Padding.Top, Math.Max(0, ListHost.ClientSize.Width - ListHost.Padding.Horizontal), Math.Max(0, Math.Min(P(48), ListHost.ClientSize.Height - ListHost.Padding.Vertical))); };
            // Docking runs from the highest index down: Strip, rule, then Fill.
            listFrame.Controls.Add(ListHost); listFrame.Controls.Add(rule); listFrame.Controls.Add(Strip);
            ListHost.BringToFront();
            layout.Controls.Add(listFrame, 0, 1);

            // Count on the left, 确定 (default) and 取消 on the right.
            var actions = new TableLayoutPanel { Name = prefix + "Actions", Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1, Margin = new Padding(0, P(8), 0, 0), Padding = Padding.Empty, BackColor = Color.White };
            actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var confirm = SmallButton("确定", true); confirm.Name = prefix + "Apply";
            var cancel = SmallButton("取消", false); cancel.Name = prefix + "Cancel";
            CountLabel.Dock = DockStyle.Fill; CountLabel.Margin = Padding.Empty;
            actions.Controls.Add(CountLabel, 0, 0); actions.Controls.Add(confirm, 1, 0); actions.Controls.Add(cancel, 2, 0);
            layout.Controls.Add(actions, 0, 2);

            root.Controls.Add(layout);
            var grip = new SizeGrip(this) { Name = prefix + "Grip", Size = new Size(P(10), P(10)) };
            root.Controls.Add(grip); grip.BringToFront();
            root.Layout += delegate { grip.Location = new Point(root.Width - 1 - grip.Width, root.Height - 1 - grip.Height); };

            host = new ToolStripControlHost(root) { AutoSize = false, Size = root.Size, Margin = Padding.Empty, Padding = Padding.Empty };
            Items.Add(host); Size = root.Size;

            search.TextChanged += delegate { glyph.ClearMode = search.TextLength > 0; Refilter(); };
            search.GotFocus += delegate { searchFrame.Border = Ui.Blue; };
            search.LostFocus += delegate { searchFrame.Border = InnerFrame; };
            search.HandleCreated += delegate { Native.SendMessage(search.Handle, Native.EM_SETCUEBANNER, (IntPtr)1, cue); };
            glyph.Click += delegate { if (search.TextLength > 0) { search.Clear(); } search.Focus(); };
            searchFrame.Click += delegate { search.Focus(); };
            confirm.Click += delegate { Commit(); }; cancel.Click += delegate { Close(); };
            Opened += delegate { BeginInvoke(new Action(delegate { if (!IsDisposed && Visible) search.Focus(); })); };
            search.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Commit(); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; if (search.TextLength > 0) search.Clear(); else Close(); }
                else if (e.KeyCode == Keys.Down) { e.SuppressKeyPress = true; FocusList(); }
            };
        }
        protected abstract void Refilter();
        protected abstract void Commit();
        protected abstract void FocusList();
        protected int P(int value) { return (int)Math.Round(value * scale); }
        protected string Term { get { return search.Text.Trim(); } }
        // Places a list control in ListHost under the "no match" label.
        protected void HostList(Control list)
        {
            list.Dock = DockStyle.Fill; list.Font = Font; list.BackColor = Color.White; list.ForeColor = Ui.Ink;
            ListHost.Controls.Add(list); EmptyNote.BringToFront();
        }
        private RoundedButton SmallButton(string text, bool primary)
        {
            var button = Ui.Button(text, 72);
            if (primary) Ui.StylePrimary(button);
            button.Font = Font; button.BackColor = Color.White; button.Radius = 6;
            button.MinimumSize = Size.Empty; button.Size = new Size(P(72), P(30));
            button.Margin = new Padding(P(8), 0, 0, 0); button.Anchor = AnchorStyles.Right;
            return button;
        }
        internal void ResizeTo(Size size)
        {
            var area = Screen.FromPoint(Bounds.Location).WorkingArea;
            size = new Size(Math.Max(minimum.Width, Math.Min(size.Width, area.Right - Left)), Math.Max(minimum.Height, Math.Min(size.Height, area.Bottom - Top)));
            if (size == Size) return;
            SuspendLayout();
            try { root.Size = size; host.Size = size; Size = size; }
            finally { ResumeLayout(true); }
            rememberedSizes96[sizeKey] = new Size((int)Math.Round(size.Width / scale), (int)Math.Round(size.Height / scale));
        }
        protected override void Dispose(bool disposing)
        { if (disposing) Tips.Dispose(); base.Dispose(disposing); }
    }

    // Workbook picker for one side of a row: search, (全选)/反选, check the
    // workbooks this row applies to, or add more. Edits stay a draft until 确定.
    internal sealed class MigrationBookPopup : MigrationFilterPopup
    {
        private sealed class Choice
        {
            internal string Path, Sheets;
            internal bool Selected;
            public override string ToString() { return System.IO.Path.GetFileName(Path); }
        }
        private readonly List<Choice> choices = new List<Choice>();
        private readonly List<Choice> matching = new List<Choice>();
        private readonly CheckedListBox list = new CheckedListBox { Name = "migrationBookList", CheckOnClick = true, IntegralHeight = false, BorderStyle = BorderStyle.None, HorizontalScrollbar = true };
        private readonly CheckBox all = new CheckBox { Name = "migrationBookAll", Text = "(全选)", AutoSize = true, AutoCheck = false, BackColor = Color.White };
        private readonly LinkLabel invert = new LinkLabel { Name = "migrationBookInvert", Text = "反选", AutoSize = true, LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Color.White };
        private readonly LinkLabel add = new LinkLabel { Name = "migrationBookAdd", Text = "添加工作簿…", AutoSize = true, LinkBehavior = LinkBehavior.HoverUnderline, BackColor = Color.White };
        private readonly Action<IList<string>> apply;
        private bool populating;
        internal CheckedListBox List { get { return list; } }
        internal int MatchCount { get { return matching.Count; } }

        internal MigrationBookPopup(DataGridView grid, IEnumerable<MigrationBook> books, IEnumerable<string> selected, string title, Action<IList<string>> onApply, Action onAdd)
            : base(grid, "migrationBook", "搜索文件名、路径或工作表", "没有匹配的工作簿", 0, new Size(340, 360))
        {
            apply = onApply;
            var chosen = new HashSet<string>(selected ?? new string[0], StringComparer.OrdinalIgnoreCase);
            foreach (var book in books) choices.Add(new Choice { Path = book.Path, Sheets = string.Join("、", book.Sheets.ToArray()), Selected = chosen.Contains(book.Path) });
            all.Font = invert.Font = add.Font = Font;
            Strip.Controls.Add(all); Strip.Controls.Add(add); Strip.Controls.Add(invert);
            Strip.Layout += delegate
            {
                var allSize = all.PreferredSize; var invertSize = invert.PreferredSize; var addSize = add.PreferredSize;
                invert.SetBounds(Strip.Width - P(10) - invertSize.Width, (Strip.Height - invertSize.Height) / 2, invertSize.Width, invertSize.Height);
                add.SetBounds(invert.Left - P(14) - addSize.Width, (Strip.Height - addSize.Height) / 2, addSize.Width, addSize.Height);
                all.SetBounds(P(5), (Strip.Height - allSize.Height) / 2, Math.Min(allSize.Width, Math.Max(0, add.Left - P(12) - P(5))), allSize.Height);
            };
            HostList(list);
            all.Click += delegate { SelectMatching(all.CheckState != CheckState.Checked); };
            invert.LinkClicked += delegate { InvertMatching(); };
            add.LinkClicked += delegate { Close(); if (onAdd != null) onAdd(); };
            list.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
            { if (!populating) { matching[e.Index].Selected = e.NewValue == CheckState.Checked; UpdateCount(); } };
            list.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                int index = list.IndexFromPoint(e.Location);
                string text = index < 0 || index >= matching.Count ? "" : matching[index].Path + (matching[index].Sheets.Length == 0 ? "" : Environment.NewLine + "工作表：" + matching[index].Sheets);
                if (Tips.GetToolTip(list) != text) Tips.SetToolTip(list, text);
            };
            list.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; ApplySelection(); } };
            Populate();
        }
        protected override void Refilter() { Populate(); }
        protected override void Commit() { ApplySelection(); }
        protected override void FocusList() { if (list.Items.Count > 0) { list.Focus(); if (list.SelectedIndex < 0) list.SelectedIndex = 0; } }
        private void Populate()
        {
            string term = Term; populating = true; list.BeginUpdate();
            try
            {
                matching.Clear(); list.Items.Clear();
                foreach (var choice in choices)
                    if (term.Length == 0 || choice.Path.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0 || choice.Sheets.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                    { matching.Add(choice); list.Items.Add(choice, choice.Selected); }
            }
            finally { list.EndUpdate(); populating = false; }
            all.Text = term.Length == 0 ? "(全选)" : "(选择所有搜索结果)";
            EmptyNote.Visible = matching.Count == 0;
            UpdateCount();
        }
        internal void SelectMatching(bool selected)
        { foreach (var choice in matching) choice.Selected = selected; Populate(); }
        internal void InvertMatching()
        { foreach (var choice in matching) choice.Selected = !choice.Selected; Populate(); }
        private void UpdateCount()
        {
            if (IsDisposed) return;
            int selected = matching.Count(c => c.Selected);
            all.Enabled = invert.Enabled = matching.Count > 0;
            all.CheckState = matching.Count == 0 || selected == 0 ? CheckState.Unchecked : selected == matching.Count ? CheckState.Checked : CheckState.Indeterminate;
            string text = "已选 " + choices.Count(c => c.Selected) + " / " + choices.Count;
            if (matching.Count != choices.Count) text += "，显示 " + matching.Count;
            CountLabel.Text = text;
        }
        internal void ApplySelection()
        {
            apply(choices.Where(c => c.Selected).Select(c => c.Path).ToList()); Close();
        }
    }

    // "工作表" dropdown: pick one worksheet that every target in scope has and
    // apply it to all of them. Names only some targets have are listed, grayed
    // with their coverage, but cannot be chosen.
    internal sealed class MigrationSheetPopup : MigrationFilterPopup
    {
        private sealed class SheetChoice
        {
            internal string Name; internal int Count; internal bool Common, Header;
            public override string ToString() { return Name; }
        }
        private readonly List<SheetChoice> choices = new List<SheetChoice>();
        private readonly List<SheetChoice> matching = new List<SheetChoice>();
        private readonly ListBox list = new ListBox { Name = "migrationSheetList", IntegralHeight = false, BorderStyle = BorderStyle.None, DrawMode = DrawMode.OwnerDrawFixed, SelectionMode = SelectionMode.None };
        private readonly Label scope = new Label { Name = "migrationSheetScope", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Ui.InkDim, BackColor = Color.White };
        private readonly Action<string> apply;
        private readonly string noun;
        private readonly int total;
        private string chosen;
        private int hover = -1;
        private bool allowPartial;
        // When set, a sheet only some workbooks have can be chosen too; the
        // caller then splits the workbooks lacking it into their own row.
        internal bool AllowPartial
        {
            get { return allowPartial; }
            set { if (allowPartial == value) return; allowPartial = value; Populate(); }
        }
        private bool Selectable(SheetChoice choice) { return !choice.Header && (choice.Common || allowPartial); }
        internal ListBox List { get { return list; } }
        internal string Chosen { get { return chosen; } }
        internal int MatchCount { get { return matching.Count(c => c.Common); } }

        internal MigrationSheetPopup(DataGridView grid, IList<string> common, IList<KeyValuePair<string, int>> partial, int scopeCount, bool scopeIsChecked, string current, Action<string> onApply, string noun = "目标")
            : base(grid, "migrationSheet", "搜索工作表名称", "没有匹配的工作表",
                grid.Columns.Contains("sheet") ? grid.Columns["sheet"].Width : 0, new Size(300, 340))
        {
            apply = onApply; total = scopeCount; this.noun = noun;
            foreach (string name in common) choices.Add(new SheetChoice { Name = name, Count = scopeCount, Common = true });
            foreach (var pair in partial) choices.Add(new SheetChoice { Name = pair.Key, Count = pair.Value });
            chosen = common.FirstOrDefault(n => string.Equals(n, current, StringComparison.OrdinalIgnoreCase));

            scope.Font = Font;
            scope.Text = scopeCount == 0 ? "没有可读取工作表的目标" : "应用到 " + scopeCount + " 个" + (scopeIsChecked ? "已勾选" : "") + noun;
            Strip.Controls.Add(scope);
            Strip.Layout += delegate { scope.SetBounds(P(6), 0, Math.Max(0, Strip.Width - P(12)), Strip.Height); };

            HostList(list);
            list.ItemHeight = Math.Max(P(22), Font.Height + P(6));
            list.DrawItem += DrawChoice;
            list.MouseMove += delegate(object sender, MouseEventArgs e)
            {
                int index = list.IndexFromPoint(e.Location);
                if (index >= matching.Count) index = -1;
                if (index != hover) { hover = index; list.Invalidate(); }
                string tip = index < 0 || matching[index].Common || matching[index].Header ? "" : allowPartial
                    ? "只有 " + matching[index].Count + "/" + total + " 个" + noun + "有这张工作表；应用后其余的拆成单独一行，保持原工作表"
                    : "只有 " + matching[index].Count + "/" + total + " 个目标有这张工作表，不能统一应用";
                if (Tips.GetToolTip(list) != tip) Tips.SetToolTip(list, tip);
            };
            list.MouseLeave += delegate { if (hover != -1) { hover = -1; list.Invalidate(); } };
            list.MouseClick += delegate(object sender, MouseEventArgs e) { Choose(list.IndexFromPoint(e.Location)); };
            list.MouseDoubleClick += delegate(object sender, MouseEventArgs e)
            { int index = list.IndexFromPoint(e.Location); if (index >= 0 && index < matching.Count && Selectable(matching[index])) { Choose(index); Commit(); } };
            list.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Commit(); }
                else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down) { e.SuppressKeyPress = true; Step(e.KeyCode == Keys.Down ? 1 : -1); }
            };
            Populate();
        }
        protected override void Refilter() { Populate(); }
        protected override void Commit() { if (chosen != null) apply(chosen); Close(); }
        protected override void FocusList() { list.Focus(); if (chosen == null) Step(1); }
        private void Populate()
        {
            string term = Term; list.BeginUpdate();
            try
            {
                matching.Clear(); list.Items.Clear(); bool header = false;
                foreach (var choice in choices)
                {
                    if (term.Length != 0 && choice.Name.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    if (!choice.Common && !header)
                    {
                        // A non-selectable caption separates the partial names.
                        var caption = new SheetChoice { Name = allowPartial ? "仅部分" + noun + "有（应用后拆分为两行）" : "仅部分目标存在（不能统一应用）", Header = true };
                        matching.Add(caption); list.Items.Add(caption); header = true;
                    }
                    matching.Add(choice); list.Items.Add(choice);
                }
            }
            finally { list.EndUpdate(); }
            hover = -1;
            EmptyNote.Visible = matching.Count == 0;
            UpdateCount();
        }
        private void Choose(int index)
        {
            if (index < 0 || index >= matching.Count || !Selectable(matching[index])) return;
            chosen = matching[index].Name; list.Invalidate(); UpdateCount();
        }
        private void Step(int direction)
        {
            var commons = matching.Select((c, i) => new { c, i }).Where(x => Selectable(x.c)).Select(x => x.i).ToList();
            if (commons.Count == 0) return;
            int current = chosen == null ? -1 : matching.FindIndex(c => Selectable(c) && c.Name == chosen);
            int position = commons.IndexOf(current);
            position = position < 0 ? (direction > 0 ? 0 : commons.Count - 1) : Math.Max(0, Math.Min(commons.Count - 1, position + direction));
            Choose(commons[position]);
            if (commons[position] < list.TopIndex) list.TopIndex = commons[position];
            else if (list.ClientSize.Height > 0 && commons[position] >= list.TopIndex + list.ClientSize.Height / list.ItemHeight) list.TopIndex = commons[position] - list.ClientSize.Height / list.ItemHeight + 1;
        }
        private void UpdateCount()
        {
            if (IsDisposed) return;
            int commons = choices.Count(c => c.Common);
            CountLabel.Text = chosen != null ? "已选“" + chosen + "”" : commons == 0 ? "没有共同工作表" : "共同工作表 " + commons + " 个";
        }
        private void DrawChoice(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= matching.Count) return;
            var choice = matching[e.Index]; var g = e.Graphics; var bounds = e.Bounds;
            bool selectable = Selectable(choice), isChosen = selectable && choice.Name == chosen, hot = selectable && e.Index == hover;
            using (var back = new SolidBrush(isChosen ? Ui.SelectBack : hot ? Color.FromArgb(244, 247, 253) : Color.White)) g.FillRectangle(back, bounds);
            var flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis;
            if (choice.Header)
            {
                var caption = new Rectangle(bounds.Left + P(2), bounds.Top, bounds.Width - P(4), bounds.Height);
                TextRenderer.DrawText(g, choice.Name, Font, caption, Ui.InkDim, flags);
                return;
            }
            var glyph = RadioButtonRenderer.GetGlyphSize(g, RadioButtonState.UncheckedNormal);
            var at = new Point(bounds.Left + P(2), bounds.Top + (bounds.Height - glyph.Height) / 2);
            RadioButtonRenderer.DrawRadioButton(g, at, !selectable ? RadioButtonState.UncheckedDisabled : isChosen ? RadioButtonState.CheckedNormal : RadioButtonState.UncheckedNormal);
            int left = at.X + glyph.Width + P(6), right = bounds.Right - P(6);
            if (!choice.Common)
            {
                string coverage = choice.Count + "/" + total;
                int width = TextRenderer.MeasureText(g, coverage, Font, Size.Empty, flags).Width;
                TextRenderer.DrawText(g, coverage, Font, new Rectangle(right - width, bounds.Top, width, bounds.Height), Ui.InkDim, flags);
                right -= width + P(8);
            }
            TextRenderer.DrawText(g, choice.Name, Font, new Rectangle(left, bounds.Top, Math.Max(0, right - left), bounds.Height), selectable ? (isChosen ? Ui.SelectInk : Ui.Ink) : SystemColors.GrayText, flags);
        }
    }
}
