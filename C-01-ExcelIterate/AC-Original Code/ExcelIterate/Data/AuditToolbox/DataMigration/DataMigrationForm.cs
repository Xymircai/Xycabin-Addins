using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace Excel_Iterate
{
    // A workbook added to the source or the target list, with its sheet names.
    internal sealed class MigrationBook
    {
        internal string Path = "", Error = "";
        internal readonly List<string> Sheets = new List<string>();
        internal string Name { get { return System.IO.Path.GetFileName(Path); } }
    }
    // A source range or a target range: which workbooks, sheet and area.
    // Used = it exists in the 源区域/目标区域 list (an unused side is a gap).
    internal sealed class MigrationSide
    {
        internal readonly List<string> Books = new List<string>();
        internal string Sheet = "", Range = "";
        internal bool Used;
        // Source side only: fill the trigger cell with each trigger value,
        // recalculate, read this range and send it to the workbook named on
        // the same row (逐行重算).
        internal MigrationTrigger Trigger;
        internal MigrationSide Copy() { var copy = new MigrationSide { Sheet = Sheet, Range = Range, Used = Used, Trigger = Trigger == null ? null : Trigger.Copy() }; copy.Books.AddRange(Books); return copy; }
    }
    // One trigger value row and the target workbook its name resolved to.
    internal sealed class MigrationTriggerRow
    {
        internal int Row;
        internal string Value = "", Name = "", Path = "";
        // 精确 / 模糊 / 多个候选 / 缺失 / 已确认 / 跳过 / 空值
        internal string Kind = "";
        internal readonly List<KeyValuePair<string, double>> Candidates = new List<KeyValuePair<string, double>>();
        internal bool Ready { get { return Path.Length > 0 && (Kind == "精确" || Kind == "已确认"); } }
    }
    internal sealed class MigrationTrigger
    {
        internal string Cell = "B2", Values = "", NameColumn = "H";
        // Sheet row -> target path the user chose ("" = skip this row).
        internal readonly Dictionary<int, string> Chosen = new Dictionary<int, string>();
        internal readonly List<MigrationTriggerRow> Rows = new List<MigrationTriggerRow>();
        internal string Error = "", Signature = "";
        internal MigrationTrigger Copy()
        {
            var copy = new MigrationTrigger { Cell = Cell, Values = Values, NameColumn = NameColumn };
            foreach (var chosen in Chosen) copy.Chosen[chosen.Key] = chosen.Value;
            return copy;
        }
    }
    // One row of 映射: the source range and target range on the same row are
    // paired; dragging in 映射 swaps one side between two rows.
    internal sealed class MigrationPair
    {
        internal MigrationSide Source = new MigrationSide(), Target = new MigrationSide();
        internal bool Enabled = true, ValuesOnly;
        // What the row expanded to in the last preview/run (one per workbook pair).
        internal readonly List<MigrationMapping> Instances = new List<MigrationMapping>();
    }
    // A source range of one workbook, as handed to the single-source engine.
    internal sealed class MigrationSourceItem
    {
        internal readonly string Id = Guid.NewGuid().ToString("N");
        internal string Path = "", Sheet = "", Range = "", Name = "";
        internal bool ValuesOnly;
        // 逐行重算: copy TriggerFrom into TriggerCell and recalculate first;
        // TriggerNameCell must still read TriggerExpect.
        internal string TriggerCell = "", TriggerFrom = "", TriggerNameCell = "", TriggerExpect = "";
        internal int TriggerRow;
        internal string Label { get { return System.IO.Path.GetFileName(Path) + (Name.Length == 0 ? "" : " · " + Name); } }
    }
    // One concrete write: a source range of one workbook into one target workbook.
    internal sealed class MigrationMapping
    {
        internal MigrationPair Pair;
        internal MigrationSourceItem Source;
        internal string TargetPath = "", TargetSheet = "", Start = "", TargetRange = "";
        internal MigrationArea Area;
        internal MigrationTarget Result;
    }
    // A pass is one source range migrated to targets that each appear once,
    // i.e. exactly one run of the existing single-source DataMigrationHost.
    internal sealed class MigrationPass
    {
        internal MigrationSourceItem Source;
        internal MigrationRequest Request;
        internal readonly List<MigrationMapping> Mappings = new List<MigrationMapping>();
    }
    internal sealed class MigrationPlan
    {
        internal readonly List<MigrationPass> Passes = new List<MigrationPass>();
        internal readonly List<MigrationMapping> Mappings = new List<MigrationMapping>();
        internal bool ReplaceTargets, AllowEmptySource, OutputCopies;
        internal readonly Dictionary<string, string> Outputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal int Blocked;
        // 逐行重算 rows left out (empty value, skipped, no target), and rows
        // whose writes land on the same target area (later one wins).
        internal readonly List<string> Skipped = new List<string>(), Overrides = new List<string>();
        // Added target workbooks no 逐行重算 row matched: not written, only reported.
        internal readonly List<string> Unused = new List<string>();
    }
    // Narrow vertical tab on the window's right edge: "◂ 区域选择" opens the
    // range panel, "▸" closes it.
    // Settings button: a gear, highlighted on hover.
    internal sealed class MigrationGearButton : Control
    {
        private bool hover;
        internal MigrationGearButton()
        {
            Cursor = Cursors.Hand; TabStop = false; Height = 28;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); Height = Ui.Px(this, 28); }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.Clear(Parent == null ? Ui.Canvas : Parent.BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
            float u = g.DpiX / 96f, size = Math.Min(Width, Height);
            float cx = Width / 2f, cy = Height / 2f;
            if (hover) using (var path = Ui.RoundRect(new RectangleF(cx - size / 2 + 1, cy - size / 2 + 1, size - 2, size - 2), 6 * u)) using (var fill = new SolidBrush(Ui.Btn2Hover)) g.FillPath(fill, path);
            Color ink = hover ? Ui.Blue : Color.FromArgb(120, 132, 152);
            float outer = 7.5f * u, inner = 5.6f * u, tooth = 2.6f * u;
            var state = g.Save();
            g.TranslateTransform(cx, cy);
            using (var brush = new SolidBrush(ink))
            {
                for (int i = 0; i < 8; i++) { g.FillRectangle(brush, -tooth / 2, -outer, tooth, outer - inner + 1.5f * u); g.RotateTransform(45); }
                g.FillEllipse(brush, -inner, -inner, inner * 2, inner * 2);
            }
            float hole = 2.4f * u;
            using (var back = new SolidBrush(hover ? Ui.Btn2Hover : (Parent == null ? Ui.Canvas : Parent.BackColor))) g.FillEllipse(back, -hole, -hole, hole * 2, hole * 2);
            g.Restore(state);
        }
    }

    internal sealed class MigrationEdgeToggle : Control
    {
        private bool open, hover;
        internal bool Open { get { return open; } set { if (open != value) { open = value; Invalidate(); } } }
        internal MigrationEdgeToggle()
        {
            TabStop = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        // Only a triangle at the top of its column: ◂ opens the range panel,
        // ▸ closes it. The control fills the column so layout cannot clip it.
        private Rectangle Tab { get { int size = Math.Max(Font.Height * 2, 24); return new Rectangle(0, 0, Width, Math.Min(Height, size)); } }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            bool inside = Tab.Contains(e.Location);
            Cursor = inside ? Cursors.Hand : Cursors.Default;
            if (inside != hover) { hover = inside; Invalidate(); }
            base.OnMouseMove(e);
        }
        protected override void OnClick(EventArgs e) { if (Tab.Contains(PointToClient(Control.MousePosition))) base.OnClick(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.Clear(Parent == null ? Ui.Canvas : Parent.BackColor); g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = Tab; float u = g.DpiX / 96f;
            if (hover) using (var path = Ui.RoundRect(new RectangleF(rect.Left + 1, rect.Top + 1, rect.Width - 2, rect.Height - 2), 6 * u)) using (var fill = new SolidBrush(Ui.Btn2Hover)) g.FillPath(fill, path);
            float cx = rect.Left + rect.Width / 2f, cy = rect.Top + rect.Height / 2f, w = 6f * u, h = 9f * u;
            var points = open
                ? new[] { new PointF(cx - w / 2, cy - h), new PointF(cx + w, cy), new PointF(cx - w / 2, cy + h) }
                : new[] { new PointF(cx + w / 2, cy - h), new PointF(cx - w, cy), new PointF(cx + w / 2, cy + h) };
            using (var brush = new SolidBrush(hover ? Ui.Blue : Color.FromArgb(120, 132, 152))) g.FillPolygon(brush, points);
        }
    }

    // Closes a dropdown or menu when the mouse is pressed anywhere outside it.
    // Inside WPS the add-in's windows get no WinForms message filter, so a
    // ToolStripDropDown would otherwise stay open after an outside click.
    internal static class MigrationDropDownWatch
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int key);
        private static bool Down() { return (GetAsyncKeyState(0x01) & 0x8000) != 0 || (GetAsyncKeyState(0x02) & 0x8000) != 0 || (GetAsyncKeyState(0x04) & 0x8000) != 0; }
        internal static void Attach(ToolStripDropDown dropDown)
        {
            if (dropDown == null || dropDown.Tag is Timer) return;
            var timer = new Timer { Interval = 30 }; bool wasDown = true;
            dropDown.Tag = timer;
            timer.Tick += delegate
            {
                if (dropDown.IsDisposed || !dropDown.Visible) { timer.Stop(); return; }
                bool down = Down();
                // Only a new press counts, so the click that opened it, or a
                // drag that started inside (resizing), never closes it.
                if (down && !wasDown && !dropDown.Bounds.Contains(Cursor.Position)) dropDown.Close(ToolStripDropDownCloseReason.AppClicked);
                wasDown = down;
            };
            dropDown.Opened += delegate { wasDown = true; timer.Start(); };
            dropDown.Closed += delegate { timer.Stop(); };
            dropDown.Disposed += delegate { timer.Stop(); timer.Dispose(); };
            if (dropDown.Visible) timer.Start();
        }
    }

    // 映射 board: one row per mapping, the source range on the left and the
    // target range on the right, each a card. Press and hold a card (or just
    // start dragging it) to lift it as a floating rounded card; drop it on
    // another row's card of the same side and the two rows swap that side.
    internal sealed class MigrationMappingBoard : Control
    {
        internal sealed class Entry
        {
            internal string SourceTitle = "", SourceDetail = "", TargetTitle = "", TargetDetail = "", Status = "", SourceTip = "", TargetTip = "", StatusTip = "";
            internal bool SourceUsed, TargetUsed, Enabled = true, ValuesOnly;
            // 0 neutral, 1 done, 2 partial, 3 problem, 4 ready
            internal int Tone;
        }
        private enum Part { None, Check, Source, Target, Mode, Status }
        private struct Cols { internal Rectangle Check, Number, Source, Arrow, Target, Mode, Status; }
        private readonly List<Entry> items = new List<Entry>();
        private readonly VScrollBar scroll = new VScrollBar { Dock = DockStyle.Right, Visible = false, TabStop = false };
        private readonly Timer hold = new Timer { Interval = 200 }, edgeScroll = new Timer { Interval = 40 };
        private readonly ToolTip tip = new ToolTip();
        private Font body = Ui.MakeFont(10f, FontStyle.Regular, null), head = Ui.MakeFont(9f, FontStyle.Bold, null);
        private float scale = 1f;
        private int current = -1, hoverRow = -1, pressRow = -1, dropRow = -1;
        private Part hoverPart, pressPart;
        private bool currentTarget, dragging;
        private Point pressPoint, mouse;
        private Size grab;
        private string tipText = "";
        private bool readOnly;
        internal string EmptyText = "";
        internal event Action CurrentChanged, DeleteRequested;
        internal event Action<int, int, bool> SwapRequested;
        internal event Action<int> EnabledToggled, ModeToggled;
        internal event Action<int, bool> SideActivated;

        internal MigrationMappingBoard()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            BackColor = Color.White; TabStop = true;
            Controls.Add(scroll);
            scroll.ValueChanged += delegate { Invalidate(); };
            scroll.VisibleChanged += delegate { Invalidate(); };
            hold.Tick += delegate { hold.Stop(); if (pressRow >= 0 && (Control.MouseButtons & MouseButtons.Left) != 0) BeginDrag(); };
            edgeScroll.Tick += delegate { AutoScrollStep(); };
        }
        internal int Count { get { return items.Count; } }
        internal Entry this[int index] { get { return items[index]; } }
        internal int Current { get { return current; } }
        internal bool CurrentTarget { get { return currentTarget; } }
        internal bool Dragging { get { return dragging; } }
        internal bool ReadOnly { get { return readOnly; } set { if (readOnly != value) { readOnly = value; if (value) CancelDrag(); Invalidate(); } } }
        // Replace every row; the selection moves to index (clamped).
        internal void SetItems(IList<Entry> list, int index, bool target)
        {
            CancelDrag(); items.Clear(); items.AddRange(list);
            current = items.Count == 0 ? -1 : Math.Max(0, Math.Min(index, items.Count - 1)); currentTarget = target;
            UpdateScroll(); EnsureVisible(current); Invalidate();
        }
        internal void UpdateItem(int index, Entry item) { if (index < 0 || index >= items.Count) return; items[index] = item; Invalidate(); }
        internal void SetCurrent(int index, bool target)
        {
            if (index >= items.Count) index = items.Count - 1;
            if (index == current && target == currentTarget) return;
            current = index; currentTarget = target; EnsureVisible(index); Invalidate();
            if (CurrentChanged != null) CurrentChanged();
        }

        // ---------------- geometry ----------------
        private int P(int value) { return (int)Math.Round(value * scale); }
        private int HeaderHeight { get { return P(38); } }
        private int RowHeight { get { return body.Height * 2 + P(24); } }
        private int ViewWidth { get { return Math.Max(0, ClientSize.Width - (scroll.Visible ? scroll.Width : 0)); } }
        private int Offset { get { return scroll.Visible ? scroll.Value : 0; } }
        private int RowTop(int index) { return HeaderHeight + index * RowHeight - Offset; }
        private Cols Columns()
        {
            int pad = P(10), check = Math.Max(P(30), TextRenderer.MeasureText("启用", head).Width + P(14)), number = Math.Max(P(30), TextRenderer.MeasureText("行", head).Width + P(14)), arrow = P(30), mode = P(86), gap = P(10);
            int width = Math.Max(P(480), ViewWidth - pad * 2);
            int rest = width - check - number - arrow - mode - gap * 2;
            int status = Math.Max(P(110), rest * 26 / 100), card = Math.Max(P(130), (rest - status) / 2);
            var c = new Cols(); int x = pad;
            c.Check = new Rectangle(x, 0, check, 0); x += check;
            c.Number = new Rectangle(x, 0, number, 0); x += number;
            c.Source = new Rectangle(x, 0, card, 0); x += card;
            c.Arrow = new Rectangle(x, 0, arrow, 0); x += arrow;
            c.Target = new Rectangle(x, 0, card, 0); x += card + gap;
            c.Mode = new Rectangle(x, 0, mode, 0); x += mode + gap;
            c.Status = new Rectangle(x, 0, Math.Max(P(40), pad + width - x), 0);
            return c;
        }
        private Rectangle Card(Rectangle column, int top) { return new Rectangle(column.Left, top + P(6), column.Width, RowHeight - P(12)); }
        private Rectangle CardOf(int row, bool target) { var c = Columns(); return Card(target ? c.Target : c.Source, RowTop(row)); }
        private Rectangle Pill(Rectangle column, int top) { int height = body.Height + P(10); return new Rectangle(column.Left, top + (RowHeight - height) / 2, column.Width, height); }
        private int RowAt(int y)
        {
            if (y < HeaderHeight) return -1;
            int row = (y - HeaderHeight + Offset) / RowHeight;
            return row >= 0 && row < items.Count ? row : -1;
        }
        private Part PartAt(Point point, out int row)
        {
            row = RowAt(point.Y); if (row < 0 || point.X >= ViewWidth) return Part.None;
            var c = Columns(); int top = RowTop(row);
            if (point.X >= c.Check.Left && point.X < c.Check.Right) return Part.Check;
            if (Card(c.Source, top).Contains(point)) return Part.Source;
            if (Card(c.Target, top).Contains(point)) return Part.Target;
            if (Pill(c.Mode, top).Contains(point)) return Part.Mode;
            if (point.X >= c.Status.Left) return Part.Status;
            return Part.None;
        }
        private bool Used(int row, bool target) { return row >= 0 && row < items.Count && (target ? items[row].TargetUsed : items[row].SourceUsed); }

        // ---------------- scrolling ----------------
        private void UpdateScroll()
        {
            int view = Math.Max(1, ClientSize.Height - HeaderHeight), total = items.Count * RowHeight;
            if (total <= view) { scroll.Value = 0; scroll.Visible = false; return; }
            scroll.Minimum = 0; scroll.SmallChange = Math.Max(1, RowHeight / 2); scroll.LargeChange = view; scroll.Maximum = total - 1;
            if (scroll.Value > total - view) scroll.Value = total - view;
            scroll.Visible = true;
        }
        private void ScrollTo(int value)
        {
            if (!scroll.Visible) return;
            scroll.Value = Math.Max(0, Math.Min(value, scroll.Maximum - scroll.LargeChange + 1));
        }
        private void EnsureVisible(int row)
        {
            if (row < 0 || !scroll.Visible) return;
            int top = row * RowHeight, view = ClientSize.Height - HeaderHeight;
            if (top < scroll.Value) ScrollTo(top);
            else if (top + RowHeight > scroll.Value + view) ScrollTo(top + RowHeight - view);
        }
        protected override void OnMouseWheel(MouseEventArgs e) { ScrollTo(Offset - e.Delta / 120 * RowHeight); base.OnMouseWheel(e); }
        protected override void OnResize(EventArgs e) { UpdateScroll(); base.OnResize(e); }
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            using (var g = CreateGraphics()) scale = g.DpiX / 96f;
            UpdateScroll();
        }

        // ---------------- painting ----------------
        private const TextFormatFlags Line = TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
        private const TextFormatFlags CardLine = TextFormatFlags.SingleLine | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
        private static GraphicsPath Round(Rectangle rect, float radius) { return Ui.RoundRect(new RectangleF(rect.X + .5f, rect.Y + .5f, rect.Width - 1f, rect.Height - 1f), radius); }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; scale = g.DpiX / 96f; g.Clear(BackColor);
            var c = Columns(); int header = HeaderHeight, width = ViewWidth;
            g.SetClip(new Rectangle(0, header, width, Math.Max(0, ClientSize.Height - header)));
            int first = Math.Max(0, Offset / RowHeight), last = Math.Min(items.Count - 1, (Offset + ClientSize.Height) / RowHeight + 1);
            for (int i = first; i <= last; i++) DrawRow(g, i, c, RowTop(i));
            if (items.Count == 0 && EmptyText.Length > 0)
                TextRenderer.DrawText(g, EmptyText, body, new Rectangle(P(24), header, Math.Max(0, width - P(48)), Math.Max(0, ClientSize.Height - header)), Ui.InkDim,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
            g.ResetClip();
            using (var brush = new SolidBrush(Ui.GridHead)) g.FillRectangle(brush, 0, 0, width, header);
            using (var pen = new Pen(Ui.GridFrame)) g.DrawLine(pen, 0, header - 1, width, header - 1);
            HeadText(g, "启用", c.Check, true); HeadText(g, "行", c.Number, true);
            HeadText(g, "源区域", new Rectangle(c.Source.Left + P(30), 0, c.Source.Width - P(30), 0), false);
            HeadText(g, "目标区域", new Rectangle(c.Target.Left + P(30), 0, c.Target.Width - P(30), 0), false);
            HeadText(g, "方式", c.Mode, true); HeadText(g, "状态", new Rectangle(c.Status.Left + P(4), 0, c.Status.Width, 0), false);
            if (dragging) DrawLifted(g);
        }
        private void HeadText(Graphics g, string text, Rectangle column, bool center)
        {
            TextRenderer.DrawText(g, text, head, new Rectangle(column.Left, 0, Math.Max(0, column.Width), HeaderHeight), Ui.GridHeadInk,
                Line | (center ? TextFormatFlags.HorizontalCenter : TextFormatFlags.Left));
        }
        private void DrawRow(Graphics g, int i, Cols c, int top)
        {
            var item = items[i]; int height = RowHeight, width = ViewWidth;
            Color back = i == current ? Color.FromArgb(246, 249, 255) : (i == hoverRow && !dragging ? Color.FromArgb(250, 251, 253) : Color.White);
            using (var brush = new SolidBrush(back)) g.FillRectangle(brush, 0, top, width, height);
            using (var pen = new Pen(Ui.GridLine)) g.DrawLine(pen, 0, top + height - 1, width, top + height - 1);
            if (i == current) using (var brush = new SolidBrush(Ui.Blue)) g.FillRectangle(brush, 0, top + P(8), P(3), height - P(16));
            g.SmoothingMode = SmoothingMode.AntiAlias;
            // 启用: rounded check box
            int box = P(16); var check = new Rectangle(c.Check.Left + (c.Check.Width - box) / 2, top + (height - box) / 2, box, box);
            using (var path = Round(check, P(4)))
            {
                if (item.Enabled) { using (var brush = new SolidBrush(readOnly ? Color.FromArgb(150, 170, 210) : Ui.Blue)) g.FillPath(brush, path); }
                else { g.FillPath(Brushes.White, path); using (var pen = new Pen(Color.FromArgb(190, 199, 214))) g.DrawPath(pen, path); }
            }
            if (item.Enabled)
                using (var pen = new Pen(Color.White, Math.Max(1.5f, 1.8f * scale)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                    g.DrawLines(pen, new[] { new PointF(check.Left + box * .25f, check.Top + box * .52f), new PointF(check.Left + box * .43f, check.Top + box * .7f), new PointF(check.Left + box * .76f, check.Top + box * .32f) });
            TextRenderer.DrawText(g, (i + 1).ToString(CultureInfo.InvariantCulture), body, new Rectangle(c.Number.Left, top, c.Number.Width, height), Ui.InkDim, Line | TextFormatFlags.HorizontalCenter);
            DrawSlot(g, i, false, Card(c.Source, top));
            DrawSlot(g, i, true, Card(c.Target, top));
            // Arrow between the two cards.
            float ax = c.Arrow.Left + c.Arrow.Width / 2f, ay = top + height / 2f, u = scale;
            using (var pen = new Pen(item.SourceUsed && item.TargetUsed ? Color.FromArgb(150, 165, 190) : Color.FromArgb(210, 216, 226), 1.4f * u) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            { g.DrawLine(pen, ax - 7 * u, ay, ax + 6 * u, ay); g.DrawLines(pen, new[] { new PointF(ax + 2 * u, ay - 4 * u), new PointF(ax + 6.5f * u, ay), new PointF(ax + 2 * u, ay + 4 * u) }); }
            // 方式: a pill, click to switch.
            var pill = Pill(c.Mode, top); bool values = item.ValuesOnly;
            bool hot = !dragging && !readOnly && i == hoverRow && hoverPart == Part.Mode;
            Color pillBack = values ? Color.FromArgb(hot ? 253 : 255, hot ? 240 : 246, hot ? 214 : 229) : (hot ? Color.FromArgb(226, 235, 254) : Color.FromArgb(238, 243, 255));
            Color pillInk = !item.Enabled ? Ui.InkDim : values ? Color.FromArgb(180, 110, 10) : Ui.Blue;
            using (var path = Round(pill, pill.Height / 2f)) using (var brush = new SolidBrush(pillBack)) g.FillPath(brush, path);
            TextRenderer.DrawText(g, values ? DataMigrationForm.ModeValues : DataMigrationForm.ModeFormula, body, pill, pillInk, Line | TextFormatFlags.HorizontalCenter);
            // 状态: coloured dot + text
            Color dot = item.Tone == 1 ? Ui.Green : item.Tone == 2 ? Ui.Amber : item.Tone == 3 ? Ui.Danger : item.Tone == 4 ? Ui.Blue : Color.FromArgb(170, 178, 192);
            int d = P(8), sx = c.Status.Left + P(4);
            using (var brush = new SolidBrush(dot)) g.FillEllipse(brush, sx, top + (height - d) / 2, d, d);
            TextRenderer.DrawText(g, item.Status, body, new Rectangle(sx + d + P(6), top, Math.Max(0, c.Status.Right - sx - d - P(6)), height),
                item.Tone == 3 ? Ui.Danger : item.Enabled ? Ui.Ink : Ui.InkDim, Line | TextFormatFlags.Left);
            g.SmoothingMode = SmoothingMode.Default;
        }
        private void DrawSlot(Graphics g, int i, bool target, Rectangle rect)
        {
            var item = items[i]; bool used = target ? item.TargetUsed : item.SourceUsed;
            bool mySide = target == (pressPart == Part.Target);
            bool lifted = dragging && i == pressRow && mySide, drop = dragging && i == dropRow && dropRow != pressRow && mySide;
            bool focus = !dragging && i == current && target == currentTarget;
            bool hover = !dragging && i == hoverRow && hoverPart == (target ? Part.Target : Part.Source);
            float radius = P(8);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (lifted)
            {
                // The slot left behind while its card is in the air.
                using (var path = Round(rect, radius)) using (var brush = new SolidBrush(Ui.Canvas)) using (var pen = new Pen(Color.FromArgb(190, 199, 214), scale) { DashStyle = DashStyle.Dash })
                { g.FillPath(brush, path); g.DrawPath(pen, path); }
                return;
            }
            if (drop)
            {
                using (var path = Round(rect, radius)) using (var brush = new SolidBrush(Color.FromArgb(232, 240, 254))) using (var pen = new Pen(Ui.Blue, 1.5f * scale) { DashStyle = DashStyle.Dash })
                { g.FillPath(brush, path); g.DrawPath(pen, path); }
                TextRenderer.DrawText(g, used ? "松开与第 " + (pressRow + 1) + " 行交换" : "松开移到这一行", body, rect, Ui.Blue, Line | TextFormatFlags.HorizontalCenter);
                return;
            }
            if (!used)
            {
                using (var path = Round(rect, radius)) using (var pen = new Pen(Color.FromArgb(205, 212, 224), scale) { DashStyle = DashStyle.Dash }) g.DrawPath(pen, path);
                TextRenderer.DrawText(g, "未配对 · 可把" + (target ? "目标" : "源") + "区域拖到这里", body, rect, Ui.InkDim, Line | TextFormatFlags.HorizontalCenter);
                return;
            }
            using (var path = Round(rect, radius))
            using (var brush = new SolidBrush(hover ? Ui.Btn2Hover : Color.White))
            using (var pen = new Pen(focus ? Ui.Blue : hover ? Color.FromArgb(175, 195, 235) : Ui.GridFrame, focus ? 1.5f * scale : scale))
            { g.FillPath(brush, path); g.DrawPath(pen, path); }
            DrawCardText(g, rect, target ? item.TargetTitle : item.SourceTitle, target ? item.TargetDetail : item.SourceDetail, item.Enabled);
        }
        private void DrawCardText(Graphics g, Rectangle rect, string title, string detail, bool enabled)
        {
            // Grip dots: the whole card drags, the dots only say so.
            float u = scale, gx = rect.Left + P(10), gy = rect.Top + rect.Height / 2f;
            using (var brush = new SolidBrush(readOnly ? Color.FromArgb(215, 220, 230) : Color.FromArgb(170, 180, 198)))
                for (int row = -1; row <= 1; row++) for (int col = 0; col < 2; col++) g.FillEllipse(brush, gx + col * 5 * u, gy + row * 5 * u - 1.2f * u, 2.6f * u, 2.6f * u);
            int left = rect.Left + P(26), width = Math.Max(0, rect.Right - P(10) - left), half = rect.Height / 2;
            TextRenderer.DrawText(g, title, body, new Rectangle(left, rect.Top + P(3), width, half - P(2)), enabled ? Ui.Ink : Ui.InkDim, CardLine | TextFormatFlags.Bottom);
            TextRenderer.DrawText(g, detail, body, new Rectangle(left, rect.Top + half, width, half - P(3)), Ui.InkDim, CardLine | TextFormatFlags.Top);
        }
        // The card in the air: follows the mouse with a soft shadow.
        private void DrawLifted(Graphics g)
        {
            bool target = pressPart == Part.Target; if (!Used(pressRow, target)) return;
            var origin = CardOf(pressRow, target);
            int y = Math.Max(HeaderHeight - origin.Height / 2, Math.Min(mouse.Y - grab.Height, ClientSize.Height - origin.Height / 2));
            var rect = new Rectangle(origin.Left + Math.Max(-P(16), Math.Min(P(16), mouse.X - grab.Width - origin.Left)), y, origin.Width, origin.Height);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float radius = P(12);
            for (int k = 4; k >= 1; k--)
            {
                var shadow = new Rectangle(rect.X - k * P(1), rect.Y - k * P(1) + P(3), rect.Width + k * P(2), rect.Height + k * P(2));
                using (var path = Round(shadow, radius + k * P(1))) using (var brush = new SolidBrush(Color.FromArgb(14, 20, 40, 80))) g.FillPath(brush, path);
            }
            using (var path = Round(rect, radius)) using (var pen = new Pen(Ui.Blue, 1.5f * scale)) { g.FillPath(Brushes.White, path); g.DrawPath(pen, path); }
            var item = items[pressRow];
            DrawCardText(g, rect, target ? item.TargetTitle : item.SourceTitle, target ? item.TargetDetail : item.SourceDetail, true);
            g.SmoothingMode = SmoothingMode.Default;
        }

        // ---------------- mouse and keys ----------------
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e); Focus();
            int row; var part = PartAt(e.Location, out row);
            if (row < 0) return;
            bool side = part == Part.Source || part == Part.Target;
            SetCurrent(row, side ? part == Part.Target : currentTarget);
            if (e.Button != MouseButtons.Left) return;
            pressRow = row; pressPart = part; pressPoint = e.Location; mouse = e.Location;
            if (side && !readOnly && Used(row, part == Part.Target))
            { var card = CardOf(row, part == Part.Target); grab = new Size(e.X - card.Left, e.Y - card.Top); hold.Start(); }
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e); mouse = e.Location;
            if (dragging) { UpdateDrop(); Invalidate(); return; }
            if (hold.Enabled && (Math.Abs(e.X - pressPoint.X) > SystemInformation.DragSize.Width || Math.Abs(e.Y - pressPoint.Y) > SystemInformation.DragSize.Height))
            { hold.Stop(); BeginDrag(); return; }
            int row; var part = PartAt(e.Location, out row);
            if (row != hoverRow || part != hoverPart) { hoverRow = row; hoverPart = part; Invalidate(); }
            bool card = (part == Part.Source || part == Part.Target) && Used(row, part == Part.Target);
            Cursor = !readOnly && (card || part == Part.Check || part == Part.Mode) ? Cursors.Hand : Cursors.Default;
            string text = row < 0 ? "" : part == Part.Source ? items[row].SourceTip : part == Part.Target ? items[row].TargetTip
                : part == Part.Status ? items[row].StatusTip : part == Part.Mode ? "点击切换：带公式 / 只写数值" : part == Part.Check ? "启用或停用这一行" : "";
            if (card && !readOnly) text = (text.Length == 0 ? "" : text + Environment.NewLine) + "按住拖动到另一行可交换；双击编辑";
            if (text != tipText) { tipText = text; tip.SetToolTip(this, text); }
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (!dragging && (hoverRow != -1 || hoverPart != Part.None)) { hoverRow = -1; hoverPart = Part.None; Cursor = Cursors.Default; Invalidate(); }
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e); hold.Stop();
            if (dragging)
            {
                int from = pressRow, to = dropRow; bool target = pressPart == Part.Target;
                EndDrag();
                if (to >= 0 && to != from && SwapRequested != null) SwapRequested(from, to, target);
                return;
            }
            if (e.Button == MouseButtons.Left && pressRow >= 0 && !readOnly)
            {
                int row; var part = PartAt(e.Location, out row);
                if (row == pressRow && part == pressPart)
                {
                    if (part == Part.Check && EnabledToggled != null) EnabledToggled(row);
                    else if (part == Part.Mode && ModeToggled != null) ModeToggled(row);
                }
            }
            pressRow = -1; pressPart = Part.None;
        }
        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            int row; var part = PartAt(e.Location, out row);
            if ((part == Part.Source || part == Part.Target) && SideActivated != null) SideActivated(row, part == Part.Target);
        }
        protected override void OnMouseCaptureChanged(EventArgs e) { base.OnMouseCaptureChanged(e); if (dragging && !Capture) CancelDrag(); }
        private void BeginDrag()
        {
            if (readOnly || pressRow < 0) return;
            dragging = true; dropRow = pressRow; Cursor = Cursors.SizeNS; tip.Hide(this); edgeScroll.Start(); Invalidate();
        }
        private void UpdateDrop()
        {
            int y = Math.Max(HeaderHeight + 1, Math.Min(mouse.Y, ClientSize.Height - 2));
            int row = RowAt(y); if (row < 0) row = items.Count - 1;
            dropRow = row;
        }
        private void AutoScrollStep()
        {
            if (!dragging) { edgeScroll.Stop(); return; }
            int edge = P(24);
            if (mouse.Y < HeaderHeight + edge) ScrollTo(Offset - P(12));
            else if (mouse.Y > ClientSize.Height - edge) ScrollTo(Offset + P(12));
            else return;
            UpdateDrop(); Invalidate();
        }
        private void EndDrag() { dragging = false; dropRow = -1; pressRow = -1; pressPart = Part.None; hold.Stop(); edgeScroll.Stop(); Cursor = Cursors.Default; Invalidate(); }
        private void CancelDrag() { if (dragging || hold.Enabled) EndDrag(); }
        protected override bool IsInputKey(Keys keyData)
        {
            Keys key = keyData & Keys.KeyCode;
            return key == Keys.Up || key == Keys.Down || key == Keys.Left || key == Keys.Right || base.IsInputKey(keyData);
        }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape && dragging) { CancelDrag(); e.Handled = true; return; }
            if (items.Count == 0 || dragging) return;
            int row = Math.Max(0, current);
            if ((e.KeyCode == Keys.Up || e.KeyCode == Keys.Down) && e.Alt)
            {
                int to = row + (e.KeyCode == Keys.Down ? 1 : -1);
                if (!readOnly && to >= 0 && to < items.Count && SwapRequested != null) SwapRequested(row, to, currentTarget);
            }
            else if (e.KeyCode == Keys.Up) SetCurrent(Math.Max(0, row - 1), currentTarget);
            else if (e.KeyCode == Keys.Down) SetCurrent(Math.Min(items.Count - 1, row + 1), currentTarget);
            else if (e.KeyCode == Keys.Left) SetCurrent(row, false);
            else if (e.KeyCode == Keys.Right) SetCurrent(row, true);
            else if (e.KeyCode == Keys.Space && !readOnly && EnabledToggled != null) EnabledToggled(row);
            else if (e.KeyCode == Keys.Enter && SideActivated != null) SideActivated(row, currentTarget);
            else if (e.KeyCode == Keys.Delete && !readOnly && DeleteRequested != null) DeleteRequested();
            else return;
            e.Handled = true; e.SuppressKeyPress = true;
        }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { hold.Dispose(); edgeScroll.Dispose(); tip.Dispose(); body.Dispose(); head.Dispose(); }
            base.Dispose(disposing);
        }
    }

    // Page switcher in the window's own style: plain text tabs with a count
    // badge, the selected one in ink with a blue underline.
    internal sealed class MigrationTabStrip : Control
    {
        private readonly List<string> titles = new List<string>();
        private readonly List<int> counts = new List<int>();
        private Font font = Ui.MakeFont(10f, FontStyle.Regular, null), bold = Ui.MakeFont(10f, FontStyle.Bold, null), small = Ui.MakeFont(8.5f, FontStyle.Regular, null);
        private int selected, hover = -1;
        internal event Action SelectedIndexChanged;
        internal MigrationTabStrip()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Ui.Canvas; Height = font.Height + 22;
        }
        internal int TabCount { get { return titles.Count; } }
        internal string Title(int index) { return titles[index]; }
        internal int Add(string title) { titles.Add(title); counts.Add(-1); Invalidate(); return titles.Count - 1; }
        internal void SetCount(int index, int count) { if (counts[index] != count) { counts[index] = count; Invalidate(); } }
        internal int SelectedIndex
        {
            get { return selected; }
            set { if (value < 0 || value >= titles.Count || value == selected) return; selected = value; Invalidate(); if (SelectedIndexChanged != null) SelectedIndexChanged(); }
        }
        private int P(int value, Graphics g) { return (int)Math.Round(value * g.DpiX / 96f); }
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); using (var g = CreateGraphics()) Height = font.Height + P(22, g); }
        private List<Rectangle> Tabs(Graphics g)
        {
            var result = new List<Rectangle>(); int x = 0;
            for (int i = 0; i < titles.Count; i++)
            {
                int width = TextRenderer.MeasureText(g, titles[i], bold, Size.Empty, TextFormatFlags.NoPadding).Width + P(28, g);
                if (counts[i] >= 0) width += TextRenderer.MeasureText(g, counts[i].ToString(CultureInfo.InvariantCulture), small, Size.Empty, TextFormatFlags.NoPadding).Width + P(18, g);
                result.Add(new Rectangle(x, 0, width, Height)); x += width + P(2, g);
            }
            return result;
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.Clear(BackColor);
            using (var pen = new Pen(Ui.GridFrame)) g.DrawLine(pen, 0, Height - 1, Width, Height - 1);
            var tabs = Tabs(g);
            for (int i = 0; i < tabs.Count; i++)
            {
                var tab = tabs[i]; bool on = i == selected, hot = i == hover && !on;
                var label = on ? bold : font;
                int textWidth = TextRenderer.MeasureText(g, titles[i], bold, Size.Empty, TextFormatFlags.NoPadding).Width;
                int x = tab.Left + P(14, g);
                TextRenderer.DrawText(g, titles[i], label, new Rectangle(x, 0, textWidth + P(4, g), Height - P(3, g)), on ? Ui.Ink : hot ? Ui.Blue : Ui.InkDim,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding);
                if (counts[i] >= 0)
                {
                    string count = counts[i].ToString(CultureInfo.InvariantCulture);
                    int cw = TextRenderer.MeasureText(g, count, small, Size.Empty, TextFormatFlags.NoPadding).Width + P(12, g), ch = small.Height + P(4, g);
                    var badge = new Rectangle(x + textWidth + P(6, g), (Height - P(3, g) - ch) / 2, Math.Max(cw, ch), ch);
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    using (var path = Ui.RoundRect(badge, ch / 2f)) using (var brush = new SolidBrush(on ? Ui.SelectBack : Color.FromArgb(234, 237, 243))) g.FillPath(brush, path);
                    g.SmoothingMode = SmoothingMode.Default;
                    TextRenderer.DrawText(g, count, small, badge, on ? Ui.Blue : Ui.InkDim, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
                }
                if (on)
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    var bar = new Rectangle(tab.Left + P(10, g), Height - P(3, g), tab.Width - P(20, g), P(3, g));
                    using (var path = Ui.RoundRect(bar, P(2, g))) using (var brush = new SolidBrush(Ui.Blue)) g.FillPath(brush, path);
                    g.SmoothingMode = SmoothingMode.Default;
                }
            }
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e); int index = -1;
            using (var g = CreateGraphics()) { var tabs = Tabs(g); for (int i = 0; i < tabs.Count; i++) if (tabs[i].Contains(e.Location)) index = i; }
            Cursor = index >= 0 ? Cursors.Hand : Cursors.Default;
            if (index != hover) { hover = index; Invalidate(); }
        }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); if (hover != -1) { hover = -1; Invalidate(); } }
        protected override void OnMouseClick(MouseEventArgs e) { base.OnMouseClick(e); if (e.Button == MouseButtons.Left && hover >= 0) SelectedIndex = hover; }
        protected override void Dispose(bool disposing) { if (disposing) { font.Dispose(); bold.Dispose(); small.Dispose(); } base.Dispose(disposing); }
    }

    // Right-click menus in the window's style: white, hairline frame,
    // rounded light-blue hover, dimmed shortcut text.
    internal sealed class MigrationMenuRenderer : ToolStripProfessionalRenderer
    {
        private sealed class Palette : ProfessionalColorTable
        {
            public override Color ToolStripDropDownBackground { get { return Color.White; } }
            public override Color ImageMarginGradientBegin { get { return Color.White; } }
            public override Color ImageMarginGradientMiddle { get { return Color.White; } }
            public override Color ImageMarginGradientEnd { get { return Color.White; } }
            public override Color MenuBorder { get { return Color.FromArgb(160, 170, 186); } }
            public override Color MenuItemBorder { get { return Color.Transparent; } }
            public override Color MenuItemSelected { get { return Ui.SelectBack; } }
            public override Color SeparatorDark { get { return Ui.GridLine; } }
            public override Color SeparatorLight { get { return Color.White; } }
        }
        internal MigrationMenuRenderer() : base(new Palette()) { RoundedEdges = false; }
        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (!e.Item.Selected || !e.Item.Enabled) return;
            float u = e.Graphics.DpiX / 96f;
            var rect = new RectangleF(2 * u, 1, e.Item.Width - 4 * u, e.Item.Height - 2);
            var smoothing = e.Graphics.SmoothingMode; e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = Ui.RoundRect(rect, 6 * u)) using (var brush = new SolidBrush(Ui.SelectBack)) e.Graphics.FillPath(brush, path);
            e.Graphics.SmoothingMode = smoothing;
        }
        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            var item = e.Item as ToolStripMenuItem;
            bool shortcut = item != null && !string.IsNullOrEmpty(item.ShortcutKeyDisplayString) && e.Text == item.ShortcutKeyDisplayString;
            e.TextColor = !e.Item.Enabled || shortcut ? Ui.InkDim : e.Item.ForeColor == Ui.Danger ? Ui.Danger : e.Item.Selected ? Ui.SelectInk : Ui.Ink;
            base.OnRenderItemText(e);
        }
    }

    // 批量迁移数据: 源区域 and 目标区域 are kept in their own pages; 映射 pairs
    // them row by row (drag ⋮⋮ there to re-pair). A range panel on the right
    // (opened from the window edge) shows the sheet so ranges are dragged.
    internal sealed class DataMigrationForm : Form, IDialogDefaultLayout
    {
        internal const string ModeFormula = "带公式", ModeValues = "只写数值";
        private readonly CheckBox allowEmpty = new CheckBox { Name = "migrationAllowEmpty", Text = "允许全空白源区域（清空目标）", AutoSize = true };
        private readonly CheckBox replaceTargets = new CheckBox { Name = "migrationReplaceTargets", Text = "更新目标原文件（自动备份）", AutoSize = true, Checked = true };
        internal readonly ResultGrid Sources = new ResultGrid { Name = "migrationSources", Dock = DockStyle.Fill };
        internal readonly ResultGrid Targets = new ResultGrid { Name = "migrationTargets", Dock = DockStyle.Fill };
        internal readonly MigrationMappingBoard Mappings = new MigrationMappingBoard { Name = "migrationMappings", Dock = DockStyle.Fill };
        private readonly List<MigrationPair> pairs = new List<MigrationPair>();
        internal readonly List<MigrationBook> SourceBooks = new List<MigrationBook>(), TargetBooks = new List<MigrationBook>();
        private readonly MigrationTabStrip tabs = new MigrationTabStrip { Name = "migrationTabs", Dock = DockStyle.Fill, Margin = Padding.Empty };
        private Panel sourcesPage, targetsPage, mappingsPage;
        private readonly Panel pageHost = new Panel { Name = "migrationPages", Dock = DockStyle.Fill, Margin = Padding.Empty, BackColor = Ui.Canvas };
        // The mapping the 映射 board has selected, kept across rebuilds.
        private MigrationPair mappingFocus;
        private readonly RichTextBox details = new RichTextBox { Name = "migrationDetails", ReadOnly = true, Dock = DockStyle.Fill, BackColor = Color.White, WordWrap = true };
        private readonly Label status = Ui.Label("在“源区域”“目标区域”页各列出要迁移的区域，“映射”页按行把它们对应起来；右侧“区域选择”可直接框选。", 9, false);
        private readonly RoundedButton preview = Ui.Button("预检与预览", 120), execute = Ui.Button("执行迁移", 110), restore = Ui.Button("还原操作", 110), clearBackups = Ui.Button("清除备份", 110), open = Ui.Button("打开结果文件夹", 150);
        private readonly List<Control> inputs = new List<Control>();
        private SplitContainer split;
        private bool splitPlaced;
        private readonly MigrationEdgeToggle edge = new MigrationEdgeToggle { Name = "migrationSidebarToggle", Dock = DockStyle.Fill };
        private readonly MigrationSheetPreview sheetPreview = new MigrationSheetPreview { Name = "migrationPreview", Dock = DockStyle.Fill };
        private readonly FlowLayoutPanel sideSwitch = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
        private readonly RadioButton showSource = new RadioButton { Name = "migrationShowSource", Text = "源", AutoSize = true, Checked = true };
        private readonly RadioButton showTarget = new RadioButton { Name = "migrationShowTarget", Text = "目标", AutoSize = true };
        private readonly ComboBox previewBook = new ComboBox { Name = "migrationPreviewBook", DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        private sealed class BookItem { internal string Path; public override string ToString() { return System.IO.Path.GetFileName(Path); } }
        private readonly Dictionary<string, MigrationSheetSnapshot> snapshots = new Dictionary<string, MigrationSheetSnapshot>(StringComparer.OrdinalIgnoreCase);
        private MigrationPlan lastPlan;
        // Targets whose original file this window replaced, newest last. A
        // restore puts the pre-migration backup back, but only while the file
        // still holds exactly the bytes the migration wrote.
        private sealed class RestorePoint { internal string Path, Backup, Hash, Displaced; internal bool Restored, Cleared; }
        private readonly List<RestorePoint> restorePoints = new List<RestorePoint>();
        private readonly ToolTip restoreTip = new ToolTip();
        // Keeps the dialog on screen while a run is in progress (see MigrationWindowScope).
        private MigrationWindowScope.HostState hostAtRun;
        private bool topMostBeforeRun;
        private readonly Timer guardTail = new Timer { Interval = 100 };
        private int guardTailTicks;
        private MigrationBookPopup bookPopup;
        private bool sidebarAvailable = true;
        internal bool SidebarAvailable { get { return sidebarAvailable; } }
        private readonly MigrationGearButton gear = new MigrationGearButton { Name = "migrationSettings", Dock = DockStyle.Top };
        private static bool? pickInWorkbook;
        private static string SettingsPath
        { get { return Path.Combine(DialogSizeStore.RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Connect).Assembly.GetName().Name, "DataMigration.settings"); } }
        // What a row's pick button does: true = open the workbook itself
        // (read-only) and select there; false = select in the right-hand panel.
        internal static bool PickInWorkbook
        {
            get
            {
                if (pickInWorkbook == null)
                {
                    bool value = true;
                    try { if (File.Exists(SettingsPath)) foreach (string line in File.ReadAllLines(SettingsPath)) if (line.Trim() == "pick=panel") value = false; }
                    catch (Exception error) { UiDiagnostics.Record("migration-settings", error); }
                    pickInWorkbook = value;
                }
                return pickInWorkbook.Value;
            }
            set
            {
                pickInWorkbook = value;
                try { Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)); File.WriteAllText(SettingsPath, "pick=" + (value ? "workbook" : "panel") + Environment.NewLine, Encoding.UTF8); }
                catch (Exception error) { UiDiagnostics.Record("migration-settings", error); }
            }
        }
        private MigrationSheetPopup sheetPopup;
        private ToolStripDropDown triggerPopup;
        // Lines of the 明细 box that open a choice for a 逐行重算 row.
        private readonly Dictionary<int, KeyValuePair<MigrationSide, MigrationTriggerRow>> detailLinks = new Dictionary<int, KeyValuePair<MigrationSide, MigrationTriggerRow>>();
        // Lines of the details listing an unused target (click: give it to a row).
        private readonly Dictionary<int, KeyValuePair<MigrationSide, string>> unusedLinks = new Dictionary<int, KeyValuePair<MigrationSide, string>>();
        private readonly Timer pump = new Timer { Interval = 15 };
        // Progress text is throttled: rewriting the status line for every
        // cell repainted and re-measured the dialog dozens of times a second.
        private readonly Stopwatch statusClock = new Stopwatch();
        private const int StatusIntervalMilliseconds = 200;
        private IEnumerator<MigrationProgress> steps;
        private Action finished;
        private Action<Exception> failed;
        private bool syncing, closeWhenDone, ticking, fittingGrid, syncingPreview;
        internal bool Running { get { return steps != null; } }
        internal event Action PreviewRequested, ExecuteRequested, CancelRequested, InputsChanged;
        // true = target workbooks, false = source workbooks.
        internal event Action<bool, string[]> BooksPicked;
        // Pick a range in the real workbook (opened read-only by the host):
        // target side?, workbook path, sheet, current range, callback with the
        // chosen address (not called when cancelled).
        internal event Action<bool, string, string, string, Action<string>> RangePickRequested;

        internal DataMigrationForm()
        {
            Text = "批量迁移数据"; DialogSizing.Prepare(this, "migration.data", "table", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 74)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 26)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            SetupSideGrid(Sources, false); SetupSideGrid(Targets, true); SetupMappingBoard();
            sourcesPage = Page(Sources, "＋ 新增源区域", delegate { AddSideFromUi(false); }, "添加源工作簿…", delegate { PickBooks(false, null); });
            targetsPage = Page(Targets, "＋ 新增目标区域", delegate { AddSideFromUi(true); }, "添加目标工作簿…", delegate { PickBooks(true, null); });
            mappingsPage = Page(Mappings, "＋ 新增映射", delegate { AddMappingRow(); }, null, null);
            tabs.Add("源区域"); tabs.Add("目标区域"); tabs.Add("映射");
            foreach (var page in new Control[] { mappingsPage, targetsPage, sourcesPage }) pageHost.Controls.Add(page);
            tabs.SelectedIndexChanged += delegate { ShowPage(tabs.SelectedIndex); };
            var pagesColumn = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = Padding.Empty, Padding = Padding.Empty, BackColor = Ui.Canvas };
            pagesColumn.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pagesColumn.RowStyles.Add(new RowStyle(SizeType.AutoSize)); pagesColumn.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pagesColumn.Controls.Add(tabs, 0, 0); pagesColumn.Controls.Add(pageHost, 0, 1);
            // Pages on the left, range panel on the right, edge toggle beyond.
            split = new SplitContainer { Name = "migrationSplit", Dock = DockStyle.Fill, Orientation = Orientation.Vertical, BackColor = Ui.Canvas, SplitterWidth = 6, Margin = Padding.Empty };
            split.Panel1.Controls.Add(pagesColumn); split.Panel2.Controls.Add(BuildSidebar()); split.Panel2.Padding = new Padding(8, 0, 0, 0);
            split.Layout += delegate { if (split.Width > 300 && !splitPlaced && !split.Panel2Collapsed) { splitPlaced = true; split.SplitterDistance = Math.Max(200, split.Width * 62 / 100); } };
            var main = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty, BackColor = Ui.Canvas };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30));
            main.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            edge.Margin = Padding.Empty; edge.Click += delegate { ShowSidebar(split.Panel2Collapsed); };
            var rail = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = new Padding(2, 0, 0, 0), Padding = Padding.Empty, BackColor = Ui.Canvas };
            rail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            rail.RowStyles.Add(new RowStyle(SizeType.Absolute, 40)); rail.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rail.HandleCreated += delegate { rail.RowStyles[0].Height = Ui.Px(rail, 40); };
            gear.Margin = new Padding(0, 4, 0, 0); gear.Click += delegate { ShowSettings(); };
            rail.Controls.Add(gear, 0, 0); rail.Controls.Add(edge, 0, 1);
            main.Controls.Add(split, 0, 0); main.Controls.Add(rail, 1, 0);
            main.HandleCreated += delegate { main.ColumnStyles[1].Width = Ui.Px(main, 30); };
            root.Controls.Add(main, 0, 0); root.Controls.Add(details, 0, 1);
            details.MouseClick += delegate(object sender, MouseEventArgs e) { if (e.Button == MouseButtons.Left) ShowTriggerChoices(e.Location); };
            details.MouseMove += delegate(object sender, MouseEventArgs e) { int at = DetailLineAt(e.Location); var cursor = detailLinks.ContainsKey(at) || unusedLinks.ContainsKey(at) ? Cursors.Hand : Cursors.IBeam; if (details.Cursor != cursor) details.Cursor = cursor; };
            status.AutoSize = true; status.MaximumSize = new Size(1100, 0); root.Controls.Add(status, 0, 2);
            var actions = Ui.ActionBar(); var close = Ui.Button("关闭", 90);
            preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); };
            restore.Name = "migrationRestore"; restore.Click += delegate { RestoreTargets(); };
            restoreTip.SetToolTip(restore, "把本窗口中以“更新目标原文件”方式迁移过的工作簿，还原为迁移前的备份。");
            clearBackups.Name = "migrationClearBackups"; clearBackups.Click += delegate { ClearBackups(); };
            restoreTip.SetToolTip(clearBackups, "永久删除本窗口迁移时生成的备份文件；删除后无法再还原。");
            restoreTip.SetToolTip(edge, "打开或收起右侧区域选择"); restoreTip.SetToolTip(gear, "预设与设置");
            guardTail.Tick += delegate { GuardWindow(); if (--guardTailTicks <= 0) guardTail.Stop(); };
            open.Click += delegate { OpenResult(); }; close.Click += delegate { Close(); };
            actions.Controls.Add(preview); actions.Controls.Add(execute); actions.Controls.Add(restore); actions.Controls.Add(clearBackups); actions.Controls.Add(open); actions.Controls.Add(close);
            replaceTargets.Margin = new Padding(16, 0, 8, 0); allowEmpty.Margin = new Padding(8, 0, 8, 0);
            actions.Controls.Add(replaceTargets); actions.Controls.Add(allowEmpty); CenterSourceRow(actions);
            root.Controls.Add(actions, 0, 3); Controls.Add(root);
            inputs.AddRange(new Control[] { allowEmpty, replaceTargets });
            allowEmpty.CheckedChanged += delegate { Changed(); };
            replaceTargets.CheckedChanged += delegate { Changed(); Message(replaceTargets.Checked ? "校验通过后更新目标原文件；原版本保存在目标目录的“迁移备份”子文件夹。请先保存并关闭目标。" : "另存“迁移结果”副本，目标原文件不修改；请重新预检。"); };
            pump.Tick += delegate { Tick(); }; execute.Enabled = false; restore.Enabled = false; clearBackups.Enabled = false; open.Enabled = false;
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (Running && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; closeWhenDone = true; if (CancelRequested != null) CancelRequested(); } };
            Disposed += delegate { ClosePopups(); restoreTip.Dispose(); guardTail.Dispose(); pump.Stop(); if (steps != null) { steps.Dispose(); steps = null; } pump.Dispose(); };
            split.Panel2Collapsed = true; edge.Open = false; ShowPage(0); UpdateTabTitles();
        }
        private static Panel Page(Control content, string addText, Action add, string booksText, Action books)
        {
            var page = new Panel { Dock = DockStyle.Fill, BackColor = Ui.Canvas, Padding = new Padding(0, 8, 0, 0), Visible = false };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Margin = Padding.Empty, BackColor = Ui.Canvas };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.Margin = Padding.Empty; layout.Controls.Add(content, 0, 0);
            if (addText != null)
            {
                var bar = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 8, 0, 0), Padding = Padding.Empty, BackColor = Ui.Canvas };
                if (booksText != null)
                {
                    var pick = Ui.Button(booksText, 150); Ui.StyleSecondary(pick); pick.Margin = new Padding(0, 0, 8, 0); pick.Click += delegate { books(); };
                    bar.Controls.Add(pick);
                }
                var button = Ui.Button(addText, 150); Ui.StyleSecondary(button); button.Margin = Padding.Empty; button.Click += delegate { add(); };
                bar.Controls.Add(button);
                layout.Controls.Add(bar, 0, 1);
            }
            page.Controls.Add(layout); return page;
        }
        private Control CurrentPage { get { return tabs.SelectedIndex == 1 ? (Control)targetsPage : tabs.SelectedIndex == 2 ? (Control)mappingsPage : sourcesPage; } }
        // 0 源区域, 1 目标区域, 2 映射.
        internal void ShowPage(int index)
        {
            if (tabs.SelectedIndex != index) { tabs.SelectedIndex = index; return; }
            var current = CurrentPage;
            foreach (var page in new Control[] { sourcesPage, targetsPage, mappingsPage }) page.Visible = page == current;
            // 映射 only pairs ranges; picking happens on the range pages.
            bool mapping = current == mappingsPage;
            if (mapping && split != null && !split.Panel2Collapsed) ShowSidebar(false);
            edge.Visible = !mapping; sidebarAvailable = !mapping;
            current.BringToFront(); ClosePopups(); UpdatePreview(); ShowPairDetails();
        }
        private void UpdateTabTitles()
        {
            tabs.SetCount(0, DistinctSides(false).Count);
            tabs.SetCount(1, DistinctSides(true).Count);
            tabs.SetCount(2, pairs.Count);
        }

        // ---------------- right-hand range panel ----------------
        private Control BuildSidebar()
        {
            var panel = new TableLayoutPanel { Name = "migrationSidebar", Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, BackColor = Ui.Canvas, Margin = Padding.Empty };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var head = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 0, 0, 4), BackColor = Ui.Canvas };
            head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var left = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = Padding.Empty };
            sideSwitch.Controls.Add(new Label { Text = "查看", AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(0, 8, 4, 0) });
            sideSwitch.Controls.Add(showSource); sideSwitch.Controls.Add(showTarget);
            showSource.Margin = showTarget.Margin = new Padding(0, 6, 8, 0);
            left.Controls.Add(sideSwitch);
            left.Controls.Add(new Label { Text = "工作簿", AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(4, 8, 4, 0) });
            left.Controls.Add(previewBook); previewBook.Margin = new Padding(0, 4, 0, 0);
            var close = new Label { Name = "migrationSidebarClose", Text = "×", AutoSize = true, Cursor = Cursors.Hand, ForeColor = Ui.InkDim, Font = Ui.MakeFont(12f, FontStyle.Regular, null), Margin = new Padding(6, 2, 0, 0) };
            close.Click += delegate { ShowSidebar(false); };
            close.MouseEnter += delegate { close.ForeColor = Ui.Blue; }; close.MouseLeave += delegate { close.ForeColor = Ui.InkDim; };
            restoreTip.SetToolTip(close, "收起区域选择（可从窗口右边缘重新打开）");
            head.Controls.Add(left, 0, 0); head.Controls.Add(close, 1, 0);
            panel.Controls.Add(head, 0, 0); panel.Controls.Add(sheetPreview, 0, 1);
            showSource.CheckedChanged += delegate { if (!showSource.Checked) return; if (Mappings.CurrentTarget) Mappings.SetCurrent(Mappings.Current, false); UpdatePreview(); };
            showTarget.CheckedChanged += delegate { if (!showTarget.Checked) return; if (!Mappings.CurrentTarget) Mappings.SetCurrent(Mappings.Current, true); UpdatePreview(); };
            previewBook.SelectedIndexChanged += delegate { if (!syncingPreview) UpdatePreview(); };
            sheetPreview.AreaPicked += PickArea;
            sheetPreview.Sync.Checked = false;
            return panel;
        }
        // Small settings popups: a clear outer frame like the filter popups,
        // so the popup stands out from the white grid behind it.
        private static void FramePopup(Control panel)
        {
            panel.Paint += delegate(object sender, PaintEventArgs e)
            {
                using (var pen = new Pen(Color.FromArgb(160, 170, 186))) e.Graphics.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1);
            };
            panel.Resize += delegate { panel.Invalidate(); };
        }
        // 齿轮: presets on top, then the pick-button mode as one switch row.
        // Every row is laid out here (not by ToolStrip) so all titles share
        // one left edge and all notes start in one column.
        private void ShowSettings()
        {
            var drop = new ToolStripDropDown { Padding = Padding.Empty, DropShadowEnabled = true, BackColor = Color.White, Renderer = new MigrationMenuRenderer() };
            var font = Ui.MakeFont(10f, FontStyle.Regular, Font);
            int pad = Ui.Px(this, 14), rowHeight = Ui.Px(this, 34), gap = Ui.Px(this, 28), inset = Ui.Px(this, 4);
            var entries = new[] { new[] { "保存为预设…", "把当前整套设置存进预设库" }, new[] { "预设管理…", "查看、应用、导入、导出、删除" } };
            int titleWidth = entries.Max(e => TextRenderer.MeasureText(e[0], font).Width), noteWidth = entries.Max(e => TextRenderer.MeasureText(e[1], font).Width);
            string pickLabel = "行末框选按钮"; int toggleWidth = Ui.Px(this, 200);
            int width = Math.Max(pad + titleWidth + gap + noteWidth + pad, pad + TextRenderer.MeasureText(pickLabel, font).Width + gap + toggleWidth + pad);
            var panel = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, BackColor = Color.White, Font = font, Padding = new Padding(inset, Ui.Px(this, 6), inset, Ui.Px(this, 8)) };
            Action<string> section = delegate(string text)
            { panel.Controls.Add(new Label { Text = text, AutoSize = false, Size = new Size(width, Ui.Px(this, 28)), TextAlign = ContentAlignment.BottomLeft, ForeColor = Ui.InkDim, Padding = new Padding(pad, 0, 0, Ui.Px(this, 3)), Margin = Padding.Empty }); };
            section("预设");
            var actions = new Action[] { delegate { SavePreset(); }, delegate { ShowPresetManager(); } };
            for (int i = 0; i < entries.Length; i++)
            {
                var action = actions[i];
                var row = new MigrationMenuRow { Title = entries[i][0], Note = entries[i][1], TextLeft = pad, NoteLeft = pad + titleWidth + gap, Size = new Size(width, rowHeight), Font = font, Margin = Padding.Empty };
                row.Click += delegate { drop.Close(); BeginInvoke(action); };
                panel.Controls.Add(row);
            }
            panel.Controls.Add(new Panel { Size = new Size(width - pad, 1), BackColor = Ui.GridLine, Margin = new Padding(pad / 2, Ui.Px(this, 6), 0, Ui.Px(this, 2)) });
            section("框选方式");
            var pick = new TableLayoutPanel { Size = new Size(width, rowHeight + Ui.Px(this, 6)), ColumnCount = 2, RowCount = 1, BackColor = Color.White, Margin = Padding.Empty, Padding = new Padding(pad, 0, pad, 0) };
            pick.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); pick.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, toggleWidth));
            pick.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            pick.Controls.Add(new Label { Text = pickLabel, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = Ui.Ink, Margin = Padding.Empty, Padding = Padding.Empty }, 0, 0);
            var toggle = new MigrationSegmented("打开工作簿", "右侧栏", PickInWorkbook ? 0 : 1) { Name = "migrationPickMode", Size = new Size(toggleWidth, Ui.Px(this, 30)), Font = font, Anchor = AnchorStyles.Right, Margin = Padding.Empty };
            toggle.SelectedChanged += delegate
            {
                PickInWorkbook = toggle.Selected == 0;
                Message(PickInWorkbook ? "框选按钮：打开工作簿，在工作表里框选（只读）。" : "框选按钮：在右侧区域选择栏框选。");
            };
            pick.Controls.Add(toggle, 1, 0);
            panel.Controls.Add(pick);
            FramePopup(panel); panel.PerformLayout();
            drop.Items.Add(new ToolStripControlHost(panel) { AutoSize = false, Size = panel.PreferredSize, Margin = Padding.Empty, Padding = Padding.Empty });
            drop.Closed += delegate { BeginInvoke(new Action(delegate { drop.Dispose(); font.Dispose(); })); };
            drop.Show(gear.PointToScreen(new Point(gear.Width, gear.Height)), ToolStripDropDownDirection.BelowLeft);
            MigrationDropDownWatch.Attach(drop);
        }
        internal void ShowSidebar(bool show)
        {
            split.Panel2Collapsed = !show; edge.Open = show;
            if (show) { if (!splitPlaced && split.Width > 300) { splitPlaced = true; split.SplitterDistance = Math.Max(200, split.Width * 62 / 100); } UpdatePreview(); }
        }

        // ---------------- grids ----------------
        private void SetupSideGrid(ResultGrid g, bool target)
        {
            PrepareGrid(g);
            string role = target ? "目标" : "源";
            AddGridColumn(g, "no", "映射行", 60, true);
            AddGridColumn(g, "book", role + "工作簿", 180, true);
            AddGridColumn(g, "sheet", role + "工作表", 140, true);
            AddGridColumn(g, "range", role + "区域", 120, false);
            AddGridColumn(g, "pick", "", 40, true);
            if (!target)
            {
                AddGridColumn(g, "trigger", "", 40, true);
                g.Columns["trigger"].Resizable = DataGridViewTriState.False; g.Columns["trigger"].ToolTipText = "逐行重算：把一列值逐个填入触发单元格，重算后写入同行文件名对应的目标工作簿";
            }
            AddGridColumn(g, "size", "尺寸", 100, true);
            g.Columns["pick"].Resizable = DataGridViewTriState.False; g.Columns["pick"].ToolTipText = "框选区域（打开方式可在右上角设置中切换）";
            g.Columns["no"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            g.Columns["sheet"].HeaderCell = new MigrationSelectionHeader { Value = role + "工作表", ToolTipText = "点击下拉箭头，把各行都存在的工作表一次应用到全部。" };
            g.ColumnHeaderMouseClick += delegate(object sender, DataGridViewCellMouseEventArgs e)
            { if (e.Button == MouseButtons.Left && g.Columns[e.ColumnIndex].Name == "sheet") ShowSheetMenu(target); };
            // Only the range is typed; workbook and sheet are picked from popups.
            g.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (syncing || e.RowIndex < 0 || e.ColumnIndex < 0 || g.Columns[e.ColumnIndex].Name != "range") return;
                var row = g.Rows[e.RowIndex]; var pair = (MigrationPair)row.Tag; var side = target ? pair.Target : pair.Source;
                side.Range = NormalizeAddress(Convert.ToString(row.Cells["range"].Value));
                pair.Instances.Clear(); WriteSideRow(row, target); UpdateMappingRow(pair); Changed(); UpdatePreview();
            };
            g.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (Running || g.ReadOnly || e.RowIndex < 0 || e.ColumnIndex < 0) return;
                string name = g.Columns[e.ColumnIndex].Name;
                if (name == "book") ShowBookPopup(g, g.Rows[e.RowIndex], target);
                else if (name == "sheet") ShowRowSheetPopup(g, g.Rows[e.RowIndex], target);
                else if (name == "pick") OpenRangePicker(g, g.Rows[e.RowIndex], target);
                else if (name == "trigger") ShowTriggerPopup(g, (MigrationPair)g.Rows[e.RowIndex].Tag, null);
            };
            g.CellMouseEnter += delegate(object sender, DataGridViewCellEventArgs e)
            {
                string name = e.ColumnIndex < 0 ? "" : g.Columns[e.ColumnIndex].Name;
                g.Cursor = e.RowIndex >= 0 && !g.ReadOnly && (name == "book" || name == "sheet" || name == "pick" || name == "trigger") ? Cursors.Hand : Cursors.Default;
                string hot = e.RowIndex >= 0 && (name == "pick" || name == "trigger") ? name + ":" + e.RowIndex : "";
                if (!object.Equals(g.Tag, hot)) { g.Tag = hot; g.Invalidate(); }
            };
            g.CellMouseLeave += delegate { g.Cursor = Cursors.Default; if (!object.Equals(g.Tag, "")) { g.Tag = ""; g.Invalidate(); } };
            g.CellPainting += PaintChoiceCell;
            g.CurrentCellChanged += delegate { UpdatePreview(); ShowPairDetails(); };
            AttachRowMenu(g, target);
            Ui.StyleGrid(g);
            // Vertical rules so each value reads under its own column.
            g.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            g.CellPainting += PaintHeaderRule;
            g.HandleCreated += delegate { g.Columns["sheet"].HeaderCell.Style.Padding = new Padding(Ui.Px(g, 6), 0, Ui.Px(g, 28), 0); g.Columns["pick"].HeaderCell.Style.Padding = new Padding(1); if (g.Columns.Contains("trigger")) g.Columns["trigger"].HeaderCell.Style.Padding = new Padding(1); };
            FitLikeTargets(g, new Dictionary<string, int> { { "book", 34 }, { "sheet", 24 }, { "range", 22 }, { "size", 20 } });
        }
        // Workbook and sheet cells: plain text with a small chevron, like the
        // header filter, instead of the system combo box.
        private static void PaintChoiceCell(object sender, DataGridViewCellPaintingEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            string name = grid.Columns[e.ColumnIndex].Name; if (name != "book" && name != "sheet" && name != "pick" && name != "trigger") return;
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.ContentForeground);
            float u = e.Graphics.DpiX / 96f; var cell = e.CellBounds;
            bool hot = object.Equals(grid.Tag, name + ":" + e.RowIndex);
            if (name == "pick") { PaintPickIcon(e.Graphics, cell, u, hot, !grid.ReadOnly); e.Handled = true; return; }
            if (name == "trigger") { PaintTriggerIcon(e.Graphics, cell, u, hot, Convert.ToString(e.Value) == "on", !grid.ReadOnly); e.Handled = true; return; }
            int arrow = (int)(24 * u), gap = (int)(8 * u);
            string text = Convert.ToString(e.FormattedValue);
            bool placeholder = text.StartsWith("（", StringComparison.Ordinal), selected = (e.State & DataGridViewElementStates.Selected) != 0;
            Color ink = placeholder ? Ui.InkDim : selected ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor;
            TextRenderer.DrawText(e.Graphics, text, e.CellStyle.Font, new Rectangle(cell.Left + gap, cell.Top, Math.Max(0, cell.Width - arrow - gap), cell.Height), ink,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            var smoothing = e.Graphics.SmoothingMode; e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = cell.Right - arrow / 2f, cy = cell.Top + cell.Height / 2f;
            using (var pen = new Pen(Color.FromArgb(140, 150, 168), 1.4f * u) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                e.Graphics.DrawLines(pen, new[] { new PointF(cx - 3.5f * u, cy - 1.5f * u), new PointF(cx, cy + 2f * u), new PointF(cx + 3.5f * u, cy - 1.5f * u) });
            e.Graphics.SmoothingMode = smoothing; e.Handled = true;
        }
        private static void PaintHeaderRule(object sender, DataGridViewCellPaintingEventArgs e)
        {
            var grid = (DataGridView)sender;
            if (e.RowIndex != -1 || e.ColumnIndex < 0 || e.ColumnIndex >= grid.Columns.Count - 1) return;
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All);
            int inset = Math.Max(1, e.CellBounds.Height / 4);
            using (var pen = new Pen(Ui.GridFrame)) e.Graphics.DrawLine(pen, e.CellBounds.Right - 1, e.CellBounds.Top + inset, e.CellBounds.Right - 1, e.CellBounds.Bottom - inset);
            e.Handled = true;
        }
        // A small sheet with a selected block in blue: "pick the range here".
        private static void PaintPickIcon(Graphics g, Rectangle cell, float u, bool hot, bool enabled)
        {
            var smoothing = g.SmoothingMode; g.SmoothingMode = SmoothingMode.AntiAlias;
            float size = 26 * u, x = cell.Left + (cell.Width - size) / 2f, y = cell.Top + (cell.Height - size) / 2f;
            if (hot && enabled)
                using (var path = Ui.RoundRect(new RectangleF(x, y, size, size), 6 * u)) using (var brush = new SolidBrush(Color.FromArgb(232, 240, 254))) g.FillPath(brush, path);
            float s = 16 * u, ix = x + (size - s) / 2f, iy = y + (size - s) / 2f, step = s / 3f;
            Color ink = !enabled ? Color.FromArgb(200, 206, 216) : hot ? Ui.Blue : Color.FromArgb(120, 132, 152);
            Color accent = !enabled ? Color.FromArgb(200, 206, 216) : Ui.Blue;
            using (var pen = new Pen(ink, 1.2f * u))
            {
                using (var frame = Ui.RoundRect(new RectangleF(ix, iy, s, s), 2.5f * u)) g.DrawPath(pen, frame);
                for (int i = 1; i < 3; i++) { g.DrawLine(pen, ix + i * step, iy, ix + i * step, iy + s); g.DrawLine(pen, ix, iy + i * step, ix + s, iy + i * step); }
            }
            var block = new RectangleF(ix + step, iy + step, step * 2, step * 2);
            using (var fill = new SolidBrush(Color.FromArgb(enabled ? 70 : 40, accent))) g.FillRectangle(fill, block);
            using (var pen = new Pen(accent, 1.6f * u)) g.DrawRectangle(pen, block.X, block.Y, block.Width, block.Height);
            // Fill handle, like a selection in a worksheet.
            float handle = 3.2f * u;
            using (var brush = new SolidBrush(accent)) g.FillRectangle(brush, block.Right - handle / 2, block.Bottom - handle / 2, handle, handle);
            g.SmoothingMode = smoothing;
        }
        // The sheet of one range: same searchable popup as 统一设置.
        private void ShowRowSheetPopup(DataGridView grid, DataGridViewRow row, bool target)
        {
            if (Running || grid.ReadOnly) return;
            var pair = (MigrationPair)row.Tag; var side = BooksSide(target ? pair.Target : pair.Source, target);
            if (side.Books.Count == 0) { Message("请先为这一行选择" + (target ? "目标" : "源") + "工作簿。"); ShowBookPopup(grid, row, target); return; }
            bool wasOpen = sheetPopup != null && sheetPopup.Visible; ClosePopups(); if (wasOpen) return;
            var common = CommonSheets(side, target); var commonSet = new HashSet<string>(common, StringComparer.OrdinalIgnoreCase);
            var partial = side.Books.Select(path => FindBook(target, path)).Where(b => b != null)
                .SelectMany(b => b.Sheets.Distinct(StringComparer.OrdinalIgnoreCase)).Where(n => !commonSet.Contains(n))
                .GroupBy(n => n, StringComparer.OrdinalIgnoreCase).Select(x => new KeyValuePair<string, int>(x.Key, x.Count())).OrderByDescending(x => x.Value).ToList();
            sheetPopup = new MigrationSheetPopup(grid, common, partial, side.Books.Count, false, side.Sheet, delegate(string sheet)
            {
                if (Running || !pairs.Contains(pair)) return;
                var now = target ? pair.Target : pair.Source;
                if (string.Equals(now.Sheet, sheet, StringComparison.Ordinal)) return;
                int moved = ApplySheet(pair, target, sheet);
                RebuildViews(); Changed();
                if (moved > 0) Message("已应用工作表“" + sheet + "”；其中 " + moved + " 个工作簿没有这张工作表，已拆到映射第 " + (pairs.IndexOf(pair) + 2) + " 行，保持原工作表。");
            }, "工作簿") { AllowPartial = true };
            var cell = grid.GetCellDisplayRectangle(grid.Columns["sheet"].Index, row.Index, true);
            sheetPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(sheetPopup);
        }
        // The pick button of a row: open its (first) workbook read-only in
        // the right-hand panel on a sheet all its workbooks share.
        private void OpenRangePicker(DataGridView grid, DataGridViewRow row, bool target)
        {
            if (Running) return;
            var pair = (MigrationPair)row.Tag; var side = BooksSide(target ? pair.Target : pair.Source, target);
            if (side.Books.Count == 0) { Message("请先为这一行选择" + (target ? "目标" : "源") + "工作簿。"); ShowBookPopup(grid, row, target); return; }
            if (side.Sheet.Length == 0)
            {
                var first = FindBook(target, side.Books[0]);
                string sheet = CommonSheets(side, target).FirstOrDefault() ?? (first == null ? null : first.Sheets.FirstOrDefault());
                if (sheet == null) { Message("读取不到这一行工作簿的工作表：" + (first == null || first.Error.Length == 0 ? "没有工作表" : first.Error)); return; }
                int moved = ApplySheet(pair, target, sheet);
                RebuildViews(); Changed();
                if (moved > 0) Message("这些工作簿没有共同的工作表：已用第一个工作簿的“" + sheet + "”，其余 " + moved + " 个拆到下一行。");
            }
            SelectRow(grid, pair);
            if (PickInWorkbook && RangePickRequested != null)
            {
                string path = side.Books[0], sheetName = side.Sheet;
                Message("请在打开的“" + Path.GetFileName(path) + "”中框选" + (target ? "目标" : "源") + "区域，然后点确定。");
                RangePickRequested(target, path, sheetName, side.Range, delegate(string address)
                {
                    MigrationArea area;
                    try { area = MigrationArea.Parse(address); } catch (Exception error) { Message("框选的区域无效：" + error.Message); return; }
                    ApplyPicked(pair, target, area, false); SelectRow(grid, pair);
                });
                return;
            }
            syncingPreview = true; try { previewBook.Items.Clear(); } finally { syncingPreview = false; }
            if (split.Panel2Collapsed) ShowSidebar(true); else UpdatePreview();
            ShowPairDetails();
        }
        private void SetupMappingBoard()
        {
            var b = Mappings;
            b.EmptyText = "还没有映射。先在“源区域”“目标区域”页新增区域，这里会按行把它们对应起来。";
            b.CurrentChanged += delegate
            {
                mappingFocus = b.Current >= 0 && b.Current < pairs.Count ? pairs[b.Current] : null;
                if (b.Current >= 0) { if (b.CurrentTarget) showTarget.Checked = true; else showSource.Checked = true; }
                UpdatePreview(); ShowPairDetails();
            };
            b.EnabledToggled += delegate(int row)
            {
                if (Running || row < 0 || row >= pairs.Count) return;
                var pair = pairs[row]; pair.Enabled = !pair.Enabled; pair.Instances.Clear(); UpdateMappingRow(pair); Changed(); ShowPairDetails();
            };
            b.ModeToggled += delegate(int row)
            {
                if (Running || row < 0 || row >= pairs.Count) return;
                var pair = pairs[row];
                if (pair.Source.Trigger != null) { Message("逐行重算固定只写数值（写入重算后的结果）。"); return; }
                pair.ValuesOnly = !pair.ValuesOnly; pair.Instances.Clear(); UpdateMappingRow(pair); Changed(); ShowPairDetails();
                Message("映射第 " + (row + 1) + " 行改为“" + (pair.ValuesOnly ? ModeValues : ModeFormula) + "”。");
            };
            b.SwapRequested += delegate(int from, int to, bool target) { if (!Running) SwapSides(from, to, target); };
            b.SideActivated += delegate(int row, bool target)
            {
                if (row < 0 || row >= pairs.Count) return;
                var pair = pairs[row]; if (!(target ? pair.Target : pair.Source).Used) return;
                ShowPage(target ? 1 : 0); SelectRow(target ? Targets : Sources, pair);
            };
            b.DeleteRequested += delegate { RemoveMappingRows(); };
            b.ContextMenuStrip = Menu(b, new[] { "同一源区域再加一个映射", "编辑源区域", "编辑目标区域", "删除这一行映射" }, new string[] { null, null, null, "Delete" }, new Action[]
            {
                delegate { AddMappingRow(); },
                delegate { if (b.Current >= 0 && pairs[b.Current].Source.Used) { ShowPage(0); SelectRow(Sources, pairs[b.Current]); } },
                delegate { if (b.Current >= 0 && pairs[b.Current].Target.Used) { ShowPage(1); SelectRow(Targets, pairs[b.Current]); } },
                RemoveMappingRows
            }, delegate { return !Running && !b.ReadOnly && b.Current >= 0; });
        }
        // Right-click menu in the window's style; the last item is the
        // destructive one and sits below a separator.
        private static ContextMenuStrip Menu(Control owner, string[] texts, string[] keys, Action[] actions, Func<bool> allowed)
        {
            var menu = new ContextMenuStrip { ShowImageMargin = false, Renderer = new MigrationMenuRenderer(), BackColor = Color.White, Font = Ui.MakeFont(10f, FontStyle.Regular, owner.Font) };
            for (int i = 0; i < texts.Length; i++)
            {
                if (i > 0 && i == texts.Length - 1) menu.Items.Add(new ToolStripSeparator());
                var action = actions[i];
                var item = new ToolStripMenuItem(texts[i]) { ShortcutKeyDisplayString = keys[i] ?? "", ShowShortcutKeys = keys[i] != null };
                if (i == texts.Length - 1) item.ForeColor = Ui.Danger;
                item.Click += delegate { action(); };
                menu.Items.Add(item);
            }
            menu.Opening += delegate(object sender, System.ComponentModel.CancelEventArgs e)
            {
                e.Cancel = !allowed(); if (e.Cancel) return;
                int h = Ui.Px(owner, 10), v = Ui.Px(owner, 5);
                menu.Padding = new Padding(Ui.Px(owner, 4));
                menu.MinimumSize = new Size(Ui.Px(owner, 180), 0);
                foreach (ToolStripItem item in menu.Items) if (item is ToolStripMenuItem) item.Padding = new Padding(h, v, h, v);
            };
            MigrationDropDownWatch.Attach(menu);
            owner.Disposed += delegate { menu.Dispose(); };
            return menu;
        }
        private void AttachRowMenu(DataGridView grid, bool target)
        {
            string role = target ? "目标" : "源";
            grid.ContextMenuStrip = Menu(grid, new[] { "新增" + role + "区域", "复制此行", "删除选中的行" }, new string[] { null, null, "Delete" },
                new Action[] { delegate { AddSideFromUi(target); }, delegate { CopySide(target); }, delegate { RemoveSides(target); } },
                delegate { return !Running && !grid.ReadOnly; });
            grid.CellMouseDown += delegate(object sender, DataGridViewCellMouseEventArgs e)
            {
                if (e.Button != MouseButtons.Right || e.RowIndex < 0 || grid.Rows[e.RowIndex].Selected) return;
                grid.ClearSelection(); grid.Rows[e.RowIndex].Selected = true;
                grid.CurrentCell = grid.Rows[e.RowIndex].Cells[Math.Max(0, e.ColumnIndex)];
            };
            grid.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.KeyCode != Keys.Delete || grid.IsCurrentCellInEditMode || Running || grid.ReadOnly) return;
                e.Handled = true; RemoveSides(target);
            };
        }
        internal List<MigrationPair> PairItems() { return pairs.ToList(); }
        // Rows created in the window paste values by default; 带公式 is a
        // per-row choice in 映射.
        private static MigrationPair NewPair() { return new MigrationPair { ValuesOnly = true }; }
        // A source range can feed several mapping rows (one source, many
        // targets): those rows share the same MigrationSide object. The
        // 源区域/目标区域 lists show each range once.
        private static MigrationSide SideOf(MigrationPair pair, bool target) { return target ? pair.Target : pair.Source; }
        private List<MigrationSide> DistinctSides(bool target)
        {
            var result = new List<MigrationSide>();
            foreach (var pair in pairs) { var side = SideOf(pair, target); if (side.Used && !result.Contains(side)) result.Add(side); }
            return result;
        }
        private List<int> RowsUsing(MigrationSide side, bool target)
        { var result = new List<int>(); for (int i = 0; i < pairs.Count; i++) if (SideOf(pairs[i], target) == side) result.Add(i); return result; }
        private IEnumerable<MigrationMapping> AllInstances() { return pairs.SelectMany(p => p.Instances); }
        private void EndEdits() { Sources.EndEdit(); Targets.EndEdit(); }
        private List<MigrationBook> Pool(bool target) { return target ? TargetBooks : SourceBooks; }
        private MigrationBook FindBook(bool target, string path) { return Pool(target).FirstOrDefault(b => string.Equals(b.Path, path, StringComparison.OrdinalIgnoreCase)); }
        private DataGridViewRow RowOf(DataGridView grid, MigrationPair pair)
        {
            bool target = grid == Targets; var side = SideOf(pair, target);
            return grid.Rows.Cast<DataGridViewRow>().FirstOrDefault(r => r.Tag == pair || (r.Tag is MigrationPair && SideOf((MigrationPair)r.Tag, target) == side));
        }
        private void SelectRow(DataGridView grid, MigrationPair pair)
        {
            var row = RowOf(grid, pair); if (row == null) return;
            try { grid.ClearSelection(); row.Selected = true; grid.CurrentCell = row.Cells["range"]; } catch (InvalidOperationException) { }
        }
        // Sheet names offered for one side: all sheets of its workbooks.
        private List<string> SheetChoices(MigrationSide side, bool target)
        {
            var result = new List<string>();
            foreach (string path in side.Books) { var book = FindBook(target, path); if (book != null) foreach (string name in book.Sheets) if (!result.Contains(name, StringComparer.OrdinalIgnoreCase)) result.Add(name); }
            if (side.Sheet.Length > 0 && !result.Contains(side.Sheet, StringComparer.OrdinalIgnoreCase)) result.Add(side.Sheet);
            return result;
        }
        // Sheets every workbook of the side has.
        private List<string> CommonSheets(MigrationSide side, bool target)
        {
            List<string> common = null;
            foreach (string path in side.Books)
            {
                var book = FindBook(target, path); if (book == null) continue;
                common = common == null ? book.Sheets.ToList() : common.Where(n => book.Sheets.Contains(n, StringComparer.OrdinalIgnoreCase)).ToList();
            }
            return common ?? new List<string>();
        }
        private static string BooksText(List<string> books)
        {
            if (books.Count == 0) return "（选择工作簿）";
            string first = Path.GetFileName(books[0]);
            return books.Count == 1 ? first : first + " 等 " + books.Count + " 个";
        }

        // ---------------- views (pairs -> rows) ----------------
        private void WriteSideRow(DataGridViewRow row, bool target)
        {
            var pair = (MigrationPair)row.Tag; var side = target ? pair.Target : pair.Source; bool was = syncing; syncing = true;
            try
            {
                row.Cells["no"].Value = string.Join("、", RowsUsing(side, target).Select(i => (i + 1).ToString(CultureInfo.InvariantCulture)).ToArray());
                bool byName = target && IsByName(side);
                row.Cells["book"].Value = byName ? "按文件名匹配（已添加 " + TargetBooks.Count + " 个）" : BooksText(side.Books);
                if (!target) row.Cells["trigger"].Value = side.Trigger != null ? "on" : "";
                row.Cells["book"].ToolTipText = side.Books.Count == 0 ? "点击选择" + (target ? "目标" : "源") + "工作簿（可多选）" : string.Join(Environment.NewLine, side.Books.ToArray());
                row.Cells["sheet"].Value = side.Sheet.Length == 0 ? "（选择工作表）" : side.Sheet;
                row.Cells["sheet"].ToolTipText = side.Books.Count == 0 ? "请先选择工作簿" : "点击选择" + (target ? "目标" : "源") + "工作表";
                if (Convert.ToString(row.Cells["range"].Value) != side.Range) row.Cells["range"].Value = side.Range;
                var area = target ? TargetArea(pair) : TryParse(side.Range);
                row.Cells["size"].Value = side.Range.Length == 0 ? "" : area == null ? "区域无效" : area.Rows + " 行 × " + area.Columns + " 列" + (target && area.Address != side.Range && side.Range.IndexOf(':') < 0 ? "（按源）" : "");
            }
            finally { syncing = was; }
        }
        // What the 映射 board shows for one mapping row.
        private MigrationMappingBoard.Entry EntryOf(MigrationPair pair)
        {
            var source = TryParse(pair.Source.Range); var target = TargetArea(pair);
            var entry = new MigrationMappingBoard.Entry
            {
                SourceUsed = pair.Source.Used, TargetUsed = pair.Target.Used, Enabled = pair.Enabled, ValuesOnly = pair.ValuesOnly || pair.Source.Trigger != null,
                SourceTitle = CardTitle(pair.Source, source), SourceDetail = CardDetail(pair.Source, source) + Shared(pair.Source, false) + TriggerTag(pair.Source),
                TargetTitle = CardTitle(pair.Target, target),
                TargetDetail = pair.Source.Trigger != null ? "按文件名匹配（已添加 " + TargetBooks.Count + " 个）" + (target == null ? "" : " · " + target.Rows + "×" + target.Columns) : CardDetail(pair.Target, target),
                SourceTip = string.Join(Environment.NewLine, pair.Source.Books.ToArray()),
                TargetTip = string.Join(Environment.NewLine, pair.Target.Books.ToArray()),
                Status = PairStatus(pair)
            };
            string errors = PairErrors(pair); entry.StatusTip = errors.Length > 0 ? errors : entry.Status;
            string s = entry.Status;
            entry.Tone = !pair.Enabled ? 0 : s.StartsWith("逐行", StringComparison.Ordinal) ? (s.Contains("模糊") || s.Contains("候选") ? 2 : 4)
                : s.StartsWith("成功", StringComparison.Ordinal) ? 1 : s.StartsWith("部分成功", StringComparison.Ordinal) ? 2
                : s.StartsWith("可执行", StringComparison.Ordinal) || s.StartsWith("就绪", StringComparison.Ordinal) ? 4
                : s.StartsWith("未配对", StringComparison.Ordinal) || s.StartsWith("已还原", StringComparison.Ordinal) || s.StartsWith("待", StringComparison.Ordinal) ? 0 : 3;
            return entry;
        }
        private string Shared(MigrationSide side, bool target)
        {
            if (!side.Used) return "";
            var rows = RowsUsing(side, target);
            return rows.Count < 2 ? "" : " · 同源 " + string.Join("、", rows.Select(i => (i + 1).ToString(CultureInfo.InvariantCulture)).ToArray()) + " 行";
        }
        private static string CardTitle(MigrationSide side, MigrationArea area)
        { return (side.Sheet.Length == 0 ? "未选工作表" : side.Sheet) + "!" + (area != null ? area.Address : side.Range.Length == 0 ? "未设区域" : side.Range); }
        private static string CardDetail(MigrationSide side, MigrationArea area)
        { return (side.Books.Count == 0 ? "未选工作簿" : BooksText(side.Books)) + (area == null ? "" : " · " + area.Rows + "×" + area.Columns); }
        private void UpdateMappingRow(MigrationPair pair)
        {
            for (int i = 0; i < pairs.Count && i < Mappings.Count; i++)
            {
                var other = pairs[i];
                if (other == pair || (pair.Source.Used && other.Source == pair.Source) || (pair.Target.Used && other.Target == pair.Target)) Mappings.UpdateItem(i, EntryOf(other));
            }
        }
        // Rows of the lists in pair order; keeps the current row if it can.
        private void RebuildViews()
        {
            EndEdits();
            var keep = new Dictionary<DataGridView, MigrationPair>();
            foreach (var grid in new DataGridView[] { Sources, Targets }) keep[grid] = grid.CurrentRow == null ? null : grid.CurrentRow.Tag as MigrationPair;
            syncing = true;
            try
            {
                foreach (var grid in new DataGridView[] { Sources, Targets })
                {
                    grid.Rows.Clear();
                    bool target = grid == Targets; var seen = new List<MigrationSide>();
                    foreach (var pair in pairs)
                    {
                        var side = SideOf(pair, target); if (!side.Used || seen.Contains(side)) continue;
                        seen.Add(side); int index = grid.Rows.Add(); grid.Rows[index].Tag = pair;
                    }
                }
            }
            finally { syncing = false; }
            int focus = mappingFocus == null ? -1 : pairs.IndexOf(mappingFocus);
            if (focus < 0) focus = Math.Min(Math.Max(0, Mappings.Current), pairs.Count - 1);
            Mappings.SetItems(pairs.Select(EntryOf).ToList(), focus, Mappings.CurrentTarget);
            mappingFocus = Mappings.Current >= 0 ? pairs[Mappings.Current] : null;
            RefreshRows();
            foreach (var pair in keep) if (pair.Value != null) SelectRow(pair.Key, pair.Value);
            UpdateTabTitles(); Ui.AutoFitColumns(Sources); Ui.AutoFitColumns(Targets);
            UpdatePreview(); ShowPairDetails();
        }
        private void RefreshRows()
        {
            foreach (DataGridViewRow row in Sources.Rows) WriteSideRow(row, false);
            foreach (DataGridViewRow row in Targets.Rows) WriteSideRow(row, true);
            for (int i = 0; i < pairs.Count && i < Mappings.Count; i++) Mappings.UpdateItem(i, EntryOf(pairs[i]));
        }
        private static MigrationArea TryParse(string text) { try { return MigrationArea.Parse(text); } catch { return null; } }
        // Target area of a row: a single start cell takes the source's size.
        private static MigrationArea TargetArea(MigrationPair pair)
        {
            var target = TryParse(pair.Target.Range); if (target == null) return null;
            var source = TryParse(pair.Source.Range);
            if (target.Rows == 1 && target.Columns == 1 && source != null) return new MigrationArea { Row = target.Row, Column = target.Column, Rows = source.Rows, Columns = source.Columns };
            return target;
        }
        // Why a mapping row cannot run yet ("" when it can be prechecked).
        private string RowProblem(MigrationPair pair)
        {
            if (!pair.Source.Used || !pair.Target.Used) return "未配对";
            if (pair.Source.Books.Count == 0) return "未选源工作簿";
            if (pair.Source.Sheet.Length == 0) return "未选源工作表";
            if (pair.Source.Range.Length == 0) return "未设源区域";
            var source = TryParse(pair.Source.Range); if (source == null) return "源区域无效：" + pair.Source.Range;
            bool byName = pair.Source.Trigger != null;
            if (byName && pair.Source.Books.Count != 1) return "逐行重算只支持一个源工作簿";
            if (byName && TargetBooks.Count == 0) return "请先添加目标工作簿（按文件名匹配）";
            if (!byName && pair.Target.Books.Count == 0) return "未选目标工作簿";
            if (pair.Target.Sheet.Length == 0) return "未选目标工作表";
            if (pair.Target.Range.Length == 0) return "未设目标区域";
            var target = TargetArea(pair); if (target == null) return "目标区域无效：" + pair.Target.Range;
            if (target.Rows != source.Rows || target.Columns != source.Columns) return "形状不一致：源 " + source.Rows + "×" + source.Columns + "，目标 " + target.Rows + "×" + target.Columns;
            if (!byName && pair.Source.Books.Count > 1 && pair.Source.Books.Count != pair.Target.Books.Count) return "源工作簿 " + pair.Source.Books.Count + " 个与目标 " + pair.Target.Books.Count + " 个无法一一对应";
            return "";
        }
        private string PairStatus(MigrationPair pair)
        {
            string problem = RowProblem(pair);
            if (!pair.Enabled) return "未启用" + (problem.Length == 0 ? "" : "；" + problem);
            if (pair.Instances.Count == 0 || pair.Instances.All(m => m.Result == null))
                return problem.Length > 0 ? problem : pair.Source.Trigger != null ? TriggerSummary(pair.Source) : "就绪（待预检）";
            var results = pair.Instances.Where(m => m.Result != null).Select(m => m.Result).ToList();
            int total = results.Count, ok = results.Count(t => t.Status == "成功"), ready = results.Count(t => t.Ready && t.Status == "可执行"), restored = results.Count(t => t.Status == "已还原");
            int bad = results.Count(t => t.Status == "失败" || t.Status == "冲突" || t.Status == "已取消");
            string head = ok > 0 ? (ok == total ? "成功 " : "部分成功 ") + ok + "/" + total
                : restored > 0 ? "已还原 " + restored + "/" + total
                : ready > 0 ? "可执行 " + ready + "/" + total
                : bad > 0 ? results.First(t => t.Status == "失败" || t.Status == "冲突" || t.Status == "已取消").Status + " " + bad + "/" + total
                : results[0].Status;
            var first = results.FirstOrDefault(t => t.Error.Length > 0 && t.Status != "成功");
            return head + (first == null ? "" : "；" + first.Error);
        }
        private static string PairErrors(MigrationPair pair)
        { return string.Join(Environment.NewLine, pair.Instances.Where(m => m.Result != null && m.Result.Error.Length > 0).Select(m => Path.GetFileName(m.TargetPath) + "：" + m.Result.Error).Take(12).ToArray()); }

        // ---------------- editing the lists ----------------
        // Put a new source/target range into the first row that lacks that
        // side, else a new row at the end. Filling gaps instead of inserting
        // keeps every existing pairing in place.
        internal MigrationPair AddSide(bool target, MigrationSide side)
        {
            side.Used = true;
            var pair = pairs.FirstOrDefault(p => !(target ? p.Target : p.Source).Used);
            if (pair == null) { pair = NewPair(); pairs.Add(pair); }
            if (target) pair.Target = side; else pair.Source = side;
            pair.Instances.Clear(); RebuildViews(); Changed();
            return pair;
        }
        // A complete mapping row (both sides present), e.g. from a saved rule.
        internal MigrationPair AddPair(MigrationPair pair)
        {
            pair.Source.Used = pair.Target.Used = true;
            pairs.Add(pair); RebuildViews(); Changed(); return pair;
        }
        // A new range inherits the workbooks and sheet of the previous one on
        // that side, so a run of ranges in the same files only needs the area.
        private void AddSideFromUi(bool target)
        {
            if (Running) return;
            var grid = target ? Targets : Sources;
            var side = new MigrationSide();
            var pair = AddSide(target, side);
            SelectRow(grid, pair);
            Message("已新增" + (target ? "目标" : "源") + "区域（映射第 " + (pairs.IndexOf(pair) + 1) + " 行）；点“" + (target ? "目标" : "源") + "工作簿”一格选择工作簿。");
        }
        private void CopySide(bool target)
        {
            if (Running) return;
            var grid = target ? Targets : Sources;
            if (grid.CurrentRow == null) { Message("请先选择要复制的行。"); return; }
            var current = (MigrationPair)grid.CurrentRow.Tag;
            var pair = AddSide(target, (target ? current.Target : current.Source).Copy());
            SelectRow(grid, pair);
            Message("已复制为映射第 " + (pairs.IndexOf(pair) + 1) + " 行的" + (target ? "目标" : "源") + "区域。");
        }
        // Removing a range leaves a gap in its mapping row (shown as 未配对),
        // so the other rows keep their pairing; an empty row disappears.
        private void RemoveSides(bool target)
        {
            if (Running) return;
            var grid = target ? Targets : Sources;
            var chosen = grid.SelectedRows.Cast<DataGridViewRow>().Select(r => r.Tag as MigrationPair).Where(p => p != null).Select(p => SideOf(p, target)).Distinct().ToList();
            if (chosen.Count == 0 && grid.CurrentRow != null) chosen.Add(SideOf((MigrationPair)grid.CurrentRow.Tag, target));
            if (chosen.Count == 0) { Message("请先选择要删除的行。"); return; }
            foreach (var pair in pairs.ToList())
            {
                if (!chosen.Contains(SideOf(pair, target))) continue;
                if (target) pair.Target = new MigrationSide(); else pair.Source = new MigrationSide();
                pair.Instances.Clear();
                if (!pair.Source.Used && !pair.Target.Used) pairs.Remove(pair);
            }
            RebuildViews(); Changed();
            Message("已删除 " + chosen.Count + " 个" + (target ? "目标" : "源") + "区域；对应映射行显示为“未配对”，可在“映射”页拖动重新对应。");
        }
        // A new mapping row for the selected row's source range (the same
        // range, shared), so one source can be written to several targets.
        // The next target range added fills this row.
        internal MigrationPair AddMappingRow()
        {
            if (Running) return null;
            var current = Mappings.Current >= 0 && Mappings.Current < pairs.Count ? pairs[Mappings.Current] : pairs.LastOrDefault(p => p.Source.Used);
            var pair = NewPair();
            if (current != null && current.Source.Used) { pair.Source = current.Source; pair.Enabled = current.Enabled; pair.ValuesOnly = current.ValuesOnly; }
            int at = current == null ? pairs.Count : (current.Source.Used ? pairs.FindLastIndex(p => p.Source == current.Source) : pairs.IndexOf(current)) + 1;
            pairs.Insert(at, pair);
            mappingFocus = pair;
            RebuildViews(); Changed(); Mappings.SetCurrent(pairs.IndexOf(pair), true);
            Message(pair.Source.Used
                ? "已新增映射第 " + (pairs.IndexOf(pair) + 1) + " 行，与第 " + (pairs.IndexOf(current) + 1) + " 行使用同一源区域；到“目标区域”页新增目标，或把其他行的目标拖到这一行。"
                : "已新增映射第 " + (pairs.IndexOf(pair) + 1) + " 行；新增的源区域和目标区域会依次补进这一行。");
            return pair;
        }
        private void RemoveMappingRows()
        {
            if (Running) return;
            if (Mappings.Current < 0 || Mappings.Current >= pairs.Count) { Message("请先选择要删除的映射行。"); return; }
            int number = Mappings.Current + 1;
            pairs.RemoveAt(Mappings.Current); mappingFocus = null;
            RebuildViews(); Changed();
            Message("已删除映射第 " + number + " 行（连同其中的源区域和目标区域）；工作簿文件不受影响。");
        }
        internal void SwapSides(int a, int b, bool target)
        {
            if (a < 0 || b < 0 || a >= pairs.Count || b >= pairs.Count || a == b) return;
            var first = pairs[a]; var second = pairs[b];
            if (target) { var side = first.Target; first.Target = second.Target; second.Target = side; }
            else { var side = first.Source; first.Source = second.Source; second.Source = side; }
            first.Instances.Clear(); second.Instances.Clear();
            // The moved card stays selected at its new row.
            mappingFocus = second; Mappings.SetCurrent(b, target);
            RebuildViews(); Changed();
            Message("已交换映射第 " + (a + 1) + " 行与第 " + (b + 1) + " 行的" + (target ? "目标" : "源") + "区域。");
        }

        // ---------------- workbooks ----------------
        // forPair: the row whose workbook popup asked for more workbooks; the
        // added workbooks are selected for that row.
        private void PickBooks(bool target, MigrationPair forPair)
        {
            try
            {
                string[] selections = MigrationTargetPicker.Show(this); if (selections.Length == 0) return;
                var scan = MigrationTargetScan.Expand(selections, "");
                if (scan.Files.Count > 0 && BooksPicked != null) BooksPicked(target, scan.Files.ToArray());
                if (forPair == null)
                {
                    int added = AddSideRows(target, scan.Files); string role = target ? "目标" : "源";
                    Message("已添加 " + scan.Files.Count + " 个" + role + "工作簿，新增 " + added + " 个" + role + "区域（每个工作簿一行）；点每行末的框选按钮选择区域。" + string.Join("；", scan.Errors.ToArray()));
                    return;
                }
                if (forPair != null && pairs.Contains(forPair))
                {
                    var side = target ? forPair.Target : forPair.Source;
                    foreach (string file in scan.Files) { string path = Path.GetFullPath(file); if (!side.Books.Contains(path, StringComparer.OrdinalIgnoreCase)) side.Books.Add(path); }
                    forPair.Instances.Clear(); RefreshRows(); Changed(); UpdatePreview();
                }
                Message("已添加 " + scan.Files.Count + " 个" + (target ? "目标" : "源") + "工作簿" + (forPair != null ? "并选入当前行" : "；新增区域后在“工作簿”一格选择") + "。文件夹只扫描当前层。" + string.Join("；", scan.Errors.ToArray()));
            }
            catch (Exception error) { Message("添加工作簿失败：" + error.Message); }
        }
        // One range per workbook, on the source and the target side alike:
        // each workbook usually has its own area. A workbook that already has
        // a single-workbook range without an area reuses that row. New ranges
        // fill the first mapping rows lacking that side, so pairs line up.
        // Returns how many rows were added.
        internal int AddSideRows(bool target, IEnumerable<string> files)
        {
            int added = 0;
            foreach (string file in files)
            {
                string path = Path.GetFullPath(file);
                if (pairs.Any(p => { var s = target ? p.Target : p.Source; return s.Used && s.Range.Length == 0 && s.Books.Count == 1 && string.Equals(s.Books[0], path, StringComparison.OrdinalIgnoreCase); })) continue;
                var book = FindBook(target, path);
                var side = new MigrationSide { Used = true, Sheet = book != null && book.Sheets.Count > 0 ? book.Sheets[0] : "" }; side.Books.Add(path);
                var pair = pairs.FirstOrDefault(p => !(target ? p.Target : p.Source).Used);
                if (pair == null) { pair = NewPair(); pairs.Add(pair); }
                if (target) pair.Target = side; else pair.Source = side;
                pair.Instances.Clear(); added++;
            }
            if (added > 0) { RebuildViews(); Changed(); ShowPage(target ? 1 : 0); }
            return added;
        }
        // Add or refresh a workbook in the list. Ranges are not changed: each
        // range picks its own workbooks.
        internal MigrationBook AddBook(bool target, string path, IEnumerable<string> sheets, string error = "")
        {
            path = Path.GetFullPath(path); var pool = Pool(target);
            var book = FindBook(target, path);
            if (book == null) { book = new MigrationBook { Path = path }; pool.Add(book); }
            book.Sheets.Clear(); book.Sheets.AddRange(sheets ?? new string[0]); book.Error = error ?? "";
            RefreshRows(); Changed(); UpdatePreview();
            return book;
        }
        // Opening the window: the active workbook becomes the first source range.
        internal void AddDefaultSource(string path, IEnumerable<string> sheets, string sheet)
        {
            AddBook(false, path, sheets);
            if (pairs.Any(p => p.Source.Used)) return;
            var side = new MigrationSide { Sheet = sheet ?? "" }; side.Books.Add(Path.GetFullPath(path));
            var pair = AddSide(false, side);
            SelectRow(Sources, pair);
        }
        internal MigrationBookPopup CreateBookPopup(MigrationPair pair, bool target)
        {
            if (Running) throw new InvalidOperationException("迁移运行期间不能更改工作簿。");
            var side = target ? pair.Target : pair.Source;
            return new MigrationBookPopup(target ? Targets : Sources, Pool(target), side.Books, target ? "目标工作簿" : "源工作簿", delegate(IList<string> chosen)
            {
                if (Running || !pairs.Contains(pair)) return;
                var now = target ? pair.Target : pair.Source;
                if (now.Books.SequenceEqual(chosen, StringComparer.OrdinalIgnoreCase)) return;
                now.Books.Clear(); now.Books.AddRange(chosen);
                pair.Instances.Clear(); RefreshRows(); Changed(); UpdatePreview();
            }, delegate { BeginInvoke(new Action(delegate { PickBooks(target, pair); })); });
        }
        private void ShowBookPopup(DataGridView grid, DataGridViewRow row, bool target)
        {
            if (Running || grid.ReadOnly) return;
            if (target && IsByName(((MigrationPair)row.Tag).Target))
            { Message("这一行与逐行重算的源配对：目标工作簿按触发值同行的文件名，在已添加的 " + TargetBooks.Count + " 个目标工作簿中自动匹配，无需选择。"); return; }
            if (Pool(target).Count == 0) { PickBooks(target, (MigrationPair)row.Tag); return; }
            bool wasOpen = bookPopup != null && bookPopup.Visible; ClosePopups(); if (wasOpen) return;
            bookPopup = CreateBookPopup((MigrationPair)row.Tag, target);
            var cell = grid.GetCellDisplayRectangle(grid.Columns["book"].Index, row.Index, true);
            bookPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(bookPopup);
        }
        private void ClosePopups()
        {
            var books = bookPopup; bookPopup = null; if (books != null) { books.Close(); books.Dispose(); }
            var sheets = sheetPopup; sheetPopup = null; if (sheets != null) { sheets.Close(); sheets.Dispose(); }
            var trigger = triggerPopup; triggerPopup = null; if (trigger != null) { trigger.Close(); trigger.Dispose(); }
        }
        // 统一设置工作表 for one side: every range of that side.
        private void ShowSheetMenu(bool target)
        {
            var grid = target ? Targets : Sources;
            if (Running || grid.ReadOnly || grid.Rows.Count == 0) return;
            bool wasOpen = sheetPopup != null && sheetPopup.Visible; ClosePopups(); if (wasOpen) return;
            EndEdits();
            var rows = pairs.Where(p => (target ? p.Target : p.Source).Used && (target ? p.Target : p.Source).Books.Count > 0).ToList();
            if (rows.Count == 0) { Message("请先为各行选择" + (target ? "目标" : "源") + "工作簿。"); return; }
            // Counted per workbook: a sheet only some workbooks have is still
            // offered; those lacking it are split off into their own rows.
            var books = rows.SelectMany(p => (target ? p.Target : p.Source).Books).ToList();
            var all = new MigrationSide(); all.Books.AddRange(books);
            var common = CommonSheets(all, target); var commonSet = new HashSet<string>(common, StringComparer.OrdinalIgnoreCase);
            var partial = books.Select(path => FindBook(target, path)).Where(b => b != null).SelectMany(b => b.Sheets.Distinct(StringComparer.OrdinalIgnoreCase)).Where(n => !commonSet.Contains(n))
                .GroupBy(n => n, StringComparer.OrdinalIgnoreCase).Select(g => new KeyValuePair<string, int>(g.Key, g.Count())).OrderByDescending(g => g.Value).ToList();
            var values = rows.Select(p => (target ? p.Target : p.Source).Sheet).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            sheetPopup = new MigrationSheetPopup(grid, common, partial, books.Count, false, values.Count == 1 ? values[0] : null, delegate(string sheet)
            {
                if (Running) return;
                int movedBooks = 0, splitRows = 0;
                foreach (var pair in rows) { if (!pairs.Contains(pair)) continue; int moved = ApplySheet(pair, target, sheet); if (moved > 0) { movedBooks += moved; splitRows++; } }
                RebuildViews(); Changed();
                Message("已将 " + rows.Count + " 个" + (target ? "目标" : "源") + "区域的工作表设为“" + sheet + "”。"
                    + (movedBooks > 0 ? "其中 " + movedBooks + " 个工作簿没有这张工作表，已拆成 " + splitRows + " 行单独保留原工作表。" : ""));
            }, "工作簿") { AllowPartial = true };
            var cell = grid.GetCellDisplayRectangle(grid.Columns["sheet"].Index, -1, true);
            sheetPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(sheetPopup);
        }
        // Set the sheet of one range. Workbooks that lack the sheet are split
        // off into a new row right below (same range, old sheet, other side
        // unpaired), so the rest of the rows keep their pairing.
        // Returns how many workbooks were split off.
        internal int ApplySheet(MigrationPair pair, bool target, string sheet)
        {
            var side = target ? pair.Target : pair.Source; string old = side.Sheet;
            if (target && IsByName(side)) { side.Sheet = sheet; pair.Instances.Clear(); return 0; }
            var lacking = side.Books.Where(path => { var book = FindBook(target, path); return book != null && book.Sheets.Count > 0 && !book.Sheets.Contains(sheet, StringComparer.OrdinalIgnoreCase); }).ToList();
            pair.Instances.Clear();
            if (lacking.Count == side.Books.Count) return 0;
            side.Sheet = sheet;
            if (lacking.Count == 0) return 0;
            side.Books.RemoveAll(path => lacking.Contains(path, StringComparer.OrdinalIgnoreCase));
            var rest = side.Copy(); rest.Books.Clear(); rest.Books.AddRange(lacking); rest.Sheet = old;
            var below = new MigrationPair { Enabled = pair.Enabled, ValuesOnly = pair.ValuesOnly };
            if (target) below.Target = rest; else below.Source = rest;
            pairs.Insert(pairs.IndexOf(pair) + 1, below);
            return lacking.Count;
        }
        private void SetGridsReadOnly(bool value)
        {
            Sources.ReadOnly = Targets.ReadOnly = value; Mappings.ReadOnly = value;
            if (value) return;
            foreach (string name in new[] { "no", "book", "sheet", "pick", "trigger", "size" })
                foreach (var grid in new DataGridView[] { Sources, Targets }) if (grid.Columns.Contains(name)) grid.Columns[name].ReadOnly = true;
        }
        private void Changed()
        {
            if (syncing || Running) return;
            lastPlan = null; execute.Enabled = false;
            foreach (var pair in pairs) if (pair.Instances.Any(m => m.Result != null && m.Result.Status != "成功")) { pair.Instances.Clear(); UpdateMappingRow(pair); }
            UpdateTabTitles();
            status.Text = "源区域 " + DistinctSides(false).Count + " 个，目标区域 " + DistinctSides(true).Count + " 个，映射 " + pairs.Count + " 行。设置好后点“预检与预览”。";
            if (InputsChanged != null) InputsChanged();
        }

        // ---------------- range panel ----------------
        // Which row and side the panel is about, by the page in front.
        private MigrationPair PanelPair(out bool target)
        {
            target = false;
            if (CurrentPage == sourcesPage) return Sources.CurrentRow == null ? null : Sources.CurrentRow.Tag as MigrationPair;
            if (CurrentPage == targetsPage) { target = true; return Targets.CurrentRow == null ? null : Targets.CurrentRow.Tag as MigrationPair; }
            target = Mappings.CurrentTarget; return Mappings.Current >= 0 && Mappings.Current < pairs.Count ? pairs[Mappings.Current] : null;
        }
        private void FillPreviewBooks(MigrationSide side)
        {
            syncingPreview = true;
            try
            {
                var selected = previewBook.SelectedItem as BookItem; string keep = selected == null ? null : selected.Path;
                previewBook.Items.Clear();
                if (side != null) foreach (string path in side.Books) previewBook.Items.Add(new BookItem { Path = path });
                int index = -1;
                for (int i = 0; i < previewBook.Items.Count; i++) if (string.Equals(((BookItem)previewBook.Items[i]).Path, keep, StringComparison.OrdinalIgnoreCase)) index = i;
                if (index < 0 && previewBook.Items.Count > 0) index = 0;
                previewBook.SelectedIndex = index;
            }
            finally { syncingPreview = false; }
        }
        private void UpdatePreview()
        {
            if (split == null || split.Panel2Collapsed || tabs == null) return;
            bool target; var pair = PanelPair(out target);
            sideSwitch.Visible = CurrentPage == mappingsPage;
            if (pair == null) { FillPreviewBooks(null); sheetPreview.Clear("在左侧选择一行，这里显示它的工作表，可直接框选区域。"); return; }
            var side = BooksSide(target ? pair.Target : pair.Source, target); string role = target ? "目标" : "源";
            FillPreviewBooks(side);
            sheetPreview.Sync.Visible = !target; sheetPreview.Sync.Text = "作为新的源区域加入"; sheetPreview.OverlapConflict = target;
            if (!side.Used) { sheetPreview.Clear("映射第 " + (pairs.IndexOf(pair) + 1) + " 行还没有" + role + "区域：到“" + role + "区域”页新增，或在本页拖动与其他行对应。"); return; }
            if (side.Books.Count == 0) { sheetPreview.Clear("这一行还没有选择" + role + "工作簿：点左侧“" + role + "工作簿”一格选择。"); return; }
            if (side.Sheet.Length == 0) { sheetPreview.Clear("请先在这一行选择" + role + "工作表。"); return; }
            var item = previewBook.SelectedItem as BookItem; string path = item == null ? side.Books[0] : item.Path;
            if (!File.Exists(path)) { sheetPreview.Clear("找不到" + role + "文件：" + path); return; }
            MigrationSheetSnapshot snapshot;
            try { snapshot = LoadSnapshot(path, side.Sheet); }
            catch (Exception error) { sheetPreview.Clear(error.Message + "（仍可在左侧直接填写" + role + "区域）"); return; }
            var source = TryParse(pair.Source.Range);
            var area = target ? TargetArea(pair) : source;
            var used = new List<KeyValuePair<MigrationArea, string>>();
            foreach (var other in pairs)
            {
                if (other == pair || !other.Enabled) continue;
                var otherSide = target ? other.Target : other.Source; if (!otherSide.Used) continue;
                if (!otherSide.Books.Contains(path, StringComparer.OrdinalIgnoreCase) || !string.Equals(otherSide.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase)) continue;
                var otherArea = target ? TargetArea(other) : TryParse(other.Source.Range);
                if (otherArea != null) used.Add(new KeyValuePair<MigrationArea, string>(otherArea, "映射第 " + (pairs.IndexOf(other) + 1) + " 行 " + otherArea.Address));
            }
            string hint = !target
                ? "拖选要迁移的区域（单击为单个单元格）。浅绿为同一工作表中其他源区域。勾选“作为新的源区域加入”可连续框选多块，每块新增一个源区域。"
                : source != null ? "单击一个单元格作为起点，按源尺寸 " + source.Rows + " 行 × " + source.Columns + " 列放置；也可拖选完整区域。浅绿为其他映射已占用，红色为重叠。"
                : "这一行的源区域尚未设置：直接拖选完整的目标区域即可；单击只记下起点。";
            sheetPreview.ShowSheet(snapshot, area, source == null ? 1 : source.Rows, source == null ? 1 : source.Columns, used,
                role + "：" + Path.GetFileName(path) + " · " + side.Sheet + "　｜　映射第 " + (pairs.IndexOf(pair) + 1) + " 行" + (area == null ? "（未设置区域）" : " " + area.Address), hint);
        }
        // A pick in the panel sets the range of the row/side the panel shows.
        private void PickArea(MigrationArea picked)
        {
            bool target; var pair = PanelPair(out target); if (pair != null) ApplyPicked(pair, target, picked, true);
        }
        // A range chosen in the panel (fromPanel) or in the workbook itself.
        private void ApplyPicked(MigrationPair pair, bool target, MigrationArea picked, bool fromPanel)
        {
            if (Running || !pairs.Contains(pair)) return;
            var side = target ? pair.Target : pair.Source; if (!side.Used) return;
            bool single = picked.Rows == 1 && picked.Columns == 1;
            if (!target)
            {
                string range = single ? MigrationArea.Cell(picked.Row, picked.Column) : picked.Address;
                if (fromPanel && sheetPreview.Sync.Checked && side.Range.Length > 0 && side.Range != range)
                {
                    var copy = side.Copy(); copy.Range = range;
                    var added = AddSide(false, copy); if (CurrentPage == sourcesPage) SelectRow(Sources, added);
                    Message("已新增源区域 " + copy.Sheet + "!" + range + "（映射第 " + (pairs.IndexOf(added) + 1) + " 行）。");
                }
                else
                {
                    side.Range = range; pair.Instances.Clear(); RefreshRows(); Changed();
                    Message("映射第 " + (pairs.IndexOf(pair) + 1) + " 行的源区域已设为 " + side.Sheet + "!" + range + "（" + SizeText(range) + "）。");
                }
                UpdatePreview(); ShowPairDetails(); return;
            }
            // Any target area is accepted: a single cell is a start (sized by
            // the source), a block is kept as chosen. Whether it fits the source
            // is shown in 映射, not enforced here.
            var source = TryParse(pair.Source.Range);
            try { picked.Check(1048576, 16384); }
            catch (Exception error) { Message(error.Message); sheetPreview.SetHint(error.Message); return; }
            side.Range = single ? MigrationArea.Cell(picked.Row, picked.Column) : picked.Address;
            string fit = source == null || single || (picked.Rows == source.Rows && picked.Columns == source.Columns) ? ""
                : "（" + picked.Rows + "×" + picked.Columns + "，与源区域 " + source.Rows + "×" + source.Columns + " 不一致，“映射”页会提示）";
            pair.Instances.Clear(); RefreshRows(); Changed();
            var mine = TargetArea(pair);
            var overlap = pairs.Where(p => p != pair && p.Enabled && p.Target.Used && string.Equals(p.Target.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase)
                    && p.Target.Books.Intersect(side.Books, StringComparer.OrdinalIgnoreCase).Any() && Overlaps(TargetArea(p), mine))
                .Select(p => "映射第 " + (pairs.IndexOf(p) + 1) + " 行").ToList();
            Message("映射第 " + (pairs.IndexOf(pair) + 1) + " 行的目标区域已设为 " + side.Sheet + "!" + side.Range + fit + "，应用到 " + side.Books.Count + " 个目标工作簿。"
                + (overlap.Count > 0 ? "注意：与 " + string.Join("、", overlap.Take(3).ToArray()) + " 重叠，预检时会被拦下。" : ""));
            UpdatePreview(); ShowPairDetails();
        }
        private static string Where(MigrationSide side, MigrationArea area)
        { return !side.Used ? "（未配对）" : (side.Sheet.Length == 0 ? "（未选工作表）" : side.Sheet) + " · " + (area != null ? area.Address : side.Range.Length == 0 ? "（未设区域）" : side.Range + "（无效）"); }
        private static string Books(MigrationSide side)
        { return side.Books.Count == 0 ? "（未选）" : string.Join("；", side.Books.Select(Path.GetFileName).ToArray()); }
        // ---------------- 逐行重算 ----------------
        // A target range paired with a 逐行重算 source takes its workbook from
        // the trigger row's name, matched among all added targets.
        private bool IsByName(MigrationSide side)
        { return side.Used && pairs.Any(p => p.Target == side && p.Source.Used && p.Source.Trigger != null); }
        // The workbooks a side stands for: a by-name target stands for all
        // added targets (for sheet lists, picking and preview).
        private MigrationSide BooksSide(MigrationSide side, bool target)
        {
            if (!target || !IsByName(side)) return side;
            var all = new MigrationSide { Sheet = side.Sheet, Range = side.Range, Used = side.Used }; all.Books.AddRange(TargetBooks.Select(b => b.Path));
            return all;
        }
        private string TriggerTag(MigrationSide side)
        {
            if (side.Trigger == null) return "";
            RefreshTrigger(side);
            return side.Trigger.Error.Length > 0 ? " · 逐行（设置无效）" : " · 逐行 " + side.Trigger.Rows.Count(r => r.Kind != "空值");
        }
        private string TriggerSummary(MigrationSide side)
        {
            var trigger = side.Trigger; RefreshTrigger(side);
            if (trigger.Error.Length > 0) return "逐行重算设置无效：" + trigger.Error;
            var rows = trigger.Rows.Where(r => r.Kind != "空值").ToList();
            int matched = rows.Count(r => r.Ready), fuzzy = rows.Count(r => r.Kind == "模糊"), many = rows.Count(r => r.Kind == "多个候选"), missing = rows.Count(r => r.Kind == "缺失"), skipped = rows.Count(r => r.Kind == "跳过");
            var parts = new List<string> { "匹配 " + matched };
            if (fuzzy > 0) parts.Add("模糊 " + fuzzy); if (many > 0) parts.Add("多个候选 " + many); if (missing > 0) parts.Add("缺 " + missing); if (skipped > 0) parts.Add("跳过 " + skipped);
            int spare = UnusedTargets(side).Count; if (spare > 0) parts.Add("未用到目标 " + spare);
            return "逐行 " + rows.Count + "：" + string.Join("，", parts.ToArray());
        }
        private static int ColumnIndex(string letters)
        {
            if (!Regex.IsMatch(letters ?? "", "^[A-Za-z]{1,3}$")) throw new InvalidOperationException("文件名列应为列字母，例如 H");
            return MigrationArea.Parse(letters.ToUpperInvariant() + "1").Column;
        }
        // Re-read the trigger values and names (from the saved file, without
        // opening WPS) and match each name against the added targets. Cached
        // until the file, settings, targets or choices change.
        private void RefreshTrigger(MigrationSide side)
        {
            var trigger = side.Trigger; if (trigger == null) return;
            string path = side.Books.Count == 1 ? side.Books[0] : "";
            long stamp = path.Length > 0 && File.Exists(path) ? File.GetLastWriteTimeUtc(path).Ticks : 0;
            string signature = path + "|" + side.Sheet + "|" + trigger.Cell + "|" + trigger.Values + "|" + trigger.NameColumn + "|" + stamp + "|"
                + string.Join("|", TargetBooks.Select(b => b.Path).ToArray()) + "|" + string.Join(",", trigger.Chosen.Select(c => c.Key + "=" + c.Value).ToArray());
            if (signature == trigger.Signature) return;
            trigger.Signature = signature; trigger.Rows.Clear(); trigger.Error = "";
            try
            {
                if (side.Books.Count != 1) throw new InvalidOperationException("逐行重算只支持一个源工作簿");
                if (side.Sheet.Length == 0) throw new InvalidOperationException("请先选择源工作表");
                var cell = MigrationArea.Parse(trigger.Cell); if (cell.Rows != 1 || cell.Columns != 1) throw new InvalidOperationException("触发单元格只能是一个单元格");
                var values = MigrationArea.Parse(trigger.Values); if (values.Columns != 1) throw new InvalidOperationException("触发值区域只能是一列");
                int nameColumn = ColumnIndex(trigger.NameColumn);
                var snapshot = LoadSnapshot(path, side.Sheet);
                var pool = TargetBooks.Select(b => b.Path).ToList();
                for (int r = values.Row; r < values.Row + values.Rows; r++)
                {
                    var row = new MigrationTriggerRow { Row = r, Value = snapshot.Value(r, values.Column).Trim(), Name = snapshot.Value(r, nameColumn).Trim() };
                    if (row.Value.Length == 0) row.Kind = "空值";
                    else
                    {
                        MatchRow(row, pool); string chosen;
                        if (trigger.Chosen.TryGetValue(r, out chosen) && (chosen.Length == 0 || pool.Contains(chosen, StringComparer.OrdinalIgnoreCase)))
                        { row.Kind = chosen.Length == 0 ? "跳过" : "已确认"; row.Path = chosen; }
                    }
                    trigger.Rows.Add(row);
                }
            }
            catch (Exception error) { trigger.Error = error.Message; }
        }
        // Name normalisation for matching: no extension, full/half width
        // unified, no "！" or leading numbering such as "00-" / "03-01-", no spaces.
        internal static string NormalizeName(string name)
        {
            string text = (name ?? "").Trim();
            foreach (string extension in new[] { ".xlsx", ".xlsm", ".xls", ".et" })
                if (text.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) { text = text.Substring(0, text.Length - extension.Length); break; }
            text = text.Normalize(NormalizationForm.FormKC);
            text = Regex.Replace(text, @"^[\s!]*(?:\d+[\s\-_.、]+)*", "");
            text = Regex.Replace(text, @"[\s!]", "");
            return text.ToLowerInvariant();
        }
        internal static double NameScore(string a, string b)
        {
            if (a.Length == 0 || b.Length == 0) return 0;
            if (a == b) return 1;
            int min = Math.Min(a.Length, b.Length), max = Math.Max(a.Length, b.Length);
            if (a.Contains(b) || b.Contains(a)) return 0.7 + 0.29 * min / max;
            var previous = new int[b.Length + 1]; var current = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) previous[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                current[0] = i;
                for (int j = 1; j <= b.Length; j++) current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + (a[i - 1] == b[j - 1] ? 0 : 1));
                var swap = previous; previous = current; current = swap;
            }
            return 0.99 * (1 - (double)previous[b.Length] / max);
        }
        // 精确: one candidate equal after normalisation. 模糊: one clearly best
        // close candidate (needs confirmation). 多个候选: no clear winner.
        // 缺失: nothing close enough.
        internal static void MatchRow(MigrationTriggerRow row, IEnumerable<string> pool)
        {
            row.Candidates.Clear(); row.Path = ""; string key = NormalizeName(row.Name);
            foreach (string path in pool)
            {
                double score = NameScore(key, NormalizeName(Path.GetFileName(path)));
                if (score >= 0.7) row.Candidates.Add(new KeyValuePair<string, double>(path, score));
            }
            row.Candidates.Sort((x, y) => y.Value.CompareTo(x.Value));
            int exact = row.Candidates.Count(c => c.Value >= 1);
            if (row.Candidates.Count == 0) row.Kind = "缺失";
            else if (exact == 1) { row.Kind = "精确"; row.Path = row.Candidates[0].Key; }
            else if (exact > 1 || (row.Candidates.Count > 1 && row.Candidates[0].Value - row.Candidates[1].Value < 0.05)) row.Kind = "多个候选";
            else { row.Kind = "模糊"; row.Path = row.Candidates[0].Key; }
        }
        // Set or change 逐行重算 for a source range. Ranges of the same
        // workbook and sheet share the setting (one recalculation feeds all).
        internal void SetTrigger(MigrationSide side, MigrationTrigger trigger)
        {
            int shared = 0;
            foreach (var other in DistinctSides(false))
            {
                if (other == side || other.Books.Count != 1 || side.Books.Count != 1 || !string.Equals(other.Books[0], side.Books[0], StringComparison.OrdinalIgnoreCase) || !string.Equals(other.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase)) continue;
                if (trigger == null && other.Trigger == null) continue;
                other.Trigger = trigger == null ? null : trigger.Copy(); shared++;
            }
            side.Trigger = trigger;
            foreach (var pair in pairs) if (pair.Source.Used && (pair.Source == side || (pair.Source.Books.Count == 1 && side.Books.Count == 1 && string.Equals(pair.Source.Books[0], side.Books[0], StringComparison.OrdinalIgnoreCase)))) pair.Instances.Clear();
            RefreshRows(); Changed(); ShowPairDetails();
            Message(trigger == null ? "已取消逐行重算" + (shared > 0 ? "（同表另外 " + shared + " 个源区域一并取消）" : "") + "。"
                : "已设置逐行重算：触发单元格 " + trigger.Cell + "，触发值 " + trigger.Values + "，文件名列 " + trigger.NameColumn + (shared > 0 ? "；已同时应用到同表的 " + shared + " 个源区域" : "") + "。在“映射”页或下方明细核对每一行的目标。");
        }
        private void ShowTriggerPopup(DataGridView grid, MigrationPair pair, MigrationTrigger draft)
        {
            if (Running || grid.ReadOnly || !pairs.Contains(pair)) return;
            var side = pair.Source;
            if (side.Books.Count != 1) { Message("逐行重算只支持一个源工作簿：请让这一行只选一个源工作簿。"); return; }
            if (side.Sheet.Length == 0) { Message("请先为这一行选择源工作表。"); return; }
            ClosePopups();
            var value = draft ?? (side.Trigger != null ? side.Trigger.Copy() : new MigrationTrigger());
            if (draft == null && side.Trigger == null)
                try
                {
                    // Guess the value list: column G from row 2 down to its last value.
                    var snapshot = LoadSnapshot(side.Books[0], side.Sheet); int last = 0;
                    for (int r = 2; r <= snapshot.MaxRow; r++) if (snapshot.Value(r, 7).Trim().Length > 0) last = r;
                    if (last >= 2) value.Values = "G2:G" + last.ToString(CultureInfo.InvariantCulture);
                }
                catch (Exception) { }
            var drop = new ToolStripDropDown { Padding = Padding.Empty, DropShadowEnabled = true, BackColor = Color.White, Renderer = new MigrationMenuRenderer() };
            var font = Ui.MakeFont(10f, FontStyle.Regular, Font);
            int p4 = Ui.Px(this, 4), p8 = Ui.Px(this, 8);
            var panel = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 3, BackColor = Color.White, Font = font, Padding = new Padding(Ui.Px(this, 16), Ui.Px(this, 14), Ui.Px(this, 16), Ui.Px(this, 12)) };
            var title = new Label { Text = "逐行重算", AutoSize = true, ForeColor = Ui.Ink, Font = Ui.MakeFont(10f, FontStyle.Bold, Font), Margin = new Padding(0, 0, 0, p8) };
            panel.Controls.Add(title, 0, 0); panel.SetColumnSpan(title, 3);
            var cellBox = new TextBox { Name = "migrationTriggerCell", Text = value.Cell, Width = Ui.Px(this, 150), Margin = new Padding(0, p4, p4, p4) };
            var valuesBox = new TextBox { Name = "migrationTriggerValues", Text = value.Values, Width = Ui.Px(this, 150), Margin = new Padding(0, p4, p4, p4) };
            var nameBox = new TextBox { Name = "migrationTriggerName", Text = value.NameColumn, Width = Ui.Px(this, 60), Margin = new Padding(0, p4, p4, p4) };
            Action keep = delegate { value.Cell = NormalizeAddress(cellBox.Text); value.Values = NormalizeAddress(valuesBox.Text); value.NameColumn = (nameBox.Text ?? "").Trim().ToUpperInvariant(); };
            Func<string, TextBox, bool, int, Control> line = delegate(string caption, TextBox box, bool pick, int at)
            {
                panel.Controls.Add(new Label { Text = caption, AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(0, Ui.Px(this, 8), p8, 0) }, 0, at);
                panel.Controls.Add(box, 1, at);
                if (!pick) return box;
                var button = new Label { Text = "", AutoSize = false, Size = new Size(Ui.Px(this, 28), Ui.Px(this, 26)), Cursor = Cursors.Hand, Margin = new Padding(0, p4, 0, p4) };
                button.Paint += delegate(object sender, PaintEventArgs e) { PaintPickIcon(e.Graphics, button.ClientRectangle, e.Graphics.DpiX / 96f, false, true); };
                button.Click += delegate
                {
                    keep(); drop.Close();
                    if (RangePickRequested == null) return;
                    bool forCell = box == cellBox;
                    RangePickRequested(false, side.Books[0], side.Sheet, forCell ? value.Cell : value.Values, delegate(string address)
                    {
                        var area = MigrationArea.Parse(address);
                        if (forCell) value.Cell = MigrationArea.Cell(area.Row, area.Column); else value.Values = area.Address;
                        ShowTriggerPopup(grid, pair, value);
                    });
                };
                restoreTip.SetToolTip(button, "在工作表中框选");
                panel.Controls.Add(button, 2, at); return box;
            };
            line("触发单元格", cellBox, true, 1);
            line("触发值", valuesBox, true, 2);
            line("目标文件名列", nameBox, false, 3);
            var note = new Label { Text = "每个触发值依次填入触发单元格，重算后读取本行源区域（只写数值），写入同一行文件名在已添加的目标工作簿中匹配到的那个工作簿。源文件只在副本中改动。", AutoSize = true, MaximumSize = new Size(Ui.Px(this, 320), 0), ForeColor = Ui.InkDim, Margin = new Padding(0, p8, 0, 0) };
            panel.Controls.Add(note, 0, 4); panel.SetColumnSpan(note, 3);
            var error = new Label { AutoSize = true, MaximumSize = new Size(Ui.Px(this, 320), 0), ForeColor = Ui.Danger, Margin = new Padding(0, p4, 0, 0) };
            panel.Controls.Add(error, 0, 5); panel.SetColumnSpan(error, 3);
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Dock = DockStyle.Fill, Margin = new Padding(0, p8, 0, 0), BackColor = Color.White };
            Func<string, bool, RoundedButton> make = delegate(string text, bool primary)
            {
                var b = Ui.Button(text, 72); if (primary) Ui.StylePrimary(b); else Ui.StyleSecondary(b);
                b.Font = font; b.MinimumSize = Size.Empty; b.Size = new Size(Ui.Px(this, text.Length > 2 ? 110 : 72), Ui.Px(this, 30)); b.Margin = new Padding(p8, 0, 0, 0); return b;
            };
            var ok = make("确定", true); var cancel = make("取消", false);
            buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
            if (side.Trigger != null)
            {
                var clear = make("取消逐行", false);
                clear.Click += delegate { drop.Close(); SetTrigger(side, null); };
                buttons.Controls.Add(clear);
            }
            panel.Controls.Add(buttons, 0, 6); panel.SetColumnSpan(buttons, 3);
            ok.Click += delegate
            {
                keep();
                try
                {
                    var cell = MigrationArea.Parse(value.Cell); if (cell.Rows != 1 || cell.Columns != 1) throw new InvalidOperationException("触发单元格只能是一个单元格");
                    var values = MigrationArea.Parse(value.Values); if (values.Columns != 1) throw new InvalidOperationException("触发值区域只能是一列");
                    int nameColumn = ColumnIndex(value.NameColumn);
                    if (nameColumn == values.Column) throw new InvalidOperationException("文件名列不能和触发值是同一列");
                    var source = TryParse(side.Range);
                    if (source != null && cell.Row >= source.Row && cell.Row < source.Row + source.Rows && cell.Column >= source.Column && cell.Column < source.Column + source.Columns)
                        throw new InvalidOperationException("触发单元格不能在源区域里");
                    value.Cell = cell.Address; value.Values = values.Address; value.Signature = "";
                }
                catch (Exception failure) { error.Text = failure.Message; return; }
                drop.Close(); SetTrigger(side, value);
            };
            cancel.Click += delegate { drop.Close(); };
            FramePopup(panel); panel.PerformLayout();
            var host = new ToolStripControlHost(panel) { AutoSize = false, Margin = Padding.Empty, Padding = Padding.Empty, Size = panel.PreferredSize };
            drop.Items.Add(host);
            error.SizeChanged += delegate { panel.PerformLayout(); host.Size = panel.PreferredSize; drop.Size = panel.PreferredSize; };
            drop.Closed += delegate { if (triggerPopup == drop) triggerPopup = null; BeginInvoke(new Action(delegate { drop.Dispose(); font.Dispose(); })); };
            triggerPopup = drop;
            var row = RowOf(grid, pair); int column = grid.Columns["trigger"].Index;
            var rect = row == null ? new Rectangle(0, 0, 0, 0) : grid.GetCellDisplayRectangle(column, row.Index, true);
            drop.Show(grid, new Point(Math.Max(0, rect.Right - panel.PreferredSize.Width), rect.Bottom));
            MigrationDropDownWatch.Attach(drop);
        }
        private static void PaintTriggerIcon(Graphics g, Rectangle cell, float u, bool hot, bool on, bool enabled)
        {
            var smoothing = g.SmoothingMode; g.SmoothingMode = SmoothingMode.AntiAlias;
            float size = 26 * u, x = cell.Left + (cell.Width - size) / 2f, y = cell.Top + (cell.Height - size) / 2f;
            if ((hot || on) && enabled)
                using (var path = Ui.RoundRect(new RectangleF(x, y, size, size), 6 * u)) using (var brush = new SolidBrush(on ? Color.FromArgb(232, 240, 254) : Color.FromArgb(242, 246, 255))) g.FillPath(brush, path);
            Color ink = !enabled ? Color.FromArgb(200, 206, 216) : on || hot ? Ui.Blue : Color.FromArgb(120, 132, 152);
            float cx = x + size / 2f, cy = y + size / 2f, r = 6.5f * u;
            using (var pen = new Pen(ink, 1.6f * u) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawArc(pen, cx - r, cy - r, r * 2, r * 2, 200, 130);
                g.DrawArc(pen, cx - r, cy - r, r * 2, r * 2, 20, 130);
            }
            using (var brush = new SolidBrush(ink))
                foreach (double degrees in new[] { 330.0, 150.0 })
                {
                    double a = degrees * Math.PI / 180; float px = cx + (float)(r * Math.Cos(a)), py = cy + (float)(r * Math.Sin(a));
                    float tx = (float)-Math.Sin(a), ty = (float)Math.Cos(a), nx = (float)Math.Cos(a), ny = (float)Math.Sin(a), head = 3.2f * u;
                    g.FillPolygon(brush, new[] { new PointF(px + tx * head, py + ty * head), new PointF(px + nx * head * 0.9f, py + ny * head * 0.9f), new PointF(px - nx * head * 0.9f, py - ny * head * 0.9f) });
                }
            g.SmoothingMode = smoothing;
        }
        // The details line under the mouse, or -1.
        private int DetailLineAt(Point location)
        {
            if ((detailLinks.Count == 0 && unusedLinks.Count == 0) || details.TextLength == 0) return -1;
            int index = details.GetCharIndexFromPosition(location); string text = details.Text;
            int line = 0; for (int i = 0; i < Math.Min(index, text.Length); i++) if (text[i] == '\n') line++;
            // The click must land on the line's text, not in the blank area below.
            var start = details.GetPositionFromCharIndex(index); if (location.Y > start.Y + details.Font.Height * 2) return -1;
            return line;
        }
        private KeyValuePair<MigrationSide, MigrationTriggerRow>? DetailLink(Point location)
        {
            KeyValuePair<MigrationSide, MigrationTriggerRow> link;
            return detailLinks.TryGetValue(DetailLineAt(location), out link) ? link : (KeyValuePair<MigrationSide, MigrationTriggerRow>?)null;
        }
        // Added targets that no row of this 逐行重算 matched and no ordinary
        // mapping writes to. They are left alone unless the user assigns one.
        private List<string> UnusedTargets(MigrationSide side)
        {
            var trigger = side.Trigger; if (trigger == null || trigger.Error.Length > 0) return new List<string>();
            var used = new HashSet<string>(trigger.Rows.Where(r => r.Path.Length > 0).Select(r => r.Path), StringComparer.OrdinalIgnoreCase);
            used.UnionWith(pairs.Where(p => p.Source.Trigger == null && p.Target.Used).SelectMany(p => p.Target.Books));
            return TargetBooks.Select(b => b.Path).Where(p => !used.Contains(p)).ToList();
        }
        private void ChooseTriggerTarget(MigrationTrigger trigger, MigrationTriggerRow row, string path)
        {
            trigger.Chosen[row.Row] = path; trigger.Signature = "";
            foreach (var pair in pairs) if (pair.Source.Trigger == trigger) pair.Instances.Clear();
            RefreshRows(); Changed(); ShowPairDetails();
            Message("触发第 " + row.Row + " 行" + (path.Length == 0 ? "已设为跳过。" : "的目标已定为 " + Path.GetFileName(path) + "。"));
        }
        // 未用到的目标: pick the row it belongs to. Rows without a target
        // (缺失, 跳过) come first; others would change their current target.
        private void ShowUnusedChoices(Point location, MigrationSide side, string path)
        {
            var trigger = side.Trigger; if (trigger == null) return; RefreshTrigger(side);
            var menu = new ContextMenuStrip { ShowImageMargin = false, Renderer = new MigrationMenuRenderer(), BackColor = Color.White, Font = Ui.MakeFont(10f, FontStyle.Regular, Font) };
            int h = Ui.Px(this, 10), v = Ui.Px(this, 5);
            menu.Items.Add(new ToolStripMenuItem("把 " + Path.GetFileName(path) + " 指定给：") { Enabled = false, Padding = new Padding(h, v, h, v) });
            var rows = trigger.Rows.Where(r => r.Kind != "空值").ToList();
            Func<MigrationTriggerRow, string> text = delegate(MigrationTriggerRow r)
            { return "第 " + r.Row + " 行  " + (r.Name.Length == 0 ? "（文件名为空）" : r.Name) + (r.Path.Length > 0 ? "    当前：" + Path.GetFileName(r.Path) : r.Kind == "跳过" ? "    已跳过" : "    无目标"); };
            var free = rows.Where(r => r.Path.Length == 0).ToList();
            foreach (var r in free.Take(12)) { var row = r; var item = new ToolStripMenuItem(text(row)) { Padding = new Padding(h, v, h, v) }; item.Click += delegate { ChooseTriggerTarget(trigger, row, path); }; menu.Items.Add(item); }
            var taken = rows.Where(r => r.Path.Length > 0).ToList();
            if (taken.Count > 0)
            {
                var other = new ToolStripMenuItem(free.Count > 0 ? "其他行（会替换它当前的目标）" : "替换某一行当前的目标") { Padding = new Padding(h, v, h, v) };
                other.DropDown.Renderer = new MigrationMenuRenderer(); other.DropDown.Font = menu.Font;
                foreach (var r in taken) { var row = r; var item = new ToolStripMenuItem(text(row)) { Padding = new Padding(h, v, h, v) }; item.Click += delegate { ChooseTriggerTarget(trigger, row, path); }; other.DropDownItems.Add(item); }
                menu.Items.Add(other);
            }
            if (rows.Count == 0) menu.Items.Add(new ToolStripMenuItem("没有可指定的触发行") { Enabled = false, Padding = new Padding(h, v, h, v) });
            menu.Closed += delegate { BeginInvoke(new Action(delegate { menu.Dispose(); })); };
            menu.Show(details, location);
            MigrationDropDownWatch.Attach(menu);
        }
        // Choose the target of one 逐行重算 row: a candidate, any added
        // target, skip, or back to the automatic match.
        private void ShowTriggerChoices(Point location)
        {
            if (Running) return;
            KeyValuePair<MigrationSide, string> spare;
            if (unusedLinks.TryGetValue(DetailLineAt(location), out spare)) { ShowUnusedChoices(location, spare.Key, spare.Value); return; }
            var link = DetailLink(location); if (link == null) return;
            var side = link.Value.Key; var row = link.Value.Value; var trigger = side.Trigger; if (trigger == null) return;
            var menu = new ContextMenuStrip { ShowImageMargin = false, Renderer = new MigrationMenuRenderer(), BackColor = Color.White, Font = Ui.MakeFont(10f, FontStyle.Regular, Font) };
            Action<string> choose = delegate(string path) { ChooseTriggerTarget(trigger, row, path); };
            int h = Ui.Px(this, 10), v = Ui.Px(this, 5);
            var head = new ToolStripMenuItem("第 " + row.Row + " 行：" + (row.Name.Length == 0 ? "（文件名为空）" : row.Name)) { Enabled = false, Padding = new Padding(h, v, h, v) };
            menu.Items.Add(head);
            foreach (var candidate in row.Candidates.Take(8))
            {
                string path = candidate.Key;
                var item = new ToolStripMenuItem(Path.GetFileName(path) + "    " + (candidate.Value >= 1 ? "精确" : "相似 " + Math.Round(candidate.Value * 100) + "%")) { Padding = new Padding(h, v, h, v), Checked = string.Equals(row.Path, path, StringComparison.OrdinalIgnoreCase) && row.Ready };
                item.Click += delegate { choose(path); }; menu.Items.Add(item);
            }
            if (row.Candidates.Count == 0)
                foreach (string spareTarget in UnusedTargets(side).Take(8))
                {
                    string path = spareTarget;
                    var item = new ToolStripMenuItem(Path.GetFileName(path) + "    未用到的目标") { Padding = new Padding(h, v, h, v) };
                    item.Click += delegate { choose(path); }; menu.Items.Add(item);
                }
            var all = new ToolStripMenuItem("从全部目标工作簿中选择") { Padding = new Padding(h, v, h, v) };
            all.DropDown.Renderer = new MigrationMenuRenderer(); all.DropDown.Font = menu.Font;
            foreach (var book in TargetBooks.OrderBy(b => b.Name, StringComparer.CurrentCulture))
            { string path = book.Path; var item = new ToolStripMenuItem(book.Name) { Padding = new Padding(h, v, h, v) }; item.Click += delegate { choose(path); }; all.DropDownItems.Add(item); }
            if (all.DropDownItems.Count > 0) menu.Items.Add(all);
            menu.Items.Add(new ToolStripSeparator());
            if (trigger.Chosen.ContainsKey(row.Row))
            {
                var auto = new ToolStripMenuItem("恢复自动匹配") { Padding = new Padding(h, v, h, v) };
                auto.Click += delegate { trigger.Chosen.Remove(row.Row); trigger.Signature = ""; foreach (var pair in pairs) if (pair.Source.Trigger == trigger) pair.Instances.Clear(); RefreshRows(); Changed(); ShowPairDetails(); };
                menu.Items.Add(auto);
            }
            if (row.Kind != "跳过") { var skip = new ToolStripMenuItem("跳过这一行") { Padding = new Padding(h, v, h, v), ForeColor = Ui.Danger }; skip.Click += delegate { choose(""); }; menu.Items.Add(skip); }
            menu.Closed += delegate { BeginInvoke(new Action(delegate { menu.Dispose(); })); };
            menu.Show(details, location);
            MigrationDropDownWatch.Attach(menu);
        }
        private void ShowPairDetails()
        {
            bool target; var pair = tabs == null ? null : PanelPair(out target); if (pair == null) { details.Text = ""; return; }
            var lines = new List<string>
            {
                "映射第 " + (pairs.IndexOf(pair) + 1) + " 行（" + (pair.Enabled ? "已启用" : "未启用") + "，" + (pair.ValuesOnly ? ModeValues : ModeFormula) + "）",
                "源：" + Where(pair.Source, TryParse(pair.Source.Range)) + "　工作簿：" + Books(pair.Source),
                "目标：" + Where(pair.Target, TargetArea(pair)) + "　工作簿：" + Books(pair.Target),
                "状态：" + PairStatus(pair)
            };
            foreach (var m in pair.Instances.Where(x => x.Result != null))
            {
                var t = m.Result;
                lines.Add("· " + Path.GetFileName(m.Source.Path) + " → " + Path.GetFileName(m.TargetPath) + "：" + t.Status + (t.Error.Length == 0 ? "" : "：" + t.Error)
                    + (t.Overwritten > 0 ? "；覆盖 " + t.Overwritten + " 格" : "") + (t.Output.Length == 0 ? "" : "；输出 " + t.Output) + (t.Backup.Length == 0 ? "" : "；备份 " + t.Backup));
                foreach (string sample in t.Samples.Take(3)) lines.Add("    " + sample);
            }
            detailLinks.Clear(); unusedLinks.Clear();
            var trigger = pair.Source.Trigger;
            if (trigger != null)
            {
                RefreshTrigger(pair.Source);
                lines.Add("逐行重算：触发单元格 " + trigger.Cell + "，触发值 " + trigger.Values + "，文件名列 " + trigger.NameColumn + "；目标按文件名在已添加的 " + TargetBooks.Count + " 个目标工作簿中匹配");
                if (trigger.Error.Length > 0) lines.Add("    设置无效：" + trigger.Error);
                foreach (var row in trigger.Rows)
                {
                    string mark = row.Kind == "精确" ? "✓ 精确" : row.Kind == "已确认" ? "✓ 已确认" : row.Kind == "模糊" ? "⚠ 模糊（点此确认或更换）"
                        : row.Kind == "多个候选" ? "⚠ 多个候选（点此选择）" : row.Kind == "缺失" ? "✗ 没有匹配的目标工作簿，将跳过（点此手动指定）"
                        : row.Kind == "跳过" ? "– 已设为跳过（点此恢复）" : "– 触发值为空，跳过";
                    if (row.Kind != "空值") detailLinks[lines.Count] = new KeyValuePair<MigrationSide, MigrationTriggerRow>(pair.Source, row);
                    lines.Add("第 " + row.Row + " 行  " + (row.Value.Length == 0 ? "（空）" : row.Value) + "  →  " + (row.Path.Length > 0 ? Path.GetFileName(row.Path) : row.Name.Length == 0 ? "（文件名为空）" : row.Name) + "    " + mark);
                }
                // The other way round: added targets no row matched.
                var unused = UnusedTargets(pair.Source);
                if (unused.Count > 0)
                {
                    lines.Add("未用到的目标工作簿 " + unused.Count + " 个：没有触发行匹配到它们，不会写入，也不会打开");
                    foreach (string path in unused)
                    {
                        unusedLinks[lines.Count] = new KeyValuePair<MigrationSide, string>(pair.Source, path);
                        lines.Add("– " + Path.GetFileName(path) + "    未用到（如需写入，点此指定给某一行）");
                    }
                }
            }
            details.Text = string.Join(Environment.NewLine, lines.ToArray());
        }

        // ---------------- plan ----------------
        // Expand every enabled row into concrete writes: one source workbook
        // to every target workbook, or N source workbooks one-to-one with N
        // target workbooks. Block overlaps and source-as-target, then group
        // the rest into passes for the existing single-source engine.
        internal MigrationPlan BuildPlan()
        {
            EndEdits();
            if (pairs.Count == 0) throw new InvalidOperationException("请先在“源区域”和“目标区域”页各新增至少一行。");
            var plan = new MigrationPlan { ReplaceTargets = replaceTargets.Checked, AllowEmptySource = allowEmpty.Checked };
            var items = new Dictionary<string, MigrationSourceItem>(StringComparer.OrdinalIgnoreCase);
            var targetPaths = new HashSet<string>(pairs.Where(p => p.Enabled && p.Target.Used).SelectMany(p => p.Target.Books), StringComparer.OrdinalIgnoreCase);
            var valid = new List<MigrationMapping>();
            for (int number = 1; number <= pairs.Count; number++)
            {
                var pair = pairs[number - 1]; pair.Instances.Clear();
                if (!pair.Enabled || RowProblem(pair).Length > 0) continue;
                var source = MigrationArea.Parse(pair.Source.Range); var area = TargetArea(pair);
                if (pair.Source.Trigger != null) { ExpandTrigger(plan, pair, number, area, targetPaths, valid); continue; }
                var books = pair.Source.Books.Count == 1 ? pair.Target.Books.Select(t => new KeyValuePair<string, string>(pair.Source.Books[0], t)).ToList()
                    : pair.Source.Books.Zip(pair.Target.Books, (s, t) => new KeyValuePair<string, string>(s, t)).ToList();
                foreach (var link in books)
                {
                    string key = link.Key + "\u001f" + pair.Source.Sheet + "\u001f" + pair.Source.Range + "\u001f" + pair.ValuesOnly; MigrationSourceItem item;
                    if (!items.TryGetValue(key, out item)) items[key] = item = new MigrationSourceItem { Path = link.Key, Sheet = pair.Source.Sheet, Range = pair.Source.Range, Name = "映射第 " + number + " 行", ValuesOnly = pair.ValuesOnly };
                    var m = new MigrationMapping { Pair = pair, Source = item, TargetPath = link.Value, TargetSheet = pair.Target.Sheet, Area = area, Start = MigrationArea.Cell(area.Row, area.Column), TargetRange = area.Address };
                    m.Result = new MigrationTarget { Path = link.Value, Sheet = pair.Target.Sheet, Start = m.Start, Status = "待预检" };
                    pair.Instances.Add(m); plan.Mappings.Add(m);
                    var sourceBook = FindBook(false, link.Key); var targetBook = FindBook(true, link.Value);
                    string problem = null;
                    if (sourceBook != null && sourceBook.Error.Length > 0) problem = "源工作簿读取失败：" + sourceBook.Error;
                    else if (targetBook != null && targetBook.Error.Length > 0) problem = "目标工作簿读取失败：" + targetBook.Error;
                    else if (sourceBook != null && sourceBook.Sheets.Count > 0 && !sourceBook.Sheets.Contains(pair.Source.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "源工作簿缺少工作表“" + pair.Source.Sheet + "”";
                    else if (targetBook != null && targetBook.Sheets.Count > 0 && !targetBook.Sheets.Contains(pair.Target.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "目标工作簿缺少工作表“" + pair.Target.Sheet + "”";
                    else if (string.Equals(link.Key, link.Value, StringComparison.OrdinalIgnoreCase) || targetPaths.Contains(link.Key)) problem = "源文件不能同时作为目标";
                    else { try { area.Check(1048576, 16384); } catch (Exception error) { problem = error.Message; } }
                    if (problem != null) Block(plan, m, problem); else valid.Add(m);
                }
            }
            // Overlap inside one target sheet is a conflict for both writes.
            var conflicted = new HashSet<MigrationMapping>();
            foreach (var group in valid.GroupBy(m => m.TargetPath + "\u001f" + m.TargetSheet.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase))
            {
                var list = group.ToList();
                for (int i = 0; i < list.Count; i++)
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        var a = list[i].Area; var b = list[j].Area;
                        int top = Math.Max(a.Row, b.Row), left = Math.Max(a.Column, b.Column);
                        int bottom = Math.Min(a.Row + a.Rows, b.Row + b.Rows), right = Math.Min(a.Column + a.Columns, b.Column + b.Columns);
                        if (top >= bottom || left >= right) continue;
                        string overlap = new MigrationArea { Row = top, Column = left, Rows = bottom - top, Columns = right - left }.Address;
                        // Two trigger rows sent to the same target area: the
                        // later row overwrites the earlier; ask, do not block.
                        if (list[i].Source.TriggerRow > 0 && list[j].Source.TriggerRow > 0)
                        {
                            plan.Overrides.Add(list[i].Source.Name + "、" + list[j].Source.Name + " → " + Path.GetFileName(list[i].TargetPath) + "  " + list[i].TargetSheet + "!" + overlap + "，最终保留后者的结果");
                            continue;
                        }
                        Note(list[i], "与" + list[j].Source.Name + "在 " + list[i].TargetSheet + "!" + overlap + " 重叠");
                        Note(list[j], "与" + list[i].Source.Name + "在 " + list[j].TargetSheet + "!" + overlap + " 重叠");
                        conflicted.Add(list[i]); conflicted.Add(list[j]);
                    }
            }
            foreach (var m in conflicted) { m.Result.Status = "冲突"; m.Result.Ready = false; plan.Blocked++; }
            Group(plan, valid.Where(m => !conflicted.Contains(m)).ToList(), plan.ReplaceTargets);
            lastPlan = plan; RefreshRows(); ShowPairDetails();
            if (plan.Passes.Count == 0) throw new InvalidOperationException(plan.Mappings.Count == 0 ? "没有可预检的映射：请在“映射”页查看各行“状态”提示的问题。" : "所有写入均未通过检查，请在“映射”页查看各行“状态”。");
            return plan;
        }
        // 逐行重算: one write per trigger value row, to the workbook its name
        // matched among the added targets. Rows not yet resolved are blocked
        // with the reason; empty, skipped and unmatched rows are left out.
        private void ExpandTrigger(MigrationPlan plan, MigrationPair pair, int number, MigrationArea area, HashSet<string> targetPaths, List<MigrationMapping> valid)
        {
            var trigger = pair.Source.Trigger; RefreshTrigger(pair.Source);
            if (trigger.Error.Length > 0) throw new InvalidOperationException("映射第 " + number + " 行的逐行重算设置无效：" + trigger.Error);
            string sourcePath = pair.Source.Books[0]; var values = MigrationArea.Parse(trigger.Values); int nameColumn = ColumnIndex(trigger.NameColumn);
            var sourceBook = FindBook(false, sourcePath);
            foreach (string spare in UnusedTargets(pair.Source)) if (!plan.Unused.Contains(spare, StringComparer.OrdinalIgnoreCase)) plan.Unused.Add(spare);
            foreach (var row in trigger.Rows)
            {
                string label = "映射第 " + number + " 行 · 触发第 " + row.Row + " 行";
                if (row.Kind == "空值" || row.Kind == "跳过" || row.Kind == "缺失")
                { plan.Skipped.Add(label + "：" + (row.Kind == "空值" ? "触发值为空" : row.Kind == "跳过" ? "已设为跳过" : "没有匹配的目标工作簿（" + row.Name + "）")); continue; }
                var item = new MigrationSourceItem
                {
                    Path = sourcePath, Sheet = pair.Source.Sheet, Range = pair.Source.Range, ValuesOnly = true, Name = label, TriggerRow = row.Row,
                    TriggerCell = MigrationArea.Parse(trigger.Cell).Address, TriggerFrom = MigrationArea.Cell(row.Row, values.Column), TriggerNameCell = MigrationArea.Cell(row.Row, nameColumn), TriggerExpect = row.Name
                };
                string path = row.Path.Length > 0 ? row.Path : "（未确定）";
                var m = new MigrationMapping { Pair = pair, Source = item, TargetPath = path, TargetSheet = pair.Target.Sheet, Area = area, Start = MigrationArea.Cell(area.Row, area.Column), TargetRange = area.Address };
                m.Result = new MigrationTarget { Path = path, Sheet = pair.Target.Sheet, Start = m.Start, Status = "待预检" };
                pair.Instances.Add(m); plan.Mappings.Add(m);
                string problem = null; var targetBook = row.Path.Length > 0 ? FindBook(true, row.Path) : null;
                if (row.Kind == "模糊") problem = "触发第 " + row.Row + " 行为模糊匹配（" + row.Name + " → " + Path.GetFileName(row.Path) + "），请在明细中确认";
                else if (row.Kind == "多个候选") problem = "触发第 " + row.Row + " 行“" + row.Name + "”有多个候选目标，请在明细中选择";
                else if (sourceBook != null && sourceBook.Error.Length > 0) problem = "源工作簿读取失败：" + sourceBook.Error;
                else if (targetBook != null && targetBook.Error.Length > 0) problem = "目标工作簿读取失败：" + targetBook.Error;
                else if (targetBook != null && targetBook.Sheets.Count > 0 && !targetBook.Sheets.Contains(pair.Target.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "目标工作簿缺少工作表“" + pair.Target.Sheet + "”";
                else if (string.Equals(path, sourcePath, StringComparison.OrdinalIgnoreCase) || targetPaths.Contains(sourcePath)) problem = "源文件不能同时作为目标";
                else { try { area.Check(1048576, 16384); } catch (Exception error) { problem = error.Message; } }
                if (problem != null) Block(plan, m, problem); else valid.Add(m);
            }
        }
        internal void RefreshResults()
        {
            bool ready = false, hasOutput = false;
            if (lastPlan != null)
            {
                var results = lastPlan.Mappings.Where(m => m.Result != null).Select(m => m.Result).ToList();
                ready = results.Any(t => t.Ready && t.Status == "可执行"); hasOutput = results.Any(t => t.Output.Length > 0);
            }
            hasOutput |= AllInstances().Any(m => m.Result != null && m.Result.Output.Length > 0);
            execute.Enabled = !Running && ready; open.Enabled = !Running && hasOutput;
            RefreshRows(); ShowPairDetails();
            if (!Running && lastPlan != null) status.Text = PlanSummary(lastPlan);
        }
        private void OpenResult()
        {
            bool side; var pair = PanelPair(out side);
            var target = pair == null ? null : pair.Instances.Select(m => m.Result).FirstOrDefault(t => t != null && File.Exists(t.Output));
            if (target == null) target = AllInstances().Select(m => m.Result).FirstOrDefault(t => t != null && File.Exists(t.Output));
            if (target != null) Process.Start(new ProcessStartInfo("explorer.exe", "/select,\"" + target.Output + "\"") { UseShellExecute = true });
        }
        private static void CenterSourceRow(FlowLayoutPanel row)
        {
            // Every item is centered on one line: fields and buttons by margin,
            // checkboxes by taking the full line height with centered text.
            foreach (var box in row.Controls.OfType<CheckBox>())
            { box.AutoSize = false; box.TextAlign = ContentAlignment.MiddleLeft; box.CheckAlign = ContentAlignment.MiddleLeft; box.Padding = Padding.Empty; }
            bool aligning = false;
            row.Layout += delegate
            {
                if (aligning || row.Controls.Count == 0) return;
                aligning = true;
                try
                {
                    var solid = row.Controls.Cast<Control>().Where(c => !(c is CheckBox)).ToList();
                    int height = solid.Count > 0 ? solid.Max(c => c.Height) : row.Controls.Cast<Control>().Max(c => c.GetPreferredSize(Size.Empty).Height);
                    foreach (Control control in row.Controls)
                    {
                        var box = control as CheckBox;
                        if (box != null)
                        {
                            var size = new Size(box.GetPreferredSize(Size.Empty).Width, height);
                            if (box.Size != size) box.Size = size;
                        }
                        int gap = height - control.Height;
                        var margin = new Padding(control.Margin.Left, gap / 2, control.Margin.Right, gap - gap / 2);
                        if (control.Margin != margin) control.Margin = margin;
                    }
                }
                finally { aligning = false; }
            };
        }
        internal const string FileFilter = "工作簿 (*.xlsx;*.xlsm;*.xls;*.et)|*.xlsx;*.xlsm;*.xls;*.et";
        private static void DrawSheetItem(object sender, DrawItemEventArgs e)
        {
            var combo = (ComboBox)sender; e.DrawBackground();
            string text = e.Index >= 0 && e.Index < combo.Items.Count ? combo.GetItemText(combo.Items[e.Index]) : combo.Text;
            var bounds = e.Bounds; int inset = Ui.Px(combo, 3); bounds.X += inset; bounds.Width = Math.Max(0, bounds.Width - inset * 2);
            TextRenderer.DrawText(e.Graphics, text, combo.Font, bounds, e.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }
        private static void AddGridColumn(DataGridView grid, string name, string title, int width, bool readOnly)
        { grid.Columns.Add(new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable }); }
        private static void AddCheckColumn(DataGridView grid, string title)
        { grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "include", HeaderText = title, Width = 56, SortMode = DataGridViewColumnSortMode.NotSortable, DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter } }); }
        private static void PrepareGrid(DataGridView grid)
        {
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = true;
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            grid.CurrentCellDirtyStateChanged += delegate { if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        }
        internal static string NormalizeAddress(string text)
        {
            // Full-width colon/space typed in an IME are normalized; anything
            // else (commas, missing colons) is left for Parse to reject.
            return (text ?? "").Trim().Replace('：', ':').Replace("　", "").ToUpperInvariant();
        }
        private static string SizeText(string range)
        {
            try { var area = MigrationArea.Parse(NormalizeAddress(range)); return area.Rows + " 行 × " + area.Columns + " 列"; }
            catch { return ""; }
        }
        private void FitLikeTargets(ResultGrid grid, Dictionary<string, int> weights)
        {
            grid.ColumnFitOverride = delegate { FitColumns(grid, weights); };
            grid.ClientSizeChanged += delegate { Ui.AutoFitColumns(grid); };
            grid.RowsRemoved += delegate { Ui.AutoFitColumns(grid); };
            grid.RowsAdded += delegate { Ui.AutoFitColumns(grid); };
            foreach (var scroll in grid.Controls.OfType<VScrollBar>())
            { scroll.VisibleChanged += delegate { Ui.AutoFitColumns(grid); }; scroll.SizeChanged += delegate { Ui.AutoFitColumns(grid); }; }
        }
        private void FitColumns(DataGridView grid, Dictionary<string, int> weights)
        {
            if (fittingGrid || grid.IsDisposed || grid.Disposing || grid.ClientSize.Width < 1) return;
            fittingGrid = true;
            try
            {
                Ui.EnsureGridHeaders(grid);
                foreach (string narrow in new[] { "pick", "trigger" }) if (grid.Columns.Contains(narrow)) grid.Columns[narrow].MinimumWidth = Ui.Px(grid, 40);
                if (grid.Columns.Contains("sheet") && grid.Columns["sheet"].HeaderCell is MigrationSelectionHeader)
                    grid.Columns["sheet"].MinimumWidth = Math.Max(grid.Columns["sheet"].MinimumWidth, Ui.Px(grid, 104));
                for (int pass = 0; pass < 3; pass++)
                {
                    int available = Math.Max(0, grid.ClientSize.Width - (grid.RowHeadersVisible ? grid.RowHeadersWidth : 0) - grid.Controls.OfType<VScrollBar>().Where(c => c.Visible).Sum(c => c.Width));
                    var columns = grid.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    var widths = columns.ToDictionary(c => c, c => c.MinimumWidth);
                    var pending = columns.Where(c => weights.ContainsKey(c.Name)).ToList();
                    int remaining = available - columns.Where(c => !weights.ContainsKey(c.Name)).Sum(c => c.MinimumWidth);
                    int weight = pending.Sum(c => weights[c.Name]);
                    while (pending.Count > 0 && weight > 0)
                    {
                        bool pinned = false;
                        foreach (var column in pending.ToList())
                            if ((double)remaining * weights[column.Name] / weight < widths[column])
                            { remaining -= widths[column]; weight -= weights[column.Name]; pending.Remove(column); pinned = true; }
                        if (pinned) continue;
                        int assigned = 0;
                        for (int i = 0; i < pending.Count; i++)
                        {
                            var column = pending[i];
                            widths[column] = i == pending.Count - 1 ? remaining - assigned : (int)((double)remaining * weights[column.Name] / weight);
                            assigned += widths[column];
                        }
                        break;
                    }
                    foreach (var column in columns) if (column.Width != widths[column]) column.Width = Math.Max(column.MinimumWidth, widths[column]);
                    int after = Math.Max(0, grid.ClientSize.Width - (grid.RowHeadersVisible ? grid.RowHeadersWidth : 0) - grid.Controls.OfType<VScrollBar>().Where(c => c.Visible).Sum(c => c.Width));
                    if (after == available) break;
                }
            }
            finally { fittingGrid = false; }
        }
        private MigrationSheetSnapshot LoadSnapshot(string path, string sheet)
        {
            string key = path + "\u001f" + sheet; MigrationSheetSnapshot cached;
            DateTime stamp = File.GetLastWriteTimeUtc(path);
            if (snapshots.TryGetValue(key, out cached) && cached.Stamp == stamp) return cached;
            var snapshot = MigrationSheetSnapshot.Load(path, sheet);
            snapshots[key] = snapshot; return snapshot;
        }
        private static bool Overlaps(MigrationArea a, MigrationArea b)
        { return a != null && b != null && a.Row < b.Row + b.Rows && b.Row < a.Row + a.Rows && a.Column < b.Column + b.Columns && b.Column < a.Column + a.Columns; }
        private static bool SameShape(MigrationMapping m)
        {
            try { var source = MigrationArea.Parse(m.Source.Range); return m.Area.Rows == source.Rows && m.Area.Columns == source.Columns; }
            catch { return false; }
        }
        private static void Block(MigrationPlan plan, MigrationMapping m, string problem)
        { m.Result.Status = "失败"; m.Result.Error = "检查：" + problem; m.Result.Ready = false; plan.Blocked++; }
        private static void Note(MigrationMapping m, string text)
        { m.Result.Error = m.Result.Error.Length == 0 ? text : m.Result.Error + "；" + text; }
        // Group mappings by source range; a target may occur once per pass.
        internal static void Group(MigrationPlan plan, List<MigrationMapping> mappings, bool replace)
        {
            foreach (var bySource in mappings.GroupBy(m => m.Source.Id))
            {
                var pending = bySource.ToList();
                while (pending.Count > 0)
                {
                    var source = pending[0].Source;
                    var pass = new MigrationPass { Source = source, Request = new MigrationRequest { SourcePath = source.Path, SourceSheet = source.Sheet, SourceRange = source.Range, AllowEmptySource = plan.AllowEmptySource, ReplaceTargets = replace, ValuesOnly = source.ValuesOnly,
                        TriggerCell = source.TriggerCell, TriggerFrom = source.TriggerFrom, TriggerNameCell = source.TriggerNameCell, TriggerExpect = source.TriggerExpect } };
                    var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var m in pending.ToList())
                    {
                        if (!used.Add(m.TargetPath)) continue;
                        pass.Mappings.Add(m); pass.Request.Targets.Add(m.Result); pending.Remove(m);
                    }
                    plan.Passes.Add(pass);
                }
            }
        }
        // The mappings that passed preview, regrouped with fresh state so
        // each pass re-checks its targets right before writing.
        internal MigrationPlan ExecutionPlan(MigrationPlan previewed)
        {
            var plan = new MigrationPlan { ReplaceTargets = previewed.ReplaceTargets, AllowEmptySource = previewed.AllowEmptySource, OutputCopies = !previewed.ReplaceTargets };
            var ready = previewed.Mappings.Where(m => m.Result != null && m.Result.Ready && m.Result.Status == "可执行").ToList();
            foreach (var m in ready) { m.Result = new MigrationTarget { Path = m.TargetPath, Sheet = m.TargetSheet, Start = m.Start, Status = "待执行" }; plan.Mappings.Add(m); }
            // Every pass updates its target in place. For result copies the
            // orchestrator first points each target at one "_迁移结果" copy,
            // so all source ranges land in the same copy.
            Group(plan, ready, true);
            lastPlan = plan; RefreshRows();
            return plan;
        }
        internal static string PlanSummary(MigrationPlan plan)
        {
            var results = plan.Mappings.Where(m => m.Result != null).Select(m => m.Result).ToList();
            int ok = results.Count(t => t.Status == "成功"), failed = results.Count(t => t.Status == "失败" || t.Status == "冲突"), ready = results.Count(t => t.Ready && t.Status == "可执行");
            int rows = plan.Mappings.Select(m => m.Pair).Distinct().Count(), skipped = plan.Skipped.Count, targets = plan.Mappings.Select(m => m.TargetPath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            return "对照 " + rows + " 行，目标工作簿 " + targets + " 个，共 " + plan.Mappings.Count + " 处写入：成功 " + ok + "，失败/冲突 " + failed + "，可执行 " + ready + "。"
                + (skipped > 0 ? "逐行重算跳过 " + skipped + " 行（见明细）。" : "")
                + (plan.Unused.Count > 0 ? "另有 " + plan.Unused.Count + " 个目标工作簿没有触发行匹配，不会写入（见明细，可点击指定）。" : "")
                + (plan.ReplaceTargets ? "保存方式：校验后更新原文件并备份。" : "保存方式：另存结果副本。");
        }
        internal void RecordRestorePoint(string path, string backup)
        {
            if (string.IsNullOrEmpty(backup) || !File.Exists(backup) || restorePoints.Any(p => string.Equals(p.Backup, backup, StringComparison.OrdinalIgnoreCase))) return;
            try { restorePoints.Add(new RestorePoint { Path = path, Backup = backup, Hash = DataMigrationHost.Digest(path) }); }
            catch (Exception error) { UiDiagnostics.Record("migration-restore-point", error); }
        }
        internal static string Summary(MigrationRequest request)
        {
            int succeeded = request.Targets.Count(t => t.Status == "成功"), failed = request.Targets.Count(t => t.Status == "失败"), ready = request.Targets.Count(t => t.Ready && t.Status == "可执行");
            return "源 " + request.SourceSheet + "!" + request.SourceRange + "：非空 " + request.NonEmptyCells + "，公式 " + request.FormulaCells
                + "；目标 " + request.Targets.Count + "：成功 " + succeeded + "，失败 " + failed + "，可执行 " + ready + "。" + (request.ReplaceTargets ? "保存方式：校验后更新原文件并备份。" : "保存方式：另存结果副本。");
        }
        internal void Run(IEnumerable<MigrationProgress> sequence, Action complete, Action<Exception> error)
        {
            if (Running) throw new InvalidOperationException("当前任务尚未结束。");
            ClosePopups();
            finished = complete; failed = error; steps = sequence.GetEnumerator();
            foreach (Control input in inputs) input.Enabled = false; SetGridsReadOnly(true); preview.Enabled = execute.Enabled = open.Enabled = restore.Enabled = clearBackups.Enabled = false;
            guardTail.Stop(); hostAtRun = MigrationWindowScope.Capture(IsHandleCreated ? Handle : IntPtr.Zero); topMostBeforeRun = TopMost;
            // Freeze the status line's size while running so changing progress
            // text never re-lays out the grid and details above it.
            status.AutoSize = false; status.AutoEllipsis = true; statusClock.Reset();
            pump.Start();
        }
        private void Tick()
        {
            if (steps == null || ticking) return;
            // Native COM calls can dispatch Windows messages on this same STA.
            // A nested timer tick must never advance/dispose the same iterator.
            ticking = true; pump.Stop();
            try
            {
                var clock = Stopwatch.StartNew();
                string message = null;
                bool pause = false;
                do { if (!steps.MoveNext()) { Finish(null); return; } message = steps.Current.Message; pause = steps.Current.Pause; } while (!pause && clock.ElapsedMilliseconds < 12);
                // A pause step announces a long blocking host call (opening or
                // saving a large workbook): show it now and paint before the
                // next step blocks this thread, so the window does not look hung.
                if (message != null && (pause || !statusClock.IsRunning || statusClock.ElapsedMilliseconds >= StatusIntervalMilliseconds) && status.Text != message)
                { status.Text = message; statusClock.Restart(); if (pause) status.Refresh(); }
                GuardWindow();
            }
            catch (Exception error) { Finish(error); }
            finally { ticking = false; if (!IsDisposed && !Disposing && steps != null) pump.Start(); }
        }
        private void Finish(Exception error)
        {
            pump.Stop(); var current = steps; steps = null; statusClock.Reset();
            status.AutoEllipsis = false; status.AutoSize = true;
            try { if (current != null) current.Dispose(); } catch (Exception cleanup) { error = error ?? cleanup; }
            foreach (Control input in inputs) input.Enabled = true; SetGridsReadOnly(false); preview.Enabled = true;
            if (error == null) { status.Text = lastPlan == null ? "请查看各行结果。" : PlanSummary(lastPlan); if (finished != null) finished(); }
            else { status.Text = error.Message; if (failed != null) failed(error); }
            RefreshResults(); UpdateRestore(); if (error != null) status.Text = error.Message;
            // WPS may still minimize/cover its frame shortly after the last
            // Close; keep watching for about a second, then release TopMost.
            GuardWindow(); TopMost = topMostBeforeRun; guardTailTicks = 10; guardTail.Start();
            if (closeWhenDone) Close();
        }
        private void GuardWindow()
        {
            if (IsDisposed || !IsHandleCreated || !Visible || closeWhenDone) return;
            MigrationWindowScope.KeepShown(Handle, hostAtRun);
            // While running and the user is in WPS, stay above a maximized
            // frame that WPS re-activated; never cover other programs.
            bool top = topMostBeforeRun || (Running && MigrationWindowScope.HostInForeground());
            if (TopMost != top) TopMost = top;
        }
        private void UpdateRestore()
        {
            restore.Enabled = !Running && restorePoints.Any(p => !p.Restored && !p.Cleared);
            clearBackups.Enabled = !Running && restorePoints.Any(p => !p.Cleared);
        }
        // Files this window created in "迁移备份": the pre-migration backup
        // and, after a restore, the displaced migrated version.
        private static IEnumerable<string> BackupFiles(RestorePoint point)
        {
            if (!string.IsNullOrEmpty(point.Backup)) yield return point.Backup;
            if (!string.IsNullOrEmpty(point.Displaced)) yield return point.Displaced;
        }
        internal void ClearBackups(bool confirm = true)
        {
            if (Running) return;
            var pending = restorePoints.Where(p => !p.Cleared).ToList();
            if (pending.Count == 0) { UpdateRestore(); return; }
            int files = pending.Sum(p => BackupFiles(p).Count(File.Exists));
            int restorable = pending.Count(p => !p.Restored);
            string question = "将永久删除本窗口迁移时生成的 " + files + " 个备份文件（涉及 " + pending.Count + " 个目标工作簿）。" + Environment.NewLine + Environment.NewLine
                + (restorable > 0 ? "删除后，这 " + restorable + " 个已更新的工作簿将无法再还原为迁移前的版本。" : "删除后无法恢复这些备份。") + Environment.NewLine
                + "此操作无法撤销。" + Environment.NewLine + Environment.NewLine + "是否清除备份？";
            if (confirm && UiDialog.Show(this, question, "清除备份", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            int cleared = 0; var failures = new List<string>();
            foreach (var point in pending)
            {
                try { MigrationOutputTransaction.DeleteBackups(BackupFiles(point).ToList()); point.Cleared = true; cleared++; }
                catch (Exception error) { failures.Add(Path.GetFileName(point.Path) + "：" + error.Message); }
            }
            UpdateRestore();
            Message((cleared > 0 ? "已清除 " + cleared + " 个目标的备份，这些工作簿不能再还原。" : "未清除任何备份。")
                + (failures.Count == 0 ? "" : " 未清除：" + string.Join("；", failures.ToArray())));
        }
        internal int RestorableCount { get { return restorePoints.Count(p => !p.Restored); } }
        internal void RestoreTargets(bool confirm = true)
        {
            if (Running) return;
            // Newest first: a file migrated twice steps back one batch at a time.
            var pending = restorePoints.Where(p => !p.Restored && !p.Cleared).Reverse().ToList();
            if (pending.Count == 0) { UpdateRestore(); return; }
            string question = "将把 " + pending.Count + " 个已更新的目标工作簿还原为迁移前的备份。" + Environment.NewLine + Environment.NewLine
                + "· 请先在 WPS 中关闭这些工作簿。" + Environment.NewLine
                + "· 迁移后的版本会另存到对应目录的“迁移备份”文件夹，不会丢失。" + Environment.NewLine
                + "· 迁移后又被修改过的文件不会还原，以免覆盖这些修改。" + Environment.NewLine + Environment.NewLine + "是否继续？";
            if (confirm && UiDialog.Show(this, question, "还原操作", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            int restored = 0; var failures = new List<string>();
            foreach (var point in pending)
            {
                string failure = null;
                try { point.Displaced = MigrationOutputTransaction.Restore(point.Path, point.Backup, point.Hash); point.Restored = true; restored++; }
                catch (Exception error) { failure = error.Message; failures.Add(Path.GetFileName(point.Path) + "：" + error.Message); }
                foreach (var target in AllInstances().Select(m => m.Result).Where(t => t != null && string.Equals(t.Path, point.Path, StringComparison.OrdinalIgnoreCase)))
                {
                    if (failure == null) { target.Status = "已还原"; target.Error = ""; target.Ready = false; }
                    else target.Error = "还原失败：" + failure;
                }
            }
            RefreshResults(); UpdateRestore();
            Message((restored > 0 ? "已还原 " + restored + " 个工作簿为迁移前的备份；如需再次迁移请重新预检。" : "未还原任何工作簿。")
                + (failures.Count == 0 ? "" : " 未还原：" + string.Join("；", failures.ToArray())));
        }
        // ---------------- presets ----------------
        // The whole setup as preset data: workbooks (absolute path plus the
        // path relative to their common folder), ranges, mappings, write
        // modes and 逐行重算 settings. Never any sheet data.
        internal Dictionary<string, object> CapturePreset(string name)
        {
            EndEdits();
            var sources = DistinctSides(false); var targets = DistinctSides(true);
            var paths = SourceBooks.Select(b => b.Path).Concat(TargetBooks.Select(b => b.Path)).Concat(sources.Concat(targets).SelectMany(s => s.Books))
                .Concat(sources.Where(s => s.Trigger != null).SelectMany(s => s.Trigger.Chosen.Values).Where(v => v.Length > 0)).ToList();
            string root = CommonFolder(paths);
            Func<string, object> book = delegate(string path) { return MigrationJson.Obj("path", path, "relative", RelativeTo(root, path)); };
            Func<MigrationSide, string, bool, object> sideData = delegate(MigrationSide side, string id, bool target)
            {
                object trigger = null;
                if (!target && side.Trigger != null)
                    trigger = MigrationJson.Obj("cell", side.Trigger.Cell, "values", side.Trigger.Values, "nameColumn", side.Trigger.NameColumn,
                        "chosen", side.Trigger.Chosen.OrderBy(c => c.Key).Select(c => (object)MigrationJson.Obj("row", (long)c.Key, "path", c.Value, "relative", c.Value.Length == 0 ? "" : RelativeTo(root, c.Value))).ToList());
                var data = MigrationJson.Obj("id", id, "books", side.Books.Select(book).ToList(), "sheet", side.Sheet, "range", side.Range);
                if (trigger != null) data["trigger"] = trigger;
                return data;
            };
            var mappings = pairs.Select(p => (object)MigrationJson.Obj(
                "source", p.Source.Used ? (object)("s" + (sources.IndexOf(p.Source) + 1)) : null,
                "target", p.Target.Used ? (object)("t" + (targets.IndexOf(p.Target) + 1)) : null,
                "enabled", p.Enabled, "valuesOnly", p.ValuesOnly)).ToList();
            return MigrationJson.Obj(
                "format", MigrationPreset.Format, "version", (long)MigrationPreset.Version, "name", name,
                "saved", DateTime.Now.ToString("s", CultureInfo.InvariantCulture), "root", root,
                "options", MigrationJson.Obj("replaceTargets", replaceTargets.Checked, "allowEmptySource", allowEmpty.Checked),
                "sourceBooks", SourceBooks.Select(b => book(b.Path)).ToList(),
                "targetBooks", TargetBooks.Select(b => book(b.Path)).ToList(),
                "sources", sources.Select((s, i) => sideData(s, "s" + (i + 1), false)).ToList(),
                "targets", targets.Select((s, i) => sideData(s, "t" + (i + 1), true)).ToList(),
                "mappings", mappings);
        }
        private static string CommonFolder(IEnumerable<string> paths)
        {
            string[] common = null;
            foreach (string path in paths)
            {
                string folder = Path.GetDirectoryName(path); if (string.IsNullOrEmpty(folder)) continue;
                var parts = folder.TrimEnd('\\').Split('\\');
                if (common == null) { common = parts; continue; }
                int n = 0; while (n < common.Length && n < parts.Length && string.Equals(common[n], parts[n], StringComparison.OrdinalIgnoreCase)) n++;
                common = common.Take(n).ToArray();
            }
            return common == null || common.Length == 0 ? "" : string.Join("\\", common) + (common.Length == 1 ? "\\" : "");
        }
        private static string RelativeTo(string root, string path)
        {
            if (root.Length == 0) return Path.GetFileName(path);
            string prefix = root.EndsWith("\\") ? root : root + "\\";
            return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? path.Substring(prefix.Length) : Path.GetFileName(path);
        }
        // Replace the whole setup with a preset's. remap: old path -> path on
        // this computer (from 定位缺失的工作簿). Only settings change: nothing
        // is read or written until the user runs 预检.
        internal void ApplyPresetData(Dictionary<string, object> data, IDictionary<string, string> remap)
        {
            if (Running) throw new InvalidOperationException("迁移运行期间不能应用预设。");
            MigrationPreset.From(data, null);
            Func<string, string> map = delegate(string path)
            {
                string found; if (path.Length == 0) return path;
                return remap != null && remap.TryGetValue(path, out found) ? found : path;
            };
            Func<Dictionary<string, object>, string> pathOf = delegate(Dictionary<string, object> b) { return map(MigrationJson.Str(b, "path")); };
            ClosePopups(); EndEdits();
            pairs.Clear(); SourceBooks.Clear(); TargetBooks.Clear(); mappingFocus = null; lastPlan = null;
            var options = MigrationJson.Map(data, "options");
            syncing = true;
            try { replaceTargets.Checked = MigrationJson.Flag(options, "replaceTargets", true); allowEmpty.Checked = MigrationJson.Flag(options, "allowEmptySource", false); }
            finally { syncing = false; }
            LoadPresetBooks(false, MigrationJson.Maps(data, "sourceBooks").Select(pathOf));
            LoadPresetBooks(true, MigrationJson.Maps(data, "targetBooks").Select(pathOf));
            var sides = new Dictionary<string, MigrationSide>(StringComparer.Ordinal);
            foreach (string list in new[] { "sources", "targets" })
                foreach (var item in MigrationJson.Maps(data, list))
                {
                    var side = new MigrationSide { Used = true, Sheet = MigrationJson.Str(item, "sheet"), Range = NormalizeAddress(MigrationJson.Str(item, "range")) };
                    foreach (var b in MigrationJson.Maps(item, "books")) { string path = pathOf(b); if (path.Length > 0) side.Books.Add(Path.GetFullPath(path)); }
                    var trigger = MigrationJson.Map(item, "trigger");
                    if (list == "sources" && trigger != null)
                    {
                        side.Trigger = new MigrationTrigger { Cell = MigrationJson.Str(trigger, "cell"), Values = MigrationJson.Str(trigger, "values"), NameColumn = MigrationJson.Str(trigger, "nameColumn") };
                        foreach (var chosen in MigrationJson.Maps(trigger, "chosen"))
                        { string path = pathOf(chosen); side.Trigger.Chosen[(int)MigrationJson.Num(chosen, "row")] = path.Length == 0 ? "" : Path.GetFullPath(path); }
                    }
                    sides[(list == "sources" ? "S:" : "T:") + MigrationJson.Str(item, "id")] = side;
                }
            foreach (var item in MigrationJson.Maps(data, "mappings"))
            {
                var pair = new MigrationPair { Enabled = MigrationJson.Flag(item, "enabled", true), ValuesOnly = MigrationJson.Flag(item, "valuesOnly", false) };
                MigrationSide side;
                if (sides.TryGetValue("S:" + MigrationJson.Str(item, "source"), out side)) pair.Source = side;
                if (sides.TryGetValue("T:" + MigrationJson.Str(item, "target"), out side)) pair.Target = side;
                if (pair.Source.Used || pair.Target.Used) pairs.Add(pair);
            }
            RebuildViews(); Changed(); ShowPage(pairs.Count > 0 ? 2 : 0);
        }
        private void LoadPresetBooks(bool target, IEnumerable<string> paths)
        {
            var list = paths.Where(p => p.Length > 0).Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var present = list.Where(File.Exists).ToArray();
            if (present.Length > 0 && BooksPicked != null) BooksPicked(target, present);
            foreach (string path in list)
                if (FindBook(target, path) == null) AddBook(target, path, new string[0], File.Exists(path) ? "" : "找不到文件，请重新选择");
        }
        private bool HasSetup { get { return pairs.Count > 0 || SourceBooks.Count > 0 || TargetBooks.Count > 0; } }
        private string SuggestedPresetName()
        {
            var source = SourceBooks.FirstOrDefault();
            return (source == null ? "迁移预设" : Path.GetFileNameWithoutExtension(source.Path)) + " " + DateTime.Now.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture).Replace(':', '.');
        }
        // 保存为预设: returns the saved file, or null when cancelled.
        internal string SavePreset()
        {
            if (Running) return null;
            if (!HasSetup) { Message("当前没有可保存的设置：先添加工作簿和区域。"); return null; }
            string name = MigrationNameDialog.Ask(this, "保存为预设", "预设名称（保存在本机预设库，可在“预设管理”中导出分享）：", SuggestedPresetName());
            if (name == null) return null;
            string file = MigrationPreset.FileFor(name);
            if (File.Exists(file) && UiDialog.Show(this, "已有同名预设“" + name + "”，是否覆盖？", "保存为预设", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return null;
            try { MigrationPreset.Write(file, CapturePreset(name)); Message("已保存预设“" + name + "”。"); return file; }
            catch (Exception error) { Message("保存预设失败：" + error.Message); return null; }
        }
        internal void ShowPresetManager()
        {
            if (Running) { Message("迁移运行期间不能管理预设。"); return; }
            ClosePopups();
            string json = HostDialogService.ShowIsolatedResult<string>(this,
                delegate { return new MigrationPresetManager(null); },
                delegate(Form dialog)
                {
                    var chosen = ((MigrationPresetManager)dialog).Chosen;
                    return chosen == null ? null : MigrationJson.Write(chosen.Data);
                }, null);
            if (json == null) return;
            ApplyPreset(MigrationPreset.From((Dictionary<string, object>)MigrationJson.Parse(json), null));
        }
        // Apply: offer to keep the current setup first, relocate workbooks
        // missing on this computer, then restore the settings only.
        private void ApplyPreset(MigrationPreset preset)
        {
            if (HasSetup)
            {
                var answer = UiDialog.Show(this, "应用预设“" + preset.Name + "”会替换当前窗口中的全部设置。" + Environment.NewLine + Environment.NewLine + "是否先把当前设置保存为预设？",
                    "应用预设", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (answer == DialogResult.Cancel) return;
                if (answer == DialogResult.Yes && SavePreset() == null) return;
            }
            var missing = preset.Files().Where(f => !File.Exists(f)).ToList();
            var remap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (missing.Count > 0)
            {
                string json = MigrationJson.Write(preset.Data);
                string[] paths = missing.ToArray();
                string[] located = HostDialogService.ShowIsolatedResult<string[]>(this,
                    delegate { return new MigrationRelocateDialog(MigrationPreset.From((Dictionary<string, object>)MigrationJson.Parse(json), null), new List<string>(paths)); },
                    delegate(Form dialog) { return ((MigrationRelocateDialog)dialog).Found.SelectMany(pair => new[] { pair.Key, pair.Value }).ToArray(); }, null);
                if (located == null) { Message("已取消应用预设，当前设置未改变。"); return; }
                for (int i = 0; i < located.Length; i += 2) remap[located[i]] = located[i + 1];
            }
            try { ApplyPresetData(preset.Data, remap); }
            catch (Exception error) { Message("应用预设失败：" + error.Message); return; }
            int unresolved = missing.Count(m => !remap.ContainsKey(m));
            Message("已应用预设“" + preset.Name + "”：只恢复了设置，没有读写任何工作簿。请检查各页后点“预检与预览”。" + (unresolved > 0 ? "仍有 " + unresolved + " 个工作簿找不到，相关区域需要重新选择。" : ""));
        }

        internal void Message(string text) { status.Text = text; }
        public Size MeasureDefaultClient() { return new Size(1280, 800); }
    }

    // Minimal JSON for presets: objects are Dictionary<string, object>,
    // arrays List<object>, numbers long or double, plus string/bool/null.
    internal static class MigrationJson
    {
        internal static string Write(object value)
        {
            var text = new StringBuilder(); Write(text, value, 0); return text.ToString();
        }
        private static void Indent(StringBuilder text, int depth) { text.Append('\n').Append(' ', depth * 2); }
        private static void Write(StringBuilder text, object value, int depth)
        {
            var map = value as IDictionary<string, object>; var list = value as IList<object>;
            if (value == null) text.Append("null");
            else if (value is string) Quote(text, (string)value);
            else if (value is bool) text.Append((bool)value ? "true" : "false");
            else if (value is int || value is long) text.Append(Convert.ToInt64(value).ToString(CultureInfo.InvariantCulture));
            else if (value is double) text.Append(((double)value).ToString("R", CultureInfo.InvariantCulture));
            else if (map != null)
            {
                if (map.Count == 0) { text.Append("{}"); return; }
                text.Append('{'); bool first = true;
                foreach (var pair in map)
                {
                    if (!first) text.Append(','); first = false;
                    Indent(text, depth + 1); Quote(text, pair.Key); text.Append(": "); Write(text, pair.Value, depth + 1);
                }
                Indent(text, depth); text.Append('}');
            }
            else if (list != null)
            {
                if (list.Count == 0) { text.Append("[]"); return; }
                text.Append('[');
                for (int i = 0; i < list.Count; i++) { if (i > 0) text.Append(','); Indent(text, depth + 1); Write(text, list[i], depth + 1); }
                Indent(text, depth); text.Append(']');
            }
            else Quote(text, Convert.ToString(value, CultureInfo.InvariantCulture));
        }
        private static void Quote(StringBuilder text, string value)
        {
            text.Append('"');
            foreach (char c in value)
            {
                if (c == '"') text.Append("\\\"");
                else if (c == '\\') text.Append("\\\\");
                else if (c == '\n') text.Append("\\n");
                else if (c == '\r') text.Append("\\r");
                else if (c == '\t') text.Append("\\t");
                else if (c < ' ') text.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                else text.Append(c);
            }
            text.Append('"');
        }
        internal static object Parse(string text)
        {
            int at = 0; var value = Value(text, ref at); Space(text, ref at);
            if (at != text.Length) throw new FormatException("JSON 末尾有多余内容（位置 " + at + "）");
            return value;
        }
        private static void Space(string text, ref int at) { while (at < text.Length && char.IsWhiteSpace(text[at])) at++; }
        private static Exception Bad(int at) { return new FormatException("JSON 格式错误（位置 " + at + "）"); }
        private static object Value(string text, ref int at)
        {
            Space(text, ref at); if (at >= text.Length) throw Bad(at);
            char c = text[at];
            if (c == '{')
            {
                var map = new Dictionary<string, object>(StringComparer.Ordinal); at++; Space(text, ref at);
                if (at < text.Length && text[at] == '}') { at++; return map; }
                while (true)
                {
                    Space(text, ref at); if (at >= text.Length || text[at] != '"') throw Bad(at);
                    string key = ReadString(text, ref at); Space(text, ref at);
                    if (at >= text.Length || text[at] != ':') throw Bad(at); at++;
                    map[key] = Value(text, ref at); Space(text, ref at);
                    if (at < text.Length && text[at] == ',') { at++; continue; }
                    if (at < text.Length && text[at] == '}') { at++; return map; }
                    throw Bad(at);
                }
            }
            if (c == '[')
            {
                var list = new List<object>(); at++; Space(text, ref at);
                if (at < text.Length && text[at] == ']') { at++; return list; }
                while (true)
                {
                    list.Add(Value(text, ref at)); Space(text, ref at);
                    if (at < text.Length && text[at] == ',') { at++; continue; }
                    if (at < text.Length && text[at] == ']') { at++; return list; }
                    throw Bad(at);
                }
            }
            if (c == '"') return ReadString(text, ref at);
            if (string.CompareOrdinal(text, at, "true", 0, 4) == 0) { at += 4; return true; }
            if (string.CompareOrdinal(text, at, "false", 0, 5) == 0) { at += 5; return false; }
            if (string.CompareOrdinal(text, at, "null", 0, 4) == 0) { at += 4; return null; }
            int start = at;
            while (at < text.Length && "+-0123456789.eE".IndexOf(text[at]) >= 0) at++;
            if (at == start) throw Bad(at);
            string number = text.Substring(start, at - start); long whole; double real;
            if (long.TryParse(number, NumberStyles.Integer, CultureInfo.InvariantCulture, out whole)) return whole;
            if (double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out real)) return real;
            throw Bad(start);
        }
        private static string ReadString(string text, ref int at)
        {
            var result = new StringBuilder(); at++;
            while (at < text.Length)
            {
                char c = text[at++];
                if (c == '"') return result.ToString();
                if (c != '\\') { result.Append(c); continue; }
                if (at >= text.Length) break;
                char e = text[at++];
                if (e == 'n') result.Append('\n'); else if (e == 'r') result.Append('\r'); else if (e == 't') result.Append('\t');
                else if (e == 'b') result.Append('\b'); else if (e == 'f') result.Append('\f');
                else if (e == 'u')
                {
                    if (at + 4 > text.Length) throw Bad(at);
                    result.Append((char)int.Parse(text.Substring(at, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture)); at += 4;
                }
                else result.Append(e);
            }
            throw Bad(at);
        }
        // Typed reads with defaults, so an older or hand-edited preset with a
        // missing field still loads.
        internal static string Str(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture) : ""; }
        internal static bool Flag(IDictionary<string, object> map, string key, bool fallback) { object v; return map != null && map.TryGetValue(key, out v) && v is bool ? (bool)v : fallback; }
        internal static long Num(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) && (v is long || v is double) ? Convert.ToInt64(v) : 0; }
        internal static Dictionary<string, object> Map(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) ? v as Dictionary<string, object> : null; }
        internal static List<Dictionary<string, object>> Maps(IDictionary<string, object> map, string key)
        { object v; var list = map != null && map.TryGetValue(key, out v) ? v as List<object> : null; return list == null ? new List<Dictionary<string, object>>() : list.OfType<Dictionary<string, object>>().ToList(); }
        internal static Dictionary<string, object> Obj(params object[] pairs)
        { var map = new Dictionary<string, object>(StringComparer.Ordinal); for (int i = 0; i + 1 < pairs.Length; i += 2) map[(string)pairs[i]] = pairs[i + 1]; return map; }
    }

    // A saved migration setup (settings only, never sheet data), kept as one
    // .xymigration.json file. The local library is a folder of these files;
    // export copies one out, import copies one in.
    internal sealed class MigrationPreset
    {
        internal const string Format = "xycabin.migration-preset", Extension = ".xymigration.json";
        internal const int Version = 1;
        internal string FilePath = "", Name = "";
        internal DateTime Saved;
        internal Dictionary<string, object> Data;
        internal int Sources, Targets, Mappings;
        internal static string Library
        { get { return System.IO.Path.Combine(DialogSizeStore.RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Connect).Assembly.GetName().Name, "迁移预设"); } }
        internal static MigrationPreset From(Dictionary<string, object> data, string file)
        {
            if (data == null || MigrationJson.Str(data, "format") != Format) throw new InvalidOperationException("不是批量迁移数据的预设文件");
            long version = MigrationJson.Num(data, "version");
            if (version < 1) throw new InvalidOperationException("预设文件缺少版本号");
            if (version > Version) throw new InvalidOperationException("预设由更新版本的加载项创建（格式版本 " + version + "），请先升级加载项");
            var preset = new MigrationPreset { FilePath = file ?? "", Data = data, Name = MigrationJson.Str(data, "name") };
            if (preset.Name.Length == 0) preset.Name = System.IO.Path.GetFileName(file ?? "未命名").Replace(Extension, "");
            DateTime saved; if (DateTime.TryParse(MigrationJson.Str(data, "saved"), CultureInfo.InvariantCulture, DateTimeStyles.None, out saved)) preset.Saved = saved;
            preset.Sources = MigrationJson.Maps(data, "sources").Count; preset.Targets = MigrationJson.Maps(data, "targets").Count; preset.Mappings = MigrationJson.Maps(data, "mappings").Count;
            return preset;
        }
        internal static MigrationPreset Read(string file)
        {
            string text;
            try { text = System.IO.File.ReadAllText(file, Encoding.UTF8); } catch (Exception error) { throw new InvalidOperationException("读取预设失败：" + error.Message, error); }
            object data;
            try { data = MigrationJson.Parse(text); } catch (FormatException error) { throw new InvalidOperationException("预设文件已损坏：" + error.Message, error); }
            return From(data as Dictionary<string, object>, file);
        }
        // Written beside the target, then swapped in, so a failed write never
        // leaves half a preset.
        internal static void Write(string file, Dictionary<string, object> data)
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file));
            string temporary = file + ".tmp";
            System.IO.File.WriteAllText(temporary, MigrationJson.Write(data), new UTF8Encoding(false));
            if (System.IO.File.Exists(file)) System.IO.File.Delete(file);
            System.IO.File.Move(temporary, file);
        }
        internal static List<MigrationPreset> List()
        {
            var result = new List<MigrationPreset>();
            if (!Directory.Exists(Library)) return result;
            foreach (string file in Directory.GetFiles(Library, "*" + Extension))
                try { result.Add(Read(file)); } catch (Exception error) { UiDiagnostics.Record("migration-preset-list", error); }
            return result.OrderByDescending(p => p.Saved).ThenBy(p => p.Name, StringComparer.CurrentCulture).ToList();
        }
        internal static string FileFor(string name)
        {
            var safe = new StringBuilder();
            foreach (char c in name.Trim()) safe.Append(System.IO.Path.GetInvalidFileNameChars().Contains(c) ? '_' : c);
            return System.IO.Path.Combine(Library, (safe.Length == 0 ? "未命名" : safe.ToString()) + Extension);
        }
        internal static string UniqueFile(string name)
        {
            string file = FileFor(name); int n = 2;
            while (System.IO.File.Exists(file)) file = FileFor(name + " (" + n++ + ")");
            return file;
        }
        // Every workbook the preset refers to (lists and 逐行重算 choices).
        internal IEnumerable<string> Files()
        {
            var paths = new List<string>();
            foreach (string list in new[] { "sourceBooks", "targetBooks" }) foreach (var book in MigrationJson.Maps(Data, list)) paths.Add(MigrationJson.Str(book, "path"));
            foreach (string list in new[] { "sources", "targets" })
                foreach (var side in MigrationJson.Maps(Data, list))
                {
                    foreach (var book in MigrationJson.Maps(side, "books")) paths.Add(MigrationJson.Str(book, "path"));
                    foreach (var chosen in MigrationJson.Maps(MigrationJson.Map(side, "trigger"), "chosen")) paths.Add(MigrationJson.Str(chosen, "path"));
                }
            return paths.Where(p => p.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase);
        }
        internal string Relative(string path)
        {
            foreach (string list in new[] { "sourceBooks", "targetBooks" })
                foreach (var book in MigrationJson.Maps(Data, list)) if (string.Equals(MigrationJson.Str(book, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(book, "relative");
            foreach (string list in new[] { "sources", "targets" })
                foreach (var side in MigrationJson.Maps(Data, list))
                {
                    foreach (var book in MigrationJson.Maps(side, "books")) if (string.Equals(MigrationJson.Str(book, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(book, "relative");
                    foreach (var chosen in MigrationJson.Maps(MigrationJson.Map(side, "trigger"), "chosen")) if (string.Equals(MigrationJson.Str(chosen, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(chosen, "relative");
                }
            return System.IO.Path.GetFileName(path);
        }
        // The manager's detail pane, as styled lines: title, a dim meta line,
        // then sections (保存方式 / 映射 / 工作簿) with blue headings,
        // numbered mapping blocks and ✓/⚠ workbook lines.
        internal enum DetailKind { Title, Meta, Heading, Number, Label, Text, Dim, Tag, Warn, Ok }
        internal sealed class DetailLine
        {
            // 0: flush left; 1: section body; 2: under a mapping number.
            internal int Indent;
            // Leave a little room above (a new section or mapping).
            internal bool Space;
            internal readonly List<KeyValuePair<DetailKind, string>> Parts = new List<KeyValuePair<DetailKind, string>>();
            internal DetailLine Add(DetailKind kind, string text) { if (!string.IsNullOrEmpty(text)) Parts.Add(new KeyValuePair<DetailKind, string>(kind, text)); return this; }
            internal string Text { get { return string.Concat(Parts.Select(p => p.Value).ToArray()); } }
        }
        private static string Circled(int index) { return index >= 1 && index <= 20 ? ((char)(0x2460 + index - 1)).ToString() : index + "."; }
        internal List<DetailLine> DescribeLines()
        {
            var lines = new List<DetailLine>();
            Func<bool, int, DetailLine> line = delegate(bool space, int indent) { var item = new DetailLine { Space = space, Indent = indent }; lines.Add(item); return item; };
            line(false, 0).Add(DetailKind.Title, Name);
            line(false, 0).Add(DetailKind.Meta, "保存于 " + (Saved == DateTime.MinValue ? "未知" : Saved.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)) + "　·　源区域 " + Sources + "　·　目标区域 " + Targets + "　·　映射 " + Mappings);
            var options = MigrationJson.Map(Data, "options");
            line(true, 0).Add(DetailKind.Heading, "保存方式");
            line(false, 1).Add(DetailKind.Text, MigrationJson.Flag(options, "replaceTargets", true) ? "更新目标原文件（自动备份）" : "另存结果副本")
                .Add(DetailKind.Dim, MigrationJson.Flag(options, "allowEmptySource", false) ? "　·　允许全空白源区域" : "");
            Func<Dictionary<string, object>, string> books = delegate(Dictionary<string, object> side)
            {
                var names = MigrationJson.Maps(side, "books").Select(b => System.IO.Path.GetFileName(MigrationJson.Str(b, "path"))).ToList();
                return names.Count == 0 ? "（未选工作簿）" : names.Count == 1 ? names[0] : names[0] + " 等 " + names.Count + " 个";
            };
            Func<Dictionary<string, object>, string> place = delegate(Dictionary<string, object> side)
            { string sheet = MigrationJson.Str(side, "sheet"), range = MigrationJson.Str(side, "range"); return (sheet.Length == 0 ? "未选工作表" : sheet) + "!" + (range.Length == 0 ? "未设区域" : range); };
            var sources = MigrationJson.Maps(Data, "sources").ToDictionary(s => MigrationJson.Str(s, "id"), s => s);
            var targets = MigrationJson.Maps(Data, "targets").ToDictionary(s => MigrationJson.Str(s, "id"), s => s);
            var mappings = MigrationJson.Maps(Data, "mappings").ToList();
            line(true, 0).Add(DetailKind.Heading, "映射（" + mappings.Count + "）");
            if (mappings.Count == 0) line(false, 1).Add(DetailKind.Dim, "（没有映射）");
            int index = 0;
            foreach (var mapping in mappings)
            {
                index++;
                Dictionary<string, object> source, target;
                sources.TryGetValue(MigrationJson.Str(mapping, "source"), out source); targets.TryGetValue(MigrationJson.Str(mapping, "target"), out target);
                var trigger = source == null ? null : MigrationJson.Map(source, "trigger");
                var head = line(index > 1, 1).Add(DetailKind.Number, Circled(index) + " ").Add(DetailKind.Label, "源　　");
                if (source == null) head.Add(DetailKind.Dim, "（未配对）"); else head.Add(DetailKind.Text, place(source)).Add(DetailKind.Dim, "　" + books(source));
                var to = line(false, 2).Add(DetailKind.Label, "目标　");
                if (target == null) to.Add(DetailKind.Dim, "（未配对）"); else to.Add(DetailKind.Text, place(target)).Add(DetailKind.Dim, "　" + (trigger != null ? "按文件名匹配" : books(target)));
                var tags = line(false, 2).Add(DetailKind.Tag, trigger != null || MigrationJson.Flag(mapping, "valuesOnly", false) ? "只写数值" : "带公式");
                if (!MigrationJson.Flag(mapping, "enabled", true)) tags.Add(DetailKind.Dim, "　·　").Add(DetailKind.Warn, "未启用");
                if (trigger != null) tags.Add(DetailKind.Dim, "　·　").Add(DetailKind.Tag, "逐行重算 ")
                    .Add(DetailKind.Dim, MigrationJson.Str(trigger, "values") + " → " + MigrationJson.Str(trigger, "cell") + "，文件名列 " + MigrationJson.Str(trigger, "nameColumn"));
            }
            var files = Files().ToList(); int missing = files.Count(f => !System.IO.File.Exists(f));
            line(true, 0).Add(DetailKind.Heading, "工作簿（" + files.Count + "）")
                .Add(missing > 0 ? DetailKind.Warn : DetailKind.Ok, missing > 0 ? "　" + missing + " 个在本机找不到，应用时可重新定位" : "　均可找到");
            foreach (string file in files)
            {
                bool found = System.IO.File.Exists(file);
                line(false, 1).Add(found ? DetailKind.Ok : DetailKind.Warn, found ? "✓  " : "⚠  ").Add(found ? DetailKind.Text : DetailKind.Warn, System.IO.Path.GetFileName(file))
                    .Add(DetailKind.Dim, "　" + System.IO.Path.GetDirectoryName(file));
            }
            return lines;
        }
        // The same content as plain text.
        internal string Describe()
        { return string.Join(Environment.NewLine, DescribeLines().Select(l => l.Text).ToArray()); }
    }

    // One clickable line of the settings popup: title at TextLeft, the grey
    // note at NoteLeft, light blue rounded hover.
    internal sealed class MigrationMenuRow : Control
    {
        internal string Title = "", Note = "";
        internal int TextLeft, NoteLeft;
        private bool hover;
        internal MigrationMenuRow()
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

    // Two-option switch used in the settings popup.
    internal sealed class MigrationSegmented : Control
    {
        private readonly string[] options;
        private int selected, hover = -1;
        internal event Action SelectedChanged;
        internal MigrationSegmented(string left, string right, int value)
        {
            options = new[] { left, right }; selected = value; Cursor = Cursors.Hand; TabStop = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
        }
        internal int Selected { get { return selected; } }
        private Rectangle Half(int index) { int w = Width / 2; return new Rectangle(index == 0 ? 0 : w, 0, index == 0 ? w : Width - w, Height); }
        protected override void OnMouseMove(MouseEventArgs e) { int h = e.X < Width / 2 ? 0 : 1; if (h != hover) { hover = h; Invalidate(); } base.OnMouseMove(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = -1; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e); int value = e.X < Width / 2 ? 0 : 1;
            if (value != selected) { selected = value; Invalidate(); if (SelectedChanged != null) SelectedChanged(); }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.Clear(BackColor); g.SmoothingMode = SmoothingMode.AntiAlias; float u = g.DpiX / 96f;
            var outer = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            using (var path = Ui.RoundRect(outer, 6 * u)) using (var pen = new Pen(Color.FromArgb(190, 205, 235))) g.DrawPath(pen, path);
            for (int i = 0; i < 2; i++)
            {
                var half = Half(i); half.Inflate(-(int)(2 * u), -(int)(2 * u));
                if (i == selected) using (var path = Ui.RoundRect(half, 5 * u)) using (var brush = new SolidBrush(Ui.Blue)) g.FillPath(brush, path);
                else if (i == hover) using (var path = Ui.RoundRect(half, 5 * u)) using (var brush = new SolidBrush(Ui.Btn2Hover)) g.FillPath(brush, path);
                TextRenderer.DrawText(g, options[i], Font, half, i == selected ? Color.White : Ui.Ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            }
        }
    }

    // Asks for a preset name.
    internal sealed class MigrationNameDialog : Form
    {
        internal readonly TextBox Input = new TextBox { Name = "migrationPresetName" };
        internal MigrationNameDialog(string title, string prompt, string value)
        {
            Text = title; DialogSizing.Prepare(this, "migration.preset.name", "table", Ui.Blue);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false; ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(18, 16, 18, 14), BackColor = Color.White };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.Controls.Add(new Label { Text = prompt, AutoSize = true, ForeColor = Ui.Ink, Margin = new Padding(0, 0, 0, 8) }, 0, 0);
            Input.Text = value; Input.Dock = DockStyle.Top; Input.Margin = new Padding(0, 0, 0, 12); layout.Controls.Add(Input, 0, 1);
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, BackColor = Color.White, Margin = Padding.Empty };
            var ok = Ui.Button("确定", 90); Ui.StylePrimary(ok); var cancel = Ui.Button("取消", 90); Ui.StyleSecondary(cancel);
            ok.Margin = new Padding(8, 0, 0, 0); cancel.Margin = new Padding(8, 0, 0, 0);
            ok.Click += delegate { if (Input.Text.Trim().Length == 0) { Input.Focus(); return; } DialogResult = DialogResult.OK; };
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            bar.Controls.Add(ok); bar.Controls.Add(cancel); layout.Controls.Add(bar, 0, 2);
            Controls.Add(layout); AcceptButton = ok; CancelButton = cancel;
            Shown += delegate { Input.Focus(); Input.SelectAll(); };
        }
        internal static string Ask(IWin32Window owner, string title, string prompt, string value)
        {
            return HostDialogService.ShowIsolatedResult<string>(owner,
                delegate { return new MigrationNameDialog(title, prompt, value); },
                delegate(Form dialog) { return ((MigrationNameDialog)dialog).Input.Text.Trim(); }, null);
        }
    }

    // Workbooks a preset names that this computer does not have: pick a
    // folder once and every missing file found there (by its path relative to
    // the preset's common folder, or by name) is relocated; any left can be
    // pointed to one by one.
    internal sealed class MigrationRelocateDialog : Form
    {
        private readonly DataGridView grid = new DataGridView { Name = "migrationRelocate", Dock = DockStyle.Fill };
        private readonly Label summary = new Label { AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(0, 8, 0, 0) };
        private readonly MigrationPreset preset;
        internal readonly Dictionary<string, string> Found = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<string> missing;
        internal MigrationRelocateDialog(MigrationPreset preset, List<string> missing)
        {
            this.preset = preset; this.missing = missing;
            Text = "定位缺失的工作簿"; DialogSizing.Prepare(this, "migration.preset.relocate", "table", Ui.Blue);
            ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent; MinimizeBox = false;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16), BackColor = Ui.Canvas };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = "预设中有 " + missing.Count + " 个工作簿在本机原位置找不到。选择存放它们的文件夹，程序会按相同的相对位置或文件名自动找回；找不到的可以逐个指定。", AutoSize = true, MaximumSize = new Size(820, 0), ForeColor = Ui.Ink, Margin = new Padding(0, 0, 0, 10) }, 0, 0);
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false; grid.ReadOnly = true;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = false;
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "name", HeaderText = "工作簿", FillWeight = 30, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "old", HeaderText = "原位置", FillWeight = 35, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "now", HeaderText = "本机位置", FillWeight = 35, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            Ui.StyleGrid(grid); layout.Controls.Add(grid, 0, 1);
            layout.Controls.Add(summary, 0, 2);
            var bar = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2, Margin = new Padding(0, 10, 0, 0), BackColor = Ui.Canvas };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            var left = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty, BackColor = Ui.Canvas };
            var right = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = Padding.Empty, BackColor = Ui.Canvas };
            var folder = Ui.Button("选择文件夹…", 130); Ui.StyleSecondary(folder);
            var one = Ui.Button("指定所选文件…", 140); Ui.StyleSecondary(one);
            var apply = Ui.Button("应用预设", 110); Ui.StylePrimary(apply);
            var cancel = Ui.Button("取消", 90); Ui.StyleSecondary(cancel);
            folder.Margin = new Padding(0, 0, 8, 0); one.Margin = new Padding(0, 0, 8, 0); apply.Margin = new Padding(8, 0, 0, 0); cancel.Margin = new Padding(8, 0, 0, 0);
            left.Controls.Add(folder); left.Controls.Add(one); right.Controls.Add(apply); right.Controls.Add(cancel);
            bar.Controls.Add(left, 0, 0); bar.Controls.Add(right, 1, 0); layout.Controls.Add(bar, 0, 3);
            Controls.Add(layout); CancelButton = cancel;
            folder.Click += delegate { PickFolder(); };
            one.Click += delegate { PickOne(); };
            grid.CellDoubleClick += delegate { PickOne(); };
            apply.Click += delegate
            {
                int left2 = missing.Count(m => !Found.ContainsKey(m));
                if (left2 > 0 && UiDialog.Show(this, "仍有 " + left2 + " 个工作簿没有找到。继续应用时它们会保留原路径并标为“找不到”，相关区域需要你稍后重新选择工作簿。" + Environment.NewLine + Environment.NewLine + "是否继续应用？",
                    "仍有缺失", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                DialogResult = DialogResult.OK;
            };
            cancel.Click += delegate { DialogResult = DialogResult.Cancel; };
            Fill();
        }
        private void Fill()
        {
            grid.Rows.Clear();
            foreach (string path in missing)
            {
                string now; bool found = Found.TryGetValue(path, out now);
                int index = grid.Rows.Add(Path.GetFileName(path), path, found ? now : "（未找到）"); grid.Rows[index].Tag = path;
                if (!found) grid.Rows[index].Cells["now"].Style.ForeColor = Ui.Danger;
            }
            int done = missing.Count(m => Found.ContainsKey(m));
            summary.Text = "已找到 " + done + " / " + missing.Count + " 个。";
        }
        private void PickFolder()
        {
            string folder = HostDialogService.Folder(this, "选择存放这些工作簿的文件夹", false);
            if (folder == null) return;
            int before = Found.Count;
            foreach (string path in missing)
            {
                if (Found.ContainsKey(path)) continue;
                string relative = preset.Relative(path), name = Path.GetFileName(path);
                foreach (string candidate in new[] { Path.Combine(folder, relative), Path.Combine(folder, name) })
                    if (File.Exists(candidate)) { Found[path] = Path.GetFullPath(candidate); break; }
            }
            Fill();
            if (Found.Count == before) UiDialog.Show(this, "这个文件夹里没有找到缺失的工作簿。", "定位缺失的工作簿", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }
        private void PickOne()
        {
            if (grid.CurrentRow == null) return; string path = (string)grid.CurrentRow.Tag;
            string[] files = HostDialogService.OpenFiles(this, DataMigrationForm.FileFilter, "指定 " + Path.GetFileName(path), false, Path.GetFileName(path));
            if (files.Length == 0) return;
            Found[path] = Path.GetFullPath(files[0]); Fill();
        }
    }

    // 预设管理: the saved presets on the left (scrolls), the full setup of
    // the selected one on the right; import, export, rename, delete, apply.
    // Applying only restores the settings; nothing runs until 预检.
    internal sealed class MigrationPresetManager : Form
    {
        private readonly ListBox list = new ListBox { Name = "migrationPresetList", Dock = DockStyle.Fill, DrawMode = DrawMode.OwnerDrawFixed, BorderStyle = BorderStyle.None, IntegralHeight = false };
        private readonly RichTextBox detail = new RichTextBox { Name = "migrationPresetDetail", Dock = DockStyle.Fill, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Color.White, WordWrap = true };
        private readonly TextBox search = new TextBox { Name = "migrationPresetSearch", Width = 220 };
        private readonly RoundedButton apply = Ui.Button("应用到窗口", 130), rename = Ui.Button("重命名", 90), remove = Ui.Button("删除", 80), export = Ui.Button("导出…", 90);
        private List<MigrationPreset> all = new List<MigrationPreset>();
        private Font bold, title, meta, gap;
        internal MigrationPreset Chosen;
        internal MigrationPresetManager(string select)
        {
            Text = "预设管理"; DialogSizing.Prepare(this, "migration.preset.manager", "table", Ui.Blue);
            ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent; MinimizeBox = false;
            bold = Ui.MakeFont(10f, FontStyle.Bold, Font); title = Ui.MakeFont(12f, FontStyle.Bold, Font); meta = Ui.MakeFont(9f, FontStyle.Regular, Font); gap = Ui.MakeFont(5f, FontStyle.Regular, Font);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(16), BackColor = Ui.Canvas };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            // Top: import on the left, search on the right.
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, Margin = new Padding(0, 0, 0, 10), BackColor = Ui.Canvas };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var import = Ui.Button("导入预设…", 120); Ui.StyleSecondary(import); import.Margin = Padding.Empty;
            var hint = new Label { Text = "预设只保存设置，不含表格数据；应用后请检查并预检。", AutoSize = true, ForeColor = Ui.InkDim, Anchor = AnchorStyles.Left, Margin = new Padding(12, 0, 0, 0) };
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
            var close = Ui.Button("关闭", 90); Ui.StyleSecondary(close);
            Ui.StyleSecondary(rename); Ui.StyleSecondary(remove); Ui.StyleSecondary(export); Ui.StylePrimary(apply);
            foreach (var b in new Control[] { rename, remove, export }) { b.Margin = new Padding(0, 0, 8, 0); leftBar.Controls.Add(b); }
            apply.Margin = new Padding(8, 0, 0, 0); close.Margin = new Padding(8, 0, 0, 0);
            rightBar.Controls.Add(close); rightBar.Controls.Add(apply);
            bottom.Controls.Add(leftBar, 0, 0); bottom.Controls.Add(rightBar, 1, 0);
            layout.Controls.Add(bottom, 0, 2);
            Controls.Add(layout); CancelButton = close;

            list.HandleCreated += delegate { list.ItemHeight = Math.Max(Ui.Px(list, 52), Font.Height * 2 + Ui.Px(list, 16)); };
            list.DrawItem += DrawPreset;
            list.SelectedIndexChanged += delegate { ShowDetail(); };
            list.DoubleClick += delegate { if (Current != null) Apply(); };
            list.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Delete) Remove(); else if (e.KeyCode == Keys.F2) Rename(); };
            search.HandleCreated += delegate { MigrationPresetManagerNative.Cue(search, "搜索预设名称"); };
            detail.HandleCreated += delegate { BeginInvoke(new Action(ShowDetail)); };
            search.TextChanged += delegate { Fill(Current == null ? null : Current.FilePath); };
            import.Click += delegate { Import(); }; export.Click += delegate { Export(); };
            rename.Click += delegate { Rename(); }; remove.Click += delegate { Remove(); };
            apply.Click += delegate { Apply(); }; close.Click += delegate { DialogResult = DialogResult.Cancel; };
            Disposed += delegate { bold.Dispose(); title.Dispose(); meta.Dispose(); gap.Dispose(); };
            Reload(select);
        }
        private static Control Card(Control content, Padding margin, Padding inner)
        {
            var card = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Padding = inner, Margin = margin };
            card.Paint += delegate(object sender, PaintEventArgs e) { using (var pen = new Pen(Ui.GridFrame)) e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1); };
            card.Resize += delegate { card.Invalidate(); };
            content.Dock = DockStyle.Fill; card.Controls.Add(content); return card;
        }
        private MigrationPreset Current { get { return list.SelectedItem as MigrationPreset; } }
        private void Reload(string select)
        {
            // Layout previews must not read the user's preset library.
            all = DialogPreviewFactory.IsConstructing ? new List<MigrationPreset>() : MigrationPreset.List(); Fill(select);
        }
        private void Fill(string select)
        {
            string term = search.Text.Trim();
            list.BeginUpdate();
            try
            {
                list.Items.Clear();
                foreach (var preset in all) if (term.Length == 0 || preset.Name.IndexOf(term, StringComparison.CurrentCultureIgnoreCase) >= 0) list.Items.Add(preset);
                int index = -1;
                for (int i = 0; i < list.Items.Count; i++) if (string.Equals(((MigrationPreset)list.Items[i]).FilePath, select, StringComparison.OrdinalIgnoreCase)) index = i;
                if (index < 0 && list.Items.Count > 0) index = 0;
                list.SelectedIndex = index;
            }
            finally { list.EndUpdate(); }
            ShowDetail();
        }
        private void ShowDetail()
        {
            var preset = Current; bool has = preset != null;
            apply.Enabled = rename.Enabled = remove.Enabled = export.Enabled = has;
            if (!has)
            {
                SetDetail(all.Count == 0 ? "还没有保存的预设。" + Environment.NewLine + Environment.NewLine + "在批量迁移窗口右上角的齿轮里点“保存为预设…”，或点上方“导入预设…”导入别人分享的预设文件。" : "没有名称匹配的预设。");
                return;
            }
            ShowLines(preset.DescribeLines());
        }
        // Each part gets its own font and color; sections and mappings are
        // separated by a short gap, bodies are indented under their heading.
        private void ShowLines(List<MigrationPreset.DetailLine> lines)
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
                            case MigrationPreset.DetailKind.Title: font = title; color = Ui.Ink; break;
                            case MigrationPreset.DetailKind.Meta: font = meta; color = Ui.InkDim; break;
                            case MigrationPreset.DetailKind.Heading: font = bold; color = Ui.Blue; break;
                            case MigrationPreset.DetailKind.Number:
                                font = bold; color = Ui.Blue;
                                number = TextRenderer.MeasureText(part.Value, bold, Size.Empty, TextFormatFlags.NoPadding).Width;
                                break;
                            case MigrationPreset.DetailKind.Label: case MigrationPreset.DetailKind.Dim: font = list.Font; color = Ui.InkDim; break;
                            case MigrationPreset.DetailKind.Tag: font = list.Font; color = Ui.Blue; break;
                            case MigrationPreset.DetailKind.Warn: font = list.Font; color = warn; break;
                            case MigrationPreset.DetailKind.Ok: font = list.Font; color = Ui.Green; break;
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
        private void SetDetail(string text)
        {
            detail.Clear(); detail.SelectionFont = list.Font; detail.SelectionColor = Ui.Ink; detail.SelectionIndent = 0;
            detail.Text = text;
            detail.SelectAll(); detail.SelectionFont = list.Font; detail.SelectionColor = Ui.Ink; detail.Select(0, 0);
        }
        private void DrawPreset(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= list.Items.Count) return;
            var preset = (MigrationPreset)list.Items[e.Index]; var g = e.Graphics; var r = e.Bounds;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            using (var back = new SolidBrush(selected ? Ui.SelectBack : Color.White)) g.FillRectangle(back, r);
            if (selected) using (var bar = new SolidBrush(Ui.Blue)) g.FillRectangle(bar, r.Left, r.Top + 6, 3, r.Height - 12);
            using (var line = new Pen(Ui.GridLine)) g.DrawLine(line, r.Left + 12, r.Bottom - 1, r.Right - 12, r.Bottom - 1);
            var flags = TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
            int pad = Ui.Px(list, 14), half = r.Height / 2;
            TextRenderer.DrawText(g, preset.Name, bold, new Rectangle(r.Left + pad, r.Top + 4, r.Width - pad * 2, half - 2), selected ? Ui.SelectInk : Ui.Ink, flags | TextFormatFlags.Bottom);
            string info = (preset.Saved == DateTime.MinValue ? "" : preset.Saved.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture) + "　·　") + "源 " + preset.Sources + "　目标 " + preset.Targets + "　映射 " + preset.Mappings;
            TextRenderer.DrawText(g, info, list.Font, new Rectangle(r.Left + pad, r.Top + half + 1, r.Width - pad * 2, half - 4), Ui.InkDim, flags | TextFormatFlags.Top);
        }
        private void Import()
        {
            string[] files = HostDialogService.OpenFiles(this, "迁移预设 (*" + MigrationPreset.Extension + ")|*" + MigrationPreset.Extension + "|JSON 文件 (*.json)|*.json", "导入预设", true, "");
            if (files.Length == 0) return;
            var failures = new List<string>(); string last = null;
            foreach (string file in files)
                try
                {
                    var preset = MigrationPreset.Read(file);
                    string target = MigrationPreset.UniqueFile(preset.Name);
                    preset.Data["name"] = Path.GetFileName(target).Replace(MigrationPreset.Extension, "");
                    MigrationPreset.Write(target, preset.Data); last = target;
                }
                catch (Exception error) { failures.Add(Path.GetFileName(file) + "：" + error.Message); }
            Reload(last);
            if (failures.Count > 0) UiDialog.Show(this, "以下文件未能导入：" + Environment.NewLine + string.Join(Environment.NewLine, failures.ToArray()), "导入预设", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
        }
        private void Export()
        {
            var preset = Current; if (preset == null) return;
            string chosen = HostDialogService.SaveFile(this, "导出预设", "迁移预设 (*" + MigrationPreset.Extension + ")|*" + MigrationPreset.Extension, Path.GetFileName(preset.FilePath));
            if (chosen == null) return;
            string file = chosen.EndsWith(MigrationPreset.Extension, StringComparison.OrdinalIgnoreCase) ? chosen : chosen + MigrationPreset.Extension;
            try { MigrationPreset.Write(file, preset.Data); UiDialog.Show(this, "已导出到：" + Environment.NewLine + file, "导出预设", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1); }
            catch (Exception error) { UiDialog.Show(this, "导出失败：" + error.Message, "导出预设", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); }
        }
        private void Rename()
        {
            var preset = Current; if (preset == null) return;
            string name = MigrationNameDialog.Ask(this, "重命名预设", "新的预设名称：", preset.Name);
            if (name == null || name == preset.Name) return;
            string target = MigrationPreset.FileFor(name);
            if (!string.Equals(target, preset.FilePath, StringComparison.OrdinalIgnoreCase) && File.Exists(target))
            { UiDialog.Show(this, "已有同名预设“" + name + "”。", "重命名预设", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1); return; }
            try
            {
                preset.Data["name"] = name; MigrationPreset.Write(target, preset.Data);
                if (!string.Equals(target, preset.FilePath, StringComparison.OrdinalIgnoreCase)) File.Delete(preset.FilePath);
                Reload(target);
            }
            catch (Exception error) { UiDialog.Show(this, "重命名失败：" + error.Message, "重命名预设", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); }
        }
        private void Remove()
        {
            var preset = Current; if (preset == null) return;
            if (UiDialog.Show(this, "删除预设“" + preset.Name + "”？" + Environment.NewLine + "只删除预设文件，不影响任何工作簿。", "删除预设", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            try { File.Delete(preset.FilePath); Reload(null); }
            catch (Exception error) { UiDialog.Show(this, "删除失败：" + error.Message, "删除预设", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); }
        }
        private void Apply() { if (Current == null) return; Chosen = Current; DialogResult = DialogResult.OK; }
    }
    internal static class MigrationPresetManagerNative
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        internal static void Cue(TextBox box, string text) { try { SendMessage(box.Handle, 0x1501, (IntPtr)1, text); } catch (Exception) { } }
    }

    // Read-only snapshot of one worksheet, read straight from the .xlsx/.xlsm
    // package: no WPS window opens and nothing is written. Only text needed to
    // recognise a position is kept (values, merges, row/column sizes).
    internal sealed class MigrationSheetSnapshot
    {
        internal string Path = "", Sheet = "";
        internal DateTime Stamp;
        internal int MaxRow, MaxColumn;
        internal readonly Dictionary<long, string> Values = new Dictionary<long, string>();
        internal readonly List<MigrationArea> Merges = new List<MigrationArea>();
        internal readonly Dictionary<int, double> Widths = new Dictionary<int, double>();
        internal readonly Dictionary<int, double> Heights = new Dictionary<int, double>();
        private List<long> ordered;
        internal const int CellLimit = 600000;
        internal static long Key(int row, int column) { return (long)row * 20000 + column; }
        internal string Value(int row, int column) { string value; return Values.TryGetValue(Key(row, column), out value) ? value : ""; }
        // Next cell (row-major) after the given one whose text contains term.
        internal bool Find(string term, int afterRow, int afterColumn, out int row, out int column)
        {
            row = column = 0;
            if (string.IsNullOrEmpty(term)) return false;
            if (ordered == null) { ordered = Values.Keys.ToList(); ordered.Sort(); }
            long after = Key(afterRow, afterColumn);
            for (int pass = 0; pass < 2; pass++)
                foreach (long key in ordered)
                {
                    if (pass == 0 ? key <= after : key > after) continue;
                    if (Values[key].IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    row = (int)(key / 20000); column = (int)(key % 20000); return true;
                }
            return false;
        }
        private static readonly string[] MainNs = { "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main" };
        private static readonly string[] RelNs = { "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships" };
        private static XmlReaderSettings Settings() { return new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = false }; }
        private static bool Is(XmlReader xml, string name) { return xml.NodeType == XmlNodeType.Element && xml.LocalName == name && MainNs.Contains(xml.NamespaceURI); }
        internal static MigrationSheetSnapshot Load(string path, string sheet)
        {
            string extension = System.IO.Path.GetExtension(path);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("窗口内预览只支持 .xlsx/.xlsm；" + extension + " 文件请直接填写目标起点，或在 WPS 中确认位置。");
            var snapshot = new MigrationSheetSnapshot { Path = path, Sheet = sheet, Stamp = File.GetLastWriteTimeUtc(path) };
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                string id = null;
                var book = package.GetEntry("xl/workbook.xml"); if (book == null) throw new InvalidOperationException("工作簿缺少工作表目录。");
                using (var xml = XmlReader.Create(book.Open(), Settings()))
                    while (xml.Read())
                        if (Is(xml, "sheet") && string.Equals(xml.GetAttribute("name"), sheet, StringComparison.OrdinalIgnoreCase))
                        { foreach (string ns in RelNs) { id = id ?? xml.GetAttribute("id", ns); } break; }
                if (id == null) throw new InvalidOperationException("找不到工作表“" + sheet + "”。");
                string target = null;
                var rels = package.GetEntry("xl/_rels/workbook.xml.rels"); if (rels == null) throw new InvalidOperationException("工作簿缺少关系目录。");
                using (var xml = XmlReader.Create(rels.Open(), Settings()))
                    while (xml.Read())
                        if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "Relationship" && xml.GetAttribute("Id") == id) { target = xml.GetAttribute("Target"); break; }
                if (string.IsNullOrEmpty(target)) throw new InvalidOperationException("工作表“" + sheet + "”的内容位置无效。");
                target = target.Replace('\\', '/');
                string entryName = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                var shared = ReadShared(package.GetEntry("xl/sharedStrings.xml"));
                var part = package.GetEntry(entryName); if (part == null) throw new InvalidOperationException("工作表“" + sheet + "”的内容缺失。");
                using (var xml = XmlReader.Create(part.Open(), Settings())) snapshot.ReadSheet(xml, shared);
            }
            return snapshot;
        }
        private static List<string> ReadShared(ZipArchiveEntry entry)
        {
            var result = new List<string>(); if (entry == null) return result;
            using (var xml = XmlReader.Create(entry.Open(), Settings()))
            {
                // Advance by hand: ReadElementContentAsString already moves
                // past the element, so an extra Read() would skip </si>.
                var text = new StringBuilder(); bool inside = false;
                while (!xml.EOF)
                {
                    if (Is(xml, "si"))
                    {
                        if (xml.IsEmptyElement) { result.Add(""); xml.Read(); continue; }
                        text.Clear(); inside = true; xml.Read(); continue;
                    }
                    if (Is(xml, "rPh")) { xml.Skip(); continue; } // phonetic guide text
                    if (inside && Is(xml, "t")) { text.Append(xml.ReadElementContentAsString()); continue; }
                    if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "si") { result.Add(text.ToString()); inside = false; }
                    xml.Read();
                }
            }
            return result;
        }
        private void ReadSheet(XmlReader xml, List<string> shared)
        {
            int row = 0, column = 0;
            while (xml.Read())
            {
                if (Is(xml, "row"))
                {
                    int r; row = int.TryParse(xml.GetAttribute("r"), out r) ? r : row + 1; column = 0;
                    double height; if (xml.GetAttribute("customHeight") == "1" && double.TryParse(xml.GetAttribute("ht"), NumberStyles.Float, CultureInfo.InvariantCulture, out height)) Heights[row] = height;
                }
                else if (Is(xml, "c"))
                {
                    string reference = xml.GetAttribute("r"), type = xml.GetAttribute("t");
                    if (!string.IsNullOrEmpty(reference)) { try { var cell = MigrationArea.Parse(reference); row = cell.Row; column = cell.Column; } catch { column++; } }
                    else column++;
                    if (xml.IsEmptyElement) continue;
                    string value = null, inline = null;
                    int depth = xml.Depth;
                    xml.Read();
                    // Stop on this cell's own end tag; never read past it.
                    while (!xml.EOF && !(xml.NodeType == XmlNodeType.EndElement && xml.Depth == depth))
                    {
                        if (Is(xml, "v")) { value = xml.ReadElementContentAsString(); continue; }
                        if (Is(xml, "t")) { inline = (inline ?? "") + xml.ReadElementContentAsString(); continue; }
                        if (Is(xml, "rPh")) { xml.Skip(); continue; }
                        xml.Read();
                    }
                    string text = Display(type, value, inline, shared);
                    if (text.Length == 0) continue;
                    if (Values.Count >= CellLimit) throw new InvalidOperationException("工作表内容过多（超过 " + CellLimit + " 个非空单元格），无法在窗口内预览。");
                    Values[Key(row, column)] = text;
                    MaxRow = Math.Max(MaxRow, row); MaxColumn = Math.Max(MaxColumn, column);
                }
                else if (Is(xml, "col"))
                {
                    int min, max; double width;
                    if (int.TryParse(xml.GetAttribute("min"), out min) && int.TryParse(xml.GetAttribute("max"), out max) && double.TryParse(xml.GetAttribute("width"), NumberStyles.Float, CultureInfo.InvariantCulture, out width))
                        for (int c = min; c <= Math.Min(max, min + 200); c++) Widths[c] = width;
                }
                else if (Is(xml, "mergeCell"))
                {
                    try { var merge = MigrationArea.Parse(xml.GetAttribute("ref")); Merges.Add(merge); MaxRow = Math.Max(MaxRow, merge.Row + merge.Rows - 1); MaxColumn = Math.Max(MaxColumn, merge.Column + merge.Columns - 1); } catch { }
                }
            }
        }
        private static string Display(string type, string value, string inline, List<string> shared)
        {
            if (type == "s") { int index; return value != null && int.TryParse(value, out index) && index >= 0 && index < shared.Count ? shared[index] : ""; }
            if (type == "inlineStr") return inline ?? "";
            if (value == null) return "";
            if (type == "b") return value == "1" ? "TRUE" : "FALSE";
            if (type == "str" || type == "e") return value;
            double number;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ? number.ToString("#,##0.####", CultureInfo.InvariantCulture) : value;
        }
    }

    // Excel-like read-only view of a snapshot. Single click places an area of
    // the source's size with its top-left there; a drag picks a full area.
    // Areas already used by other mappings are tinted; overlap turns red.
    internal sealed class MigrationSheetPreview : UserControl
    {
        private sealed class SheetGrid : DataGridView
        { internal SheetGrid() { DoubleBuffered = true; } }
        private readonly SheetGrid grid = new SheetGrid { Name = "migrationPreviewGrid", Dock = DockStyle.Fill };
        private readonly Label title = new Label { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, ForeColor = Ui.Ink };
        private readonly Label hint = new Label { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoEllipsis = true, ForeColor = Ui.InkDim };
        private readonly Label empty = new Label { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Ui.InkDim, BackColor = Color.White };
        private readonly TextBox address = new TextBox { Name = "migrationPreviewAddress", Width = 64 };
        private readonly TextBox find = new TextBox { Name = "migrationPreviewFind", Width = 104 };
        internal readonly CheckBox Sync = new CheckBox { Name = "migrationPreviewSync", Text = "同步到同源其他目标", AutoSize = true, Checked = true };
        private MigrationSheetSnapshot snapshot;
        private MigrationArea ghost;
        private int sourceRows = 1, sourceColumns = 1;
        private List<KeyValuePair<MigrationArea, string>> others = new List<KeyValuePair<MigrationArea, string>>();
        private float zoom = 1f;
        // Source mode: overlapping other ranges is allowed, so no red tint.
        internal bool OverlapConflict = true;
        private Font baseFont;
        private int foundRow, foundColumn;
        internal event Action<MigrationArea> AreaPicked;
        internal MigrationSheetSnapshot Snapshot { get { return snapshot; } }

        internal MigrationSheetPreview()
        {
            BackColor = Ui.Canvas;
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Margin = Padding.Empty, BackColor = Ui.Canvas };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            layout.Controls.Add(title, 0, 0);
            var bar = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Top, WrapContents = true, Margin = new Padding(0, 0, 0, 4) };
            var go = Ui.Button("跳转", 50); go.Click += delegate { GoTo(); };
            var next = Ui.Button("下一个", 64); next.Click += delegate { FindNext(); };
            var smaller = Ui.Button("−", 30); smaller.Click += delegate { SetZoom(zoom - 0.2f); };
            var larger = Ui.Button("+", 30); larger.Click += delegate { SetZoom(zoom + 0.2f); };
            foreach (var button in new[] { go, next, smaller, larger }) { button.MinimumSize = Size.Empty; button.Height = 28; button.Margin = new Padding(0, 0, 4, 0); }
            bar.Controls.Add(Caption("地址")); bar.Controls.Add(address); bar.Controls.Add(go);
            bar.Controls.Add(Caption("查找")); bar.Controls.Add(find); bar.Controls.Add(next);
            bar.Controls.Add(smaller); bar.Controls.Add(larger); bar.Controls.Add(Sync);
            Sync.Margin = new Padding(8, 5, 0, 0);
            layout.Controls.Add(bar, 0, 1);
            var host = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            host.Controls.Add(grid); host.Controls.Add(empty); empty.BringToFront();
            layout.Controls.Add(host, 0, 2); layout.Controls.Add(hint, 0, 3);
            Controls.Add(layout);
            address.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; GoTo(); } };
            find.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; FindNext(); } };
            find.TextChanged += delegate { foundRow = foundColumn = 0; };

            grid.VirtualMode = true; grid.ReadOnly = true; grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false;
            grid.AllowUserToResizeRows = false; grid.AllowUserToOrderColumns = false; grid.MultiSelect = true;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect; grid.BackgroundColor = Color.White; grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false; grid.GridColor = Color.FromArgb(226, 230, 237);
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Ui.GridHead; grid.ColumnHeadersDefaultCellStyle.ForeColor = Ui.GridHeadInk;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.RowHeadersDefaultCellStyle.BackColor = Ui.GridHead; grid.RowHeadersDefaultCellStyle.ForeColor = Ui.GridHeadInk;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(215, 228, 252); grid.DefaultCellStyle.SelectionForeColor = Ui.Ink;
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            grid.ShowCellToolTips = true;
            grid.CellValueNeeded += delegate(object sender, DataGridViewCellValueEventArgs e) { e.Value = snapshot == null ? "" : snapshot.Value(e.RowIndex + 1, e.ColumnIndex + 1); };
            grid.CellToolTipTextNeeded += delegate(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                string owner = OwnerAt(e.RowIndex + 1, e.ColumnIndex + 1);
                e.ToolTipText = MigrationArea.Cell(e.RowIndex + 1, e.ColumnIndex + 1) + (owner.Length == 0 ? "" : "  ·  已分配给：" + owner);
            };
            grid.RowPostPaint += delegate(object sender, DataGridViewRowPostPaintEventArgs e)
            {
                var bounds = new Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
                TextRenderer.DrawText(e.Graphics, (e.RowIndex + 1).ToString(CultureInfo.InvariantCulture), grid.RowHeadersDefaultCellStyle.Font ?? grid.Font, bounds, Ui.GridHeadInk, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            };
            grid.CellPainting += PaintCell;
            // MouseUp on the grid itself fires wherever a drag ends (also past
            // the last cell); CellMouseUp does not always.
            grid.MouseUp += delegate(object sender, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Left || snapshot == null) return;
                var hit = grid.HitTest(e.X, e.Y);
                if (hit.Type == DataGridViewHitTestType.ColumnHeader || hit.Type == DataGridViewHitTestType.RowHeader || hit.Type == DataGridViewHitTestType.TopLeftHeader) return;
                var cells = grid.SelectedCells.Cast<DataGridViewCell>().ToList(); if (cells.Count == 0) return;
                int top = cells.Min(c => c.RowIndex) + 1, left = cells.Min(c => c.ColumnIndex) + 1, bottom = cells.Max(c => c.RowIndex) + 1, right = cells.Max(c => c.ColumnIndex) + 1;
                var area = new MigrationArea { Row = top, Column = left, Rows = bottom - top + 1, Columns = right - left + 1 };
                grid.ClearSelection();
                if (AreaPicked != null) AreaPicked(area);
            };
            Clear("在左侧选择一条映射，这里会显示它的目标工作表。");
        }
        private static Label Caption(string text) { return new Label { Text = text, AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(2, 7, 2, 0) }; }
        internal void Clear(string text)
        {
            snapshot = null; ghost = null; others.Clear();
            grid.RowCount = 0; grid.ColumnCount = 0;
            empty.Text = text; empty.Visible = true; title.Text = "区域预览"; hint.Text = "";
        }
        internal void SetHint(string text) { hint.Text = text; }
        internal void ShowSheet(MigrationSheetSnapshot sheet, MigrationArea area, int rows, int columns, List<KeyValuePair<MigrationArea, string>> used, string caption, string hintText)
        {
            bool reload = snapshot != sheet;
            snapshot = sheet; ghost = area; sourceRows = Math.Max(1, rows); sourceColumns = Math.Max(1, columns);
            others = used ?? new List<KeyValuePair<MigrationArea, string>>();
            title.Text = caption;
            hint.Text = hintText;
            empty.Visible = false;
            if (reload) Rebuild();
            grid.Invalidate();
            if (reload && ghost != null) Reveal(ghost.Row, ghost.Column);
        }
        private void Rebuild()
        {
            if (baseFont == null) baseFont = Font;
            grid.RowCount = 0; grid.ColumnCount = 0;
            int columns = Math.Min(Math.Max(snapshot.MaxColumn, ghost == null ? 1 : ghost.Column + ghost.Columns - 1) + 6, 400);
            int rows = Math.Min(Math.Max(snapshot.MaxRow, ghost == null ? 1 : ghost.Row + ghost.Rows - 1) + 40, 200000);
            var font = new Font(baseFont.FontFamily, Math.Max(6f, baseFont.SizeInPoints * 0.9f * zoom), GraphicsUnit.Point);
            grid.DefaultCellStyle.Font = font; grid.ColumnHeadersDefaultCellStyle.Font = font; grid.RowHeadersDefaultCellStyle.Font = font;
            grid.RowTemplate.Height = Math.Max(14, (int)Math.Round((font.Height + 6) * 1.0));
            grid.ColumnHeadersHeight = grid.RowTemplate.Height + 2;
            grid.RowHeadersWidth = Math.Max(44, TextRenderer.MeasureText(rows.ToString(CultureInfo.InvariantCulture), font).Width + 16);
            var list = new DataGridViewColumn[columns];
            for (int c = 1; c <= columns; c++)
            {
                double chars; int width = snapshot.Widths.TryGetValue(c, out chars) ? (int)Math.Round(chars * 7 + 5) : 64;
                // FillWeight must stay tiny: the grid caps the total at 65535.
                list[c - 1] = new DataGridViewTextBoxColumn { HeaderText = FormulaRefRewriter.ColumnLetters(c), SortMode = DataGridViewColumnSortMode.NotSortable, FillWeight = 1, Width = Math.Max(24, (int)Math.Round(Math.Min(width, 400) * zoom)), MinimumWidth = 10 };
            }
            grid.Columns.AddRange(list);
            grid.RowCount = rows;
        }
        private void SetZoom(float value)
        {
            value = Math.Max(0.6f, Math.Min(1.8f, value)); if (Math.Abs(value - zoom) < 0.01f || snapshot == null) { zoom = value; return; }
            zoom = value; int row = grid.FirstDisplayedScrollingRowIndex, column = grid.FirstDisplayedScrollingColumnIndex;
            Rebuild();
            if (row >= 0 && row < grid.RowCount) grid.FirstDisplayedScrollingRowIndex = row;
            if (column >= 0 && column < grid.ColumnCount) grid.FirstDisplayedScrollingColumnIndex = column;
        }
        private void Reveal(int row, int column)
        {
            if (snapshot == null || grid.RowCount == 0) return;
            try
            {
                grid.FirstDisplayedScrollingRowIndex = Math.Max(0, Math.Min(grid.RowCount - 1, row - 4));
                grid.FirstDisplayedScrollingColumnIndex = Math.Max(0, Math.Min(grid.ColumnCount - 1, column - 3));
            }
            catch (InvalidOperationException) { } // not yet laid out
        }
        private void GoTo()
        {
            if (snapshot == null) return;
            try
            {
                var cell = MigrationArea.Parse(DataMigrationForm.NormalizeAddress(address.Text));
                if (cell.Row > grid.RowCount || cell.Column > grid.ColumnCount) { hint.Text = "地址超出预览范围（预览到第 " + grid.RowCount + " 行、" + FormulaRefRewriter.ColumnLetters(grid.ColumnCount) + " 列）。"; return; }
                Reveal(cell.Row, cell.Column); grid.ClearSelection(); grid.Rows[cell.Row - 1].Cells[cell.Column - 1].Selected = true;
            }
            catch (Exception error) { hint.Text = error.Message; }
        }
        private void FindNext()
        {
            if (snapshot == null) return;
            int row, column;
            if (!snapshot.Find(find.Text.Trim(), foundRow, foundColumn, out row, out column)) { hint.Text = "没有找到“" + find.Text.Trim() + "”。"; return; }
            foundRow = row; foundColumn = column;
            Reveal(row, column); grid.ClearSelection();
            if (row <= grid.RowCount && column <= grid.ColumnCount) grid.Rows[row - 1].Cells[column - 1].Selected = true;
            hint.Text = "找到 " + MigrationArea.Cell(row, column) + "：" + snapshot.Value(row, column) + "。在它附近单击即可放置。";
        }
        private static bool Inside(MigrationArea area, int row, int column)
        { return area != null && row >= area.Row && row < area.Row + area.Rows && column >= area.Column && column < area.Column + area.Columns; }
        private string OwnerAt(int row, int column)
        {
            foreach (var other in others) if (Inside(other.Key, row, column)) return other.Value;
            return Inside(ghost, row, column) ? "当前映射" : "";
        }
        private void PaintCell(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            int row = e.RowIndex + 1, column = e.ColumnIndex + 1;
            bool mine = Inside(ghost, row, column), taken = others.Any(o => Inside(o.Key, row, column));
            if (!mine && !taken) return;
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All);
            bool clash = mine && taken && OverlapConflict;
            Color tint = clash ? Color.FromArgb(90, 220, 38, 38) : mine ? Color.FromArgb(70, 37, 99, 235) : Color.FromArgb(70, 22, 163, 74);
            using (var brush = new SolidBrush(tint)) e.Graphics.FillRectangle(brush, e.CellBounds);
            if (mine)
            {
                var b = e.CellBounds;
                using (var pen = new Pen(clash ? Color.FromArgb(220, 38, 38) : Ui.Blue, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash })
                {
                    if (row == ghost.Row) e.Graphics.DrawLine(pen, b.Left, b.Top + 1, b.Right, b.Top + 1);
                    if (row == ghost.Row + ghost.Rows - 1) e.Graphics.DrawLine(pen, b.Left, b.Bottom - 2, b.Right, b.Bottom - 2);
                    if (column == ghost.Column) e.Graphics.DrawLine(pen, b.Left + 1, b.Top, b.Left + 1, b.Bottom);
                    if (column == ghost.Column + ghost.Columns - 1) e.Graphics.DrawLine(pen, b.Right - 2, b.Top, b.Right - 2, b.Bottom);
                }
            }
            e.Handled = true;
        }
    }
}
