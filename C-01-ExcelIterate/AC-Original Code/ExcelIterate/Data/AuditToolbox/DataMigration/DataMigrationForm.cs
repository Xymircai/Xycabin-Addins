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
        internal sealed class MatchSummary
        {
            internal int Rows, Matched, Missing;
            internal string Additional = "";
            internal string Heading { get { return "目标文件区域：" + Rows + "行"; } }
            internal string MatchedText { get { return "已匹配" + Matched + "行"; } }
            internal string MissingText { get { return "缺失" + Missing + "行"; } }
            internal string SecondLine { get { return MatchedText + "（" + MissingText + "）。" + Additional; } }
            internal string Text { get { return Heading + "；" + SecondLine; } }
        }
        internal sealed class Entry
        {
            internal string SourceTitle = "", SourceDetail = "", TargetTitle = "", TargetDetail = "", Status = "", SourceTip = "", TargetTip = "", StatusTip = "";
            internal bool SourceUsed, TargetUsed, Enabled = true, ValuesOnly, KeepFormat, FormatOnly;
            // 0 neutral, 1 done, 2 partial, 3 problem, 4 ready
            internal int Tone;
            internal MatchSummary Match;
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
            int pad = P(10), check = Math.Max(P(30), TextRenderer.MeasureText("启用", head).Width + P(14)), number = Math.Max(P(30), TextRenderer.MeasureText("行", head).Width + P(14)), arrow = P(30), mode = P(96), gap = P(10);
            int width = Math.Max(P(480), ViewWidth - pad * 2);
            int rest = width - check - number - arrow - mode - gap * 2;
            int status = Math.Max(P(110), rest * 26 / 100), card = Math.Max(P(130), (rest - status) / 2);
            foreach (var item in items.Where(item => item.Match != null))
            {
                int desired = Math.Max(TextRenderer.MeasureText(item.Match.Heading, body).Width, TextRenderer.MeasureText(item.Match.SecondLine, body).Width) + P(22);
                status = Math.Max(status, Math.Min(desired, Math.Max(P(110), rest - P(260))));
            }
            card = Math.Max(P(130), (rest - status) / 2);
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
            TextRenderer.DrawText(g, MigrationPair.ModeText(values, item.KeepFormat, item.FormatOnly), body, pill, pillInk, Line | TextFormatFlags.HorizontalCenter);
            // 状态: coloured dot + text
            Color dot = item.Tone == 1 ? Ui.Green : item.Tone == 2 ? Ui.Amber : item.Tone == 3 ? Ui.Danger : item.Tone == 4 ? Ui.Blue : Color.FromArgb(170, 178, 192);
            int d = P(8), sx = c.Status.Left + P(4);
            using (var brush = new SolidBrush(dot)) g.FillEllipse(brush, sx, top + (height - d) / 2, d, d);
            var statusBounds = new Rectangle(sx + d + P(6), top, Math.Max(0, c.Status.Right - sx - d - P(6)), height);
            Color statusInk = item.Tone == 3 ? Ui.Danger : item.Enabled ? Ui.Ink : Ui.InkDim;
            if (item.Match != null)
            {
                int lineHeight = TextRenderer.MeasureText("目标", body).Height;
                int firstTop = top + Math.Max(0, (height - lineHeight * 2) / 2);
                TextRenderer.DrawText(g, item.Match.Heading, body,
                    new Rectangle(statusBounds.Left, firstTop, statusBounds.Width, lineHeight), statusInk, CardLine);
                DrawMatchSummary(g, item.Match, new Rectangle(statusBounds.Left, firstTop + lineHeight, statusBounds.Width, lineHeight), statusInk);
            }
            else TextRenderer.DrawText(g, item.Status, body, statusBounds, statusInk, Line | TextFormatFlags.Left);
            g.SmoothingMode = SmoothingMode.Default;
        }
        private void DrawMatchSummary(Graphics g, MatchSummary summary, Rectangle bounds, Color normal)
        {
            const TextFormatFlags flags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.PreserveGraphicsClipping;
            int full = TextRenderer.MeasureText(g, summary.SecondLine, body, Size.Empty, flags).Width;
            bool ellipsis = full > bounds.Width;
            int dots = ellipsis ? TextRenderer.MeasureText(g, "…", body, Size.Empty, flags).Width : 0;
            var saved = g.Save();
            try
            {
                g.SetClip(new Rectangle(bounds.Left, bounds.Top, Math.Max(0, bounds.Width - dots), bounds.Height), CombineMode.Intersect);
                string[] parts = { summary.MatchedText, "（", summary.MissingText, "）。" + summary.Additional };
                Color[] colors = { Ui.Green, normal, Ui.Danger, normal }; int x = bounds.Left;
                for (int i = 0; i < parts.Length; i++)
                {
                    int width = TextRenderer.MeasureText(g, parts[i], body, Size.Empty, flags).Width;
                    TextRenderer.DrawText(g, parts[i], body, new Rectangle(x, bounds.Top, width, bounds.Height), colors[i], flags); x += width;
                }
            }
            finally { g.Restore(saved); }
            if (ellipsis) TextRenderer.DrawText(g, "…", body, new Rectangle(bounds.Right - dots, bounds.Top, dots, bounds.Height), normal, flags);
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
                : part == Part.Status ? items[row].StatusTip : part == Part.Mode ? "点击切换：只写数值 → 数值+格式 → 只写公式 → 公式+格式" : part == Part.Check ? "启用或停用这一行" : "";
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
            // Item padding pushes the text towards the top; centre it on the
            // item's full height instead.
            var text = e.TextRectangle; e.TextRectangle = new Rectangle(text.Left, 0, text.Width, e.Item.Height);
            e.TextFormat = (e.TextFormat & ~(TextFormatFlags.Top | TextFormatFlags.Bottom)) | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine;
            base.OnRenderItemText(e);
        }
    }

    // 批量复制数据: 源区域 and 目标区域 are kept in their own pages; 映射 pairs
    // them row by row (drag ⋮⋮ there to re-pair). A range panel on the right
    // (opened from the window edge) shows the sheet so ranges are dragged.
    internal sealed class DataMigrationForm : Form, IDialogDefaultLayout
    {
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
        private readonly Label status = Ui.Label("在“源区域”“目标区域”页各列出要复制的区域，“映射”页按行把它们对应起来；右侧“区域选择”可直接框选。", 9, true);
        private readonly RoundedButton preview = Ui.Button("预检与预览", 120), execute = Ui.Button("执行复制", 110), restore = Ui.Button("还原操作", 110), clearBackups = Ui.Button("清除备份", 110), open = Ui.Button("打开结果文件夹", 150);
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
        private readonly MigrationRecovery recovery = new MigrationRecovery();
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
            Text = "批量复制数据"; DialogSizing.Prepare(this, "migration.data", "table", Ui.Blue);
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
            status.ForeColor = Ui.Blue; status.AutoSize = true; status.AutoEllipsis = false;
            root.Controls.Add(status, 0, 2);
            root.SizeChanged += delegate { status.MaximumSize = new Size(Math.Max(100, root.ClientSize.Width - root.Padding.Horizontal), 0); };
            var actions = Ui.ActionBar(); var close = Ui.Button("关闭", 90);
            preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); };
            restore.Name = "migrationRestore"; restore.Click += delegate { RestoreTargets(); };
            restoreTip.SetToolTip(restore, "把本窗口中以“更新目标原文件”方式复制过的工作簿，还原为复制前的备份。");
            clearBackups.Name = "migrationClearBackups"; clearBackups.Click += delegate { ClearBackups(); };
            restoreTip.SetToolTip(clearBackups, "永久删除本窗口复制时生成的备份文件；删除后无法再还原。");
            restoreTip.SetToolTip(edge, "打开或收起右侧区域选择"); restoreTip.SetToolTip(gear, "预设与设置");
            guardTail.Tick += delegate { GuardWindow(); if (--guardTailTicks <= 0) guardTail.Stop(); };
            open.Click += delegate { OpenResult(); }; close.Click += delegate { Close(); };
            actions.Controls.Add(preview); actions.Controls.Add(execute); actions.Controls.Add(restore); actions.Controls.Add(clearBackups); actions.Controls.Add(open); actions.Controls.Add(close);
            replaceTargets.Margin = new Padding(16, 0, 8, 0); allowEmpty.Margin = new Padding(8, 0, 8, 0);
            actions.Controls.Add(replaceTargets); actions.Controls.Add(allowEmpty); CenterSourceRow(actions);
            root.Controls.Add(actions, 0, 3); Controls.Add(root);
            inputs.AddRange(new Control[] { allowEmpty, replaceTargets });
            allowEmpty.CheckedChanged += delegate { Changed(); };
            replaceTargets.CheckedChanged += delegate { Changed(); Message(replaceTargets.Checked ? "校验通过后更新目标原文件；原版本保存在目标目录的“复制备份”子文件夹。请先保存并关闭目标。" : "另存“复制结果”副本，目标原文件不修改；请重新预检。"); };
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
            string pickLabel="行末框选按钮"; int toggleWidth=Ui.Px(this,200);
            int minimumWidth=pad+TextRenderer.MeasureText(pickLabel,font).Width+gap+toggleWidth+pad;
            var panel=PresetMenuView.Rows(this,drop,font,delegate { SavePreset(); },delegate { ShowPresetManager(); },minimumWidth);
            int width=panel.Controls[0].Width;
            Action<string> section=delegate(string text) { panel.Controls.Add(new Label { Text=text,AutoSize=false,Size=new Size(width,Ui.Px(this,28)),TextAlign=ContentAlignment.BottomLeft,ForeColor=Ui.InkDim,Padding=new Padding(pad,0,0,Ui.Px(this,3)),Margin=Padding.Empty }); };
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
            g.Columns["range"].HeaderCell = new MigrationSelectionHeader { Value = role + "区域", ToolTipText = "点击下拉箭头，从已输入的区域中选一个，应用到选中的行（只选一行时应用到全部）。" };
            g.ColumnHeaderMouseClick += delegate(object sender, DataGridViewCellMouseEventArgs e)
            {
                if (e.Button != MouseButtons.Left) return;
                if (g.Columns[e.ColumnIndex].Name == "sheet") ShowSheetMenu(target);
                else if (g.Columns[e.ColumnIndex].Name == "range") ShowRangePopup(g, target);
            };
            // Only the range is typed; workbook and sheet are picked from popups.
            g.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
            {
                if (syncing || e.RowIndex < 0 || e.ColumnIndex < 0 || g.Columns[e.ColumnIndex].Name != "range") return;
                var row = g.Rows[e.RowIndex]; var pair = (MigrationPair)row.Tag; var side = target ? pair.Target : pair.Source;
                side.Range = MigrationPlanner.NormalizeAddress(Convert.ToString(row.Cells["range"].Value));
                pair.Instances.Clear();
                // Rows of one by-name target share its range: update them all.
                foreach (DataGridViewRow other in g.Rows) if (other.Tag is MigrationPair && SideOf((MigrationPair)other.Tag, target) == side) WriteSideRow(other, target);
                UpdateMappingRow(pair); Changed(); UpdatePreview();
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
            g.HandleCreated += delegate { g.Columns["sheet"].HeaderCell.Style.Padding = new Padding(Ui.Px(g, 6), 0, Ui.Px(g, 28), 0); g.Columns["range"].HeaderCell.Style.Padding = new Padding(Ui.Px(g, 6), 0, Ui.Px(g, 28), 0); g.Columns["pick"].HeaderCell.Style.Padding = new Padding(1); if (g.Columns.Contains("trigger")) g.Columns["trigger"].HeaderCell.Style.Padding = new Padding(1); };
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
        // Header drop-down of 区域: the ranges already entered (this page
        // first, then the other one), applied to the selected rows, or to all
        // rows when at most one is selected (like 统一设置工作表).
        private void ShowRangePopup(DataGridView grid, bool target)
        {
            if (Running || grid.ReadOnly || grid.Rows.Count == 0) return;
            bool wasOpen = sheetPopup != null && sheetPopup.Visible; ClosePopups(); if (wasOpen) return;
            EndEdits();
            var ranges = DistinctSides(target).Concat(DistinctSides(!target)).Select(x => x.Range).Where(r => r.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (ranges.Count == 0) { Message("还没有输入过区域：在各行的区域格直接输入，或点行末的框选按钮。"); return; }
            var selected = grid.SelectedRows.Cast<DataGridViewRow>().Select(r => r.Tag as MigrationPair).Where(p => p != null).Select(p => SideOf(p, target)).Distinct().ToList();
            var sides = selected.Count > 1 ? selected : DistinctSides(target);
            var values = sides.Select(x => x.Range).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            sheetPopup = MigrationSheetPopup.ForRanges(grid, ranges, values.Count == 1 ? values[0] : null, delegate(string range)
            {
                if (Running) return;
                foreach (var side in sides) side.Range = MigrationPlanner.NormalizeAddress(range);
                foreach (var pair in pairs) if (sides.Contains(SideOf(pair, target))) pair.Instances.Clear();
                RefreshRows(); Changed(); UpdatePreview(); ShowPairDetails();
                Message("已将 " + sides.Count + " 个" + (target ? "目标" : "源") + "区域设为 " + range + "。");
            });
            var cell = grid.GetCellDisplayRectangle(grid.Columns["range"].Index, -1, true);
            sheetPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(sheetPopup);
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
                // Click cycles 只写数值 → 数值+格式 → 只写公式 → 公式+格式;
                // 逐行重算 only toggles the formats (it writes values).
                var pair = pairs[row]; bool trigger = pair.Source.Trigger != null;
                // → 只迁移格式 → back to 只写数值.
                if (trigger) { pair.ValuesOnly = true; pair.FormatOnly = false; pair.KeepFormat = !pair.KeepFormat; }
                else if (pair.FormatOnly) { pair.FormatOnly = false; pair.ValuesOnly = true; pair.KeepFormat = false; }
                else if (pair.ValuesOnly) { if (!pair.KeepFormat) pair.KeepFormat = true; else { pair.ValuesOnly = false; pair.KeepFormat = false; } }
                else { if (!pair.KeepFormat) pair.KeepFormat = true; else { pair.FormatOnly = true; pair.ValuesOnly = true; pair.KeepFormat = true; } }
                pair.Instances.Clear(); UpdateMappingRow(pair); Changed(); ShowPairDetails();
                Message("映射第 " + (row + 1) + " 行改为“" + MigrationPair.ModeText(pair.ValuesOnly || trigger, pair.KeepFormat, pair.FormatOnly && !trigger) + "”。");
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
            grid.ContextMenuStrip.Opening += delegate
            {
                int count = grid.SelectedRows.Cast<DataGridViewRow>().Select(r => r.Tag as MigrationPair).Where(p => p != null).Select(p => SideOf(p, target)).Distinct().Count();
                grid.ContextMenuStrip.Items[1].Text = count > 1 ? "复制选中的 " + count + " 行" : "复制此行";
                grid.ContextMenuStrip.Items[3].Text = count > 1 ? "删除选中的 " + count + " 行" : "删除此行";
            };
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
        // Rows created in the window paste values without formats by default;
        // the other 写入方式 are a per-row choice in 映射.
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
                var matched = byName ? MatchedFiles(side) : null;
                // Target rows a 逐行重算 source writes to carry a 【重算】 tag.
                string tag = target && pairs.Any(p => p.Target == side && p.Source.Used && p.Source.Trigger != null) ? "【重算】" : "";
                row.Cells["book"].Value = tag + (byName ? (matched.Count > 0 ? BooksText(matched) : "（按文件名匹配：未匹配到目标）") : BooksText(side.Books));
                if (!target) row.Cells["trigger"].Value = side.Trigger != null ? "on" : "";
                row.Cells["book"].ToolTipText = byName ? (matched.Count > 0 ? "由逐行重算按文件名匹配：" + Environment.NewLine + string.Join(Environment.NewLine, matched.ToArray()) : "逐行重算没有匹配到已添加的目标工作簿")
                    : side.Books.Count == 0 ? "点击选择" + (target ? "目标" : "源") + "工作簿（可多选）" : string.Join(Environment.NewLine, side.Books.ToArray());
                row.Cells["sheet"].Value = side.Sheet.Length == 0 ? "（选择工作表）" : side.Sheet;
                row.Cells["sheet"].ToolTipText = side.Books.Count == 0 ? "请先选择工作簿" : "点击选择" + (target ? "目标" : "源") + "工作表";
                if (Convert.ToString(row.Cells["range"].Value) != side.Range) row.Cells["range"].Value = side.Range;
                var area = target ? MigrationPlanner.TargetArea(pair) : MigrationPlanner.TryParse(side.Range);
                row.Cells["size"].Value = side.Range.Length == 0 ? "" : area == null ? "区域无效" : area.Rows + " 行 × " + area.Columns + " 列" + (target && area.Address != side.Range && side.Range.IndexOf(':') < 0 ? "（按源）" : "");
            }
            finally { syncing = was; }
        }
        // What the 映射 board shows for one mapping row.
        private MigrationMappingBoard.Entry EntryOf(MigrationPair pair)
        {
            var source = MigrationPlanner.TryParse(pair.Source.Range); var target = MigrationPlanner.TargetArea(pair);
            MigrationMappingBoard.MatchSummary match; string statusText = PairStatus(pair, out match);
            var entry = new MigrationMappingBoard.Entry
            {
                SourceUsed = pair.Source.Used, TargetUsed = pair.Target.Used, Enabled = pair.Enabled, ValuesOnly = pair.ValuesOnly || pair.Source.Trigger != null, KeepFormat = pair.KeepFormat, FormatOnly = pair.FormatOnly && pair.Source.Trigger == null,
                SourceTitle = CardTitle(pair.Source, source), SourceDetail = CardDetail(pair.Source, source) + Shared(pair.Source, false) + TriggerTag(pair.Source),
                TargetTitle = CardTitle(pair.Target, target),
                TargetDetail = IsByName(pair.Target) ? "按文件名匹配（已添加 " + TargetBooks.Count + " 个）" + (target == null ? "" : " · " + target.Rows + "×" + target.Columns) : CardDetail(pair.Target, target),
                SourceTip = string.Join(Environment.NewLine, pair.Source.Books.ToArray()),
                TargetTip = string.Join(Environment.NewLine, pair.Target.Books.ToArray()),
                Status = statusText, Match = match
            };
            string errors = PairErrors(pair); entry.StatusTip = errors.Length > 0 ? errors : entry.Status;
            string s = entry.Status;
            entry.Tone = !pair.Enabled ? 0 : s.StartsWith("目标文件", StringComparison.Ordinal) ? (s.Contains("待选") ? 2 : 4)
                : s.StartsWith("成功", StringComparison.Ordinal) ? 1 : s.StartsWith("部分成功", StringComparison.Ordinal) ? 2
                : s.StartsWith("可执行", StringComparison.Ordinal) || s.StartsWith("就绪", StringComparison.Ordinal) || s.StartsWith("执行中", StringComparison.Ordinal) ? 4
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
            PrepareMatches();
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
                Targets.Rows.Clear(); FillTargetRows(TargetLayout());
                Sources.Rows.Clear();
                foreach (var pair in pairs.Where(p => p.Source.Used).GroupBy(p => p.Source).Select(g => g.First()))
                {
                    int index = Sources.Rows.Add(); Sources.Rows[index].Tag = pair;
                }
            }
            finally { syncing = false; }
            RefreshRows();
            int focus = mappingFocus == null ? -1 : pairs.IndexOf(mappingFocus);
            if (focus < 0) focus = Math.Min(Math.Max(0, Mappings.Current), pairs.Count - 1);
            Mappings.SetItems(pairs.Select(EntryOf).ToList(), focus, Mappings.CurrentTarget);
            mappingFocus = Mappings.Current >= 0 ? pairs[Mappings.Current] : null;
            foreach (var pair in keep) if (pair.Value != null) SelectRow(pair.Key, pair.Value);
            UpdateTabTitles(); Ui.AutoFitColumns(Sources); Ui.AutoFitColumns(Targets);
            UpdatePreview(); ShowPairDetails();
        }
        private void RefreshRows()
        {
            PrepareMatches();
            SyncTargetRows();
            foreach (DataGridViewRow row in Sources.Rows) WriteSideRow(row, false);
            foreach (DataGridViewRow row in Targets.Rows) WriteSideRow(row, true);
            for (int i = 0; i < pairs.Count && i < Mappings.Count; i++) Mappings.UpdateItem(i, EntryOf(pairs[i]));
        }
        private string PairStatus(MigrationPair pair)
        { MigrationMappingBoard.MatchSummary match; return PairStatus(pair, out match); }
        private string PairStatus(MigrationPair pair, out MigrationMappingBoard.MatchSummary match)
        {
            match = null;
            string problem = RowProblem(pair);
            if (!pair.Enabled) return "未启用" + (problem.Length == 0 ? "" : "；" + problem);
            if (pair.Instances.Count == 0 || pair.Instances.All(m => m.Result == null))
            {
                if (problem.Length > 0) return problem;
                if (pair.Source.Trigger == null) return "就绪（待预检）";
                return TriggerSummary(pair.Source, out match);
            }
            var results = pair.Instances.Where(m => m.Result != null).Select(m => m.Result).ToList();
            int total = results.Count, ok = results.Count(t => t.Status == "成功"), ready = results.Count(t => t.Ready && t.Status == "可执行"), restored = results.Count(t => t.Status == "已还原");
            int bad = results.Count(t => t.Status == "失败" || t.Status == "冲突" || t.Status == "已取消");
            // Writes still under way (待执行, 复核…, 写入…): the row is running,
            // not failed, even if a write was already left out at precheck.
            var settled = new HashSet<string> { "成功", "失败", "冲突", "已取消", "已还原", "可执行", "待预检", "需重新预检" };
            int running = results.Count(t => !settled.Contains(t.Status));
            string head = running > 0 ? "执行中 " + ok + "/" + (ok + running)
                : ok > 0 ? (ok == total ? "成功 " : "部分成功 ") + ok + "/" + total
                : restored > 0 ? "已还原 " + restored + "/" + total
                : ready > 0 ? "可执行 " + ready + "/" + total
                : bad > 0 ? results.First(t => t.Status == "失败" || t.Status == "冲突" || t.Status == "已取消").Status + " " + bad + "/" + total
                : results[0].Status;
            var first = results.FirstOrDefault(t => t.Error.Length > 0 && t.Status != "成功" && settled.Contains(t.Status));
            if (first == null) return head;
            // A write left out at precheck (检查：…) is reported as such beside
            // a running or successful row, not as the row's result.
            bool aside = running > 0 || ok > 0 || ready > 0;
            return head + (aside ? "；另有 " + bad + " 处未写入：" : "；") + first.Error;
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
            // Every selected range (each shared range once), in list order.
            var chosen = grid.SelectedRows.Cast<DataGridViewRow>().OrderBy(r => r.Index).Select(r => r.Tag as MigrationPair).Where(p => p != null).Select(p => SideOf(p, target)).Distinct().ToList();
            if (chosen.Count == 0 && grid.CurrentRow != null) chosen.Add(SideOf((MigrationPair)grid.CurrentRow.Tag, target));
            if (chosen.Count == 0) { Message("请先选择要复制的行。"); return; }
            MigrationPair last = null;
            foreach (var side in chosen) last = AddSide(target, side.Copy());
            grid.ClearSelection(); SelectRow(grid, last);
            Message(chosen.Count == 1 ? "已复制为映射第 " + (pairs.IndexOf(last) + 1) + " 行的" + (target ? "目标" : "源") + "区域。" : "已复制 " + chosen.Count + " 个" + (target ? "目标" : "源") + "区域。");
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
                if (target) pair.Target = Remembered(pair.Target); else pair.Source = new MigrationSide();
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
            if (current != null && current.Source.Used) { pair.Source = current.Source; pair.Enabled = current.Enabled; pair.ValuesOnly = current.ValuesOnly; pair.KeepFormat = current.KeepFormat; pair.FormatOnly = current.FormatOnly; }
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
            if (Running) throw new InvalidOperationException("复制运行期间不能更改工作簿。");
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
            if (target && IsByName(((MigrationPair)row.Tag).Target)) { ShowMatchedBookMenu(grid, row); return; }
            if (Pool(target).Count == 0) { PickBooks(target, (MigrationPair)row.Tag); return; }
            bool wasOpen = bookPopup != null && bookPopup.Visible; ClosePopups(); if (wasOpen) return;
            bookPopup = CreateBookPopup((MigrationPair)row.Tag, target);
            var cell = grid.GetCellDisplayRectangle(grid.Columns["book"].Index, row.Index, true);
            bookPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(bookPopup);
        }
        // A by-name target row: the same workbook list as other rows. Ticked =
        // the files 逐行重算 writes to. Untick a file: its trigger rows are
        // skipped. Tick another: it takes the trigger row without a target
        // with the same name, else the first one without a target (or gets
        // back its automatic match).
        private void ShowMatchedBookMenu(DataGridView grid, DataGridViewRow gridRow)
        {
            var pair = (MigrationPair)gridRow.Tag; var side = pair.Target;
            var triggers = pairs.Where(x => x.Target == side && x.Source.Used && x.Source.Trigger != null).Select(x => x.Source).Distinct().ToList();
            foreach (var source in triggers) RefreshTrigger(source);
            if (TargetBooks.Count == 0) { PickTargetPool(); return; }
            bool wasOpen = bookPopup != null && bookPopup.Visible; ClosePopups(); if (wasOpen) return;
            var before = MatchedFiles(side);
            bookPopup = new MigrationBookPopup(grid, TargetBooks, before, "目标工作簿", delegate(IList<string> chosen)
            {
                if (Running || !pairs.Contains(pair)) return;
                var keep = new HashSet<string>(chosen, StringComparer.OrdinalIgnoreCase);
                int skipped = 0, assigned = 0; var unplaced = new List<string>();
                foreach (var source in triggers)
                    foreach (var row in source.Trigger.Rows)
                        if (row.Path.Length > 0 && !keep.Contains(row.Path)) { source.Trigger.Chosen[row.Row] = ""; skipped++; }
                var index = TargetIndex();
                foreach (string file in chosen.Where(f => !before.Contains(f, StringComparer.OrdinalIgnoreCase)))
                {
                    MigrationTrigger best = null; MigrationTriggerRow bestRow = null; double score = -1;
                    foreach (var source in triggers)
                        foreach (var row in source.Trigger.Rows.Where(r => r.Kind != "空值" && r.Path.Length == 0))
                        {
                            // A row skipped by hand whose automatic match is this file gets it back first.
                            var matches = index.Find(row.Name);
                            double value = matches.Count == 1 && string.Equals(matches[0], file, StringComparison.OrdinalIgnoreCase) ? 2 : matches.Contains(file, StringComparer.OrdinalIgnoreCase) ? 1 : 0;
                            if (value > score) { score = value; best = source.Trigger; bestRow = row; }
                        }
                    if (bestRow == null) { unplaced.Add(Path.GetFileName(file)); continue; }
                    best.Chosen[bestRow.Row] = file; best.Signature = ""; bestRow.Path = file; assigned++;
                }
                foreach (var source in triggers) source.Trigger.Signature = "";
                foreach (var p in pairs) if (p.Source.Trigger != null) p.Instances.Clear();
                RefreshRows(); Changed(); ShowPairDetails();
                Message((skipped > 0 ? "不写入 " + skipped + " 个触发行。" : "") + (assigned > 0 ? "新增写入 " + assigned + " 个工作簿。" : "")
                    + (unplaced.Count > 0 ? "没有空着的触发行可写入：" + string.Join("、", unplaced.ToArray()) + "。" : "") + (skipped + assigned + unplaced.Count == 0 ? "未改变。" : ""));
            }, delegate { BeginInvoke(new Action(PickTargetPool)); });
            var cell = grid.GetCellDisplayRectangle(grid.Columns["book"].Index, gridRow.Index, true);
            bookPopup.Show(grid, new Point(cell.Left, cell.Bottom));
            MigrationDropDownWatch.Attach(bookPopup);
        }
        // 添加工作簿 for 逐行重算: the files only join the target list and are
        // matched by name; no row is chosen for them by hand.
        private void PickTargetPool()
        {
            try
            {
                string[] selections = MigrationTargetPicker.Show(this); if (selections.Length == 0) return;
                var scan = MigrationTargetScan.Expand(selections, "");
                if (scan.Files.Count > 0 && BooksPicked != null) BooksPicked(true, scan.Files.ToArray());
                foreach (var side in DistinctSides(false).Where(x => x.Trigger != null)) side.Trigger.Signature = "";
                foreach (var p in pairs) if (p.Source.Trigger != null) p.Instances.Clear();
                foreach (var side in DistinctSides(false).Where(x => x.Trigger != null)) SplitByFile(side);
                RebuildViews(); Changed(); ShowPairDetails();
                Message("已添加 " + scan.Files.Count + " 个目标工作簿，逐行重算按文件名重新匹配。" + string.Join("；", scan.Errors.ToArray()));
            }
            catch (Exception error) { Message("添加工作簿失败：" + error.Message); }
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
            // A by-name target (逐行重算) counts with the files it matched.
            var rows = pairs.Where(p => (target ? p.Target : p.Source).Used && BooksSide(target ? p.Target : p.Source, target).Books.Count > 0).ToList();
            if (rows.Count == 0) { Message("请先为各行选择" + (target ? "目标" : "源") + "工作簿。"); return; }
            // Counted per workbook: a sheet only some workbooks have is still
            // offered; those lacking it are split off into their own rows.
            var books = rows.Select(p => BooksSide(target ? p.Target : p.Source, target)).Distinct().SelectMany(s => s.Books).ToList();
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
            var below = new MigrationPair { Enabled = pair.Enabled, ValuesOnly = pair.ValuesOnly, KeepFormat = pair.KeepFormat, FormatOnly = pair.FormatOnly };
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
            var source = MigrationPlanner.TryParse(pair.Source.Range);
            var area = target ? MigrationPlanner.TargetArea(pair) : source;
            var used = new List<KeyValuePair<MigrationArea, string>>();
            foreach (var other in pairs)
            {
                if (other == pair || !other.Enabled) continue;
                var otherSide = target ? other.Target : other.Source; if (!otherSide.Used) continue;
                if (!otherSide.Books.Contains(path, StringComparer.OrdinalIgnoreCase) || !string.Equals(otherSide.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase)) continue;
                var otherArea = target ? MigrationPlanner.TargetArea(other) : MigrationPlanner.TryParse(other.Source.Range);
                if (otherArea != null) used.Add(new KeyValuePair<MigrationArea, string>(otherArea, "映射第 " + (pairs.IndexOf(other) + 1) + " 行 " + otherArea.Address));
            }
            string hint = !target
                ? "拖选要复制的区域（单击为单个单元格）。浅绿为同一工作表中其他源区域。勾选“作为新的源区域加入”可连续框选多块，每块新增一个源区域。"
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
                string range = single ? CellAddress.Name(picked.Row, picked.Column) : picked.Address;
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
            var source = MigrationPlanner.TryParse(pair.Source.Range);
            try { picked.Check(1048576, 16384); }
            catch (Exception error) { Message(error.Message); sheetPreview.SetHint(error.Message); return; }
            side.Range = single ? CellAddress.Name(picked.Row, picked.Column) : picked.Address;
            string fit = source == null || single || (picked.Rows == source.Rows && picked.Columns == source.Columns) ? ""
                : "（" + picked.Rows + "×" + picked.Columns + "，与源区域 " + source.Rows + "×" + source.Columns + " 不一致，“映射”页会提示）";
            pair.Instances.Clear(); RefreshRows(); Changed();
            var mine = MigrationPlanner.TargetArea(pair);
            var overlap = pairs.Where(p => p != pair && p.Enabled && p.Target.Used && string.Equals(p.Target.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase)
                    && p.Target.Books.Intersect(side.Books, StringComparer.OrdinalIgnoreCase).Any() && Overlaps(MigrationPlanner.TargetArea(p), mine))
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
        // A target range of 逐行重算 that has no workbook of its own yet: it
        // stands for the matched files until 应用 gives each file its own row.
        private bool IsByName(MigrationSide side)
        { return side.Used && side.Books.Count == 0 && pairs.Any(p => p.Target == side && p.Source.Used && p.Source.Trigger != null); }
        // After 应用: the 逐行重算 target becomes one ordinary row whose
        // workbooks are the matched files (shown as "01-… 等 5 个"); its sheet
        // and range apply to all of them. Files another row of this source
        // already lists stay with that row.
        private int SplitByFile(MigrationSide source)
        {
            if (source.Trigger == null) return 0;
            RefreshTrigger(source); if (source.Trigger.Error.Length > 0) return 0;
            var mine = pairs.Where(p => p.Source == source && p.Target.Used).ToList();
            var template = mine.FirstOrDefault(p => p.Target.Books.Count == 0) ?? mine.FirstOrDefault();
            if (template == null)
            {
                // Its target row was deleted or cleared: 应用 brings it back
                // with the sheet and range it had, if remembered.
                template = pairs.FirstOrDefault(p => p.Source == source && !p.Target.Used); if (template == null) return 0;
                if (pairs.Any(p => p.Target == template.Target && p != template)) template.Target = Remembered(template.Target);
                template.Target.Used = true; template.Target.Books.Clear(); template.Instances.Clear();
                mine.Add(template);
            }
            var covered = new HashSet<string>(mine.Where(p => p.Target != template.Target).SelectMany(p => p.Target.Books), StringComparer.OrdinalIgnoreCase);
            var files = new List<string>();
            foreach (var row in source.Trigger.Rows) if (row.Path.Length > 0 && !covered.Contains(row.Path) && !files.Contains(row.Path, StringComparer.OrdinalIgnoreCase)) files.Add(row.Path);
            if (files.Count == 0) return 0;
            template.Target.Recalc = true;
            if (files.Count == template.Target.Books.Count && files.All(f => template.Target.Books.Contains(f, StringComparer.OrdinalIgnoreCase))) return 0;
            template.Target.Books.Clear(); template.Target.Books.AddRange(files);
            foreach (var p in pairs.Where(x => x.Target == template.Target)) p.Instances.Clear();
            return files.Count;
        }
        // The workbooks a side stands for: a by-name target stands for all
        // added targets (for sheet lists, picking and preview).
        private MigrationSide BooksSide(MigrationSide side, bool target)
        {
            if (!target || !IsByName(side)) return side;
            var all = new MigrationSide { Sheet = side.Sheet, Range = side.Range, Used = side.Used };
            var matched = MatchedFiles(side);
            all.Books.AddRange(matched.Count > 0 ? matched : TargetBooks.Select(b => b.Path).ToList());
            return all;
        }
        // The files the 逐行重算 rows feeding this target side matched, in row order.
        private List<string> MatchedFiles(MigrationSide side)
        {
            var files = new List<string>();
            foreach (var pair in pairs.Where(p => p.Target == side && p.Source.Used && p.Source.Trigger != null))
            {
                foreach (var row in pair.Source.Trigger.Rows) if (row.Path.Length > 0 && !files.Contains(row.Path, StringComparer.OrdinalIgnoreCase)) files.Add(row.Path);
            }
            return files;
        }
        private List<MigrationPair> TargetLayout()
        {
            return pairs.Where(p => p.Target.Used).GroupBy(p => p.Target).Select(g => g.First()).ToList();
        }
        private void FillTargetRows(List<MigrationPair> layout)
        {
            foreach (var pair in layout) { int index = Targets.Rows.Add(); Targets.Rows[index].Tag = pair; }
        }
        private bool targetLayoutPending;
        // Rebuild target rows only when their side identities or order change.
        private bool TargetLayoutMatches(List<MigrationPair> layout)
        {
            if (layout.Count != Targets.Rows.Count) return false;
            for (int i = 0; i < layout.Count; i++)
                if (Targets.Rows[i].Tag != layout[i]) return false;
            return true;
        }
        // Never rebuilt inside the caller: RefreshRows also runs from grid
        // events and popups, and clearing rows there is re-entrant (WPS
        // closed). The rebuild is posted once and guarded.
        private void SyncTargetRows()
        {
            if (TargetLayoutMatches(TargetLayout())) return;
            if (!IsHandleCreated) { RebuildTargetRows(); return; }
            if (targetLayoutPending) return;
            targetLayoutPending = true;
            BeginInvoke(new Action(delegate
            {
                targetLayoutPending = false;
                if (IsDisposed) return;
                try
                {
                    if (TargetLayoutMatches(TargetLayout())) return;
                    if (Targets.IsCurrentCellInEditMode) Targets.EndEdit();
                    RebuildTargetRows();
                    foreach (DataGridViewRow row in Targets.Rows) WriteSideRow(row, true);
                }
                catch (Exception error) { UiDiagnostics.Record("migration-target-rows", error); }
            }));
        }
        private void RebuildTargetRows()
        {
            var layout = TargetLayout();
            var current = Targets.CurrentRow == null ? null : Targets.CurrentRow.Tag as MigrationPair; bool was = syncing; syncing = true;
            try { Targets.Rows.Clear(); FillTargetRows(layout); }
            finally { syncing = was; }
            if (current != null) SelectRow(Targets, current);
            UpdateTabTitles();
        }
        // Added target workbooks nothing uses: no 逐行重算 row matched them or
        // lists them as a candidate, and no ordinary mapping writes to them.
        private int DropUnmatchedTargets()
        {
            // Target rows superseded by the by-name rows: not by-name, not a
            // 逐行重算 row, and either unpaired or writing only to files
            // 逐行重算 matched. A cleared row keeps its sheet and range
            // (unused), so a later 逐行重算 of that source reuses them.
            var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var source in DistinctSides(false).Where(x => x.Trigger != null)) { RefreshTrigger(source); foreach (var row in source.Trigger.Rows) if (row.Path.Length > 0) matched.Add(row.Path); }
            int cleared = 0;
            if (matched.Count > 0)
                foreach (var pair in pairs.ToList())
                {
                    var side = pair.Target; if (!side.Used || side.Recalc || pairs.Any(x => x.Target == side && x.Source.Trigger != null)) continue;
                    if (pair.Source.Used && side.Books.Any(b => !matched.Contains(b))) continue;
                    foreach (var other in pairs.Where(x => x.Target == side).ToList()) other.Target = Remembered(side);
                    cleared++;
                }
            pairs.RemoveAll(x => !x.Source.Used && !x.Target.Used);
            var keep = new HashSet<string>(pairs.Where(p => p.Target.Used && !IsByName(p.Target)).SelectMany(p => p.Target.Books), StringComparer.OrdinalIgnoreCase);
            foreach (var side in DistinctSides(false).Where(s => s.Trigger != null))
            {
                RefreshTrigger(side); if (side.Trigger.Error.Length > 0 || side.Trigger.Rows.Count == 0) return cleared;
                foreach (var row in side.Trigger.Rows) { if (row.Path.Length > 0) keep.Add(row.Path); foreach (var c in row.Candidates) keep.Add(c.Key); }
            }
            var drop = TargetBooks.Where(b => !keep.Contains(b.Path)).ToList();
            foreach (var book in drop) TargetBooks.Remove(book);
            if (drop.Count > 0) { foreach (var side in DistinctSides(false).Where(s => s.Trigger != null)) side.Trigger.Signature = ""; foreach (var pair in pairs) pair.Instances.Clear(); }
            return drop.Count + cleared;
        }
        // An unused target side that remembers a removed row's sheet and range.
        private static MigrationSide Remembered(MigrationSide side)
        { return new MigrationSide { Sheet = side.Sheet, Range = side.Range }; }
        private string TriggerTag(MigrationSide side)
        {
            if (side.Trigger == null) return "";
            return side.Trigger.Error.Length > 0 ? " · 逐行（设置无效）" : " · 逐行 " + side.Trigger.Rows.Count(r => r.Kind != "空值");
        }
        private sealed class MatchView
        {
            internal MigrationMappingBoard.MatchSummary Summary;
            internal List<string> Unused;
        }
        private readonly Dictionary<MigrationSide, MatchView> matches = new Dictionary<MigrationSide, MatchView>();
        private string TriggerSummary(MigrationSide side, out MigrationMappingBoard.MatchSummary summary)
        {
            MatchView view; matches.TryGetValue(side, out view);
            summary = view == null ? null : view.Summary;
            return side.Trigger.Error.Length > 0 ? "逐行重算设置无效：" + side.Trigger.Error : summary == null ? "待刷新匹配" : summary.Text;
        }
        private void PrepareMatches()
        {
            matches.Clear();
            var sides = DistinctSides(false).Where(side => side.Trigger != null).ToList();
            if (sides.Count == 0) return;
            var index = TargetIndex();
            foreach (var side in sides)
            {
                RefreshTrigger(side, index);
                var unused = MigrationPlanner.UnusedTargets(side, pairs, TargetBooks);
                var rows = side.Trigger.Rows.Where(row => row.Kind != "空值").ToList();
                int many = rows.Count(row => row.Kind == "多个候选"), skipped = rows.Count(row => row.Kind == "跳过");
                var parts = new List<string>();
                if (many > 0) parts.Add("同名待选 " + many + " 行"); if (skipped > 0) parts.Add("跳过 " + skipped + " 行");
                if (unused.Count > 0) parts.Add("另有 " + unused.Count + " 个目标工作簿未用到");
                matches.Add(side, new MatchView { Unused = unused, Summary = new MigrationMappingBoard.MatchSummary {
                    Rows = rows.Count, Matched = rows.Count(row => row.Ready), Missing = rows.Count(row => row.Kind == "缺失"),
                    Additional = parts.Count == 0 ? "" : " " + string.Join("，", parts.ToArray()) } });
            }
        }
        // Re-read the trigger values and names (from the saved file, without
        // opening WPS) and match each name against the added targets. Cached
        // until the file, settings, targets or choices change.
        private MigrationNameIndex nameIndex;
        private MigrationNameIndex TargetIndex()
        {
            var paths = TargetBooks.Select(b => b.Path);
            if (nameIndex == null || !nameIndex.Paths.SequenceEqual(paths, StringComparer.OrdinalIgnoreCase)) nameIndex = new MigrationNameIndex(paths);
            return nameIndex;
        }
        private void RefreshTrigger(MigrationSide side, MigrationNameIndex index = null)
        {
            var trigger = side.Trigger; if (trigger == null) return;
            index = index ?? TargetIndex();
            string path = side.Books.Count == 1 ? side.Books[0] : "";
            long stamp = path.Length > 0 && File.Exists(path) ? File.GetLastWriteTimeUtc(path).Ticks : 0;
            string signature = path + "|" + side.Sheet + "|" + trigger.Cell + "|" + trigger.Values + "|" + trigger.NameColumn + "|" + stamp + "|"
                + index.Signature + "|" + string.Join(",", trigger.Chosen.Select(c => c.Key + "=" + c.Value).ToArray());
            if (signature == trigger.Signature) return;
            trigger.Signature = signature; trigger.Rows.Clear(); trigger.Error = "";
            try
            {
                if (side.Books.Count != 1) throw new InvalidOperationException("逐行重算只支持一个源工作簿");
                if (side.Sheet.Length == 0) throw new InvalidOperationException("请先选择源工作表");
                var cell = MigrationArea.Parse(trigger.Cell); if (cell.Rows != 1 || cell.Columns != 1) throw new InvalidOperationException("触发单元格只能是一个单元格");
                var values = MigrationArea.Parse(trigger.Values); if (values.Columns != 1) throw new InvalidOperationException("触发值区域只能是一列");
                int nameColumn = MigrationPlanner.NameColumnIndex(trigger.NameColumn);
                var snapshot = LoadSnapshot(path, side.Sheet);
                var pool = index.Paths;
                for (int r = values.Row; r < values.Row + values.Rows; r++)
                {
                    var row = new MigrationTriggerRow { Row = r, Value = snapshot.Value(r, values.Column).Trim(), Name = snapshot.Value(r, nameColumn).Trim() };
                    if (row.Value.Length == 0) row.Kind = "空值";
                    else
                    {
                        index.Match(row); string chosen;
                        if (trigger.Chosen.TryGetValue(r, out chosen) && (chosen.Length == 0 || pool.Contains(chosen, StringComparer.OrdinalIgnoreCase)))
                        { row.Kind = chosen.Length == 0 ? "跳过" : "已确认"; row.Path = chosen; }
                    }
                    trigger.Rows.Add(row);
                }
            }
            catch (Exception error) { trigger.Error = error.Message; }
        }
        // Set or change 逐行重算 for a source range. One workbook and sheet
        // may have several independent sets (different trigger cells or
        // value lists); ranges already using this range's set change with it.
        internal void SetTrigger(MigrationSide side, MigrationTrigger trigger)
        {
            int shared = 0;
            if (side.Trigger != null)
                foreach (var other in DistinctSides(false))
                {
                    if (other == side || !SameTriggerSet(other, side)) continue;
                    other.Trigger = trigger == null ? null : trigger.Copy(); shared++;
                }
            side.Trigger = trigger;
            foreach (var pair in pairs) if (pair.Source.Used && (pair.Source == side || (pair.Source.Books.Count == 1 && side.Books.Count == 1 && string.Equals(pair.Source.Books[0], side.Books[0], StringComparison.OrdinalIgnoreCase)))) pair.Instances.Clear();
            RefreshRows(); Changed(); ShowPairDetails();
            int removed = trigger == null ? 0 : DropUnmatchedTargets();
            int split = 0;
            if (trigger != null) foreach (var source in DistinctSides(false).Where(x => x.Trigger != null)) split += SplitByFile(source);
            if (removed > 0 || split > 0) RebuildViews();
            Message(trigger == null ? "已取消应用逐行重算" + (shared > 0 ? "（使用同一套设置的另外 " + shared + " 个源区域一并取消）" : "") + "。"
                : "已应用逐行重算：触发值 " + trigger.Values + " → " + trigger.Cell + string.Concat(trigger.Extra.Select(c => "，" + c.Text).ToArray()) + "，目标文件区域 " + trigger.NameColumn + (shared > 0 ? "；使用同一套设置的 " + shared + " 个源区域一并更新" : "")
                  + (removed > 0 ? "；目标区域已清理，只保留逐行重算匹配到的工作簿" : "") + "。目标区域页这一行的工作簿即匹配到的文件。");
        }
        private static bool SameTriggerSet(MigrationSide a, MigrationSide b)
        {
            if (a.Trigger == null || b.Trigger == null || a.Books.Count != 1 || b.Books.Count != 1) return false;
            return string.Equals(a.Books[0], b.Books[0], StringComparison.OrdinalIgnoreCase) && string.Equals(a.Sheet, b.Sheet, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.Trigger.Cell, b.Trigger.Cell, StringComparison.OrdinalIgnoreCase) && string.Equals(a.Trigger.Values, b.Trigger.Values, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.Trigger.NameColumn, b.Trigger.NameColumn, StringComparison.OrdinalIgnoreCase)
                && string.Join("|", a.Trigger.Extra.Select(c => c.Text).ToArray()) == string.Join("|", b.Trigger.Extra.Select(c => c.Text).ToArray());
        }
        // Other 逐行重算 sets on this source's workbook and sheet.
        private int OtherTriggerSets(MigrationSide side)
        {
            var sets = new List<MigrationSide>();
            foreach (var other in DistinctSides(false))
                if (other != side && other.Trigger != null && other.Books.Count == 1 && side.Books.Count == 1 && string.Equals(other.Books[0], side.Books[0], StringComparison.OrdinalIgnoreCase)
                    && string.Equals(other.Sheet, side.Sheet, StringComparison.OrdinalIgnoreCase) && !SameTriggerSet(other, side) && !sets.Any(s => SameTriggerSet(s, other))) sets.Add(other);
            return sets.Count;
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
            var panel = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 4, BackColor = Color.White, Font = font, Padding = new Padding(Ui.Px(this, 16), Ui.Px(this, 14), Ui.Px(this, 16), Ui.Px(this, 12)) };
            var title = new Label { Text = "逐行重算", AutoSize = true, ForeColor = Ui.Ink, Font = Ui.MakeFont(10f, FontStyle.Bold, Font), Margin = new Padding(0, 0, 0, p8) };
            panel.Controls.Add(title, 0, 0); panel.SetColumnSpan(title, 3);
            // 触发单元格 1 / 触发值 1, then 触发单元格 2 / 触发值 2 … (all
            // filled per row before one recalculation), then 目标文件.
            Func<TextBox> box = delegate { return new TextBox { Width = Ui.Px(this, 150), Margin = new Padding(0, p4, p4, p4) }; };
            var cellBox = box(); cellBox.Name = "migrationTriggerCell"; cellBox.Text = value.Cell;
            var valuesBox = box(); valuesBox.Name = "migrationTriggerValues"; valuesBox.Text = value.Values;
            var nameBox = box(); nameBox.Name = "migrationTriggerName";
            // Older settings kept only the file-name column: show it as the range beside 触发值 1.
            string nameText = value.NameColumn;
            if (Regex.IsMatch(nameText ?? "", "^[A-Za-z]{1,3}$")) { var rows = MigrationPlanner.TryParse(value.Values); if (rows != null) nameText = CellAddress.Name(rows.Row, MigrationPlanner.ColumnIndex(nameText)) + ":" + CellAddress.Name(rows.Row + rows.Rows - 1, MigrationPlanner.ColumnIndex(nameText)); }
            nameBox.Text = nameText;
            var extraCells = new List<TextBox>(); var extraValues = new List<TextBox>();
            Action keep = delegate
            {
                value.Cell = MigrationPlanner.NormalizeAddress(cellBox.Text); value.Values = MigrationPlanner.NormalizeAddress(valuesBox.Text); value.NameColumn = MigrationPlanner.NormalizeAddress(nameBox.Text);
                for (int i = 0; i < value.Extra.Count && i < extraCells.Count; i++)
                {
                    string from = (extraValues[i].Text ?? "").Trim();
                    value.Extra[i].Cell = MigrationPlanner.NormalizeAddress(extraCells[i].Text); value.Extra[i].Column = "";
                    bool isRange = MigrationPlanner.TryParse(from) != null;
                    value.Extra[i].Range = isRange ? MigrationPlanner.NormalizeAddress(from) : ""; value.Extra[i].Value = isRange ? "" : from;
                }
            };
            Action<string, TextBox, int, bool, Action<MigrationArea>> line = delegate(string caption, TextBox field, int at, bool single, Action<MigrationArea> picked)
            {
                panel.Controls.Add(new Label { Text = caption, AutoSize = true, ForeColor = Ui.InkDim, Margin = new Padding(0, Ui.Px(this, 8), p8, 0) }, 0, at);
                panel.Controls.Add(field, 1, at);
                var button = new Label { Text = "", AutoSize = false, Size = new Size(Ui.Px(this, 28), Ui.Px(this, 26)), Cursor = Cursors.Hand, Margin = new Padding(0, p4, 0, p4) };
                button.Paint += delegate(object sender, PaintEventArgs e) { PaintPickIcon(e.Graphics, button.ClientRectangle, e.Graphics.DpiX / 96f, false, true); };
                button.Click += delegate
                {
                    keep(); drop.Close();
                    if (RangePickRequested == null) return;
                    RangePickRequested(false, side.Books[0], side.Sheet, field.Text, delegate(string address)
                    {
                        var area = MigrationArea.Parse(address);
                        picked(single ? new MigrationArea { Row = area.Row, Column = area.Column, Rows = 1, Columns = 1 } : area);
                        ShowTriggerPopup(grid, pair, value);
                    });
                };
                restoreTip.SetToolTip(button, "在工作表中框选");
                panel.Controls.Add(button, 2, at);
            };
            line("触发单元格 1", cellBox, 1, true, delegate(MigrationArea area) { value.Cell = area.Address; });
            line("触发值 1", valuesBox, 2, false, delegate(MigrationArea area) { value.Values = area.Address; });
            int next = 3;
            for (int i = 0; i < value.Extra.Count; i++)
            {
                var condition = value.Extra[i]; int index = i, number = i + 2;
                var c = box(); c.Text = condition.Cell;
                var v = box(); v.Text = condition.Range.Length > 0 ? condition.Range : condition.Column.Length > 0 ? condition.Column : condition.Value;
                restoreTip.SetToolTip(v, "一列区域，与触发值 1 同行逐行取值；只选一个单元格则每一行都用它的值");
                extraCells.Add(c); extraValues.Add(v);
                line("触发单元格 " + number, c, next, true, delegate(MigrationArea area) { condition.Cell = area.Address; });
                var remove = new LinkLabel { Text = "删除", AutoSize = true, LinkBehavior = LinkBehavior.HoverUnderline, LinkColor = Ui.InkDim, ActiveLinkColor = Ui.Danger, Margin = new Padding(p4, Ui.Px(this, 8), 0, 0), BackColor = Color.White };
                remove.LinkClicked += delegate { keep(); drop.Close(); value.Extra.RemoveAt(index); ShowTriggerPopup(grid, pair, value); };
                restoreTip.SetToolTip(remove, "删除触发单元格 " + number + " 和触发值 " + number);
                panel.Controls.Add(remove, 3, next); next++;
                line("触发值 " + number, v, next, false, delegate(MigrationArea area) { condition.Range = area.Address; condition.Column = ""; condition.Value = ""; }); next++;
            }
            var addCondition = new LinkLabel { Text = "＋ 添加触发单元格", AutoSize = true, LinkBehavior = LinkBehavior.HoverUnderline, Margin = new Padding(0, p4, 0, p4), BackColor = Color.White };
            addCondition.LinkClicked += delegate { keep(); drop.Close(); value.Extra.Add(new MigrationCondition()); ShowTriggerPopup(grid, pair, value); };
            restoreTip.SetToolTip(addCondition, "再加一组：例如年度、期间，每一行重算前和触发值 1 一起填入");
            panel.Controls.Add(addCondition, 1, next); next++;
            line("目标文件区域", nameBox, next, false, delegate(MigrationArea area) { value.NameColumn = area.Address; }); next++;
            int others = OtherTriggerSets(side);
            var note = new Label { Text = "触发值依次填入触发单元格，并执行重算操作后再复制；目标文件区域内的文件名会自动匹配所选择的工作簿，只对匹配到的文件执行复制，忽略缺失或溢出文件。"
                + (value.Extra.Count > 0 ? Environment.NewLine + "多组触发单元格在每一行重算前一并填入，取同一行的值；某组触发值只选一个单元格时，每一行都用它的值。" : "")
                + (others > 0 ? Environment.NewLine + "这张源表已有另外 " + others + " 套逐行重算设置，各自独立。" : ""), AutoSize = true, MaximumSize = new Size(Ui.Px(this, 320), 0), ForeColor = Ui.InkDim, Margin = new Padding(0, p8, 0, 0) };
            panel.Controls.Add(note, 0, next); panel.SetColumnSpan(note, 3);
            var error = new Label { AutoSize = true, MaximumSize = new Size(Ui.Px(this, 320), 0), ForeColor = Ui.Danger, Margin = new Padding(0, p4, 0, 0) };
            panel.Controls.Add(error, 0, next + 1); panel.SetColumnSpan(error, 3);
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Dock = DockStyle.Fill, Margin = new Padding(0, p8, 0, 0), BackColor = Color.White };
            Func<string, bool, RoundedButton> make = delegate(string text, bool primary)
            {
                var b = Ui.Button(text, 72); if (primary) Ui.StylePrimary(b); else Ui.StyleSecondary(b);
                b.Font = font; b.MinimumSize = Size.Empty; b.Size = new Size(Ui.Px(this, text.Length > 2 ? 110 : 72), Ui.Px(this, 30)); b.Margin = new Padding(p8, 0, 0, 0); return b;
            };
            var ok = make("应用", true); var cancel = make("取消应用", false);
            buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
            panel.Controls.Add(buttons, 0, next + 2); panel.SetColumnSpan(buttons, 3);
            ok.Click += delegate
            {
                keep();
                try
                {
                    var cell = MigrationPlanner.TryParse(value.Cell); if (cell == null || cell.Rows != 1 || cell.Columns != 1) throw new InvalidOperationException("触发单元格 1 只能是一个单元格，如 B2");
                    var values = MigrationPlanner.TryParse(value.Values); if (values == null || values.Columns != 1) throw new InvalidOperationException("触发值 1 应为一列区域，如 G2:G30");
                    var names = MigrationPlanner.TryParse(value.NameColumn);
                    if (names == null || names.Columns != 1) throw new InvalidOperationException("目标文件区域应为一列区域，如 H2:H30");
                    if (names.Column == values.Column) throw new InvalidOperationException("目标文件区域不能和触发值 1 是同一列");
                    if (names.Row > values.Row || names.Row + names.Rows < values.Row + values.Rows) throw new InvalidOperationException("目标文件区域应与触发值 1 同行（覆盖 " + values.Row + "–" + (values.Row + values.Rows - 1) + " 行）");
                    var source = MigrationPlanner.TryParse(side.Range);
                    if (source != null && cell.Row >= source.Row && cell.Row < source.Row + source.Rows && cell.Column >= source.Column && cell.Column < source.Column + source.Columns)
                        throw new InvalidOperationException("触发单元格不能在源区域里");
                    var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { cell.Address };
                    for (int i = 0; i < value.Extra.Count; i++)
                    {
                        var condition = value.Extra[i]; int number = i + 2;
                        var target = MigrationPlanner.TryParse(condition.Cell);
                        if (target == null || target.Rows != 1 || target.Columns != 1) throw new InvalidOperationException("触发单元格 " + number + " 只能是一个单元格，如 C2");
                        if (source != null && target.Row >= source.Row && target.Row < source.Row + source.Rows && target.Column >= source.Column && target.Column < source.Column + source.Columns)
                            throw new InvalidOperationException("触发单元格 " + number + " 不能在源区域里");
                        if (!used.Add(target.Address)) throw new InvalidOperationException("触发单元格 " + number + "（" + target.Address + "）已被使用");
                        condition.Cell = target.Address;
                        if (condition.Range.Length > 0)
                        {
                            var range = MigrationPlanner.TryParse(condition.Range);
                            if (range == null || range.Columns != 1) throw new InvalidOperationException("触发值 " + number + " 应为一列区域或一个单元格");
                            bool one = range.Rows == 1;
                            if (!one && (range.Row > values.Row || range.Row + range.Rows < values.Row + values.Rows)) throw new InvalidOperationException("触发值 " + number + " 应与触发值 1 同行（覆盖 " + values.Row + "–" + (values.Row + values.Rows - 1) + " 行），或只选一个单元格");
                            condition.Range = range.Address;
                        }
                        else if (condition.Column.Length == 0 && condition.Value.Length == 0) throw new InvalidOperationException("请填写触发值 " + number);
                    }
                    value.Cell = cell.Address; value.Values = values.Address; value.NameColumn = names.Address; value.Signature = "";
                }
                catch (Exception failure) { error.Text = failure.Message; return; }
                drop.Close(); SetTrigger(side, value);
            };
            // 取消应用: remove this range's 逐行重算; without one, just close.
            cancel.Click += delegate { drop.Close(); if (side.Trigger != null) SetTrigger(side, null); };
            FramePopup(panel); panel.PerformLayout();
            var host = new ToolStripControlHost(panel) { AutoSize = false, Margin = Padding.Empty, Padding = Padding.Empty, Size = panel.PreferredSize };
            drop.Items.Add(host);
            error.SizeChanged += delegate { panel.PerformLayout(); host.Size = panel.PreferredSize; drop.Size = panel.PreferredSize; };
            drop.Closed += delegate { if (triggerPopup == drop) triggerPopup = null; BeginInvoke(new Action(delegate { drop.Dispose(); font.Dispose(); })); };
            triggerPopup = drop;
            // Production controls can be rendered in the isolated preview without opening a native popup.
            if (DialogPreviewFactory.IsConstructing) return;
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
                var item = new ToolStripMenuItem(Path.GetFileName(path) + "    " + Path.GetDirectoryName(path)) { Padding = new Padding(h, v, h, v), Checked = string.Equals(row.Path, path, StringComparison.OrdinalIgnoreCase) && row.Ready };
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
                "映射第 " + (pairs.IndexOf(pair) + 1) + " 行（" + (pair.Enabled ? "已启用" : "未启用") + "，" + MigrationPair.ModeText(pair.ValuesOnly || pair.Source.Trigger != null, pair.KeepFormat, pair.FormatOnly && pair.Source.Trigger == null) + "）",
                "源：" + Where(pair.Source, MigrationPlanner.TryParse(pair.Source.Range)) + "　工作簿：" + Books(pair.Source),
                "目标：" + Where(pair.Target, MigrationPlanner.TargetArea(pair)) + "　工作簿：" + Books(pair.Target),
                "状态：" + PairStatus(pair)
            };
            foreach (var m in pair.Instances.Where(x => x.Result != null))
            {
                var t = m.Result;
                lines.Add("· " + (t.Finished == DateTime.MinValue ? "" : DataMigrationHost.Stamp(t.Finished) + "  ") + Path.GetFileName(m.Source.Path) + " → " + Path.GetFileName(m.TargetPath) + "：" + t.Status + (t.Error.Length == 0 ? "" : "：" + t.Error)
                    + (t.Overwritten > 0 ? "；覆盖 " + t.Overwritten + " 格" : "") + (t.Output.Length == 0 ? "" : "；输出 " + t.Output) + (t.Backup.Length == 0 ? "" : "；备份 " + t.Backup));
                foreach (string sample in t.Samples.Take(3)) lines.Add("    " + sample);
            }
            detailLinks.Clear(); unusedLinks.Clear();
            var trigger = pair.Source.Trigger;
            if (trigger != null)
            {
                lines.Add("逐行重算：触发值 " + trigger.Values + " → " + trigger.Cell + string.Concat(trigger.Extra.Select(c => "，" + c.Text).ToArray()) + "，目标文件区域 " + trigger.NameColumn + "；目标按文件名在已添加的 " + TargetBooks.Count + " 个目标工作簿中匹配");
                if (trigger.Error.Length > 0) lines.Add("    设置无效：" + trigger.Error);
                foreach (var row in trigger.Rows)
                {
                    string mark = row.Kind == "精确" ? "✓ 已匹配" : row.Kind == "已确认" ? "✓ 已指定"
                        : row.Kind == "多个候选" ? "⚠ 有多个同名工作簿（点此选择）" : row.Kind == "缺失" ? "✗ 没有同名的目标工作簿，将跳过（点此手动指定）"
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
        internal MigrationPlan BuildPlan()
        {
            EndEdits();
            foreach (var side in DistinctSides(false)) RefreshTrigger(side);
            var snapshot = new MigrationConfiguration(pairs, SourceBooks, TargetBooks, replaceTargets.Checked, allowEmpty.Checked);
            var plan = new MigrationPlanner(snapshot).Build();
            snapshot.Attach(plan); lastPlan = plan; RefreshRows(); ShowPairDetails();
            if (plan.Passes.Count == 0) throw new InvalidOperationException(plan.Mappings.Count == 0 ? "没有可预检的映射：请在“映射”页查看各行“状态”提示的问题。" : "所有写入均未通过检查，请在“映射”页查看各行“状态”。");
            return plan;
        }
        private string RowProblem(MigrationPair pair) { return MigrationPlanner.RowProblem(pair, TargetBooks.Count); }
        private List<string> UnusedTargets(MigrationSide side)
        { MatchView view; return matches.TryGetValue(side, out view) ? view.Unused : new List<string>(); }
        internal static string PlanSummary(MigrationPlan plan)
        {
            var results = plan.Mappings.Where(m => m.Result != null).Select(m => m.Result).ToList();
            int ok = results.Count(t => t.Status == "成功"), failed = results.Count(t => t.Status == "失败" || t.Status == "冲突"), ready = results.Count(t => t.Ready && t.Status == "可执行");
            int rows = plan.Mappings.Select(m => m.Pair).Distinct().Count(), skipped = plan.Skipped.Count, targets = plan.Mappings.Select(m => m.TargetPath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            return "对照 " + rows + " 行，目标工作簿 " + targets + " 个，共 " + plan.Mappings.Count + " 处写入：成功 " + ok + "，失败/冲突 " + failed + "，可执行 " + ready + "。"
                + (skipped > 0 ? "逐行重算跳过 " + skipped + " 行（见明细）。" : "")
                + (plan.Unused.Count > 0 ? "另有 " + plan.Unused.Count + " 个目标工作簿没有触发行匹配，不会写入（见明细，可点击指定）。" : "")
                + (plan.ReplaceTargets ? "保存方式：校验后更新原文件并备份。" : "保存方式：另存结果副本。")
                + (plan.Finished == DateTime.MinValue ? "" : plan.Executed
                    ? "执行完成于 " + DataMigrationHost.Stamp(plan.Finished) + "，用时 " + DataMigrationHost.DurationText(plan.RunMilliseconds) + "。"
                    : "预检完成于 " + DataMigrationHost.Stamp(plan.Finished) + (ready > 0 ? "，预计执行约 " + DataMigrationHost.DurationText(DataMigrationHost.EstimateMilliseconds(plan.Passes.Select(p => p.Request), plan.PrecheckMilliseconds)) : "") + "。");
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
        private static void AddGridColumn(DataGridView grid, string name, string title, int width, bool readOnly)
        { grid.Columns.Add(new DataGridViewTextBoxColumn { Name = name, HeaderText = title, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable }); }
        private static void PrepareGrid(DataGridView grid)
        {
            grid.AllowUserToAddRows = false; grid.AllowUserToDeleteRows = false; grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; grid.MultiSelect = true;
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            grid.CurrentCellDirtyStateChanged += delegate { if (grid.IsCurrentCellDirty && grid.CurrentCell is DataGridViewCheckBoxCell) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
        }
        private static string SizeText(string range)
        {
            try { var area = MigrationArea.Parse(MigrationPlanner.NormalizeAddress(range)); return area.Rows + " 行 × " + area.Columns + " 列"; }
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
        // The mappings that passed preview, regrouped with fresh state so
        // each pass re-checks its targets right before writing.
        internal MigrationPlan ExecutionPlan(MigrationPlan previewed)
        {
            var plan = new MigrationPlan { ReplaceTargets = previewed.ReplaceTargets, AllowEmptySource = previewed.AllowEmptySource, OutputCopies = !previewed.ReplaceTargets };
            var ready = previewed.Mappings.Where(m => m.Result != null && m.Result.Ready && m.Result.Status == "可执行").ToList();
            foreach (var m in ready) { m.Result = new MigrationTarget { Path = m.TargetPath, Sheet = m.TargetSheet, Start = m.Start, Status = "待执行" }; plan.Mappings.Add(m); }
            // Every pass updates its target in place. For result copies the
            // orchestrator first points each target at one "_复制结果" copy,
            // so all source ranges land in the same copy.
            MigrationPlanner.Group(plan, ready, true);
            lastPlan = plan; RefreshRows();
            return plan;
        }
        internal void RecordRestorePoint(string path, string backup)
        {
            try { recovery.Record(path, backup); }
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
            restore.Enabled = !Running && recovery.Pending.Count > 0;
            clearBackups.Enabled = !Running && recovery.Backups.Count > 0;
        }
        internal void ClearBackups(bool confirm = true)
        {
            if (Running) return;
            var pending = recovery.Backups;
            if (pending.Count == 0) { UpdateRestore(); return; }
            int files = pending.Sum(p => MigrationRecovery.Files(p).Count(File.Exists));
            int restorable = pending.Count(p => !p.Restored);
            string question = "将永久删除本窗口复制时生成的 " + files + " 个备份文件（涉及 " + pending.Count + " 个目标工作簿）。" + Environment.NewLine + Environment.NewLine
                + (restorable > 0 ? "删除后，这 " + restorable + " 个已更新的工作簿将无法再还原为复制前的版本。" : "删除后无法恢复这些备份。") + Environment.NewLine
                + "此操作无法撤销。" + Environment.NewLine + Environment.NewLine + "是否清除备份？";
            if (confirm && UiDialog.Show(this, question, "清除备份", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            var results = MigrationRecovery.Clear(pending);
            int cleared = results.Count(p => p.Value.Length == 0);
            var failures = results.Where(p => p.Value.Length > 0).Select(p => Path.GetFileName(p.Key.Path) + "：" + p.Value).ToList();
            UpdateRestore();
            Message((cleared > 0 ? "已清除 " + cleared + " 个目标的备份，这些工作簿不能再还原。" : "未清除任何备份。")
                + (failures.Count == 0 ? "" : " 未清除：" + string.Join("；", failures.ToArray())));
        }
        internal int RestorableCount { get { return recovery.RestorableCount; } }
        internal void RestoreTargets(bool confirm = true)
        {
            if (Running) return;
            // Newest first: a file migrated twice steps back one batch at a time.
            var pending = recovery.Pending;
            if (pending.Count == 0) { UpdateRestore(); return; }
            string question = "将把 " + pending.Count + " 个已更新的目标工作簿还原为复制前的备份。" + Environment.NewLine + Environment.NewLine
                + "· 请先在 WPS 中关闭这些工作簿。" + Environment.NewLine
                + "· 复制后的版本会另存到对应目录的“复制备份”文件夹，不会丢失。" + Environment.NewLine
                + "· 复制后又被修改过的文件不会还原，以免覆盖这些修改。" + Environment.NewLine + Environment.NewLine + "是否继续？";
            if (confirm && UiDialog.Show(this, question, "还原操作", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            var results = MigrationRecovery.Restore(pending);
            int restored = results.Count(p => p.Value.Length == 0);
            var failures = results.Where(p => p.Value.Length > 0).Select(p => Path.GetFileName(p.Key.Path) + "：" + p.Value).ToList();
            foreach (var result in results)
                foreach (var target in AllInstances().Select(m => m.Result).Where(t => t != null && string.Equals(t.Path, result.Key.Path, StringComparison.OrdinalIgnoreCase)))
                {
                    if (result.Value.Length == 0) { target.Status = "已还原"; target.Error = ""; target.Ready = false; }
                    else target.Error = "还原失败：" + result.Value;
                }
            RefreshResults(); UpdateRestore();
            Message((restored > 0 ? "已还原 " + restored + " 个工作簿为复制前的备份；如需再次复制请重新预检。" : "未还原任何工作簿。")
                + (failures.Count == 0 ? "" : " 未还原：" + string.Join("；", failures.ToArray())));
        }
        internal Dictionary<string, object> CapturePreset(string name)
        {
            EndEdits();
            return MigrationPresetCodec.Capture(name, new MigrationConfiguration(pairs, SourceBooks, TargetBooks, replaceTargets.Checked, allowEmpty.Checked));
        }
        internal void ApplyPresetData(Dictionary<string, object> data, IDictionary<string, string> remap)
        {
            if (Running) throw new InvalidOperationException("复制运行期间不能应用预设。");
            var configuration = MigrationPresetCodec.Read(data, remap);
            ClosePopups(); EndEdits();
            pairs.Clear(); SourceBooks.Clear(); TargetBooks.Clear(); mappingFocus = null; lastPlan = null;
            syncing = true;
            try { replaceTargets.Checked = configuration.ReplaceTargets; allowEmpty.Checked = configuration.AllowEmptySource; }
            finally { syncing = false; }
            LoadPresetBooks(false, configuration.SourceBooks.Select(b => b.Path));
            LoadPresetBooks(true, configuration.TargetBooks.Select(b => b.Path));
            pairs.AddRange(configuration.Pairs);
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
            return (source == null ? "复制预设" : Path.GetFileNameWithoutExtension(source.Path)) + " " + DateTime.Now.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture).Replace(':', '.');
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
            if (Running) { Message("复制运行期间不能管理预设。"); return; }
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

    // A saved migration setup (settings only, never sheet data), kept as one
    // .xymigration.json file. The local library is a folder of these files;
    // export copies one out, import copies one in.

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
            PresetNameView.Build(this, Input, prompt, value, delegate { return true; });
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
        private readonly PresetManagerView view;
        internal MigrationPreset Chosen;
        internal MigrationPresetManager(string select)
        {
            Text = "预设管理"; DialogSizing.Prepare(this, "migration.preset.manager", "table", Ui.Blue);
            ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent; MinimizeBox = false;
            view = new PresetManagerView(this,list,detail,search,apply,rename,remove,export,
                item => ((MigrationPreset)item).Name,
                item => { var p=(MigrationPreset)item; return (p.Saved==DateTime.MinValue ? "" : p.Saved.ToString("MM-dd HH:mm",CultureInfo.InvariantCulture)+"　·　")+"源 "+p.Sources+"　目标 "+p.Targets+"　映射 "+p.Mappings; });
            var import=view.Import; var close=view.Close;
            list.SelectedIndexChanged += delegate { ShowDetail(); };
            list.DoubleClick += delegate { if (Current != null) Apply(); };
            list.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Delete) Remove(); else if (e.KeyCode == Keys.F2) Rename(); };
            detail.HandleCreated += delegate { BeginInvoke(new Action(ShowDetail)); };
            search.TextChanged += delegate { Fill(Current == null ? null : Current.FilePath); };
            import.Click += delegate { Import(); }; export.Click += delegate { Export(); };
            rename.Click += delegate { Rename(); }; remove.Click += delegate { Remove(); };
            apply.Click += delegate { Apply(); }; close.Click += delegate { DialogResult = DialogResult.Cancel; };
            Reload(select);
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
                SetDetail(all.Count == 0 ? "还没有保存的预设。" + Environment.NewLine + Environment.NewLine + "在批量复制窗口右上角的齿轮里点“保存为预设…”，或点上方“导入预设…”导入别人分享的预设文件。" : "没有名称匹配的预设。");
                return;
            }
            ShowLines(preset.DescribeLines());
        }
        private void ShowLines(List<PresetDetailLine> lines) { view.ShowLines(lines); }
        private void SetDetail(string text) { view.SetDetail(text); }
        private void Import()
        {
            string[] files = HostDialogService.OpenFiles(this, "复制预设 (*" + MigrationPreset.Extension + ")|*" + MigrationPreset.Extension + "|JSON 文件 (*.json)|*.json", "导入预设", true, "");
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
            string chosen = HostDialogService.SaveFile(this, "导出预设", "复制预设 (*" + MigrationPreset.Extension + ")|*" + MigrationPreset.Extension, Path.GetFileName(preset.FilePath));
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
    // Read-only snapshot of one worksheet, read straight from the .xlsx/.xlsm
    // package: no WPS window opens and nothing is written. Only text needed to
    // recognise a position is kept (values, merges, row/column sizes).

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
                e.ToolTipText = CellAddress.Name(e.RowIndex + 1, e.ColumnIndex + 1) + (owner.Length == 0 ? "" : "  ·  已分配给：" + owner);
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
                list[c - 1] = new DataGridViewTextBoxColumn { HeaderText = CellAddress.ColumnLetters(c), SortMode = DataGridViewColumnSortMode.NotSortable, FillWeight = 1, Width = Math.Max(24, (int)Math.Round(Math.Min(width, 400) * zoom)), MinimumWidth = 10 };
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
                var cell = MigrationArea.Parse(MigrationPlanner.NormalizeAddress(address.Text));
                if (cell.Row > grid.RowCount || cell.Column > grid.ColumnCount) { hint.Text = "地址超出预览范围（预览到第 " + grid.RowCount + " 行、" + CellAddress.ColumnLetters(grid.ColumnCount) + " 列）。"; return; }
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
            hint.Text = "找到 " + CellAddress.Name(row, column) + "：" + snapshot.Value(row, column) + "。在它附近单击即可放置。";
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


