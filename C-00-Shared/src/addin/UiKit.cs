using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Shared WinForms UI, compiled into both the add-in and its standalone installer.
    // Both products compile this single implementation from C-00-Shared.
    internal interface IGridColumnFit { bool TryFitColumns(); }
    internal static class Ui
    {
        private sealed class ColumnMeasure
        {
            internal int Version,Measured=-1;
            internal string Signature;
            internal ColumnMeasure(DataGridView grid)
            {
                grid.CellValueChanged+=delegate{Version++;};grid.RowsAdded+=delegate{Version++;};grid.RowsRemoved+=delegate{Version++;};
                grid.ColumnAdded+=delegate{Version++;};grid.ColumnRemoved+=delegate{Version++;};grid.FontChanged+=delegate{Version++;};
                grid.CellStyleContentChanged+=delegate{Version++;};grid.DefaultCellStyleChanged+=delegate{Version++;};
                grid.ColumnDefaultCellStyleChanged+=delegate{Version++;};grid.ColumnHeadersDefaultCellStyleChanged+=delegate{Version++;};
            }
        }
        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<DataGridView,ColumnMeasure> columnMeasures=new System.Runtime.CompilerServices.ConditionalWeakTable<DataGridView,ColumnMeasure>();
        [ThreadStatic] private static HashSet<DataGridView> updatingGridHeaders;
        private const TextFormatFlags GridHeaderTextFlags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix;
        public static readonly Color Blue = Color.FromArgb(37, 99, 235);
        public static readonly Color BlueHover = Color.FromArgb(29, 78, 216);
        public static readonly Color BlueDown = Color.FromArgb(30, 64, 175);
        public static readonly Color Amber = Color.FromArgb(245, 158, 11);
        public static readonly Color Green = Color.FromArgb(22, 135, 111);
        public static readonly Color GreenHover = Color.FromArgb(17, 112, 91);
        public static readonly Color GreenDown = Color.FromArgb(13, 91, 74);
        public static readonly Color Danger = Color.FromArgb(196, 65, 82);
        public static readonly Color DangerHover = Color.FromArgb(174, 50, 66);
        public static readonly Color DangerDown = Color.FromArgb(148, 39, 54);
        public static readonly Color Canvas = Color.FromArgb(245, 247, 251);
        public static readonly Color BarBack = Canvas;
        public static readonly Color GridHead = Color.FromArgb(247, 249, 252);
        public static readonly Color GridHeadInk = Color.FromArgb(80, 95, 117);
        public static readonly Color GridLine = Color.FromArgb(239, 242, 247);
        public static readonly Color GridFrame = Color.FromArgb(228, 233, 242);
        public static readonly Color SelectBack = Color.FromArgb(234, 241, 255);
        public static readonly Color SelectInk = Color.FromArgb(30, 64, 140);
        public static readonly Color ZebraBack = Color.FromArgb(252, 253, 255);
        public static readonly Color SelfRowBack = Color.FromArgb(255, 250, 236);
        public static readonly Color SelfRowSel = Color.FromArgb(253, 239, 196);
        public static readonly Color Ink = Color.FromArgb(36, 50, 71);
        public static readonly Color InkDim = Color.FromArgb(102, 112, 133);
        public static readonly Color Btn2Back = Color.White;
        public static readonly Color Btn2Hover = Color.FromArgb(242, 246, 255);
        public static readonly Color Btn2Down = Color.FromArgb(229, 237, 252);
        public static readonly Color Btn2Border = GridFrame;

        public static Font MakeFont(float size, FontStyle style, Font fallback)
        {
            return FindFont(new string[] { "Microsoft YaHei UI", "Microsoft YaHei", "微软雅黑", "Segoe UI" }, size, style, fallback);
        }

        public static Font MonoFont(float size, Font fallback)
        {
            return FindFont(new string[] { "Consolas", "Cascadia Mono", "Microsoft YaHei UI" }, size, FontStyle.Regular, fallback);
        }

        private static Font FindFont(string[] names, float size, FontStyle style, Font fallback)
        {
            foreach (string name in names)
            {
                try
                {
                    using (FontFamily family = new FontFamily(name))
                        return new Font(family, size, family.IsStyleAvailable(style) ? style : FontStyle.Regular, GraphicsUnit.Point);
                }
                catch { }
            }
            return new Font((fallback ?? SystemFonts.MessageBoxFont).FontFamily, size, style, GraphicsUnit.Point);
        }

        public static int Px(Control control, int value)
        {
            using (Graphics g = control.CreateGraphics()) return (int)Math.Round(value * g.DpiX / 96f);
        }

        public static void PrepareForm(Form form, Size preferred, Size minimum)
        {
#if WPS_WORD
            PrepareForm(form, preferred, minimum, "document", Blue);
#else
            PrepareForm(form, preferred, minimum, "table", Blue);
#endif
        }

        internal static Func<Form, string, bool> DialogSizeApplier;
        internal static Action<Form,string> DialogFormatApplier;
        internal static Func<Form,bool> PreviewActivationGuard;
        internal static Func<IWin32Window,string,string,MessageBoxButtons,MessageBoxIcon,MessageBoxDefaultButton,DialogResult> HostDialogPresenter;
        // Runtime width drags persist through this injected callback; the installer (which
        // compiles UiKit but not DialogFormatting) leaves it null, so grips simply skip saving.
        internal static Action<Form,Control,int> SaveControlWidth;
        // Runtime row-height drags persist through this injected callback. The key is a
        // synthetic id (e.g. "row.sources") owned by the dialog's layout, not by a control.
        internal static Action<Form,string,int> SaveRowHeight;
        internal static bool IsInertPreview(Form form) {return PreviewActivationGuard!=null&&PreviewActivationGuard(form);}

        public static void PrepareForm(Form form, Size preferred, Size minimum, string iconKind, Color accent, string dialogId = null)
        {
            form.AutoScaleDimensions = new SizeF(96f, 96f);
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.Font = MakeFont(10f, FontStyle.Regular, null);
            form.BackColor = Canvas;
            form.ClientSize = preferred;
            form.MinimumSize = minimum;
            form.StartPosition = FormStartPosition.CenterScreen;
            UiWindowChrome.Attach(form, iconKind, accent);
            if(dialogId!=null&&DialogFormatApplier!=null)DialogFormatApplier(form,dialogId);
            bool managed=dialogId!=null&&DialogSizeApplier!=null&&DialogSizeApplier(form,dialogId);
            if(!managed)form.Shown += delegate
            {
                Rectangle area = Screen.FromControl(form).WorkingArea;
                // Host process DPI awareness is owned by WPS; never change it globally.
                form.MinimumSize = new Size(Math.Min(form.MinimumSize.Width, area.Width - 24), Math.Min(form.MinimumSize.Height, area.Height - 24));
                form.Size = new Size(Math.Min(form.Width, area.Width - 24), Math.Min(form.Height, area.Height - 24));
                form.Location = new Point(area.Left + (area.Width - form.Width) / 2, area.Top + (area.Height - form.Height) / 2);
            };
        }

        public static Label Label(string text, float size, bool bold)
        {
            Label label = new Label();
            label.Text = text;
            label.Font = MakeFont(size, bold ? FontStyle.Bold : FontStyle.Regular, null);
            label.ForeColor = Ink;
            label.AutoSize = false;
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            label.AutoEllipsis = true;
            label.Margin = Padding.Empty;
            return label;
        }

        public static UiCard Card(Control content, string title)
        {
            UiCard card = new UiCard();
            card.Dock = DockStyle.Fill;
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowCount = string.IsNullOrEmpty(title) ? 1 : 2;
            if (!string.IsNullOrEmpty(title))
            {
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
                layout.Controls.Add(Label(title, 10f, true), 0, 0);
            }
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            content.Dock = DockStyle.Fill;
            content.Margin = Padding.Empty;
            layout.Controls.Add(content, 0, layout.RowCount - 1);
            card.Controls.Add(layout);
            return card;
        }

        public static FlowLayoutPanel ActionBar()
        {
            FlowLayoutPanel bar = new FlowLayoutPanel();
            bar.Dock = DockStyle.Fill;
            bar.AutoSize = true;
            bar.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            bar.WrapContents = true;
            bar.Margin = Padding.Empty;
            bar.Padding = new Padding(0, 4, 0, 0);
            bar.BackColor = Canvas;
            return bar;
        }

        public static RoundedButton Button(string text, int width)
        {
            RoundedButton button = new RoundedButton();
            button.Text = text;
            button.Size = new Size(width, 38);
            StyleSecondary(button);
            return button;
        }

        public static void StylePrimary(RoundedButton button) { StylePrimary(button, false); }
        public static void StylePrimary(RoundedButton button, bool green)
        {
            button.BaseColor = green ? Green : Blue;
            button.HoverColor = green ? GreenHover : BlueHover;
            button.DownColor = green ? GreenDown : BlueDown;
            button.ForeColor = Color.White;
            button.BorderColor = Color.Empty;
            button.Font = MakeFont(10f, FontStyle.Regular, button.Font);
        }

        public static void StyleSecondary(RoundedButton button)
        {
            button.BaseColor = Btn2Back;
            button.HoverColor = Btn2Hover;
            button.DownColor = Btn2Down;
            button.ForeColor = Ink;
            button.BorderColor = Btn2Border;
            button.Font = MakeFont(10f, FontStyle.Regular, button.Font);
        }

        public static void StyleDanger(RoundedButton button)
        {
            button.BaseColor = Danger;
            button.HoverColor = DangerHover;
            button.DownColor = DangerDown;
            button.ForeColor = Color.White;
            button.BorderColor = Color.Empty;
            button.Font = MakeFont(10f, FontStyle.Regular, button.Font);
        }

        public static void StyleGrid(DataGridView grid)
        {
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = 38;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            grid.ColumnHeadersDefaultCellStyle.BackColor = GridHead;
            grid.ColumnHeadersDefaultCellStyle.ForeColor = GridHeadInk;
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = GridHead;
            grid.ColumnHeadersDefaultCellStyle.SelectionForeColor = GridHeadInk;
            grid.ColumnHeadersDefaultCellStyle.Font = MakeFont(9f, FontStyle.Bold, null);
            grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 0, 8, 0);
            grid.RowTemplate.Height = 34;
            grid.DefaultCellStyle.BackColor = Color.White;
            grid.DefaultCellStyle.ForeColor = Ink;
            grid.DefaultCellStyle.Font = MakeFont(10f, FontStyle.Regular, null);
            grid.DefaultCellStyle.Padding = new Padding(8, 0, 8, 0);
            grid.DefaultCellStyle.SelectionBackColor = SelectBack;
            grid.DefaultCellStyle.SelectionForeColor = SelectInk;
            grid.AlternatingRowsDefaultCellStyle.BackColor = ZebraBack;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.GridColor = GridLine;
            grid.HandleCreated += delegate
            {
                grid.ColumnHeadersHeight = Px(grid, 38);
                grid.RowTemplate.Height = Px(grid, 34);
                if(!grid.VirtualMode)foreach (DataGridViewRow row in grid.Rows) row.Height = Px(grid, 34);
                int padding = Px(grid, 8);
                grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(padding, 0, padding, 0);
                grid.DefaultCellStyle.Padding = new Padding(padding, 0, padding, 0);
                AutoFitColumns(grid);
            };
        }

        public static void AutoFitColumns(DataGridView grid)
        {
            // Disposal clears the columns one by one and can raise layout and
            // row events on the way; a grid being torn down is never fitted.
            if (grid == null || grid.IsDisposed || grid.Disposing || grid.Columns.Count == 0) return;
            var local = grid as IGridColumnFit;
            if (local != null && local.TryFitColumns()) return;
            EnsureGridHeaders(grid);
            float scale;
            using (Graphics g = grid.CreateGraphics())
                scale = Math.Max(g.DpiX / 96f, (grid.DefaultCellStyle.Font ?? grid.Font).Size / 10f);
            var measure=columnMeasures.GetValue(grid,g=>new ColumnMeasure(g));
            var signature=new System.Text.StringBuilder(scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach(DataGridViewColumn column in grid.Columns)signature.Append('|').Append(column.HeaderText).Append(':').Append(column.MinimumWidth).Append(':').Append(column.Tag).Append(':').Append(column.AutoSizeMode);
            string stamp=signature.ToString();
            if(measure.Measured==measure.Version&&measure.Signature==stamp)return;
            foreach (DataGridViewColumn col in Attached(grid))
            {
                int need = col.MinimumWidth;
                col.MinimumWidth = Math.Max((int)(50 * scale), need);
                if (col.AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill) continue;
                grid.AutoResizeColumn(col.Index, DataGridViewAutoSizeColumnMode.AllCells);
                int cap = col.Tag is int ? Math.Max(need, (int)((int)col.Tag * scale)) : int.MaxValue;
                col.Width = Math.Max(col.MinimumWidth, Math.Min(col.Width, cap));
            }
            measure.Measured=measure.Version;
            // Minimum widths may grow during this pass; cache the resulting state.
            signature.Clear().Append(scale.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach(DataGridViewColumn column in grid.Columns)signature.Append('|').Append(column.HeaderText).Append(':').Append(column.MinimumWidth).Append(':').Append(column.Tag).Append(':').Append(column.AutoSizeMode);
            measure.Signature=signature.ToString();
        }

        // Headers are single-line labels. Their font may be explicitly customized,
        // including regular weight; width caps apply to body text, never to a header.
        // This helper has no dependency on an add-in form or ResultGrid, since UiKit
        // is also compiled into the standalone installer.
        // A snapshot of the columns still attached to the grid: a column being
        // removed has no DataGridView and its inherited style cannot be read.
        private static List<DataGridViewColumn> Attached(DataGridView grid)
        {
            var result = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn column in grid.Columns) if (column != null && column.DataGridView == grid) result.Add(column);
            return result;
        }
        internal static void EnsureGridHeaders(DataGridView grid)
        {
            if (grid == null || grid.IsDisposed || grid.Disposing) return;
            if (updatingGridHeaders == null) updatingGridHeaders = new HashSet<DataGridView>();
            if (!updatingGridHeaders.Add(grid)) return;
            try
            {
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
                int height = Px(grid, 38);
                using (Graphics graphics = grid.CreateGraphics())
                {
                    foreach (DataGridViewColumn column in Attached(grid))
                    {
                        column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        column.HeaderCell.Style.WrapMode = DataGridViewTriState.False;
                        var style = column.HeaderCell.InheritedStyle;
                        Font font = style.Font ?? grid.Font;
                        Size text = TextRenderer.MeasureText(graphics, column.HeaderText ?? "", font, Size.Empty, GridHeaderTextFlags);
                        int sorting = column.SortMode == DataGridViewColumnSortMode.NotSortable ? 0 : Px(grid, 18);
                        column.MinimumWidth = Math.Max(Px(grid, 50), text.Width + style.Padding.Horizontal + Px(grid, 12) + sorting);
                        height = Math.Max(height, text.Height + style.Padding.Vertical + Px(grid, 12));
                    }
                }
                grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                if (grid.ColumnHeadersHeight != height) grid.ColumnHeadersHeight = height;
            }
            finally { updatingGridHeaders.Remove(grid); }
        }

        internal static GraphicsPath RoundRect(RectangleF rect, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float d = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height));
            if (d <= 0) { path.AddRectangle(rect); return path; }
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        // Wraps an input with an optional leading label and a trailing WidthGrip inside a
        // non-wrapping FlowLayoutPanel. Every interactive dialog uses this so dragged
        // widths persist through DialogFormatting. Path-selection boxes stay ungripped.
        public static FlowLayoutPanel Grip(Control input, Control leading)
        {
            var field = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 12, 0) };
            // 按钮的 Margin.Bottom=8（RoundedButton 构造里设的）是为底部按钮栏准备的；
            // 放在行内会把这一格撑到 46px（同行的文本框格只有 34px），
            // 单行布局按最高格取高 → 每个输入组下方凭空多出约 12px 空白。这里清掉底边距。
            if (input is Button) input.Margin = new Padding(0, 0, 8, 0);
            else input.Margin = new Padding(3, 0, 3, 0);   // 收紧上下边距，行内不需要
            // 前导标签上边距原为 8，会把包装器撑得比输入框高出一截（输入框顶对齐后下方留白）
            if (leading != null) field.Controls.Add(leading);
            field.Controls.Add(input);
            field.Controls.Add(new WidthGrip(input));
            return field;
        }

        public static FlowLayoutPanel Grip(Control input, string label)
        {
            Label lead = string.IsNullOrEmpty(label) ? null : new Label { Text = label, AutoSize = true, Margin = new Padding(3, 6, 6, 2) };
            return Grip(input, lead);
        }
    }

    // A thin drag handle placed after an input; dragging left/right resizes the input
    // so long values can be revealed without touching the saved dialog formats.
    // Committing writes through DialogFormatting so any registered form persists.
    // Two placement modes: embedded (target in a FlowLayoutPanel, grip flows after it)
    // and floating (target anywhere else, grip joins target.Parent and tracks its edge).
    // A drag is provisional until left-button release; capture/lifetime loss restores it.
    internal sealed class WidthGrip : Control
    {
        private const int MinimumWidth = 40;
        private readonly Control target;
        private int anchorX, anchorWidth;
        private bool dragging, floating, repositioning;
        internal WidthGrip(Control target)
        {
            this.target = target;
            Width = 7; Margin = new Padding(0); Cursor = Cursors.SizeWE;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.Opaque | ControlStyles.ResizeRedraw, true);
            target.SizeChanged += OnTargetBoundsChanged;
            target.LocationChanged += OnTargetBoundsChanged;
            target.ParentChanged += OnTargetParentChanged;
            target.VisibleChanged += OnTargetVisibilityChanged;
            target.Disposed += OnTargetDisposed;
            Height = target.Height;
            Visible = target.Visible;
            OnTargetParentChanged(null, null);
        }
        private void OnTargetVisibilityChanged(object sender, EventArgs e) { if (!target.Visible) CancelDrag(); Visible = target.Visible; }
        private void OnTargetDisposed(object sender, EventArgs e) { CancelDrag(); }
        private void OnTargetParentChanged(object sender, EventArgs e)
        {
            CancelDrag();
            floating = !(target.Parent is FlowLayoutPanel);
            if (floating)
            {
                var host = target.Parent;
                if (host != null && Parent != host)
                {
                    if (Parent != null) Parent.Controls.Remove(this);
                    host.Controls.Add(this);
                    BringToFront();
                }
            }
            OnTargetBoundsChanged(null, null);
        }
        private void OnTargetBoundsChanged(object sender, EventArgs e)
        {
            if (IsDisposed || target == null) return;
            Height = target.Height;
            if (floating && target.Parent != null && !repositioning)
            {
                repositioning = true;
                try
                {
                    int x = target.Right;
                    int y = target.Top + (target.Height - Height) / 2;
                    if (Left != x || Top != y) SetBounds(x, y, Width, target.Height);
                }
                finally { repositioning = false; }
            }
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && !IsDisposed && !target.IsDisposed)
            {
                CancelDrag(); anchorX = PointToScreen(e.Location).X; anchorWidth = target.Width;
                dragging = true; Capture = true;
                if (!Capture) CancelDrag();
            }
            base.OnMouseDown(e);
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (dragging)
            {
                if (!Capture || (e.Button & MouseButtons.Left) == 0 || target.IsDisposed) CancelDrag();
                else target.Width = Math.Max(MinimumWidth, anchorWidth + PointToScreen(e.Location).X - anchorX);
            }
            base.OnMouseMove(e);
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (dragging)
            {
                if (e.Button == MouseButtons.Left && Capture && !target.IsDisposed)
                { dragging = false; Capture = false; CommitWidth(); }
                else CancelDrag();
            }
            base.OnMouseUp(e);
        }
        private void CommitWidth()
        {
            var form = FindForm();
            if (form == null || form.IsDisposed) return;
            if (Ui.SaveControlWidth == null) return;
            float scale;
            using (var g = form.CreateGraphics()) scale = g.DpiX / 96f;
            Ui.SaveControlWidth(form, target, (int)Math.Round(target.Width / scale));
        }
        private void CancelDrag()
        {
            if (!dragging) return;
            dragging = false;
            if (!target.IsDisposed) target.Width = anchorWidth;
            Capture = false;
        }
        protected override void OnMouseCaptureChanged(EventArgs e) { if (dragging && !Capture) CancelDrag(); base.OnMouseCaptureChanged(e); }
        protected override void OnVisibleChanged(EventArgs e) { if (!Visible) CancelDrag(); base.OnVisibleChanged(e); }
        protected override void OnEnabledChanged(EventArgs e) { if (!Enabled) CancelDrag(); base.OnEnabledChanged(e); }
        protected override void OnParentChanged(EventArgs e) { CancelDrag(); base.OnParentChanged(e); }
        protected override void OnHandleDestroyed(EventArgs e) { CancelDrag(); base.OnHandleDestroyed(e); }
        protected override void OnPaint(PaintEventArgs e)
        {
            // 与 RowGrip 一致，先清画布再画装饰：本控件声明了 Opaque 且 OnPaintBackground 为空，
            // 表面完全自绘——「打开即最大化」会让拖柄随布局骤变整块重绘，
            // 不清底就会露出未初始化的 GDI 表面（实测整条变黑）。
            e.Graphics.Clear(Ui.Canvas);
            using (var brush = new SolidBrush(Color.FromArgb(150, SystemColors.ControlDark)))
                for (int i = 1; i <= 3; i++) e.Graphics.FillRectangle(brush, Width * i / 4, Height / 4, 1, Height / 2);
        }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // 双保险：OnPaint 已清一次，这里再清一次，避免任何自绘时序差异露出未初始化表面。
            using (var brush = new SolidBrush(Ui.Canvas)) e.Graphics.FillRectangle(brush, ClientRectangle);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && target != null)
            {
                CancelDrag();
                target.SizeChanged -= OnTargetBoundsChanged;
                target.LocationChanged -= OnTargetBoundsChanged;
                target.ParentChanged -= OnTargetParentChanged;
                target.VisibleChanged -= OnTargetVisibilityChanged;
                target.Disposed -= OnTargetDisposed;
            }
            base.Dispose(disposing);
        }
    }

    // A horizontal drag strip placed under a section. Dragging up/down resizes the two
    // TableLayoutPanel rows it spans (the box above and the next content box below), so a
    // long list can be given more room. Heights are stored under synthetic keys that no
    // control owns, so the generic DialogControlSizing.Apply never re-lays them out; the
    // dialog itself reads them back on load.
    // Cancellation restores RowStyle types and values, not just measured pixel heights.
    internal sealed class RowGrip : Control
    {
        private const int GripHeight = 8;
        private readonly TableLayoutPanel host;
        private readonly int above, below, minAbove, minBelow; // Minimums are 96-DPI logical pixels.
        private readonly string saveKey;
        private readonly bool fixedBelow;   // true：只改 above 行高，below 行（如按钮栏 AutoSize）保持不动
        private int anchorY, aboveHeight, belowHeight;
        private SizeType aboveType, belowType;
        private float aboveStyleHeight, belowStyleHeight;
        private bool dragging;
        internal RowGrip(TableLayoutPanel layout, int aboveRow, int belowRow, string key, int minimumAbove, int minimumBelow)
            : this(layout, aboveRow, belowRow, key, minimumAbove, minimumBelow, false) { }
        internal RowGrip(TableLayoutPanel layout, int aboveRow, int belowRow, string key, int minimumAbove, int minimumBelow, bool fixedBelowRow)
        {
            host = layout; above = aboveRow; below = belowRow; saveKey = key;
            minAbove = minimumAbove; minBelow = minimumBelow; fixedBelow = fixedBelowRow;
            Height = GripHeight; Margin = new Padding(0); Dock = DockStyle.Fill; Cursor = Cursors.SizeNS;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.Opaque | ControlStyles.ResizeRedraw, true);
            host.Disposed += OnHostDisposed;
        }
        private void OnHostDisposed(object sender, EventArgs e) { CancelDrag(); }
        private void Fix(int row, float height)
        { host.RowStyles[row].SizeType = SizeType.Absolute; host.RowStyles[row].Height = height; }
        private bool Begin(Point location)
        {
            if (host.IsDisposed || above < 0 || below < 0 || above >= host.RowStyles.Count || below >= host.RowStyles.Count) return false;
            int[] heights = host.GetRowHeights();
            if (above >= heights.Length || below >= heights.Length) return false;
            anchorY = PointToScreen(location).Y; aboveHeight = heights[above]; belowHeight = heights[below];
            // A paired resize redistributes its original height; impossible
            // minimums must not expand the surrounding layout on mouse-down.
            if (!fixedBelow && (long)aboveHeight + belowHeight < (long)Ui.Px(this, minAbove) + Ui.Px(this, minBelow)) return false;
            aboveType = host.RowStyles[above].SizeType; aboveStyleHeight = host.RowStyles[above].Height;
            belowType = host.RowStyles[below].SizeType; belowStyleHeight = host.RowStyles[below].Height;
            Fix(above, aboveHeight);
            if (!fixedBelow) Fix(below, belowHeight);
            return true;
        }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && !IsDisposed)
            {
                CancelDrag();
                if (Begin(e.Location)) { dragging = true; Capture = true; if (!Capture) CancelDrag(); }
            }
            base.OnMouseDown(e);
        }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            if (dragging)
            {
                if (!Capture || (e.Button & MouseButtons.Left) == 0 || host.IsDisposed) { CancelDrag(); base.OnMouseMove(e); return; }
                int delta = PointToScreen(e.Location).Y - anchorY;
                float total = (float)aboveHeight + belowHeight;
                int minimumAbove = Ui.Px(this, minAbove), minimumBelow = Ui.Px(this, minBelow);
                if (!fixedBelow && total < (long)minimumAbove + minimumBelow) { CancelDrag(); base.OnMouseMove(e); return; }
                float next = Math.Max(minimumAbove, (float)aboveHeight + delta);
                if (!fixedBelow) next = Math.Min(next, total - minimumBelow);
                host.RowStyles[above].Height = next;
                if (!fixedBelow)
                {
                    host.RowStyles[below].Height = total - next;
                }
                host.PerformLayout();
            }
            base.OnMouseMove(e);
        }
        protected override void OnMouseUp(MouseEventArgs e)
        {
            if (dragging)
            {
                if (e.Button == MouseButtons.Left && Capture && !host.IsDisposed)
                { dragging = false; Capture = false; Commit(); }
                else CancelDrag();
            }
            base.OnMouseUp(e);
        }
        private void Commit()
        {
            Form form = FindForm();
            if (form == null || form.IsDisposed || string.IsNullOrEmpty(saveKey) || Ui.SaveRowHeight == null) return;
            float scale;
            using (Graphics g = form.CreateGraphics()) scale = g.DpiX / 96f;
            Ui.SaveRowHeight(form, saveKey, (int)Math.Round(host.RowStyles[above].Height / scale));
        }
        private void CancelDrag()
        {
            if (!dragging) return;
            dragging = false;
            if (!host.IsDisposed)
            {
                host.SuspendLayout();
                try
                {
                    if (above < host.RowStyles.Count) { host.RowStyles[above].SizeType = aboveType; host.RowStyles[above].Height = aboveStyleHeight; }
                    if (!fixedBelow && below < host.RowStyles.Count) { host.RowStyles[below].SizeType = belowType; host.RowStyles[below].Height = belowStyleHeight; }
                }
                finally { host.ResumeLayout(true); }
            }
            Capture = false;
        }
        protected override void OnMouseCaptureChanged(EventArgs e) { if (dragging && !Capture) CancelDrag(); base.OnMouseCaptureChanged(e); }
        protected override void OnVisibleChanged(EventArgs e) { if (!Visible) CancelDrag(); base.OnVisibleChanged(e); }
        protected override void OnEnabledChanged(EventArgs e) { if (!Enabled) CancelDrag(); base.OnEnabledChanged(e); }
        protected override void OnParentChanged(EventArgs e) { CancelDrag(); base.OnParentChanged(e); }
        protected override void OnHandleDestroyed(EventArgs e) { CancelDrag(); base.OnHandleDestroyed(e); }
        protected override void Dispose(bool disposing)
        {
            if (disposing) { CancelDrag(); host.Disposed -= OnHostDisposed; }
            base.Dispose(disposing);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(Ui.Canvas);
            using (Pen pen = new Pen(Ui.GridFrame)) e.Graphics.DrawLine(pen, 0, Height / 2 - 1, Width, Height / 2 - 1);
            using (SolidBrush brush = new SolidBrush(Color.FromArgb(150, SystemColors.ControlDark)))
                for (int i = 0; i < 3; i++) e.Graphics.FillRectangle(brush, Width / 2 - 6 + i * 5, Height / 2 - 1, 2, 2);
        }
        protected override void OnPaintBackground(PaintEventArgs e) { }
    }

    // Native caption buttons, resizing, system menu and snap remain owned by Windows.
    // This binding follows the form's handle and owns only its generated icon.
    internal sealed class UiWindowChrome : NativeWindow
    {
        private readonly Form _owner;
        private readonly Icon _icon;

        private UiWindowChrome(Form owner, string iconKind, Color accent)
        {
            _owner = owner;
            _icon = UiIcons.CreateWindowIcon(iconKind, accent);
            owner.Icon = _icon;
            owner.ShowIcon = true;
            owner.HandleCreated += delegate { BindHandle(); };
            owner.HandleDestroyed += delegate { if (Handle != IntPtr.Zero) ReleaseHandle(); };
            owner.Activated += delegate { ApplyPalette(); };
            owner.Deactivate += delegate { ApplyPalette(); };
            owner.Shown += delegate { ApplyPalette(); };
            owner.Disposed += delegate {
                if (Handle != IntPtr.Zero) ReleaseHandle();
                if (owner.Icon == _icon) owner.Icon = null;
                _icon.Dispose();
            };
            if (owner.IsHandleCreated) BindHandle();
        }

        internal static void Attach(Form owner, string iconKind, Color accent)
        {
            if(Ui.IsInertPreview(owner))return;
            new UiWindowChrome(owner, iconKind, accent);
        }

        private void BindHandle()
        {
            if (Handle != IntPtr.Zero) ReleaseHandle();
            AssignHandle(_owner.Handle);
            ApplyPalette();
        }

        private void ApplyPalette()
        {
            if (Handle != IntPtr.Zero) TryApplyPalette(Handle, SystemInformation.HighContrast);
        }

        protected override void WndProc(ref Message message)
        {
            base.WndProc(ref message);
            // WM_SETTINGCHANGE, WM_THEMECHANGED, WM_DWMCOMPOSITIONCHANGED, WM_DPICHANGED.
            if (message.Msg == 0x001A || message.Msg == 0x031A || message.Msg == 0x031E || message.Msg == 0x02E0)
                ApplyPalette();
        }

        internal static bool TryApplyPalette(IntPtr handle, bool highContrast)
        {
            if (handle == IntPtr.Zero) return false;
            // COLORREF is RGB, not the ARGB value returned by Color.ToArgb().
            int border = highContrast ? -1 : ColorRef(Ui.GridFrame);
            int caption = highContrast ? -1 : ColorRef(Ui.Canvas);
            int text = highContrast ? -1 : ColorRef(Ui.Ink);
            if (!highContrast) SetAttribute(handle, 20, 0); // Keep the light UI's native caption light.
            SetAttribute(handle, 33, highContrast ? 0 : 2); // Default / rounded native corners.
            bool borderSet = SetAttribute(handle, 34, border);
            bool captionSet = SetAttribute(handle, 35, caption);
            bool textSet = SetAttribute(handle, 36, text);
            return borderSet && captionSet && textSet;
        }

        private static int ColorRef(Color color) { return color.R | (color.G << 8) | (color.B << 16); }

        private static bool SetAttribute(IntPtr handle, int attribute, int value)
        {
            try { return DwmSetWindowAttribute(handle, attribute, ref value, sizeof(int)) == 0; }
            catch (DllNotFoundException) { return false; }
            catch (EntryPointNotFoundException) { return false; }
        }

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr handle, int attribute, ref int value, int size);
    }

    internal sealed class BufferedTableLayoutPanel : TableLayoutPanel
    {
        internal BufferedTableLayoutPanel() { DoubleBuffered = true; ResizeRedraw = true; }
    }

    internal sealed class BufferedFlowLayoutPanel : FlowLayoutPanel
    {
        internal BufferedFlowLayoutPanel() { DoubleBuffered = true; ResizeRedraw = true; }
    }

    internal class UiCard : Panel
    {
        public UiCard()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            BackColor = Color.White;
            Padding = new Padding(12);
            Margin = new Padding(0, 0, 0, 12);
        }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent != null ? Parent.BackColor : Ui.Canvas);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = Ui.RoundRect(new RectangleF(.5f, .5f, Width - 1f, Height - 1f), 12f * e.Graphics.DpiX / 96f))
            using (SolidBrush brush = new SolidBrush(BackColor))
            using (Pen pen = new Pen(Ui.GridFrame))
            { e.Graphics.FillPath(brush, path); e.Graphics.DrawPath(pen, path); }
        }
    }

    internal class UiProgressBar : ProgressBar
    {
        public UiProgressBar()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            ForeColor = Ui.Green;
        }
        protected override void WndProc(ref Message message)
        {
            base.WndProc(ref message);
            if (message.Msg == 0x0402) Invalidate(); // PBM_SETPOS
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Ui.Canvas);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float height = Math.Min(8f * g.DpiX / 96f, Height);
            RectangleF track = new RectangleF(0, (Height - height) / 2f, Width, height);
            using (GraphicsPath path = Ui.RoundRect(track, height / 2))
            using (SolidBrush brush = new SolidBrush(Ui.GridFrame)) g.FillPath(brush, path);
            float fraction = Maximum > Minimum ? (float)(Value - Minimum) / (Maximum - Minimum) : 0;
            if (fraction > 0)
            {
                RectangleF fill = new RectangleF(track.X, track.Y, Math.Max(height, track.Width * fraction), height);
                using (GraphicsPath path = Ui.RoundRect(fill, height / 2))
                using (SolidBrush brush = new SolidBrush(ForeColor)) g.FillPath(brush, path);
            }
        }
    }

    internal class RoundedButton : Button
    {
        private bool _hover, _pressed;
        private int _radius = 8;
        public Color BaseColor { get; set; }
        public Color HoverColor { get; set; }
        public Color DownColor { get; set; }
        public Color BorderColor { get; set; }
        public Color DisabledBackColor { get; set; }
        public Color DisabledForeColor { get; set; }
        public bool ShowFocusRect { get; set; }
        public int Radius { get { return _radius; } set { _radius = Math.Max(0, value); Invalidate(); } }

        public RoundedButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
            SetStyle(ControlStyles.SupportsTransparentBackColor, false);
            BackColor = Ui.Canvas;
            BaseColor = Ui.Blue; HoverColor = Ui.BlueHover; DownColor = Ui.BlueDown;
            DisabledBackColor = Color.FromArgb(235, 238, 244);
            DisabledForeColor = Color.FromArgb(128, 139, 158);
            ForeColor = Color.White;
            Cursor = Cursors.Hand;
            MinimumSize = new Size(68, 34);
            Size = new Size(100, 38);
            Margin = new Padding(0, 0, 8, 8);
            ShowFocusRect = true;
        }
        protected override void OnMouseEnter(EventArgs e) { _hover = true; base.OnMouseEnter(e); Invalidate(); }
        protected override void OnMouseLeave(EventArgs e) { _hover = _pressed = false; base.OnMouseLeave(e); Invalidate(); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) _pressed = true; base.OnMouseDown(e); Invalidate(); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; base.OnMouseUp(e); Invalidate(); }
        protected override void OnKeyDown(KeyEventArgs e) { if (e.KeyCode == Keys.Space) _pressed = true; base.OnKeyDown(e); Invalidate(); }
        protected override void OnKeyUp(KeyEventArgs e) { _pressed = false; base.OnKeyUp(e); Invalidate(); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { _pressed = false; base.OnLostFocus(e); Invalidate(); }
        protected override void OnEnabledChanged(EventArgs e) { _pressed = false; base.OnEnabledChanged(e); Invalidate(); }
        protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); Invalidate(); }
        protected override void OnParentChanged(EventArgs e) { base.OnParentChanged(e); Invalidate(); }
        protected override void OnParentBackColorChanged(EventArgs e) { base.OnParentBackColorChanged(e); Invalidate(); }

        private void PaintOpaqueBackground(Graphics graphics)
        {
            Color background = Ui.Canvas;
            for (Control container = Parent; container != null; container = container.Parent)
            {
                if (container.BackColor.A == 255)
                {
                    background = container.BackColor;
                    break;
                }
            }
            using (SolidBrush brush = new SolidBrush(background))
                graphics.FillRectangle(brush, ClientRectangle);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            // ButtonBase's transparent/double-buffered background can leave black corners.
            PaintOpaqueBackground(e.Graphics);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            // Do not call base.OnPaint: ButtonBase would draw a rectangular 3D surface.
            Graphics g = e.Graphics;
            // Also initialize the buffer when WM_PRINTCLIENT or a partial repaint skips
            // the background pass. Antialiased edges must always blend into opaque color.
            PaintOpaqueBackground(g);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = new Rectangle(1, 1, Width - 3, Height - 3);
            using (GraphicsPath path = Ui.RoundRect(rect, _radius * g.DpiX / 96f))
            using (SolidBrush brush = new SolidBrush(!Enabled ? DisabledBackColor : _pressed ? DownColor : _hover ? HoverColor : BaseColor))
            {
                g.FillPath(brush, path);
                if (BorderColor != Color.Empty) using (Pen pen = new Pen(BorderColor)) g.DrawPath(pen, path);
                if (ShowFocusRect && Focused && ShowFocusCues)
                    using (Pen pen = new Pen(BaseColor.GetBrightness() > .8f ? Ui.Blue : ForeColor, 2))
                    using (GraphicsPath focus = Ui.RoundRect(new Rectangle(4, 4, Width - 9, Height - 9), Math.Max(2, _radius - 3)))
                        g.DrawPath(pen, focus);
            }
            TextRenderer.DrawText(g, Text, Font, rect, Enabled ? ForeColor : DisabledForeColor,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        }
    }

    internal static class UiDialog
    {
        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
        { return Show(null, text, caption, buttons, icon, defaultButton); }

        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon,
            MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
        {
            if (Ui.HostDialogPresenter != null) return Ui.HostDialogPresenter(owner,text,caption,buttons,icon,defaultButton);
            return ShowDirect(owner,text,caption,buttons,icon,defaultButton);
        }
        internal static DialogResult ShowDirect(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
        {
            if (buttons != MessageBoxButtons.OK && buttons != MessageBoxButtons.YesNo && buttons != MessageBoxButtons.OKCancel)
                return MessageBox.Show(owner, text, caption, buttons, icon, defaultButton);
            try
            {
                using (UiMessageForm form = new UiMessageForm(text, caption, buttons, icon, defaultButton))
                    return owner == null ? form.ShowDialog() : form.ShowDialog(owner);
            }
            catch { return MessageBox.Show(owner, text, caption, buttons, icon, defaultButton); }
        }
    }

    internal class UiMessageForm : Form
    {
        private readonly bool _confirm;
        private readonly bool _yesNo;
        internal readonly RoundedButton Positive;
        internal readonly RoundedButton Negative;
        internal DialogResult DismissResult { get { return _yesNo ? DialogResult.No : _confirm ? DialogResult.Cancel : DialogResult.OK; } }

        internal UiMessageForm(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton)
        {
            _yesNo = buttons == MessageBoxButtons.YesNo;
            _confirm = _yesNo || buttons == MessageBoxButtons.OKCancel;
            Text = caption;
            bool warning = icon == MessageBoxIcon.Error || icon == MessageBoxIcon.Warning;
            Ui.PrepareForm(this, new Size(560, 360), new Size(440, 280),
                warning ? "alert" : icon == MessageBoxIcon.Question ? "question" : "info", warning ? Ui.Danger : Ui.Blue,
                _yesNo ? "message.yesno" : _confirm ? "message.okcancel" : "message.ok");
            ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = false;
            KeyPreview = true;
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill; root.Padding = new Padding(20); root.ColumnCount = 1; root.RowCount = 3;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Label title = Ui.Label(caption, 16f, true);
            title.ForeColor = icon == MessageBoxIcon.Error || icon == MessageBoxIcon.Warning ? Ui.Danger : Ui.Blue;
            TableLayoutPanel titleRow = new TableLayoutPanel(); titleRow.Dock = DockStyle.Fill;
            titleRow.ColumnCount = 2; titleRow.RowCount = 1; titleRow.Margin = Padding.Empty;
            titleRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); titleRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72));
            Label severity = Ui.Label(icon == MessageBoxIcon.Error ? "错误" : icon == MessageBoxIcon.Warning ? "注意" : icon == MessageBoxIcon.Question ? "请确认" : "提示", 9f, false);
            severity.ForeColor = title.ForeColor; severity.TextAlign = ContentAlignment.MiddleCenter;
            severity.BackColor = Ui.SelectBack; severity.Margin = new Padding(0, 8, 0, 8);
            titleRow.Controls.Add(title, 0, 0); titleRow.Controls.Add(severity, 1, 0);
            root.Controls.Add(titleRow, 0, 0);
            TextBox body = new TextBox();
            body.Multiline = true; body.ReadOnly = true; body.BorderStyle = BorderStyle.None;
            body.ScrollBars = ScrollBars.Vertical; body.Text = text; body.BackColor = Color.White;
            body.ForeColor = Ui.Ink; body.Font = Ui.MakeFont(10f, FontStyle.Regular, null); body.TabStop = true;
            root.Controls.Add(Ui.Card(body, null), 0, 1);
            FlowLayoutPanel bar = Ui.ActionBar();
            RoundedButton copy = Ui.Button("复制详情", 104);
            copy.Click += delegate { try { Clipboard.SetDataObject(text ?? "", true); copy.Text = "已复制"; } catch { copy.Text = "复制失败"; } };
            bar.Controls.Add(copy);
            Positive = Ui.Button(_confirm ? "确认" : "确定", 100);
            Ui.StylePrimary(Positive);
            if (_confirm && (caption.IndexOf("中断", StringComparison.Ordinal) >= 0 || caption == "卸载")) Ui.StyleDanger(Positive);
            Positive.DialogResult = _yesNo ? DialogResult.Yes : DialogResult.OK;
            bar.Controls.Add(Positive);
            if (_confirm)
            {
                Negative = Ui.Button("取消", 100); Negative.DialogResult = DismissResult;
                bar.Controls.Add(Negative); CancelButton = Negative;
                // Never bind Enter to a potentially destructive confirmation.
                AcceptButton = null;
                Shown += delegate { if(Ui.IsInertPreview(this))return; if (defaultButton == MessageBoxDefaultButton.Button2) Negative.Focus(); else Positive.Focus(); };
            }
            else { AcceptButton = Positive; CancelButton = Positive; }
            KeyDown += delegate(object s, KeyEventArgs e)
            { if (e.KeyCode == Keys.Escape) { e.Handled = true; DialogResult = DismissResult; Close(); } };
            FormClosing += delegate { if (DialogResult == DialogResult.None || DialogResult == DialogResult.Cancel) DialogResult = DismissResult; };
            root.Controls.Add(bar, 0, 2); Controls.Add(root);
        }
    }

    internal static class UiIcons
    {
        public static Icon CreateWindowIcon(string kind, Color accent)
        {
            // A managed ICO containing both PNG frames keeps small captions crisp and
            // larger task-switcher icons sharp without a white background or native HICON leaks.
            byte[][] frames = new byte[2][];
            int[] sizes = new int[] { 16, 32 };
            for (int i = 0; i < sizes.Length; i++)
            {
                using (Bitmap bitmap = Create(kind, accent, sizes[i], true))
                using (MemoryStream png = new MemoryStream())
                { bitmap.Save(png, ImageFormat.Png); frames[i] = png.ToArray(); }
            }
            using (MemoryStream data = new MemoryStream())
            using (BinaryWriter writer = new BinaryWriter(data))
            {
                writer.Write((ushort)0); writer.Write((ushort)1); writer.Write((ushort)2);
                int offset = 6 + 16 * frames.Length;
                for (int i = 0; i < frames.Length; i++)
                {
                    writer.Write((byte)sizes[i]); writer.Write((byte)sizes[i]);
                    writer.Write((byte)0); writer.Write((byte)0);
                    writer.Write((ushort)1); writer.Write((ushort)32);
                    writer.Write(frames[i].Length); writer.Write(offset); offset += frames[i].Length;
                }
                foreach (byte[] frame in frames) writer.Write(frame);
                writer.Flush();
                using (MemoryStream input = new MemoryStream(data.ToArray()))
                using (Icon icon = new Icon(input)) return (Icon)icon.Clone();
            }
        }

        public static Bitmap Create(string kind, Color accent, int size)
        {
            return Create(kind, accent, size, false);
        }

        // Ribbon geometry is independent of the original 16/32 caption ICO layout.
        // Draw at the calibrated physical resolution, then let OLE expose 16/32 logical px.
        public static Bitmap CreateRibbon(string kind, Color accent, int width, int height)
        {
            Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White); g.SmoothingMode = SmoothingMode.AntiAlias;
                g.ScaleTransform(width / 32f, height / 32f);
                Color tint = Color.FromArgb(238 + (accent.R - 238) / 12, 242 + (accent.G - 242) / 12, 250 + (accent.B - 250) / 12);
                using (GraphicsPath tile = Ui.RoundRect(new RectangleF(1, 1, 30, 30), 8))
                using (SolidBrush fill = new SolidBrush(tint)) g.FillPath(fill, tile);
                using (Pen pen = new Pen(accent, 2f))
                {
                    pen.StartCap = pen.EndCap = LineCap.Round; pen.LineJoin = LineJoin.Round;
                    if (kind == "update")
                    { g.DrawArc(pen, 4, 4, 24, 24, 35, 280); g.DrawLines(pen, new PointF[] { new PointF(22, 3), new PointF(28, 7), new PointF(24, 13) }); }
                    else if (kind == "filter")
                    { g.DrawLines(pen, new PointF[] { new PointF(4, 5), new PointF(28, 5), new PointF(19, 16), new PointF(19, 25), new PointF(13, 28), new PointF(13, 16), new PointF(4, 5) }); }
                    else if (kind == "link")
                    { g.DrawArc(pen, 3, 3, 17, 14, 65, 275); g.DrawArc(pen, 12, 15, 17, 14, 245, 275); g.DrawLine(pen, 11, 13, 21, 22); }
                    else if (kind == "directory")
                    { using (GraphicsPath page = Ui.RoundRect(new RectangleF(4, 3, 24, 26), 3)) g.DrawPath(pen, page); for (int y=9;y<=23;y+=7) { g.DrawEllipse(pen,9,y,2,2); g.DrawLine(pen,15,y+1,23,y+1); } }
                    else if (kind == "settings")
                    { for (int i=0;i<3;i++) { int x=7+i*9, y=9+i*6; g.DrawLine(pen,x,4,x,28); g.DrawEllipse(pen,x-3,y-3,6,6); } }
                    else if (kind == "web")
                    { g.DrawEllipse(pen, 4, 4, 24, 24); g.DrawEllipse(pen, 10, 4, 12, 24); g.DrawLine(pen, 4, 16, 28, 16); g.DrawArc(pen, 5, 6, 22, 12, 5, 170); g.DrawArc(pen, 5, 14, 22, 12, 185, 170); }
                    else if (kind == "agent")
                    { using (GraphicsPath robot = Ui.RoundRect(new RectangleF(4, 8, 24, 20), 6)) g.DrawPath(pen, robot); g.DrawLine(pen, 16, 4, 16, 8); g.DrawEllipse(pen, 14, 2, 4, 4); g.DrawEllipse(pen, 9, 14, 3, 3); g.DrawEllipse(pen, 20, 14, 3, 3); g.DrawLine(pen, 11, 23, 21, 23); }
                    else if (kind == "document")
                    { g.DrawLines(pen, new PointF[] { new PointF(20, 3), new PointF(5, 3), new PointF(5, 29), new PointF(27, 29), new PointF(27, 10), new PointF(20, 3), new PointF(20, 10), new PointF(27, 10) }); g.DrawLine(pen, 10, 16, 22, 16); g.DrawLine(pen, 10, 22, 20, 22); }
                    else if (kind == "formula")
                    {
                        // Code-native rounded fx, independent of installed symbol fonts.
                        using (GraphicsPath f = new GraphicsPath())
                        { f.AddBezier(4, 28, 10, 31, 9, 6, 14, 4); f.AddBezier(14, 4, 16, 3, 17, 4, 18, 5); g.DrawPath(pen, f); }
                        g.DrawLine(pen, 5, 13, 17, 13);
                        g.DrawLine(pen, 18, 13, 28, 26); g.DrawLine(pen, 28, 13, 18, 26);
                    }
                    else if (kind == "mail")
                    { using (GraphicsPath box = Ui.RoundRect(new RectangleF(3, 5, 26, 22), 4)) g.DrawPath(pen, box); g.DrawLines(pen, new PointF[] { new PointF(4, 7), new PointF(16, 17), new PointF(28, 7) }); }
                    else if (kind == "paste")
                    { using (GraphicsPath box = Ui.RoundRect(new RectangleF(3, 3, 19, 19), 2)) g.DrawPath(pen, box); g.DrawLine(pen, 3, 12, 22, 12); g.DrawLine(pen, 12, 3, 12, 22); g.DrawEllipse(pen, 16, 22, 11, 6); g.DrawEllipse(pen, 22, 16, 6, 11); }
                    else if (kind == "table")
                    { using (GraphicsPath box = Ui.RoundRect(new RectangleF(4, 4, 24, 24), 2)) g.DrawPath(pen, box); g.DrawLine(pen, 4, 12, 28, 12); g.DrawLine(pen, 12, 4, 12, 28); g.DrawLines(pen, new PointF[] { new PointF(17, 22), new PointF(21, 26), new PointF(27, 18) }); }
                    else if (kind == "precedents" || kind == "dependents")
                    {
                        using (GraphicsPath source = Ui.RoundRect(new RectangleF(3.5f, 3.5f, 10, 10), 1.5f)) g.DrawPath(pen, source);
                        using (GraphicsPath target = Ui.RoundRect(new RectangleF(18.5f, 18.5f, 10, 10), 1.5f)) g.DrawPath(pen, target);
                        g.DrawLine(pen, 13.5f, 13.5f, 18.5f, 18.5f);
                        if (kind == "precedents") g.DrawLines(pen, new PointF[] { new PointF(21, 14.5f), new PointF(14.5f, 14.5f), new PointF(14.5f, 21) });
                        else g.DrawLines(pen, new PointF[] { new PointF(17, 17.5f), new PointF(23.5f, 17.5f), new PointF(23.5f, 11) });
                    }
                    else
                    { using (GraphicsPath page = Ui.RoundRect(new RectangleF(4, 3, 16, 23), 2)) g.DrawPath(pen, page); g.DrawLine(pen, 8, 9, 16, 9); g.DrawLine(pen, 8, 14, 14, 14); g.DrawEllipse(pen, 17, 17, 11, 11); g.DrawLine(pen, 26, 26, 29, 29); }
                }
            }
            return bitmap;
        }

        private static Bitmap Create(string kind, Color accent, int size, bool transparent)
        {
            Bitmap bitmap = new Bitmap(size, size, transparent ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(transparent ? Color.Transparent : Color.White); g.SmoothingMode = SmoothingMode.AntiAlias;
                g.ScaleTransform(size / 32f, size / 32f);
                Color tint = Color.FromArgb(238 + (accent.R - 238) / 12, 242 + (accent.G - 242) / 12, 250 + (accent.B - 250) / 12);
                using (GraphicsPath path = Ui.RoundRect(new RectangleF(1, 1, 30, 30), 8))
                using (SolidBrush brush = new SolidBrush(tint)) g.FillPath(brush, path);
                using (Pen pen = new Pen(accent, 2f))
                {
                    pen.StartCap = pen.EndCap = LineCap.Round; pen.LineJoin = LineJoin.Round;
                    if (kind == "update")
                    { g.DrawArc(pen, 7, 7, 18, 18, 35, 280); g.DrawLines(pen, new PointF[] { new PointF(20, 6), new PointF(25, 9), new PointF(22, 14) }); }
                    else if (kind == "mail")
                    { using (GraphicsPath box = Ui.RoundRect(new RectangleF(6, 9, 20, 15), 3)) g.DrawPath(pen, box); g.DrawLines(pen, new PointF[] { new PointF(7, 10), new PointF(16, 17), new PointF(25, 10) }); }
                    else if (kind == "paste")
                    { g.DrawRectangle(pen, 6, 6, 14, 14); g.DrawLine(pen, 6, 13, 20, 13); g.DrawLine(pen, 13, 6, 13, 20); g.DrawEllipse(pen, 16, 20, 9, 6); g.DrawEllipse(pen, 21, 16, 6, 9); }
                    else if (kind == "table")
                    { g.DrawRectangle(pen, 6, 7, 20, 18); g.DrawLine(pen, 6, 13, 26, 13); g.DrawLine(pen, 13, 7, 13, 25); g.DrawLines(pen, new PointF[] { new PointF(17, 20), new PointF(20, 23), new PointF(26, 17) }); }
                    else if (kind == "link")
                    { g.DrawArc(pen, 5, 7, 14, 11, 65, 275); g.DrawArc(pen, 13, 15, 14, 11, 245, 275); g.DrawLine(pen, 12, 14, 20, 21); }
                    else if (kind == "directory")
                    { g.DrawRectangle(pen, 7, 5, 18, 22); for (int y=10;y<=22;y+=6) { g.DrawLine(pen,11,y,12,y); g.DrawLine(pen,16,y,22,y); } }
                    else if (kind == "document")
                    { g.DrawLines(pen, new PointF[] { new PointF(20, 5), new PointF(7, 5), new PointF(7, 27), new PointF(25, 27), new PointF(25, 10), new PointF(20, 5), new PointF(20, 10), new PointF(25, 10) }); g.DrawLine(pen, 11, 15, 21, 15); g.DrawLine(pen, 11, 20, 19, 20); }
                    else if (kind == "settings")
                    { for (int i=0;i<3;i++) { int x=8+i*8, y=10+i*5; g.DrawLine(pen,x,6,x,26); g.DrawEllipse(pen,x-2,y-2,4,4); } }
                    else if (kind == "alert" || kind == "info" || kind == "question")
                    {
                        g.DrawEllipse(pen, 6, 6, 20, 20);
                        if (kind == "question") { g.DrawArc(pen, 12, 10, 8, 8, 190, 230); g.DrawLine(pen, 16, 17, 16, 19); }
                        else g.DrawLine(pen, 16, kind == "info" ? 15 : 11, 16, kind == "info" ? 22 : 17);
                        using (SolidBrush dot = new SolidBrush(accent)) g.FillEllipse(dot, 15, kind == "info" ? 10 : 21, 2, 2);
                    }
                    else if (kind == "precedents" || kind == "dependents")
                    {
                        g.DrawRectangle(pen, 5, 5, 9, 9); g.DrawRectangle(pen, 19, 19, 8, 8);
                        if (kind == "precedents") g.DrawLines(pen, new PointF[] { new PointF(21, 15), new PointF(15, 15), new PointF(15, 21) });
                        else g.DrawLines(pen, new PointF[] { new PointF(17, 17), new PointF(23, 17), new PointF(23, 11) });
                        g.DrawLine(pen, 14, 14, 19, 19);
                    }
                    else
                    { g.DrawRectangle(pen, 7, 5, 13, 20); g.DrawLine(pen, 10, 10, 17, 10); g.DrawLine(pen, 10, 14, 15, 14); g.DrawEllipse(pen, 16, 17, 10, 10); g.DrawLine(pen, 24, 25, 28, 29); }
                }
            }
            return bitmap;
        }
    }
}
