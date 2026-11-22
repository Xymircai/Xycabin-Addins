using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel_Iterate;

internal static class ConversionInteractionProbe
{
    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr keysAndDelta, IntPtr location);

    private static int checks, hostThread;
    private static volatile CellConversionForm editor;
    private static CellConversionResult result;
    private static long ticket;
    private static void Check(bool value, string message) { checks++; if (!value) throw new Exception(message); }
    private static void Pump(Func<bool> done)
    { var stop = DateTime.UtcNow.AddSeconds(12); while (!done()) { if (DateTime.UtcNow > stop) throw new Exception("Timed out waiting for conversion UI"); Application.DoEvents(); Thread.Sleep(5); } }
    private static void OnUi(Action<CellConversionForm> action)
    { var form = editor; form.BeginInvoke((MethodInvoker)delegate { action(form); }); }
    private static void Settled() { Pump(() => !ConversionInteraction.HasPending && ConversionWindowThread.Count == 0); }
    private static void ApplyAndClose(Host host)
    {
        OnUi(f => f.Apply.PerformClick());
        Pump(() => result != null && editor.Enabled);
        Check(ConversionInteraction.HasPending && ConversionWindowThread.Count == 1 && host.Disposals == 0,
            "apply keeps the conversion window and its target alive");
        OnUi(f => f.Cancel.PerformClick()); Settled();
    }
    private sealed class Cell : IConversionCell
    {
        internal ConversionCellState Current = new ConversionCellState { IsFormula = true, Formula = "=$A$2", Value = 7.0, Display = "7", NumberFormat = "General" };
        public string Address { get { return "A1"; } }
        public ConversionCellState Original { get; private set; }
        internal int Writes;
        internal Cell() { Original = Current.Copy(); }
        internal void Snapshot() { Original = Current.Copy(); }
        public void CheckWritable() { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "preflight on host STA"); }
        public ConversionCellState Read() { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "read on host STA"); return Current.Copy(); }
        public void Write(ConversionCellState state) { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "write on host STA"); Writes++; Current = state.Copy(); }
    }
    private sealed class Host : IConversionSessionHost
    {
        internal Cell Cell = new Cell();
        internal int Disposals, Picks, PreviewReads;
        internal readonly List<string> PreviewCells = new List<string>();
        internal string Picked;
        internal bool PickFailure, ScopeFailure;
        public string Address { get; private set; }
        public string Description { get { return "工作簿：测试.xlsx\r\n工作表：数据"; } }
        public bool EnableEvents { get; set; }
        public bool ScreenUpdating { get; set; }
        internal Host() { Address = "A1"; EnableEvents = ScreenUpdating = true; }
        public IList<IConversionCell> Collect() { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "collect on host STA"); Cell.Snapshot(); return new IConversionCell[] { Cell }; }
        public void Calculate(IList<IConversionCell> targets) { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "calculate on host STA"); }
        public ConversionPreview ReadPreview(string address) { RecheckScope(); PreviewReads++; PreviewCells.Add(ConversionRangeAddress.FirstCell(address)); return new ConversionPreview(address, Cell.Read()); }
        public void RecheckScope() { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "scope on host STA"); if (ScopeFailure) throw new InvalidOperationException("scope changed"); }
        public void Dispose() { Check(Thread.CurrentThread.ManagedThreadId == hostThread, "dispose on host STA"); Disposals++; }
        public void SetAddress(string address) { Address = ConversionRangeAddress.Normalize(address); }
        public string PickRange() { Check(Thread.CurrentThread.ManagedThreadId == hostThread && !editor.Enabled, "picker on host STA while editor disabled"); Picks++; if (PickFailure) throw new InvalidOperationException("picker failure"); return Picked; }
    }
    private static Host Open(CellConversionKind kind = CellConversionKind.ReferenceStyle, Host supplied = null)
    {
        editor = null; result = null; var host = supplied ?? new Host(); ticket = ConversionInteraction.Open(host, kind, IntPtr.Zero);
        Pump(() => editor != null && editor.IsHandleCreated); return host;
    }
    private static void InteractionChecks()
    {
        hostThread = Thread.CurrentThread.ManagedThreadId;
        ConversionWindowThread.HiddenForProbe = true;
        ConversionWindowThread.CreatedForProbe = delegate(CellConversionForm form) { Check(Thread.CurrentThread.ManagedThreadId != hostThread && Thread.CurrentThread.GetApartmentState() == ApartmentState.STA, "editor uses independent STA"); editor = form; };
        ConversionInteraction.ResultForProbe = delegate(CellConversionResult value) { result = value; };
        var host = Open(); OnUi(f => f.Cancel.PerformClick()); Settled(); Check(host.Cell.Writes == 0 && host.Disposals == 1, "cancel releases session with zero writes");
        host = Open(); var same = new Host(); long existing = ConversionInteraction.Open(same, CellConversionKind.ReferenceStyle, IntPtr.Zero);
        Check(existing == ticket && same.Disposals == 1, "repeated command keeps a single parameter session"); OnUi(f => f.Close()); Settled();
        host = Open();
        var switched = Open(CellConversionKind.TextFilter);
        Check(host.Disposals == 1 && switched.Disposals == 0 && editor.Text == "文本提取与过滤", "a distinct conversion command replaces the old editor");
        OnUi(f => f.Close()); Settled();
        host = Open(); OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 1 && editor.Enabled);
        Check(host.Address == "A1" && host.Cell.Writes == 0 && ConversionInteraction.HasPending, "cancelled native picker restores pending window and original address");
        host.PickFailure = true; OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 2 && editor.Enabled && editor.Validation.Text.Contains("picker failure"));
        Check(host.Cell.Writes == 0, "picker exception restores editor without writes");
        host.PickFailure = false; host.Picked = "B1"; OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 3 && editor.Enabled && editor.RangeInput.Text == "B1" && editor.Apply.Enabled);
        ApplyAndClose(host); Check(result != null && !result.Failed && host.Address == "B1" && host.Cell.Current.Formula == "=A2" && host.Disposals == 1, "accepted picker fills address then explicit apply commits");
        long completed = ticket; int writes = host.Cell.Writes; ConversionInteraction.Receive(completed, false, new ConversionRequest { Kind = CellConversionKind.ValuesOnly, Address = "A1" }); Application.DoEvents(); Check(host.Cell.Writes == writes, "duplicate ticket ignored");
        host = Open(); ConversionInteraction.Receive(completed, false, null); Application.DoEvents(); Check(ConversionInteraction.HasPending, "old ticket cannot cancel a new session");
        Check(Connect.TryBeginConversion(), "fixture holds unrelated busy operation"); OnUi(f => f.Pick.PerformClick()); Pump(() => editor.Enabled && editor.Validation.Text.Contains("其他操作"));
        Check(host.Picks == 0 && !Connect.TryBeginConversion(), "busy rejection retains another operation guard"); Connect.EndConversion(); OnUi(f => f.Cancel.PerformClick()); Settled();
        host = Open(); host.ScopeFailure = true; ApplyAndClose(host); Check(result.Failed && host.Cell.Writes == 0 && host.Disposals == 1, "changed target is zero-write failure");
        host = Open(CellConversionKind.TextFilter); bool inspected = false;
        OnUi(delegate(CellConversionForm f)
        {
            Check(f.Request().Extract && f.Request().Kinds == CellCharacterKinds.Chinese, "default extract Chinese");
            f.RangeInput.Text = "Other!A1"; Check(!f.Apply.Enabled, "invalid address disables apply"); f.RangeInput.Text = "B1";
            for (int i = 0; i < 5; i++) ((CheckBox)f.Controls.Find("conversionKind" + i, true)[0]).Checked = false;
            Check(!f.Apply.Enabled, "no character categories disables apply"); ((CheckBox)f.Controls.Find("conversionKind3", true)[0]).Checked = true;
            Check(!f.Apply.Enabled, "manual address waits for fresh preview"); inspected = true;
        });
        Pump(() => inspected && editor.Apply.Enabled); ApplyAndClose(host); Check(!result.Failed && (string)host.Cell.Current.Value == "7" && host.Cell.Current.NumberFormat == "@", "manual address and text settings delivered");
        PreviewChecks();
        BatchUiChecks();
        for (int i = 0; i < 3; i++)
        {
            host = Open(); long stale = ticket; ConversionInteraction.Shutdown(); Settled();
            ConversionInteraction.Receive(stale, false, new ConversionRequest { Kind = CellConversionKind.ValuesOnly, Address = "A1" }); Application.DoEvents();
            Check(host.Cell.Writes == 0 && host.Disposals == 1, "shutdown invalidates pending ticket " + i);
        }
        ConversionWindowThread.CreatedForProbe = delegate { throw new InvalidOperationException("startup fixture"); };
        var failed = new Host(); ConversionInteraction.Open(failed, CellConversionKind.TextFilter, IntPtr.Zero); Settled();
        Check(failed.Disposals == 1 && failed.Cell.Writes == 0, "failed UI startup releases target"); ConversionInteraction.Shutdown();
    }
    private static string Before(CellConversionForm f) { return f.Controls.Find("conversionBefore", true)[0].Text; }
    private static string After(CellConversionForm f) { return f.Controls.Find("conversionPreview", true)[0].Text; }
    private static void BatchUiChecks()
    {
        var host = Open(CellConversionKind.BatchInsertTrimDelete); bool done = false;
        OnUi(delegate(CellConversionForm f)
        {
            var tabs = (TabControl)f.Controls.Find("batchEditTabs", true)[0];
            Check(tabs.TabPages.Count == 3 && tabs.TabPages[0].Text == "插入" && tabs.TabPages[1].Text == "截取" && tabs.TabPages[2].Text == "删除", "three batch edit pages");
            Check(!f.Apply.Enabled && Before(f).Contains("=$A$2"), "empty insertion disables execute but shows actual formula source");
            Check(!Descendants(f).Any(c => c is Label && c.Text.StartsWith("可输入 A1:B10", StringComparison.Ordinal)), "batch dialog omits the range explanation requested by user");
            var panel = f.Controls.Find("batchEditPanel", true)[0]; int insertionHeight = panel.Height;
            ((TextBox)f.Controls.Find("batchInsertText", true)[0]).Text = "X";
            Check(f.Apply.Enabled && After(f).Contains("X=$A$2"), "insertion preview edits formula source");
            ((NumericUpDown)f.Controls.Find("batch0Count", true)[0]).Value = 2;
            tabs.SelectedIndex = 1; ((RadioButton)f.Controls.Find("batch1Middle", true)[0]).Checked = true;
            Check(panel.Height < insertionHeight, "trim page compacts the parameter area without insertion whitespace");
            ((NumericUpDown)f.Controls.Find("batch1Last", true)[0]).Value = 5; ((NumericUpDown)f.Controls.Find("batch1First", true)[0]).Value = 3;
            Check(f.Apply.Enabled && After(f).Contains("A$2"), "inclusive formula preview counts leading equal sign");
            ((NumericUpDown)f.Controls.Find("batch1First", true)[0]).Value = 6; Check(!f.Apply.Enabled, "reversed middle disables execute");
            ((NumericUpDown)f.Controls.Find("batch1First", true)[0]).Value = 3;
            tabs.SelectedIndex = 2; Check(f.Request().TextEdit.Count == 1, "delete has independent default count");
            tabs.SelectedIndex = 0; Check(f.Request().TextEdit.Text == "X" && f.Request().TextEdit.Count == 2, "page switching retains insertion parameters");
            tabs.SelectedIndex = 1; Check(f.Request().TextEdit.Start == 3 && f.Request().TextEdit.End == 5, "page switching retains trim parameters"); done = true;
        });
        Pump(() => done); OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 1 && editor.Enabled);
        Check(host.Cell.Writes == 0 && editor.RangeInput.Text == "A1", "batch picker cancel retains range without writes");
        ApplyAndClose(host);
        Check(!result.Failed && (string)host.Cell.Current.Value == "A$2" && !host.Cell.Current.IsFormula && result.PlainResults == 1 && result.FormulasAsText == 0 && host.Cell.Current.NumberFormat == "General", "batch controls submit ordinary result without a text-format fallback");
        host = Open(CellConversionKind.BatchInsertTrimDelete); done = false;
        OnUi(delegate(CellConversionForm f)
        {
            ((TextBox)f.Controls.Find("batchInsertText", true)[0]).Text = "X";
            f.ReceivePreview(0, "A1", new ConversionPreview("A1", new ConversionCellState { Value = null, Display = "", NumberFormat = "General" }), null);
            Check(After(f).Contains("（空白）") && f.Validation.Text.Contains("跳过") && f.Apply.Enabled, "true blank preview stays blank while batch remains usable"); done = true;
        });
        Pump(() => done); OnUi(f => f.Cancel.PerformClick()); Settled(); Check(host.Cell.Writes == 0, "batch cancel is zero write");
    }
    private static void PreviewChecks()
    {
        var host = new Host(); host.Cell.Current = new ConversionCellState { Value = "MIRCAI文化旅游公司2025!!", Display = "MIRCAI文化旅游公司2025!!", NumberFormat = "General" };
        host = Open(CellConversionKind.TextFilter, host); bool done = false;
        OnUi(delegate(CellConversionForm f)
        {
            Check(Before(f) == "修改前示例：MIRCAI文化旅游公司2025!!", "text before example uses host data");
            ((CheckBox)f.Controls.Find("conversionKind1", true)[0]).Checked = true;
            Check(After(f) == "修改后示例：MIRCAI文化旅游公司", "text preview combines actual English and Chinese");
            f.RangeInput.Text = "J1:J5"; Check(!f.Apply.Enabled && !Before(f).Contains("MIRCAI"), "address edit clears stale source immediately"); done = true;
        });
        Pump(() => done && editor.Apply.Enabled && host.PreviewCells.Contains("J1"));
        Check(host.Cell.Writes == 0 && host.PreviewReads == 2, "manual preview is debounced and zero-write");
        host.Picked = "J1:J5"; host.Cell.Current.Display = "新示例2026"; host.Cell.Current.Value = "新示例2026";
        OnUi(f => f.Pick.PerformClick()); Pump(() => editor.Enabled && editor.Apply.Enabled && Before(editor).Contains("新示例"));
        Check(host.PreviewReads == 3 && host.Cell.Writes == 0, "same-address picker refreshes first cell without writes");
        done = false; OnUi(delegate(CellConversionForm f)
        {
            f.ReceivePreview(-1, "J1:J5", new ConversionPreview("J1:J5", new ConversionCellState { Display = "stale" }), null);
            Check(!Before(f).Contains("stale"), "late preview revision cannot replace current source");
            f.RangeInput.Text = "invalid!A1"; done = true;
        });
        Pump(() => done && !editor.Apply.Enabled && editor.Validation.Text.Contains("有效"));
        Check(!Before(editor).Contains("新示例"), "invalid address clears old example");
        OnUi(f => f.Cancel.PerformClick()); Settled();
        host = new Host(); host.Cell.Current.Formula = "=H23+H24+H26"; host = Open(CellConversionKind.ReferenceStyle, host); done = false;
        OnUi(delegate(CellConversionForm f) { f.RangeInput.Text = "G11:G14,G16:G18"; ((RadioButton)f.Controls.Find("conversionReference2", true)[0]).Checked = true; done = true; });
        Pump(() => done && editor.Apply.Enabled && host.PreviewCells.Contains("G11"));
        Check(Before(editor) == "修改前示例：=H23+H24+H26" && After(editor) == "修改后示例：=$H23+$H24+$H26", "reference before/after uses actual first area formula");
        Check(host.Cell.Writes == 0, "reference rule changes never write");
        host.Picked = "G11:G14,G16:G18"; host.Cell.Current = new ConversionCellState { Value = null, Display = "", NumberFormat = "General" };
        OnUi(f => f.Pick.PerformClick()); Pump(() => editor.Enabled && editor.Apply.Enabled && Before(editor).Contains("（空白）"));
        Check(After(editor).Contains("（空白）") && editor.Validation.Text.Contains("首格不是公式"), "blank first cell is not replaced by a later formula");
        host.ScopeFailure = true; OnUi(f => f.RangeInput.Text = "B1"); Pump(() => !editor.Apply.Enabled && editor.Validation.Text.Contains("scope changed"));
        OnUi(f => f.Cancel.PerformClick()); Settled(); Check(host.Cell.Writes == 0, "invalidated preview target stays zero-write");
    }
    private static void Render(string folder)
    {
        Directory.CreateDirectory(folder);
        foreach (string id in new[] { "conversion.text", "formula.references", "conversion.batchInsertTrimDelete" })
        using (var form = DialogPreviewFactory.CreateProgramDefault(id))
        using (var parking = new Panel { Size = form.Size })
        {
            Size client = form.ClientSize; form.FormBorderStyle = FormBorderStyle.None; form.ClientSize = client;
            form.TopLevel = false; parking.Controls.Add(form); form.Show();
            ((CellConversionForm)form).SuppressActivation = true;
            typeof(Form).GetMethod("OnShown", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { EventArgs.Empty }); form.PerformLayout();
            if (id == "conversion.text") ((CheckBox)form.Controls.Find("conversionKind1", true)[0]).Checked = true;
            else if (id == "formula.references") ((RadioButton)form.Controls.Find("conversionReference2", true)[0]).Checked = true;
            using (var overlay = new PreviewSelectionOverlay(form))
            {
                overlay.RefreshSnapshot();
                using (var bitmap = new Bitmap(overlay.Width, overlay.Height)) { overlay.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(folder, id + ".png"), ImageFormat.Png); }
            }
            var range = form.Controls.Find("conversionRange", true)[0];
            Check(range.Width >= 200 && range.Height >= 18 && range.Text == (id == "conversion.text" ? "J1:J5" : id == "formula.references" ? "G11:G14,G16:G18" : "A1:A5"), "usable range editor bounds for " + id);
            Check(form.Controls.Find("conversionApply", true).Length == 1, "real preview controls for " + id);
            if (id == "conversion.batchInsertTrimDelete")
            {
                var tabs = (TabControl)form.Controls.Find("batchEditTabs", true)[0];
                for (int page = 1; page <= 2; page++)
                {
                    tabs.SelectedIndex = page; ((RadioButton)form.Controls.Find("batch" + page + "Middle", true)[0]).Checked = true;
                    ((NumericUpDown)form.Controls.Find("batch" + page + "Last", true)[0]).Value = 6; ((NumericUpDown)form.Controls.Find("batch" + page + "First", true)[0]).Value = 3;
                    form.PerformLayout(); using (var pageOverlay = new PreviewSelectionOverlay(form)) { pageOverlay.RefreshSnapshot(); using (var bitmap = new Bitmap(pageOverlay.Width, pageOverlay.Height)) { pageOverlay.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(folder, id + "." + page + ".png"), ImageFormat.Png); } }
                }
            }
        }
        using (var form = new WorksheetBatchForm("Notes", "B3:F20", true))
        using (var parking = new Panel { Size = form.Size })
        {
            Size client = form.ClientSize; form.FormBorderStyle = FormBorderStyle.None; form.ClientSize = client;
            form.TopLevel = false; parking.Controls.Add(form); form.Show(); form.PerformLayout();
            RenderWorksheetBatch(form, Path.Combine(folder, "conversion.worksheetBatch.png"));
            var viewport = (Panel)form.Controls.Find("batchFormatViewport", true)[0];
            viewport.AutoScrollPosition = new Point(0, viewport.VerticalScroll.Maximum);
            RenderWorksheetBatch(form, Path.Combine(folder, "conversion.worksheetBatch.scrolled.png"));
        }
    }
    private static void RenderWorksheetBatch(Form form, string path)
    {
        using (var overlay = new PreviewSelectionOverlay(form))
        {
            overlay.RefreshSnapshot();
            using (var bitmap = new Bitmap(overlay.Width, overlay.Height))
            {
                overlay.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size));
                bitmap.Save(path, ImageFormat.Png);
            }
        }
    }

    private static void Wheel(Control control, int delta)
    {
        Point screen = control.PointToScreen(new Point(control.Width / 2, control.Height / 2));
        int location = (screen.Y << 16) | (screen.X & 0xffff);
        SendMessage(control.Handle, 0x020A, new IntPtr(delta << 16), new IntPtr(location));
    }

    private static void WorksheetBatchLayout()
    {
        using (var form = DialogPreviewFactory.CreateProgramDefault("conversion.worksheetBatch"))
        using (var parking = new Panel { Size = form.Size })
        {
            form.TopLevel = false; parking.Controls.Add(form); form.Show(); form.PerformLayout();
            var review = (ListView)form.Controls.Find("batchReview", true)[0];
            var sheet = (TextBox)form.Controls.Find("batchSheetName", true)[0];
            var apply = (Button)form.Controls.Find("batchApply", true)[0];
            Check(review.Items.Count == 3 && review.Height >= 120 && !apply.Enabled,
                "worksheet batch sample presents an inert preview " + review.Bounds + ", items=" + review.Items.Count);
            var undo = (WorksheetUndoButton)form.Controls.Find("worksheetUndo", true).Single();
            Check(undo.Parent == apply.Parent && !undo.Enabled && !Descendants(form).Any(control => control is Label && control.Text == form.Text),
                "worksheet batch shows one window title and an adjacent idle undo control");
            Check(sheet.Parent is RoundedEditor && sheet.Width >= 180, "worksheet batch sheet field has a rounded frame and editable width");
            var textStyle = (ComboBox)form.Controls.Find("batchTextStyle", true)[0];
            var fonts = (ComboBox)form.Controls.Find("batchFontFamily", true)[0];
            Check(textStyle.Items.Cast<string>().SequenceEqual(new[] { "保持原样", "加粗", "斜体", "下划线", "删除线" }) &&
                textStyle.Parent is RoundedEditor && fonts.Items.Count > 1 && fonts.AutoCompleteSource == AutoCompleteSource.ListItems,
                "text styles and searchable local fonts use rounded input frames");
        }
        using (var form = new WorksheetBatchForm("Notes", "A1:A2", true))
        using (var parking = new Panel { Size = form.Size })
        {
            form.TopLevel = false; parking.Controls.Add(form); form.Show(); form.PerformLayout();
            var apply = (Button)form.Controls.Find("batchApply", true)[0];
            var review = (ListView)form.Controls.Find("batchReview", true)[0];
            var sheet = (TextBox)form.Controls.Find("batchSheetName", true)[0];
            Check(apply.Enabled, "changed preview enables apply");
            var clear = (CheckBox)form.Controls.Find("batchClearFormatting", true)[0];
            var alignment = (ComboBox)form.Controls.Find("batchHorizontal", true)[0];
            var style = (ComboBox)form.Controls.Find("batchTextStyle", true)[0];
            var replacement = (TextBox)form.Controls.Find("batchReplacement", true)[0];
            var vertical = (ComboBox)form.Controls.Find("batchVertical", true)[0];
            var viewport = (Panel)form.Controls.Find("batchFormatViewport", true)[0];
            var replaceCard = (UiCard)form.Controls.Find("batchTextCard", true)[0];
            var formatCard = (UiCard)form.Controls.Find("batchFormatCard", true)[0];
            Check(replaceCard.Parent == formatCard.Parent &&
                replaceCard.Top == formatCard.Top && replaceCard.Right < formatCard.Left,
                "text replacement and formatting are side-by-side");
            Point leftPosition = replacement.PointToScreen(Point.Empty), clearPosition = clear.PointToScreen(Point.Empty);
            Check(viewport.VerticalScroll.Visible && !viewport.HorizontalScroll.Visible && viewport.Contains(vertical),
                "the right-hand viewport contains the additional format selectors and scrolls only vertically");
            int styleIndex = style.SelectedIndex;
            Wheel(style, -120);
            Check(viewport.AutoScrollPosition.Y < 0 && style.SelectedIndex == styleIndex && apply.Enabled,
                "wheel on a closed selector scrolls the panel without changing choices or preview");
            Check(replacement.PointToScreen(Point.Empty) == leftPosition && clear.PointToScreen(Point.Empty) == clearPosition,
                "replacement and clear-format controls stay in place when the format list scrolls");
            Wheel(viewport, 120);
            Check(viewport.AutoScrollPosition.Y == 0, "upward wheel returns to the initial format options");
            form.ActiveControl = vertical;
            Point end = viewport.PointToClient(vertical.PointToScreen(new Point(0, vertical.Height)));
            Check(viewport.AutoScrollPosition.Y < 0 && end.Y <= viewport.ClientSize.Height,
                "focusing the last selector brings it into the scroll viewport");
            int oldIndex = alignment.SelectedIndex;
            clear.Checked = true;
            Check(!apply.Enabled && review.Items.Count == 0 && !alignment.Enabled && !vertical.Enabled && replacement.Enabled,
                "clear invalidates preview and locks appearance selectors while text replacement stays available");
            WorksheetBatchRequest submitted = null;
            form.Run += (request, previous) =>
            {
                submitted = request;
                var result = new WorksheetBatchReport();
                result.Rows.Add(new WorksheetBatchRow { Workbook = "sample.xlsx", Changed = 1, Detail = "可处理" });
                return result;
            };
            typeof(WorksheetBatchForm).GetMethod("RunBatch", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { false });
            Check(submitted != null && submitted.ClearFormatting && submitted.ReplacesText && submitted.TextStyle == WorksheetBatchTextStyle.Keep &&
                submitted.FontName == null && !submitted.Horizontal.HasValue && submitted.NumberFormat == null && apply.Enabled,
                "clear submits only the requested reset and text replacement");
            clear.Checked = false;
            Check(alignment.Enabled && alignment.SelectedIndex == oldIndex && !apply.Enabled, "leaving clear mode restores choices and requires a new preview");
            vertical.SelectedItem = "居中";
            typeof(WorksheetBatchForm).GetMethod("RunBatch", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, new object[] { false });
            Check(submitted.Horizontal == WorksheetBatchHorizontal.Justify && submitted.Vertical == WorksheetBatchVertical.Center &&
                submitted.TextStyle == WorksheetBatchTextStyle.Italic && !submitted.ClearFormatting, "preview receives selected style and both alignments");
            form.ClientSize = new Size(760, 620);
            form.PerformLayout();
            Check(form.AutoScrollMinSize.Height >= form.ClientSize.Height && alignment.Width >= 100 && apply.Width >= 100 &&
                viewport.VerticalScroll.Visible && !viewport.HorizontalScroll.Visible,
                "compact size retains reachable controls through scrolling");
            sheet.Text = "Different";
            Check(!apply.Enabled && review.Items.Count == 0, "editing target invalidates previous preview");
        }
    }
    private static IEnumerable<Control> Descendants(Control root)
    { foreach (Control item in root.Controls) { yield return item; foreach (Control child in Descendants(item)) yield return child; } }
    private static void LayoutChecks()
    {
        foreach (string id in new[] { "conversion.text", "formula.references", "conversion.batchInsertTrimDelete" })
        foreach (float factor in new[] { 1f, 1.25f, 1.5f, 2f })
        {
            var fonts = new List<Font>();
            try
            {
                using (var f = (CellConversionForm)DialogPreviewFactory.CreateProgramDefault(id))
                using (var parking = new Panel { Size = new Size(2200, 1800) })
                {
                    f.FormBorderStyle = FormBorderStyle.None; f.TopLevel = false; f.SuppressActivation = true; parking.Controls.Add(f); f.Show();
                    // The invisible parking panel intentionally prevents the native Shown event.
                    typeof(Form).GetMethod("OnShown", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(f, new object[] { EventArgs.Empty });
                    var sizes = new Dictionary<Control, float>(); foreach (Control control in Descendants(f)) sizes[control] = control.Font.Size;
                    if (factor != 1)
                    {
                        f.Scale(new SizeF(factor, factor));
                        foreach (var item in sizes) { var font = new Font(item.Key.Font.FontFamily, item.Value * factor, item.Key.Font.Style); fonts.Add(font); item.Key.Font = font; }
                    }
                    string address = f.RangeInput.Text;
                    var state = new ConversionCellState { Display = "MIRCAI文化旅游公司2025!!", Value = "MIRCAI文化旅游公司2025!!", NumberFormat = "General" };
                    if (id == "formula.references") { state.IsFormula = true; state.Formula = "=H23+H24+H26"; state.Value = 100.0; }
                    f.ReceivePreview(0, address, new ConversionPreview(address, state), null); f.PerformLayout();
                    if (id == "conversion.batchInsertTrimDelete") Check(f.Request().TextEdit.Operation == BatchTextEditOperation.Insert && After(f).Contains("示例"), "factory batch preview stays valid during native tab creation " + factor);
                    var bar = f.Apply.Parent; var root = bar.Parent;
                    int buttonBottom = root.Top + bar.Bottom;
                    Check(Math.Max(f.Undo.Bottom, Math.Max(f.Apply.Bottom, f.Cancel.Bottom)) <= bar.Height, "footer buttons are contained at scale " + factor);
                    Check(f.Undo.Parent == bar && f.Undo.Width > 0 && !f.Undo.Enabled, "adjacent conversion undo stays disabled in an inert preview " + id + " " + factor);
                    if (id != "formula.references") Check(!Descendants(f).Any(control => control is Label && control.Text == f.Text), "conversion uses one window title " + id);
                    int gap = f.ClientSize.Height - buttonBottom;
                    Check(f.AutoScrollMinSize.Height > f.ClientSize.Height || gap >= 0 && gap <= root.Padding.Bottom + 2, "compact footer in scaled layout " + id + " " + factor + ": gap=" + gap + "; padding=" + root.Padding.Bottom + "; preferred=" + f.AutoScrollMinSize + "; client=" + f.ClientSize + "; root=" + root.Bounds + "; bar=" + bar.Bounds + "; apply=" + f.Apply.Bounds);
                    Check(f.RangeInput.Width >= 180 && f.Apply.Width > 0 && f.Controls.Find("conversionBefore", true)[0].Height > 0, "visible real controls in scaled layout " + id + " " + factor);
                    if (factor == 1)
                    {
                        typeof(Form).GetMethod("OnResizeBegin", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(f, new object[] { EventArgs.Empty });
                        f.ClientSize = new Size(f.ClientSize.Width + 60, f.ClientSize.Height + 50);
                        typeof(Form).GetMethod("OnResizeEnd", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(f, new object[] { EventArgs.Empty });
                        Size resized = f.Size;
                        f.ReceivePreview(0, address, new ConversionPreview(address, state), null);
                        Check(f.Size == resized, "refresh preserves manually resized window " + id);
                        state.Display = new string('中', 3000); state.Formula = "=" + new string('A', 3000);
                        f.ReceivePreview(0, address, new ConversionPreview(address, state), null);
                        Check(f.AutoScrollMinSize.Height > f.ClientSize.Height && f.Size == resized, "long sample wraps and scrolls without replacing user size " + id + "; content=" + f.AutoScrollMinSize + "; outer=" + f.Size + "; saved=" + resized);
                    }
                }
            }
            finally { foreach (Font font in fonts) font.Dispose(); }
        }
        foreach (string id in new[] { "conversion.text", "formula.references", "conversion.batchInsertTrimDelete" })
        using (var form = (CellConversionForm)DialogPreviewFactory.Create(id, new Size(900, 680)))
        using (var parking = new Panel { Size = form.Size })
        {
            form.TopLevel = false; parking.Controls.Add(form); form.Show(); Size saved = form.ClientSize;
            form.ReceivePreview(0, form.RangeInput.Text, new ConversionPreview(form.RangeInput.Text, new ConversionCellState { Display = "更改示例", Value = "更改示例" }), null);
            Check(DialogSizing.HasOverride(form) && saved == form.ClientSize, "refresh preserves explicit saved size " + id);
        }
    }
    [STAThread] private static int Main(string[] args)
    {
        string settings = Path.Combine(Path.GetTempPath(), "ConversionUiProbe-" + Guid.NewGuid().ToString("N")); DialogSizeStore.RootOverride = settings;
        try { InteractionChecks(); LayoutChecks(); WorksheetBatchLayout(); if (args.Length == 2 && args[0] == "--render") Render(args[1]); Console.WriteLine("PASS ConversionInteractionProbe: " + checks + " assertions; hidden real controls, STA previews, scaled layouts, picker restore, busy and shutdown; no host connection"); return 0; }
        catch (Exception error) { Console.WriteLine("FAIL ConversionInteractionProbe: " + error); return 1; }
        finally { ConversionInteraction.Shutdown(); ConversionWindowThread.CreatedForProbe = null; ConversionInteraction.ResultForProbe = null; }
    }
}
