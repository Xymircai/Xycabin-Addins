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
            RenderForm(form, Path.Combine(folder, "conversion.worksheetBatch.png"));
            var scroll = form.Controls.Find("batchFormatScroll", true)[0];
            PressKey(scroll, Keys.End);
            RenderForm(form, Path.Combine(folder, "conversion.worksheetBatch.scrolled.png"));
        }
    }
    private static void RenderForm(Form form, string path)
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

    private static void FormulaWindowCaptions(string renderFolder)
    {
        if (renderFolder != null) Directory.CreateDirectory(renderFolder);
        foreach (var dialog in DialogCatalog.All.Where(item => item.Id.StartsWith("formula.", StringComparison.Ordinal)))
        using (var form = DialogPreviewFactory.CreateProgramDefault(dialog.Id))
        {
            string caption = DialogPreviewFactory.OriginalTitle(form);
            Check(!string.IsNullOrWhiteSpace(caption) && form.Text.EndsWith(caption, StringComparison.Ordinal),
                "formula keeps its window caption " + dialog.Id);
            Check(!Descendants(form).OfType<Label>().Any(label => label.Text == caption),
                "formula does not repeat the window caption in the body " + dialog.Id);
            if (renderFolder == null) continue;
            using (var parking = new Panel { Size = new Size(2400, 1800) })
            {
                Size client = form.ClientSize;
                form.FormBorderStyle = FormBorderStyle.None;
                form.ClientSize = client;
                form.TopLevel = false;
                parking.Controls.Add(form);
                form.Show();
                form.PerformLayout();
                RenderForm(form, Path.Combine(renderFolder, dialog.Id + ".png"));
            }
        }
    }

    private static void Wheel(Control control, int delta)
    {
        Point screen = control.PointToScreen(new Point(control.Width / 2, control.Height / 2));
        int location = (screen.Y << 16) | (screen.X & 0xffff);
        SendMessage(control.Handle, 0x020A, new IntPtr(delta << 16), new IntPtr(location));
    }

    private static KeyEventArgs PressKey(Control control, Keys key)
    {
        var e = new KeyEventArgs(key);
        typeof(Control).GetMethod("OnKeyDown", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(control, new object[] { e });
        return e;
    }

    private static Rectangle In(Control container, Control child)
    { return new Rectangle(container.PointToClient(child.PointToScreen(Point.Empty)), child.Size); }

    private static bool TextIsCentered(Bitmap bitmap)
    {
        int left = bitmap.Width, right = -1;
        for (int y = 1; y < bitmap.Height - 2; y++)
        for (int x = 1; x < bitmap.Width - 1; x++)
            if (bitmap.GetPixel(x, y).GetBrightness() < 0.55f)
            {
                left = Math.Min(left, x);
                right = Math.Max(right, x);
            }
        return left <= right && Math.Abs(left + right + 1 - bitmap.Width) <= 10;
    }

    private static void ReviewAlignment(ListView review)
    {
        var events = BindingFlags.Instance | BindingFlags.NonPublic;
        var header = typeof(ListView).GetMethod("OnDrawColumnHeader", events);
        var cell = typeof(ListView).GetMethod("OnDrawSubItem", events);
        ListViewItem item = review.Items[0];
        foreach (ColumnHeader column in review.Columns)
        using (var bitmap = new Bitmap(220, 32))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            var bounds = new Rectangle(Point.Empty, bitmap.Size);
            var title = new DrawListViewColumnHeaderEventArgs(graphics, bounds, column.Index, column,
                ListViewItemStates.Default, review.ForeColor, review.BackColor, review.Font);
            header.Invoke(review, new object[] { title });
            Check(!title.DrawDefault && (SystemInformation.HighContrast ||
                (bitmap.GetPixel(3, 3).ToArgb() == Color.FromArgb(253, 232, 235).ToArgb() && TextIsCentered(bitmap))),
                "preview headers are centered on a light red fill " + column.Text);
            var subItem = new DrawListViewSubItemEventArgs(graphics, bounds, item, item.SubItems[column.Index],
                0, column.Index, column, ListViewItemStates.Default);
            cell.Invoke(review, new object[] { subItem });
            Check(!subItem.DrawDefault && (SystemInformation.HighContrast || TextIsCentered(bitmap)),
                "preview text is centered even in the first column " + column.Text);
        }
        Check(review.ShowItemToolTips && item.ToolTipText.Contains(item.Text),
            "clipped preview values retain the workbook and outcome tooltip");
    }

    private static void CheckBlueChrome(Control choice, Control scroll)
    {
        if (SystemInformation.HighContrast) return;
        using (var normal = new Bitmap(choice.Width, choice.Height))
        using (var hovered = new Bitmap(choice.Width, choice.Height))
        using (var track = new Bitmap(scroll.Width, scroll.Height))
        {
            Rectangle bounds = new Rectangle(Point.Empty, choice.Size);
            choice.DrawToBitmap(normal, bounds);
            typeof(Control).GetMethod("OnMouseEnter", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(choice, new object[] { EventArgs.Empty });
            choice.DrawToBitmap(hovered, bounds);
            scroll.DrawToBitmap(track, new Rectangle(Point.Empty, track.Size));
            int valueX = choice.Width / 2, middle = choice.Height / 2;
            int arrowX = choice.Width - Ui.Px(choice, 18);
            Check(normal.GetPixel(valueX, middle).ToArgb() == choice.BackColor.ToArgb() &&
                hovered.GetPixel(valueX, middle).ToArgb() == choice.BackColor.ToArgb(),
                "the closed choice keeps a plain background on hover");
            Check(normal.GetPixel(arrowX, middle).ToArgb() == Ui.SelectBack.ToArgb() &&
                hovered.GetPixel(arrowX, middle).ToArgb() == Ui.Btn2Down.ToArgb(),
                "the closed choice paints only its arrow with the blue hover style");
            Check(track.GetPixel(scroll.Width / 2, scroll.Height / 4).ToArgb() == Ui.Blue.ToArgb(),
                "the formatting scrollbar paints a blue thumb");
        }
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
            ReviewAlignment(review);
            var undo = (WorksheetUndoButton)form.Controls.Find("worksheetUndo", true).Single();
            Check(undo.Parent == apply.Parent && !undo.Enabled && !Descendants(form).Any(control => control is Label && control.Text == form.Text),
                "worksheet batch shows one window title and an adjacent idle undo control");
            Check(sheet.Parent is RoundedEditor && sheet.Width >= 180, "worksheet batch sheet field has a rounded frame and editable width");
            var textStyle = (ComboBox)form.Controls.Find("batchTextStyle", true)[0];
            var fonts = (ComboBox)form.Controls.Find("batchFontFamily", true)[0];
            Check(textStyle.Items.Cast<string>().SequenceEqual(new[] { "保持原样", "加粗", "斜体", "下划线", "删除线" }) &&
                textStyle.Parent is RoundedEditor && fonts.Items.Count > 1 && fonts.AutoCompleteSource == AutoCompleteSource.ListItems,
                "text styles and searchable local fonts use rounded input frames");
            var methods = BindingFlags.Instance | BindingFlags.NonPublic;
            int closedHeight = textStyle.Height;
            textStyle.GetType().GetMethod("MeasureList", methods).Invoke(textStyle, null);
            int rowHeight = textStyle.GetItemHeight(0);
            Check(textStyle.DrawMode == DrawMode.OwnerDrawVariable && rowHeight > textStyle.Font.Height &&
                textStyle.DropDownHeight >= rowHeight * textStyle.Items.Count && textStyle.Height == closedHeight,
                "expanded row height " + rowHeight + ", list height " + textStyle.DropDownHeight +
                ", closed editor " + closedHeight + " -> " + textStyle.Height);
            if (!SystemInformation.HighContrast)
            using (var bitmap = new Bitmap(170, rowHeight))
            using (var graphics = Graphics.FromImage(bitmap))
            {
                int selected = textStyle.SelectedIndex;
                var bounds = new Rectangle(Point.Empty, bitmap.Size);
                textStyle.GetType().GetMethod("OnDrawItem", methods).Invoke(textStyle, new object[] {
                    new DrawItemEventArgs(graphics, textStyle.Font, bounds, selected, DrawItemState.Selected)
                });
                Check(bitmap.GetPixel(0, 0).ToArgb() == textStyle.BackColor.ToArgb() &&
                    bitmap.GetPixel(153, bitmap.Height / 2).ToArgb() == textStyle.BackColor.ToArgb(),
                    "native selection does not paint a blue background without a pointer over the row");
            }
        }
        using (var form = new WorksheetBatchForm("Notes", "A1"))
        {
            var number = (ComboBox)form.Controls.Find("batchNumberFormat", true)[0];
            Check(number.Items.Cast<string>().SequenceEqual(new[]
            {
                "保持原样", "常规格式", "文本格式", "千分位 · 两位小数", "百分比 · 两位小数", "日期 · 年-月-日"
            }), "number format choices have the requested labels and order");
            var formats = new[] { null, "General", "@", "#,##0.00", "0.00%", "yyyy-mm-dd" };
            var request = typeof(WorksheetBatchForm).GetMethod("Request", BindingFlags.Instance | BindingFlags.NonPublic);
            for (int i = 0; i < formats.Length; i++)
            {
                number.SelectedIndex = i;
                Check(((WorksheetBatchRequest)request.Invoke(form, null)).NumberFormat == formats[i],
                    "number format choice maps to the workbook format " + number.Text);
            }
            var fonts = (ComboBox)form.Controls.Find("batchFontFamily", true)[0];
            Check(fonts.Items.Count == 1, "opening the working dialog defers the installed font scan");
            fonts.GetType().GetMethod("EnsureItems", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public).Invoke(fonts, null);
            Check(fonts.Items.Count > 1 && fonts.SelectedIndex == 0 && fonts.AutoCompleteSource == AutoCompleteSource.ListItems,
                "the font list loads on demand with type-ahead and the original selection");
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
            var scroll = form.Controls.Find("batchFormatScroll", true)[0];
            var formatTable = (TableLayoutPanel)style.Parent.Parent;
            var clip = formatTable.Parent;
            var source = (ComboBox)form.Controls.Find("batchSource", true)[0];
            var replaceCard = (UiCard)form.Controls.Find("batchTextCard", true)[0];
            var formatCard = (UiCard)form.Controls.Find("batchFormatCard", true)[0];
            Check(clear.Text == "清除格式（重置常用格式，保留字号、边框、条件格式）" &&
                !Descendants(formatCard).OfType<Label>().Any(label => label.Text.StartsWith("重置常用格式")),
                "clear formatting has one checkbox caption with the full inline explanation");
            Check(apply.Parent.Controls.Cast<Control>().Select(button => button.Name)
                .SequenceEqual(new[] { "batchApply", "worksheetUndo", "batchCheck" }),
                "the footer has three actions and the window chrome provides close");
            Check(replaceCard.Parent == formatCard.Parent &&
                replaceCard.Top == formatCard.Top && replaceCard.Right < formatCard.Left,
                "text replacement and formatting are side-by-side");
            Point leftPosition = replacement.PointToScreen(Point.Empty), clearPosition = clear.PointToScreen(Point.Empty);
            Check(scroll.Visible && scroll.AccessibleRole == AccessibleRole.ScrollBar &&
                !viewport.HorizontalScroll.Visible && viewport.Contains(vertical) && clip.Contains(vertical),
                "the right-hand viewport contains the additional format selectors and scrolls only vertically");
            int sourceIndex = source.SelectedIndex;
            Control active = form.ActiveControl;
            CheckBlueChrome(source, scroll);
            Check(source.SelectedIndex == sourceIndex && form.ActiveControl == active && apply.Enabled,
                "hovering the target workbook choice does not change selection, focus or preview");
            int styleIndex = style.SelectedIndex;
            Wheel(style, -120);
            bool wheelEnabled = SystemInformation.MouseWheelScrollLines != 0;
            Check((formatTable.Top < 0) == wheelEnabled && style.SelectedIndex == styleIndex && apply.Enabled,
                "wheel on a closed selector scrolls the panel without changing choices or preview");
            Check(replacement.PointToScreen(Point.Empty) == leftPosition && clear.PointToScreen(Point.Empty) == clearPosition,
                "replacement and clear-format controls stay in place when the format list scrolls");
            Wheel(viewport, 120);
            Check(formatTable.Top == 0, "upward wheel returns to the initial format options");
            Check(PressKey(scroll, Keys.End).Handled && formatTable.Top < 0 && clip.ClientRectangle.Contains(In(clip, vertical)),
                "End on the scrollbar brings the last selector into view");
            int bottom = formatTable.Top;
            Check(PressKey(scroll, Keys.PageUp).Handled && formatTable.Top > bottom && PressKey(scroll, Keys.Home).Handled && formatTable.Top == 0,
                "Page Up and Home move the custom scrollbar");
            int x = scroll.Width / 2, y = scroll.Height / 4;
            var mouse = BindingFlags.Instance | BindingFlags.NonPublic;
            typeof(Control).GetMethod("OnMouseDown", mouse).Invoke(scroll, new object[] { new MouseEventArgs(MouseButtons.Left, 1, x, y, 0) });
            typeof(Control).GetMethod("OnMouseMove", mouse).Invoke(scroll, new object[] { new MouseEventArgs(MouseButtons.Left, 0, x, scroll.Height - 2, 0) });
            typeof(Control).GetMethod("OnMouseUp", mouse).Invoke(scroll, new object[] { new MouseEventArgs(MouseButtons.Left, 1, x, scroll.Height - 2, 0) });
            Check(formatTable.Top == bottom && !scroll.Capture, "dragging the painted thumb reaches the bottom and releases capture");
            PressKey(scroll, Keys.Home);
            form.ActiveControl = vertical;
            Check(formatTable.Top < 0 && clip.ClientRectangle.Contains(In(clip, vertical)),
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
            Check(alignment.Width >= 96 && apply.Width >= 100 && scroll.Visible && !viewport.HorizontalScroll.Visible,
                "compact size retains reachable controls through scrolling");
            sheet.Text = "Different";
            Check(!apply.Enabled && review.Items.Count == 0, "editing target invalidates previous preview");
        }
    }

    private static void WorksheetBatchScaledLayout(string renderFolder)
    {
        foreach (float factor in new[] { 1f, 1.25f, 1.5f, 2f })
        {
            var fonts = new List<Font>();
            try
            {
                using (var form = new WorksheetBatchForm("Notes", "B3:F20", true))
                using (var parking = new Panel { Size = new Size(2400, 2000) })
                {
                    Size client = form.ClientSize;
                    form.FormBorderStyle = FormBorderStyle.None;
                    form.ClientSize = client;
                    form.TopLevel = false;
                    parking.Controls.Add(form);
                    form.Show();
                    if (factor != 1)
                    {
                        var sizes = new[] { (Control)form }.Concat(Descendants(form)).ToDictionary(control => control, control => control.Font.SizeInPoints);
                        form.Scale(new SizeF(factor, factor));
                        foreach (var item in sizes)
                        {
                            var scaled = new Font(item.Key.Font.FontFamily, item.Value * factor, item.Key.Font.Style);
                            fonts.Add(scaled);
                            item.Key.Font = scaled;
                        }
                    }
                    string samplePath = @"C:\演示目录\" + string.Concat(Enumerable.Repeat("工作簿批量处理目录\\", 12));
                    typeof(WorksheetBatchForm).GetField("folder", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(form, samplePath);
                    var source = (ComboBox)form.Controls.Find("batchSource", true)[0];
                    source.SelectedIndex = (int)WorksheetBatchSource.Folder;
                    foreach (int width in new[] { 900, 760 })
                    {
                        form.ClientSize = new Size((int)(width * factor), (int)(700 * factor));
                        form.PerformLayout();
                        string size = width + " at " + factor;
                        var target = form.Controls.Find("batchTargetCard", true)[0];
                        var detail = Descendants(target).OfType<Label>().Single(label => label.Text == samplePath);
                        var select = Descendants(target).OfType<Button>().Single(button => button.Text.Contains("文件夹"));
                        Rectangle choice = In(target, source), frame = In(target, source.Parent), detailBounds = In(target, detail);
                        Check(Descendants(target).OfType<Label>().Any(label => label.Text == "目标与区域") &&
                            Descendants(target).OfType<Label>().Any(label => label.Text == "目标工作簿"),
                            "worksheet batch uses the new target captions " + size);
                        Check(frame.Contains(choice) && In(target, select).Bottom <= detailBounds.Top && frame.Bottom <= detailBounds.Top &&
                            detailBounds.Bottom <= In(target, form.Controls.Find("batchSheetName", true)[0].Parent).Top,
                            "target choice, chooser and long source details stay in separate rows " + size + ": " + choice + " in " + frame + " before " + detailBounds);
                        Check(detail.Height >= TextRenderer.MeasureText(detail.Text, detail.Font,
                            new Size(detail.Width, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height,
                            "the source details are tall enough for the wrapped path " + size);
                        var format = form.Controls.Find("batchFormatCard", true)[0];
                        var text = form.Controls.Find("batchTextCard", true)[0];
                        Check(text.Top == format.Top && text.Right < format.Left, "format card stays to the right " + size);
                        foreach (string caption in new[] { "文本格式", "数字格式", "字体", "字体颜色", "填充颜色", "水平对齐", "垂直对齐" })
                        {
                            var label = Descendants(format).OfType<Label>().Single(item => item.Text == caption);
                            Size line = TextRenderer.MeasureText(caption, label.Font, Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
                            Check(label.Width >= line.Width && label.Height == line.Height,
                                "format label fits on one line " + caption + " " + size);
                        }
                        CheckBatchBounds(form, size);
                        if (renderFolder != null && factor == 1.5f && width == 900)
                            RenderForm(form, Path.Combine(renderFolder, "conversion.worksheetBatch.scaled.png"));
                    }
                }
            }
            finally { foreach (var font in fonts) font.Dispose(); }
        }
    }

    private static void CheckBatchBounds(Form form, string size)
    {
        var card = form.Controls.Find("batchFormatCard", true)[0];
        var viewport = form.Controls.Find("batchFormatViewport", true)[0];
        var scroll = form.Controls.Find("batchFormatScroll", true)[0];
        var clear = (CheckBox)form.Controls.Find("batchClearFormatting", true)[0];
        Rectangle content = card.ClientRectangle;
        content.Width -= card.Padding.Right;
        Check(content.Contains(In(card, viewport)) && content.Contains(In(card, scroll)) &&
            content.Contains(In(card, clear)) && clear.Width >= clear.GetPreferredSize(Size.Empty).Width,
            "the viewport, scrollbar and full clear caption fit the format card " + size);
        foreach (string name in new[] { "batchTextStyle", "batchNumberFormat", "batchFontColor", "batchFillColor" })
        {
            Control editor = form.Controls.Find(name, true)[0];
            var fields = (TableLayoutPanel)editor.Parent.Parent;
            var clip = fields.Parent;
            var swatch = fields.GetControlFromPosition(2, fields.GetRow(editor.Parent));
            Check(clip.ClientRectangle.Contains(In(clip, editor.Parent)) &&
                editor.Parent.ClientRectangle.Contains(In(editor.Parent, editor)) &&
                (swatch == null || swatch == editor.Parent || clip.ClientRectangle.Contains(In(clip, swatch))),
                "the complete format row, including the color button, fits " + name + " " + size);
        }
        var textCard = form.Controls.Find("batchTextCard", true)[0];
        var find = form.Controls.Find("batchFind", true)[0];
        Check(In(textCard, find.Parent).Right <= textCard.ClientSize.Width - textCard.Padding.Right,
            "the entire find editor fits the text replacement card " + size);
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateDesktop(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr attributes);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetThreadDesktop(IntPtr desktop);
    [DllImport("user32.dll")]
    private static extern bool CloseDesktop(IntPtr desktop);
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();
    [DllImport("user32.dll")]
    private static extern bool GetComboBoxInfo(IntPtr combo, ref ComboWindowInfo info);
    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr keys, ref NativeRect rectangle);
    [DllImport("user32.dll")]
    private static extern bool ValidateRect(IntPtr window, IntPtr rect);
    [DllImport("user32.dll")]
    private static extern bool GetUpdateRect(IntPtr window, out NativeRect rect, bool erase);

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        internal int Left, Top, Right, Bottom;

        internal Rectangle Bounds { get { return Rectangle.FromLTRB(Left, Top, Right, Bottom); } }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ComboWindowInfo
    {
        internal int Size;
        internal NativeRect Item, Button;
        internal int ButtonState;
        internal IntPtr Combo, Edit, List;
    }

    private static void PressNativeKey(Control control, Keys key)
    {
        const int KeyDown = 0x0100;
        var message = Message.Create(control.Handle, KeyDown, (IntPtr)key, IntPtr.Zero);
        if (!control.PreProcessMessage(ref message))
            SendMessage(control.Handle, KeyDown, (IntPtr)key, IntPtr.Zero);
    }

    private static Rectangle ListItemBounds(IntPtr list, int index)
    {
        const int GetItemRect = 0x0198;
        NativeRect rect = new NativeRect();
        if (SendMessage(list, GetItemRect, (IntPtr)index, ref rect).ToInt64() == -1)
            throw new InvalidOperationException("Could not read native list item " + index);
        return rect.Bounds;
    }

    private static int ListTopIndex(IntPtr list)
    {
        const int GetTopIndex = 0x018e;
        return SendMessage(list, GetTopIndex, IntPtr.Zero, IntPtr.Zero).ToInt32();
    }

    private static void MoveOverListItem(IntPtr list, int index, int offset = 0)
    {
        Rectangle row = ListItemBounds(list, index);
        const int MouseMove = 0x0200;
        SendMessage(list, MouseMove, IntPtr.Zero, (IntPtr)(((row.Top + row.Height / 2) << 16) | (row.Left + row.Width / 2 + offset)));
    }

    private static void CheckListPaint(ComboBox choice, IntPtr list, int expectedRow, int visibleRows)
    {
        const int PrintClient = 0x0318, PrintClientArea = 4;
        int top = ListTopIndex(list);
        var bounds = ListItemBounds(list, top + visibleRows - 1);
        using (var bitmap = new Bitmap(bounds.Right, bounds.Bottom))
        using (var graphics = Graphics.FromImage(bitmap))
        {
            IntPtr dc = graphics.GetHdc();
            try { SendMessage(list, PrintClient, dc, (IntPtr)PrintClientArea); }
            finally { graphics.ReleaseHdc(dc); }
            Color plain = SystemInformation.HighContrast ? SystemColors.Window : choice.BackColor;
            Color blue = SystemInformation.HighContrast ? SystemColors.Highlight : Ui.SelectBack;
            for (int index = top; index < top + visibleRows; index++)
            {
                Rectangle row = ListItemBounds(list, index);
                Color pixel = bitmap.GetPixel(row.Right - 14, row.Top + row.Height / 2);
                Check(pixel.ToArgb() == (index == expectedRow ? blue : plain).ToArgb(),
                    choice.Name + " native list row " + index + " expects hover on " + expectedRow + ", painted " + pixel);
            }
        }
    }

    private static IntPtr OpenNativeList(ComboBox choice)
    {
        choice.Focus();
        PressNativeKey(choice, Keys.F4);
        var info = new ComboWindowInfo { Size = Marshal.SizeOf(typeof(ComboWindowInfo)) };
        Check(choice.DroppedDown && GetComboBoxInfo(choice.Handle, ref info) && info.List != IntPtr.Zero,
            choice.Name + " opens its native list from the keyboard");
        return info.List;
    }

    private static void CheckNativeChoice(Form form, string name)
    {
        var choice = (ComboBox)form.Controls.Find(name, true)[0];
        int commits = 0;
        choice.SelectionChangeCommitted += delegate { commits++; };
        IntPtr list = OpenNativeList(choice);
        int visible = Math.Min(choice.Items.Count, choice.MaxDropDownItems);
        CheckListPaint(choice, list, -1, visible);
        MoveOverListItem(list, 1);
        CheckListPaint(choice, list, 1, visible);
        ValidateRect(list, IntPtr.Zero);
        MoveOverListItem(list, 1, 5);
        MoveOverListItem(list, 1, -5);
        NativeRect update;
        Check(!GetUpdateRect(list, out update, false), name + " moving within the highlighted row schedules no repaint");
        MoveOverListItem(list, 2);
        Check(!GetUpdateRect(list, out update, false) ||
            Rectangle.Union(ListItemBounds(list, 1), ListItemBounds(list, 2)).Contains(update.Bounds),
            name + " crossing into a new row repaints only the two affected rows");
        CheckListPaint(choice, list, 2, visible);
        const int MouseLeave = 0x02a3;
        SendMessage(list, MouseLeave, IntPtr.Zero, IntPtr.Zero);
        CheckListPaint(choice, list, -1, visible);
        Check(commits == 0, name + " mouse movement does not commit a choice");
        PressNativeKey(choice, Keys.Escape);
        Check(!choice.DroppedDown && !form.IsDisposed, name + " Esc closes only the drop-down");
        int selected = choice.SelectedIndex;
        list = OpenNativeList(choice);
        CheckListPaint(choice, list, -1, visible);
        PressNativeKey(choice, Keys.Down);
        CheckListPaint(choice, list, -1, visible);
        PressNativeKey(choice, Keys.Enter);
        Check(!choice.DroppedDown && choice.SelectedIndex == selected + 1, name + " commits its keyboard selection");
    }

    private static void CheckNativeFontList(Form form)
    {
        var fonts = (ComboBox)form.Controls.Find("batchFontFamily", true)[0];
        IntPtr list = OpenNativeList(fonts);
        Check(fonts.AutoCompleteSource == AutoCompleteSource.ListItems && fonts.Items.Count > fonts.MaxDropDownItems,
            "installed fonts load on first open, with type-ahead and additional scrollable rows");
        int first = ListTopIndex(list);
        MoveOverListItem(list, first + 1);
        CheckListPaint(fonts, list, first + 1, fonts.MaxDropDownItems);
        const int VerticalScroll = 0x0115, LineDown = 1;
        SendMessage(list, VerticalScroll, (IntPtr)LineDown, IntPtr.Zero);
        int next = ListTopIndex(list);
        Check(next > first, "the native font list scrolls");
        CheckListPaint(fonts, list, next + 1, fonts.MaxDropDownItems);
        PressNativeKey(fonts, Keys.Escape);
        Check(!fonts.DroppedDown && !form.IsDisposed, "Esc dismisses the font list");
    }

    private static void WorksheetBatchNative()
    {
        SetProcessDPIAware();
        IntPtr desktop = CreateDesktop("worksheet-batch-probe-" + Guid.NewGuid().ToString("N"), IntPtr.Zero, IntPtr.Zero,
            0, 0x000F01FF, IntPtr.Zero);
        if (desktop == IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        Exception failure = null;
        try
        {
            var thread = new Thread(delegate()
            {
                try
                {
                    if (!SetThreadDesktop(desktop)) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
                    Application.EnableVisualStyles();
                    using (var form = new WorksheetBatchForm("Sheet1", "V14"))
                    {
                        form.Show();
                        Application.DoEvents();
                        float scale = DialogSizing.Scale(form);
                        Console.WriteLine("Worksheet batch window at " + scale.ToString("P0") + " DPI scaling");
                        CheckBatchBounds(form, "native default window");
                        Check(!form.HorizontalScroll.Visible, "the default window fits both cards without horizontal scrolling");
                        CheckNativeChoice(form, "batchSource");
                        CheckNativeChoice(form, "batchTextStyle");
                        CheckNativeChoice(form, "batchNumberFormat");
                        int frameWidth = form.Width - form.ClientSize.Width, frameHeight = form.Height - form.ClientSize.Height;
                        form.Size = new Size((int)Math.Round(760 * scale) + frameWidth, (int)Math.Round(620 * scale) + frameHeight);
                        Application.DoEvents();
                        CheckBatchBounds(form, "native compact window with vertical scrolling");
                        Check(form.VerticalScroll.Visible && !form.HorizontalScroll.Visible,
                            "the short compact window fits the width and scrolls vertically");
                        var apply = form.Controls.Find("batchApply", true)[0];
                        form.ScrollControlIntoView(apply);
                        Check(form.ClientRectangle.Contains(In(form, apply)), "the footer is reachable in the short window");
                        CheckNativeFontList(form);
                        PressNativeKey(form.ActiveControl ?? form, Keys.Escape);
                        Check(form.IsDisposed, "Esc closes the idle window");
                    }
                    using (var form = new WorksheetBatchForm("Sheet1", "A1"))
                    {
                        var settings = new CustomizationData();
                        settings.Formats.Add("conversion.worksheetBatch::batchClearFormatting.1", new FormatRule { Width96 = 95 });
                        form.Load += delegate { DialogFormatting.Apply(form, "conversion.worksheetBatch", settings); };
                        form.Show();
                        CheckBatchBounds(form, "saved checkbox width from before the longer caption");
                        const int SystemCommand = 0x0112, CloseCommand = 0xf060;
                        bool ran = false;
                        form.Run += delegate
                        {
                            ran = true;
                            PressNativeKey(form, Keys.Escape);
                            SendMessage(form.Handle, SystemCommand, (IntPtr)CloseCommand, IntPtr.Zero);
                            Check(!form.IsDisposed, "Esc and title-bar close respect the busy state");
                            return new WorksheetBatchReport();
                        };
                        ((Button)form.Controls.Find("batchCheck", true)[0]).PerformClick();
                        Check(ran, "the preview action runs from the footer");
                        SendMessage(form.Handle, SystemCommand, (IntPtr)CloseCommand, IntPtr.Zero);
                        Check(form.IsDisposed, "the native title-bar close command closes the window");
                    }
                }
                catch (Exception error) { failure = error; }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(TimeSpan.FromSeconds(45))) throw new TimeoutException("Native worksheet batch probe did not finish");
            if (failure != null) throw new InvalidOperationException("Native worksheet batch probe", failure);
        }
        finally { CloseDesktop(desktop); }
    }

    private static IEnumerable<Control> Descendants(Control root)
    { foreach (Control item in root.Controls) { yield return item; foreach (Control child in Descendants(item)) yield return child; } }

    private static void FormulaCompactLayout()
    {
        string[] formulas = { "formula.abs", "formula.unwrap", "formula.custom", "formula.round", "formula.iferror", "formula.add",
            "formula.formulaAsText", "formula.restoreTextFormula", "formula.resultAsText", "formula.valuesOnly" };
        var events = BindingFlags.Instance | BindingFlags.NonPublic;
        foreach (string id in formulas)
        using (var form = DialogPreviewFactory.CreateProgramDefault(id))
        using (var parking = new Panel { Size = new Size(1400, 1100) })
        {
            var editor = form as IFormulaOperationDialog;
            if (editor != null && id != "formula.custom")
                editor.SetTarget(@"目标选区 C:\演示目录\" + new string('中', 95) + ".xlsx / 工作表 · $A$1:$D$30", true);
            form.TopLevel = false;
            parking.Controls.Add(form);
            form.Show();
            form.PerformLayout();
            int startingHeight = form.ClientSize.Height;
            Check(form.FormBorderStyle == FormBorderStyle.Sizable && Math.Abs(startingHeight - form.AutoScrollMinSize.Height) <= 2,
                "default formula height fits actual content " + id + " client=" + form.ClientSize + " content=" + form.AutoScrollMinSize);

            typeof(Form).GetMethod("OnResizeBegin", events).Invoke(form, new object[] { EventArgs.Empty });
            int minimum = (int)Math.Round(DialogCatalog.Find(id).Minimum.Height * DialogSizing.Scale(form));
            form.ClientSize = new Size(form.ClientSize.Width, minimum);
            typeof(Form).GetMethod("OnResizeEnd", events).Invoke(form, new object[] { EventArgs.Empty });
            Size resized = form.ClientSize;
            Check(resized.Height < startingHeight && form.AutoScrollMinSize.Height > resized.Height,
                "formula can shrink below its old floor with content still scrollable " + id);
            if (editor != null) editor.Resume("这是重新计算后的校验提示。");
            else
            {
                var conversion = (CellConversionForm)form;
                conversion.ReceivePreview(0, conversion.RangeInput.Text,
                    new ConversionPreview(conversion.RangeInput.Text, new ConversionCellState { Display = "示例", Value = "示例" }), null);
            }
            Check(form.ClientSize == resized, "live content preserves a manually resized formula window " + id);
            var close = (Control)form.CancelButton;
            form.ScrollControlIntoView(close);
            Check(form.ClientRectangle.Contains(In(form, close)),
                "the formula footer remains reachable when the window is shortened " + id + " " + In(form, close));
        }

        foreach (string id in new[] { "formula.abs", "formula.custom", "formula.valuesOnly" })
        using (var form = DialogPreviewFactory.Create(id, new Size(740, 410)))
        using (var parking = new Panel { Size = new Size(1400, 1100) })
        {
            form.TopLevel = false;
            parking.Controls.Add(form);
            form.Show();
            Check(DialogSizing.HasOverride(form) && DialogSizing.LogicalClient(form) == new Size(740, 410),
                "the shared fit respects an explicit formula size " + id);
        }
    }

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
                    Check(!Descendants(f).OfType<Label>().Any(label => label.Text == DialogPreviewFactory.OriginalTitle(f)), "conversion uses one window title " + id);
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
        FormatSession.SourceOverride = Path.Combine(settings, "Customization.xml");
        try
        {
            if (args.Length == 1 && args[0] == "--worksheet-batch-native")
            {
                WorksheetBatchNative();
                Console.WriteLine("PASS WorksheetBatchNative: " + checks + " assertions; private desktop, native list and top-level layout");
                return 0;
            }
            string renderFolder = args.Length == 2 && args[0] == "--render" ? args[1] : null;
            if (renderFolder != null) Directory.CreateDirectory(renderFolder);
            InteractionChecks();
            LayoutChecks();
            FormulaCompactLayout();
            WorksheetBatchLayout();
            WorksheetBatchScaledLayout(renderFolder);
            FormulaWindowCaptions(renderFolder);
            if (renderFolder != null) Render(renderFolder);
            Console.WriteLine("PASS ConversionInteractionProbe: " + checks + " assertions; hidden real controls, STA previews, scaled layouts, picker restore, busy and shutdown; no host connection");
            return 0;
        }
        catch (Exception error) { Console.WriteLine("FAIL ConversionInteractionProbe: " + error); return 1; }
        finally { ConversionInteraction.Shutdown(); ConversionWindowThread.CreatedForProbe = null; ConversionInteraction.ResultForProbe = null; }
    }
}
