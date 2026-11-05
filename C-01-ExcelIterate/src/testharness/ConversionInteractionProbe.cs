using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Excel_Iterate;

internal static class ConversionInteractionProbe
{
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
        host = Open(); var same = new Host(); long existing = ConversionInteraction.Open(same, CellConversionKind.TextFilter, IntPtr.Zero);
        Check(existing == ticket && same.Disposals == 1, "repeated command keeps a single parameter session"); OnUi(f => f.Close()); Settled();
        host = Open(); OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 1 && editor.Enabled);
        Check(host.Address == "A1" && host.Cell.Writes == 0 && ConversionInteraction.HasPending, "cancelled native picker restores pending window and original address");
        host.PickFailure = true; OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 2 && editor.Enabled && editor.Validation.Text.Contains("picker failure"));
        Check(host.Cell.Writes == 0, "picker exception restores editor without writes");
        host.PickFailure = false; host.Picked = "B1"; OnUi(f => f.Pick.PerformClick()); Pump(() => host.Picks == 3 && editor.Enabled && editor.RangeInput.Text == "B1" && editor.Apply.Enabled);
        OnUi(f => f.Apply.PerformClick()); Settled(); Check(result != null && !result.Failed && host.Address == "B1" && host.Cell.Current.Formula == "=A2" && host.Disposals == 1, "accepted picker fills address then explicit apply commits");
        long completed = ticket; int writes = host.Cell.Writes; ConversionInteraction.Receive(completed, false, new ConversionRequest { Kind = CellConversionKind.ValuesOnly, Address = "A1" }); Application.DoEvents(); Check(host.Cell.Writes == writes, "duplicate ticket ignored");
        host = Open(); ConversionInteraction.Receive(completed, false, null); Application.DoEvents(); Check(ConversionInteraction.HasPending, "old ticket cannot cancel a new session");
        Check(Connect.TryBeginConversion(), "fixture holds unrelated busy operation"); OnUi(f => f.Pick.PerformClick()); Pump(() => editor.Enabled && editor.Validation.Text.Contains("其他操作"));
        Check(host.Picks == 0 && !Connect.TryBeginConversion(), "busy rejection retains another operation guard"); Connect.EndConversion(); OnUi(f => f.Cancel.PerformClick()); Settled();
        host = Open(); host.ScopeFailure = true; OnUi(f => f.Apply.PerformClick()); Settled(); Check(result.Failed && host.Cell.Writes == 0 && host.Disposals == 1, "changed target is zero-write failure");
        host = Open(CellConversionKind.TextFilter); bool inspected = false;
        OnUi(delegate(CellConversionForm f)
        {
            Check(f.Request().Extract && f.Request().Kinds == CellCharacterKinds.Chinese, "default extract Chinese");
            f.RangeInput.Text = "Other!A1"; Check(!f.Apply.Enabled, "invalid address disables apply"); f.RangeInput.Text = "B1";
            for (int i = 0; i < 5; i++) ((CheckBox)f.Controls.Find("conversionKind" + i, true)[0]).Checked = false;
            Check(!f.Apply.Enabled, "no character categories disables apply"); ((CheckBox)f.Controls.Find("conversionKind3", true)[0]).Checked = true;
            Check(!f.Apply.Enabled, "manual address waits for fresh preview"); inspected = true;
        });
        Pump(() => inspected && editor.Apply.Enabled); OnUi(f => f.Apply.PerformClick()); Settled(); Check(!result.Failed && (string)host.Cell.Current.Value == "7" && host.Cell.Current.NumberFormat == "@", "manual address and text settings delivered");
        PreviewChecks();
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
        foreach (string id in new[] { "conversion.text", "formula.references" })
        using (var form = DialogPreviewFactory.CreateProgramDefault(id))
        using (var parking = new Panel { Size = form.Size })
        {
            Size client = form.ClientSize; form.FormBorderStyle = FormBorderStyle.None; form.ClientSize = client;
            form.TopLevel = false; parking.Controls.Add(form); form.Show();
            ((CellConversionForm)form).SuppressActivation = true;
            typeof(Form).GetMethod("OnShown", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { EventArgs.Empty }); form.PerformLayout();
            if (id == "conversion.text") ((CheckBox)form.Controls.Find("conversionKind1", true)[0]).Checked = true;
            else ((RadioButton)form.Controls.Find("conversionReference2", true)[0]).Checked = true;
            using (var overlay = new PreviewSelectionOverlay(form))
            {
                overlay.RefreshSnapshot();
                using (var bitmap = new Bitmap(overlay.Width, overlay.Height)) { overlay.DrawToBitmap(bitmap, new Rectangle(Point.Empty, bitmap.Size)); bitmap.Save(Path.Combine(folder, id + ".png"), ImageFormat.Png); }
            }
            var range = form.Controls.Find("conversionRange", true)[0];
            Check(range.Width >= 200 && range.Height >= 18 && range.Text == (id == "conversion.text" ? "J1:J5" : "G11:G14,G16:G18"), "usable range editor bounds for " + id);
            Check(form.Controls.Find("conversionApply", true).Length == 1, "real preview controls for " + id);
        }
    }
    private static IEnumerable<Control> Descendants(Control root)
    { foreach (Control item in root.Controls) { yield return item; foreach (Control child in Descendants(item)) yield return child; } }
    private static void LayoutChecks()
    {
        foreach (string id in new[] { "conversion.text", "formula.references" })
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
                    var bar = f.Apply.Parent; var root = bar.Parent;
                    int buttonBottom = root.Top + bar.Top + Math.Max(f.Apply.Bottom, f.Cancel.Bottom);
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
        foreach (string id in new[] { "conversion.text", "formula.references" })
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
        try { InteractionChecks(); LayoutChecks(); if (args.Length == 2 && args[0] == "--render") Render(args[1]); Console.WriteLine("PASS ConversionInteractionProbe: " + checks + " assertions; hidden real controls, STA previews, scaled layouts, picker restore, busy and shutdown; no host connection"); return 0; }
        catch (Exception error) { Console.WriteLine("FAIL ConversionInteractionProbe: " + error); return 1; }
        finally { ConversionInteraction.Shutdown(); ConversionWindowThread.CreatedForProbe = null; ConversionInteraction.ResultForProbe = null; }
    }
}
