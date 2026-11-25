using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

// Reflection-only offline UI validation. Never connects to WPS or runs installation.
// Usage: UiPreviewProbe.exe addin.dll [installer.exe] [output-directory]
namespace WpsUiChecks
{
    public class Coll
    {
        public List<object> Values = new List<object>();
        public int Count { get { return Values.Count; } }
        public object this[int i] { get { return Values[i - 1]; } }
    }
    public class Code { public string Text; public int Start; }
    public class Field
    {
        public int Type = 56;
        public Code Code;
        public Field(string text, int start) { Code = new Code { Text = text, Start = start }; }
    }
    public class Range
    {
        public Coll Fields = new Coll();
        public Coll InlineShapes = new Coll();
        public object NextStoryRange { get { return null; } }
    }
    public class Document
    {
        public string Name = "2026年度财务报告.docx";
        public Coll StoryRanges = new Coll();
        public Coll Shapes = new Coll();
        public Range Main = new Range();
        public Document() { StoryRanges.Values.Add(Main); }
    }
    public class WordApp { public Document ActiveDocument; public WordApp(Document d) { ActiveDocument = d; } }
    public class Workbook
    {
        public string Name = "2026年度资金测算.xlsx";
        public string FullName = @"E:\项目资料\2026年度\资金分析\2026年度资金测算.xlsx";
    }
    public class Sheet { public string Name = "资金测算"; }
    public class Cell
    {
        public string Formula = "='资产负债表'!$D$18+SUM(B3:B8)+'E:\\项目资料\\2026年度\\[期初数据.xlsx]余额表'!F32+$C$6";
        public string Address(bool row, bool column) { return "H12"; }
    }
    public class ExcelApp
    {
        public Workbook ActiveWorkbook = new Workbook();
        public Sheet ActiveSheet = new Sheet();
        public Cell ActiveCell = new Cell();
        public Coll Workbooks = new Coll();
    }

    public static class UiPreviewProbe
    {
        private static int failed, passed, skipped;
        private static readonly BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
        private static readonly BindingFlags All = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr window, int attribute, out int value, int size);
        [DllImport("dwmapi.dll")]
        private static extern int DwmFlush();
        [DllImport("dwmapi.dll")]
        private static extern int DwmGetWindowAttribute(IntPtr window, int attribute, out NativeRect value, int size);
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr window, int attribute, ref int value, int size);
        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr window, int flags);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool GetMonitorInfo(IntPtr monitor, ref NativeMonitor info);
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point point);
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, int flags);
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect { internal int Left, Top, Right, Bottom; internal Rectangle Bounds { get { return Rectangle.FromLTRB(Left, Top, Right, Bottom); } } }
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeMonitor { internal int Size; internal NativeRect Monitor, Work; internal int Flags; }

        [STAThread]
        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try
            {
                Assembly addin = Assembly.LoadFrom(Path.GetFullPath(args[0]));
                string ns = Path.GetFileNameWithoutExtension(args[0]);
                string output = args.Length > 2 ? Path.GetFullPath(args[2]) : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui-preview");
                Directory.CreateDirectory(output);
                foreach (string pattern in new string[] { "Agent-README-*.png", "Agent-使用说明-*.png" })
                    foreach (string retired in WpsShared.SourcePaths.GetFiles(output, pattern)) File.Delete(retired);
                addin.GetType(ns+".DialogSizeStore").GetField("RootOverride", All | BindingFlags.Static).SetValue(null, Path.Combine(output,"preferences"));
                addin.GetType(ns+".DialogSizing").GetMethod("Initialize",All | BindingFlags.Static).Invoke(null,null);
                bool word = ns == "WordUpdater";
                foreach (float factor in new float[] { 1f, 1.25f, 1.5f, 2f })
                {
                    foreach(bool minimum in new[]{false,true})
                    using(Form sizes=(Form)Activator.CreateInstance(addin.GetType(ns+".DialogSizesForm"),All,null,new object[0],null))
                    {
                        sizes.GetType().GetMethod("ExpandAll",All).Invoke(sizes,null);
                        Render(sizes,factor,minimum,Path.Combine(output,minimum?"交互弹窗设置-最小":"交互弹窗设置"));
                        CheckLayout(sizes,"dialog size settings "+minimum+" "+factor);
                        Check("all catalog rows visible via scroll "+factor,((DataGridView)Get(sizes,"Grid")).Rows.Count>(word?7:18));
                        CheckDialogSizeTable(sizes,factor,minimum);
                    }
                    foreach (bool guide in new bool[] { false, true })
                    foreach (bool minimum in new bool[] { false, true })
                    using (Form reader = (Form)Activator.CreateInstance(addin.GetType(ns + ".AgentDocumentForm"), All, null, new object[] { guide }, null))
                    {
                        Render(reader, factor, minimum, Path.Combine(output, guide ? (minimum ? "Agent-用户预览-最小" : "Agent-用户预览-概览") : (minimum ? "Agent-开发预览-最小" : "Agent-开发预览-概览")));
                        CheckLayout(reader, "Agent reader " + guide + " " + minimum + " " + factor);
                        RichTextBox content = null;
                        foreach (Control control in Descendants(reader)) if (control is RichTextBox) content = (RichTextBox)control;
                        Check("Agent reader readable body " + factor, content != null && content.ReadOnly && content.Width >= 400 * factor && content.Height >= 220 * factor);
                        Check("Agent reader focus and Esc " + factor, content.TabStop && reader.CancelButton != null);
                        Check("Agent reader Chinese and no marker " + factor, content.Text.Contains("加载项") && !content.Text.Contains("<!--"));
                        CheckOverviewFormatting(reader, content, factor);
                        content.Select(0, 0); content.ScrollToCaret();
                    }
                    using (Form form = CreateMain(addin, ns, word, false))
                    {
                        Render(form, factor, false, Path.Combine(output, word ? "查看链接" : "追踪引用"));
                        CheckLayout(form, "main " + factor);
                        DataGridView grid = Grid(form, word ? "gridSrc" : "_grid");
                        Check("populated rows " + factor, grid.Rows.Count > 0);
                        Check("header stays single line " + factor, grid.ColumnHeadersDefaultCellStyle.WrapMode == DataGridViewTriState.False);
                        Check("horizontal overflow uses scrollbar " + factor, grid.ScrollBars == ScrollBars.Both);
                        Check("last column fills available width " + factor, grid.Columns[grid.Columns.Count - 1].AutoSizeMode == DataGridViewAutoSizeColumnMode.Fill);
                    }
                    using (Form form = CreateMain(addin, ns, word, true))
                    {
                        Render(form, factor, true, Path.Combine(output, word ? "查看链接-空列表" : "追踪引用-空列表"));
                        CheckLayout(form, "empty/minimum " + factor);
                        Check("empty state text " + factor, ((Label)Get(form, word ? "lblEmpty" : "_lblEmpty")).Text.Length > 0);
                    }
                    using (Form dialog = CreateDialog(addin, ns, true))
                    {
                        Render(dialog, factor, true, Path.Combine(output, "确认提示"));
                        CheckLayout(dialog, "confirm " + factor);
                        Check("no Enter action on confirm " + factor, dialog.AcceptButton == null);
                        Check("confirmation defaults to cancel " + factor, dialog.ActiveControl == (Control)Get(dialog, "Negative"));
                        Check("Esc cancels without confirming " + factor, (DialogResult)dialog.GetType().GetProperty("DismissResult", Private).GetValue(dialog, null) == DialogResult.No);
                        InvokeBase(dialog, "OnKeyDown", new KeyEventArgs(Keys.Escape));
                        Check("Esc result is No " + factor, dialog.DialogResult == DialogResult.No);
                    }
                    if (!word)
                    {
                        foreach (string mode in new string[] { "IFError", "Round", "Add", "Subtract", "Multiply", "Divide" })
                        using (Form parameters = CreateFormulaOptions(addin, ns, mode))
                        {
                            ((TextBox)Get(parameters, "ValueInput")).Text = mode == "IFError" ? "错误" : "10";
                            Render(parameters, factor, true, Path.Combine(output, "一键公式-自定义" + mode));
                            CheckLayout(parameters, "formula parameters " + mode + " " + factor);
                            CheckFormulaSpacing(parameters, mode, factor);
                            Check("no implicit Enter formula commit " + mode, parameters.AcceptButton == null);
                            Check("parameter cancel is keyboard accessible " + mode, parameters.CancelButton != null);
                            Check("parameter preview is copyable " + mode, ((TextBox)Get(parameters, "Preview")).ReadOnly);
                            TextBox preview = (TextBox)Get(parameters, "Preview");
                            Check("parameter preview has a visible text row " + mode + " " + factor,
                                preview.ClientSize.Height >= preview.Font.Height + 4);
                            Check("parameter editor text is visible " + mode + " " + factor,
                                ((TextBox)Get(parameters, "ValueInput")).ClientSize.Height >= ((TextBox)Get(parameters, "ValueInput")).Font.Height);
                            ((TextBox)Get(parameters, "ValueInput")).Text = mode == "IFError" ? "A1,B1" : mode == "Round" ? "2.5" : "NaN";
                            if (mode == "IFError") ((ComboBox)Get(parameters, "ValueType")).SelectedIndex = 2;
                            parameters.PerformLayout();
                            CheckLayout(parameters, "invalid formula parameters " + mode + " " + factor);
                            CheckFormulaSpacing(parameters, mode, factor);
                            Check("validation is measured and cannot commit " + mode + " " + factor,
                                ((Label)Get(parameters, "Validation")).Visible && !((Button)Get(parameters, "Apply")).Enabled);
                            parameters.GetType().GetMethod("OnKeyDown", Private).Invoke(parameters, new object[] { new KeyEventArgs(Keys.Escape) });
                            Check("parameter Esc returns Cancel " + mode, parameters.DialogResult == DialogResult.Cancel && parameters.GetType().GetProperty("Request", Private).GetValue(parameters, null) == null);
                        }
                        foreach(string feature in new[]{"OrganizeFilesForm","RenameFilesForm"})
                        foreach(bool narrow in new[]{false,true})
                        using(Form files=(Form)Activator.CreateInstance(addin.GetType(ns+"."+feature),All,null,feature=="RenameFilesForm"?new object[]{null}:new object[0],null))
                        {
                            ((TextBox)Get(files,"FolderPath")).Text="C:\\隔离演示文件\\测试目录";
                            ((RichTextBox)Get(files,"LogView")).AppendText("文件操作预览；此窗口不连接宿主，不执行实际文件处理。\r\n完整路径、进度和实际结果可复制。");
                            if(narrow){files.MinimumSize=Size.Empty;files.ClientSize=new Size(550,480);}
                            Render(files,factor,false,Path.Combine(output,feature+(narrow?"-窄窗口":"")));
                            CheckLayout(files,"file tools "+feature+" "+narrow+" "+factor);
                            Check("file controls in window "+feature+" "+factor,files.ClientRectangle.Contains(files.RectangleToClient(((RichTextBox)Get(files,"LogView")).RectangleToScreen(((RichTextBox)Get(files,"LogView")).ClientRectangle))));
                            CheckFileTools(files,feature,factor,narrow);
                        }
                        using(Form filter = CreateColumnFilter(addin,ns))
                        {
                            Render(filter,factor,true,Path.Combine(output,"按列筛选"));
                            CheckLayout(filter,"column filtering "+factor);
                            DataGridView grid=(DataGridView)Get(filter,"Grid");
                            Check("column filter has checkboxes "+factor,grid.Columns[0] is DataGridViewCheckBoxColumn && grid.Rows.Count==30);
                            Check("column header full hover text "+factor,grid.Rows[0].Cells[2].ToolTipText.Length>80);
                            Check("no implicit Enter column commit "+factor,filter.AcceptButton==null && filter.CancelButton!=null);
                            grid.Rows[0].Cells[0].Value=false;
                            TextBox search = (TextBox)Get(filter, "Search");
                            search.Text = "数据";
                            Check("keyword finds both related titles " + factor, grid.Rows[4].Visible && grid.Rows[5].Visible && !grid.Rows[0].Visible);
                            Check("search preserves unmatched choices " + factor, !Convert.ToBoolean(grid.Rows[0].Cells[0].Value));
                            ((Button)Get(filter,"SelectAll")).PerformClick();
                            Check("visible select all excludes hidden rows " + factor, !Convert.ToBoolean(grid.Rows[0].Cells[0].Value));
                            search.Clear();
                            Check("clear search restores all rows " + factor, grid.Rows.GetRowCount(DataGridViewElementStates.Visible) == 30);
                            ((Button)Get(filter,"Reset")).PerformClick(); Check("restore original checks "+factor,Convert.ToBoolean(grid.Rows[0].Cells[0].Value));
                            ((Button)Get(filter,"SelectAll")).PerformClick(); Check("select all checks "+factor,Convert.ToBoolean(grid.Rows[1].Cells[0].Value));
                            filter.GetType().GetMethod("OnKeyDown",Private).Invoke(filter,new object[]{new KeyEventArgs(Keys.Escape)});
                            Check("column Esc closes without apply "+factor,filter.IsDisposed || !filter.Visible);
                        }
                        using (Form assistant = (Form)Activator.CreateInstance(addin.GetType(ns + ".LinkAssistantForm"), All, null, new object[] { null }, null))
                        {
                            FillAssistantLog(assistant);
                            Render(assistant, factor, false, Path.Combine(output, "链接助手"));
                            CheckLayout(assistant, "link assistant " + factor);
                            Check("link assistant starts without implicit commit " + factor, !((Button)Get(assistant, "ExecuteButton")).Enabled && assistant.AcceptButton == null && assistant.CancelButton == null);
                            Check("link assistant path is readable " + factor, ((TextBox)Get(assistant, "LocationBox")).ClientSize.Height >= ((TextBox)Get(assistant, "LocationBox")).Font.Height);
                            CheckLinkAssistant(assistant, factor, true);
                            foreach(string mode in new[]{"Other","Folder","Current"})
                            {((RadioButton)Get(assistant,mode)).Checked=true;Application.DoEvents();CheckLinkAssistant(assistant,factor,true);}
                        }
                        using (Form assistant = (Form)Activator.CreateInstance(addin.GetType(ns + ".LinkAssistantForm"), All, null, new object[] { null }, null))
                        {
                            ((TextBox)Get(assistant, "LocationBox")).Text = @"E:\项目资料\" + new string('长', 120) + "\\链接来源.xlsx";
                            ((Label)Get(assistant, "Status")).Text = "已更新 128 条链接至「外部链接」，填写新链接后点击执行替换，程序自动保存。已打开目标保持打开，处理结果会写入清单。";
                            FillAssistantLog(assistant);
                            Render(assistant, factor, true, Path.Combine(output, "链接助手-最小窗口"));
                            CheckLayout(assistant, "link assistant minimum " + factor);
                            CheckLinkAssistant(assistant, factor, false);
                            ((RichTextBox)Get(assistant, "LogView")).Focus();
                            Check("link assistant guidance accepts keyboard focus " + factor, assistant.ActiveControl != null);
                            RichTextBox journal = (RichTextBox)Get(assistant, "LogView"); journal.Focus(); journal.SelectAll();
                            Check("journal can receive focus and select complete copyable text " + factor, journal.Focused && journal.SelectedText == journal.Text && journal.TextLength > 0);
                            ((Button)Get(assistant, "CopyLog")).Focus();
                            Check("copy journal button receives keyboard focus " + factor, ((Button)Get(assistant, "CopyLog")).Focused);
                            assistant.GetType().GetMethod("ProcessDialogKey", Private).Invoke(assistant, new object[] { Keys.Escape });
                            Check("link assistant Esc closes without executing " + factor, assistant.IsDisposed || !assistant.Visible);
                        }
                        // A smaller work area can reduce the usual minimum. Verify
                        // the fallback actually wraps buttons instead of cropping them.
                        using (Form assistant = (Form)Activator.CreateInstance(addin.GetType(ns + ".LinkAssistantForm"), All, null, new object[] { null }, null))
                        {
                            assistant.MinimumSize = Size.Empty; assistant.ClientSize = new Size(360, 700);
                            FillAssistantLog(assistant);
                            Render(assistant, factor, false, Path.Combine(output, "链接助手-窄窗口"));
                            CheckLayout(assistant, "link assistant narrow " + factor);
                            CheckLinkAssistant(assistant, factor, false);
                            Control actions = (Control)Get(assistant, "actions");
                            Check("narrow footer keeps both actions visible " + factor, actions.Controls.Cast<Control>().All(control => control.Visible));
                        }
                        using (Form form = CreateDependent(addin, ns))
                        {
                            Render(form, factor, false, Path.Combine(output, "追踪从属"));
                            CheckLayout(form, "dependent " + factor);
                        }
                        using (Form form = CreateDependent(addin, ns, true))
                        {
                            Render(form, factor, true, Path.Combine(output, "从属追踪-扫描中"));
                            CheckLayout(form, "scanning " + factor);
                            ((Button)Get(form, "_btnCancel")).PerformClick();
                            Check("cancel sets request without COM " + factor, (bool)Get(form, "_cancelRequested"));
                        }
                    }
                }
                using (Form close = CreateDialog(addin, ns, true))
                {
                    InvokeBase(close, "OnFormClosing", new FormClosingEventArgs(CloseReason.UserClosing, false));
                    Check("title-bar close cancels confirmation", close.DialogResult == DialogResult.No);
                }
                using (Form result = CreateDialog(addin, ns, false))
                {
                    Render(result, 1f, false, Path.Combine(output, "结果提示"));
                    Check("result text remains copyable and scrollable", FindTextBox(result) != null && FindTextBox(result).ReadOnly && FindTextBox(result).Multiline && FindTextBox(result).ScrollBars == ScrollBars.Vertical);
                }
                using (Form confirmation = CreateDialog(addin, ns, true)) Render(confirmation, 1f, false, Path.Combine(output, "确认提示-完整"));
                // DrawToBitmap above validates client layouts. These separate captures include
                // the real desktop-composed native caption, border, and title icon.
                using (Form form = CreateMain(addin, ns, word, false))
                    RenderNativeFrame(form, ns, Path.Combine(output, word ? "查看链接-原生标题栏" : "追踪引用-原生标题栏"), true);
                if (!word) using (Form form = CreateDependent(addin, ns))
                    RenderNativeFrame(form, ns, Path.Combine(output, "追踪从属-原生标题栏"), false);
                using (Form form = CreateDialog(addin, ns, true))
                    RenderNativeFrame(form, ns, Path.Combine(output, "确认提示-原生标题栏"), false);
                using (Form form = CreateDialog(addin, ns, false))
                    RenderNativeFrame(form, ns, Path.Combine(output, "结果提示-原生标题栏"), false);
                if (!word) foreach (string mode in new string[] { "IFError", "Round", "Add", "Subtract", "Multiply", "Divide" })
                using (Form parameters = CreateFormulaOptions(addin, ns, mode))
                {
                    ((TextBox)Get(parameters, "ValueInput")).Text = mode == "IFError" ? "错误" : "10";
                    RenderNativeFrame(parameters, ns, Path.Combine(output, "一键公式-自定义" + mode + "-原生标题栏"), false);
                }
                if (!word) using(Form filter=CreateColumnFilter(addin,ns)) RenderNativeFrame(filter,ns,Path.Combine(output,"按列筛选-原生标题栏"),false);
                if (!word) using(Form assistant=(Form)Activator.CreateInstance(addin.GetType(ns+".LinkAssistantForm"),All,null,new object[]{null},null))
                    RenderNativeFrame(assistant,ns,Path.Combine(output,"链接助手-原生标题栏"),false);
                if (!word) foreach(string feature in new[]{"OrganizeFilesForm","RenameFilesForm"})
                using(Form form=(Form)Activator.CreateInstance(addin.GetType(ns+"."+feature),All,null,feature=="RenameFilesForm"?new object[]{null}:new object[0],null))
                    RenderNativeFrame(form,ns,Path.Combine(output,feature+"-原生标题栏"),false);
                RenderIcons(addin, ns, output);
                RenderRibbonLayout(addin, ns, output);
                RenderAgentLayout(addin, ns, output);
                if (args.Length > 1 && WpsShared.SourcePaths.Exists(args[1]))
                {
                    Assembly setup = Assembly.LoadFrom(Path.GetFullPath(args[1]));
                    string version = addin.GetName().Version.ToString();
                    Check("installer assembly version matches addin", setup.GetName().Version.ToString() == version);
                    System.Diagnostics.FileVersionInfo setupVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(Path.GetFullPath(args[1]));
                    Check("installer file version matches addin", setupVersion.FileVersion == version && setupVersion.ProductVersion == version);
                    string setupNs = word ? "WordUpdaterSetup" : "Excel_IterateSetup";
                    Check("installer visible version matches addin", ((string)setup.GetType(setupNs + ".InstallerText").GetField("Copyright", BindingFlags.Static | BindingFlags.NonPublic).GetRawConstantValue()).Contains("Version：" + version));
                    Type main = setup.GetType(setupNs + ".MainForm");
                    if (main == null) main = setup.GetType("ExcelIterateSetup.MainForm");
                    Check("installer form can be reflected", main != null);
                    if (main != null) foreach (float factor in new float[] { 1f, 1.25f, 1.5f, 2f })
                    {
                        using (Form form = (Form)Activator.CreateInstance(main, true))
                        {
                            Render(form, factor, true, Path.Combine(output, "安装器"));
                            CheckLayout(form, "installer " + factor);
                        }
                    }
                    if (main != null) using (Form full = (Form)Activator.CreateInstance(main, true))
                        Render(full, 1f, false, Path.Combine(output, "安装器-默认尺寸"));
                    if (main != null) using (Form form = (Form)Activator.CreateInstance(main, true))
                        RenderNativeFrame(form, ns, Path.Combine(output, "安装器-原生标题栏"), true);
                    // Render the actual shared prerequisite guidance without invoking installation.
                    string guidance = (string)setup.GetType("WpsShared.RuntimeLocator")
                        .GetField("MissingMessage", BindingFlags.Static | BindingFlags.NonPublic).GetRawConstantValue();
                    foreach (float factor in new float[] { 1f, 1.25f, 1.5f, 2f })
                    {
                        using (Form dialog = (Form)Activator.CreateInstance(setup.GetType(ns + ".UiMessageForm"), All, null,
                            new object[] { guidance, "安装", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1 }, null))
                        {
                            Render(dialog, factor, false, Path.Combine(output, "缺少VSTO运行库"));
                            CheckLayout(dialog, "runtime guidance " + factor);
                            TextBox body = FindTextBox(dialog);
                            Check("complete prerequisite guidance remains scrollable and copyable " + factor,
                                body != null && body.Text == guidance && body.ReadOnly && body.Multiline
                                && body.ScrollBars == ScrollBars.Vertical);
                        }
                    }
                }
                Console.WriteLine("UI checks: PASS=" + passed + " FAIL=" + failed + " SKIP=" + skipped);
                return failed == 0 ? 0 : 1;
            }
            catch (Exception ex) { Console.WriteLine(ex); return 2; }
        }

        private static Form CreateMain(Assembly assembly, string ns, bool word, bool empty)
        {
            Form form;
            if (word)
            {
                Document doc = new Document();
                if (!empty)
                {
                    for (int i = 1; i <= 6; i++) doc.Main.Fields.Values.Add(new Field(
                        " LINK Excel.Sheet.12 \"E:\\项目资料\\2026年度\\财务报告编制\\基础数据\\年度财务数据汇总.xlsx\" \"资产负债表!R" + i + "C4\" \\t ", i * 20));
                    for (int i = 1; i <= 3; i++) doc.Main.Fields.Values.Add(new Field(
                        " LINK Excel.Sheet.12 \"E:\\项目资料\\2026年度\\现金流量测算.xlsx\" \"经营分析!R" + i + "C2\" \\t ", 200 + i * 20));
                }
                form = (Form)Activator.CreateInstance(assembly.GetType(ns + ".LinkViewForm"), All, null, new object[] { new WordApp(doc) }, null);
            }
            else
            {
                ExcelApp app = new ExcelApp();
                if (empty) app.ActiveCell.Formula = "42";
                form = (Form)Activator.CreateInstance(assembly.GetType(ns + ".TrackerForm"), All, null,
                    new object[] { app, Enum.Parse(assembly.GetType(ns + ".TrackerMode"), "Precedents") }, null);
            }
            InvokeBase(form, "OnLoad", EventArgs.Empty);
            return form;
        }

        private static Form CreateDependent(Assembly assembly, string ns, bool scanning = false)
        {
            Form form = (Form)Activator.CreateInstance(assembly.GetType(ns + ".TrackerForm"), All, null,
                new object[] { new ExcelApp(), Enum.Parse(assembly.GetType(ns + ".TrackerMode"), "Dependents") }, null);
            if (scanning)
            {
                ((Label)Get(form, "_lblSummary")).Text = "正在扫描：工作簿 2/4 · 工作表 3/6 · 现金流量表 · 已找到 5 处";
                ((ProgressBar)Get(form, "_progress")).Value = 45;
                return form;
            }
            object result = Activator.CreateInstance(assembly.GetType(ns + ".DependentResult"));
            Set(result, "Status", Enum.Parse(assembly.GetType(ns + ".TrackerStatus"), "Completed"));
            Set(result, "TargetWorkbookName", "2026年度资金测算.xlsx"); Set(result, "TargetSheetName", "资金测算"); Set(result, "TargetCellAddress", "H12");
            Set(result, "ScannedCells", 12000); Set(result, "ScannedFormulas", 1560);
            IList items = (IList)result.GetType().GetField("Items").GetValue(result);
            for (int i = 1; i <= 5; i++)
            {
                object item = Activator.CreateInstance(assembly.GetType(ns + ".DependentItem"));
                Set(item, "Index", i); Set(item, "WorkbookName", "财务数据汇总.xlsx"); Set(item, "SheetName", i < 3 ? "经营分析" : "现金流量表");
                Set(item, "CellAddressA1", "F" + (i * 7)); Set(item, "MatchFragment", "资金测算!H12");
                Set(item, "SourceFormula", "=资金测算!H12*0.8+SUM(B3:B28)"); Set(item, "WorkbookPath", @"E:\项目资料\2026年度\财务数据汇总.xlsx"); items.Add(item);
            }
            form.GetType().GetMethod("FillDependents", Private).Invoke(form, new object[] { result });
            return form;
        }

        private static void FillAssistantLog(Form form)
        {
            object journal = Get(form, "Operations");
            MethodInfo record = journal.GetType().GetMethod("Record", Private);
            for(int i=1;i<=18;i++) record.Invoke(journal,new object[]{i==18?"扫描完成":"扫描工作簿",i==18?"已扫描 3 个工作簿，共 12 条链接；工作表「外部链接」":@"E:\项目资料\年度财务报告\长路径示例\工作簿"+i+".xlsx（离线示例）"});
        }

        private static void CheckLinkAssistant(Form form, float factor, bool full)
        {
            Label status = (Label)Get(form, "Status"), scope = (Label)Get(form, "ScopeLabel");
            RichTextBox log=(RichTextBox)Get(form,"LogView");Button copy=(Button)Get(form,"CopyLog");Control card=(Control)Get(form,"logCard");
            Control scopeCard=(Control)Get(form,"scope");
            Check("journal below scope and above footer "+factor,card.Top>=scopeCard.Bottom&&card.Bottom<=status.Top);
            Check("rich journal selectable scrollable "+factor,log.ReadOnly&&log.WordWrap&&log.ScrollBars==RichTextBoxScrollBars.Vertical&&log.TabStop);
            Check("journal readable rows "+factor,log.ClientSize.Height>=log.Font.Height*3);
            if(full && form.ClientSize.Height >= (int)(630*factor)) {
                int last=log.Text.IndexOf("仅支持",StringComparison.Ordinal);
                if(last>=0)Check("all introductory lines visible initially "+factor,log.GetPositionFromCharIndex(last).Y+log.Font.Height<=log.ClientSize.Height);
            }
            string guide=(string)Get(form,"Guidance");string history=(string)Get(form,"Operations").GetType().GetProperty("Text",Private).GetValue(Get(form,"Operations"),null);
            Check("guidance merged with complete accumulated log "+factor,log.Text.Replace("\n","\r\n").Contains(guide)||log.Text.StartsWith("1. 选择范围"));
            Check("all operation records retained "+factor,log.Text.EndsWith(history)||log.Text.Replace("\n","\r\n").EndsWith(history));
            Check("instructions describe saved results and automatic rollback " + factor,
                guide.Contains("结果自动保存") && guide.Contains("失败时会尝试回退"));
            var help = (Label)Get(form, "help");
            int gap = scopeCard.ClientSize.Height - help.Parent.Top - help.Bottom;
            Check("scope card ends at its bottom padding " + factor, Math.Abs(gap - scopeCard.Padding.Bottom) <= 1);
            Color blue=Color.FromArgb(37,99,235);
            foreach(string field in new[]{"Current","Other","Folder"}){var opt=(RadioButton)Get(form,field);Check("blue bold scope "+field+factor,opt.ForeColor==blue&&opt.Font.Bold);Check("scope complete "+field+factor,opt.Width>=opt.GetPreferredSize(Size.Empty).Width);}
            int sh=TextRenderer.MeasureText(scope.Text,scope.Font).Width;Check("scope single line "+factor,scope.Width>=sh);
            Check("status blue left "+factor,status.ForeColor==blue&&status.Font.Bold&&status.Left==scopeCard.Left);
            Check("status complete "+factor,status.Height>=TextRenderer.MeasureText(status.Text,status.Font,new Size(status.Width,int.MaxValue),TextFormatFlags.WordBreak|TextFormatFlags.NoPrefix|TextFormatFlags.TextBoxControl).Height);
            Check("status actions separate "+factor,!status.Bounds.IntersectsWith(((Button)Get(form,"ScanButton")).Parent.Bounds));
            var actions = (Control)Get(form, "actions");
            Check("scope footer contains only scan and execute " + factor, actions.Controls.Count == 2 &&
                actions.Controls[0] == Get(form, "ScanButton") && actions.Controls[1] == Get(form, "ExecuteButton"));
            foreach(Control control in new Control[]{scope,status,(Control)Get(form,"options"),actions,card,log,copy})Check("assistant inside parent "+control.GetType().Name+factor,control.Left>=0&&control.Top>=0&&control.Right<=control.Parent.ClientSize.Width&&control.Bottom<=control.Parent.ClientSize.Height);
            CheckSurfaces(form,new[]{"Current","Other","Folder","ScopeLabel","LocationBox","Browse","help","logTitle","CopyLog","LogView","Status","ScanButton","ExecuteButton"},"assistant "+factor);
        }
        private static void CheckSurfaces(Form form,string[] names,string scenario)
        {
            var surfaces=names.Select(n=>(Control)Get(form,n)).Where(c=>c.Visible).ToList();
            foreach(Control control in surfaces)
            {
                Rectangle rect=form.RectangleToClient(control.RectangleToScreen(control.ClientRectangle));
                Check(scenario+" visible in form "+control.Text,form.ClientRectangle.Contains(rect));
                for(Control ancestor=control.Parent;ancestor!=null&&ancestor!=form;ancestor=ancestor.Parent)
                    Check(scenario+" not clipped by ancestor "+control.GetType().Name,ancestor.ClientRectangle.Contains(ancestor.RectangleToClient(control.RectangleToScreen(control.ClientRectangle))));
                foreach(Control other in surfaces.Where(c=>c!=control))
                    Check(scenario+" no surface occlusion "+control.GetType().Name+" / "+other.GetType().Name,!rect.IntersectsWith(form.RectangleToClient(other.RectangleToScreen(other.ClientRectangle))));
            }
        }
        private static void CheckFileTools(Form form,string feature,float factor,bool narrow)
        {
            RichTextBox log=(RichTextBox)Get(form,"LogView");TextBox folder=(TextBox)Get(form,"FolderPath");
            Check("compact default log "+feature+factor,narrow||log.Height<=230*factor&&log.Height>=100*factor);
            Check("wrapped copyable file log "+feature+factor,log.ReadOnly&&log.WordWrap&&log.ScrollBars==RichTextBoxScrollBars.Vertical);
            CheckSurfaces(form,new[]{"FolderPath","Browse","LogView","Stop","Confirm"},feature+factor);
            if(feature=="RenameFilesForm")
            {
                Check("numbered rename guidance "+factor,log.Lines.Length>=2&&log.Lines[0].StartsWith("1. ")&&log.Lines[1].StartsWith("2. "));
                log.Select(0,5);Check("rename introduction bold "+factor,log.SelectionFont!=null&&log.SelectionFont.Bold);log.Select(log.TextLength,0);
                Rectangle body=form.RectangleToClient(log.RectangleToScreen(log.ClientRectangle));Rectangle path=form.RectangleToClient(folder.RectangleToScreen(folder.ClientRectangle));
                Check("rename path without blank block "+factor,body.Top-path.Bottom<45*factor);
            }
            else
            {
                var title=(Label)Get(form,"CustomTitle");var types=(FlowLayoutPanel)Get(form,"Types");
                Check("blue custom title "+factor,title.Text=="自定义文件批量整理选项"&&title.Font.Bold&&title.ForeColor==Color.FromArgb(37,99,235));
                Check("single row extensions "+factor,!types.WrapContents&&types.Controls.Cast<Control>().Select(c=>c.Top).Distinct().Count()==1);
                Check("extension height content-sized "+factor,types.Height<=types.Controls.Cast<Control>().Max(c=>c.Height+c.Margin.Vertical)+SystemInformation.HorizontalScrollBarHeight+2);
                // Many types must retain one row and expose horizontal scrolling.
                string[] original=types.Controls.Cast<Control>().Select(c=>(string)c.Tag).ToArray();
                form.GetType().GetMethod("SetTypes",Private).Invoke(form,new object[]{Enumerable.Range(0,40).Select(i=>"extension"+i)});types.PerformLayout();
                Check("many types scroll one row "+factor,types.HorizontalScroll.Visible&&!types.WrapContents&&types.Controls.Cast<Control>().Select(c=>c.Top).Distinct().Count()==1);
                form.GetType().GetMethod("SetTypes",Private).Invoke(form,new object[]{original});types.AutoScrollPosition=Point.Empty;types.PerformLayout();
                form.GetType().GetMethod("SetTypes",Private).Invoke(form,new object[]{new string[0]});Check("no empty extension block "+factor,!types.Visible);
                form.GetType().GetMethod("SetTypes",Private).Invoke(form,new object[]{original});types.PerformLayout();
            }
        }

        private static void CheckFormulaSpacing(Form form, string mode, float factor)
        {
            TextBox input = (TextBox)Get(form, "ValueInput"), preview = (TextBox)Get(form, "Preview");
            Label hint = (Label)Get(form, "Hint"), validation = (Label)Get(form, "Validation");
            ComboBox type = (ComboBox)Get(form, "ValueType");
            int gap;
            using (Graphics g = form.CreateGraphics()) gap = (int)Math.Round(8 * g.DpiX / 96f * form.Font.SizeInPoints / 10f);
            Check("helper gap is eight logical pixels " + mode + " " + factor, Math.Abs(hint.Top - input.Bottom - gap) <= 1);
            if (mode != "Round")
                Check("editor is centered between type and helper " + factor, Math.Abs(input.Top - type.Bottom - gap) <= 1);
            else Check("ROUND has no redundant type selector " + factor, !type.Visible);
            foreach (Label label in new Label[] { hint, validation })
                if (label.Visible) Check("complete measured helper " + label.Text + " " + factor,
                    label.Height >= TextRenderer.MeasureText(label.Text, label.Font, new Size(label.Width, int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix).Height);
            Check("preview has three readable lines " + mode + " " + factor, preview.ClientSize.Height >= preview.Font.Height * 3);
        }

        private static Form CreateColumnFilter(Assembly assembly,string ns)
        {
            Type choice=assembly.GetType(ns+".ColumnChoice");
            IList items=(IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(choice));
            for(int i=1;i<=30;i++)
            {
                object item=Activator.CreateInstance(choice,true);
                Set(item,"Number",i); Set(item,"Letter",(string)assembly.GetType(ns+".ComFormulaHost").GetMethod("Address",BindingFlags.Static | BindingFlags.NonPublic).Invoke(null,new object[]{1,i}).ToString().TrimEnd('1')); Set(item,"Header",i==1?new string('长',120):i==3?"（空标题）":i==4?"日期 [合并 $D$1:$E$1]":"数据列 "+i);
                Set(item,"OriginalHidden",i==2); Set(item,"Visible",i!=2); items.Add(item);
            }
            return (Form)Activator.CreateInstance(assembly.GetType(ns+".ColumnFilterForm"),All,null,new object[]{items},null);
        }

        private static Form CreateFormulaOptions(Assembly assembly, string ns, string mode)
        {
            object operation = Enum.Parse(assembly.GetType(ns + ".FormulaOperation"), mode);
            return (Form)Activator.CreateInstance(assembly.GetType(ns + ".FormulaOptionsForm"), BindingFlags.Instance | BindingFlags.NonPublic,
                null, new object[] { operation }, null);
        }

        private static Form CreateDialog(Assembly assembly, string ns, bool confirm)
        {
            string text = confirm ? "即将中断本文档中的全部 9 条链接。\r\n\r\n中断后，链接会变成静态内容，不再随源文件变化。\r\n\r\n请确认已保存文档。" : "更新完成：成功 9 条，失败 0 条。\r\n\r\n检测到 3 处内容变化，已临时标黄。\r\n\r\n点击确定后开始标黄计时。";
            return (Form)Activator.CreateInstance(assembly.GetType(ns + ".UiMessageForm"), All, null,
                new object[] { text, confirm ? "中断所有链接" : "更新链接", confirm ? MessageBoxButtons.YesNo : MessageBoxButtons.OK,
                    confirm ? MessageBoxIcon.Warning : MessageBoxIcon.Information, MessageBoxDefaultButton.Button2 }, null);
        }

        private static string FirstBodyLine(string markdown)
        {
            foreach (string line in markdown.Replace("\r\n", "\n").Split('\n'))
            {
                string text = line.Trim();
                if (text.Length == 0 || text.StartsWith("<!--") || text.StartsWith("```") || Regex.IsMatch(text, "^#{1,6}\\s")) continue;
                return Regex.Replace(text, "\\[([^\\]]+)\\]\\(([^)]+)\\)", "$1（$2）").Replace("**", "").Replace("`", "");
            }
            throw new InvalidOperationException("Overview has no body text");
        }

        private static void CheckOverviewFormatting(Form reader, RichTextBox content, float factor)
        {
            string overview = (string)reader.GetType().GetProperty("CurrentSource", All).GetValue(reader, null);
            string body = FirstBodyLine(overview); int start = content.Text.IndexOf(body, StringComparison.Ordinal);
            Check("Agent actual overview body is rendered " + factor, start >= 0);
            if (start < 0) return;
            content.Select(start, Math.Min(body.Length, 4));
            Check("Agent actual body at scale " + factor, content.SelectionFont != null && !content.SelectionFont.Bold && Math.Abs(content.SelectionFont.Size - 10f * factor) < 0.1f && content.SelectionColor == Color.FromArgb(36,50,71));
            // A hidden sample checks heading styling even when the overview is plain prose.
            using (Form sample = (Form)Activator.CreateInstance(reader.GetType(), All, null, new object[] { false }, null))
            {
                sample.GetType().GetField("overview", All).SetValue(sample, "# 样例标题\n样例正文");
                sample.GetType().GetMethod("SetView", All).Invoke(sample, new object[] { false });
                if (factor != 1f) ScaleFonts(sample, factor);
                RichTextBox text = (RichTextBox)Get(sample, "document");
                text.Select(text.Text.IndexOf("样例标题", StringComparison.Ordinal), 2);
                Check("Agent explicit Markdown heading at scale " + factor, text.SelectionFont != null && text.SelectionFont.Bold && Math.Abs(text.SelectionFont.Size - 13f * factor) < 0.1f && text.SelectionColor == Color.FromArgb(37,99,235));
                text.Select(text.Text.IndexOf("样例正文", StringComparison.Ordinal), 2);
                Check("Agent explicit Markdown body at scale " + factor, text.SelectionFont != null && !text.SelectionFont.Bold && Math.Abs(text.SelectionFont.Size - 10f * factor) < 0.1f && text.SelectionColor == Color.FromArgb(36,50,71));
            }
        }

        private static void Render(Form form, float factor, bool minimum, string stem)
        {
            if (minimum) form.Size = form.MinimumSize;
            // Show a fully transparent test window so DrawToBitmap includes child controls.
            // Remove Load first: dependent mode must never connect to a real WPS instance.
            RemoveLoadHandlers(form);
            form.Opacity = 0;
            form.Show();
            Application.DoEvents();
            if(minimum){form.Size=form.MinimumSize;Application.DoEvents();}
            if (factor != 1f)
            {
                form.Scale(new SizeF(factor, factor));
                ScaleFonts(form, factor);
            }
            form.PerformLayout();
            foreach (Control control in Descendants(form)) control.PerformLayout();
            using (Bitmap bitmap = new Bitmap(form.Width, form.Height))
            {
                form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
                // WinForms DrawToBitmap does not print native RichEdit text. Use
                // its print interface, preserving its actual formatted runs.
                foreach (Control child in Descendants(form))
                    if (child is RichTextBox) PrintRichText((RichTextBox)child, form, bitmap);
                bitmap.Save(stem + "-" + (int)(factor * 100) + ".png", ImageFormat.Png);
            }
        }

        private static void CheckDialogSizeTable(Form form,float factor,bool minimum)
        {
            DataGridView grid=(DataGridView)Get(form,"Grid");
            int available=grid.ClientSize.Width-SystemInformation.VerticalScrollBarWidth;
            int total=0;foreach(DataGridViewColumn column in grid.Columns)total+=column.Width;
            bool scroll=minimum||total>available;
            if(!scroll)
            {
                Rectangle header=grid.GetCellDisplayRectangle(7,-1,false);
                Check("last dialog header fully visible "+factor,header.Left>=0&&header.Right<=available);
            }
            else Check("narrow or screen-limited table keeps horizontal scrolling "+minimum+factor,grid.ScrollBars==ScrollBars.Both&&total>available);
            DataGridViewRow action=null;
            foreach(DataGridViewRow row in grid.Rows)if(Convert.ToString(row.Cells[7].Value)=="自定调整"){action=row;break;}
            Check("rounded action cell exists "+factor,action!=null&&action.Cells[7].GetType().Name=="RoundedAdjustmentCell");
            grid.CurrentCell=action.Cells[7];grid.FirstDisplayedScrollingRowIndex=action.Index;grid.HorizontalScrollingOffset=0;
            if(scroll)grid.FirstDisplayedScrollingColumnIndex=7;
            Application.DoEvents();Rectangle cell=grid.GetCellDisplayRectangle(7,action.Index,false);
            Rectangle lastHeader=grid.GetCellDisplayRectangle(7,-1,false);
            Check("last header reachable in available viewport "+minimum+factor,lastHeader.Left>=0&&lastHeader.Right<=available);
            var type=form.GetType().Assembly.GetType(form.GetType().Namespace+".RoundedAdjustmentCell");
            Rectangle pill=(Rectangle)type.GetMethod("ButtonBounds",All|BindingFlags.Static).Invoke(null,new object[]{cell,grid.DefaultCellStyle.Font,96f});
            Check("adjustment pill visible after scroll "+minimum+factor,pill.Left>=0&&pill.Right<=grid.ClientSize.Width&&pill.Top>=grid.ColumnHeadersHeight&&pill.Bottom<=grid.ClientSize.Height);
            using(var font=new Font(grid.DefaultCellStyle.Font,FontStyle.Bold))
                Check("adjustment pill has full caption "+minimum+factor,pill.Width>=TextRenderer.MeasureText("自定调整",font).Width&&pill.Height>=28*factor);
            grid.CurrentCell=grid.Rows[0].Cells[0];grid.FirstDisplayedScrollingRowIndex=0;grid.HorizontalScrollingOffset=0;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RichRect { internal int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)]
        private struct RichRange { internal int First, Last; }
        [StructLayout(LayoutKind.Sequential)]
        private struct RichFormat { internal IntPtr Hdc, Target; internal RichRect Area, Page; internal RichRange Range; }
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr FormatRichText(IntPtr handle, int message, IntPtr wParam, ref RichFormat data);
        [DllImport("user32.dll", EntryPoint = "SendMessageW")]
        private static extern IntPtr ClearRichFormat(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);
        private static void PrintRichText(RichTextBox box, Form form, Bitmap target)
        {
            using (Bitmap content = new Bitmap(box.ClientSize.Width, box.ClientSize.Height))
            using (Graphics g = Graphics.FromImage(content))
            {
                g.Clear(box.BackColor);
                IntPtr hdc = g.GetHdc();
                try
                {
                    RichRect area = new RichRect { Right = (int)Math.Round(content.Width * 1440.0 / content.HorizontalResolution), Bottom = (int)Math.Round(content.Height * 1440.0 / content.VerticalResolution) };
                    RichFormat format = new RichFormat { Hdc = hdc, Target = hdc, Area = area, Page = area, Range = new RichRange { First = 0, Last = -1 } };
                    FormatRichText(box.Handle, 0x439, new IntPtr(1), ref format);
                }
                finally { g.ReleaseHdc(hdc); ClearRichFormat(box.Handle, 0x439, IntPtr.Zero, IntPtr.Zero); }
                int ink = 0;
                for (int y = 0; y < content.Height; y += 2)
                for (int x = 0; x < content.Width; x += 2)
                    if (content.GetPixel(x, y).ToArgb() != box.BackColor.ToArgb()) ink++;
                Check("RichEdit preview contains rendered text", box.TextLength == 0 || ink > 30);
                using (Graphics output = Graphics.FromImage(target))
                {
                    Point position = box.PointToScreen(Point.Empty);
                    position.Offset(-form.Left, -form.Top); // bitmap includes native frame
                    output.DrawImageUnscaled(content, position);
                }
            }
        }

        private static void RemoveLoadHandlers(Form form)
        {
            System.ComponentModel.EventHandlerList events = (System.ComponentModel.EventHandlerList)typeof(System.ComponentModel.Component)
                .GetProperty("Events", Private).GetValue(form, null);
            FieldInfo loadKey = typeof(Form).GetField("EVENT_LOAD", BindingFlags.Static | BindingFlags.NonPublic);
            if (loadKey == null) throw new InvalidOperationException("WinForms Load event key not found; refusing to show test form.");
            object key = loadKey.GetValue(null);
            events.RemoveHandler(key, events[key]);
        }

        private static void RenderNativeFrame(Form form, string ns, string stem, bool exerciseLifecycle)
        {
            // Only these freshly constructed offline forms become briefly visible. Load is
            // removed before Show, so no real WPS scan or installer action can be started.
            RemoveLoadHandlers(form);
            Rectangle work = Screen.PrimaryScreen.WorkingArea;
            Size limit = new Size(Math.Max(1, work.Width - 48), Math.Max(1, work.Height - 48));
            form.MinimumSize = new Size(Math.Min(form.MinimumSize.Width, limit.Width), Math.Min(form.MinimumSize.Height, limit.Height));
            form.Size = new Size(Math.Min(form.Width, limit.Width), Math.Min(form.Height, limit.Height));
            form.StartPosition = FormStartPosition.Manual;
            // Keep our test frames away from floating windows near screen center.
            // Captures still require all sampled points to belong to this form.
            form.Location = new Point(work.Left + 24, work.Top + 24);
            form.Opacity = 1;
            form.TopMost = true;
            int disableTransitions = 1;
            DwmSetWindowAttribute(form.Handle, 3, ref disableTransitions, sizeof(int));
            try
            {
                form.Show();
                Application.DoEvents();
                if(form.GetType().Name=="LinkAssistantForm")CheckLinkAssistant(form,1f,false);
                if(form.GetType().Name=="OrganizeFilesForm"||form.GetType().Name=="RenameFilesForm")
                    CheckFileTools(form,form.GetType().Name,1f,false);
                CheckNativeFrame(form, ns, exerciseLifecycle);
                // Production Shown handlers center and size their windows. Apply
                // the probe-only position after those handlers have completed.
                form.Location = new Point(work.Left + 24, work.Top + 24);
                // RecreateHandle also creates a new compositor animation state.
                DwmSetWindowAttribute(form.Handle, 3, ref disableTransitions, sizeof(int));
                form.BringToFront();
                form.Activate();
                Application.DoEvents();
                // Let pending show/recreate animations finish before checking opacity.
                System.Threading.Thread.Sleep(250);
                Application.DoEvents();
                try { DwmFlush(); } catch (DllNotFoundException) { } catch (EntryPointNotFoundException) { }
                // Form.Bounds is DPI-virtualized in this .NET Framework probe. Use the
                // compositor's physical bounds and temporarily physical GDI coordinates.
                Point[] alternatePositions = new Point[] {
                    new Point(Math.Max(work.Left + 24, work.Right - form.Width - 24), work.Top + 24),
                    new Point(work.Left + 24, Math.Max(work.Top + 24, work.Bottom - form.Height - 24)),
                    new Point(Math.Max(work.Left + 24, work.Right - form.Width - 24), Math.Max(work.Top + 24, work.Bottom - form.Height - 24)) };
                IntPtr previousDpi = SetThreadDpiAwarenessContext(new IntPtr(-4));
                try
                {
                    NativeRect frame;
                    if (DwmGetWindowAttribute(form.Handle, 9, out frame, Marshal.SizeOf(typeof(NativeRect))) != 0)
                        throw new InvalidOperationException("Cannot obtain physical bounds for the probe's own window.");
                    Rectangle bounds = frame.Bounds;
                    NativeMonitor monitor = new NativeMonitor(); monitor.Size = Marshal.SizeOf(typeof(NativeMonitor));
                    if (!GetMonitorInfo(MonitorFromWindow(form.Handle, 2), ref monitor) || bounds.Width <= 0 || bounds.Height <= 0 || !monitor.Work.Bounds.Contains(bounds))
                        throw new InvalidOperationException("Refusing a clipped native-frame screenshot: " + form.Text);
                    Check("physical native frame is inside work area " + form.Text, true);
                    bool unobstructed = OwnsFramePoints(form, bounds);
                    for (int attempt = 0; !unobstructed && attempt < alternatePositions.Length; attempt++)
                    {
                        if (previousDpi != IntPtr.Zero) SetThreadDpiAwarenessContext(previousDpi);
                        form.Location = alternatePositions[attempt]; form.BringToFront();
                        Application.DoEvents(); System.Threading.Thread.Sleep(250); Application.DoEvents();
                        try { DwmFlush(); } catch (DllNotFoundException) { } catch (EntryPointNotFoundException) { }
                        SetThreadDpiAwarenessContext(new IntPtr(-4));
                        if (DwmGetWindowAttribute(form.Handle, 9, out frame, Marshal.SizeOf(typeof(NativeRect))) != 0)
                            throw new InvalidOperationException("Cannot read moved probe frame.");
                        bounds = frame.Bounds;
                        unobstructed = monitor.Work.Bounds.Contains(bounds) && OwnsFramePoints(form, bounds);
                    }
                    if (!unobstructed)
                    {
                        if (WpsShared.SourcePaths.Exists(stem + "-100.png")) File.Delete(stem + "-100.png");
                        Skip("native frame capture occluded by another window " + form.Text); return;
                    }
                    using (Bitmap bitmap = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb))
                    using (Graphics graphics = Graphics.FromImage(bitmap))
                    {
                        // Capture only our full opaque window, never Form.Bounds-scaled
                        // portions of the desktop or other applications.
                        graphics.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size, CopyPixelOperation.SourceCopy);
                        if (!SystemInformation.HighContrast && PaletteSupported(form, ns, false))
                        {
                            bool opaque = bitmap.GetPixel(bitmap.Width / 2, Math.Min(14, bitmap.Height - 1)).ToArgb() == Color.FromArgb(245, 247, 251).ToArgb();
                            Check("native caption visibly matches canvas " + form.Text, opaque);
                            if (!opaque)
                            {
                                if (WpsShared.SourcePaths.Exists(stem + "-100.png")) File.Delete(stem + "-100.png");
                                throw new InvalidOperationException("Native frame did not settle to the expected opaque caption: " + form.Text);
                            }
                        }
                        bitmap.Save(stem + "-100.png", ImageFormat.Png);
                    }
                }
                finally { if (previousDpi != IntPtr.Zero) SetThreadDpiAwarenessContext(previousDpi); }
            }
            finally { form.TopMost = false; form.Hide(); }
        }

        private static bool OwnsFramePoints(Form form, Rectangle bounds)
        {
            foreach (Point point in new Point[] {
                new Point(bounds.Left + bounds.Width / 2, bounds.Top + 14),
                new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2),
                new Point(bounds.Left + 16, bounds.Top + 16), new Point(bounds.Right - 16, bounds.Top + 16),
                new Point(bounds.Left + 16, bounds.Bottom - 16), new Point(bounds.Right - 16, bounds.Bottom - 16) })
                if (GetAncestor(WindowFromPoint(point), 2) != form.Handle) return false;
            return true;
        }

        private static void CheckNativeFrame(Form form, string ns, bool exerciseLifecycle)
        {
            string caption = form.Text;
            Check("system caption and resize border retained " + caption, form.FormBorderStyle == FormBorderStyle.Sizable && form.ControlBox);
            CheckWindowIcon(form.Icon, caption);
            bool supported = CheckCaptionPalette(form, SystemInformation.HighContrast, "initial " + caption);
            if (!exerciseLifecycle) return;

            int created = 0;
            EventHandler countCreated = delegate { created++; };
            form.HandleCreated += countCreated;
            try
            {
                typeof(Control).GetMethod("RecreateHandle", Private).Invoke(form, null);
                Application.DoEvents();
                Check("native frame survives handle recreation " + caption, created > 0 && form.IsHandleCreated && form.ControlBox);
                CheckWindowIcon(form.Icon, "recreated " + caption);
                if (supported) CheckCaptionPalette(form, SystemInformation.HighContrast, "recreated " + caption);
            }
            finally { form.HandleCreated -= countCreated; }

            if (!supported) return;
            Type chrome = form.GetType().Assembly.GetType(ns + ".UiWindowChrome");
            MethodInfo apply = chrome == null ? null : chrome.GetMethod("TryApplyPalette", BindingFlags.Static | BindingFlags.NonPublic);
            Check("native palette fallback is testable " + caption, apply != null);
            if (apply == null) return;
            try
            {
                Check("high contrast reset succeeds " + caption, (bool)apply.Invoke(null, new object[] { form.Handle, true }));
                CheckCaptionPalette(form, true, "high contrast reset " + caption);
            }
            finally
            {
                Check("native palette reapplies after reset " + caption,
                    (bool)apply.Invoke(null, new object[] { form.Handle, SystemInformation.HighContrast }));
                CheckCaptionPalette(form, SystemInformation.HighContrast, "restored " + caption);
            }
        }

        private static bool CheckCaptionPalette(Form form, bool highContrast, string scenario)
        {
            int border, caption, text;
            if (!ReadDwmColor(form.Handle, 34, out border) || !ReadDwmColor(form.Handle, 35, out caption) || !ReadDwmColor(form.Handle, 36, out text))
            {
                // These attributes are documented for Set, not for Get. E_INVALIDARG
                // on readback does not mean the window cannot use the new colors.
                if (PaletteSupported(form, form.GetType().Namespace.Replace("Setup", ""), highContrast))
                { Console.WriteLine("INFO DWM palette accepted; color readback unavailable " + scenario); return true; }
                Skip("DWM caption colors unsupported; system frame fallback " + scenario);
                return false;
            }
            Check("native border palette " + scenario, border == (highContrast ? -1 : ColorRef(Color.FromArgb(228, 233, 242))));
            Check("native caption palette " + scenario, caption == (highContrast ? -1 : ColorRef(Color.FromArgb(245, 247, 251))));
            Check("native caption text palette " + scenario, text == (highContrast ? -1 : ColorRef(Color.FromArgb(36, 50, 71))));
            return true;
        }

        private static bool PaletteSupported(Form form, string ns, bool highContrast)
        {
            MethodInfo apply = form.GetType().Assembly.GetType(ns + ".UiWindowChrome")
                .GetMethod("TryApplyPalette", BindingFlags.Static | BindingFlags.NonPublic);
            return (bool)apply.Invoke(null, new object[] { form.Handle, highContrast });
        }

        private static bool ReadDwmColor(IntPtr window, int attribute, out int value)
        {
            value = 0;
            try { return DwmGetWindowAttribute(window, attribute, out value, sizeof(int)) == 0; }
            catch (DllNotFoundException) { return false; }
            catch (EntryPointNotFoundException) { return false; }
        }

        private static int ColorRef(Color color) { return color.R | (color.G << 8) | (color.B << 16); }

        private static void CheckWindowIcon(Icon icon, string scenario)
        {
            Check("native title icon supplied " + scenario, icon != null);
            if (icon == null) return;
            byte[] data;
            using (MemoryStream stream = new MemoryStream())
            {
                icon.Save(stream);
                data = stream.ToArray();
                int count = data.Length >= 6 ? BitConverter.ToUInt16(data, 4) : 0;
                bool small = false, large = false;
                for (int i = 0; i < count && 6 + (i + 1) * 16 <= data.Length; i++)
                {
                    int offset = 6 + i * 16;
                    small |= data[offset] == 16 && data[offset + 1] == 16;
                    large |= data[offset] == 32 && data[offset + 1] == 32;
                }
                Check("native title icon contains 16/32 frames " + scenario, small && large);
            }
            foreach (int size in new int[] { 16, 32 })
            using (Bitmap image = ReadPngIconFrame(data, size))
            using (Icon legacy = new Icon(SystemIcons.Application, size, size))
            using (Bitmap old = IconBitmap(legacy, "system default " + scenario + " " + size))
            {
                Check("native icon transparent corners " + scenario + " " + size,
                    image.GetPixel(0, 0).A == 0 && image.GetPixel(image.Width - 1, 0).A == 0
                    && image.GetPixel(0, image.Height - 1).A == 0 && image.GetPixel(image.Width - 1, image.Height - 1).A == 0);
                bool different = image.Size != old.Size;
                for (int y = 0; !different && y < image.Height; y++)
                    for (int x = 0; !different && x < image.Width; x++)
                        different = image.GetPixel(x, y).ToArgb() != old.GetPixel(x, y).ToArgb();
                Check("native icon replaces old default " + scenario + " " + size, different);
            }
        }

        private static Bitmap ReadPngIconFrame(byte[] data, int size)
        {
            // .NET Framework's Icon.ToBitmap treats some PNG ICO frames as a DIB.
            // Decode the encoded PNG directly; native Windows rendering is checked separately.
            int count = data.Length >= 6 ? BitConverter.ToUInt16(data, 4) : 0;
            for (int i = 0; i < count && 6 + (i + 1) * 16 <= data.Length; i++)
            {
                int entry = 6 + i * 16;
                if (data[entry] != size || data[entry + 1] != size) continue;
                int length = BitConverter.ToInt32(data, entry + 8);
                int offset = BitConverter.ToInt32(data, entry + 12);
                if (length < 8 || offset < 6 + count * 16 || offset > data.Length - length)
                    throw new InvalidDataException("Invalid icon frame bounds: " + size);
                if (data[offset] != 137 || data[offset + 1] != 80 || data[offset + 2] != 78 || data[offset + 3] != 71)
                    throw new InvalidDataException("Expected PNG icon frame: " + size);
                using (MemoryStream png = new MemoryStream(data, offset, length))
                using (Bitmap image = new Bitmap(png)) return new Bitmap(image);
            }
            throw new InvalidDataException("Missing icon frame: " + size);
        }

        private static Bitmap IconBitmap(Icon icon, string scenario)
        {
            try { return icon.ToBitmap(); }
            catch (Exception ex) { throw new InvalidOperationException("Cannot inspect native icon: " + scenario, ex); }
        }

        private sealed class RichRun
        {
            internal int Start, Length; internal string Family; internal float Size;
            internal FontStyle Style; internal Color Color;
        }
        private static void ScaleFonts(Control root, float factor)
        {
            Dictionary<Control, Font> fonts = new Dictionary<Control, Font>();
            fonts.Add(root, root.Font);
            foreach (Control control in Descendants(root)) fonts.Add(control, control.Font);
            var richRuns = new Dictionary<RichTextBox, List<RichRun>>();
            foreach (Control control in Descendants(root))
            {
                RichTextBox rich = control as RichTextBox;
                if (rich == null) continue;
                var runs = new List<RichRun>(); int start = 0;
                // Reader formatting is constant within each authored line.
                foreach (string line in rich.Text.Split('\n'))
                {
                    int length = Math.Min(line.Length + 1, rich.TextLength - start);
                    if (length <= 0) break;
                    rich.Select(start, 1);
                    using (Font font = rich.SelectionFont)
                        runs.Add(new RichRun { Start = start, Length = length, Family = font.Name,
                            Size = font.Size, Style = font.Style, Color = rich.SelectionColor });
                    start += length;
                }
                richRuns.Add(rich, runs);
            }
            foreach (KeyValuePair<Control, Font> entry in fonts)
                entry.Key.Font = new Font(entry.Value.FontFamily, entry.Value.Size * factor, entry.Value.Style);
            foreach (KeyValuePair<RichTextBox, List<RichRun>> entry in richRuns)
            {
                foreach (RichRun run in entry.Value)
                {
                    entry.Key.Select(run.Start, run.Length);
                    using (Font font = new Font(run.Family, run.Size * factor, run.Style)) entry.Key.SelectionFont = font;
                    entry.Key.SelectionColor = run.Color;
                }
                entry.Key.Select(0, 0); entry.Key.ScrollToCaret();
            }
            foreach (Control control in Descendants(root))
            {
                if (control is DataGridView)
                {
                    DataGridView grid = (DataGridView)control;
                    grid.ColumnHeadersDefaultCellStyle.Font = new Font(grid.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f * factor, FontStyle.Bold);
                    grid.DefaultCellStyle.Font = new Font(grid.DefaultCellStyle.Font.FontFamily, 10f * factor);
                    grid.DefaultCellStyle.Padding = new Padding((int)(8 * factor), 0, (int)(8 * factor), 0);
                    grid.ColumnHeadersDefaultCellStyle.Padding = grid.DefaultCellStyle.Padding;
                    grid.ColumnHeadersHeight = (int)(38 * factor);
                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        row.Height = (int)(34 * factor);
                        if (row.DefaultCellStyle.Font != null)
                            row.DefaultCellStyle.Font = new Font(row.DefaultCellStyle.Font.FontFamily, row.DefaultCellStyle.Font.Size * factor, row.DefaultCellStyle.Font.Style);
                        foreach (DataGridViewCell cell in row.Cells)
                            if (cell.Style.Font != null)
                                cell.Style.Font = new Font(cell.Style.Font.FontFamily, cell.Style.Font.Size * factor, cell.Style.Font.Style);
                    }
                    foreach (DataGridViewColumn col in grid.Columns)
                        if (col.DefaultCellStyle.Font != null)
                            col.DefaultCellStyle.Font = new Font(col.DefaultCellStyle.Font.FontFamily, 10f * factor);
                    if(root.GetType().Name=="DialogSizesForm")root.GetType().GetMethod("FitColumns",All).Invoke(root,null);
                    else root.GetType().Assembly.GetType(root.GetType().Namespace.Replace("Setup", "") + ".Ui").GetMethod("AutoFitColumns", BindingFlags.Public | BindingFlags.Static).Invoke(null, new object[] { grid });
                }
            }
        }

        private static void CheckLayout(Form form, string scenario)
        {
            foreach (Control control in Descendants(form))
            {
                DataGridView grid = control as DataGridView;
                if (grid != null) foreach (DataGridViewColumn column in grid.Columns)
                    Check(scenario + " full header " + column.HeaderText,
                        column.Width >= TextRenderer.MeasureText(column.HeaderText, grid.ColumnHeadersDefaultCellStyle.Font).Width + grid.ColumnHeadersDefaultCellStyle.Padding.Horizontal);
                Button button = control as Button;
                if (button == null) continue;
                Size needed = TextRenderer.MeasureText(button.Text, button.Font, Size.Empty, TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
                Check(scenario + " button text " + button.Text, button.Width >= needed.Width + 8 && button.Height >= needed.Height + 6);
                Check(scenario + " button stays inside parent " + button.Text, button.Left >= 0 && button.Top >= 0 && button.Right <= button.Parent.ClientSize.Width && button.Bottom <= button.Parent.ClientSize.Height);
                foreach (Control other in button.Parent.Controls)
                    if (other is Button && other != button) Check(scenario + " no button overlap", !button.Bounds.IntersectsWith(other.Bounds));
            }
        }

        private static void RenderIcons(Assembly assembly, string ns, string output)
        {
            Type icons = assembly.GetType(ns + ".UiIcons");
            MethodInfo legacy = icons.GetMethod("Create", BindingFlags.Static | BindingFlags.Public);
            MethodInfo create = icons.GetMethod("CreateRibbon", BindingFlags.Static | BindingFlags.Public);
            MethodInfo window = icons.GetMethod("CreateWindowIcon", BindingFlags.Static | BindingFlags.Public);
            string[] kinds = { "update", "paste", "view", "table", "mail", "precedents", "dependents", "formula", "filter", "link", "directory" };
            string[] labels = { "更新链接", "链接Excel", "查看链接", "全选表格", "问题反馈", "追踪引用", "追踪从属", "公式 fx", "筛选增强", "链接助手", "生成目录" };
            using (Bitmap sheet = new Bitmap(1470, 390))
            using (Graphics g = Graphics.FromImage(sheet))
            using (Font title = new Font("Microsoft YaHei UI", 14f))
            using (Font font = new Font("Microsoft YaHei UI", 9f))
            {
                g.Clear(Color.FromArgb(245, 247, 251));
                g.DrawString("功能区图标 · 尺寸校准与图形主体对照", title, Brushes.DarkSlateGray, 20, 16);
                g.DrawString("代码渲染预览；WPS 按钮留白与文字宽度由宿主控制", font, Brushes.SlateGray, 20, 48);
                g.DrawString("原图形 / 32 画布", font, Brushes.SlateGray, 20, 112);
                g.DrawString("新图形 / 32 画布", font, Brushes.SlateGray, 20, 172);
                g.DrawString("150% 下显示比例", font, Brushes.SlateGray, 20, 250);
                g.DrawString("每列左旧 / 右新", font, Brushes.SlateGray, 20, 273);
                g.DrawString("标题栏 16 / 32", font, Brushes.SlateGray, 20, 332);
                for (int i = 0; i < kinds.Length; i++)
                {
                    Color accent = i == 4 ? Color.FromArgb(245,158,11) : i == 6 ? Color.FromArgb(22, 135, 111) : Color.FromArgb(37, 99, 235);
                    int x = 176 + i * 113;
                    g.DrawString(labels[i], font, Brushes.DarkSlateGray, x - 10, 80);
                    if (kinds[i] == "formula") g.DrawString("新增组件", font, Brushes.SlateGray, x, 112);
                    else using (Bitmap old = (Bitmap)legacy.Invoke(null, new object[] { kinds[i], accent, 32 }))
                    { g.DrawImageUnscaled(old, x + 16, 108); g.DrawImageUnscaled(old, x - 4, 242); }
                    using (Bitmap current = (Bitmap)create.Invoke(null, new object[] { kinds[i], accent, 32, 32 }))
                        g.DrawImageUnscaled(current, x + 16, 168);
                    using (Bitmap current = (Bitmap)create.Invoke(null, new object[] { kinds[i], accent, 48, 48 }))
                        g.DrawImageUnscaled(current, x + 36, 234);
                    foreach (int size in new int[] { 16, 32, 40, 48, 64 })
                    using (Bitmap icon = (Bitmap)create.Invoke(null, new object[] { kinds[i], accent, size, size }))
                    {
                        Check("Ribbon physical dimensions " + kinds[i] + " " + size, icon.Width == size && icon.Height == size);
                        Rectangle bounds = AccentBounds(icon, accent);
                        Check("Ribbon safe edge margin " + kinds[i] + " " + size, bounds.Left > 0 && bounds.Top > 0 && bounds.Right < size && bounds.Bottom < size);
                        if (size == 32) Check("Ribbon main body " + kinds[i], bounds.Width >= 24 && bounds.Height >= (kinds[i] == "mail" ? 22 : 24));
                    }
                    if (kinds[i] == "formula") { g.DrawString("仅功能区", font, Brushes.SlateGray, x, 332); continue; }
                    using (Icon caption = (Icon)window.Invoke(null, new object[] { kinds[i], accent }))
                    using (MemoryStream bytes = new MemoryStream())
                    {
                        caption.Save(bytes);
                        foreach (int size in new int[] { 16, 32 })
                        using (Bitmap image = ReadPngIconFrame(bytes.ToArray(), size))
                        using (Bitmap original = (Bitmap)icons.GetMethod("Create", BindingFlags.Static | BindingFlags.NonPublic,
                            null, new Type[] { typeof(string), typeof(Color), typeof(int), typeof(bool) }, null)
                            .Invoke(null, new object[] { kinds[i], accent, size, true }))
                        {
                            Check("caption remains " + kinds[i] + " " + size, image.Size == original.Size);
                            // Compare through the same PNG decode/clone path: GDI+ rounds
                            // premultiplied colors when cloning translucent antialiased pixels.
                            using (MemoryStream png = new MemoryStream())
                            {
                                original.Save(png, ImageFormat.Png); png.Position = 0;
                                using (Bitmap decoded = new Bitmap(png))
                                using (Bitmap normalized = new Bitmap(decoded))
                                {
                                    bool same = true;
                                    for (int y = 0; y < size; y++) for (int k = 0; k < size; k++)
                                        same &= image.GetPixel(k, y).ToArgb() == normalized.GetPixel(k, y).ToArgb();
                                    Check("caption keeps original geometry " + kinds[i] + " " + size, same);
                                }
                            }
                            g.DrawImageUnscaled(image, x + (size == 16 ? 0 : 24), 326);
                        }
                    }
                }
                sheet.Save(Path.Combine(output, "功能区图标.png"), ImageFormat.Png);
                sheet.Save(Path.Combine(output, "功能区图标-前后对照.png"), ImageFormat.Png);
            }
            if (ns == "Excel_Iterate") RenderFormulaMenu(assembly, ns, output);
            MethodInfo read = assembly.GetType(ns + ".RibbonImages").GetMethod("ReadLogicalSize", BindingFlags.Static | BindingFlags.NonPublic);
            object connect = Activator.CreateInstance(assembly.GetType(ns + ".Connect"));
            foreach (MethodInfo method in connect.GetType().GetMethods())
                if (method.Name.StartsWith("GetImage_", StringComparison.Ordinal))
                {
                    object picture = method.Invoke(connect, new object[] { null });
                    try
                    {
                        SizeF actual = (SizeF)read.Invoke(null, new object[] { picture });
                        int expected = (int)assembly.GetType(ns + ".RibbonConfig").GetMethod("GetImageSize", BindingFlags.Static | BindingFlags.NonPublic)
                            .Invoke(null, new object[] { method.Name.Substring("GetImage_".Length) });
                        Check("real Ribbon logical size " + method.Name, Math.Abs(actual.Width - expected) <= 1 && Math.Abs(actual.Height - expected) <= 1);
                    }
                    finally { if (picture != null && Marshal.IsComObject(picture)) Marshal.FinalReleaseComObject(picture); }
                }
        }

        private static void RenderRibbonLayout(Assembly assembly, string ns, string output)
        {
            string xml = (string)assembly.GetType(ns + ".RibbonConfig").GetProperty("RibbonXml").GetValue(null, null);
            XmlDocument doc = new XmlDocument(); doc.LoadXml(xml);
            MethodInfo create = assembly.GetType(ns + ".UiIcons").GetMethod("CreateRibbon");
            using (Bitmap bitmap = new Bitmap(1240, 320))
            using (Graphics g = Graphics.FromImage(bitmap))
            using (Font title = new Font("Microsoft YaHei UI", 16))
            using (Font text = new Font("Microsoft YaHei UI", 10))
            using (Font note = new Font("Microsoft YaHei UI", 9))
            {
                g.Clear(Color.FromArgb(245, 247, 251));
                g.DrawString(ns + " · 0.9.7.0 功能区布局", title, Brushes.DarkSlateGray, 24, 22);
                g.DrawString("来自生成的 Ribbon XML；三行图文横排 + 大图标；16/32 逻辑像素（示意按 1.5 倍绘制）", note, Brushes.SlateGray, 24, 62);
                int x = 24;
                foreach (XmlElement group in doc.GetElementsByTagName("tab")[0].ChildNodes)
                {
                    int groupLeft = x;
                    if (((XmlElement)group.FirstChild).GetAttribute("size") == "normal")
                    {
                        g.FillRectangle(Brushes.White, x, 106, 180, 132);
                        int y = 116;
                        foreach (XmlElement item in group.ChildNodes) { DrawRibbonItem(g, create, item, text, x + 12, y, 24); y += 38; }
                        x += 184;
                        g.DrawString(group.GetAttribute("label"), note, Brushes.SlateGray, groupLeft, 244);
                        g.DrawLine(Pens.LightGray, x + 8, 110, x + 8, 238); x += 24;
                        continue;
                    }
                    foreach (XmlElement entrance in group.ChildNodes)
                    {
                        bool column = entrance.LocalName == "box";
                        int width = column ? 172 : Math.Max(140, (int)g.MeasureString(entrance.GetAttribute("label"), text).Width + 56);
                        g.FillRectangle(Brushes.White, x, 106, width, 132);
                        if (column)
                        {
                            int y = 124;
                            foreach (XmlElement row in entrance.ChildNodes)
                            { DrawRibbonItem(g, create, row, text, x + 12, y, 24); y += 38; }
                        }
                        else if (entrance.GetAttribute("size") == "normal")
                            DrawRibbonItem(g, create, entrance, text, x + 12, 124, 24);
                        else
                        {
                            Color accent = entrance.GetAttribute("id") == "btnFeedback" ? Color.FromArgb(245,158,11) : Color.FromArgb(37, 99, 235);
                            string kind = (string)assembly.GetType(ns + ".RibbonConfig").GetMethod("BuiltinIconKind", All | BindingFlags.Static).Invoke(null, new object[] { entrance.GetAttribute("id") });
                            using (Bitmap icon = (Bitmap)create.Invoke(null, new object[] { kind, accent, 48, 48 })) g.DrawImageUnscaled(icon, x + (width - 48) / 2, 124);
                            string label = entrance.GetAttribute("label") + (entrance.LocalName == "menu" ? " ▾" : "");
                            g.DrawString(label, text, Brushes.DarkSlateGray, x + (width - g.MeasureString(label, text).Width) / 2, 186);
                        }
                        x += width + 4;
                    }
                    g.DrawString(group.GetAttribute("label"), note, Brushes.SlateGray, groupLeft, 244);
                    g.DrawLine(Pens.LightGray, x + 8, 110, x + 8, 238); x += 24;
                }
                g.DrawString("布局示意；WPS 控制整体高度，实际排列和点击需用户重装后验收", note, Brushes.SlateGray, 24, 282);
                bitmap.Save(Path.Combine(output, "功能区-紧凑布局-100.png"), ImageFormat.Png);
            }
        }
        private static void RenderAgentLayout(Assembly assembly, string ns, string output)
        {
            Type cfg = assembly.GetType(ns + ".RibbonConfig");
            XmlDocument doc = new XmlDocument(); doc.LoadXml((string)cfg.GetProperty("RibbonXml").GetValue(null, null));
            using (Bitmap bitmap = new Bitmap(1050, 540))
            using (Graphics g = Graphics.FromImage(bitmap))
            using (Font title = new Font("Microsoft YaHei UI", 18, FontStyle.Bold))
            using (Font text = new Font("Microsoft YaHei UI", 11))
            {
                g.Clear(Color.FromArgb(245,247,251));
                g.DrawString("Agent-Evo · 原生菜单结构示意", title, Brushes.RoyalBlue, 24, 18);
                g.DrawString("主选项卡 → 扩展 → Agent扩展：点击显示，再次点击隐藏", text, Brushes.DimGray, 24, 60);
                int x = 24;
                foreach (XmlElement menu in doc.GetElementsByTagName("tab")[1].FirstChild.ChildNodes)
                {
                    g.FillRectangle(Brushes.White, x, 100, 478, 410);
                    g.DrawString(menu.GetAttribute("label") + " ▼", title, Brushes.RoyalBlue, x + 20, 116);
                    int y = 163;
                    foreach (XmlElement leaf in menu.ChildNodes)
                    { g.DrawString("• " + leaf.GetAttribute("label"), text, Brushes.Black, x + 20, y); y += 35; }
                    x += 520;
                }
                bitmap.Save(Path.Combine(output, "Agent-Evo-菜单结构-100.png"), ImageFormat.Png);
            }
        }

        private static void DrawRibbonItem(Graphics g, MethodInfo create, XmlElement row, Font text, int x, int y, int size)
        {
            string id = row.GetAttribute("id");
            string kind = id == "btnUpdateLinks" ? "update" : id == "btnLinkToExcel" ? "paste" : id == "btnTracePrecedents" ? "precedents" :
                id == "btnTraceDependents" ? "dependents" : id == "btnViewLinks" ? "view" : id == "btnSelectAllTables" ? "table" :
                id == "menuFilterEnhanced" || id == "btnFilterColumns" || id == "btnQuickAdvancedFilter" ? "filter" : row.LocalName == "menu" ? "formula" : "mail";
            Color accent = id == "btnFeedback" ? Color.FromArgb(245,158,11) : id == "btnTraceDependents" ? Color.FromArgb(22, 135, 111) : Color.FromArgb(37, 99, 235);
            using (Bitmap icon = (Bitmap)create.Invoke(null, new object[] { kind, accent, size, size })) g.DrawImageUnscaled(icon, x, y);
            g.DrawString(row.GetAttribute("label") + (row.LocalName == "menu" ? " ▾" : ""), text, Brushes.DarkSlateGray, x + size + 8, y + 1);
        }
        private static void RenderFormulaMenu(Assembly assembly, string ns, string output)
        {
            string xml = (string)assembly.GetType(ns + ".RibbonConfig").GetProperty("RibbonXml").GetValue(null, null);
            XmlDocument doc = new XmlDocument(); doc.LoadXml(xml);
            XmlNamespaceManager names = new XmlNamespaceManager(doc.NameTable); names.AddNamespace("r", doc.DocumentElement.NamespaceURI);
            XmlElement root = (XmlElement)doc.SelectSingleNode("//r:menu[@id='menuOneClickFormula']", names);
            List<string> lines = new List<string>(); MenuLines(root, 0, lines);
            using (Bitmap bitmap = new Bitmap(980, 130 + lines.Count * 30))
            using (Graphics g = Graphics.FromImage(bitmap))
            using (Font title = new Font("Microsoft YaHei UI", 16))
            using (Font text = new Font("Microsoft YaHei UI", 11))
            {
                g.Clear(Color.FromArgb(245,247,251)); g.DrawString("一键公式 · 配置生成的完整菜单树", title, Brushes.DarkSlateGray, 24, 20);
                g.DrawString("大图标入口 32，嵌套图标 16；实际悬停与 Esc 由用户重装验收", text, Brushes.SlateGray, 24, 60);
                for(int i=0;i<lines.Count;i++) g.DrawString(lines[i], text, Brushes.DarkSlateGray, 32, 104+i*30);
                bitmap.Save(Path.Combine(output, "一键公式-菜单结构-100.png"), ImageFormat.Png);
            }
        }
        private static void MenuLines(XmlElement element, int depth, List<string> lines)
        {
            lines.Add(new string(' ',depth*6) + (depth > 0 ? "└─ " : "") + element.GetAttribute("label") + (element.LocalName == "menu" ? "  ▶" : ""));
            foreach(XmlElement child in element.ChildNodes) MenuLines(child,depth+1,lines);
        }

        private static Rectangle AccentBounds(Bitmap image, Color accent)
        {
            int left = image.Width, top = image.Height, right = -1, bottom = -1;
            for (int y = 0; y < image.Height; y++) for (int x = 0; x < image.Width; x++)
            {
                Color p = image.GetPixel(x, y);
                int r = p.R - accent.R, g = p.G - accent.G, b = p.B - accent.B;
                if (r * r + g * g + b * b < 9000)
                { left = Math.Min(left, x); top = Math.Min(top, y); right = Math.Max(right, x); bottom = Math.Max(bottom, y); }
            }
            return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }

        private static TextBox FindTextBox(Control root)
        { foreach (Control control in Descendants(root)) if (control is TextBox) return (TextBox)control; return null; }
        private static IEnumerable<Control> Descendants(Control root)
        { foreach (Control control in root.Controls) { yield return control; foreach (Control child in Descendants(control)) yield return child; } }
        private static object Get(object target, string field) { return target.GetType().GetField(field, Private).GetValue(target); }
        private static DataGridView Grid(object target, string field) { return (DataGridView)Get(target, field); }
        private static void Set(object target, string field, object value) { target.GetType().GetField(field, All).SetValue(target, value); }
        private static void InvokeBase(Form form, string method, object arg) { typeof(Form).GetMethod(method, Private).Invoke(form, new object[] { arg }); }
        private static void Check(string name, bool success)
        { if (success) passed++; else { failed++; Console.WriteLine("FAIL " + name); } }
        private static void Skip(string name) { skipped++; Console.WriteLine("SKIP " + name); }
    }
}
