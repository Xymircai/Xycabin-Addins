using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class PdfToolRequest
    {
        internal string Mode, Output, Text, Ranges, Name, Scope, Color = "808080", Layout = "center";
        internal string[] Files, Folders;
        internal decimal Opacity, Size, Angle;
        internal XmlDocument ToXml()
        {
            var doc = new XmlDocument(); var root = doc.CreateElement("PdfJob"); doc.AppendChild(root);
            root.SetAttribute("mode", Mode); root.SetAttribute("output", Path.GetFullPath(Output));
            root.SetAttribute("name", Name ?? ""); root.SetAttribute("color", Color); root.SetAttribute("layout", Layout);
            root.SetAttribute("text", Text ?? ""); root.SetAttribute("ranges", Ranges ?? "");
            root.SetAttribute("opacity", Opacity.ToString(CultureInfo.InvariantCulture));
            root.SetAttribute("size", Size.ToString(CultureInfo.InvariantCulture)); root.SetAttribute("angle", Angle.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < Files.Length; i++) { var child = doc.CreateElement("File"); child.InnerText = Path.GetFullPath(Files[i]); if (Folders != null && i < Folders.Length) child.SetAttribute("folder", Folders[i] ?? ""); root.AppendChild(child); }
            return doc;
        }
    }
    internal static class PdfWorkerBridge
    {
        private static readonly object gate = new object();
        private static string Digest(byte[] data) { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", ""); }
        private static string Extract()
        {
            using (var stream = typeof(PdfWorkerBridge).Assembly.GetManifestResourceStream("WpsShared.PdfWorker.zip"))
            {
                if (stream == null) throw new InvalidOperationException("PDF 工具未打包，请重新构建加载项。");
                byte[] bytes; using (var copy = new MemoryStream()) { stream.CopyTo(copy); bytes = copy.ToArray(); }
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XyCabin", "PdfWorker", Digest(bytes));
                lock (gate)
                {
                    Directory.CreateDirectory(folder);
                    using (var input = new MemoryStream(bytes)) using (var archive = new ZipArchive(input, ZipArchiveMode.Read))
                    foreach (var entry in archive.Entries)
                    {
                        if (entry.FullName != Path.GetFileName(entry.FullName)) throw new InvalidDataException("PDF 工具包含无效路径。");
                        string target = Path.Combine(folder, entry.FullName); byte[] contents;
                        using (var source = entry.Open()) using (var copy = new MemoryStream()) { source.CopyTo(copy); contents = copy.ToArray(); }
                        if (File.Exists(target) && Digest(File.ReadAllBytes(target)) == Digest(contents)) continue;
                        string temp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                        try { File.WriteAllBytes(temp, contents); if (File.Exists(target)) File.Replace(temp, target, null); else File.Move(temp, target); }
                        finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (Exception cleanup) { Trace.WriteLine("PDF 缓存临时文件清理失败：" + temp + "：" + cleanup); } }
                    }
                }
                return Path.Combine(folder, "PdfWorker.exe");
            }
        }
        internal static string Run(PdfToolRequest job, Func<bool> cancelled)
        {
            string executable = Extract();
            string folder = Path.Combine(Path.GetTempPath(), "ExcelIterate-PDF-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            string request = Path.Combine(folder, "request.xml"); job.ToXml().Save(request);
            File.WriteAllText(request + ".bridge.log", DateTime.UtcNow.ToString("o") + " starting worker" + Environment.NewLine);
            try
            {
                if (cancelled()) return "已取消，未开始处理。";
                var start = new ProcessStartInfo(executable, "\"" + request + "\"") { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(executable) };
                using (var process = Process.Start(start))
                {
                    File.AppendAllText(request + ".bridge.log", "worker pid=" + process.Id + Environment.NewLine);
                    while (!process.WaitForExit(200)) if (cancelled() && !File.Exists(request + ".cancel")) File.WriteAllText(request + ".cancel", "cancel");
                    if (!File.Exists(request + ".result")) throw new InvalidOperationException("PDF 工具未返回结果，退出码：" + process.ExitCode);
                    var result = new XmlDocument { XmlResolver = null };
                    using (var reader = XmlReader.Create(request + ".result", new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) result.Load(reader);
                    var text = new StringBuilder(); int success = result.SelectNodes("/Result/Success").Count, failed = result.SelectNodes("/Result/Failure").Count;
                    text.AppendLine((process.ExitCode == 4 ? "批次已暂停。" : "") + "已生成 " + success + " 个 PDF，失败 " + failed + " 项。" + (process.ExitCode == 3 ? "已取消。" : ""));
                    foreach (XmlElement line in result.DocumentElement.ChildNodes) text.AppendLine((line.Name == "Success" ? "已保存：" : "提示：") + line.InnerText);
                    text.AppendLine("诊断记录：" + folder);
                    return text.ToString();
                }
            }
            finally { Trace.WriteLine("PDF 诊断记录保留：" + folder); }
        }
    }
    internal sealed class PdfToolsForm : Form, IDialogDefaultLayout
    {
        private string mode;
        private readonly bool conversionWindow;
        private readonly PdfFileSelection selection = new PdfFileSelection();
        private readonly ListBox files = new ListBox { Name = "pdfFiles", Dock = DockStyle.Fill, HorizontalScrollbar = true, SelectionMode = SelectionMode.MultiExtended };
        private readonly TextBox output = new TextBox { Name = "pdfOutput", Dock = DockStyle.Fill, ReadOnly = true, Text = "每次处理时选择目录" };
        private readonly TextBox watermark = new TextBox { Name = "watermarkText", Text = "内部资料", Width = 180, MaxLength = 200 };
        private readonly TextBox ranges = new TextBox { Name = "pdfRanges", Width = 200 }, currentRanges = new TextBox { Name = "currentPdfRanges", Width = 160 };
        private readonly NumericUpDown opacity = Number("watermarkOpacity", 1, 100, 20), size = Number("watermarkSize", 8, 144, 48), angle = Number("watermarkAngle", -180, 180, -35);
        private readonly ComboBox layoutChoice = Choice("watermarkLayout", new[] { "居中", "平铺" });
        private readonly ComboBox scope = Choice("currentScope", CurrentScopes());
        private readonly ComboBox conversion = Choice("conversionOutput", new[] { "分别输出 PDF", "转换后合并为一个 PDF" });
        private readonly TextBox log = new TextBox { Name = "pdfLog", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
        private readonly CheckBox recursive = new CheckBox { Name = "pdfRecursive", Text = "包含子文件夹", AutoSize = true }, preserve = new CheckBox { Name = "pdfPreserveFolders", Text = "保留来源子目录结构", AutoSize = true, Checked = true };
        private readonly FlowLayoutPanel toolbar = Ui.ActionBar(), options = Ui.ActionBar(), extensionsBar = Ui.ActionBar();
        private readonly Label hint = Ui.Label("", 9, false), counts = Ui.Label("", 9, false);
        private readonly TabControl tabs = new TabControl { Name = "pdfModes", Dock = DockStyle.Fill, Height = 30 };
        private readonly List<Control> editing = new List<Control>();
        private readonly RoundedButton run = Ui.Button("开始处理…", 120), cancel = Ui.Button("取消处理", 110), colorButton = Ui.Button("水印颜色…", 120);
        private readonly HashSet<string> selectedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool hasExtensionPreference, building, preview;
        private string currentName = "当前文档（执行时绑定）";
        private int watermarkColor = 0x808080;
        private volatile bool cancellation;
        private bool working;
        internal Func<PdfToolRequest, string> ExportCurrent;
        internal Action CloseCurrent;
        internal Func<string> BindCurrent;
        internal string Mode { get { return mode; } }
        public Size MeasureDefaultClient() { PerformLayout(); return new Size(DialogSizing.LogicalClient(this).Width, 700); }
        private static string[] CurrentScopes()
        {
#if WPS_WORD
            return new[] { "整篇文档", "指定页码范围" };
#else
            return new[] { "整个工作簿", "打开窗口时的工作表", "打开窗口时的选中区域" };
#endif
        }
        private static ComboBox Choice(string name, string[] values)
        { var box = new ComboBox { Name = name, Width = 220, DropDownStyle = ComboBoxStyle.DropDownList }; box.Items.AddRange(values); box.SelectedIndex = 0; return box; }
        private static NumericUpDown Number(string name, int min, int max, int value)
        { return new NumericUpDown { Name = name, Minimum = min, Maximum = max, Value = value, Width = 66 }; }
        internal void SetCurrentSource(string name) { currentName = name; RefreshFiles(); }
        internal void ShowCurrentError(string message) { log.Text = "当前文档未绑定：" + message + "\r\n仍可切换到批量文件。"; }
        internal void SetPreview()
        {
            preview = true;
            selection.Sources.Add(new PdfSource { Path = conversionWindow ? @"C:\示例资料\年度报告.docx" : @"C:\示例资料\年度报告.pdf", Folder = "示例资料" });
            selection.Sources.Add(new PdfSource { Path = conversionWindow ? @"C:\示例资料\附件.xlsx" : @"C:\示例资料\报告附件.pdf", Folder = "示例资料" });
            currentName = @"C:\示例资料\当前文档";
            selectedExtensions.Clear(); hasExtensionPreference = false; RebuildExtensions(); RefreshFiles();
            log.Text = "尺寸预览：不读取或处理文件。\r\n保留源文件，同名自动编号。";
        }
        internal static string Title(string value) { return value == "current" || value == "convert" ? "转为 PDF" : "编辑 PDF"; }
        internal void SelectMode(string operation)
        {
            if (working) return;
            int index = conversionWindow ? (operation == "current" ? 0 : 1) : operation == "split" ? 1 : operation == "merge" ? 2 : 0;
            tabs.SelectedIndex = index;
        }
        internal PdfToolsForm(string operation)
        {
            mode = operation; conversionWindow = mode == "current" || mode == "convert"; Text = Title(mode);
            // Keep legacy IDs and their independent preview/size customizations.
            DialogSizing.Prepare(this, "file.pdf." + operation, "document", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 10 };
            for (int i = 0; i < 10; i++) root.RowStyles.Add(new RowStyle(i == 4 || i == 8 ? SizeType.Percent : SizeType.AutoSize, i == 4 ? 55 : i == 8 ? 45 : 0));
            foreach (string name in conversionWindow ? new[] { "当前文档", "批量文件" } : new[] { "添加水印", "拆分", "合并" }) tabs.TabPages.Add(name);
            tabs.SelectedIndex = conversionWindow ? (mode == "current" ? 0 : 1) : mode == "split" ? 1 : mode == "merge" ? 2 : 0;
            tabs.MinimumSize = new Size(0, 32); root.Controls.Add(tabs, 0, 0); editing.Add(tabs);
            hint.AutoSize = true; hint.Dock = DockStyle.Fill; root.Controls.Add(hint, 0, 1);
            toolbar.AutoSize = true;
            AddButton(toolbar, "添加文件", AddFiles, 120); AddButton(toolbar, "添加文件夹", AddFolder, 136);
            AddButton(toolbar, "移除", delegate { foreach (var item in files.SelectedItems.Cast<PdfSource>().ToArray()) selection.Remove(item); RebuildExtensions(); RefreshFiles(); });
            AddButton(toolbar, "清空", delegate { selection.Clear(); RebuildExtensions(); RefreshFiles(); });
            AddButton(toolbar, "上移", delegate { MoveFile(-1); }); AddButton(toolbar, "下移", delegate { MoveFile(1); });
            var fileMenu = new ContextMenuStrip();
            fileMenu.Items.Add("复制完整路径", null, delegate { try { if (files.SelectedItem is PdfSource) Clipboard.SetText(((PdfSource)files.SelectedItem).Path); } catch (Exception e) { log.AppendText(e.Message + "\r\n"); } });
            fileMenu.Items.Add("按名称排序", null, delegate { if (!working) { selection.Sources.Sort((a, b) => StringComparer.OrdinalIgnoreCase.Compare(a.Path, b.Path)); RefreshFiles(); } });
            files.ContextMenuStrip = fileMenu; Disposed += delegate { fileMenu.Dispose(); };
            toolbar.Controls.Add(recursive); root.Controls.Add(toolbar, 0, 2); editing.Add(toolbar);
            extensionsBar.AutoSize = true; root.Controls.Add(extensionsBar, 0, 3); editing.Add(extensionsBar);
            root.Controls.Add(CompactCard(files, "来源文件（按列表顺序；右键复制路径或排序）"), 0, 4); editing.Add(files);
            counts.AutoSize = true; root.Controls.Add(counts, 0, 5);
            var destination = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(0, 4, 0, 4) };
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); destination.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            destination.Controls.Add(Ui.Label("本次输出目录", 10, false), 0, 0); destination.Controls.Add(output, 1, 0); root.Controls.Add(destination, 0, 6);
            options.AutoSize = true; root.Controls.Add(options, 0, 7); editing.Add(options);
            root.Controls.Add(CompactCard(log, "处理结果（保留源文件，同名自动编号）"), 0, 8);
            var actions = Ui.ActionBar(); actions.AutoSize = true; Ui.StylePrimary(run); cancel.Enabled = false;
            run.Click += delegate { Start(); }; cancel.Click += delegate { cancellation = true; cancel.Enabled = false; log.AppendText("正在取消，等待当前处理结束…\r\n"); };
            var open = Ui.Button("打开输出目录", 140); open.Click += delegate { try { if (Directory.Exists(output.Text)) Process.Start(new ProcessStartInfo(output.Text) { UseShellExecute = true }); } catch (Exception e) { log.AppendText(e.Message); } };
            actions.Controls.Add(run); actions.Controls.Add(cancel); actions.Controls.Add(open); root.Controls.Add(actions, 0, 9); Controls.Add(root);
            LoadPreferences();
            tabs.Selecting += delegate(object sender, TabControlCancelEventArgs e) { if (working) e.Cancel = true; };
            tabs.SelectedIndexChanged += delegate {
                mode = conversionWindow ? (tabs.SelectedIndex == 0 ? "current" : "convert") : new[] { "watermark", "split", "merge" }[tabs.SelectedIndex];
                log.Text = mode == "current" ? "当前文档：点击开始后保存并导出绑定文档。" : "已切换任务，请点击开始执行。";
                if (mode == "current" && ExportCurrent == null && BindCurrent != null) try { SetCurrentSource(BindCurrent()); } catch (Exception e) { log.Text = "绑定当前文档失败：" + e.Message; }
                RenderMode();
            };
            recursive.CheckedChanged += delegate { if (!building && selection.Inputs.Count > 0) Scan(); };
            conversion.SelectedIndexChanged += delegate { preserve.Visible = conversion.SelectedIndex == 0; };
            scope.SelectedIndexChanged += delegate { currentRanges.Enabled = scope.SelectedIndex == 1; };
            colorButton.Click += delegate { int? value = HostDialogService.PdfColor(this, watermarkColor); if (value.HasValue) { watermarkColor = value.Value; colorButton.Text = "颜色 #" + watermarkColor.ToString("X6"); } };
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (working) { cancellation = true; if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; log.AppendText("正在取消，请等待处理结束后关闭窗口。\r\n"); } } else SavePreferences(); };
            Disposed += delegate { cancellation = true; if (CloseCurrent != null) { CloseCurrent(); CloseCurrent = null; } };
            KeyPreview = true; KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { if (working) cancellation = true; else Close(); e.Handled = true; } };
            RenderMode();
        }
        private void RenderMode()
        {
            options.SuspendLayout();
            // Reparent retained input controls before disposing their old wrappers.
            foreach (Control control in new Control[] { watermark, opacity, size, angle, layoutChoice, colorButton, ranges, scope, currentRanges, conversion, preserve }) if (control.Parent != null) control.Parent.Controls.Remove(control);
            foreach (Control old in options.Controls.Cast<Control>().ToArray()) old.Dispose();
            options.Controls.Clear();
            toolbar.Visible = extensionsBar.Visible = mode != "current";
            if (mode == "current")
            {
                hint.Text = "先保存绑定文档，再导出；取消保存则停止。每次处理时选择输出目录。";
                options.Controls.Add(Ui.Grip(scope, "导出范围"));
#if WPS_WORD
                options.Controls.Add(Ui.Grip(currentRanges, "页码（如 1-3,5）")); currentRanges.Enabled = scope.SelectedIndex == 1;
#endif
            }
            else if (mode == "convert")
            { hint.Text = "添加文件或文件夹，按后缀筛选。已打开的来源需先保存并关闭；合并必须全部转换成功。"; options.Controls.Add(conversion); options.Controls.Add(preserve); preserve.Visible = conversion.SelectedIndex == 0; }
            else if (mode == "watermark")
            {
                hint.Text = "为每页添加文字水印；与拆分、合并共用来源及顺序。";
                options.Controls.Add(Ui.Grip(watermark, "文字")); options.Controls.Add(Ui.Grip(opacity, "不透明度 %")); options.Controls.Add(Ui.Grip(size, "字号")); options.Controls.Add(Ui.Grip(angle, "角度 °")); options.Controls.Add(layoutChoice); options.Controls.Add(colorButton); options.Controls.Add(preserve); preserve.Visible = true;
            }
            else if (mode == "split") { hint.Text = "页码从 1 开始；留空逐页拆分，或输入 1-3,4-6,7。"; options.Controls.Add(Ui.Grip(ranges, "拆分范围")); options.Controls.Add(preserve); preserve.Visible = true; }
            else hint.Text = "按列表从上到下合并；使用上移、下移调整顺序，仅输出一个 PDF。";
            options.ResumeLayout(true); RefreshFiles();
        }
        private void AddButton(FlowLayoutPanel bar, string text, Action action, int width = 86)
        {
            var button = Ui.Button(text, width);
            button.TextAlign = ContentAlignment.MiddleCenter;
            button.AutoEllipsis = false;
            if (width > 86) button.MinimumSize = new Size(width, 40);
            button.Click += delegate { action(); }; bar.Controls.Add(button);
        }
        private static Control CompactCard(Control content, string title)
        {
            var card = Ui.Card(content, title); card.Padding = new Padding(8); card.Margin = new Padding(0, 0, 0, 6);
            ((TableLayoutPanel)card.Controls[0]).RowStyles[0].Height = 28;
            return card;
        }
        private void AddFiles()
        {
            string[] paths = HostDialogService.OpenFiles(this, "所有文件 (*.*)|*.*");
            if (paths == null || paths.Length == 0) return;
            foreach (string path in paths) selection.AddInput(path);
            Scan();
        }
        private void AddFolder()
        { string path = HostDialogService.Folder(this, "选择来源文件夹", false); if (path != null) { selection.AddInput(path); Scan(); } }
        private void Scan()
        {
            if (working || preview || DialogPreviewFactory.IsPreview(this)) return;
            bool include = recursive.Checked; SetWorking(true); cancellation = false; log.Text = "正在扫描…\r\n";
            // Native Office/WPS message loops may have no WinForms synchronization context.
            // Create the handle here on the owner thread before starting asynchronous work.
            IntPtr ownerHandle = Handle;
            var worker = new BackgroundWorker();
            worker.DoWork += delegate(object sender, DoWorkEventArgs e) { var warnings = new List<string>(); var result = selection.Scan(include, () => cancellation, warnings); e.Result = new object[] { result, warnings }; };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                worker.Dispose(); CompleteOnUi(delegate {
                if (e.Error == null && !cancellation)
                { var result = (object[])e.Result; selection.Commit((List<PdfSource>)result[0]); log.Text = string.Join("\r\n", (List<string>)result[1]); RebuildExtensions(); }
                else log.Text = e.Error is OperationCanceledException || cancellation ? "扫描已取消，保留上次清单。" : "扫描失败：" + e.Error.Message;
                SetWorking(false); RefreshFiles();
                });
            };
            worker.RunWorkerAsync();
        }
        private void CompleteOnUi(Action completion)
        {
            if (IsDisposed || Disposing || !IsHandleCreated) return;
            try
            {
                BeginInvoke(new MethodInvoker(delegate {
                    if (IsDisposed || Disposing) return;
                    try { completion(); }
                    catch (Exception error)
                    {
                        Trace.WriteLine("PDF 界面更新失败：" + error);
                        log.Text = "界面更新失败：" + error.Message;
                        SetWorking(false);
                    }
                }));
            }
            catch (InvalidOperationException error)
            {
                // A queued completion may race with window disposal or host shutdown.
                if (!IsDisposed && !Disposing && IsHandleCreated) Trace.WriteLine("PDF 回调投递失败：" + error);
            }
        }
        private void RebuildExtensions()
        {
            building = true;
            foreach (Control control in extensionsBar.Controls.Cast<Control>().ToArray()) control.Dispose();
            extensionsBar.Controls.Clear();
            foreach (var group in selection.Sources.GroupBy(s => Path.GetExtension(s.Path).ToLowerInvariant()).OrderBy(g => g.Key))
            {
                string ext = group.Key; bool supported = PdfFileSelection.Supported("file" + ext, conversionWindow);
                if (!hasExtensionPreference && supported) selectedExtensions.Add(ext);
                var check = new CheckBox { AutoSize = true, Text = (ext.Length == 0 ? "无后缀" : ext.TrimStart('.')) + "（" + group.Count() + "）" + (supported ? "" : " 不支持"), Checked = supported && selectedExtensions.Contains(ext), Enabled = supported, Tag = ext };
                check.CheckedChanged += delegate { if (building) return; hasExtensionPreference = true; if (check.Checked) selectedExtensions.Add(ext); else selectedExtensions.Remove(ext); RefreshFiles(); };
                extensionsBar.Controls.Add(check);
            }
            building = false;
        }
        private void RefreshFiles()
        {
            string[] selected = files.SelectedItems.Cast<object>().OfType<PdfSource>().Select(s => s.Path).ToArray();
            files.BeginUpdate(); files.Items.Clear();
            if (mode == "current") files.Items.Add(currentName);
            else foreach (var source in selection.Tasks(selectedExtensions, conversionWindow)) files.Items.Add(source);
            for (int i = 0; i < files.Items.Count; i++) { var source = files.Items[i] as PdfSource; if (source != null && selected.Contains(source.Path, StringComparer.OrdinalIgnoreCase)) files.SetSelected(i, true); }
            files.EndUpdate();
            counts.Text = mode == "current" ? "当前文档与批量文件分别保留设置；关闭窗口后清空来源。" : "扫描总数 " + selection.Sources.Count + "，当前参与 " + files.Items.Count + "；未勾选项仅隐藏，不修改源文件。";
            run.Enabled = !working && (mode == "current" || files.Items.Count >= (mode == "merge" ? 2 : 1));
        }
        private void MoveFile(int direction)
        {
            if (files.SelectedIndices.Count != 1) return;
            int index = files.SelectedIndex, target = index + direction;
            if (target < 0 || target >= files.Items.Count) return;
            var source = (PdfSource)files.Items[index]; var neighbor = (PdfSource)files.Items[target];
            int a = selection.Sources.IndexOf(source), b = selection.Sources.IndexOf(neighbor);
            selection.Sources[a] = neighbor; selection.Sources[b] = source; RefreshFiles(); files.SelectedItem = source;
        }
        private void SetWorking(bool value)
        { working = value; foreach (var control in editing) control.Enabled = !value; cancel.Enabled = value; RefreshFiles(); }
        private string PreferencesPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XyCabin", typeof(PdfToolsForm).Namespace, "pdf-" + (conversionWindow ? "convert" : "edit") + "-preferences.xml"); } }
        private void LoadPreferences()
        {
            if (DialogPreviewFactory.IsConstructing) return;
            try
            {
                if (!File.Exists(PreferencesPath)) return;
                var xml = new XmlDocument { XmlResolver = null };
                using (var reader = XmlReader.Create(PreferencesPath, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) xml.Load(reader);
                if (xml.DocumentElement.Name != "PdfPreferences" || xml.DocumentElement.GetAttribute("version") != "1") return;
                recursive.Checked = xml.DocumentElement.GetAttribute("recursive") == "true"; preserve.Checked = xml.DocumentElement.GetAttribute("preserve") != "false";
                hasExtensionPreference = xml.DocumentElement.GetAttribute("extensionsSet") == "true";
                foreach (XmlElement item in xml.SelectNodes("/PdfPreferences/Extension")) selectedExtensions.Add(item.InnerText);
            }
            catch (Exception e) { log.Text = "偏好读取失败，使用默认值：" + e.Message; }
        }
        private void SavePreferences()
        {
            if (preview || DialogPreviewFactory.IsPreview(this)) return;
            string temp = PreferencesPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(PreferencesPath));
                var xml = new XmlDocument(); var root = xml.CreateElement("PdfPreferences"); xml.AppendChild(root);
                root.SetAttribute("version", "1"); root.SetAttribute("recursive", recursive.Checked ? "true" : "false"); root.SetAttribute("preserve", preserve.Checked ? "true" : "false"); root.SetAttribute("extensionsSet", hasExtensionPreference ? "true" : "false");
                foreach (string ext in selectedExtensions.OrderBy(s => s)) { var item = xml.CreateElement("Extension"); item.InnerText = ext; root.AppendChild(item); }
                xml.Save(temp); if (File.Exists(PreferencesPath)) File.Replace(temp, PreferencesPath, null); else File.Move(temp, PreferencesPath);
            }
            catch (Exception e) { System.Diagnostics.Trace.WriteLine("PDF 偏好保存失败：" + e); }
            finally { try { if (File.Exists(temp)) File.Delete(temp); } catch (Exception e) { System.Diagnostics.Trace.WriteLine(e); } }
        }
        private void Start()
        {
            if (working || preview || DialogPreviewFactory.IsPreview(this)) return;
            if (mode == "current" && ExportCurrent == null)
            {
                try { if (BindCurrent != null) SetCurrentSource(BindCurrent()); }
                catch (Exception e) { log.Text = "绑定当前文档失败：" + e.Message; return; }
                if (ExportCurrent == null) { log.Text = "请从 Word／Excel 加载项打开；当前入口未绑定用户文档。"; return; }
            }
            var tasks = selection.Tasks(selectedExtensions, conversionWindow);
            if (mode != "current" && tasks.Length < (mode == "merge" ? 2 : 1)) { log.Text = "没有足够的参与文件，请检查后缀筛选。"; return; }
            if (mode == "watermark" && string.IsNullOrWhiteSpace(watermark.Text)) { log.Text = "请输入水印文字。"; return; }
#if WPS_WORD
            if (mode == "current" && scope.SelectedIndex == 1 && string.IsNullOrWhiteSpace(currentRanges.Text)) { log.Text = "请输入要导出的页码范围。"; return; }
#endif
            string selected = HostDialogService.Folder(this, "选择本次 PDF 输出文件夹", true); if (selected == null) return;
            output.Text = selected; selection.OutputFolders.Add(selected);
            var request = new PdfToolRequest { Mode = mode == "convert" && conversion.SelectedIndex == 1 ? "convertmerge" : mode, Files = tasks.Select(s => s.Path).ToArray(), Folders = preserve.Checked && mode != "current" && mode != "merge" && !(mode == "convert" && conversion.SelectedIndex == 1) ? tasks.Select(s => s.Folder).ToArray() : null, Output = selected, Text = watermark.Text, Ranges = mode == "current" ? currentRanges.Text : ranges.Text, Opacity = opacity.Value, Size = size.Value, Angle = angle.Value, Color = watermarkColor.ToString("X6"), Layout = layoutChoice.SelectedIndex == 1 ? "tile" : "center" };
#if WPS_WORD
            request.Scope = "document"; if (mode == "current" && scope.SelectedIndex == 0) request.Ranges = "";
#else
            request.Scope = new[] { "book", "sheet", "selection" }[scope.SelectedIndex];
#endif
            SetWorking(true); cancellation = false; log.Text = "正在处理，请稍候…\r\n"; string scratch = null;
            IntPtr ownerHandle = Handle;
            try
            {
                if (mode == "current")
                {
                    scratch = ExportCurrent(request);
                    if (request.Mode != "publish" || request.Files == null || request.Files.Length != 1 ||
                        !string.Equals(Path.GetExtension(request.Files[0]), ".pdf", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidOperationException("当前文档未生成有效的 PDF，已停止；请重新打开导出窗口。");
                }
            }
            catch (Exception error) { SetWorking(false); log.Text = "已停止：" + error.Message; return; }
            var worker = new BackgroundWorker();
            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                string cleanup = "";
                try { e.Result = PdfWorkerBridge.Run(request, () => cancellation); }
                finally { if (scratch != null) try { Directory.Delete(scratch, true); } catch (Exception error) { cleanup = "临时目录清理失败：" + scratch + "：" + error.Message; System.Diagnostics.Trace.WriteLine(cleanup); } }
                if (cleanup.Length > 0) e.Result = Convert.ToString(e.Result) + "\r\n" + cleanup;
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                worker.Dispose(); CompleteOnUi(delegate {
                log.Text = e.Error == null ? Convert.ToString(e.Result) : "处理失败：" + e.Error.Message;
                foreach (string line in log.Lines) if (line.StartsWith("已保存：")) selection.Generated.Add(line.Substring(4));
                SetWorking(false);
                });
            };
            worker.RunWorkerAsync();
        }
    }
}
