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
        internal string[] Files;
        internal decimal Opacity, Size, Angle;
        internal XmlDocument ToXml()
        {
            var doc = new XmlDocument(); var root = doc.CreateElement("PdfJob"); doc.AppendChild(root);
            root.SetAttribute("mode", Mode); root.SetAttribute("output", Path.GetFullPath(Output));
            root.SetAttribute("name", Name ?? ""); root.SetAttribute("color", Color); root.SetAttribute("layout", Layout);
            root.SetAttribute("text", Text ?? ""); root.SetAttribute("ranges", Ranges ?? "");
            root.SetAttribute("opacity", Opacity.ToString(CultureInfo.InvariantCulture));
            root.SetAttribute("size", Size.ToString(CultureInfo.InvariantCulture)); root.SetAttribute("angle", Angle.ToString(CultureInfo.InvariantCulture));
            foreach (string file in Files) { var child = doc.CreateElement("File"); child.InnerText = Path.GetFullPath(file); root.AppendChild(child); }
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
                        finally { if (File.Exists(temp)) File.Delete(temp); }
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
            try
            {
                if (cancelled()) return "已取消，未开始处理。";
                var start = new ProcessStartInfo(executable, "\"" + request + "\"") { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(executable) };
                using (var process = Process.Start(start))
                {
                    while (!process.WaitForExit(200)) if (cancelled() && !File.Exists(request + ".cancel")) File.WriteAllText(request + ".cancel", "cancel");
                    if (!File.Exists(request + ".result")) throw new InvalidOperationException("PDF 工具未返回结果，退出码：" + process.ExitCode);
                    var result = new XmlDocument { XmlResolver = null };
                    using (var reader = XmlReader.Create(request + ".result", new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) result.Load(reader);
                    var text = new StringBuilder(); int success = result.SelectNodes("/Result/Success").Count, failed = result.SelectNodes("/Result/Failure").Count;
                    text.AppendLine((process.ExitCode == 4 ? "批次已暂停。" : "") + "已生成 " + success + " 个 PDF，失败 " + failed + " 项。" + (process.ExitCode == 3 ? "已取消。" : ""));
                    foreach (XmlElement line in result.DocumentElement.ChildNodes) text.AppendLine((line.Name == "Success" ? "已保存：" : "提示：") + line.InnerText);
                    return text.ToString();
                }
            }
            finally { try { Directory.Delete(folder, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
        }
    }
    internal sealed class PdfToolsForm : Form, IDialogDefaultLayout
    {
        private readonly string mode;
        private readonly ListBox files = new ListBox { Name = "pdfFiles", Dock = DockStyle.Fill, HorizontalScrollbar = true, SelectionMode = SelectionMode.MultiExtended };
        private readonly TextBox output = new TextBox { Name = "pdfOutput", Dock = DockStyle.Fill, ReadOnly = true };
        private readonly TextBox watermark = new TextBox { Name = "watermarkText", Text = "内部资料", Width = 200, MaxLength = 200 };
        private readonly TextBox ranges = new TextBox { Name = "pdfRanges", Width = 250 };
        private readonly NumericUpDown opacity = Number("watermarkOpacity", 1, 100, 20), size = Number("watermarkSize", 8, 144, 48), angle = Number("watermarkAngle", -180, 180, -35);
        private readonly ComboBox layoutChoice = Choice("watermarkLayout", new[] { "居中", "平铺" });
        private readonly ComboBox scope = Choice("currentScope", CurrentScopes());
        private readonly ComboBox conversion = Choice("conversionOutput", new[] { "分别输出 PDF", "转换后合并为一个 PDF" });
        private readonly TextBox log = new TextBox { Name = "pdfLog", Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false };
        private readonly List<Control> editing = new List<Control>();
        private readonly RoundedButton run = Ui.Button("开始处理…", 120), cancel = Ui.Button("取消处理", 110), colorButton = Ui.Button("水印颜色…", 120);
        private int watermarkColor = 0x808080;
        private volatile bool cancellation;
        private bool working;
        internal Func<PdfToolRequest, string> ExportCurrent;
        internal Action CloseCurrent;
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
        internal void SetCurrentSource(string name) { files.Items.Clear(); files.Items.Add(name); }
        internal void SetPreview()
        {
            files.Items.Add(mode == "convert" || mode == "current" ? @"C:\示例资料\年度报告.docx" : @"C:\示例资料\年度报告.pdf");
            if (mode == "merge") files.Items.Add(@"C:\示例资料\报告附件.pdf");
            output.Text = "每次点击开始处理时选择目录";
            if (mode == "split") ranges.Text = "1-3,4-6,7";
            log.Text = "尺寸预览：不读取或处理文件。\r\n输出为新文件，同名自动编号。";
        }
        internal static string Title(string value) { return value == "current" ? "当前文档转 PDF" : value == "convert" ? "批量文件转 PDF" : value == "watermark" ? "PDF 添加水印" : value == "split" ? "PDF 拆分" : "PDF 合并"; }
        private static NumericUpDown Number(string name, int min, int max, int value)
        { return new NumericUpDown { Name = name, Minimum = min, Maximum = max, Value = value, Width = 66 }; }
        internal PdfToolsForm(string operation)
        {
            mode = operation; Text = Title(mode);
            DialogSizing.Prepare(this, "file.pdf." + mode, "document", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 7 };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 55)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 45)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var hint = Ui.Label(mode == "current" ? "导出前先保存原文档；取消保存或保存失败则停止。每次处理都选择输出目录。" : mode == "convert" ? "按列表顺序处理。已打开的来源需先保存并关闭；合并必须全部转换成功。" : mode == "merge" ? "按列表从上到下合并；使用上移、下移调整顺序。" : mode == "split" ? "页码从 1 开始；留空逐页拆分，或输入 1-3,4-6,7。" : "给每页添加文字水印，支持颜色、不透明度、字号、角度、居中或平铺。", 9, false);
            hint.AutoSize = true; hint.Dock = DockStyle.Fill; root.Controls.Add(hint, 0, 0);
            var toolbar = Ui.ActionBar(); toolbar.AutoSize = true;
            if (mode != "current")
            {
                AddButton(toolbar, "添加文件…", AddFiles); AddButton(toolbar, "移除", delegate { foreach (int index in files.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToArray()) files.Items.RemoveAt(index); });
                AddButton(toolbar, "上移", delegate { MoveFile(-1); }); AddButton(toolbar, "下移", delegate { MoveFile(1); });
            }
            if (mode == "convert") toolbar.Controls.Add(conversion);
            editing.Add(toolbar); root.Controls.Add(toolbar, 0, 1); root.Controls.Add(Ui.Card(files, mode == "current" ? "绑定的当前文档" : "来源文件（按列表顺序）"), 0, 2); editing.Add(files);
            var destination = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(0, 8, 0, 8) };
            destination.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); destination.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            destination.Controls.Add(Ui.Label("本次输出目录", 10, false), 0, 0); destination.Controls.Add(output, 1, 0); root.Controls.Add(destination, 0, 3);
            var options = Ui.ActionBar(); options.AutoSize = true;
            if (mode == "watermark")
            {
                options.Controls.Add(Ui.Grip(watermark, "文字"));
                options.Controls.Add(Ui.Grip(opacity, "不透明度 %"));
                options.Controls.Add(Ui.Grip(size, "字号"));
                options.Controls.Add(Ui.Grip(angle, "角度 °"));
                options.Controls.Add(layoutChoice); options.Controls.Add(colorButton);
                colorButton.Click += delegate { int? selected = HostDialogService.PickColor(this, watermarkColor); if (selected.HasValue) { watermarkColor = selected.Value; colorButton.Text = "颜色 #" + watermarkColor.ToString("X6"); } };
            }
            if (mode == "split") { options.Controls.Add(Ui.Grip(ranges, "拆分范围")); }
            if (mode == "current")
            {
                options.Controls.Add(Ui.Grip(scope, "导出范围"));
#if WPS_WORD
                options.Controls.Add(Ui.Grip(ranges, "页码（如 1-3,5）")); ranges.Enabled = false;
                scope.SelectedIndexChanged += delegate { ranges.Enabled = scope.SelectedIndex == 1; };
#endif
            }
            editing.Add(options); root.Controls.Add(options, 0, 4); root.Controls.Add(Ui.Card(log, "处理结果（保留源文件，同名自动编号）"), 0, 5);
            var actions = Ui.ActionBar(); actions.AutoSize = true; Ui.StylePrimary(run); cancel.Enabled = false;
            run.Click += delegate { Start(); }; cancel.Click += delegate { cancellation = true; cancel.Enabled = false; log.AppendText("正在取消，等待当前处理结束…\r\n"); };
            var open = Ui.Button("打开输出目录", 140); open.Click += delegate { try { if (Directory.Exists(output.Text)) Process.Start(new ProcessStartInfo(output.Text) { UseShellExecute = true }); } catch (Exception e) { log.AppendText(e.Message); } };
            actions.Controls.Add(run); actions.Controls.Add(cancel); actions.Controls.Add(open); root.Controls.Add(actions, 0, 6); Controls.Add(root);
            FormClosing += delegate(object sender, FormClosingEventArgs e) { if (working) { cancellation = true; if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; log.AppendText("正在取消，请等待当前处理结束后关闭窗口。\r\n"); } } };
            Disposed += delegate { cancellation = true; if (CloseCurrent != null) { CloseCurrent(); CloseCurrent = null; } };
            KeyPreview = true; KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { if (working) cancellation = true; else Close(); e.Handled = true; } };
        }
        private void AddButton(FlowLayoutPanel bar, string text, Action action)
        { var button = Ui.Button(text, 100); button.Click += delegate { action(); }; bar.Controls.Add(button); }
        private void AddFiles()
        {
            foreach (string path in HostDialogService.OpenFiles(this, mode == "convert" ? "文档和图片|*.doc;*.docx;*.rtf;*.odt;*.wps;*.xls;*.xlsx;*.xlsm;*.xlsb;*.ods;*.et;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.gif" : "PDF 文件 (*.pdf)|*.pdf"))
                if (!files.Items.Cast<string>().Contains(path, StringComparer.OrdinalIgnoreCase)) files.Items.Add(path);
        }
        private void MoveFile(int direction)
        {
            if (files.SelectedIndices.Count != 1) return; int index = files.SelectedIndex, target = index + direction;
            if (target < 0 || target >= files.Items.Count) return;
            object value = files.Items[index]; files.Items.RemoveAt(index); files.Items.Insert(target, value); files.SelectedIndex = target;
        }
        private void SetWorking(bool value)
        { working = value; foreach (var control in editing) control.Enabled = !value; run.Enabled = !value; cancel.Enabled = value; }
        private void Start()
        {
            if (working || DialogPreviewFactory.IsPreview(this)) return;
            if (mode == "current" && ExportCurrent == null) { log.Text = "请从 Word／Excel 加载项的当前文档入口打开；独立试用程序不绑定用户文档。"; return; }
            if (mode != "current" && (files.Items.Count == 0 || (mode == "merge" && files.Items.Count < 2))) { log.Text = "请添加" + (mode == "merge" ? "至少两个 PDF。" : "来源文件。"); return; }
            if (mode == "watermark" && string.IsNullOrWhiteSpace(watermark.Text)) { log.Text = "请输入水印文字。"; return; }
#if WPS_WORD
            if (mode == "current" && scope.SelectedIndex == 1 && string.IsNullOrWhiteSpace(ranges.Text)) { log.Text = "请输入要导出的页码范围。"; return; }
#endif
            string selected = HostDialogService.Folder(this, "选择本次 PDF 输出文件夹", true);
            if (selected == null) return;
            output.Text = selected;
            var request = new PdfToolRequest { Mode = mode == "convert" && conversion.SelectedIndex == 1 ? "convertmerge" : mode, Files = files.Items.Cast<string>().ToArray(), Output = selected, Text = watermark.Text, Ranges = ranges.Text, Opacity = opacity.Value, Size = size.Value, Angle = angle.Value, Color = watermarkColor.ToString("X6"), Layout = layoutChoice.SelectedIndex == 1 ? "tile" : "center" };
#if WPS_WORD
            request.Scope = "document"; if (mode == "current" && scope.SelectedIndex == 0) request.Ranges = "";
#else
            request.Scope = new[] { "book", "sheet", "selection" }[scope.SelectedIndex];
#endif
            SetWorking(true); cancellation = false; log.Text = "正在处理，请稍候…\r\n";
            string scratch = null;
            try { if (mode == "current") scratch = ExportCurrent(request); }
            catch (Exception error) { SetWorking(false); log.Text = "已停止：" + error.Message; return; }
            var worker = new BackgroundWorker();
            worker.DoWork += delegate(object sender, DoWorkEventArgs e)
            {
                try { e.Result = PdfWorkerBridge.Run(request, () => cancellation); }
                finally { if (scratch != null) { foreach (string file in Directory.GetFiles(scratch)) File.Delete(file); Directory.Delete(scratch); } }
            };
            worker.RunWorkerCompleted += delegate(object sender, RunWorkerCompletedEventArgs e)
            {
                working = false; worker.Dispose(); if (IsDisposed || Disposing) return;
                log.Text = e.Error == null ? Convert.ToString(e.Result) : "处理失败：" + e.Error.Message; SetWorking(false);
            };
            worker.RunWorkerAsync();
        }
    }
}
