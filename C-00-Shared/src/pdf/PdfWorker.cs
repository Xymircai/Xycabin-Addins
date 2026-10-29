using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Xml;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Fonts;

// Runs outside WPS: PDF library versions and Office automation never enter its AppDomain.
internal static class PdfWorker
{
    private static string cancelFile;
    private static void Log(string message)
    {
        if (cancelFile == null) return;
        try { File.AppendAllText(cancelFile + ".log", DateTime.UtcNow.ToString("o") + " " + message + Environment.NewLine); }
        catch (IOException) { } catch (UnauthorizedAccessException) { }
    }
    private static readonly List<string> cleanupWarnings = new List<string>();
    private static void Cleanup(Action action, string context)
    { try { Log(context); action(); } catch (Exception error) { cleanupWarnings.Add(context + "：" + Details(error)); } }
    private static string Details(Exception error)
    { return error.GetType().Name + "：" + error.Message + "（错误码 0x" + error.HResult.ToString("X8") + "）"; }
    private static string OutputFolder(string output, XmlElement item)
    {
        string relative = item.GetAttribute("folder");
        if (relative.Length == 0) return output;
        if (Path.IsPathRooted(relative) || relative.Split('\\', '/').Any(p => p == ".." || p == "." || p.Length == 0 || p.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
            throw new ArgumentException("无效输出相对目录：" + relative);
        string target = Path.GetFullPath(Path.Combine(output, relative));
        if (!target.StartsWith(output.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("输出目录越界。");
        Directory.CreateDirectory(target); return target;
    }
    private static void CheckCancel() { if (File.Exists(cancelFile)) throw new OperationCanceledException("已取消；已生成的文件保留。"); }
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length != 1) return 2;
        string request = Path.GetFullPath(args[0]); cancelFile = request + ".cancel";
        var report = new XmlDocument(); var root = report.CreateElement("Result"); report.AppendChild(root);
        int exit = 0; string context = "读取请求：" + request, outputContext = null;
        try
        {
            Log("worker-start pid=" + Process.GetCurrentProcess().Id);
            var xml = new XmlDocument { XmlResolver = null };
            using (var reader = XmlReader.Create(request, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) xml.Load(reader);
            XmlElement job = xml.DocumentElement;
            string mode = job.GetAttribute("mode"), output = Path.GetFullPath(job.GetAttribute("output"));
            outputContext = output;
            if (!Directory.Exists(output)) throw new DirectoryNotFoundException("输出文件夹不存在：" + output);
            var inputs = job.SelectNodes("File").Cast<XmlElement>().ToArray();
            var files = inputs.Select(n => Path.GetFullPath(n.InnerText)).ToArray();
            context = "阶段：" + mode + "；来源：" + string.Join("；", files) + "；输出：" + output;
            if (files.Length == 0) throw new ArgumentException("请添加文件。");
            if (mode == "convert" || mode == "convertmerge")
                foreach (string file in files) { Log("preflight " + file); EnsureClosed(file); }
            if (mode == "convertmerge")
            {
                string scratch = Path.Combine(Path.GetTempPath(), "XyPdfMerge-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(scratch);
                try
                {
                    var converted = new List<string>();
                    foreach (string file in files) { CheckCancel(); converted.Add(ConvertFile(file, scratch)); }
                    using (var merged = new PdfDocument())
                    {
                        foreach (string file in converted) using (var source = OpenPdf(file, PdfDocumentOpenMode.Import))
                            foreach (PdfPage page in source.Pages) { CheckCancel(); merged.AddPage(page); }
                        Add(report, root, "Success", Save(merged, output, "转换合并结果"));
                    }
                }
                finally { Cleanup(() => Directory.Delete(scratch, true), "清理合并临时目录 " + scratch); }
            }
            else if (mode == "publish")
            {
                if (files.Length != 1) throw new ArgumentException("当前文档导出来源必须唯一。");
                string name = job.GetAttribute("name");
                if (name.Length == 0 || name != Path.GetFileName(name) || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) throw new ArgumentException("无效输出名称。");
                using (var source = OpenPdf(files[0], PdfDocumentOpenMode.Import))
                {
                    string ranges = job.GetAttribute("ranges");
                    if (string.IsNullOrWhiteSpace(ranges))
                    {
                        string temp = Path.Combine(output, ".pdf-" + Guid.NewGuid().ToString("N") + ".tmp");
                        try { File.Copy(files[0], temp); CheckCancel(); Add(report, root, "Success", Publish(temp, output, name)); }
                        finally { Cleanup(() => { if (File.Exists(temp)) File.Delete(temp); }, "清理临时文件 " + temp); }
                    }
                    else using (var part = new PdfDocument())
                    {
                        foreach (int page in Groups(ranges, source.PageCount).SelectMany(g => g)) { CheckCancel(); part.AddPage(source.Pages[page - 1]); }
                        Add(report, root, "Success", Save(part, output, name));
                    }
                }
            }
            else if (mode == "merge")
            {
                if (files.Length < 2) throw new ArgumentException("合并至少需要两个 PDF。");
                using (var merged = new PdfDocument())
                {
                    foreach (string file in files) { CheckCancel(); using (var source = OpenPdf(file, PdfDocumentOpenMode.Import)) foreach (PdfPage page in source.Pages) { CheckCancel(); merged.AddPage(page); } }
                    string saved = Save(merged, output, "合并结果"); Add(report, root, "Success", saved);
                }
            }
            else foreach (XmlElement input in inputs)
            {
                CheckCancel();
                string file = Path.GetFullPath(input.InnerText), destination = output, stage = "创建输出目录";
                try
                {
                    destination = OutputFolder(output, input); stage = "处理 " + mode;
                    string basis = Path.GetFileNameWithoutExtension(file);
                    if (mode == "convert")
                    {
                        Add(report, root, "Success", ConvertFile(file, destination));
                    }
                    else if (mode == "watermark")
                    {
                        string text = job.GetAttribute("text").Trim();
                        if (text.Length == 0 || text.Length > 200) throw new ArgumentException("水印文字需为 1–200 个字符。");
                        double opacity = Number(job, "opacity", 1, 100) / 100, size = Number(job, "size", 8, 144), angle = Number(job, "angle", -180, 180);
                        using (var doc = OpenPdf(file, PdfDocumentOpenMode.Modify))
                        {
                            if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new WatermarkFontResolver();
                            var font = new XFont("WatermarkFont", size, XFontStyleEx.Regular);
                            foreach (PdfPage page in doc.Pages)
                            {
                                CheckCancel();
                                using (var g = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append))
                                {
                                    double w = page.Width.Point, h = page.Height.Point;
                                    string rgb = job.GetAttribute("color"); if (rgb.Length == 0) rgb = "808080";
                                    int color; if (rgb.Length != 6 || !int.TryParse(rgb, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out color)) throw new ArgumentException("水印颜色需为六位十六进制 RGB。");
                                    string layout = job.GetAttribute("layout"); if (layout.Length == 0) layout = "center";
                                    if (layout != "center" && layout != "tile") throw new ArgumentException("未知水印布局。");
                                    g.TranslateTransform(w / 2, h / 2); g.RotateTransform(angle);
                                    double width = g.MeasureString(text, font).Width;
                                    double fit = Math.Min(1, Math.Min(w, h) * 0.9 / Math.Max(1, width));
                                    g.ScaleTransform(fit);
                                    var brush = new XSolidBrush(XColor.FromArgb((int)Math.Round(opacity * 255), (color >> 16) & 255, (color >> 8) & 255, color & 255));
                                    if (layout == "center") g.DrawString(text, font, brush, new XPoint(0, 0), XStringFormats.Center);
                                    else
                                    {
                                        double extent = Math.Sqrt(w * w + h * h) / fit;
                                        for (double y = -extent; y <= extent; y += size * 3 + 32)
                                            for (double x = -extent; x <= extent; x += width + 64)
                                            { CheckCancel(); g.DrawString(text, font, brush, new XPoint(x, y), XStringFormats.Center); }
                                    }
                                }
                            }
                            Add(report, root, "Success", Save(doc, destination, basis + "_水印"));
                        }
                    }
                    else if (mode == "split")
                    {
                        using (var doc = OpenPdf(file, PdfDocumentOpenMode.Import))
                        {
                            List<int[]> groups = Groups(job.GetAttribute("ranges"), doc.PageCount);
                            for (int i = 0; i < groups.Count; i++)
                            {
                                CheckCancel(); using (var part = new PdfDocument())
                                {
                                    foreach (int page in groups[i]) { CheckCancel(); part.AddPage(doc.Pages[page - 1]); }
                                    Add(report, root, "Success", Save(part, destination, basis + "_" + (i + 1).ToString("D3")));
                                }
                            }
                        }
                    }
                    else throw new ArgumentException("未知 PDF 操作。");
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception error) { exit = 1; Add(report, root, "Failure", "阶段：" + stage + "；来源：" + file + "；输出：" + destination + "；" + Details(error)); DiagnoseOutput(report, root, destination); }
            }
        }
        catch (SourceOpenException error) { exit = 4; Add(report, root, "Blocked", error.Message); }
        catch (OperationCanceledException error) { exit = 3; Add(report, root, "Cancelled", error.Message); }
        catch (Exception error) { exit = 1; Add(report, root, "Failure", context + "；" + Details(error)); if (outputContext != null) DiagnoseOutput(report, root, outputContext); }
        foreach (string warning in cleanupWarnings) Add(report, root, "Warning", warning);
        report.Save(request + ".result"); Log("worker-complete exit=" + exit); return exit;
    }
    // Diagnose only after a failure, without changing permissions or retrying elsewhere.
    private static void DiagnoseOutput(XmlDocument report, XmlElement root, string output)
    {
        string probe = Path.Combine(output, ".pdf-" + Guid.NewGuid().ToString("N") + ".tmp");
        try { using (var stream = new FileStream(probe, FileMode.CreateNew, FileAccess.Write, FileShare.None)) stream.WriteByte(0); }
        catch (Exception error) { Add(report, root, "Warning", "失败后输出目录检测：" + output + "；" + Details(error)); }
        finally { Cleanup(() => { if (File.Exists(probe)) File.Delete(probe); }, "清理目录检测文件 " + probe); }
    }
    private sealed class SourceOpenException : IOException { internal SourceOpenException(string message) : base(message) { } }
    private static bool IsOffice(string path)
    { return new[] { ".doc", ".docx", ".rtf", ".odt", ".wps", ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".et" }.Contains(Path.GetExtension(path).ToLowerInvariant()); }
    private static void EnsureClosed(string path)
    {
        if (!IsOffice(path) || !File.Exists(path)) return;
        string marker = Path.Combine(Path.GetDirectoryName(path), "~$" + Path.GetFileName(path));
        if (File.Exists(marker)) throw new SourceOpenException("批次已暂停：请先保存并关闭来源文件，再重试：" + path);
        try { using (var input = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.None)) { } }
        catch (IOException) { throw new SourceOpenException("批次已暂停：来源文件正在打开或被占用，请保存并关闭后重试：" + path); }
        bool sourceWord = new[] { ".doc", ".docx", ".rtf", ".odt", ".wps" }.Contains(Path.GetExtension(path).ToLowerInvariant());
        foreach (string id in sourceWord ? new[] { "KWPS.Application", "WPS.Application", "Word.Application" } : new[] { "KET.Application", "et.Application", "Excel.Application" })
        {
            object app = null, items = null;
            try
            {
                try { Log("preflight-active " + id); app = Marshal.GetActiveObject(id); } catch (COMException) { continue; }
                bool word = id == "KWPS.Application" || id == "WPS.Application" || id == "Word.Application";
                items = word ? (object)((dynamic)app).Documents : (object)((dynamic)app).Workbooks;
                for (int i = 1; i <= Convert.ToInt32(((dynamic)items).Count); i++)
                {
                    object item = ((dynamic)items).Item(i);
                    try { if (string.Equals(Convert.ToString(((dynamic)item).FullName), path, StringComparison.OrdinalIgnoreCase)) throw new SourceOpenException("批次已暂停：请保存并关闭来源文件，再重试：" + path); }
                    finally { Release(item); }
                }
            }
            catch (COMException error)
            {
                if (error.ErrorCode != unchecked((int)0x80010001) && error.ErrorCode != unchecked((int)0x8001010A)) throw;
                throw new SourceOpenException("批次已暂停：宿主忙或存在待处理对话框，无法确认来源占用状态。请处理宿主提示后重试（" + id + "）：" + path);
            }
            finally { Release(items); Release(app); }
        }
    }
    private static string ConvertFile(string file, string output)
    {
        string basis = Path.GetFileNameWithoutExtension(file), ext = Path.GetExtension(file).ToLowerInvariant();
        if (new[] { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif" }.Contains(ext))
        { using (var doc = ImagePdf(file)) return Save(doc, output, basis); }
        return OfficePdf(file, output, basis);
    }
    private static double Number(XmlElement job, string name, double min, double max)
    {
        double value; if (!double.TryParse(job.GetAttribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out value) || double.IsNaN(value) || value < min || value > max)
            throw new ArgumentException("无效参数：" + name);
        return value;
    }
    private static void Add(XmlDocument doc, XmlElement root, string kind, string text)
    { var item = doc.CreateElement(kind); item.InnerText = text; root.AppendChild(item); }
    private static PdfDocument OpenPdf(string file, PdfDocumentOpenMode mode)
    {
        var doc = PdfReader.Open(file, mode);
        // 6.2.4 resets SecuritySettings.IsEncrypted after reading, and may not detect
        // a direct /Encrypt dictionary. Inspect the parsed trailer, never copy it as plaintext.
        var property = typeof(PdfDocument).GetProperty("Trailer", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
        var trailer = property == null ? null : property.GetValue(doc, null) as PdfDictionary;
        if (trailer == null) { doc.Dispose(); throw new InvalidOperationException("无法检查 PDF 加密结构。"); }
        if (trailer.Elements.ContainsKey("/Encrypt")) { doc.Dispose(); throw new InvalidOperationException("请先解密 PDF 后再处理。"); }
        if (doc.PageCount == 0) { doc.Dispose(); throw new InvalidDataException("PDF 没有页面。"); }
        return doc;
    }
    private sealed class WatermarkFontResolver : IFontResolver
    {
        public FontResolverInfo ResolveTypeface(string familyName, bool bold, bool italic) { return new FontResolverInfo("WindowsSimHei"); }
        public byte[] GetFont(string faceName)
        {
            // Use the installed Chinese TTF; YaHei is a TTC collection that this engine cannot embed.
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "simhei.ttf");
            if (!File.Exists(path)) throw new InvalidOperationException("未找到 Windows 黑体字体（simhei.ttf），请安装中文补充字体后重试。");
            return File.ReadAllBytes(path);
        }
    }
    internal static List<int[]> Groups(string input, int count)
    {
        if (count < 1) throw new ArgumentException("PDF 没有页面。");
        if (string.IsNullOrWhiteSpace(input)) return Enumerable.Range(1, count).Select(n => new[] { n }).ToList();
        var result = new List<int[]>();
        foreach (string token in input.Replace('，', ',').Split(','))
        {
            string[] ends = token.Trim().Split('-'); int first, last;
            if (ends.Length > 2 || !int.TryParse(ends[0], out first)) throw new ArgumentException("页码格式示例：1-3,4-6,7；留空逐页拆分。");
            last = first;
            if (ends.Length == 2 && !int.TryParse(ends[1], out last)) throw new ArgumentException("页码范围不完整。");
            if (first < 1 || last < first || last > count) throw new ArgumentException("页码超出范围或起止顺序有误，总页数：" + count);
            result.Add(Enumerable.Range(first, last - first + 1).ToArray());
        }
        return result;
    }
    // Write beside the destination, then move without overwrite. Never truncate an existing file.
    private static string Save(PdfDocument doc, string output, string basis)
    {
        CheckCancel(); string temp = Path.Combine(output, ".pdf-" + Guid.NewGuid().ToString("N") + ".tmp");
        try { doc.Save(temp); CheckCancel(); return Publish(temp, output, basis); }
        finally { Cleanup(() => { if (File.Exists(temp)) File.Delete(temp); }, "清理临时文件 " + temp); }
    }
    private static string Publish(string temp, string output, string basis)
    {
        for (int i = 0; i < 10000; i++)
        {
            string path = Path.Combine(output, basis + (i == 0 ? "" : " (" + i + ")") + ".pdf");
            if (File.Exists(path) || Directory.Exists(path)) continue;
            try { File.Move(temp, path); return path; }
            catch (IOException) { if (!File.Exists(path) && !Directory.Exists(path)) throw; }
        }
        throw new IOException("同名输出文件过多，请更换输出目录。");
    }
    private static PdfDocument ImagePdf(string path)
    {
        var doc = new PdfDocument();
        try
        {
            using (var image = Image.FromFile(path))
            {
                // TIFF pages are preserved; animated GIF is treated as a single still image.
                bool tiff = image.RawFormat.Guid == ImageFormat.Tiff.Guid;
                int count = tiff ? image.GetFrameCount(FrameDimension.Page) : 1;
                for (int i = 0; i < count; i++)
                {
                    CheckCancel(); if (tiff) image.SelectActiveFrame(FrameDimension.Page, i);
                    using (var bitmap = new Bitmap(image.Width, image.Height))
                    {
                        using (var drawing = Graphics.FromImage(bitmap)) { drawing.Clear(Color.White); drawing.DrawImage(image, 0, 0, image.Width, image.Height); }
                        using (var stream = new MemoryStream())
                        {
                            bitmap.Save(stream, ImageFormat.Png); stream.Position = 0;
                            using (var source = XImage.FromStream(stream))
                            {
                                var page = doc.AddPage(); bool landscape = image.Width > image.Height;
                                page.Width = XUnit.FromPoint(landscape ? 841.89 : 595.28); page.Height = XUnit.FromPoint(landscape ? 595.28 : 841.89);
                                double scale = Math.Min((page.Width.Point - 36) / image.Width, (page.Height.Point - 36) / image.Height);
                                double w = image.Width * scale, h = image.Height * scale;
                                using (var g = XGraphics.FromPdfPage(page)) g.DrawImage(source, (page.Width.Point - w) / 2, (page.Height.Point - h) / 2, w, h);
                            }
                        }
                    }
                }
            }
            return doc;
        }
        catch { doc.Dispose(); throw; }
    }
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    private static string OfficePdf(string path, string output, string basis)
    {
        string temp = Path.Combine(output, ".pdf-" + Guid.NewGuid().ToString("N") + ".pdf");
        try
        {
            ExportOfficePdf(path, temp); CheckCancel();
            using (var check = OpenPdf(temp, PdfDocumentOpenMode.Import)) { }
            return Publish(temp, output, basis);
        }
        finally { Cleanup(() => { if (File.Exists(temp)) File.Delete(temp); }, "清理转换临时文件 " + temp); }
    }
    private static bool CanOpenOfficeSource(bool word, uint pid, bool newProcess, int documentCount)
    {
        // WPS Writer can have no application HWND before opening its first document.
        // An empty, unidentified Writer may own a document, never the application process.
        return documentCount == 0 && (newProcess || (word && pid == 0));
    }
    private static void ExportOfficePdf(string path, string temp)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        bool word = new[] { ".doc", ".docx", ".rtf", ".odt", ".wps" }.Contains(ext);
        if (!word && !new[] { ".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".et" }.Contains(ext)) throw new ArgumentException("不支持的转换格式：" + ext);
        object app = null, items = null, doc = null, previous = null; bool ownProcess = false, securitySet = false; object security = null;
        try
        {
            var existing = new HashSet<int>(Process.GetProcesses().Select(p => { int id = p.Id; p.Dispose(); return id; }));
            foreach (string id in word ? new[] { "KWPS.Application", "WPS.Application", "Word.Application" } : new[] { "KET.Application", "et.Application", "Excel.Application" })
            {
                Type type = Type.GetTypeFromProgID(id); if (type == null) continue;
                try { Log("create-host " + id); app = Activator.CreateInstance(type); if (app != null) break; } catch (COMException) { }
            }
            if (app == null) throw new InvalidOperationException("未找到可用的 WPS／Office，无法导出此格式。");
            dynamic a = app; uint pid = 0;
            try { GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(a.Hwnd)), out pid); ownProcess = pid != 0 && !existing.Contains((int)pid); }
            catch (Exception error) { Log("application-hwnd unavailable " + Details(error)); }
            Log("host pid=" + pid + " newProcess=" + ownProcess);
            items = word ? (object)a.Documents : (object)a.Workbooks;
            int documentCount = Convert.ToInt32(((dynamic)items).Count);
            if (!CanOpenOfficeSource(word, pid, ownProcess, documentCount))
            {
                ownProcess = false;
                throw new SourceOpenException("批次已暂停：转换宿主已有文档或无法安全识别；不会打开、导出或退出该宿主。");
            }
            if (!ownProcess) Log("empty-writer without application HWND; document-only ownership, application Quit disabled");
            for (int i = 1; i <= Convert.ToInt32(((dynamic)items).Count); i++)
            {
                object opened = ((dynamic)items).Item(i);
                try { if (string.Equals(Convert.ToString(((dynamic)opened).FullName), path, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("来源正在 WPS／Office 中打开，请先关闭该文件再转换。"); }
                finally { Release(opened); }
            }
            // Fail closed if macros cannot be disabled; restore the setting even on export failure.
            Log("open-source " + path);
            security = a.AutomationSecurity; securitySet = true; a.AutomationSecurity = 3;
            if (word) doc = ((dynamic)items).Open(FileName: path, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false, PasswordDocument: "", WritePasswordDocument: "", Visible: false, NoEncodingDialog: true);
            else doc = ((dynamic)items).Open(Filename: path, UpdateLinks: 0, ReadOnly: true, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, AddToMru: false);
            if (word)
            {
                object window = null;
                try
                {
                    window = ((dynamic)doc).ActiveWindow;
                    GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd)), out pid);
                    Log("document-window pid=" + pid);
                    if (pid != 0 && existing.Contains((int)pid))
                    {
                        ownProcess = false;
                        throw new SourceOpenException("批次已暂停：来源窗口属于已有宿主，不继续导出或退出该宿主。");
                    }
                }
                catch (SourceOpenException) { throw; }
                catch (Exception error) { Log("document-hwnd unavailable " + Details(error)); }
                finally { Release(window); }
            }
            CheckCancel();
            Log("export-source " + path);
            if (word) ((dynamic)doc).ExportAsFixedFormat(temp, 17, false);
            else ((dynamic)doc).ExportAsFixedFormat(0, temp, 0, true, false, Type.Missing, Type.Missing, false);
            CheckCancel();
        }
        finally
        {
            Cleanup(() => { if (doc != null) { if (word) ((dynamic)doc).Close(0); else ((dynamic)doc).Close(false); } }, "关闭本批次来源 " + path);
            Cleanup(() => Release(doc), "释放来源引用 " + path);
            Cleanup(() => Release(items), "释放文档集合 " + path);
            Cleanup(() => { if (app != null && securitySet) ((dynamic)app).AutomationSecurity = security; }, "恢复宏设置 " + path);
            Cleanup(() => { if (previous != null) ((dynamic)previous).Activate(); }, "恢复原活动文档 " + path);
            Cleanup(() => Release(previous), "释放原文档引用 " + path);
            Cleanup(() => { if (ownProcess && app != null)
                {
                    object remaining = null;
                    try
                    {
                        remaining = word ? (object)((dynamic)app).Documents : (object)((dynamic)app).Workbooks;
                        if (Convert.ToInt32(((dynamic)remaining).Count) == 0) ((dynamic)app).Quit();
                        else cleanupWarnings.Add("宿主中仍有文档，保留进程，不调用 Quit。");
                    }
                    finally { Release(remaining); }
                } }, "退出自建宿主 " + path);
            Cleanup(() => Release(app), "释放宿主引用 " + path);
        }
    }
}
