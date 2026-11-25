using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

// Opt-in production-session regression on copies. Works with the preserved old
// DLL as a red test; no ribbon click or installation is performed.
internal static class NameLinksLiveProbe
{
    private static readonly BindingFlags Hidden = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
    private static Assembly assembly;
    private static int checks;
    private static object Get(object value, string name) { return value.GetType().GetField(name, Hidden).GetValue(value); }
    private static object Call(object value, string name, params object[] args) { return value.GetType().GetMethod(name, Hidden).Invoke(value, args); }
    private static object Package(string name, params object[] args) { return assembly.GetType("Excel_Iterate.XlsxLinkPackage").GetMethod(name, Hidden).Invoke(null, args); }
    private static string Hash(string path) { return (string)Package("Hash", path); }
    private static List<object> Links(string path) { return ((IEnumerable)Package("Scan", path)).Cast<object>().ToList(); }
    private static bool First(string path) { return path != null && path.EndsWith("\\11资产1", StringComparison.OrdinalIgnoreCase); }
    private static bool Second(string path) { return path != null && path.Contains("私有云") && path.EndsWith("\\Worksheet in 5440 Stock Breakdown", StringComparison.OrdinalIgnoreCase); }
    private static bool AddressNameSource(string path) { return path != null && path.EndsWith("\\审核2.XLS", StringComparison.OrdinalIgnoreCase); }
    private static void Check(bool value, string text) { if (!value) throw new InvalidOperationException(text); checks++; Console.WriteLine("PASS: " + text); }
    private static XmlDocument Xml(ZipArchive zip, string part)
    { var doc = new XmlDocument { PreserveWhitespace = true, XmlResolver = null }; using (var stream = zip.GetEntry(part).Open()) doc.Load(stream); return doc; }
    private static string Style(XmlDocument styles, string id, bool baseStyle = false)
    {
        var root = styles.DocumentElement.ChildNodes.OfType<XmlElement>().ToDictionary(n => n.LocalName);
        int index = string.IsNullOrEmpty(id) ? 0 : int.Parse(id);
        var xf = (XmlElement)root[baseStyle ? "cellStyleXfs" : "cellXfs"].ChildNodes.OfType<XmlElement>().ElementAt(index).CloneNode(true);
        foreach (string attr in new[] { "fontId", "fillId", "borderId" })
            if (xf.HasAttribute(attr)) xf.SetAttribute(attr, root[attr == "fontId" ? "fonts" : attr == "fillId" ? "fills" : "borders"].ChildNodes.OfType<XmlElement>().ElementAt(int.Parse(xf.GetAttribute(attr))).OuterXml);
        if (!baseStyle && xf.HasAttribute("xfId")) xf.SetAttribute("xfId", Style(styles, xf.GetAttribute("xfId"), true));
        XmlElement formats;
        if (root.TryGetValue("numFmts", out formats))
        {
            var custom = formats.ChildNodes.OfType<XmlElement>().FirstOrDefault(n => n.GetAttribute("numFmtId") == xf.GetAttribute("numFmtId"));
            if (custom != null) xf.SetAttribute("numFmtId", custom.GetAttribute("formatCode"));
        }
        return xf.OuterXml;
    }
    private static string Signature(string path, bool skipReport)
    {
        Type type = assembly.GetType("Excel_Iterate.LinkNameAnalysis", false);
        object analysis = type == null ? null : type.GetMethod("Read", Hidden).Invoke(null, new object[] { path });
        Func<string, string> semantic = text => analysis == null ? text : (string)Call(analysis, "Semantic", text);
        var result = new StringBuilder();
        using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
        {
            var strings = new List<string>();
            if (zip.GetEntry("xl/sharedStrings.xml") != null)
                foreach (XmlElement si in Xml(zip, "xl/sharedStrings.xml").DocumentElement.ChildNodes.OfType<XmlElement>()) strings.Add(si.OuterXml);
            var styles = Xml(zip, "xl/styles.xml"); var usedStyles = new Dictionary<string, string>();
            Func<string, string> style = id => { id = string.IsNullOrEmpty(id) ? "0" : id; string text; if (!usedStyles.TryGetValue(id, out text)) using (var sha = SHA256.Create()) usedStyles[id] = text = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Style(styles, id)))); return text; };
            string reportPart = null;
            if (skipReport)
            {
                var workbook = Xml(zip, "xl/workbook.xml"); var rels = Xml(zip, "xl/_rels/workbook.xml.rels");
                var report = workbook.SelectNodes("//*[local-name()='sheet']").Cast<XmlElement>().Single(n => n.GetAttribute("name") == "外部链接");
                string id = report.Attributes.Cast<XmlAttribute>().Single(a => a.LocalName == "id").Value;
                string target = rels.DocumentElement.ChildNodes.OfType<XmlElement>().Single(n => n.GetAttribute("Id") == id).GetAttribute("Target");
                reportPart = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
            }
            foreach (var part in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).OrderBy(e => e.FullName, StringComparer.Ordinal))
            {
                if (part.FullName == reportPart) continue;
                result.Append(part.FullName);
                foreach (XmlElement cell in Xml(zip, part.FullName).SelectNodes("//*[local-name()='c']"))
                {
                    result.Append('|').Append(cell.GetAttribute("r")).Append(':').Append(style(cell.GetAttribute("s"))).Append(':').Append(cell.GetAttribute("t"));
                    foreach (XmlElement node in cell.ChildNodes.OfType<XmlElement>())
                    {
                        string text = node.InnerText;
                        if (node.LocalName == "f") text = semantic(text);
                        else if (node.LocalName == "v" && cell.GetAttribute("t") == "s") text = strings[int.Parse(text)];
                        result.Append('/').Append(node.LocalName).Append(':').Append(text);
                    }
                }
                var sheet = Xml(zip, part.FullName);
                foreach (XmlElement row in sheet.SelectNodes("//*[local-name()='row']"))
                { var attrs = row.Attributes.Cast<XmlAttribute>().Select(a => a.LocalName + "=" + (a.LocalName == "s" ? style(a.Value) : a.Value)); result.Append("row:").Append(string.Join(";", attrs)); }
                foreach (XmlElement node in sheet.DocumentElement.ChildNodes.OfType<XmlElement>().Where(n => new[] { "cols", "mergeCells", "sheetFormatPr" }.Contains(n.LocalName)))
                {
                    var copy = (XmlElement)node.CloneNode(true);
                    foreach (XmlElement col in copy.SelectNodes(".//*[local-name()='col']")) if (col.HasAttribute("style")) col.SetAttribute("style", style(col.GetAttribute("style")));
                    result.Append(copy.OuterXml);
                }
            }
            foreach (XmlElement color in styles.DocumentElement.ChildNodes.OfType<XmlElement>().Where(n => n.LocalName == "colors")) result.Append(color.OuterXml);
            if (zip.GetEntry("xl/theme/theme1.xml") != null) result.Append(Xml(zip, "xl/theme/theme1.xml").OuterXml);
        }
        return result.ToString();
    }
    private static void Track(ProbeSession session)
    { foreach (string file in WpsShared.SourcePaths.GetFiles(session.DirectoryPath)) if (Path.GetFileName(file).StartsWith(".EI_", StringComparison.Ordinal)) session.TrackFile(file); }
    private static void Case(ProbeSession session, object application, string input, int number, bool first, bool second, bool closed, bool sameReport = false, bool addressName = false)
    {
        using (var refs = session.References())
        {
            dynamic app = application, books = refs.Take(app.Workbooks);
            string target = session.CopyInput(input, "target-" + number + ".xlsx");
            object targetBook = books.Open(target, 0, false); bool targetClosed = false;
            session.Own(targetBook, () => { if (!targetClosed) ((dynamic)targetBook).Close(false); });
            ((dynamic)targetBook).Save();
            if (closed) { ((dynamic)targetBook).Close(false); targetClosed = true; }
            object receiver = sameReport ? targetBook : books.Add();
            if (!sameReport) { session.Own(receiver, () => ((dynamic)receiver).Close(false)); string report = session.Output("report-" + number + ".xlsx"); ((dynamic)receiver).SaveAs(report, 51); }
            ((dynamic)receiver).Activate();
            Type sessionType = assembly.GetType("Excel_Iterate.LinkAssistantSession", true);
            object engine = Activator.CreateInstance(sessionType, Hidden, null, new[] { application }, null);
            int count = (int)Call(engine, "Scan", (object)new[] { target });
            var rows = ((IEnumerable)Get(engine, "Links")).Cast<object>().ToList();
            var selected = rows.Where(r => (first && First((string)Get(r, "SourcePath"))) || (second && Second((string)Get(r, "SourcePath"))) || (addressName && AddressNameSource((string)Get(r, "SourcePath")))).ToList();
            Check(selected.Count == (first && second ? 2 : 1), "scan identifies the exact requested source paths");
            if (assembly.GetType("Excel_Iterate.LinkNameAnalysis", false) != null)
            {
                Check(selected.All(r => ((string)Get(r, "Status")).Contains("仅名称管理器定义引用")), "scan explains name-only links in the existing status column");
                Check(selected.All(r => ((string)Get(r, "Status")).Contains(AddressNameSource((string)Get(r, "SourcePath")) ? "（2 个名称）" : First((string)Get(r, "SourcePath")) ? "（5 个名称）" : "（44 个名称）")), "the sample's exact scoped name-definition counts are reported");
            }
            var beforeSources = new HashSet<string>(Links(target).Select(l => (string)Get(l, "SourcePath")), StringComparer.OrdinalIgnoreCase);
            dynamic sheets = refs.Take(((dynamic)receiver).Worksheets), sheet = refs.Take(sheets.Item("外部链接"));
            string command = (string)sessionType.GetField("BreakCommand", Hidden).GetRawConstantValue();
            for (int row = 2; row <= count + 1; row++)
            {
                dynamic old = refs.Take(sheet.Range["B" + row]); string path = Convert.ToString(old.Value2);
                if (selected.Any(r => string.Equals((string)Get(r, "SourcePath"), path, StringComparison.OrdinalIgnoreCase)))
                { dynamic edit = refs.Take(sheet.Range["C" + row]); edit.Value2 = command; }
            }
            ((dynamic)receiver).Save();
            string beforeContent = Signature(target, sameReport);
            Call(engine, "Execute", new Action<string>(Console.WriteLine));
            object batch = Get(engine, "Batch"); Track(session); Console.WriteLine(Get(batch, "UserLog"));
            Console.WriteLine("CASE=" + number + " closed=" + closed + " first=" + first + " second=" + second + " addressName=" + addressName + " STATE=" + Get(batch, "State"));
            var expected = new HashSet<string>(beforeSources, StringComparer.OrdinalIgnoreCase);
            foreach (object row in selected) expected.Remove((string)Get(row, "SourcePath"));
            Check(expected.SetEquals(new HashSet<string>(Links(target).Select(l => (string)Get(l, "SourcePath")), StringComparer.OrdinalIgnoreCase)), "selected name-only sources disappear and all other source paths remain");
            Check(Signature(target, sameReport) == beforeContent, "all data-sheet cell values, formulas, applied styles, row/column sizes and merged cells are preserved");
            if (!targetClosed) { ((dynamic)targetBook).Close(false); targetClosed = true; }
            object reopened = books.Open(target, 0, false); bool reopenedClosed = false;
            session.Own(reopened, () => { if (!reopenedClosed) ((dynamic)reopened).Close(false); });
            Check(expected.SetEquals(new HashSet<string>(Links(target).Select(l => (string)Get(l, "SourcePath")), StringComparer.OrdinalIgnoreCase)), "reopening does not recreate the broken sources");
            ((dynamic)reopened).Close(false); reopenedClosed = true;
            Check((string)Get(batch, "State") == "Completed" &&
                !WpsShared.SourcePaths.GetFiles(session.DirectoryPath, ".EI_LinkBackup_*.xlsx").Any(),
                "successful break leaves the result saved and clears the protection copy");
        }
    }
    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        if (args.Length < 1 || args.Length > 2 || (args.Length == 2 && args[1] != "--y-only")) { Console.WriteLine("Usage: NameLinksLiveProbe original.xlsx [--y-only] (only temporary copies are modified)"); return 2; }
        assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel_Iterate.dll"));
        string input = Path.GetFullPath(args[0]), before = Hash(input);
        using (var session = new ProbeSession("name-links-live"))
        {
            try
            {
                object app = session.Take(ProbeSession.Connect(new[] { "KET.Application", "Excel.Application" }, false)); if (app == null) return session.Finish(2);
                dynamic host = app;
                session.Override(() => (object)host.DisplayAlerts, v => host.DisplayAlerts = v, false, "DisplayAlerts");
                session.Override(() => (object)host.ScreenUpdating, v => host.ScreenUpdating = v, false, "ScreenUpdating");
                session.Override(() => (object)host.EnableEvents, v => host.EnableEvents = v, false, "EnableEvents");
                if (args.Length == 1)
                {
                    Case(session, app, input, 1, true, false, false); Case(session, app, input, 2, false, true, false);
                    Case(session, app, input, 3, true, true, false); Case(session, app, input, 4, true, true, true);
                    Case(session, app, input, 5, true, true, false, true);
                }
                Case(session, app, input, 6, false, false, false, false, true);
                Check(Hash(input) == before, "the supplied original workbook remains unchanged");
                Console.WriteLine("NameLinksLiveProbe: " + checks + " PASS"); return session.Finish(0);
            }
            catch (Exception error) { Track(session); session.Fail(error.ToString()); return session.Finish(1); }
        }
    }
}
