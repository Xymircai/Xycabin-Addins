using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using Excel_Iterate;

partial class WorkbookToolsProbe
{
    private static string Esc(string text) { return System.Security.SecurityElement.Escape(text); }
    private static void NamePackage(string file, string a, string b, string formula, string extraNames = "", string objectFormula = null, bool missing = false, string firstSheet = "S1")
    {
        const string s = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        using (var output = File.Create(file)) using (var zip = new ZipArchive(output, ZipArchiveMode.Create))
        {
            Part(zip, "xl/workbook.xml", "<workbook xmlns='" + s + "' xmlns:r='" + r + "'><sheets><sheet name='" + Esc(firstSheet) + "' r:id='s1'/><sheet name='S2' r:id='s2'/></sheets>"
                + "<externalReferences><externalReference r:id='a'/><externalReference r:id='b'/></externalReferences><definedNames>"
                + "<definedName name='Unused'>[1]Data!#REF!</definedName><definedName name='Same' localSheetId='0'>[1]Data!$A$1</definedName>"
                + "<definedName name='Same' localSheetId='1'>[2]Data!$B$1</definedName><definedName name='Keep'>S1!$A$8</definedName>" + extraNames + "</definedNames></workbook>");
            Part(zip, "xl/_rels/workbook.xml.rels", "<Relationships><Relationship Id='a' Target='externalLinks/externalLink9.xml'/><Relationship Id='b' Target='externalLinks/externalLink4.xml'/>"
                + "<Relationship Id='s1' Target='worksheets/sheet1.xml'/><Relationship Id='s2' Target='worksheets/sheet2.xml'/></Relationships>");
            foreach (int part in new[] { 9, 4 })
            {
                Part(zip, "xl/externalLinks/externalLink" + part + ".xml", "<externalLink xmlns='" + s + "' xmlns:r='" + r + "'><externalBook r:id='p'/></externalLink>");
                Part(zip, "xl/externalLinks/_rels/externalLink" + part + ".xml.rels", "<Relationships><Relationship Id='p' Type='" + (missing && part == 4 ? "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing" : r + "/externalLinkPath") + "' Target='" + Esc(part == 9 ? a : b) + "' TargetMode='External'/></Relationships>");
            }
            Part(zip, "xl/worksheets/sheet1.xml", "<worksheet xmlns='" + s + "'><sheetData><row r='1'><c r='A1'><f>" + Esc(formula) + "</f><v>17</v></c></row></sheetData></worksheet>");
            Part(zip, "xl/worksheets/sheet2.xml", "<worksheet xmlns='" + s + "'><sheetData><row r='1'><c r='A1'><f>Same</f><v>9</v></c></row></sheetData></worksheet>");
            if (objectFormula != null) Part(zip, "xl/charts/chart1.xml", "<chart><f>" + Esc(objectFormula) + "</f></chart>");
        }
    }
    private static void NameLinkChecks(string folder, string a, string b, string report)
    {
        AddressNameChecks(folder, a, b, report);
        MissingPathReportChecks(folder, a, b, report);
        string target = Path.Combine(folder, "names.xlsx");
        NamePackage(target, a, b, "\"Unused\"&UnusedMore");
        var analysis = LinkNameAnalysis.Read(target); var plan = analysis.Plan(a);
        Check(plan.Count == 2 && plan.Single(n => n.Name == "Same").Sheet == "S1", "external-reference order resolves source even when part numbers are 9 and 4");
        Reject(() => analysis.Plan(b), "a local name used on S2 is not deleted");
        Check(plan.All(n => n.Sources.Count == 1 && n.Sources.Contains(a)), "only the exact selected source's scoped names are planned");
        var described = XlsxLinkPackage.Scan(target); LinkNameAnalysis.Describe(target, described);
        Check(described.First(n => n.SourcePath == a).Status == "仅名称管理器定义引用（2 个名称）", "scan explains name-only links and their count");
        Check(analysis.Semantic("='" + Path.GetDirectoryName(a) + Path.DirectorySeparatorChar + "[" + Path.GetFileName(a) + "]Data'!#REF!") == analysis.Semantic("[1]Data!#REF!"), "native full-path and package index forms identify the same source");
        foreach (string use in new[] { "Unused", "SUM(Unused)", "S1!Same", "'S1'!Same", "INDIRECT(\"Unused\")", "EVALUATE(A1)" })
        { NamePackage(target, a, b, use); var value = LinkNameAnalysis.Read(target); Reject(() => value.Plan(a), "used or dynamic name is refused: " + use); }
        NamePackage(target, a, b, "S2!Same"); Check(LinkNameAnalysis.Read(target).Plan(a).Count == 2, "explicit other-sheet scope does not use the selected sheet's same name");
        foreach (string value in new[] { "Table1[Unused]", "Table1[[#Headers],[Unused]]", "[2]!Unused" })
        { NamePackage(target, a, b, value); Check(LinkNameAnalysis.Read(target).Plan(a).Count == 2, "structured columns and another workbook's name are not local name uses: " + value); }
        NamePackage(target, a, b, "17", "<definedName name='Wrapper'>Unused</definedName>"); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "other defined-name dependencies are protected");
        NamePackage(target, a, b, "17", "", "S1!Same"); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "chart/object formula dependencies are protected");
        NamePackage(target, a, b, "17", "<definedName name='Mixed'>[1]Data!A1+[2]Data!A1</definedName>"); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "mixed-source names are refused");
        NamePackage(target, a, b, "[1]Data!A1+[2]Data!B1"); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "one cell using selected and other sources is refused before native conversion");
        NamePackage(target, a, b, "17", "<definedName name='Unused'>[1]Data!A2</definedName>"); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "duplicate identity in the same scope is refused");
        NamePackage(target, a, b, "17", "", null, true);
        Check(LinkNameAnalysis.Read(target).Plan(a).Count == 2, "xlPathMissing metadata for another source does not hide valid names");
        NamePackage(target, a, b, "17");
        using (var stream = File.Open(target, FileMode.Open, FileAccess.ReadWrite)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Update)) Part(zip, "xl/embeddings/opaque.bin", "unreadable object");
        Reject(() => LinkNameAnalysis.Read(target).Plan(a), "opaque embedded objects prevent unproven name cleanup");
        string sameA = Path.Combine(folder, "one", "same.xlsx"), sameB = Path.Combine(folder, "two", "same.xlsx");
        NamePackage(target, sameA, sameB, "17");
        Check(LinkNameAnalysis.Read(target).Plan(sameA).All(n => n.Sources.Contains(sameA) && !n.Sources.Contains(sameB)), "same leaf filename in another directory is not the selected source");
        NamePackage(target, a, b, "17");
        using (var stream = File.Open(target, FileMode.Open, FileAccess.ReadWrite)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Update)) zip.GetEntry("xl/externalLinks/externalLink9.xml").Delete();
        Reject(() => LinkNameAnalysis.Read(target), "required package read failure never becomes an empty safe plan");
        NamePackage(target, a, b, "17", "<definedName name='Overflow'>[99999999999999999999]Data!A1</definedName>");
        Reject(() => LinkNameAnalysis.Read(target), "overflowing numeric source indices cannot masquerade as filenames");
        NamePackage(target, a, b, "17", "", null, false, "A'B");
        var quoted = LinkNameAnalysis.Read(target); var wrongScope = new NameAwareBook { FullName = target, StripScopeQuote = true };
        using (var refs = new ComBag()) Reject(() => quoted.NativeNames(wrongScope, quoted.Plan(a), refs), "a sheet named A-apostrophe-B never aliases a different sheet AB");

        foreach (string mode in new[] { "ok", "read", "identity", "formula", "delete", "save", "used", "edit" })
        {
            NamePackage(target, a, b, mode == "used" ? "Unused" : "17"); byte[] before = File.ReadAllBytes(target);
            var session = JournalSession(report, new[] { target }, a, "  断开链接  ", true);
            var app = (LinkFakeApp)session.App; var previous = app.Workbooks.Items[1];
            var book = new NameAwareBook { FullName = target, Owner = app.Workbooks, Sheet = previous.Sheet,
                FailNameRead = mode == "read", WrongNameIdentity = mode == "identity", WrongNameFormula = mode == "formula",
                FailNameDeleteAt = mode == "delete" ? 2 : 0, FailNameSave = mode == "save", EditOnNameRead = mode == "edit" };
            app.Workbooks.Items[1] = book; session.Execute(null);
            if (mode == "ok")
            {
                Check(book.NameDeletes == 2 && session.Batch.Changes[0].Applied, "production session deletes both unused scoped names and commits native break");
                Check(XlsxLinkPackage.Scan(target).Select(l => l.SourcePath).SequenceEqual(new[] { b }), "native model preserves the other source after index renumbering");
                Check(LinkNameAnalysis.Read(target).Names.Count == 2 && LinkNameAnalysis.Read(target).Names.Any(n => n.Name == "Same" && n.Sheet == "S2"), "other scope's same name and internal names survive");
            }
            else
            {
                Check(!session.Batch.Changes[0].Applied && File.ReadAllBytes(target).SequenceEqual(before), "name failure never falsely commits original bytes: " + mode);
                if (mode == "used") Check(session.Batch.Checkpoints.Count == 0 && book.NameDeletes == 0, "unsafe dependency is rejected before checkpoint or deletion");
                if (mode == "read" || mode == "identity" || mode == "formula") Check(book.NameDeletes == 0, "all native name handles are verified before first deletion: " + mode);
                if (mode == "edit") Check(book.NameDeletes == 0 && !book.Saved && book.SaveCount == 0, "new edits during native name reads are not saved or overwritten");
                if (mode == "delete" || mode == "save")
                    Check(session.Batch.State == "UndoFailed" && File.Exists(session.Batch.Checkpoints[0].Backup),
                        "unsafe automatic rollback retains its original copy: " + mode);
                Check(!session.Links[0].Status.Contains("原链接未改变"), "failure status makes no unsupported unchanged claim: " + mode);
            }
            if (mode == "ok" || session.Batch.State == "Undone")
                Check(session.Batch.Checkpoints.Count == 0, "successful save or automatic rollback clears the temporary protection: " + mode);
        }
        foreach (string command in new[] { "请断开该链接", "断开链接！", "请断开链接", "断开 链接" })
        {
            NamePackage(target, a, b, "17"); var session = JournalSession(report, new[] { target }, a, command, true); string hash = XlsxLinkPackage.Hash(target);
            Reject(() => session.ReadMappings(), "only the new exact command is accepted: " + command);
            Check(!session.Links[0].BreakRequested && XlsxLinkPackage.Hash(target) == hash, "rejected command never requests break or alters the workbook");
            if (command == "请断开该链接")
            { string message = null; try { session.ReadMappings(); } catch (InvalidOperationException error) { message = error.Message; } Check(message != null && message.Contains("断链指令已改为「断开链接」"), "old command gives a specific migration instruction before any write"); }
        }
        NamePackage(target, a, b, "17"); byte[] noopBefore = File.ReadAllBytes(target);
        var noop = JournalSession(report, new[] { target }, a, "断开链接", true); var noopApp = (LinkFakeApp)noop.App;
        var noopBook = new NameAwareBook { FullName = target, Owner = noopApp.Workbooks, Sheet = noopApp.Workbooks.Items[1].Sheet, NoNameDelete = true }; noopApp.Workbooks.Items[1] = noopBook;
        noop.Execute(null);
        Check(!noop.Batch.Changes[0].Applied && XlsxLinkPackage.Scan(target).Any(l => l.SourcePath == a) &&
            File.ReadAllBytes(target).SequenceEqual(noopBefore), "silent no-op name deletion is caught and leaves the original bytes");
    }
    private static void MissingPathReportChecks(string folder, string source, string replacement, string report)
    {
        string workbook = Path.Combine(folder, "missing-path-report.xlsx");
        foreach (string target in new[] { XlsxLinkPackage.WpsFileUri(replacement), "未定位来源.xlsx" })
        {
            NamePackage(workbook, source, target, "[2]Data!A1", missing: true);
            string before = XlsxLinkPackage.Hash(workbook);
            var session = JournalSession(report, new[] { workbook }, source, replacement, true);
            Check(session.Scan(new[] { workbook }) == 2, "report includes ordinary and xlPathMissing relationships");
            var app = (LinkFakeApp)session.App;
            var reopened = new LinkAssistantSession(app);
            var missing = reopened.Links.Single(link => link.IsMissingPath);
            var sheet = ((LinkFakeBook)app.ActiveWorkbook).Sheet;
            int missingRow = reopened.Links.IndexOf(missing) + 1;
            int normalRow = reopened.Links.FindIndex(link => link.SourcePath == source) + 1;
            Check(missing.SourcePath == null && (string)sheet.Values[missingRow, 1] == target
                && ((string)sheet.Values[missingRow, 3]).Contains("可填写新链接"), "missing source keeps its original target and report status after reopen");
            sheet.Values[normalRow, 2] = report;
            sheet.Values[missingRow, 2] = LinkAssistantSession.BreakCommand;
            Reject(() => reopened.Execute(null), "editing an unverified source is stopped before applying any mapping");
            Check(XlsxLinkPackage.Hash(workbook) == before, "rejected missing-source edit leaves the workbook intact");
            sheet.Values[missingRow, 2] = "";
            reopened.Execute(null);
            Check(XlsxLinkPackage.Scan(workbook).Single().SourcePath == report
                && XlsxLinkPackage.ScanReport(workbook).Single(link => link.IsMissingPath).RawTarget == target,
                "a blank missing-source row allows ordinary replacement and preserves the missing relationship");
            NamePackage(workbook, source, target, "[2]Data!A1", missing: true);
            reopened.Scan(new[] { workbook });
            session = reopened;
            missingRow = session.Links.FindIndex(link => link.IsMissingPath) + 1;
            sheet.Values[missingRow, 2] = Path.GetDirectoryName(replacement) + @"\\[" + Path.GetFileName(replacement) + "]";
            var book = app.Workbooks.Items[1];
            Check(session.Execute(null).Contains("成功1条") && book.CloseCount == 0,
                "native replacement accepts formula-style paths and keeps the open workbook open");
            Check(XlsxLinkPackage.Scan(workbook).Select(l => l.SourcePath).SequenceEqual(new[] { source, replacement })
                && !XlsxLinkPackage.ScanReport(workbook).Any(l => l.IsMissingPath),
                "only the selected missing path is rebound to an ordinary external source");
            Check(session.Batch.Checkpoints.Count == 0, "missing-source replacement clears the temporary protection after success");
        }
        NamePackage(workbook, Path.Combine(folder, "另一目录", "同名.xlsx"), "同名.xlsx", "[2]Data!A1", missing: true);
        var links = XlsxLinkPackage.ScanReportParts(workbook);
        Reject(() => LinkAssistantSession.MatchMissingHostName(new[] { "同名.xlsx" }, links.Single(l => l.IsMissingPath), links),
            "a host basename cannot select a missing relationship when another saved source uses the same filename");
        const string normalizedTarget = "保存后改写.xlsx";
        foreach (string target in new[] { XlsxLinkPackage.WpsFileUri(Path.Combine(folder, normalizedTarget)), normalizedTarget })
        {
            NamePackage(workbook, source, target, "[2]Data!A1", missing: true);
            var session = JournalSession(report, new[] { workbook }, source, replacement, true);
            session.Scan(new[] { workbook });
            var app = (LinkFakeApp)session.App;
            int row = session.Links.FindIndex(link => link.IsMissingPath) + 1;
            var sheet = ((LinkFakeBook)app.ActiveWorkbook).Sheet;
            sheet.Values[row, 2] = replacement;
            NamePackage(workbook, source, normalizedTarget, "[2]Data!A1");
            string before = XlsxLinkPackage.Hash(workbook);
            StopAfterSave(session, () =>
            {
                Check(session.Batch.Changes.Single().Applied,
                    "equivalent WPS save normalization preserves the report row and applies its replacement");
                app.Workbooks.Items[1].Close(false);
            });
            Check(XlsxLinkPackage.Hash(workbook) == before && session.Batch.Checkpoints.Count == 0,
                "automatic rollback restores the actual saved pre-replacement identity");
        }
    }
    private static void AddressNameChecks(string folder, string a, string b, string report)
    {
        string target = Path.Combine(folder, "address-names.xlsx");
        string y = "<definedName name='Y' localSheetId='0'>[1]Data!#REF!</definedName>";
        foreach (string formula in new[] { "Y$5", "$Y$5", "$Y5", "Y5", "Y$5:AA$9", "$Y$5:$AA$9", "Y:Y", "$Y:$Y", "Y:$Y", "$Y:Y", "S1!Y$5", "'S1'!$Y$5", "SUM(S1!Y$5,'S1'!$Y:$Y)", "[2]Data!$Y$5" })
        {
            NamePackage(target, a, b, formula, y);
            Check(LinkNameAnalysis.Read(target).Plan(a).Count == 3, "A1 cell or whole-column reference is not a defined-name Y use: " + formula);
        }
        foreach (string formula in new[] { "Y", "SUM(Y)", "S1!Y", "'S1'!Y", "Y+Y$5", "$Y$5+Y" })
        { NamePackage(target, a, b, formula, y); Reject(() => LinkNameAnalysis.Read(target).Plan(a), "real defined-name Y references remain protected: " + formula); }
        foreach (string column in new[] { "A", "AB", "XFD" })
        {
            NamePackage(target, a, b, "$" + column + "$5", "<definedName name='" + column + "' localSheetId='0'>[1]Data!#REF!</definedName>");
            Check(LinkNameAnalysis.Read(target).Plan(a).Count == 3, "absolute A1 references do not use same-letter names: " + column);
        }
        foreach (string name in new[] { "Y0", "Y1048577", "XFE5" })
        {
            NamePackage(target, a, b, name, "<definedName name='" + name + "' localSheetId='0'>[1]Data!#REF!</definedName>");
            Reject(() => LinkNameAnalysis.Read(target).Plan(a), "names outside valid A1 bounds remain protected: " + name);
        }
        string filter = "<definedName name='Y'>[1]Data!#REF!</definedName><definedName name='_xlnm._FilterDatabase' localSheetId='1'>S2!$A$8:$Y$271</definedName>";
        NamePackage(target, a, b, "Y$5", y + filter);
        Check(LinkNameAnalysis.Read(target).Plan(a).Count == 4, "filter database absolute ranges do not use global or local Y names");
        var engine = JournalSession(report, new[] { target }, a, "断开链接", true);
        var app = (LinkFakeApp)engine.App; var book = new NameAwareBook { FullName = target, Owner = app.Workbooks, Sheet = app.Workbooks.Items[1].Sheet }; app.Workbooks.Items[1] = book;
        engine.Execute(null);
        Check(engine.Batch.Changes[0].Applied && book.NameDeletes == 4, "production session removes only unused names after excluding cell/range address spans");
        var after = LinkNameAnalysis.Read(target);
        Check(after.Names.Any(n => n.Name == "_xlnm._FilterDatabase" && n.Formula == "S2!$A$8:$Y$271"), "filter-name definition itself is preserved");
        Check(engine.Batch.Checkpoints.Count == 0 && book.CloseCount == 0,
            "address-name break saves without closing or leaving a protection copy");
    }
}
public partial class LinkFakeBook
{
    private IEnumerable<ExternalLinkInfo> MissingHostLinks()
    {
        return XlsxLinkPackage.ScanReportParts(FullName).Where(link => link.IsMissingPath);
    }
    private static string NativeMissingName(ExternalLinkInfo link)
    {
        return XlsxLinkPackage.AbsoluteFileTarget(link.RawTarget) ?? Uri.UnescapeDataString(link.RawTarget);
    }
    private bool ChangeMissingHostLink(string name, string replacement)
    {
        bool missing = !File.Exists(replacement);
        var source = XlsxLinkPackage.ScanReportParts(FullName).SingleOrDefault(link =>
            link.IsMissingPath ? NativeMissingName(link) == name : missing && link.SourcePath == name);
        if (source == null) return false;
        using (var stream = File.Open(FullName, FileMode.Open, FileAccess.ReadWrite))
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update))
        {
            var entry = zip.GetEntry(source.Part);
            XmlDocument xml;
            using (var input = entry.Open()) xml = XlsxLinkPackage.Load(input);
            var relationship = xml.DocumentElement.ChildNodes.OfType<XmlElement>().Single(item => item.GetAttribute("Id") == source.Relationship);
            relationship.SetAttribute("Type", missing
                ? "http://schemas.microsoft.com/office/2006/relationships" + XlsxLinkPackage.MissingPathSuffix
                : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath");
            relationship.SetAttribute("Target", XlsxLinkPackage.WpsFileUri(replacement));
            entry.Delete();
            using (var output = zip.CreateEntry(source.Part).Open()) xml.Save(output);
        }
        return true;
    }
}

// Native WPS model for named links: BreakLink retains a source while a name
// still points to it. XML rewriting below belongs only to the fake host.
public sealed class NameAwareBook : LinkFakeBook
{
    public byte[] Pending;
    public bool FailNameRead, FailNameSave, WrongNameIdentity, WrongNameFormula, NoNameDelete, EditOnNameRead, StripScopeQuote;
    public int FailNameDeleteAt, NameDeletes;
    public new object Names { get { return new NameAwareNames { Book = this }; } }
    internal byte[] Bytes { get { return Pending ?? File.ReadAllBytes(FullName); } }
    internal static XmlDocument Doc(ZipArchive zip, string part)
    { var doc = new XmlDocument { XmlResolver = null }; using (var stream = zip.GetEntry(part).Open()) doc.Load(stream); return doc; }
    internal static void Put(ZipArchive zip, string part, XmlDocument document)
    { var old = zip.GetEntry(part); old.Delete(); using (var stream = zip.CreateEntry(part).Open()) document.Save(stream); }
    internal static string Identity(XmlElement name, XmlDocument workbook)
    {
        string scope = name.GetAttribute("localSheetId");
        var sheets = workbook.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']").Cast<XmlElement>().ToList();
        return scope.Length == 0 ? name.GetAttribute("name") : "'" + sheets[int.Parse(scope)].GetAttribute("name").Replace("'", "''") + "'!" + name.GetAttribute("name");
    }
    internal void Edit(Action<ZipArchive> action)
    {
        using (var memory = new MemoryStream())
        { byte[] bytes = Bytes; memory.Write(bytes, 0, bytes.Length); using (var zip = new ZipArchive(memory, ZipArchiveMode.Update, true)) action(zip); Pending = memory.ToArray(); }
        Saved = false;
    }
    public new void Save()
    {
        if (FailNameSave) throw new IOException("injected name Save failure");
        if (Pending != null) { File.WriteAllBytes(FullName, Pending); Pending = null; }
        base.Save();
    }
    public new void BreakLink(string source, int type)
    {
        Edit(zip =>
        {
            var workbook = Doc(zip, "xl/workbook.xml"); var rels = Doc(zip, "xl/_rels/workbook.xml.rels");
            var refs = workbook.SelectNodes("//*[local-name()='externalReference']").Cast<XmlElement>().ToList();
            int index = 0; string part = null, relPart = null, relationId = null;
            for (int i = 0; i < refs.Count; i++)
            {
                string id = refs[i].Attributes.Cast<XmlAttribute>().Single(x => x.LocalName == "id").Value;
                string relative = rels.DocumentElement.ChildNodes.OfType<XmlElement>().Single(r => r.GetAttribute("Id") == id).GetAttribute("Target");
                string full = "xl/" + relative, rp = "xl/externalLinks/_rels/" + Path.GetFileName(relative) + ".rels";
                var link = Doc(zip, rp).DocumentElement.ChildNodes.OfType<XmlElement>().Single();
                if (string.Equals(XlsxLinkPackage.Resolve(FullName, link.GetAttribute("Target")), source, StringComparison.OrdinalIgnoreCase)) { index = i + 1; part = full; relPart = rp; relationId = id; break; }
            }
            if (index == 0) throw new IOException("fake source missing");
            if (workbook.SelectNodes("//*[local-name()='definedName']").Cast<XmlElement>().Any(n => n.InnerText.Contains("[" + index + "]"))) return;
            refs[index - 1].ParentNode.RemoveChild(refs[index - 1]);
            var relation = rels.DocumentElement.ChildNodes.OfType<XmlElement>().Single(r => r.GetAttribute("Id") == relationId); relation.ParentNode.RemoveChild(relation);
            foreach (XmlElement n in workbook.SelectNodes("//*[local-name()='definedName']")) n.InnerText = Shift(n.InnerText, index);
            Put(zip, "xl/workbook.xml", workbook); Put(zip, "xl/_rels/workbook.xml.rels", rels); zip.GetEntry(part).Delete(); zip.GetEntry(relPart).Delete();
            foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/") && e.FullName.EndsWith(".xml")).ToArray())
            {
                var doc = Doc(zip, entry.FullName);
                foreach (XmlElement f in doc.SelectNodes("//*[local-name()='f']").Cast<XmlElement>().ToArray())
                    if (f.InnerText.Contains("[" + index + "]")) f.ParentNode.RemoveChild(f); else f.InnerText = Shift(f.InnerText, index);
                Put(zip, entry.FullName, doc);
            }
        });
    }
    private static string Shift(string value, int removed)
    { return System.Text.RegularExpressions.Regex.Replace(value, @"\[([0-9]+)\]", m => "[" + (int.Parse(m.Groups[1].Value) > removed ? int.Parse(m.Groups[1].Value) - 1 : int.Parse(m.Groups[1].Value)) + "]"); }
}
public sealed class NameAwareNames
{
    public NameAwareBook Book;
    public object Item(string name) { return new NameAwareName { Book = Book, Lookup = name, Name = Book.WrongNameIdentity ? "wrong!" + name : Book.StripScopeQuote ? name.Replace("'", "") : name }; }
}
public sealed class NameAwareName
{
    public NameAwareBook Book;
    public string Name, Lookup;
    public string RefersTo
    {
        get
        {
            if (Book.FailNameRead) throw new IOException("injected name read failure");
            if (Book.EditOnNameRead) Book.Saved = false;
            using (var memory = new MemoryStream(Book.Bytes)) using (var zip = new ZipArchive(memory, ZipArchiveMode.Read))
            { var doc = NameAwareBook.Doc(zip, "xl/workbook.xml"); string value = doc.SelectNodes("//*[local-name()='definedName']").Cast<XmlElement>().Single(n => NameAwareBook.Identity(n, doc) == Lookup).InnerText; return Book.WrongNameFormula ? value + "+1" : value; }
        }
    }
    public void Delete()
    {
        if (++Book.NameDeletes == Book.FailNameDeleteAt) throw new IOException("injected name Delete failure");
        if (Book.NoNameDelete) return;
        Book.Edit(zip => { var doc = NameAwareBook.Doc(zip, "xl/workbook.xml"); var node = doc.SelectNodes("//*[local-name()='definedName']").Cast<XmlElement>().Single(n => NameAwareBook.Identity(n, doc) == Name); node.ParentNode.RemoveChild(node); NameAwareBook.Put(zip, "xl/workbook.xml", doc); });
    }
}
