using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

namespace Excel_Iterate
{
    // Native WPS serialization normalizes number-format codes and removes data
    // stored beneath a merge. Restore these from the saved source in OUR copy,
    // then verify the persisted package as well as the native COM round-trip.
    internal sealed class MigrationPackagePreservation
    {
        private sealed class Format { internal int Id; internal string Code; internal string Key { get { return Code == null ? "b:" + Id : "c:" + Code; } } }
        private readonly Dictionary<string, Format> formats = new Dictionary<string, Format>();
        private readonly Dictionary<string, XmlElement> hidden = new Dictionary<string, XmlElement>();
        private MigrationArea area;
        internal bool HasHiddenValues { get { return hidden.Count > 0; } }
        private static XmlDocument Read(ZipArchiveEntry entry)
        {
            if (entry == null) throw new InvalidOperationException("缺少工作簿数据部件。");
            var doc = new XmlDocument { XmlResolver = null, PreserveWhitespace = true };
            using (var xml = XmlReader.Create(entry.Open(), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, CloseInput = true, MaxCharactersInDocument = 64L * 1024 * 1024 })) doc.Load(xml);
            return doc;
        }
        private static XmlNamespaceManager Ns(XmlDocument doc)
        { var ns = new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("s", doc.DocumentElement.NamespaceURI); return ns; }
        private static XmlElement Child(XmlElement node, string name)
        { return node.ChildNodes.OfType<XmlElement>().FirstOrDefault(x => x.LocalName == name); }
        private static string SheetPart(ZipArchive zip, string name)
        {
            var book = Read(zip.GetEntry("xl/workbook.xml"));
            var sheet = book.SelectNodes("//s:sheets/s:sheet", Ns(book)).OfType<XmlElement>().FirstOrDefault(x => string.Equals(x.GetAttribute("name"),name,StringComparison.OrdinalIgnoreCase));
            if (sheet == null) throw new InvalidOperationException("文件中缺少工作表：" + name);
            string id = sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if(id.Length==0)id=sheet.GetAttribute("id","http://purl.oclc.org/ooxml/officeDocument/relationships");
            var rels = Read(zip.GetEntry("xl/_rels/workbook.xml.rels"));
            var rel = rels.DocumentElement.ChildNodes.OfType<XmlElement>().FirstOrDefault(x => x.GetAttribute("Id") == id);
            if (rel == null) throw new InvalidOperationException("工作表关系缺失。");
            string part = rel.GetAttribute("Target").Replace('\\', '/');
            return Uri.UnescapeDataString(new Uri(new Uri("http://package/xl/workbook.xml"),part).AbsolutePath.TrimStart('/'));
        }
        internal void RequireCompatibleTarget(string path)
        {
            if(!HasHiddenValues)return;
            using(var stream=File.Open(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))
            {if(stream.ReadByte()!='P'||stream.ReadByte()!='K')throw new InvalidOperationException("源合并区包含隐藏数据，该二进制格式暂不能无损保留；请选择 xlsx/xlsm 目标。");}
        }
        private static Dictionary<string, XmlElement> Cells(XmlDocument sheet)
        { return sheet.SelectNodes("//s:sheetData/s:row/s:c", Ns(sheet)).OfType<XmlElement>().ToDictionary(x => x.GetAttribute("r"), StringComparer.OrdinalIgnoreCase); }
        private static Format NumberFormat(XmlDocument styles, int index)
        {
            var xf = styles.SelectNodes("//s:cellXfs/s:xf", Ns(styles))[index] as XmlElement;
            int id = int.Parse(xf.GetAttribute("numFmtId").Length == 0 ? "0" : xf.GetAttribute("numFmtId"), CultureInfo.InvariantCulture);
            var custom = styles.SelectNodes("//s:numFmts/s:numFmt", Ns(styles)).OfType<XmlElement>().FirstOrDefault(x => x.GetAttribute("numFmtId") == id.ToString(CultureInfo.InvariantCulture));
            return new Format { Id = id, Code = custom == null ? null : custom.GetAttribute("formatCode") };
        }
        private static int Style(XmlElement cell) { return cell == null || cell.GetAttribute("s").Length == 0 ? 0 : int.Parse(cell.GetAttribute("s"), CultureInfo.InvariantCulture); }
        internal static MigrationPackagePreservation Capture(string path, string sheetName, MigrationArea range, bool keepFormat = true)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (stream.ReadByte() != 'P' || stream.ReadByte() != 'K') return null; stream.Position = 0;
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    if (zip.GetEntry("xl/workbook.xml") == null) return null;
                    var result = new MigrationPackagePreservation { area = range };
                    var sheet = Read(zip.GetEntry(SheetPart(zip, sheetName))); var cells = Cells(sheet); var styles = Read(zip.GetEntry("xl/styles.xml"));
                    var rows = sheet.SelectNodes("//s:sheetData/s:row[@s]", Ns(sheet)).OfType<XmlElement>().ToDictionary(x => int.Parse(x.GetAttribute("r")), x => int.Parse(x.GetAttribute("s")));
                    var columns = new Dictionary<int, int>();
                    foreach (XmlElement col in sheet.SelectNodes("//s:cols/s:col[@style]", Ns(sheet)))
                        for (int c = int.Parse(col.GetAttribute("min")); c <= int.Parse(col.GetAttribute("max")) && c <= 16384; c++) columns[c] = int.Parse(col.GetAttribute("style"));
                    var formatCache = new Dictionary<int, Format>();
                    // Content-only modes still preserve data beneath merged cells,
                    // but must not reapply source number formats after saving.
                    if (keepFormat) for (int r = range.Row; r < range.Row + range.Rows; r++) for (int c = range.Column; c < range.Column + range.Columns; c++)
                    {
                        string address = MigrationArea.Cell(r, c); XmlElement cell; cells.TryGetValue(address, out cell); int style;
                        if (cell != null && cell.HasAttribute("s")) style = Style(cell);
                        else if (!rows.TryGetValue(r, out style)) columns.TryGetValue(c, out style);
                        Format format;if(!formatCache.TryGetValue(style,out format))formatCache[style]=format=NumberFormat(styles,style);result.formats[address]=format;
                    }
                    var sharedEntry = zip.GetEntry("xl/sharedStrings.xml");
                    var shared = sharedEntry == null ? new XmlElement[0] : Read(sharedEntry).DocumentElement.ChildNodes.OfType<XmlElement>().ToArray();
                    foreach (XmlElement merge in sheet.SelectNodes("//s:mergeCells/s:mergeCell", Ns(sheet)))
                    {
                        var m = MigrationArea.Parse(merge.GetAttribute("ref")); if (!range.Contains(m)) continue;
                        for (int r = m.Row; r < m.Row + m.Rows; r++) for (int c = m.Column; c < m.Column + m.Columns; c++)
                        {
                            if (r == m.Row && c == m.Column) continue;
                            string address = MigrationArea.Cell(r, c); XmlElement cell;
                            if (!cells.TryGetValue(address, out cell))continue;
                            if (Child(cell, "f") != null) throw new InvalidOperationException("合并从属格包含隐藏公式，暂不能可靠迁移：" + address);
                            if (Child(cell, "v") == null && Child(cell, "is") == null)continue;
                            var copy = (XmlElement)cell.CloneNode(true);
                            if (copy.GetAttribute("t") == "s")
                            {
                                int idx = int.Parse(Child(copy, "v").InnerText, CultureInfo.InvariantCulture); copy.RemoveChild(Child(copy, "v"));
                                var inline = copy.OwnerDocument.CreateElement("is", copy.NamespaceURI); inline.InnerXml = shared[idx].InnerXml; copy.AppendChild(inline); copy.SetAttribute("t", "inlineStr");
                            }
                            result.hidden[address] = copy;
                        }
                    }
                    return result;
                }
            }
        }
        private string Destination(string address, MigrationTarget target)
        { var p = MigrationArea.Parse(address); return MigrationArea.Cell(p.Row + target.Area.Row - area.Row, p.Column + target.Area.Column - area.Column); }
        private static XmlElement EnsureCell(XmlDocument doc, Dictionary<string, XmlElement> cells, string address)
        {
            XmlElement cell; if (cells.TryGetValue(address, out cell)) return cell;
            var point = MigrationArea.Parse(address); var data = (XmlElement)doc.SelectSingleNode("//s:sheetData", Ns(doc));
            var row = data.ChildNodes.OfType<XmlElement>().FirstOrDefault(x => x.GetAttribute("r") == point.Row.ToString(CultureInfo.InvariantCulture));
            if (row == null) { row = doc.CreateElement("row", doc.DocumentElement.NamespaceURI); row.SetAttribute("r", point.Row.ToString(CultureInfo.InvariantCulture)); var next = data.ChildNodes.OfType<XmlElement>().FirstOrDefault(x => int.Parse(x.GetAttribute("r")) > point.Row); data.InsertBefore(row, next); }
            cell = doc.CreateElement("c", doc.DocumentElement.NamespaceURI); cell.SetAttribute("r", address);
            var following = row.ChildNodes.OfType<XmlElement>().FirstOrDefault(x => MigrationArea.Parse(x.GetAttribute("r")).Column > point.Column); row.InsertBefore(cell, following); cells[address] = cell; return cell;
        }
        private static void Save(ZipArchiveEntry entry, XmlDocument doc)
        { using (var stream = entry.Open()) { stream.SetLength(0); using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false })) doc.Save(writer); } }
        internal void Apply(string ownedPath, MigrationTarget target)
        {
            using (var file = File.Open(ownedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                if (file.ReadByte() != 'P' || file.ReadByte() != 'K') { if (HasHiddenValues) throw new InvalidOperationException("该二进制格式暂不能无损保存合并区隐藏数据，请选择 xlsx/xlsm。"); return; } file.Position = 0;
                using (var zip = new ZipArchive(file, ZipArchiveMode.Update))
                {
                    var sheetEntry = zip.GetEntry(SheetPart(zip, target.Sheet)); var sheet = Read(sheetEntry); var cells = Cells(sheet);
                    var styleEntry = zip.GetEntry("xl/styles.xml"); var styles = Read(styleEntry); var xfs = (XmlElement)styles.SelectSingleNode("//s:cellXfs", Ns(styles));
                    var cache = new Dictionary<string, int>(); var fmtIds = new Dictionary<string, int>();var readCache=new Dictionary<int,Format>();
                    var num = (XmlElement)styles.SelectSingleNode("//s:numFmts", Ns(styles));
                    if (num == null) { num = styles.CreateElement("numFmts", styles.DocumentElement.NamespaceURI); styles.DocumentElement.InsertBefore(num, styles.DocumentElement.FirstChild); }
                    int nextId = 164;
                    foreach (XmlElement fmt in num.ChildNodes.OfType<XmlElement>()) { int id = int.Parse(fmt.GetAttribute("numFmtId")); fmtIds[fmt.GetAttribute("formatCode")] = id; nextId = Math.Max(nextId, id + 1); }
                    foreach (var pair in formats)
                    {
                        string address = Destination(pair.Key, target); XmlElement existing; cells.TryGetValue(address, out existing); int old = Style(existing);
                        Format actual;if(!readCache.TryGetValue(old,out actual))readCache[old]=actual=NumberFormat(styles,old);if(actual.Key==pair.Value.Key)continue;
                        int index; string key = old + "|" + pair.Value.Key;
                        if (!cache.TryGetValue(key, out index))
                        {
                            int id = pair.Value.Id;
                            if (pair.Value.Code != null && !fmtIds.TryGetValue(pair.Value.Code, out id)) { id = nextId++; fmtIds[pair.Value.Code] = id; var fmt = styles.CreateElement("numFmt", styles.DocumentElement.NamespaceURI); fmt.SetAttribute("numFmtId", id.ToString()); fmt.SetAttribute("formatCode", pair.Value.Code); num.AppendChild(fmt); }
                            if (xfs.ChildNodes.OfType<XmlElement>().Count() >= 65530) throw new InvalidOperationException("目标单元格格式数超出安全上限。");
                            var xf = (XmlElement)xfs.ChildNodes.OfType<XmlElement>().ElementAt(old).CloneNode(true); xf.SetAttribute("numFmtId", id.ToString()); xf.SetAttribute("applyNumberFormat", "1"); index = xfs.ChildNodes.OfType<XmlElement>().Count(); xfs.AppendChild(xf); cache[key] = index;
                        }
                        EnsureCell(sheet, cells, address).SetAttribute("s", index.ToString());
                    }
                    foreach (var pair in hidden)
                    {
                        var cell = EnsureCell(sheet, cells, Destination(pair.Key, target));
                        foreach (var node in cell.ChildNodes.OfType<XmlElement>().Where(x => x.LocalName == "v" || x.LocalName == "is" || x.LocalName == "f").ToArray()) cell.RemoveChild(node);
                        cell.SetAttribute("t", pair.Value.GetAttribute("t").Length == 0 ? "n" : pair.Value.GetAttribute("t"));
                        foreach (var node in pair.Value.ChildNodes.OfType<XmlElement>().Where(x => x.LocalName == "v" || x.LocalName == "is")) cell.AppendChild(sheet.ImportNode(node, true));
                    }
                    num.SetAttribute("count", num.ChildNodes.OfType<XmlElement>().Count().ToString()); xfs.SetAttribute("count", xfs.ChildNodes.OfType<XmlElement>().Count().ToString());
                    Save(styleEntry, styles); Save(sheetEntry, sheet);
                }
            }
        }
        internal void Verify(string path, MigrationTarget target)
        {
            using (var stream = File.OpenRead(path))
            {
                if (stream.ReadByte() != 'P' || stream.ReadByte() != 'K') return; stream.Position = 0;
                using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
                {
                    var sheet = Read(zip.GetEntry(SheetPart(zip, target.Sheet))); var cells = Cells(sheet); var styles = Read(zip.GetEntry("xl/styles.xml"));
                    var cache=new Dictionary<int,Format>();foreach (var pair in formats) { XmlElement cell; cells.TryGetValue(Destination(pair.Key, target), out cell);int index=Style(cell);Format actual;if(!cache.TryGetValue(index,out actual))cache[index]=actual=NumberFormat(styles,index);if(actual.Key!=pair.Value.Key)throw new IOException("保存文件中的数字格式不一致：" + Destination(pair.Key, target)); }
                    foreach (var pair in hidden)
                    {
                        XmlElement cell; if (!cells.TryGetValue(Destination(pair.Key, target), out cell)) throw new IOException("合并区隐藏内容缺失。");
                        string expected = Child(pair.Value, "is") == null ? Child(pair.Value, "v").InnerText : Child(pair.Value, "is").InnerText;
                        string actual = Child(cell, "is") == null ? (Child(cell, "v") == null ? "" : Child(cell, "v").InnerText) : Child(cell, "is").InnerText;
                        if (expected != actual) throw new IOException("合并区隐藏内容核对失败：" + Destination(pair.Key, target));
                    }
                }
            }
        }
    }
}
