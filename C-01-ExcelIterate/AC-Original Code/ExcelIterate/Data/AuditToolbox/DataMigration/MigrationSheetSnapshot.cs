using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.IO.Compression;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class MigrationSheetSnapshot
    {
        internal string Path = "", Sheet = "";
        internal DateTime Stamp;
        internal int MaxRow, MaxColumn;
        internal readonly Dictionary<long, string> Values = new Dictionary<long, string>();
        internal readonly List<MigrationArea> Merges = new List<MigrationArea>();
        internal readonly Dictionary<int, double> Widths = new Dictionary<int, double>();
        internal readonly Dictionary<int, double> Heights = new Dictionary<int, double>();
        private List<long> ordered;
        internal const int CellLimit = 600000;
        internal static long Key(int row, int column) { return (long)row * 20000 + column; }
        internal string Value(int row, int column) { string value; return Values.TryGetValue(Key(row, column), out value) ? value : ""; }
        // Next cell (row-major) after the given one whose text contains term.
        internal bool Find(string term, int afterRow, int afterColumn, out int row, out int column)
        {
            row = column = 0;
            if (string.IsNullOrEmpty(term)) return false;
            if (ordered == null) { ordered = Values.Keys.ToList(); ordered.Sort(); }
            long after = Key(afterRow, afterColumn);
            for (int pass = 0; pass < 2; pass++)
                foreach (long key in ordered)
                {
                    if (pass == 0 ? key <= after : key > after) continue;
                    if (Values[key].IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    row = (int)(key / 20000); column = (int)(key % 20000); return true;
                }
            return false;
        }
        private static readonly string[] MainNs = { "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "http://purl.oclc.org/ooxml/spreadsheetml/main" };
        private static readonly string[] RelNs = { "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "http://purl.oclc.org/ooxml/officeDocument/relationships" };
        private static XmlReaderSettings Settings() { return new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, IgnoreWhitespace = false }; }
        private static bool Is(XmlReader xml, string name) { return xml.NodeType == XmlNodeType.Element && xml.LocalName == name && MainNs.Contains(xml.NamespaceURI); }
        internal static MigrationSheetSnapshot Load(string path, string sheet)
        {
            string extension = System.IO.Path.GetExtension(path);
            if (!extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("窗口内预览只支持 .xlsx/.xlsm；" + extension + " 文件请直接填写目标起点，或在 WPS 中确认位置。");
            var snapshot = new MigrationSheetSnapshot { Path = path, Sheet = sheet, Stamp = File.GetLastWriteTimeUtc(path) };
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                string id = null;
                var book = package.GetEntry("xl/workbook.xml"); if (book == null) throw new InvalidOperationException("工作簿缺少工作表目录。");
                using (var xml = XmlReader.Create(book.Open(), Settings()))
                    while (xml.Read())
                        if (Is(xml, "sheet") && string.Equals(xml.GetAttribute("name"), sheet, StringComparison.OrdinalIgnoreCase))
                        { foreach (string ns in RelNs) { id = id ?? xml.GetAttribute("id", ns); } break; }
                if (id == null) throw new InvalidOperationException("找不到工作表“" + sheet + "”。");
                string target = null;
                var rels = package.GetEntry("xl/_rels/workbook.xml.rels"); if (rels == null) throw new InvalidOperationException("工作簿缺少关系目录。");
                using (var xml = XmlReader.Create(rels.Open(), Settings()))
                    while (xml.Read())
                        if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "Relationship" && xml.GetAttribute("Id") == id) { target = xml.GetAttribute("Target"); break; }
                if (string.IsNullOrEmpty(target)) throw new InvalidOperationException("工作表“" + sheet + "”的内容位置无效。");
                target = target.Replace('\\', '/');
                string entryName = target.StartsWith("/") ? target.TrimStart('/') : "xl/" + target;
                var shared = ReadShared(package.GetEntry("xl/sharedStrings.xml"));
                var part = package.GetEntry(entryName); if (part == null) throw new InvalidOperationException("工作表“" + sheet + "”的内容缺失。");
                using (var xml = XmlReader.Create(part.Open(), Settings())) snapshot.ReadSheet(xml, shared);
            }
            return snapshot;
        }
        private static List<string> ReadShared(ZipArchiveEntry entry)
        {
            var result = new List<string>(); if (entry == null) return result;
            using (var xml = XmlReader.Create(entry.Open(), Settings()))
            {
                // Advance by hand: ReadElementContentAsString already moves
                // past the element, so an extra Read() would skip </si>.
                var text = new StringBuilder(); bool inside = false;
                while (!xml.EOF)
                {
                    if (Is(xml, "si"))
                    {
                        if (xml.IsEmptyElement) { result.Add(""); xml.Read(); continue; }
                        text.Clear(); inside = true; xml.Read(); continue;
                    }
                    if (Is(xml, "rPh")) { xml.Skip(); continue; } // phonetic guide text
                    if (inside && Is(xml, "t")) { text.Append(xml.ReadElementContentAsString()); continue; }
                    if (xml.NodeType == XmlNodeType.EndElement && xml.LocalName == "si") { result.Add(text.ToString()); inside = false; }
                    xml.Read();
                }
            }
            return result;
        }
        private void ReadSheet(XmlReader xml, List<string> shared)
        {
            int row = 0, column = 0;
            while (xml.Read())
            {
                if (Is(xml, "row"))
                {
                    int r; row = int.TryParse(xml.GetAttribute("r"), out r) ? r : row + 1; column = 0;
                    double height; if (xml.GetAttribute("customHeight") == "1" && double.TryParse(xml.GetAttribute("ht"), NumberStyles.Float, CultureInfo.InvariantCulture, out height)) Heights[row] = height;
                }
                else if (Is(xml, "c"))
                {
                    string reference = xml.GetAttribute("r"), type = xml.GetAttribute("t");
                    if (!string.IsNullOrEmpty(reference)) { try { var cell = MigrationArea.Parse(reference); row = cell.Row; column = cell.Column; } catch { column++; } }
                    else column++;
                    if (xml.IsEmptyElement) continue;
                    string value = null, inline = null;
                    int depth = xml.Depth;
                    xml.Read();
                    // Stop on this cell's own end tag; never read past it.
                    while (!xml.EOF && !(xml.NodeType == XmlNodeType.EndElement && xml.Depth == depth))
                    {
                        if (Is(xml, "v")) { value = xml.ReadElementContentAsString(); continue; }
                        if (Is(xml, "t")) { inline = (inline ?? "") + xml.ReadElementContentAsString(); continue; }
                        if (Is(xml, "rPh")) { xml.Skip(); continue; }
                        xml.Read();
                    }
                    string text = Display(type, value, inline, shared);
                    if (text.Length == 0) continue;
                    if (Values.Count >= CellLimit) throw new InvalidOperationException("工作表内容过多（超过 " + CellLimit + " 个非空单元格），无法在窗口内预览。");
                    Values[Key(row, column)] = text;
                    MaxRow = Math.Max(MaxRow, row); MaxColumn = Math.Max(MaxColumn, column);
                }
                else if (Is(xml, "col"))
                {
                    int min, max; double width;
                    if (int.TryParse(xml.GetAttribute("min"), out min) && int.TryParse(xml.GetAttribute("max"), out max) && double.TryParse(xml.GetAttribute("width"), NumberStyles.Float, CultureInfo.InvariantCulture, out width))
                        for (int c = min; c <= Math.Min(max, min + 200); c++) Widths[c] = width;
                }
                else if (Is(xml, "mergeCell"))
                {
                    try { var merge = MigrationArea.Parse(xml.GetAttribute("ref")); Merges.Add(merge); MaxRow = Math.Max(MaxRow, merge.Row + merge.Rows - 1); MaxColumn = Math.Max(MaxColumn, merge.Column + merge.Columns - 1); } catch { }
                }
            }
        }
        private static string Display(string type, string value, string inline, List<string> shared)
        {
            if (type == "s") { int index; return value != null && int.TryParse(value, out index) && index >= 0 && index < shared.Count ? shared[index] : ""; }
            if (type == "inlineStr") return inline ?? "";
            if (value == null) return "";
            if (type == "b") return value == "1" ? "TRUE" : "FALSE";
            if (type == "str" || type == "e") return value;
            double number;
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ? number.ToString("#,##0.####", CultureInfo.InvariantCulture) : value;
        }
    }
}
