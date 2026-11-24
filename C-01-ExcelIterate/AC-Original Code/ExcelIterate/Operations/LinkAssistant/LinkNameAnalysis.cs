using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class LinkDefinedName
    {
        internal string Name, Sheet, Formula;
        internal readonly HashSet<string> Sources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal string Key { get { return (Sheet ?? "") + "\n" + Name; } }
        internal string HostName { get { return Sheet == null ? Name : "'" + Sheet.Replace("'", "''") + "'!" + Name; } }
    }

    // Read-only package analysis. Numeric [n] references are resolved through
    // workbook/externalReferences order, never through externalLink filenames.
    internal sealed class LinkNameAnalysis
    {
        private sealed class FormulaUse
        {
            internal string Key, Text, Sheet;
            internal bool Cell;
            internal HashSet<string> Sources;
        }
        private static readonly Regex Prefix = new Regex(@"'(?:[^']|'')*'!|\[[^\[\]\r\n]+\][\p{L}\p{N}_\.]*!", RegexOptions.Compiled);
        private static readonly Regex Qualifier = new Regex(@"'(?:[^']|'')*'!|(?:\[[^\[\]\r\n]+\][\p{L}\p{N}_\.]*|[\p{L}\p{N}_\.]+)!", RegexOptions.Compiled);
        private static readonly Regex DynamicName = new Regex(@"(?<![\p{L}\p{N}_\.])(?:_xlfn\.)?(?:INDIRECT|EVALUATE|GET\.CELL|GET\.WORKBOOK|GET\.WINDOW|CALL|REGISTER)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private readonly Dictionary<int, string> sources = new Dictionary<int, string>();
        private readonly List<FormulaUse> formulas = new List<FormulaUse>();
        internal readonly List<LinkDefinedName> Names = new List<LinkDefinedName>();
        private string owner, uncertain;

        private static string PartPath(string target)
        { return new Uri(new Uri("http://package/xl/workbook.xml"), target).AbsolutePath.TrimStart('/'); }
        private static XmlDocument Document(ZipArchive zip, string part)
        {
            ZipArchiveEntry entry = zip.GetEntry(part);
            if (entry == null) throw new IOException("名称引用所需部件缺失：" + part);
            using (Stream stream = entry.Open()) return XlsxLinkPackage.Load(stream);
        }
        private static IEnumerable<XmlElement> Children(XmlNode node, string name)
        { return node.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == name); }
        private static string RelId(XmlElement node)
        { return node.Attributes.Cast<XmlAttribute>().Where(a => a.LocalName == "id").Select(a => a.Value).SingleOrDefault(); }
        private static string NormalPath(string path) { return Path.GetFullPath(path).ToUpperInvariant(); }
        private static string MaskStrings(string text)
        {
            var result = new StringBuilder(text); bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '"')
                {
                    result[i] = ' ';
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"') { result[++i] = ' '; continue; }
                    quoted = !quoted;
                }
                else if (quoted) result[i] = ' ';
            }
            if (quoted) throw new IOException("公式字符串不完整，无法安全分析名称引用。");
            return result.ToString();
        }

        private bool SourcePrefix(string prefix, out string source, out string sheet)
        {
            source = null; sheet = null;
            string text = prefix.Substring(0, prefix.Length - 1);
            if (text.StartsWith("'")) text = text.Substring(1, text.Length - 2).Replace("''", "'");
            int left = text.IndexOf('['), right = text.IndexOf(']');
            if (left < 0) return false;
            if (right <= left || text.IndexOf('[', left + 1) >= 0) throw new IOException("无法可靠解析名称中的工作簿引用。");
            string book = text.Substring(left + 1, right - left - 1); int index;
            if (left == 0 && Regex.IsMatch(book, "^[0-9]+$"))
            {
                if (!int.TryParse(book, out index) || !sources.TryGetValue(index, out source) || source == null) throw new IOException("名称中的外部引用编号无有效文件来源：" + book);
            }
            else
            {
                string path = text.Substring(0, left) + book;
                // Native RefersTo contains Windows paths, not package URIs.
                source = path.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? XlsxLinkPackage.Resolve(owner, path)
                    : Path.GetFullPath(Path.IsPathRooted(path) ? path : Path.Combine(Path.GetDirectoryName(owner), path));
            }
            sheet = text.Substring(right + 1);
            return !string.Equals(source, owner, StringComparison.OrdinalIgnoreCase);
        }
        private HashSet<string> FormulaSources(string formula)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in Prefix.Matches(MaskStrings(formula)))
            { string path, sheet; if (SourcePrefix(formula.Substring(match.Index, match.Length), out path, out sheet)) result.Add(path); }
            return result;
        }
        internal string Semantic(string formula)
        {
            return NormalizeFormula(formula, null, null);
        }
        private string NormalizeFormula(string formula, string previous, string replacement)
        {
            formula = (formula ?? "").Trim(); if (formula.StartsWith("=")) formula = formula.Substring(1);
            var result = new StringBuilder(formula);
            foreach (Match match in Prefix.Matches(MaskStrings(formula)).Cast<Match>().Reverse())
            {
                string path, sheet;
                if (SourcePrefix(formula.Substring(match.Index, match.Length), out path, out sheet))
                {
                    if (string.Equals(path, previous, StringComparison.OrdinalIgnoreCase)) path = replacement;
                    result.Remove(match.Index, match.Length);
                    result.Insert(match.Index, "'[" + NormalPath(path) + "]" + sheet.ToUpperInvariant().Replace("'", "''") + "'!");
                }
            }
            return result.ToString();
        }

        internal static LinkNameAnalysis Read(string file)
        {
            var result = new LinkNameAnalysis { owner = Path.GetFullPath(file) };
            using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                XmlDocument workbook = Document(zip, "xl/workbook.xml");
                XmlElement root = workbook.DocumentElement;
                var sheetNodes = Children(root, "sheets").SelectMany(n => Children(n, "sheet")).ToList();
                var refs = Children(root, "externalReferences").SelectMany(n => Children(n, "externalReference")).ToList();
                var rels = new Dictionary<string, string>(StringComparer.Ordinal);
                if (refs.Count > 0 || sheetNodes.Count > 0)
                    foreach (XmlElement rel in Document(zip, "xl/_rels/workbook.xml.rels").DocumentElement.ChildNodes.OfType<XmlElement>())
                        rels.Add(rel.GetAttribute("Id"), rel.GetAttribute("Target"));
                for (int i = 0; i < refs.Count; i++)
                {
                    string id = RelId(refs[i]), part;
                    if (id == null || !rels.TryGetValue(id, out part)) throw new IOException("工作簿外部引用关系无法对应。");
                    part = PartPath(part);
                    string relPart = part.Substring(0, part.LastIndexOf('/') + 1) + "_rels/" + part.Substring(part.LastIndexOf('/') + 1) + ".rels";
                    XmlDocument link = Document(zip, part);
                    var book = Children(link.DocumentElement, "externalBook").SingleOrDefault();
                    string path = null;
                    if (book != null)
                    {
                        string linkId = RelId(book);
                        var target = Document(zip, relPart).DocumentElement.ChildNodes.OfType<XmlElement>()
                            .SingleOrDefault(r => r.GetAttribute("Id") == linkId && (r.GetAttribute("Type").EndsWith("/externalLinkPath", StringComparison.Ordinal)
                                || r.GetAttribute("Type").EndsWith(XlsxLinkPackage.MissingPathSuffix, StringComparison.Ordinal)));
                        if (target == null) throw new IOException("名称的外部工作簿路径无法读取。");
                        path = XlsxLinkPackage.Resolve(file, target.GetAttribute("Target"));
                    }
                    result.sources.Add(i + 1, path);
                }
                var sheetParts = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (XmlElement sheet in sheetNodes)
                { string target; if (rels.TryGetValue(RelId(sheet) ?? "", out target)) sheetParts[PartPath(target)] = sheet.GetAttribute("name"); }
                foreach (XmlElement element in Children(root, "definedNames").SelectMany(n => Children(n, "definedName")))
                {
                    string scope = null; int index;
                    if (element.HasAttribute("localSheetId"))
                    {
                        if (!int.TryParse(element.GetAttribute("localSheetId"), out index) || index < 0 || index >= sheetNodes.Count) throw new IOException("名称的工作表作用域无效。");
                        scope = sheetNodes[index].GetAttribute("name");
                    }
                    var name = new LinkDefinedName { Name = element.GetAttribute("name"), Sheet = scope, Formula = element.InnerText };
                    name.Sources.UnionWith(result.FormulaSources(name.Formula)); result.Names.Add(name);
                    result.AddFormula("name:" + name.Key, name.Formula, scope, false);
                }
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    string part = entry.FullName;
                    if (part.StartsWith("xl/externalLinks/", StringComparison.Ordinal) || part == "xl/workbook.xml") continue;
                    if (part.StartsWith("xl/activeX/", StringComparison.Ordinal) || part.StartsWith("xl/embeddings/", StringComparison.Ordinal)
                        || part.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)) result.uncertain = "存在无法可靠分析名称依赖的嵌入或宏对象";
                    if (!part.StartsWith("xl/", StringComparison.Ordinal) || !(part.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || part.EndsWith(".vml", StringComparison.OrdinalIgnoreCase))) continue;
                    XmlDocument doc; using (Stream input = entry.Open()) doc = XlsxLinkPackage.Load(input);
                    string scope; sheetParts.TryGetValue(part, out scope); int ordinal = 0;
                    foreach (XmlElement element in doc.SelectNodes("//*").OfType<XmlElement>())
                    {
                        string tag = element.LocalName.ToLowerInvariant();
                        if (tag == "f" || tag.Contains("formula") || tag.StartsWith("fmla", StringComparison.Ordinal))
                        {
                            bool cell = tag == "f" && element.ParentNode is XmlElement && ((XmlElement)element.ParentNode).LocalName == "c" && scope != null;
                            string key = cell ? part + ":" + ((XmlElement)element.ParentNode).GetAttribute("r") : part + ":object:" + (++ordinal);
                            result.AddFormula(key, element.InnerText, scope, cell);
                        }
                        foreach (XmlAttribute attr in element.Attributes)
                            if (attr.LocalName.ToLowerInvariant().Contains("fmla") || attr.LocalName.ToLowerInvariant().Contains("formula")
                                || (tag == "worksheetsource" && attr.LocalName == "name"))
                                result.AddFormula(part + ":attribute:" + (++ordinal), attr.Value, scope, false);
                    }
                }
            }
            return result;
        }
        private void AddFormula(string key, string text, string sheet, bool cell)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (DynamicName.IsMatch(MaskStrings(text))) uncertain = "存在动态名称引用，无法证明删除名称不会影响数据";
            formulas.Add(new FormulaUse { Key = key, Text = text, Sheet = sheet, Cell = cell, Sources = FormulaSources(text) });
        }

        private bool UsesName(FormulaUse use, LinkDefinedName name)
        {
            if (use.Text.IndexOf(name.Name, StringComparison.OrdinalIgnoreCase) < 0) return false;
            string masked = MaskStrings(use.Text);
            var qualifiers = Qualifier.Matches(masked).Cast<Match>().ToList();
            var visible = new StringBuilder(masked);
            for (int i = 0; i < masked.Length; i++)
            {
                if (masked[i] != '[' || qualifiers.Any(q => i >= q.Index && i < q.Index + q.Length)) continue;
                int depth = 0;
                do { if (masked[i] == '[') depth++; else if (masked[i] == ']') depth--; visible[i++] = ' '; }
                while (i < masked.Length && depth > 0);
                if (depth != 0) throw new IOException("结构化引用不完整，无法安全分析名称依赖。");
                i--;
            }
            masked = visible.ToString();
            string pattern = @"(?<![\p{L}\p{N}_\.\\])" + Regex.Escape(name.Name) + @"(?![\p{L}\p{N}_\.\\])";
            List<RefSpan> addresses = null;
            foreach (Match match in Regex.Matches(masked, pattern, RegexOptions.IgnoreCase))
            {
                if (qualifiers.Any(q => match.Index >= q.Index && match.Index < q.Index + q.Length)) continue;
                int end = match.Index + match.Length; while (end < masked.Length && char.IsWhiteSpace(masked[end])) end++;
                if (end < masked.Length && (masked[end] == '(' || masked[end] == '!')) continue;
                // A column token in Y$5, $Y$5 or Y:Y is an address, not name Y.
                // Reuse the A1 grammar; bare Y and Sheet!Y still reach scope checks.
                if (addresses == null) addresses = FormulaRefRewriter.Scan(masked);
                if (addresses.Any(a => match.Index >= a.Start && match.Index + match.Length <= a.Start + a.Length
                    && (a.IsWholeColumn || (a.Row1 >= 1 && a.Row1 <= FormulaRefRewriter.MaxRow
                        && (!a.IsRange || (a.Row2 >= 1 && a.Row2 <= FormulaRefRewriter.MaxRow)))))) continue;
                int start = match.Index; while (start > 0 && char.IsWhiteSpace(masked[start - 1])) start--;
                Match qualifier = qualifiers.FirstOrDefault(q => q.Index + q.Length == start);
                string scope = use.Sheet;
                if (qualifier != null)
                {
                    string text = qualifier.Value.TrimEnd('!');
                    if (text.Contains("[")) continue; // An external book's name is not this workbook's name.
                    scope = text.StartsWith("'") ? text.Substring(1, text.Length - 2).Replace("''", "'") : text;
                }
                if (scope == null) return true; // Object/global-name context cannot rule out a local name.
                if (name.Sheet != null) { if (string.Equals(scope, name.Sheet, StringComparison.OrdinalIgnoreCase)) return true; }
                else if (!Names.Any(n => n.Sheet != null && string.Equals(n.Sheet, scope, StringComparison.OrdinalIgnoreCase) && string.Equals(n.Name, name.Name, StringComparison.OrdinalIgnoreCase))) return true;
            }
            return false;
        }
        internal List<LinkDefinedName> Plan(string source)
        {
            var names = Names.Where(n => n.Sources.Contains(source)).ToList();
            if (formulas.Any(f => f.Cell && f.Sources.Contains(source) && f.Sources.Count > 1))
                throw new IOException("同一公式同时引用所选及其他来源，无法保证仅断开指定来源。");
            if (names.Count == 0) return names;
            if (uncertain != null) throw new IOException(uncertain + "，未开始断链。");
            foreach (var name in names)
            {
                if (name.Sources.Count != 1 || Names.Count(n => string.Equals(n.Key, name.Key, StringComparison.OrdinalIgnoreCase)) != 1)
                    throw new IOException("名称含混合来源或作用域不唯一，不能安全清理：" + name.HostName);
                if (formulas.Any(f => f.Key != "name:" + name.Key && UsesName(f, name)))
                    throw new IOException("检测到名称引用或无法排除依赖，不能安全删除：" + name.HostName);
            }
            return names;
        }
        internal static void Describe(string file, IEnumerable<ExternalLinkInfo> links)
        {
            try
            {
                var analysis = Read(file);
                foreach (var link in links.Where(l => l.SourcePath != null))
                {
                    int count = analysis.Names.Count(n => n.Sources.Contains(link.SourcePath));
                    if (count == 0) continue;
                    bool other = analysis.formulas.Any(f => !f.Key.StartsWith("name:", StringComparison.Ordinal) && f.Sources.Contains(link.SourcePath));
                    link.Status = (other ? "公式及名称定义引用" : "仅名称管理器定义引用") + "（" + count + " 个名称）";
                }
            }
            catch (Exception error)
            { foreach (var link in links.Where(l => l.SourcePath != null)) link.Status = "来源记录存在，名称引用待核对：" + error.Message; }
        }
        internal List<object> NativeNames(object book, List<LinkDefinedName> plan, ComBag references)
        {
            var result = new List<object>(); if (plan.Count == 0) return result;
            dynamic names = references.Own(((dynamic)book).Names);
            foreach (var planned in plan)
            {
                dynamic actual = references.Own(names.Item(planned.HostName));
                string label = Convert.ToString(actual.Name), formula = Convert.ToString(actual.RefersTo);
                int bang = label.LastIndexOf('!'); string scope = bang < 0 ? "" : label.Substring(0, bang);
                if (scope.StartsWith("'") && scope.EndsWith("'")) scope = scope.Substring(1, scope.Length - 2).Replace("''", "'");
                string identity = scope + "\n" + (bang < 0 ? label : label.Substring(bang + 1));
                if (!string.Equals(identity, planned.Key, StringComparison.OrdinalIgnoreCase)
                    || Semantic(formula) != Semantic(planned.Formula)) throw new IOException("待清理名称的身份或引用已变化：" + planned.HostName);
                result.Add((object)actual);
            }
            return result;
        }
        internal void Verify(LinkNameAnalysis after, string source, List<LinkDefinedName> removed)
        {
            var keys = new HashSet<string>(removed.Select(n => n.Key), StringComparer.OrdinalIgnoreCase);
            if (after.Names.Any(n => keys.Contains(n.Key) || n.Sources.Contains(source))) throw new IOException("保存后所选来源的名称引用仍存在。");
            Func<LinkNameAnalysis, IEnumerable<LinkDefinedName>, Dictionary<string, string>> signatures = (a, items) => items.GroupBy(n => n.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => string.Join("\n", g.Select(n => a.Semantic(n.Formula)).OrderBy(t => t, StringComparer.Ordinal)), StringComparer.OrdinalIgnoreCase);
            var beforeNames = signatures(this, Names.Where(n => !keys.Contains(n.Key)));
            var afterNames = signatures(after, after.Names);
            if (beforeNames.Count != afterNames.Count || beforeNames.Any(n => !afterNames.ContainsKey(n.Key) || afterNames[n.Key] != n.Value)) throw new IOException("保存后其他名称定义发生变化，备份已保留。");
            var cells = after.formulas.Where(f => f.Cell).ToDictionary(f => f.Key, StringComparer.Ordinal);
            foreach (var before in formulas.Where(f => f.Cell && !f.Sources.Contains(source)))
            {
                FormulaUse current;
                if (!cells.TryGetValue(before.Key, out current) || Semantic(before.Text) != after.Semantic(current.Text)) throw new IOException("保存后其他单元格的公式映射发生变化：" + before.Key);
            }
        }
        internal void VerifyReplacement(LinkNameAnalysis after, string source, string destination)
        {
            var saved = after.formulas.Where(UsesCellOrName).ToLookup(FormulaIdentity, StringComparer.OrdinalIgnoreCase);
            foreach (var before in formulas.Where(formula => UsesCellOrName(formula) && formula.Sources.Contains(source)))
            {
                string expected = NormalizeFormula(before.Text, source, destination);
                if (saved[FormulaIdentity(before)].Count(formula => after.Semantic(formula.Text) == expected) != 1)
                    throw new IOException("保存后无法核对外链公式，恢复备份已保留：" + FormulaIdentity(before));
            }
        }
        private static bool UsesCellOrName(FormulaUse formula)
        {
            return formula.Cell || formula.Key.StartsWith("name:", StringComparison.Ordinal);
        }
        private static string FormulaIdentity(FormulaUse formula)
        {
            return formula.Cell ? formula.Sheet + "!" + formula.Key.Substring(formula.Key.LastIndexOf(':') + 1) : formula.Key;
        }
    }
}
