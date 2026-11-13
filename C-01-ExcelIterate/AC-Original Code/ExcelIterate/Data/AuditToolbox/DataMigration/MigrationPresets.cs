using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel_Iterate
{
    internal static class MigrationJson
    {
        internal static string Write(object value)
        {
            var text = new StringBuilder(); Write(text, value, 0); return text.ToString();
        }
        private static void Indent(StringBuilder text, int depth) { text.Append('\n').Append(' ', depth * 2); }
        private static void Write(StringBuilder text, object value, int depth)
        {
            var map = value as IDictionary<string, object>; var list = value as IList<object>;
            if (value == null) text.Append("null");
            else if (value is string) Quote(text, (string)value);
            else if (value is bool) text.Append((bool)value ? "true" : "false");
            else if (value is int || value is long) text.Append(Convert.ToInt64(value).ToString(CultureInfo.InvariantCulture));
            else if (value is double) text.Append(((double)value).ToString("R", CultureInfo.InvariantCulture));
            else if (map != null)
            {
                if (map.Count == 0) { text.Append("{}"); return; }
                text.Append('{'); bool first = true;
                foreach (var pair in map)
                {
                    if (!first) text.Append(','); first = false;
                    Indent(text, depth + 1); Quote(text, pair.Key); text.Append(": "); Write(text, pair.Value, depth + 1);
                }
                Indent(text, depth); text.Append('}');
            }
            else if (list != null)
            {
                if (list.Count == 0) { text.Append("[]"); return; }
                text.Append('[');
                for (int i = 0; i < list.Count; i++) { if (i > 0) text.Append(','); Indent(text, depth + 1); Write(text, list[i], depth + 1); }
                Indent(text, depth); text.Append(']');
            }
            else Quote(text, Convert.ToString(value, CultureInfo.InvariantCulture));
        }
        private static void Quote(StringBuilder text, string value)
        {
            text.Append('"');
            foreach (char c in value)
            {
                if (c == '"') text.Append("\\\"");
                else if (c == '\\') text.Append("\\\\");
                else if (c == '\n') text.Append("\\n");
                else if (c == '\r') text.Append("\\r");
                else if (c == '\t') text.Append("\\t");
                else if (c < ' ') text.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                else text.Append(c);
            }
            text.Append('"');
        }
        internal static object Parse(string text)
        {
            int at = 0; var value = Value(text, ref at); Space(text, ref at);
            if (at != text.Length) throw new FormatException("JSON 末尾有多余内容（位置 " + at + "）");
            return value;
        }
        private static void Space(string text, ref int at) { while (at < text.Length && char.IsWhiteSpace(text[at])) at++; }
        private static Exception Bad(int at) { return new FormatException("JSON 格式错误（位置 " + at + "）"); }
        private static object Value(string text, ref int at)
        {
            Space(text, ref at); if (at >= text.Length) throw Bad(at);
            char c = text[at];
            if (c == '{')
            {
                var map = new Dictionary<string, object>(StringComparer.Ordinal); at++; Space(text, ref at);
                if (at < text.Length && text[at] == '}') { at++; return map; }
                while (true)
                {
                    Space(text, ref at); if (at >= text.Length || text[at] != '"') throw Bad(at);
                    string key = ReadString(text, ref at); Space(text, ref at);
                    if (at >= text.Length || text[at] != ':') throw Bad(at); at++;
                    map[key] = Value(text, ref at); Space(text, ref at);
                    if (at < text.Length && text[at] == ',') { at++; continue; }
                    if (at < text.Length && text[at] == '}') { at++; return map; }
                    throw Bad(at);
                }
            }
            if (c == '[')
            {
                var list = new List<object>(); at++; Space(text, ref at);
                if (at < text.Length && text[at] == ']') { at++; return list; }
                while (true)
                {
                    list.Add(Value(text, ref at)); Space(text, ref at);
                    if (at < text.Length && text[at] == ',') { at++; continue; }
                    if (at < text.Length && text[at] == ']') { at++; return list; }
                    throw Bad(at);
                }
            }
            if (c == '"') return ReadString(text, ref at);
            if (string.CompareOrdinal(text, at, "true", 0, 4) == 0) { at += 4; return true; }
            if (string.CompareOrdinal(text, at, "false", 0, 5) == 0) { at += 5; return false; }
            if (string.CompareOrdinal(text, at, "null", 0, 4) == 0) { at += 4; return null; }
            int start = at;
            while (at < text.Length && "+-0123456789.eE".IndexOf(text[at]) >= 0) at++;
            if (at == start) throw Bad(at);
            string number = text.Substring(start, at - start); long whole; double real;
            if (long.TryParse(number, NumberStyles.Integer, CultureInfo.InvariantCulture, out whole)) return whole;
            if (double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out real)) return real;
            throw Bad(start);
        }
        private static string ReadString(string text, ref int at)
        {
            var result = new StringBuilder(); at++;
            while (at < text.Length)
            {
                char c = text[at++];
                if (c == '"') return result.ToString();
                if (c != '\\') { result.Append(c); continue; }
                if (at >= text.Length) break;
                char e = text[at++];
                if (e == 'n') result.Append('\n'); else if (e == 'r') result.Append('\r'); else if (e == 't') result.Append('\t');
                else if (e == 'b') result.Append('\b'); else if (e == 'f') result.Append('\f');
                else if (e == 'u')
                {
                    if (at + 4 > text.Length) throw Bad(at);
                    result.Append((char)int.Parse(text.Substring(at, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture)); at += 4;
                }
                else result.Append(e);
            }
            throw Bad(at);
        }
        // Typed reads with defaults, so an older or hand-edited preset with a
        // missing field still loads.
        internal static string Str(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) && v != null ? Convert.ToString(v, CultureInfo.InvariantCulture) : ""; }
        internal static bool Flag(IDictionary<string, object> map, string key, bool fallback) { object v; return map != null && map.TryGetValue(key, out v) && v is bool ? (bool)v : fallback; }
        internal static long Num(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) && (v is long || v is double) ? Convert.ToInt64(v) : 0; }
        internal static Dictionary<string, object> Map(IDictionary<string, object> map, string key) { object v; return map != null && map.TryGetValue(key, out v) ? v as Dictionary<string, object> : null; }
        internal static List<Dictionary<string, object>> Maps(IDictionary<string, object> map, string key)
        { object v; var list = map != null && map.TryGetValue(key, out v) ? v as List<object> : null; return list == null ? new List<Dictionary<string, object>>() : list.OfType<Dictionary<string, object>>().ToList(); }
        internal static Dictionary<string, object> Obj(params object[] pairs)
        { var map = new Dictionary<string, object>(StringComparer.Ordinal); for (int i = 0; i + 1 < pairs.Length; i += 2) map[(string)pairs[i]] = pairs[i + 1]; return map; }
    }
    internal sealed class MigrationPreset
    {
        internal const string Format = "xycabin.migration-preset", Extension = ".xymigration.json";
        internal const int Version = 2;
        internal string FilePath = "", Name = "";
        internal DateTime Saved;
        internal Dictionary<string, object> Data;
        internal int Sources, Targets, Mappings;
        internal static string Library
        {
            get
            {
                string root = System.IO.Path.Combine(DialogSizeStore.RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Connect).Assembly.GetName().Name);
                string folder = System.IO.Path.Combine(root, "复制预设"), old = System.IO.Path.Combine(root, "迁移预设");
                try { if (!Directory.Exists(folder) && Directory.Exists(old)) Directory.Move(old, folder); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                return Directory.Exists(folder) || !Directory.Exists(old) ? folder : old;
            }
        }
        internal static MigrationPreset From(Dictionary<string, object> data, string file)
        {
            if (data == null || MigrationJson.Str(data, "format") != Format) throw new InvalidOperationException("不是批量复制数据的预设文件");
            long version = MigrationJson.Num(data, "version");
            if (version < 1) throw new InvalidOperationException("预设文件缺少版本号");
            if (version > Version) throw new InvalidOperationException("预设由更新版本的加载项创建（格式版本 " + version + "），请先升级加载项");
            var preset = new MigrationPreset { FilePath = file ?? "", Data = data, Name = MigrationJson.Str(data, "name") };
            if (preset.Name.Length == 0) preset.Name = System.IO.Path.GetFileName(file ?? "未命名").Replace(Extension, "");
            DateTime saved; if (DateTime.TryParse(MigrationJson.Str(data, "saved"), CultureInfo.InvariantCulture, DateTimeStyles.None, out saved)) preset.Saved = saved;
            preset.Sources = MigrationJson.Maps(data, "sources").Count; preset.Targets = MigrationJson.Maps(data, "targets").Count; preset.Mappings = MigrationJson.Maps(data, "mappings").Count;
            return preset;
        }
        internal static MigrationPreset Read(string file)
        {
            string text;
            try { text = System.IO.File.ReadAllText(file, Encoding.UTF8); } catch (Exception error) { throw new InvalidOperationException("读取预设失败：" + error.Message, error); }
            object data;
            try { data = MigrationJson.Parse(text); } catch (FormatException error) { throw new InvalidOperationException("预设文件已损坏：" + error.Message, error); }
            return From(data as Dictionary<string, object>, file);
        }
        // Written beside the target, then swapped in, so a failed write never
        // leaves half a preset.
        internal static void Write(string file, Dictionary<string, object> data)
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(file));
            string temporary = file + ".tmp";
            System.IO.File.WriteAllText(temporary, MigrationJson.Write(data), new UTF8Encoding(false));
            if (System.IO.File.Exists(file)) System.IO.File.Delete(file);
            System.IO.File.Move(temporary, file);
        }
        internal static List<MigrationPreset> List()
        {
            var result = new List<MigrationPreset>();
            if (!Directory.Exists(Library)) return result;
            foreach (string file in Directory.GetFiles(Library, "*" + Extension))
                try { result.Add(Read(file)); } catch (Exception error) { UiDiagnostics.Record("migration-preset-list", error); }
            return result.OrderByDescending(p => p.Saved).ThenBy(p => p.Name, StringComparer.CurrentCulture).ToList();
        }
        internal static string FileFor(string name)
        {
            var safe = new StringBuilder();
            foreach (char c in name.Trim()) safe.Append(System.IO.Path.GetInvalidFileNameChars().Contains(c) ? '_' : c);
            return System.IO.Path.Combine(Library, (safe.Length == 0 ? "未命名" : safe.ToString()) + Extension);
        }
        internal static string UniqueFile(string name)
        {
            string file = FileFor(name); int n = 2;
            while (System.IO.File.Exists(file)) file = FileFor(name + " (" + n++ + ")");
            return file;
        }
        // Every workbook the preset refers to (lists and 逐行重算 choices).
        internal IEnumerable<string> Files()
        {
            var paths = new List<string>();
            foreach (string list in new[] { "sourceBooks", "targetBooks" }) foreach (var book in MigrationJson.Maps(Data, list)) paths.Add(MigrationJson.Str(book, "path"));
            foreach (string list in new[] { "sources", "targets" })
                foreach (var side in MigrationJson.Maps(Data, list))
                {
                    foreach (var book in MigrationJson.Maps(side, "books")) paths.Add(MigrationJson.Str(book, "path"));
                    foreach (var chosen in MigrationJson.Maps(MigrationJson.Map(side, "trigger"), "chosen")) paths.Add(MigrationJson.Str(chosen, "path"));
                }
            return paths.Where(p => p.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase);
        }
        internal string Relative(string path)
        {
            foreach (string list in new[] { "sourceBooks", "targetBooks" })
                foreach (var book in MigrationJson.Maps(Data, list)) if (string.Equals(MigrationJson.Str(book, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(book, "relative");
            foreach (string list in new[] { "sources", "targets" })
                foreach (var side in MigrationJson.Maps(Data, list))
                {
                    foreach (var book in MigrationJson.Maps(side, "books")) if (string.Equals(MigrationJson.Str(book, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(book, "relative");
                    foreach (var chosen in MigrationJson.Maps(MigrationJson.Map(side, "trigger"), "chosen")) if (string.Equals(MigrationJson.Str(chosen, "path"), path, StringComparison.OrdinalIgnoreCase)) return MigrationJson.Str(chosen, "relative");
                }
            return System.IO.Path.GetFileName(path);
        }
        // The manager's detail pane, as styled lines: title, a dim meta line,
        // then sections (保存方式 / 映射 / 工作簿) with blue headings,
        // numbered mapping blocks and ✓/⚠ workbook lines.
        private static string Circled(int index) { return index >= 1 && index <= 20 ? ((char)(0x2460 + index - 1)).ToString() : index + "."; }
        internal List<PresetDetailLine> DescribeLines()
        {
            var lines = new List<PresetDetailLine>();
            Func<bool, int, PresetDetailLine> line = delegate(bool space, int indent) { var item = new PresetDetailLine { Space = space, Indent = indent }; lines.Add(item); return item; };
            line(false, 0).Add(PresetDetailKind.Title, Name);
            line(false, 0).Add(PresetDetailKind.Meta, "保存于 " + (Saved == DateTime.MinValue ? "未知" : Saved.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)) + "　·　源区域 " + Sources + "　·　目标区域 " + Targets + "　·　映射 " + Mappings);
            var options = MigrationJson.Map(Data, "options");
            line(true, 0).Add(PresetDetailKind.Heading, "保存方式");
            line(false, 1).Add(PresetDetailKind.Text, MigrationJson.Flag(options, "replaceTargets", true) ? "更新目标原文件（自动备份）" : "另存结果副本")
                .Add(PresetDetailKind.Dim, MigrationJson.Flag(options, "allowEmptySource", false) ? "　·　允许全空白源区域" : "");
            Func<Dictionary<string, object>, string> books = delegate(Dictionary<string, object> side)
            {
                var names = MigrationJson.Maps(side, "books").Select(b => System.IO.Path.GetFileName(MigrationJson.Str(b, "path"))).ToList();
                return names.Count == 0 ? "（未选工作簿）" : names.Count == 1 ? names[0] : names[0] + " 等 " + names.Count + " 个";
            };
            Func<Dictionary<string, object>, string> place = delegate(Dictionary<string, object> side)
            { string sheet = MigrationJson.Str(side, "sheet"), range = MigrationJson.Str(side, "range"); return (sheet.Length == 0 ? "未选工作表" : sheet) + "!" + (range.Length == 0 ? "未设区域" : range); };
            var sources = MigrationJson.Maps(Data, "sources").ToDictionary(s => MigrationJson.Str(s, "id"), s => s);
            var targets = MigrationJson.Maps(Data, "targets").ToDictionary(s => MigrationJson.Str(s, "id"), s => s);
            var mappings = MigrationJson.Maps(Data, "mappings").ToList();
            line(true, 0).Add(PresetDetailKind.Heading, "映射（" + mappings.Count + "）");
            if (mappings.Count == 0) line(false, 1).Add(PresetDetailKind.Dim, "（没有映射）");
            int index = 0;
            foreach (var mapping in mappings)
            {
                index++;
                Dictionary<string, object> source, target;
                sources.TryGetValue(MigrationJson.Str(mapping, "source"), out source); targets.TryGetValue(MigrationJson.Str(mapping, "target"), out target);
                var trigger = source == null ? null : MigrationJson.Map(source, "trigger");
                var head = line(index > 1, 1).Add(PresetDetailKind.Number, Circled(index) + " ").Add(PresetDetailKind.Label, "源　　");
                if (source == null) head.Add(PresetDetailKind.Dim, "（未配对）"); else head.Add(PresetDetailKind.Text, place(source)).Add(PresetDetailKind.Dim, "　" + books(source));
                var to = line(false, 2).Add(PresetDetailKind.Label, "目标　");
                if (target == null) to.Add(PresetDetailKind.Dim, "（未配对）"); else to.Add(PresetDetailKind.Text, place(target)).Add(PresetDetailKind.Dim, "　" + (trigger != null ? "按文件名匹配" : books(target)));
                var tags = line(false, 2).Add(PresetDetailKind.Tag, MigrationPair.ModeText(trigger != null || MigrationJson.Flag(mapping, "valuesOnly", false), MigrationJson.Flag(mapping, "keepFormat", true), trigger == null && MigrationJson.Flag(mapping, "formatOnly", false), trigger == null && MigrationJson.Flag(mapping, "pasteLinks", false)));
                if (!MigrationJson.Flag(mapping, "enabled", true)) tags.Add(PresetDetailKind.Dim, "　·　").Add(PresetDetailKind.Warn, "未启用");
                if (trigger != null) tags.Add(PresetDetailKind.Dim, "　·　").Add(PresetDetailKind.Tag, "逐行重算 ")
                    .Add(PresetDetailKind.Dim, MigrationJson.Str(trigger, "values") + " → " + MigrationJson.Str(trigger, "cell") + "，目标文件区域 " + MigrationJson.Str(trigger, "nameColumn")
                        + string.Concat(MigrationJson.Maps(trigger, "conditions").Select(c => "，" + MigrationJson.Str(c, "range") + (MigrationJson.Str(c, "range").Length == 0 && MigrationJson.Str(c, "column").Length > 0 ? MigrationJson.Str(c, "column") + " 列" : MigrationJson.Str(c, "range").Length == 0 ? MigrationJson.Str(c, "value") : "") + " → " + MigrationJson.Str(c, "cell")).ToArray()));
            }
            var files = Files().ToList(); int missing = files.Count(f => !System.IO.File.Exists(f));
            line(true, 0).Add(PresetDetailKind.Heading, "工作簿（" + files.Count + "）")
                .Add(missing > 0 ? PresetDetailKind.Warn : PresetDetailKind.Ok, missing > 0 ? "　" + missing + " 个在本机找不到，应用时可重新定位" : "　均可找到");
            foreach (string file in files)
            {
                bool found = System.IO.File.Exists(file);
                line(false, 1).Add(found ? PresetDetailKind.Ok : PresetDetailKind.Warn, found ? "✓  " : "⚠  ").Add(found ? PresetDetailKind.Text : PresetDetailKind.Warn, System.IO.Path.GetFileName(file))
                    .Add(PresetDetailKind.Dim, "　" + System.IO.Path.GetDirectoryName(file));
            }
            return lines;
        }
        // The same content as plain text.
        internal string Describe()
        { return string.Join(Environment.NewLine, DescribeLines().Select(l => l.Text).ToArray()); }
    }
    internal static class MigrationPresetCodec
    {
        // ---------------- presets ----------------
        // The whole setup as preset data: workbooks (absolute path plus the
        // path relative to their common folder), ranges, mappings, write
        // modes and 逐行重算 settings. Never any sheet data.
        internal static Dictionary<string, object> Capture(string name, MigrationConfiguration configuration)
        {
            var pairs = configuration.Pairs;
            var sources = pairs.Select(p => p.Source).Where(s => s.Used).Distinct().ToList();
            var targets = pairs.Select(p => p.Target).Where(s => s.Used).Distinct().ToList();
            var paths = configuration.SourceBooks.Select(b => b.Path).Concat(configuration.TargetBooks.Select(b => b.Path)).Concat(sources.Concat(targets).SelectMany(s => s.Books))
                .Concat(sources.Where(s => s.Trigger != null).SelectMany(s => s.Trigger.Chosen.Values).Where(v => v.Length > 0)).ToList();
            string root = CommonFolder(paths);
            Func<string, object> book = delegate(string path) { return MigrationJson.Obj("path", path, "relative", RelativeTo(root, path)); };
            Func<MigrationSide, string, bool, object> sideData = delegate(MigrationSide side, string id, bool target)
            {
                object trigger = null;
                if (!target && side.Trigger != null)
                    trigger = MigrationJson.Obj("cell", side.Trigger.Cell, "values", side.Trigger.Values, "nameColumn", side.Trigger.NameColumn,
                        "conditions", side.Trigger.Extra.Select(c => (object)MigrationJson.Obj("cell", c.Cell, "range", c.Range, "column", c.Column, "value", c.Value)).ToList(),
                        "chosen", side.Trigger.Chosen.OrderBy(c => c.Key).Select(c => (object)MigrationJson.Obj("row", (long)c.Key, "path", c.Value, "relative", c.Value.Length == 0 ? "" : RelativeTo(root, c.Value))).ToList());
                var data = MigrationJson.Obj("id", id, "books", side.Books.Select(book).ToList(), "sheet", side.Sheet, "range", side.Range);
                if (trigger != null) data["trigger"] = trigger;
                return data;
            };
            var mappings = pairs.Select(p => (object)MigrationJson.Obj(
                "source", p.Source.Used ? (object)("s" + (sources.IndexOf(p.Source) + 1)) : null,
                "target", p.Target.Used ? (object)("t" + (targets.IndexOf(p.Target) + 1)) : null,
                "enabled", p.Enabled, "valuesOnly", p.ValuesOnly, "keepFormat", p.KeepFormat, "formatOnly", p.FormatOnly, "pasteLinks", p.PasteLinks && p.Source.Trigger == null)).ToList();
            return MigrationJson.Obj(
                "format", MigrationPreset.Format, "version", (long)MigrationPreset.Version, "name", name,
                "saved", DateTime.Now.ToString("s", CultureInfo.InvariantCulture), "root", root,
                "options", MigrationJson.Obj("replaceTargets", configuration.ReplaceTargets, "allowEmptySource", configuration.AllowEmptySource),
                "sourceBooks", configuration.SourceBooks.Select(b => book(b.Path)).ToList(),
                "targetBooks", configuration.TargetBooks.Select(b => book(b.Path)).ToList(),
                "sources", sources.Select((s, i) => sideData(s, "s" + (i + 1), false)).ToList(),
                "targets", targets.Select((s, i) => sideData(s, "t" + (i + 1), true)).ToList(),
                "mappings", mappings);
        }
        internal static MigrationConfiguration Read(Dictionary<string, object> data, IDictionary<string, string> remap)
        {
            MigrationPreset.From(data, null);
            Func<string, string> map = delegate(string path)
            {
                string found; if (path.Length == 0) return path;
                return remap != null && remap.TryGetValue(path, out found) ? found : path;
            };
            Func<Dictionary<string, object>, string> pathOf = delegate(Dictionary<string, object> b) { return map(MigrationJson.Str(b, "path")); };
            var options = MigrationJson.Map(data, "options");
            var pairs = new List<MigrationPair>();
            var sides = new Dictionary<string, MigrationSide>(StringComparer.Ordinal);
            foreach (string list in new[] { "sources", "targets" })
                foreach (var item in MigrationJson.Maps(data, list))
                {
                    var side = new MigrationSide { Used = true, Sheet = MigrationJson.Str(item, "sheet"), Range = MigrationPlanner.NormalizeAddress(MigrationJson.Str(item, "range")) };
                    foreach (var b in MigrationJson.Maps(item, "books")) { string path = pathOf(b); if (path.Length > 0) side.Books.Add(Path.GetFullPath(path)); }
                    var trigger = MigrationJson.Map(item, "trigger");
                    if (list == "sources" && trigger != null)
                    {
                        side.Trigger = new MigrationTrigger { Cell = MigrationJson.Str(trigger, "cell"), Values = MigrationJson.Str(trigger, "values"), NameColumn = MigrationJson.Str(trigger, "nameColumn") };
                        foreach (var condition in MigrationJson.Maps(trigger, "conditions"))
                            side.Trigger.Extra.Add(new MigrationCondition { Cell = MigrationJson.Str(condition, "cell"), Range = MigrationJson.Str(condition, "range"), Column = MigrationJson.Str(condition, "column"), Value = MigrationJson.Str(condition, "value") });
                        foreach (var chosen in MigrationJson.Maps(trigger, "chosen"))
                        { string path = pathOf(chosen); side.Trigger.Chosen[(int)MigrationJson.Num(chosen, "row")] = path.Length == 0 ? "" : Path.GetFullPath(path); }
                    }
                    sides[(list == "sources" ? "S:" : "T:") + MigrationJson.Str(item, "id")] = side;
                }
            foreach (var item in MigrationJson.Maps(data, "mappings"))
            {
                // Presets saved before 写入方式 had four choices always copied formats.
                var pair = new MigrationPair { Enabled = MigrationJson.Flag(item, "enabled", true), ValuesOnly = MigrationJson.Flag(item, "valuesOnly", false), KeepFormat = MigrationJson.Flag(item, "keepFormat", true), FormatOnly = MigrationJson.Flag(item, "formatOnly", false), PasteLinks = MigrationJson.Flag(item, "pasteLinks", false) };
                MigrationSide side;
                if (sides.TryGetValue("S:" + MigrationJson.Str(item, "source"), out side)) pair.Source = side;
                if (sides.TryGetValue("T:" + MigrationJson.Str(item, "target"), out side)) pair.Target = side;
                if (pair.Source.Trigger != null) pair.SelectMode(pair.KeepFormat ? MigrationPair.ModeValuesFormat : MigrationPair.ModeValues);
                else if (pair.PasteLinks) pair.SelectMode(MigrationPair.ModeLinks);
                if (pair.Source.Used || pair.Target.Used) pairs.Add(pair);
            }
            Func<string, List<MigrationBook>> books = key => MigrationJson.Maps(data, key).Select(pathOf).Where(p => p.Length > 0).Select(p => new MigrationBook { Path = Path.GetFullPath(p) }).ToList();
            return new MigrationConfiguration(pairs, books("sourceBooks"), books("targetBooks"), MigrationJson.Flag(options, "replaceTargets", true), MigrationJson.Flag(options, "allowEmptySource", false));
        }
        private static string CommonFolder(IEnumerable<string> paths)
        {
            string[] common = null;
            foreach (string path in paths)
            {
                string folder = Path.GetDirectoryName(path); if (string.IsNullOrEmpty(folder)) continue;
                var parts = folder.TrimEnd('\\').Split('\\');
                if (common == null) { common = parts; continue; }
                int n = 0; while (n < common.Length && n < parts.Length && string.Equals(common[n], parts[n], StringComparison.OrdinalIgnoreCase)) n++;
                common = common.Take(n).ToArray();
            }
            return common == null || common.Length == 0 ? "" : string.Join("\\", common) + (common.Length == 1 ? "\\" : "");
        }
        private static string RelativeTo(string root, string path)
        {
            if (root.Length == 0) return Path.GetFileName(path);
            string prefix = root.EndsWith("\\") ? root : root + "\\";
            return path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? path.Substring(prefix.Length) : Path.GetFileName(path);
        }
    }
}

