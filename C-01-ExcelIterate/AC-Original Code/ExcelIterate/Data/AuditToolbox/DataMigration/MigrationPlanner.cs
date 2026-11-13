using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class MigrationPlanner
    {
        private readonly MigrationConfiguration configuration;
        private IList<MigrationPair> pairs { get { return configuration.Pairs; } }
        private IList<MigrationBook> TargetBooks { get { return configuration.TargetBooks; } }
        internal MigrationPlanner(MigrationConfiguration snapshot) { configuration = snapshot; }
        private MigrationBook FindBook(bool target, string path)
        { return (target ? configuration.TargetBooks : configuration.SourceBooks).FirstOrDefault(b => string.Equals(b.Path, path, StringComparison.OrdinalIgnoreCase)); }
        private string RowProblem(MigrationPair pair) { return RowProblem(pair, TargetBooks.Count); }

        // ---------------- plan ----------------
        // Expand every enabled row into concrete writes: one source workbook
        // to every target workbook, or N source workbooks one-to-one with N
        // target workbooks. Block overlaps and source-as-target, then group
        // the rest into passes for the existing single-source engine.
        internal MigrationPlan Build()
        {
            if (pairs.Count == 0) throw new InvalidOperationException("请先在“源区域”和“目标区域”页各新增至少一行。");
            var plan = new MigrationPlan { ReplaceTargets = configuration.ReplaceTargets, AllowEmptySource = configuration.AllowEmptySource };
            var items = new Dictionary<string, MigrationSourceItem>(StringComparer.OrdinalIgnoreCase);
            var targetPaths = new HashSet<string>(pairs.Where(p => p.Enabled && p.Target.Used).SelectMany(p => p.Target.Books), StringComparer.OrdinalIgnoreCase);
            var valid = new List<MigrationMapping>();
            for (int number = 1; number <= pairs.Count; number++)
            {
                var pair = pairs[number - 1]; pair.Instances.Clear();
                if (!pair.Enabled || RowProblem(pair).Length > 0) continue;
                var source = MigrationArea.Parse(pair.Source.Range); var area = TargetArea(pair);
                if (pair.Source.Trigger != null) { ExpandTrigger(plan, pair, number, area, targetPaths, valid); continue; }
                var books = pair.Source.Books.Count == 1 ? pair.Target.Books.Select(t => new KeyValuePair<string, string>(pair.Source.Books[0], t)).ToList()
                    : pair.Source.Books.Zip(pair.Target.Books, (s, t) => new KeyValuePair<string, string>(s, t)).ToList();
                foreach (var link in books)
                {
                    string key = link.Key + "\u001f" + pair.Source.Sheet + "\u001f" + pair.Source.Range + "\u001f" + pair.ValuesOnly + "\u001f" + pair.KeepFormat + "\u001f" + pair.FormatOnly + "\u001f" + pair.PasteLinks; MigrationSourceItem item;
                    if (!items.TryGetValue(key, out item)) items[key] = item = new MigrationSourceItem { Path = link.Key, Sheet = pair.Source.Sheet, Range = pair.Source.Range, Name = "映射第 " + number + " 行", ValuesOnly = pair.ValuesOnly, KeepFormat = pair.KeepFormat, FormatOnly = pair.FormatOnly, PasteLinks = pair.PasteLinks };
                    var m = new MigrationMapping { Pair = pair, Source = item, TargetPath = link.Value, TargetSheet = pair.Target.Sheet, Area = area, Start = CellAddress.Name(area.Row, area.Column), TargetRange = area.Address };
                    m.Result = new MigrationTarget { Path = link.Value, Sheet = pair.Target.Sheet, Start = m.Start, Status = "待预检" };
                    pair.Instances.Add(m); plan.Mappings.Add(m);
                    var sourceBook = FindBook(false, link.Key); var targetBook = FindBook(true, link.Value);
                    string problem = null;
                    if (sourceBook != null && sourceBook.Error.Length > 0) problem = "源工作簿读取失败：" + sourceBook.Error;
                    else if (targetBook != null && targetBook.Error.Length > 0) problem = "目标工作簿读取失败：" + targetBook.Error;
                    else if (sourceBook != null && sourceBook.Sheets.Count > 0 && !sourceBook.Sheets.Contains(pair.Source.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "源工作簿缺少工作表“" + pair.Source.Sheet + "”";
                    else if (targetBook != null && targetBook.Sheets.Count > 0 && !targetBook.Sheets.Contains(pair.Target.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "目标工作簿缺少工作表“" + pair.Target.Sheet + "”";
                    else if (string.Equals(link.Key, link.Value, StringComparison.OrdinalIgnoreCase) || targetPaths.Contains(link.Key)) problem = "源文件不能同时作为目标";
                    else { try { area.Check(1048576, 16384); } catch (Exception error) { problem = error.Message; } }
                    if (problem != null) Block(plan, m, problem); else valid.Add(m);
                }
            }
            // Overlap inside one target sheet is a conflict for both writes.
            var conflicted = new HashSet<MigrationMapping>();
            foreach (var group in valid.GroupBy(m => m.TargetPath + "\u001f" + m.TargetSheet.ToUpperInvariant(), StringComparer.OrdinalIgnoreCase))
            {
                var list = group.ToList();
                for (int i = 0; i < list.Count; i++)
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        var a = list[i].Area; var b = list[j].Area;
                        int top = Math.Max(a.Row, b.Row), left = Math.Max(a.Column, b.Column);
                        int bottom = Math.Min(a.Row + a.Rows, b.Row + b.Rows), right = Math.Min(a.Column + a.Columns, b.Column + b.Columns);
                        if (top >= bottom || left >= right) continue;
                        string overlap = new MigrationArea { Row = top, Column = left, Rows = bottom - top, Columns = right - left }.Address;
                        // Two trigger rows sent to the same target area: the
                        // later row overwrites the earlier; ask, do not block.
                        if (list[i].Source.TriggerRow > 0 && list[j].Source.TriggerRow > 0)
                        {
                            plan.Overrides.Add(list[i].Source.Name + "、" + list[j].Source.Name + " → " + Path.GetFileName(list[i].TargetPath) + "  " + list[i].TargetSheet + "!" + overlap + "，最终保留后者的结果");
                            continue;
                        }
                        Note(list[i], "与" + list[j].Source.Name + "在 " + list[i].TargetSheet + "!" + overlap + " 重叠");
                        Note(list[j], "与" + list[i].Source.Name + "在 " + list[j].TargetSheet + "!" + overlap + " 重叠");
                        conflicted.Add(list[i]); conflicted.Add(list[j]);
                    }
            }
            foreach (var m in conflicted) { m.Result.Status = "冲突"; m.Result.Ready = false; plan.Blocked++; }
            Group(plan, valid.Where(m => !conflicted.Contains(m)).ToList(), plan.ReplaceTargets);
            return plan;
        }
        // 逐行重算: one write per trigger value row, to the workbook its name
        // matched among the added targets. Rows not yet resolved are blocked
        // with the reason; empty, skipped and unmatched rows are left out.
        private void ExpandTrigger(MigrationPlan plan, MigrationPair pair, int number, MigrationArea area, HashSet<string> targetPaths, List<MigrationMapping> valid)
        {
            var trigger = pair.Source.Trigger;
            if (trigger.Error.Length > 0) throw new InvalidOperationException("映射第 " + number + " 行的逐行重算设置无效：" + trigger.Error);
            string sourcePath = pair.Source.Books[0]; var values = MigrationArea.Parse(trigger.Values); int nameColumn = NameColumnIndex(trigger.NameColumn);
            var sourceBook = FindBook(false, sourcePath);
            // Several target rows share this source (one per file): rows are
            // reported once, by the first usable one, and each file is written
            // by the row that owns it.
            var lead = pairs.FirstOrDefault(p => p.Source == pair.Source && p.Enabled && p.Target.Used && RowProblem(p).Length == 0) == pair;
            foreach (string spare in UnusedTargets(pair.Source, pairs, TargetBooks)) if (!plan.Unused.Contains(spare, StringComparer.OrdinalIgnoreCase)) plan.Unused.Add(spare);
            foreach (var row in trigger.Rows)
            {
                string label = "映射第 " + number + " 行 · 触发第 " + row.Row + " 行";
                if (row.Kind == "空值" || row.Kind == "跳过" || row.Kind == "缺失")
                { if (lead) plan.Skipped.Add(label + "：" + (row.Kind == "空值" ? "触发值为空" : row.Kind == "跳过" ? "已设为跳过" : "没有匹配的目标工作簿（" + row.Name + "）")); continue; }
                if (row.Path.Length > 0)
                {
                    var owner = TriggerOwner(pair.Source, row.Path);
                    if (owner == null && lead) plan.Skipped.Add(label + "：" + Path.GetFileName(row.Path) + " 没有对应的目标区域行");
                    if (owner != pair) continue;
                }
                else if (!lead) continue;
                var item = new MigrationSourceItem
                {
                    Path = sourcePath, Sheet = pair.Source.Sheet, Range = pair.Source.Range, ValuesOnly = true, KeepFormat = pair.KeepFormat, Name = label, TriggerRow = row.Row,
                    TriggerCell = MigrationArea.Parse(trigger.Cell).Address, TriggerFrom = CellAddress.Name(row.Row, values.Column), TriggerNameCell = CellAddress.Name(row.Row, nameColumn), TriggerExpect = row.Name
                };
                foreach (var condition in trigger.Extra)
                {
                    string from = "";
                    if (condition.Range.Length > 0)
                    {
                        var range = MigrationArea.Parse(condition.Range);
                        from = range.Rows == 1 && range.Columns == 1 ? CellAddress.Name(range.Row, range.Column) : CellAddress.Name(row.Row, range.Column);
                    }
                    else if (condition.Column.Length > 0) from = CellAddress.Name(row.Row, ColumnIndex(condition.Column));
                    item.TriggerExtra.Add(new MigrationTriggerSet { Cell = MigrationArea.Parse(condition.Cell).Address, From = from, Value = condition.Value });
                }
                string path = row.Path.Length > 0 ? row.Path : "（未确定）";
                var m = new MigrationMapping { Pair = pair, Source = item, TargetPath = path, TargetSheet = pair.Target.Sheet, Area = area, Start = CellAddress.Name(area.Row, area.Column), TargetRange = area.Address };
                m.Result = new MigrationTarget { Path = path, Sheet = pair.Target.Sheet, Start = m.Start, Status = "待预检" };
                pair.Instances.Add(m); plan.Mappings.Add(m);
                string problem = null; var targetBook = row.Path.Length > 0 ? FindBook(true, row.Path) : null;
                if (row.Kind == "多个候选") problem = "触发第 " + row.Row + " 行“" + row.Name + "”有多个同名目标工作簿，请在明细中选择";
                else if (sourceBook != null && sourceBook.Error.Length > 0) problem = "源工作簿读取失败：" + sourceBook.Error;
                else if (targetBook != null && targetBook.Error.Length > 0) problem = "目标工作簿读取失败：" + targetBook.Error;
                else if (targetBook != null && targetBook.Sheets.Count > 0 && !targetBook.Sheets.Contains(pair.Target.Sheet, StringComparer.OrdinalIgnoreCase)) problem = "目标工作簿缺少工作表“" + pair.Target.Sheet + "”";
                else if (string.Equals(path, sourcePath, StringComparison.OrdinalIgnoreCase) || targetPaths.Contains(sourcePath)) problem = "源文件不能同时作为目标";
                else { try { area.Check(1048576, 16384); } catch (Exception error) { problem = error.Message; } }
                if (problem != null) Block(plan, m, problem); else valid.Add(m);
            }
        }
        // Why a mapping row cannot run yet ("" when it can be prechecked).
        internal static string RowProblem(MigrationPair pair, int targetCount)
        {
            if (!pair.Source.Used || !pair.Target.Used) return "未配对";
            if (pair.Source.Books.Count == 0) return "未选源工作簿";
            if (pair.Source.Sheet.Length == 0) return "未选源工作表";
            if (pair.Source.Range.Length == 0) return "未设源区域";
            var source = TryParse(pair.Source.Range); if (source == null) return "源区域无效：" + pair.Source.Range;
            bool byName = pair.Source.Trigger != null;
            if (byName && pair.Source.Books.Count != 1) return "逐行重算只支持一个源工作簿";
            if (byName && targetCount == 0) return "请先添加目标工作簿（按文件名匹配）";
            if (!byName && pair.Target.Books.Count == 0) return "未选目标工作簿";
            if (pair.Target.Sheet.Length == 0) return "未选目标工作表";
            if (pair.Target.Range.Length == 0) return "未设目标区域";
            var target = TargetArea(pair); if (target == null) return "目标区域无效：" + pair.Target.Range;
            if (target.Rows != source.Rows || target.Columns != source.Columns) return "形状不一致：源 " + source.Rows + "×" + source.Columns + "，目标 " + target.Rows + "×" + target.Columns;
            if (!byName && pair.Source.Books.Count > 1 && pair.Source.Books.Count != pair.Target.Books.Count) return "源工作簿 " + pair.Source.Books.Count + " 个与目标 " + pair.Target.Books.Count + " 个无法一一对应";
            return "";
        }
        // Target area of a row: a single start cell takes the source's size.
        internal static MigrationArea TargetArea(MigrationPair pair)
        {
            var target = TryParse(pair.Target.Range); if (target == null) return null;
            var source = TryParse(pair.Source.Range);
            if (target.Rows == 1 && target.Columns == 1 && source != null) return new MigrationArea { Row = target.Row, Column = target.Column, Rows = source.Rows, Columns = source.Columns };
            return target;
        }
        // 目标文件: a one-column range (H2:H30), or the column letters older
        // presets saved (H).
        internal static int NameColumnIndex(string spec)
        {
            spec = (spec ?? "").Trim();
            if (Regex.IsMatch(spec, "^[A-Za-z]{1,3}$")) return ColumnIndex(spec);
            MigrationArea area;
            try { area = MigrationArea.Parse(spec); } catch (Exception) { throw new InvalidOperationException("目标文件区域应为一列区域，例如 H2:H30"); }
            if (area.Columns != 1) throw new InvalidOperationException("目标文件区域只能是一列");
            return area.Column;
        }
        internal static int ColumnIndex(string letters)
        {
            if (!Regex.IsMatch(letters ?? "", "^[A-Za-z]{1,3}$")) throw new InvalidOperationException("文件名列应为列字母，例如 H");
            return MigrationArea.Parse(letters.ToUpperInvariant() + "1").Column;
        }
        internal static MigrationArea TryParse(string text) { try { return MigrationArea.Parse(text); } catch { return null; } }
        internal static string NormalizeAddress(string text)
        {
            // Full-width colon/space typed in an IME are normalized; anything
            // else (commas, missing colons) is left for Parse to reject.
            return (text ?? "").Trim().Replace('：', ':').Replace("　", "").ToUpperInvariant();
        }
        internal static void Block(MigrationPlan plan, MigrationMapping m, string problem)
        { m.Result.Status = "失败"; m.Result.Error = "检查：" + problem; m.Result.Ready = false; plan.Blocked++; }
        internal static void Note(MigrationMapping m, string text)
        { m.Result.Error = m.Result.Error.Length == 0 ? text : m.Result.Error + "；" + text; }
        // 逐行重算 rows whose file has a target row of its own (same source,
        // that file among its workbooks) write there; others fall back to a
        // by-name row of that source, if any.
        private MigrationPair TriggerOwner(MigrationSide source, string path)
        {
            var candidates = pairs.Where(p => p.Source == source && p.Enabled && p.Target.Used && RowProblem(p).Length == 0).ToList();
            return candidates.FirstOrDefault(p => p.Target.Books.Contains(path, StringComparer.OrdinalIgnoreCase)) ?? candidates.FirstOrDefault(p => p.Target.Books.Count == 0);
        }
        // Added targets that no row of this 逐行重算 matched and no ordinary
        // mapping writes to. They are left alone unless the user assigns one.
        internal static List<string> UnusedTargets(MigrationSide side, IEnumerable<MigrationPair> pairs, IEnumerable<MigrationBook> targets)
        {
            var trigger = side.Trigger; if (trigger == null || trigger.Error.Length > 0) return new List<string>();
            var used = new HashSet<string>(trigger.Rows.Where(r => r.Path.Length > 0).Select(r => r.Path), StringComparer.OrdinalIgnoreCase);
            used.UnionWith(pairs.Where(p => p.Source.Trigger == null && p.Target.Used).SelectMany(p => p.Target.Books));
            return targets.Select(b => b.Path).Where(p => !used.Contains(p)).ToList();
        }
        // Group mappings by source range; a target may occur once per pass.
        internal static void Group(MigrationPlan plan, List<MigrationMapping> mappings, bool replace)
        {
            foreach (var bySource in mappings.GroupBy(m => m.Source.Id))
            {
                var passes = new List<MigrationPass>();
                var occurrences = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var m in bySource)
                {
                    int index; occurrences.TryGetValue(m.TargetPath, out index);
                    occurrences[m.TargetPath] = index + 1;
                    if (index == passes.Count)
                    {
                        var source = m.Source;
                        var pass = new MigrationPass { Source = source, Request = new MigrationRequest { SourcePath = source.Path, SourceSheet = source.Sheet, SourceRange = source.Range, AllowEmptySource = plan.AllowEmptySource, ReplaceTargets = replace, ValuesOnly = !source.PasteLinks && (source.ValuesOnly || source.FormatOnly), KeepFormat = !source.PasteLinks && (source.KeepFormat || source.FormatOnly), FormatOnly = !source.PasteLinks && source.FormatOnly, PasteLinks = source.PasteLinks,
                            TriggerCell = source.TriggerCell, TriggerFrom = source.TriggerFrom, TriggerNameCell = source.TriggerNameCell, TriggerExpect = source.TriggerExpect } };
                        pass.Request.TriggerExtra.AddRange(source.TriggerExtra);
                        passes.Add(pass); plan.Passes.Add(pass);
                    }
                    passes[index].Mappings.Add(m); passes[index].Request.Targets.Add(m.Result);
                }
            }
        }
    }
}
