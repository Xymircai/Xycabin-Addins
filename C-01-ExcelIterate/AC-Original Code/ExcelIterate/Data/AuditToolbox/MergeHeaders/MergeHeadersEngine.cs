using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel_Iterate
{
    // =====================================================================
    // 跨工作簿表头匹配合并 —— 纯逻辑层（零 COM）
    //
    // 职责：表头归一化与匹配、缺列收集、写入计划分块、执行报告渲染。
    // 所有 COM 读写在 MergeHeadersHost.cs，界面在 MergeHeadersForm.cs。
    //
    // 语义定稿（与用户逐条确认）：
    //   1. 多源 → 单目标，按表头「文本」匹配，不按物理列位置
    //   2. 目标缺失的列名不自动补建，先告知用户并由用户确认
    //   3. 每一列独立追加到「该列已有数据的末尾 + 1」
    //   4. 公式保留并重写同表引用（相对列按映射表、相对行按 delta）；
    //      跨表/跨簿引用原文不变
    // =====================================================================

    internal static class MergeLimits
    {
        internal const int MaxSources = 20;
        internal const int MaxCells = 500000;
        internal const int MaxRow = 1048576;
        internal const int MaxColumn = 16384;
    }

    internal static class MergeKey
    {
        /// <summary>匹配键：去首尾空格（比较时再忽略大小写）。</summary>
        internal static string Normalize(string text)
        { return text == null ? "" : text.Trim(); }
    }

    internal sealed class MergeColumn
    {
        internal int Number;
        internal string Letter = "";
        internal string Text = "";       // 显示原文
        internal string Key = "";        // 归一化后的匹配键
        internal string Signature = "";  // 复核用（值/公式/文本）
    }

    /// <summary>一张工作表「表头行」的完整快照。</summary>
    internal sealed class MergeHeaderSet
    {
        internal string Path = "";
        internal string SheetName = "";
        internal int HeaderRow;
        internal int FirstColumn;
        internal int LastColumn;
        internal int LastUsedRow;        // UsedRange 底行
        internal readonly List<MergeColumn> Columns = new List<MergeColumn>();
        internal readonly List<string> Duplicates = new List<string>();
        internal readonly List<int> Blanks = new List<int>();
        private readonly Dictionary<string, int> index =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        /// <summary>建索引并统计重复/空表头。每次填充 Columns 后调用。</summary>
        internal void Seal()
        {
            index.Clear();
            Duplicates.Clear();
            Blanks.Clear();
            for (int i = 0; i < Columns.Count; i++)
            {
                MergeColumn column = Columns[i];
                if (column.Key.Length == 0) { Blanks.Add(column.Number); continue; }
                if (index.ContainsKey(column.Key))
                {
                    if (!Duplicates.Contains(column.Text)) Duplicates.Add(column.Text);
                    continue;
                }
                index.Add(column.Key, column.Number);
            }
        }

        /// <summary>按归一化列名查物理列号；0 表示不存在。</summary>
        internal int ColumnOf(string key)
        {
            int number;
            return index.TryGetValue(key, out number) ? number : 0;
        }

        internal MergeColumn At(int column)
        {
            for (int i = 0; i < Columns.Count; i++) if (Columns[i].Number == column) return Columns[i];
            return null;
        }

        /// <summary>返回阻断级错误文本；null 表示通过。</summary>
        internal string Validate(string role)
        {
            if (Columns.Count == 0) return role + "的表头行没有可识别的列。";
            if (Blanks.Count > 0)
            {
                List<string> letters = new List<string>();
                foreach (int number in Blanks)
                {
                    MergeColumn column = At(number);
                    letters.Add(column == null ? number.ToString(CultureInfo.InvariantCulture) : column.Letter);
                }
                return role + "的表头行存在空列名：" + string.Join("、", letters.ToArray())
                    + "。请补全列名后重试。";
            }
            if (Duplicates.Count > 0)
                return role + "的表头行存在重复列名：" + string.Join("、", Duplicates.ToArray())
                    + "。列名需唯一（忽略大小写、忽略首尾空格）才能匹配。";
            return null;
        }
    }

    internal sealed class MergeWorkbookSpec
    {
        internal string Path = "";
        internal string DisplayName = "";
        internal string SheetName = "";      // 空 = 第一张可见工作表
        internal int HeaderRow = 1;
        internal List<string> SheetNames = new List<string>();
        internal string SheetError = "";
        internal bool SheetNamesLoaded;

        internal static string Display(string path)
        {
            try { return System.IO.Path.GetFileName(path); }
            catch { return path ?? ""; }
        }
    }

    internal sealed class MergeRequest
    {
        internal MergeWorkbookSpec Target = new MergeWorkbookSpec();
        internal readonly List<MergeWorkbookSpec> Sources = new List<MergeWorkbookSpec>();

        /// <summary>纯 IO 层校验（不打开工作簿）。返回错误文本；null 表示通过。</summary>
        internal string Validate()
        {
            if (Target == null || string.IsNullOrEmpty(Target.Path))
                return "请先选择目标工作簿。";
            if (!System.IO.File.Exists(Target.Path)) return "目标工作簿不存在：" + Target.Path;
            if (Target.HeaderRow < 1 || Target.HeaderRow > MergeLimits.MaxRow)
                return "目标表头行需在 1 到 " + MergeLimits.MaxRow + " 之间。";
            if (Sources.Count == 0) return "请至少添加一个来源工作簿。";
            if (Sources.Count > MergeLimits.MaxSources)
                return "来源工作簿最多 " + MergeLimits.MaxSources + " 个，当前 " + Sources.Count + " 个。";
            List<string> seen = new List<string>();
            string targetFull;
            try { targetFull = System.IO.Path.GetFullPath(Target.Path); }
            catch { return "目标工作簿路径无效。"; }
            seen.Add(targetFull.ToLowerInvariant());
            for (int i = 0; i < Sources.Count; i++)
            {
                MergeWorkbookSpec source = Sources[i];
                if (string.IsNullOrEmpty(source.Path)) return "第 " + (i + 1) + " 个来源未选择工作簿。";
                if (!System.IO.File.Exists(source.Path)) return "来源工作簿不存在：" + source.Path;
                if (source.HeaderRow < 1 || source.HeaderRow > MergeLimits.MaxRow)
                    return "第 " + (i + 1) + " 个来源的表头行需在 1 到 " + MergeLimits.MaxRow + " 之间。";
                string full;
                try { full = System.IO.Path.GetFullPath(source.Path); }
                catch { return "第 " + (i + 1) + " 个来源路径无效。"; }
                string lowered = full.ToLowerInvariant();
                if (seen.Contains(lowered)) return "目标工作簿与来源不能重复，也不能重复添加同一来源：" + source.Path;
                seen.Add(lowered);
            }
            return null;
        }
    }

    /// <summary>源列 → 目标列的一条映射，以及该列待写内容。</summary>
    internal sealed class MergeColumnMap
    {
        internal int SourceColumn;
        internal int TargetColumn;
        internal string SourceLetter = "";
        internal string TargetLetter = "";
        internal string Header = "";
        internal int RowCount;
        internal int SourceFirstRow;      // 源数据首行 = 源表头行 + 1
        internal int TargetStartRow;      // 目标写入首行 = 该目标列末行 + 1
        internal int RowDelta { get { return TargetStartRow - SourceFirstRow; } }
        internal object[] Cells;          // 公式格为带前导 = 的字符串，常量格为原值
        internal int FormulaCount;
        internal readonly List<string> Notes = new List<string>();
    }

    internal sealed class MergeSourcePlan
    {
        internal string Path = "";
        internal string DisplayName = "";
        internal string SheetName = "";
        internal int HeaderRow;
        internal int RowCount;
        internal readonly List<MergeColumnMap> Maps = new List<MergeColumnMap>();
        internal readonly List<string> Missing = new List<string>();
        internal string Note = "";
        internal bool StartRowsDiffer;    // 同一来源各列起始行不一致（决策 4 的固有风险）
    }

    internal sealed class MergeMatchResult
    {
        internal readonly List<MergeSourcePlan> Sources = new List<MergeSourcePlan>();
        internal readonly List<string> Blocking = new List<string>();
        internal readonly List<string> Warnings = new List<string>();
        internal bool HasMissing
        {
            get
            {
                foreach (MergeSourcePlan plan in Sources) if (plan.Missing.Count > 0) return true;
                return false;
            }
        }
        internal int MappedColumns
        {
            get
            {
                int total = 0;
                foreach (MergeSourcePlan plan in Sources) total += plan.Maps.Count;
                return total;
            }
        }
        internal int TotalCells
        {
            get
            {
                int total = 0;
                foreach (MergeSourcePlan plan in Sources)
                    foreach (MergeColumnMap map in plan.Maps) total += map.RowCount;
                return total;
            }
        }
    }

    internal sealed class MergeWriteBlock
    {
        internal int FirstColumn;
        internal int ColumnCount;
        internal int StartRow;
        internal int RowCount;
        internal object[,] Cells;
        internal bool AnyFormula;
    }

    internal sealed class MergeWritePlan
    {
        internal readonly List<MergeWriteBlock> Blocks = new List<MergeWriteBlock>();
        internal int TotalCells;
        internal int FormulaCells;
        internal string TargetSheetName = "";
    }

    internal sealed class MergeRewriteSample
    {
        internal string Source = "";
        internal string Cell = "";
        internal string Before = "";
        internal string After = "";
        internal string Note = "";
    }

    internal sealed class MergeReport
    {
        internal int Sources;
        internal int MappedColumns;
        internal int SkippedColumns;
        internal int WrittenCells;
        internal int FormulaCells;
        internal int RewrittenFormulas;
        internal readonly List<string> Issues = new List<string>();
        internal readonly List<MergeRewriteSample> Samples = new List<MergeRewriteSample>();
        internal string BackupPath = "";
        internal bool RolledBack;
        internal string RecoveryLog = "";
        internal string Text = "";
    }

    internal static class MergeMatcher
    {
        /// <summary>
        /// 用已读好的表头快照做匹配。target/sources 必须已 Seal()。
        /// 只做纯逻辑，不碰 COM；阻断项写入 Blocking。
        /// </summary>
        internal static MergeMatchResult Match(MergeHeaderSet target, IList<MergeHeaderSet> sources)
        {
            MergeMatchResult result = new MergeMatchResult();
            string targetError = target.Validate("目标工作簿");
            if (targetError != null) { result.Blocking.Add(targetError); return result; }

            for (int i = 0; i < sources.Count; i++)
            {
                MergeHeaderSet source = sources[i];
                MergeSourcePlan plan = new MergeSourcePlan();
                plan.Path = source.Path;
                plan.DisplayName = MergeWorkbookSpec.Display(source.Path);
                plan.SheetName = source.SheetName;
                plan.HeaderRow = source.HeaderRow;
                plan.RowCount = Math.Max(0, source.LastUsedRow - source.HeaderRow);
                result.Sources.Add(plan);

                string error = source.Validate("来源 " + plan.DisplayName);
                if (error != null) { result.Blocking.Add(error); continue; }
                if (plan.RowCount == 0)
                {
                    plan.Note = "表头行下方没有数据行，已跳过。";
                    result.Warnings.Add(plan.DisplayName + "：" + plan.Note);
                    continue;
                }

                for (int c = 0; c < source.Columns.Count; c++)
                {
                    MergeColumn sourceColumn = source.Columns[c];
                    int targetColumn = target.ColumnOf(sourceColumn.Key);
                    if (targetColumn == 0) { plan.Missing.Add(sourceColumn.Text); continue; }
                    MergeColumnMap map = new MergeColumnMap();
                    map.SourceColumn = sourceColumn.Number;
                    map.TargetColumn = targetColumn;
                    map.SourceLetter = sourceColumn.Letter;
                    map.TargetLetter = CellAddress.ColumnLetters(targetColumn);
                    map.Header = sourceColumn.Text;
                    map.RowCount = plan.RowCount;
                    map.SourceFirstRow = source.HeaderRow + 1;
                    plan.Maps.Add(map);
                }
                if (plan.Missing.Count > 0)
                    result.Warnings.Add(plan.DisplayName + "：有 " + plan.Missing.Count
                        + " 列在目标工作簿中不存在，将不写入。");
            }
            return result;
        }
    }

    internal static class MergePlanner
    {
        /// <summary>把已填好 Cells 的映射列表切成写入块（相邻且起始行/行数相同的列合并）。</summary>
        internal static MergeWritePlan Build(IList<MergeColumnMap> maps, int headerRow)
        {
            MergeWritePlan plan = new MergeWritePlan();
            if (maps == null || maps.Count == 0) return plan;

            List<MergeColumnMap> ordered = new List<MergeColumnMap>(maps);
            ordered.Sort(delegate(MergeColumnMap a, MergeColumnMap b)
            { return a.TargetColumn.CompareTo(b.TargetColumn); });

            int i = 0;
            while (i < ordered.Count)
            {
                int j = i;
                while (j + 1 < ordered.Count
                    && ordered[j + 1].TargetColumn == ordered[j].TargetColumn + 1
                    && ordered[j + 1].TargetStartRow == ordered[i].TargetStartRow
                    && ordered[j + 1].RowCount == ordered[i].RowCount) j++;

                int columnCount = j - i + 1;
                int rows = ordered[i].RowCount;
                MergeWriteBlock block = new MergeWriteBlock();
                block.FirstColumn = ordered[i].TargetColumn;
                block.ColumnCount = columnCount;
                block.StartRow = ordered[i].TargetStartRow;
                block.RowCount = rows;
                block.Cells = new object[rows, columnCount];
                for (int r = 0; r < rows; r++)
                    for (int k = 0; k <= j - i; k++)
                    {
                        MergeColumnMap map = ordered[i + k];
                        object value = map.Cells != null && r < map.Cells.Length ? map.Cells[r] : null;
                        block.Cells[r, k] = value;
                        if (IsFormula(value)) { block.AnyFormula = true; plan.FormulaCells++; }
                    }
                plan.TotalCells += rows * columnCount;
                plan.Blocks.Add(block);
                i = j + 1;
            }
            return plan;
        }

        internal static bool IsFormula(object value)
        {
            string text = value as string;
            return text != null && text.StartsWith("=", StringComparison.Ordinal);
        }

        /// <summary>总量保护：超过阈值直接拒绝，避免一次性改出无法核对的大范围。</summary>
        internal static string Guard(MergeWritePlan plan)
        {
            if (plan == null || plan.Blocks.Count == 0) return "没有可写入的内容。";
            if (plan.TotalCells > MergeLimits.MaxCells)
                return "本次将写入 " + plan.TotalCells.ToString(CultureInfo.InvariantCulture)
                    + " 个单元格，超过单次上限 " + MergeLimits.MaxCells.ToString(CultureInfo.InvariantCulture)
                    + " 个。请拆分来源或减少列数后重试。";
            foreach (MergeWriteBlock block in plan.Blocks)
            {
                int last = block.StartRow + block.RowCount - 1;
                if (block.RowCount > 0 && last > MergeLimits.MaxRow)
                    return "写入将超出工作表最后一行（第 " + MergeLimits.MaxRow.ToString(CultureInfo.InvariantCulture)
                        + " 行）。请检查目标列已有数据的长度。";
                if (block.FirstColumn + block.ColumnCount - 1 > MergeLimits.MaxColumn)
                    return "写入将超出工作表最后一列。";
            }
            return null;
        }
    }

    internal static class MergePlanText
    {
        /// <summary>缺列告知（需求 3：必须明确告诉用户哪几列缺失）。</summary>
        internal static string MissingText(MergeMatchResult match)
        {
            if (!match.HasMissing) return "";
            StringBuilder text = new StringBuilder();
            text.AppendLine("以下列名在目标工作簿中不存在，本次不会写入这些列：");
            text.AppendLine();
            foreach (MergeSourcePlan plan in match.Sources)
            {
                if (plan.Missing.Count == 0) continue;
                text.AppendLine("· " + plan.DisplayName + "（" + plan.SheetName + "）：");
                List<string> shown = new List<string>();
                for (int i = 0; i < plan.Missing.Count && i < 20; i++) shown.Add(plan.Missing[i]);
                text.AppendLine("    " + string.Join("、", shown.ToArray())
                    + (plan.Missing.Count > 20 ? " …等共 " + plan.Missing.Count + " 列" : ""));
            }
            text.AppendLine();
            text.AppendLine("目标工作簿不会新增列。若这些列本应写入，请先核对目标表头列名是否一致。");
            return text.ToString();
        }

        /// <summary>执行前的二次确认全文。</summary>
        internal static string ConfirmationText(MergeRequest request, MergeMatchResult match)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("目标：" + MergeWorkbookSpec.Display(request.Target.Path)
                + "（" + SheetText(request.Target) + "，表头第 " + request.Target.HeaderRow + " 行）");
            text.AppendLine();
            foreach (MergeSourcePlan plan in match.Sources)
            {
                text.AppendLine("来源：" + plan.DisplayName + "（" + plan.SheetName
                    + "，表头第 " + plan.HeaderRow + " 行）");
                text.AppendLine("    " + plan.Maps.Count + " 列匹配、" + plan.Missing.Count + " 列缺失、"
                    + plan.RowCount + " 行数据");
                if (plan.StartRowsDiffer)
                    text.AppendLine("    ⚠ 该来源各列的目标起始行不一致，写入后同一条记录会横向错开（按各列独立追加的既定语义）。");
                for (int i = 0; i < plan.Maps.Count && i < 12; i++)
                {
                    MergeColumnMap map = plan.Maps[i];
                    text.AppendLine("    " + map.Header + "：" + map.SourceLetter + " → " + map.TargetLetter
                        + "，写入第 " + map.TargetStartRow + " 行起 " + map.RowCount + " 行");
                }
                if (plan.Maps.Count > 12)
                    text.AppendLine("    …其余 " + (plan.Maps.Count - 12) + " 列见上方列表");
                text.AppendLine();
            }
            if (match.HasMissing) { text.AppendLine(MissingText(match)); text.AppendLine(); }
            text.AppendLine("共 " + match.MappedColumns + " 列、" + match.TotalCells + " 个单元格将被写入。");
            text.AppendLine("写入前会备份目标工作簿；完成后可点「取消操作」还原。");
            return text.ToString();
        }

        internal static string SheetText(MergeWorkbookSpec spec)
        {
            return string.IsNullOrEmpty(spec.SheetName) ? "第一张工作表" : spec.SheetName;
        }

        /// <summary>执行报告正文。</summary>
        internal static string ReportText(MergeReport report, string targetName)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("跨工作簿表头匹配合并 · 执行报告");
            text.AppendLine();
            text.AppendLine("目标工作簿：" + targetName);
            text.AppendLine("来源： " + report.Sources + " 个；匹配列 " + report.MappedColumns
                + " 列；跳过 " + report.SkippedColumns + " 列");
            text.AppendLine("写入： " + report.WrittenCells + " 个单元格，其中公式 "
                + report.FormulaCells + " 个（重写 " + report.RewrittenFormulas + " 个）");
            if (report.BackupPath.Length > 0)
                text.AppendLine("备份： " + report.BackupPath);
            if (report.RolledBack)
                text.AppendLine("状态： 已还原到写入前。");
            if (report.RecoveryLog.Length > 0)
                text.AppendLine("恢复日志： " + report.RecoveryLog);
            if (report.Issues.Count > 0)
            {
                text.AppendLine();
                text.AppendLine("未完整重映射、已保持原样的引用（" + report.Issues.Count + " 类）：");
                for (int i = 0; i < report.Issues.Count && i < 30; i++) text.AppendLine("  · " + report.Issues[i]);
                if (report.Issues.Count > 30) text.AppendLine("  …等共 " + report.Issues.Count + " 类");
            }
            if (report.Samples.Count > 0)
            {
                text.AppendLine();
                text.AppendLine("公式重写样本（最多 50 条）：");
                text.AppendLine("  来源 | 源单元格 | 原公式 → 新公式");
                for (int i = 0; i < report.Samples.Count && i < 50; i++)
                {
                    MergeRewriteSample sample = report.Samples[i];
                    text.AppendLine("  " + sample.Source + " | " + sample.Cell + " | "
                        + sample.Before + "  →  " + sample.After
                        + (sample.Note.Length > 0 ? "（" + sample.Note + "）" : ""));
                }
            }
            return text.ToString();
        }
    }
}
