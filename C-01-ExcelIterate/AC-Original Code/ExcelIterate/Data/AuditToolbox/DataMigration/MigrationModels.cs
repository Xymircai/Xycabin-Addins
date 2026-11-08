using System;
using System.Collections.Generic;
using System.IO;

namespace Excel_Iterate
{
    // A workbook added to the source or the target list, with its sheet names.
    internal sealed class MigrationBook
    {
        internal string Path = "", Error = "";
        internal readonly List<string> Sheets = new List<string>();
        internal string Name { get { return System.IO.Path.GetFileName(Path); } }
    }
    // A source range or a target range: which workbooks, sheet and area.
    // Used = it exists in the 源区域/目标区域 list (an unused side is a gap).
    internal sealed class MigrationSide
    {
        internal readonly List<string> Books = new List<string>();
        internal string Sheet = "", Range = "";
        internal bool Used;
        // Source side only: fill the trigger cell with each trigger value,
        // recalculate, read this range and send it to the workbook named on
        // the same row (逐行重算).
        internal MigrationTrigger Trigger;
        internal MigrationSide Copy() { var copy = new MigrationSide { Sheet = Sheet, Range = Range, Used = Used, Trigger = Trigger == null ? null : Trigger.Copy() }; copy.Books.AddRange(Books); return copy; }
    }
    // One trigger value row and the target workbook its name resolved to.
    internal sealed class MigrationTriggerRow
    {
        internal int Row;
        internal string Value = "", Name = "", Path = "";
        // 精确 / 多个候选 / 缺失 / 已确认 / 跳过 / 空值
        internal string Kind = "";
        // Added targets whose name equals this row's (more than one: 多个候选).
        internal readonly List<KeyValuePair<string, double>> Candidates = new List<KeyValuePair<string, double>>();
        internal bool Ready { get { return Path.Length > 0 && (Kind == "精确" || Kind == "已确认"); } }
    }
    // 触发单元格 2… / 触发值 2…: another cell filled with each row's value
    // from Range (one column, same rows as 触发值 1; a single cell gives
    // every row that cell's value). Column (letters) and Value (fixed) are
    // what earlier presets saved.
    internal sealed class MigrationCondition
    {
        internal string Cell = "", Range = "", Column = "", Value = "";
        internal MigrationCondition Copy() { return new MigrationCondition { Cell = Cell, Range = Range, Column = Column, Value = Value }; }
        internal string Text { get { return Cell + (Range.Length > 0 ? " ← " + Range : Column.Length > 0 ? " ← " + Column + " 列" : " = " + Value); } }
    }
    internal sealed class MigrationTrigger
    {
        internal string Cell = "B2", Values = "", NameColumn = "H";
        internal readonly List<MigrationCondition> Extra = new List<MigrationCondition>();
        // Sheet row -> target path the user chose ("" = skip this row).
        internal readonly Dictionary<int, string> Chosen = new Dictionary<int, string>();
        internal readonly List<MigrationTriggerRow> Rows = new List<MigrationTriggerRow>();
        internal string Error = "", Signature = "";
        internal MigrationTrigger Copy()
        {
            var copy = new MigrationTrigger { Cell = Cell, Values = Values, NameColumn = NameColumn };
            foreach (var chosen in Chosen) copy.Chosen[chosen.Key] = chosen.Value;
            foreach (var condition in Extra) copy.Extra.Add(condition.Copy());
            return copy;
        }
    }
    // One row of 映射: the source range and target range on the same row are
    // paired; dragging in 映射 swaps one side between two rows.
    internal sealed class MigrationPair
    {
        // 写入方式: values or formulas, each with or without the source formats.
        internal const string ModeFormula = "只写公式", ModeValues = "只写数值", ModeValuesFormat = "数值+格式", ModeFormulaFormat = "公式+格式";
        internal const string ModeFormatOnly = "只迁移格式";
        internal static string ModeText(bool values, bool format, bool formatOnly = false) { return formatOnly ? ModeFormatOnly : values ? (format ? ModeValuesFormat : ModeValues) : (format ? ModeFormulaFormat : ModeFormula); }
        internal MigrationSide Source = new MigrationSide(), Target = new MigrationSide();
        internal bool Enabled = true, ValuesOnly;
        // 数值+格式 / 公式+格式: also copy the source cell formats.
        internal bool KeepFormat;
        // 只迁移格式: formats only, the target's content stays.
        internal bool FormatOnly;
        // What the row expanded to in the last preview/run (one per workbook pair).
        internal readonly List<MigrationMapping> Instances = new List<MigrationMapping>();
    }
    // A source range of one workbook, as handed to the single-source engine.
    internal sealed class MigrationSourceItem
    {
        internal readonly string Id = Guid.NewGuid().ToString("N");
        internal string Path = "", Sheet = "", Range = "", Name = "";
        internal bool ValuesOnly, KeepFormat, FormatOnly;
        // 逐行重算: copy TriggerFrom into TriggerCell and recalculate first;
        // TriggerNameCell must still read TriggerExpect.
        internal string TriggerCell = "", TriggerFrom = "", TriggerNameCell = "", TriggerExpect = "";
        internal int TriggerRow;
        // 触发单元格 2… for this row: cell, and the address to copy from or a fixed value.
        internal readonly List<MigrationTriggerSet> TriggerExtra = new List<MigrationTriggerSet>();
        internal string Label { get { return System.IO.Path.GetFileName(Path) + (Name.Length == 0 ? "" : " · " + Name); } }
    }
    // One concrete write: a source range of one workbook into one target workbook.
    internal sealed class MigrationMapping
    {
        internal MigrationPair Pair;
        internal MigrationSourceItem Source;
        internal string TargetPath = "", TargetSheet = "", Start = "", TargetRange = "";
        internal MigrationArea Area;
        internal MigrationTarget Result;
    }
    // A pass is one source range migrated to targets that each appear once,
    // i.e. exactly one run of the existing single-source DataMigrationHost.
    internal sealed class MigrationPass
    {
        internal MigrationSourceItem Source;
        internal MigrationRequest Request;
        internal readonly List<MigrationMapping> Mappings = new List<MigrationMapping>();
    }
    internal sealed class MigrationPlan
    {
        internal readonly List<MigrationPass> Passes = new List<MigrationPass>();
        internal readonly List<MigrationMapping> Mappings = new List<MigrationMapping>();
        internal bool ReplaceTargets, AllowEmptySource, OutputCopies;
        internal readonly Dictionary<string, string> Outputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal int Blocked;
        // 逐行重算 rows left out (empty value, skipped, no target), and rows
        // whose writes land on the same target area (later one wins).
        internal readonly List<string> Skipped = new List<string>(), Overrides = new List<string>();
        // Added target workbooks no 逐行重算 row matched: not written, only reported.
        internal readonly List<string> Unused = new List<string>();
        // How long the precheck took (the base of 预计耗时); when the last
        // precheck or run ended and how long it took.
        internal long PrecheckMilliseconds, RunMilliseconds;
        internal DateTime Finished;
        internal bool Executed;
    }
}
