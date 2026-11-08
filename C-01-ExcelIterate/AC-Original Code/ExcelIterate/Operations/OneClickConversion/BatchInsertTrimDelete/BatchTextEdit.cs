using System;
using System.Collections.Generic;

namespace Excel_Iterate
{
    internal enum BatchTextEditOperation { Insert, Trim, Delete }
    internal enum BatchTextCountOrigin { Start, End, Middle }
    internal sealed class BatchTextEditOptions
    {
        internal BatchTextEditOperation Operation;
        internal BatchTextCountOrigin Origin;
        internal int Count, Start = 1, End = 1;
        internal string Text = "";
        internal void Validate()
        {
            if (Operation < BatchTextEditOperation.Insert || Operation > BatchTextEditOperation.Delete || Origin < BatchTextCountOrigin.Start || Origin > BatchTextCountOrigin.Middle)
                throw new ArgumentException("请选择有效的操作和计数方式。");
            if (Operation == BatchTextEditOperation.Insert && Origin == BatchTextCountOrigin.Middle) throw new ArgumentException("插入仅支持从开端或末尾计数。");
            if (Origin == BatchTextCountOrigin.Middle)
            { if (Start < 1 || End < Start) throw new ArgumentException("中间区间从第 1 个字符起算，末尾位置不得小于开头位置。"); }
            else if (Count < 0) throw new ArgumentException("位置或字符数必须为非负整数。");
            if (Operation == BatchTextEditOperation.Insert && string.IsNullOrEmpty(Text)) throw new ArgumentException("请输入要插入的文本。");
        }
    }
    internal static class BatchTextEdit
    {
        internal const int MaximumTextLength = 32767;
        internal static string Transform(string text, BatchTextEditOptions options)
        {
            options.Validate(); text = text ?? "";
            // Offsets are UTF-16 boundaries for Unicode code points, never halves of a surrogate pair.
            var offsets = new List<int>();
            for (int i = 0; i < text.Length; i++)
            { offsets.Add(i); if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++; }
            offsets.Add(text.Length); int count = offsets.Count - 1;
            string result;
            if (options.Operation == BatchTextEditOperation.Insert)
            {
                int at = Math.Min(options.Count, count);
                if (options.Origin == BatchTextCountOrigin.End) at = count - at;
                result = text.Insert(offsets[at], options.Text);
            }
            else
            {
                int first, last;
                if (options.Origin == BatchTextCountOrigin.Middle) { first = Math.Min(options.Start - 1, count); last = Math.Min(options.End, count); }
                else if (options.Origin == BatchTextCountOrigin.End) { first = count - Math.Min(options.Count, count); last = count; }
                else { first = 0; last = Math.Min(options.Count, count); }
                int begin = offsets[first], length = offsets[last] - begin;
                result = options.Operation == BatchTextEditOperation.Trim ? text.Substring(begin, length) : text.Remove(begin, length);
            }
            if (result.Length > MaximumTextLength) throw new InvalidOperationException("处理结果超过单元格 32767 字符上限，请减少插入文本。");
            return result;
        }
        internal static ConversionCellState TransformState(ConversionCellState original, BatchTextEditOptions options)
        {
            if (original.Value == null && !original.IsFormula) return null;
            string before = original.IsFormula ? original.Formula : original.Display ?? "";
            if (!original.IsFormula && CellConversionEngine.IsNumeric(original.Value) && before.Trim().Length > 0 && before.Trim().Trim('#').Length == 0)
                throw new InvalidOperationException("数值显示为 ####，请先调宽该列后重试。");
            string after = Transform(before, options);
            if (after == before) return null;
            var next = original.Copy();
            if (original.IsFormula && KeepsFormulaMarker(before, options) && after.StartsWith("=", StringComparison.Ordinal))
            { next.Formula = after; if (next.NumberFormat == "@") next.NumberFormat = "General"; }
            else { next.IsFormula = false; next.Formula = null; next.Value = after.Length == 0 ? null : (object)after; }
            return next;
        }
        private static bool KeepsFormulaMarker(string text, BatchTextEditOptions options)
        {
            if (options.Operation == BatchTextEditOperation.Insert) return true;
            int count = 0;
            for (int i = 0; i < text.Length; i++) { count++; if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) i++; }
            int first = options.Origin == BatchTextCountOrigin.Middle ? Math.Min(options.Start - 1, count)
                : options.Origin == BatchTextCountOrigin.End ? count - Math.Min(options.Count, count) : 0;
            int last = options.Origin == BatchTextCountOrigin.Middle ? Math.Min(options.End, count)
                : options.Origin == BatchTextCountOrigin.End ? count : Math.Min(options.Count, count);
            return options.Operation == BatchTextEditOperation.Trim ? first == 0 && last > 0 : first != 0 || last == 0;
        }
    }
}
