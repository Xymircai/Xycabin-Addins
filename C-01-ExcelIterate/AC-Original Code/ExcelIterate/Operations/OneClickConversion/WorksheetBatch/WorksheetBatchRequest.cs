using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal enum WorksheetBatchSource { OpenWorkbooks, Folder, Files }
    internal enum WorksheetBatchTextStyle { Keep, Bold, Italic, Underline, Strikethrough }
    internal enum WorksheetBatchHorizontal
    {
        General = 1, Left = -4131, Center = -4108, Right = -4152, Justify = -4130, Distributed = -4117
    }
    internal enum WorksheetBatchVertical { Top = -4160, Center = -4108, Bottom = -4107 }

    internal sealed class WorksheetBatchRequest
    {
        private static readonly HashSet<string> Extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { ".xlsx", ".xls", ".xlsm", ".xlsb", ".et" };
        internal const string FileFilter = "工作簿 (*.xlsx;*.xls;*.xlsm;*.xlsb;*.et)|*.xlsx;*.xls;*.xlsm;*.xlsb;*.et";
        internal WorksheetBatchSource Source;
        internal string Folder, SheetName, Address;
        internal string[] Files;
        internal string Find, Replacement = "";
        internal bool MatchCase, WholeCell;
        internal WorksheetBatchTextStyle TextStyle;
        internal string FontName;
        internal WorksheetBatchHorizontal? Horizontal;
        internal WorksheetBatchVertical? Vertical;
        internal int? FontColor, FillColor;
        internal bool ClearFill, ClearFormatting;
        internal string NumberFormat;

        internal bool ReplacesText { get { return Find != null; } }
        private bool SetsFormat
        {
            get
            {
                return TextStyle != WorksheetBatchTextStyle.Keep || FontName != null || Horizontal.HasValue || Vertical.HasValue ||
                    FontColor.HasValue || FillColor.HasValue || ClearFill || NumberFormat != null;
            }
        }
        internal bool FormatsCells { get { return ClearFormatting || SetsFormat; } }

        internal IList<string> Validate()
        {
            if (string.IsNullOrWhiteSpace(SheetName)) throw new ArgumentException("请填写要处理的工作表名称。");
            if (ReplacesText && Find.Length == 0) throw new ArgumentException("请填写查找文字。");
            if (ReplacesText && Replacement.Length > CellConversionEngine.MaximumCellTextLength) throw new ArgumentException("替换文字超过单元格的字符上限。");
            if (ClearFormatting && SetsFormat) throw new ArgumentException("请在清除格式和设置格式中选择一种。");
            if (!ReplacesText && !FormatsCells) throw new ArgumentException("请至少设置一个文字替换或单元格格式选项。");
            return ConversionRangeAddress.Cells(Address);
        }

        internal IList<string> FilePaths()
        {
            if (Source == WorksheetBatchSource.OpenWorkbooks) return new string[0];
            IEnumerable<string> candidates;
            if (Source == WorksheetBatchSource.Folder)
            {
                if (!Directory.Exists(Folder)) throw new ArgumentException("请选择可访问的工作簿文件夹。");
                candidates = Directory.EnumerateFiles(Folder);
            }
            else candidates = Files ?? new string[0];
            var paths = candidates.Where(path => Extensions.Contains(Path.GetExtension(path)) && !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToList();
            if (paths.Count == 0) throw new InvalidOperationException("所选位置没有可处理的工作簿。");
            return paths;
        }

        internal string Replace(string text)
        {
            if (text == null) return null;
            StringComparison comparison = MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            if (WholeCell) return string.Equals(text, Find, comparison) ? Replacement : text;
            int start = 0, match = text.IndexOf(Find, comparison);
            if (match < 0) return text;
            var result = new System.Text.StringBuilder();
            do
            {
                result.Append(text, start, match - start).Append(Replacement);
                if (result.Length > CellConversionEngine.MaximumCellTextLength) throw new InvalidOperationException("替换结果超过单元格的字符上限。");
                start = match + Find.Length;
                match = text.IndexOf(Find, start, comparison);
            } while (match >= 0);
            result.Append(text, start, text.Length - start);
            if (result.Length > CellConversionEngine.MaximumCellTextLength) throw new InvalidOperationException("替换结果超过单元格的字符上限。");
            return result.ToString();
        }

        private static readonly Regex PlainNumber = new Regex(@"^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?\z",
            RegexOptions.CultureInvariant);

        internal static string SearchText(object value)
        {
            if (value is string) return (string)value;
            if (!CellConversionEngine.IsNumeric(value)) return null;
            double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (double.IsNaN(number) || double.IsInfinity(number)) return null;
            string raw = value is double ? number.ToString("R", CultureInfo.InvariantCulture) :
                value is float ? ((float)value).ToString("R", CultureInfo.InvariantCulture) :
                Convert.ToString(value, CultureInfo.InvariantCulture);
            return ExpandExponent(raw);
        }

        internal static object ReplacementValue(object original, string text)
        {
            if (text.Length == 0) return null;
            double number;
            if (!(original is string) && PlainNumber.IsMatch(text) &&
                double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
                !double.IsInfinity(number) && !double.IsNaN(number) &&
                NormalizeDecimal(text) == NormalizeDecimal(ExpandExponent(number.ToString("G15", CultureInfo.InvariantCulture))))
                return number;
            return text;
        }

        private static string NormalizeDecimal(string text)
        { return text.IndexOf('.') < 0 ? text : text.TrimEnd('0').TrimEnd('.'); }

        private static string ExpandExponent(string text)
        {
            int exponent = text.IndexOfAny(new[] { 'E', 'e' });
            if (exponent < 0) return text;
            int shift = int.Parse(text.Substring(exponent + 1), CultureInfo.InvariantCulture);
            string sign = text[0] == '-' ? "-" : "";
            string mantissa = text.Substring(sign.Length, exponent - sign.Length);
            int point = mantissa.IndexOf('.');
            string digits = mantissa.Replace(".", "");
            int position = (point < 0 ? digits.Length : point) + shift;
            if (position <= 0) return sign + "0." + new string('0', -position) + digits;
            if (position >= digits.Length) return sign + digits + new string('0', position - digits.Length);
            return sign + digits.Insert(position, ".");
        }

        internal static bool IsTemporalFormat(string format)
        {
            for (int i = 0; i < format.Length; i++)
            {
                char token = char.ToLowerInvariant(format[i]);
                if (token == '\\' || token == '_' || token == '*') { i++; continue; }
                if (token == '"')
                {
                    int end = format.IndexOf('"', i + 1);
                    if (end < 0) return true;
                    i = end; continue;
                }
                if (token == '[')
                {
                    int end = format.IndexOf(']', i + 1);
                    if (end < 0) return true;
                    string section = format.Substring(i + 1, end - i - 1).ToLowerInvariant();
                    if (section == "h" || section == "hh" || section == "m" || section == "mm" ||
                        section == "s" || section == "ss" || section == "$-f800" || section == "$-f400")
                        return true;
                    i = end; continue;
                }
                if (token == 'g' && i + 7 <= format.Length &&
                    string.Compare(format, i, "General", 0, 7, StringComparison.OrdinalIgnoreCase) == 0)
                { i += 6; continue; }
                if (token == 'y' || token == 'd' || token == 'h' || token == 's' || token == 'm' || token == 'g' ||
                    (token == 'e' && (i + 1 == format.Length || (format[i + 1] != '+' && format[i + 1] != '-'))))
                    return true;
            }
            return false;
        }
    }
}
