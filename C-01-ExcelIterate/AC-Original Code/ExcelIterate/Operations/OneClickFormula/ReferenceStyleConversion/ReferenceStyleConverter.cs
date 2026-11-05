using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal static class ReferenceStyleConverter
    {
        private static readonly Regex Address = new Regex(@"\G(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)", RegexOptions.Compiled);
        private static readonly Regex Number = new Regex(@"\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", RegexOptions.Compiled);
        private static bool Ident(char c) { return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\' || c == '$'; }

        internal static string Convert(string formula, ReferenceStyle style)
        {
            if (style < ReferenceStyle.Absolute || style > ReferenceStyle.Relative) throw new ArgumentException("引用方式无效。");
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return formula;
            var output = new StringBuilder(); int i = 0;
            while (i < formula.Length)
            {
                int start = i;
                if (formula[i] == '"' || formula[i] == '\'')
                {
                    char quote = formula[i++]; bool closed = false;
                    while (i < formula.Length)
                    {
                        if (formula[i++] != quote) continue;
                        if (i < formula.Length && formula[i] == quote) { i++; continue; }
                        closed = true; break;
                    }
                    if (!closed) throw new FormatException("公式引号未闭合。");
                    output.Append(formula, start, i - start); continue;
                }
                // Bracketed table expressions and external workbook names are opaque.
                if (formula[i] == '[')
                {
                    int depth = 0;
                    do
                    {
                        char c = formula[i++];
                        if (c == '\'' && i < formula.Length) { i++; continue; }
                        if (c == '[') depth++; else if (c == ']') depth--;
                    } while (i < formula.Length && depth > 0);
                    if (depth != 0) throw new FormatException("公式方括号未闭合。");
                    output.Append(formula, start, i - start); continue;
                }
                Match address = Address.Match(formula, i);
                int end = address.Success ? i + address.Length : i;
                int next = end; while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                bool before = i == 0 || !Ident(formula[i - 1]);
                bool after = end == formula.Length || !Ident(formula[end]) && formula[end] != '!' && formula[end] != '[';
                if (address.Success && before && after && (next == formula.Length || formula[next] != '('))
                {
                    string rendered;
                    if (Render(address.Value, style, out rendered)) { output.Append(rendered); i = end; continue; }
                }
                Match number = Number.Match(formula, i);
                if (number.Success) { output.Append(number.Value); i += number.Length; continue; }
                if (Ident(formula[i]))
                {
                    while (i < formula.Length && Ident(formula[i])) i++;
                    output.Append(formula, start, i - start); continue;
                }
                output.Append(formula[i++]);
            }
            return output.ToString();
        }

        private static bool Render(string token, ReferenceStyle style, out string rendered)
        {
            bool columnAbsolute = style == ReferenceStyle.Absolute || style == ReferenceStyle.AbsoluteColumnRelativeRow;
            bool rowAbsolute = style == ReferenceStyle.Absolute || style == ReferenceStyle.RelativeColumnAbsoluteRow;
            var output = new StringBuilder();
            foreach (string part in token.Split(':'))
            {
                string value = part.Replace("$", ""); int split = 0;
                while (split < value.Length && char.IsLetter(value[split])) split++;
                string column = value.Substring(0, split), row = value.Substring(split); int col = 0; long rowNumber;
                foreach (char c in column) col = col * 26 + char.ToUpperInvariant(c) - 'A' + 1;
                if (column.Length > 0 && (col < 1 || col > 16384) || row.Length > 0 && (!long.TryParse(row, NumberStyles.None, CultureInfo.InvariantCulture, out rowNumber) || rowNumber < 1 || rowNumber > 1048576))
                { rendered = null; return false; }
                if (output.Length > 0) output.Append(':');
                if (column.Length > 0) output.Append(columnAbsolute ? "$" : "").Append(column);
                if (row.Length > 0) output.Append(rowAbsolute ? "$" : "").Append(row);
            }
            rendered = output.ToString(); return true;
        }
    }
}
