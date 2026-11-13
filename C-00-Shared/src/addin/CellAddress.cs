using System;
using System.Globalization;
using System.Text.RegularExpressions;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Coordinate syntax only; selection limits and write policies belong to callers.
    internal static class CellAddress
    {
        private static readonly Regex Cell = new Regex(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$", RegexOptions.Compiled);
        internal static int ColumnNumber(string letters, int maximum = 16384)
        {
            int value = 0;
            foreach (char c in letters ?? "")
            {
                if (c == '$') continue;
                int digit = c >= 'A' && c <= 'Z' ? c - 'A' + 1 : c >= 'a' && c <= 'z' ? c - 'a' + 1 : 0;
                if (digit == 0 || value > (maximum - digit) / 26) return 0;
                value = value * 26 + digit;
            }
            return value;
        }
        internal static string ColumnLetters(int column)
        {
            string result = "";
            while (column > 0)
            {
                column--;
                result = (char)('A' + column % 26) + result;
                column /= 26;
            }
            return result;
        }
        internal static string Name(int row, int column)
        { return ColumnLetters(column) + row.ToString(CultureInfo.InvariantCulture); }
        internal static int[] Coordinates(string token)
        {
            Match match = Cell.Match(token.Trim()); long row;
            if (!match.Success || !long.TryParse(match.Groups[2].Value, out row) || row > 1048576)
                throw new FormatException("请输入有效的 A1 单元格或矩形区域，例如 A1:B10；多区域以英文逗号分隔。");
            int column = ColumnNumber(match.Groups[1].Value);
            if (column == 0) throw new FormatException("区域列号超过工作表上限。");
            return new[] { (int)row, column };
        }
    }
}
