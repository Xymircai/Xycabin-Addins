using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    // Pure migration rules. In particular, external addresses do NOT follow the
    // physical offset, unlike the existing header-merge rewriter.
    internal sealed class MigrationArea
    {
        internal int Row, Column, Rows, Columns;
        internal string Address { get { return CellAddress.Name(Row, Column) + ":" + CellAddress.Name(Row + Rows - 1, Column + Columns - 1); } }
        internal static MigrationArea Parse(string text)
        {
            Match m = Regex.Match((text ?? "").Trim(), @"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$");
            if (!m.Success) throw new InvalidOperationException("请输入单个矩形区域，例如 A1:D20，或起始单元格 B100。");
            int col = CellAddress.ColumnNumber(m.Groups[1].Value), row, lastRow;
            if (!int.TryParse(m.Groups[2].Value, out row)) throw new InvalidOperationException("行号超出范围。");
            int lastCol = m.Groups[3].Success ? CellAddress.ColumnNumber(m.Groups[3].Value) : col;
            if (!m.Groups[4].Success) lastRow = row;
            else if (!int.TryParse(m.Groups[4].Value, out lastRow)) throw new InvalidOperationException("行号超出范围。");
            var area = new MigrationArea { Row = row, Column = col, Rows = lastRow - row + 1, Columns = lastCol - col + 1 };
            area.Check(1048576, 16384); return area;
        }
        internal void Check(int maxRow, int maxColumn)
        {
            if (Row < 1 || Column < 1 || Rows < 1 || Columns < 1 || (long)Row + Rows - 1 > maxRow || (long)Column + Columns - 1 > maxColumn)
                throw new InvalidOperationException("区域超出工作表实际行列边界。");
        }
        internal bool Contains(MigrationArea other)
        { return other.Row >= Row && other.Column >= Column && (long)other.Row + other.Rows <= (long)Row + Rows && (long)other.Column + other.Columns <= (long)Column + Columns; }
    }

    internal sealed class MigrationName
    {
        internal string Scope = "", Name = "", Formula = "";
        internal bool Visible = true;
        internal string Key { get { return Scope + "\u001f" + Name; } }
    }
    // A formula refers to a sheet the target workbook does not have. The host
    // can then paste that cell's calculated value instead (after asking).
    internal sealed class MigrationMissingSheetException : InvalidOperationException
    {
        internal readonly string Sheet;
        internal MigrationMissingSheetException(string sheet) : base("目标缺少内部引用工作表：" + sheet) { Sheet = sheet; }
    }
    internal sealed class MigrationFormulaContext
    {
        internal string SourcePath = "", SourceSheet = "", TargetSheet = "";
        internal int RowDelta, ColumnDelta, MaxRow = 1048576, MaxColumn = 16384;
        internal HashSet<string> TargetSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal Dictionary<string, string> ExternalFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // File names the workbook links to under more than one path. Only a
        // formula that actually refers to one of them is ambiguous.
        internal HashSet<string> AmbiguousFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal List<MigrationName> Names = new List<MigrationName>();
        internal string MapSheet(string name)
        {
            string mapped = string.Equals(name, SourceSheet, StringComparison.OrdinalIgnoreCase) ? TargetSheet : name;
            if (!TargetSheets.Contains(mapped)) throw new MigrationMissingSheetException(mapped);
            return mapped;
        }
        internal MigrationName FindName(string name, string scope)
        {
            MigrationName global = null;
            foreach (var item in Names)
                if (string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    if (scope.Length != 0 && string.Equals(item.Scope, scope, StringComparison.OrdinalIgnoreCase)) return item;
                    if (item.Scope.Length == 0) global = item;
                }
            return global;
        }
    }
    internal sealed class MigrationFormula
    {
        internal string Text;
        internal readonly List<MigrationName> Dependencies = new List<MigrationName>();
    }

    internal static class DataMigrationEngine
    {
        private static readonly Regex AddressToken = new Regex(@"\G(?:\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?[1-9][0-9]*:\$?[1-9][0-9]*)", RegexOptions.Compiled);
        private static readonly Regex NumberToken = new Regex(@"\G(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?", RegexOptions.Compiled);
        private static bool Ident(char c) { return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\\' || c == '$'; }
        private static string Quote(string text) { return "'" + text.Replace("'", "''") + "'"; }
        internal static bool SupportedFile(string path)
        { string ext = Path.GetExtension(path); return Array.Exists(new[] { ".xlsx", ".xlsm", ".xls", ".et" }, e => string.Equals(e, ext, StringComparison.OrdinalIgnoreCase)); }
        internal static string OutputPath(string target, Func<string, bool> exists)
        {
            string stem = Path.Combine(Path.GetDirectoryName(target), Path.GetFileNameWithoutExtension(target) + "_复制结果"), ext = Path.GetExtension(target);
            for (int i = 1; ; i++) { string candidate = stem + (i == 1 ? "" : "_" + i) + ext; if (!exists(candidate)) return candidate; }
        }
        // Range endpoint order is not semantic: WPS saves $C$6:B5 as B5:$C$6.
        // Keep row/column dollar locks while canonicalizing each axis, and never
        // rewrite text literals or quoted workbook/sheet names in comparisons.
        internal static string ComparisonFormula(string formula, MigrationFormulaContext context)
        {
            string text = Translate(formula, context, context.SourceSheet).Text;
            var output = new StringBuilder(); int i = 0;
            while (i < text.Length)
            {
                if (text[i] == '"')
                {
                    int start = i++;
                    while (i < text.Length) { if (text[i++] != '"') continue; if (i < text.Length && text[i] == '"') { i++; continue; } break; }
                    output.Append(text.Substring(start, i - start)); continue;
                }
                string qualifier; int end;
                if (ReadQualifier(text, i, out qualifier, out end)) { output.Append(text.Substring(i, end - i)); i = end; continue; }
                string address;
                if (ReadAddress(text, i, out address)) { output.Append(CanonicalRange(address)); i += address.Length; continue; }
                output.Append(text[i++]);
            }
            return output.ToString();
        }
        private static string CanonicalRange(string address)
        {
            string[] ends = address.Split(':'); if (ends.Length != 2) return address;
            Match a = Regex.Match(ends[0], @"^(\$?[A-Za-z]+)?(\$?[0-9]+)?$"), b = Regex.Match(ends[1], @"^(\$?[A-Za-z]+)?(\$?[0-9]+)?$");
            if (!a.Success || !b.Success) return address;
            string ac = a.Groups[1].Value, bc = b.Groups[1].Value, ar = a.Groups[2].Value, br = b.Groups[2].Value;
            bool col = ac.Length == 0 || CellAddress.ColumnNumber(ac.Replace("$", "")) <= CellAddress.ColumnNumber(bc.Replace("$", ""));
            bool row = ar.Length == 0 || long.Parse(ar.Replace("$", ""), CultureInfo.InvariantCulture) <= long.Parse(br.Replace("$", ""), CultureInfo.InvariantCulture);
            return (col ? ac : bc) + (row ? ar : br) + ":" + (col ? bc : ac) + (row ? br : ar);
        }
        internal static bool DirectSelfReference(string formula, string sheet, int row, int column)
        {
            string text = formula.Substring(1).Trim(), qualifier; int end;
            if (text.Length == 0) return false;
            if (ReadQualifier(text, 0, out qualifier, out end))
            { if (!string.Equals(qualifier, sheet, StringComparison.OrdinalIgnoreCase)) return false; text = text.Substring(end); }
            return string.Equals(text.Replace("$", ""), CellAddress.Name(row, column), StringComparison.OrdinalIgnoreCase);
        }
        internal static MigrationFormula Translate(string formula, MigrationFormulaContext context, string scope)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') throw new InvalidOperationException("无法识别的公式。");
            var result = new MigrationFormula(); var output = new StringBuilder("="); int i = 1;
            while (i < formula.Length)
            {
                char c = formula[i];
                if (c == '"')
                {
                    int start = i++; bool closed = false;
                    while (i < formula.Length) { if (formula[i++] != '"') continue; if (i < formula.Length && formula[i] == '"') { i++; continue; } closed = true; break; }
                    if (!closed) throw new InvalidOperationException("公式字符串未闭合。"); output.Append(formula.Substring(start, i - start)); continue;
                }
                if (c == '{' || c == '}' || c == '@' || c == '#' || c == '|') throw new InvalidOperationException("暂不复制数组、溢出、错误引用或动态/特殊引用公式。");
                // A qualified token consumes its address as one unit, including a
                // range, so the second endpoint of an external range stays frozen.
                string qualifier; int end;
                if (ReadQualifier(formula, i, out qualifier, out end))
                {
                    string prefix = NormalizeQualifier(qualifier, context);
                    bool external = prefix != null;
                    string originalSheet = qualifier;
                    if (!external)
                    {
                        int bracket = originalSheet.LastIndexOf(']'); if (bracket >= 0) originalSheet = originalSheet.Substring(bracket + 1);
                        prefix = context.MapSheet(originalSheet);
                    }
                    i = end;
                    string address;
                    if (ReadAddress(formula, i, out address))
                    {
                        var named = external ? null : context.FindName(address, originalSheet);
                        if (named != null) { AddDependency(result, context, address, originalSheet); output.Append(Quote(prefix)).Append('!').Append(named.Name); }
                        else output.Append(Quote(prefix)).Append('!').Append(external ? address : Shift(address, context));
                        i += address.Length;
                    }
                    else
                    {
                        int start = i; while (i < formula.Length && Ident(formula[i])) i++;
                        string name = formula.Substring(start, i - start);
                        if (name.Length == 0) throw new InvalidOperationException("无法解析带工作表或工作簿前缀的公式。");
                        if (!external) AddDependency(result, context, name, originalSheet);
                        output.Append(Quote(prefix)).Append('!').Append(name);
                    }
                    continue;
                }
                string localAddress;
                if (ReadAddress(formula, i, out localAddress))
                {
                    var named = context.FindName(localAddress, scope);
                    if (named != null) { AddDependency(result, context, localAddress, scope); output.Append(named.Name); }
                    else output.Append(Shift(localAddress, context));
                    i += localAddress.Length; continue;
                }
                Match number = NumberToken.Match(formula, i);
                if (number.Success) { output.Append(number.Value); i += number.Length; continue; }
                if (Ident(c))
                {
                    int start = i; while (i < formula.Length && Ident(formula[i])) i++;
                    string token = formula.Substring(start, i - start); int next = i; while (next < formula.Length && char.IsWhiteSpace(formula[next])) next++;
                    if (next < formula.Length && formula[next] == '(')
                    {
                        string function = token.ToUpperInvariant();
                        if (function == "INDIRECT" || function == "EVALUATE" || function.StartsWith("_XLFN.", StringComparison.Ordinal))
                            throw new InvalidOperationException("无法可靠复制动态引用或扩展函数：" + token);
                    }
                    else if (!string.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase) && !string.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase)) AddDependency(result, context, token, scope);
                    output.Append(token); continue;
                }
                if (c == '[' || c == ']' || c == '\'' || c == '!') throw new InvalidOperationException("暂不复制结构化表引用或无法解析的外链。");
                if (c == ':') throw new InvalidOperationException("暂不复制三维或跨前缀区域引用。");
                if (!char.IsWhiteSpace(c) && "+-*/^&=<>(),%;".IndexOf(c) < 0) throw new InvalidOperationException("无法解析的公式字符：" + c);
                output.Append(c); i++;
            }
            result.Text = output.ToString(); return result;
        }
        private static void AddDependency(MigrationFormula result, MigrationFormulaContext context, string name, string scope)
        {
            MigrationName found = context.FindName(name, scope);
            if (found == null) throw new InvalidOperationException("无法解析的名称：" + (scope.Length == 0 ? "" : scope + "!") + name);
            if (!result.Dependencies.Contains(found)) result.Dependencies.Add(found);
        }
        private static bool ReadAddress(string text, int index, out string address)
        {
            Match match = AddressToken.Match(text, index); address = null;
            if (!match.Success) return false;
            int end = index + match.Length;
            if (end < text.Length && (Ident(text[end]) || text[end] == '(' || text[end] == '[')) return false;
            // LOG10 (A1) is a function rather than a cell followed by whitespace.
            int next = end; while (next < text.Length && char.IsWhiteSpace(text[next])) next++;
            if (next < text.Length && text[next] == '(') return false;
            address = match.Value; return true;
        }
        private static bool ReadQualifier(string text, int index, out string qualifier, out int end)
        {
            qualifier = null; end = index; int i = index;
            if (text[i] == '\'')
            {
                var body = new StringBuilder(); i++; bool closed = false;
                while (i < text.Length)
                {
                    char c = text[i++]; if (c != '\'') { body.Append(c); continue; }
                    if (i < text.Length && text[i] == '\'') { body.Append('\''); i++; continue; } closed = true; break;
                }
                if (!closed || i >= text.Length || text[i] != '!') throw new InvalidOperationException("带引号的引用前缀无效。");
                qualifier = body.ToString(); end = i + 1;
            }
            else
            {
                if (text[i] == '[') { int close = text.IndexOf(']', i + 1); if (close < 0) return false; i = close + 1; }
                while (i < text.Length && Ident(text[i])) i++;
                if (i == index || i >= text.Length || text[i] != '!') return false;
                qualifier = text.Substring(index, i - index); end = i + 1;
            }
            if (qualifier.Length == 0) throw new InvalidOperationException("引用前缀为空。");
            int bracketEnd = qualifier.LastIndexOf(']'); string sheet = bracketEnd >= 0 ? qualifier.Substring(bracketEnd + 1) : qualifier;
            if (sheet.IndexOf(':') >= 0) throw new InvalidOperationException("暂不复制三维工作表引用。");
            return true;
        }
        // null means this workbook. Relative external paths are resolved against
        // the SOURCE directory, never the directory of a different target.
        private static string NormalizeQualifier(string qualifier, MigrationFormulaContext context)
        {
            int left = qualifier.IndexOf('['), right = qualifier.IndexOf(']');
            if (left < 0)
            {
                if (qualifier.IndexOf('\\') >= 0 || qualifier.IndexOf('/') >= 0 || SupportedFile(qualifier))
                    throw new InvalidOperationException("暂不复制无工作表的外部名称引用：" + qualifier);
                return null;
            }
            if (right <= left || qualifier.IndexOf('[', left + 1) >= 0) throw new InvalidOperationException("外链前缀无效。");
            string book = qualifier.Substring(left + 1, right - left - 1), directory = qualifier.Substring(0, left), sheet = qualifier.Substring(right + 1);
            int numeric; if (int.TryParse(book, out numeric)) throw new InvalidOperationException("无法确定数字编号外链的真实来源：" + qualifier);
            string full;
            if (directory.Length == 0)
            {
                if (!context.ExternalFiles.TryGetValue(book, out full))
                {
                    if (context.AmbiguousFiles.Contains(book)) throw new InvalidOperationException("外链中存在同名不同路径文件，无法可靠判定来源：" + book);
                    if (string.Equals(book, Path.GetFileName(context.SourcePath), StringComparison.OrdinalIgnoreCase)) return null;
                    throw new InvalidOperationException("无法确定外链完整路径：" + book);
                }
            }
            else full = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(context.SourcePath), directory, book));
            if (string.Equals(full, Path.GetFullPath(context.SourcePath), StringComparison.OrdinalIgnoreCase)) return null;
            if (sheet.Length == 0) throw new InvalidOperationException("外部引用缺少工作表。");
            return Path.Combine(Path.GetDirectoryName(full), "[" + Path.GetFileName(full) + "]" + sheet);
        }
        private static string Shift(string address, MigrationFormulaContext context)
        {
            bool wholeRow = Regex.IsMatch(address, @"^\$?[0-9]+:");
            bool wholeColumn = Regex.IsMatch(address, @"^\$?[A-Za-z]+:");
            string[] endpoints = address.Split(':');
            for (int i = 0; i < endpoints.Length; i++)
            {
                Match m = Regex.Match(endpoints[i], @"^(\$?)([A-Za-z]*)(\$?)([0-9]*)$");
                if (!m.Success) throw new InvalidOperationException("无法换算地址：" + address);
                string letters = m.Groups[2].Value, digits = m.Groups[4].Value;
                if (!wholeRow)
                {
                    int col = CellAddress.ColumnNumber(letters);
                    long shifted = col + (m.Groups[1].Value.Length == 0 ? context.ColumnDelta : 0);
                    if (col < 1 || shifted < 1 || shifted > context.MaxColumn) throw new InvalidOperationException("公式列引用越界：" + address);
                    letters = m.Groups[1].Value + CellAddress.ColumnLetters((int)shifted);
                }
                if (!wholeColumn)
                {
                    long row; if (!long.TryParse(digits, out row)) throw new InvalidOperationException("公式行引用无效：" + address);
                    // For a whole-row token the leading dollar occupies group 1.
                    string locked = wholeRow ? m.Groups[1].Value : m.Groups[3].Value;
                    long shifted = row + (locked.Length == 0 ? context.RowDelta : 0);
                    if (row < 1 || shifted < 1 || shifted > context.MaxRow) throw new InvalidOperationException("公式行引用越界：" + address);
                    digits = locked + shifted.ToString(CultureInfo.InvariantCulture);
                }
                endpoints[i] = letters + digits;
            }
            return string.Join(":", endpoints);
        }
        internal static List<MigrationName> RequiredNames(IEnumerable<MigrationFormula> formulas, MigrationFormulaContext context)
        {
            var result = new List<MigrationName>(); var visiting = new HashSet<string>(StringComparer.OrdinalIgnoreCase); var done = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Action<MigrationName> visit = null;
            visit = delegate(MigrationName name)
            {
                if (done.Contains(name.Key)) return;
                if (!visiting.Add(name.Key)) throw new InvalidOperationException("名称存在循环依赖：" + name.Name);
                // RefersTo relative names depend on the active cell at read time;
                // accepting those would silently change their meaning.
                foreach (RefSpan span in FormulaRefRewriter.Scan(name.Formula))
                    if (span.Kind != RefKind.CrossBook && (!span.Column1Absolute || (!span.IsWholeColumn && !span.Row1Absolute) || (span.IsRange && (!span.Column2Absolute || (!span.IsWholeColumn && !span.Row2Absolute)))))
                        throw new InvalidOperationException("暂不复制依赖活动单元格的相对定义名称：" + name.Name);
                MigrationFormula translated = Translate(name.Formula, context, name.Scope);
                foreach (var dependency in translated.Dependencies) visit(dependency);
                result.Add(new MigrationName { Scope = name.Scope.Length == 0 ? "" : context.MapSheet(name.Scope), Name = name.Name, Formula = translated.Text, Visible = name.Visible });
                visiting.Remove(name.Key); done.Add(name.Key);
            };
            foreach (var formula in formulas) foreach (var name in formula.Dependencies) visit(name);
            return result;
        }
    }
}
