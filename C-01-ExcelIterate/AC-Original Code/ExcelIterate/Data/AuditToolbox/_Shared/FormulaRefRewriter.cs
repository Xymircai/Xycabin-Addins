using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    // =====================================================================
    // 公式引用重写（纯字符串处理，不碰 COM）
    //
    // 用途：把源工作簿某列的公式搬到目标工作簿的另一列时，按「Excel 原生
    //       复制粘贴语义」重写其中的同表引用：
    //
    //         相对列引用（A1）→ 按列映射表换算
    //         绝对列引用（$A1）→ 列不动
    //         相对行引用（A1）→ 按 delta 偏移
    //         绝对行引用（A$1）→ 行不动
    //         跨表 / 跨簿引用（含 ! 前缀）→ 地址部分：列保持不变（不做映射）、
    //             行照常偏移；绝对锁定（$A$1）仍完全不变
    //
    // 为什么不用现成的 FormulaRefParser：
    //   它的 Tokenize 丢弃运算符、且不记录 token 在原串中的位置，无法还原
    //   回公式字符串。这里另写一套「带位置的 span 扫描 + 从后往前反向替换」，
    //   仅复用它的引号状态机规则（"" 与 '' 转义、! 与 : 不作分隔符）。
    // =====================================================================

    internal enum RefIssueKind { UnmappedColumn, ColumnOutOfRange, RowOutOfRange }

    /// <summary>引用分类：本表 / 本工作簿跨表 / 跨工作簿。</summary>
    internal enum RefKind { Local, CrossSheet, CrossBook }

    /// <summary>一处未能完整重映射的引用（保持原样，交由报告告知用户）。</summary>
    internal sealed class RefIssue
    {
        internal RefIssueKind Kind;
        internal int Column;
        internal int Row;
        internal string Fragment = "";

        internal string Describe()
        {
            switch (Kind)
            {
                case RefIssueKind.UnmappedColumn:
                    return "引用的列未参与本次搬移，已保持原样：" + Fragment;
                case RefIssueKind.ColumnOutOfRange:
                    return "换算后的列号超出工作表范围，已保持原样：" + Fragment;
                default:
                    return "换算后的行号超出工作表范围，已保持原样：" + Fragment;
            }
        }

        /// <summary>聚合去重用的键（同一公式同一片段只报一次）。</summary>
        internal string Key()
        {
            return ((int)Kind).ToString(CultureInfo.InvariantCulture) + "|" + Fragment;
        }
    }

    /// <summary>公式里一处「本地引用」的位置与构成。</summary>
    internal sealed class RefSpan
    {
        internal int Start;
        internal int Length;
        internal string Text = "";
        internal int Column1;
        internal int Row1;
        internal int Column2;
        internal int Row2;
        internal bool Column1Absolute;
        internal bool Row1Absolute;
        internal bool Column2Absolute;
        internal bool Row2Absolute;
        internal bool IsRange;
        internal bool IsWholeColumn;
        internal RefKind Kind;        // 本表 / 跨表 / 跨簿三分类
        internal int PrefixStart;     // 跨表/跨簿引用的前缀起始（含前导 '），Local 时无意义
        internal int PrefixLength;    // 前缀长度（到 ! 之前，不含 !），Local 时无意义
    }

    /// <summary>源列号 → 目标列号（0 表示未映射）。</summary>
    internal interface IMergeColumnMapper
    {
        int MapColumn(int sourceColumn);
    }

    internal sealed class RewriteResult
    {
        internal string Text = "";
        internal bool Changed;
        internal readonly List<RefIssue> Issues = new List<RefIssue>();
    }

    /// <summary>重写上下文：物理列差 + 源工作簿路径（供跨表前缀加路径用）。</summary>
    internal sealed class RewriteContext
    {
        /// <summary>物理列差 = 目标列号 − 源列号（跨表/跨簿引用地址随单元格整体平移用）。</summary>
        internal int ColumnDelta;
        /// <summary>源工作簿完整路径（跨表引用前缀改写为带路径跨簿引用时用）。可为 null。</summary>
        internal string SourcePath;
    }

    internal static class FormulaRefRewriter
    {
        internal const int MaxColumn = 16384;
        internal const int MaxRow = 1048576;

        // ---- 掩码：正确跳过不该改写的区域 ----

        /// <summary>双引号字符串字面量（含两端引号）；"" 视为转义。</summary>
        internal static bool[] Literal(string s)
        {
            bool[] mask = new bool[s.Length];
            bool inside = false;
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '"')
                {
                    mask[i] = true;
                    if (inside && i + 1 < s.Length && s[i + 1] == '"') { i++; mask[i] = true; }
                    else inside = !inside;
                }
                else if (inside) mask[i] = true;
                i++;
            }
            return mask;
        }

        /// <summary>单引号区域（含两端引号）；'' 视为转义。</summary>
        internal static bool[] Quoted(string s)
        {
            bool[] mask = new bool[s.Length];
            bool inside = false;
            int i = 0;
            while (i < s.Length)
            {
                char c = s[i];
                if (c == '\'')
                {
                    mask[i] = true;
                    if (inside && i + 1 < s.Length && s[i + 1] == '\'') { i++; mask[i] = true; }
                    else inside = !inside;
                }
                else if (inside) mask[i] = true;
                i++;
            }
            return mask;
        }

        /// <summary>
        /// 「前缀!」掩码 —— 只掩码跨表/跨簿引用的前缀（含 !），不掩码 ! 后面的地址，
        /// 让地址部分可被 RxCellRange / RxWholeColumn 扫描后按「列不变、行偏移」重写。
        /// 覆盖 'My Sheet'!、Sheet2!、[Book1.xlsx]Sheet1! 三种前缀形态。
        /// </summary>
        internal static bool[] Qualified(string s, bool[] literal, bool[] quoted)
        {
            bool[] mask = new bool[s.Length];
            for (int p = 0; p < s.Length; p++)
            {
                if (s[p] != '!' || literal[p] || quoted[p]) continue;
                int left = p - 1;
                if (left >= 0 && quoted[left])
                {
                    // 引号表名：向左退到该引号段的起始引号。
                    while (left > 0 && quoted[left - 1]) left--;
                }
                else
                {
                    // 无引号表名 / [工作簿]工作表名 的工作表名部分。
                    while (left >= 0 && !literal[left] && !quoted[left] && IsSheetNameChar(s[left])) left--;
                    left++;
                }
                // 只掩码前缀 + !（不含 ! 后面的地址）。
                for (int k = Math.Max(0, left); k <= p; k++) mask[k] = true;
            }
            return mask;
        }

        private static bool IsSheetNameChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_' || c == '.' || c > '\x0080';
        }

        /// <summary>判定「!」前缀是否属于跨簿（前缀文本含 [ 即 [文件]Sheet 或引号内路径/[）。</summary>
        internal static bool PrefixHasBracket(string body, bool[] quoted, int bang)
        {
            int left = bang - 1;
            if (left >= 0 && quoted[left])
            {
                while (left >= 0 && quoted[left]) left--;   // 引号段起始引号
                left++;
            }
            else
            {
                while (left >= 0 && (IsSheetNameChar(body[left]) || body[left] == '[' || body[left] == ']' || body[left] == ':')) left--;
                left++;
            }
            for (int k = left; k < bang; k++)
                if (body[k] == '[') return true;
            return false;
        }

        /// <summary>定位「!」前缀的起止（含前导单引号），用于跨表前缀改写。返回起始下标；无前缀返回 -1。</summary>
        private static int PrefixStartOf(string body, bool[] quoted, int bang)
        {
            int left = bang - 1;
            if (left < 0) return -1;
            if (quoted[left])
            {
                while (left >= 0 && quoted[left]) left--;
                return left + 1;   // 起始引号
            }
            while (left >= 0 && (IsSheetNameChar(body[left]) || body[left] == '[' || body[left] == ']' || body[left] == ':')) left--;
            return left + 1;
        }

        // ---- 扫描：定位本地引用 ----

        // 单元格 / 区域：A1 / $A$1 / A1:B2。
        // 前后断言挡住 LOG10(、Total1、1.5E10、"A1"、[A1.xlsx]、Table1[Col] 等误判。
        // 注意用 [0-9] 而非 \d —— .NET 的 \d 会匹配全角等 Unicode 数字。
        // lookbehind 不再排除 !：让 Sheet1!A1 的 A1 也能被扫描到（前缀由 qualified 掩码保护）。
        private static readonly Regex RxCellRange = new Regex(
            @"(?<![A-Za-z0-9_.$\[\]'])"
            + @"(?<c1>\$?[A-Za-z]{1,3})(?<r1>\$?[0-9]{1,7})"
            + @"(?::(?<c2>\$?[A-Za-z]{1,3})(?<r2>\$?[0-9]{1,7}))?"
            + @"(?![A-Za-z0-9_(['])", RegexOptions.Compiled);

        // 整列：A:A / $A:$A（不搬就会静默指向错列，必须支持）。
        private static readonly Regex RxWholeColumn = new Regex(
            @"(?<![A-Za-z0-9_.$\[\]'])"
            + @"(?<c1>\$?[A-Za-z]{1,3}):(?<c2>\$?[A-Za-z]{1,3})"
            + @"(?![A-Za-z0-9_(['])", RegexOptions.Compiled);

        internal static List<RefSpan> Scan(string body)
        {
            List<RefSpan> spans = new List<RefSpan>();
            if (string.IsNullOrEmpty(body)) return spans;
            bool[] literal = Literal(body);
            bool[] quoted = Quoted(body);
            bool[] qualified = Qualified(body, literal, quoted);
            Collect(spans, body, RxCellRange, literal, quoted, qualified, false);
            Collect(spans, body, RxWholeColumn, literal, quoted, qualified, true);
            spans.Sort(delegate(RefSpan a, RefSpan b) { return a.Start.CompareTo(b.Start); });
            return spans;
        }

        private static void Collect(List<RefSpan> spans, string body, Regex pattern,
            bool[] literal, bool[] quoted, bool[] qualified, bool wholeColumn)
        {
            for (Match m = pattern.Match(body); m.Success; m = m.NextMatch())
            {
                bool blocked = false;
                for (int k = m.Index; k < m.Index + m.Length; k++)
                    if (literal[k] || quoted[k] || qualified[k]) { blocked = true; break; }
                if (blocked) continue;
                RefSpan span = Parse(m, wholeColumn);
                if (span == null) continue;
                // 地址起始前一个字符是 ! 且该 ! 属于前缀掩码 → 跨表/跨簿地址。
                bool external = m.Index > 0 && body[m.Index - 1] == '!' && qualified[m.Index - 1];
                if (external)
                {
                    span.Kind = PrefixHasBracket(body, quoted, m.Index - 1) ? RefKind.CrossBook : RefKind.CrossSheet;
                    span.PrefixStart = PrefixStartOf(body, quoted, m.Index - 1);
                    span.PrefixLength = (m.Index - 1) - span.PrefixStart;   // 前缀到 ! 之前
                }
                else span.Kind = RefKind.Local;
                foreach (RefSpan existing in spans)
                    if (span.Start < existing.Start + existing.Length
                        && existing.Start < span.Start + span.Length) { blocked = true; break; }
                if (!blocked) spans.Add(span);
            }
        }

        private static RefSpan Parse(Match m, bool wholeColumn)
        {
            string c1 = m.Groups["c1"].Value;
            int column1 = CellAddress.ColumnNumber(c1);
            if (column1 <= 0) return null;
            RefSpan span = new RefSpan();
            span.Start = m.Index;
            span.Length = m.Length;
            span.Text = m.Value;
            span.Column1 = column1;
            span.Column1Absolute = c1.StartsWith("$", StringComparison.Ordinal);
            span.IsWholeColumn = wholeColumn;
            if (!wholeColumn)
            {
                string r1 = m.Groups["r1"].Value;
                int row1;
                if (!int.TryParse(r1.TrimStart('$'), NumberStyles.None, CultureInfo.InvariantCulture, out row1)) return null;
                span.Row1 = row1;
                span.Row1Absolute = r1.StartsWith("$", StringComparison.Ordinal);
            }
            if (m.Groups["c2"].Success)
            {
                int column2 = CellAddress.ColumnNumber(m.Groups["c2"].Value);
                if (column2 <= 0) return null;
                span.IsRange = true;
                span.Column2 = column2;
                span.Column2Absolute = m.Groups["c2"].Value.StartsWith("$", StringComparison.Ordinal);
                if (!wholeColumn)
                {
                    string r2 = m.Groups["r2"].Value;
                    int row2;
                    if (!int.TryParse(r2.TrimStart('$'), NumberStyles.None, CultureInfo.InvariantCulture, out row2)) return null;
                    span.Row2 = row2;
                    span.Row2Absolute = r2.StartsWith("$", StringComparison.Ordinal);
                }
            }
            return span;
        }

        // ---- 重写 ----

        /// <summary>
        /// 重写公式中的同表引用。
        /// </summary>
        /// <param name="formula">源公式（可带前导 =）。</param>
        /// <param name="map">源列 → 目标列映射；null 表示只做行偏移。</param>
        /// <param name="rowDelta">该列的行偏移 = 目标起始行 − 源数据首行。</param>
        /// <param name="excludedNames">工作簿定义名；命中则不改写（避免误判定义名）。可为 null。</param>
        internal static RewriteResult Rewrite(string formula, IMergeColumnMapper map,
            int rowDelta, HashSet<string> excludedNames)
        {
            return Rewrite(formula, map, rowDelta, excludedNames, null);
        }

        /// <summary>
        /// 重写公式中的引用。
        /// </summary>
        /// <param name="formula">源公式（可带前导 =）。</param>
        /// <param name="map">源列 → 目标列映射；null 表示本表引用只做行偏移。</param>
        /// <param name="rowDelta">该列的行偏移 = 目标起始行 − 源数据首行。</param>
        /// <param name="excludedNames">工作簿定义名；命中则不改写。可为 null。</param>
        /// <param name="context">物理列差 + 源工作簿路径；null 时跨表/跨簿列不做偏移、前缀不加路径。</param>
        internal static RewriteResult Rewrite(string formula, IMergeColumnMapper map,
            int rowDelta, HashSet<string> excludedNames, RewriteContext context)
        {
            RewriteResult result = new RewriteResult();
            if (string.IsNullOrEmpty(formula)) { result.Text = formula ?? ""; return result; }
            string body = formula.Trim();
            bool leading = body.StartsWith("=", StringComparison.Ordinal);
            if (leading) body = body.Substring(1);

            List<RefSpan> spans = Scan(body);
            if (spans.Count == 0) { result.Text = formula; return result; }

            StringBuilder builder = new StringBuilder(body);
            // 从后往前替换，保证前面的索引不发生偏移。
            for (int i = spans.Count - 1; i >= 0; i--)
            {
                RefSpan span = spans[i];
                if (excludedNames != null && excludedNames.Contains(span.Text)) continue;
                string text = Render(span, map, rowDelta, context, result.Issues);
                if (string.Equals(text, span.Text, StringComparison.Ordinal)) continue;
                builder.Remove(span.Start, span.Length);
                builder.Insert(span.Start, text);
            }
            string rewritten = builder.ToString();
            result.Changed = !string.Equals(rewritten, body, StringComparison.Ordinal);
            result.Text = leading ? "=" + rewritten : rewritten;

            // 跨表前缀改写：把 Sheet2! / 'My Sheet'! 等前缀替换为带完整路径的跨簿引用前缀。
            if (context != null && !string.IsNullOrEmpty(context.SourcePath))
                result.Text = RewriteCrossSheetPrefixes(result.Text, context.SourcePath);

            return result;
        }

        private static string Render(RefSpan span, IMergeColumnMapper map, int rowDelta, RewriteContext context, List<RefIssue> issues)
        {
            string c1 = MapColumn(span.Column1, span.Column1Absolute, span.Kind, map, context, issues, span.Text);
            string c2 = span.IsRange
                ? MapColumn(span.Column2, span.Column2Absolute, span.Kind, map, context, issues, span.Text) : null;
            string r1 = null, r2 = null;
            if (!span.IsWholeColumn)
            {
                r1 = MapRow(span.Row1, span.Row1Absolute, rowDelta, issues, span.Text);
                r2 = span.IsRange
                    ? MapRow(span.Row2, span.Row2Absolute, rowDelta, issues, span.Text) : null;
            }
            return c1 + r1 + (span.IsRange ? ":" + c2 + r2 : "");
        }

        private static string MapColumn(int column, bool absolute, RefKind kind, IMergeColumnMapper map,
            RewriteContext context, List<RefIssue> issues, string fragment)
        {
            string letters = CellAddress.ColumnLetters(column);
            if (absolute) return "$" + letters;              // 绝对列：任何类别都原样不动

            if (kind == RefKind.Local)
            {
                // 本表引用：按源列→目标列映射表换算。
                if (map == null) return letters;
                int target = map.MapColumn(column);
                if (target <= 0)
                {
                    issues.Add(new RefIssue { Kind = RefIssueKind.UnmappedColumn, Column = column, Fragment = fragment });
                    return letters;
                }
                if (target > MaxColumn)
                {
                    issues.Add(new RefIssue { Kind = RefIssueKind.ColumnOutOfRange, Column = column, Fragment = fragment });
                    return letters;
                }
                return CellAddress.ColumnLetters(target);
            }

            // 跨表/跨簿引用：只遵循相对引用规律，随单元格整体平移（物理列差）。
            if (context == null) return letters;
            int shifted = column + context.ColumnDelta;
            if (shifted < 1 || shifted > MaxColumn)
            {
                issues.Add(new RefIssue { Kind = RefIssueKind.ColumnOutOfRange, Column = column, Fragment = fragment });
                return letters;
            }
            return CellAddress.ColumnLetters(shifted);
        }

        private static string MapRow(int row, bool absolute, int delta, List<RefIssue> issues, string fragment)
        {
            if (absolute) return "$" + row.ToString(CultureInfo.InvariantCulture);   // 绝对行：原样不动
            int target = row + delta;
            if (target < 1 || target > MaxRow)
            {
                issues.Add(new RefIssue { Kind = RefIssueKind.RowOutOfRange, Row = row, Fragment = fragment });
                return row.ToString(CultureInfo.InvariantCulture);
            }
            return target.ToString(CultureInfo.InvariantCulture);
        }

        // ---- 跨表前缀改写 ----

        /// <summary>
        /// 把「本工作簿跨表引用」的前缀（Sheet2!、'My Sheet'!、Sheet1:Sheet3! 等）改写为
        /// 带完整路径的跨簿引用前缀：'路径[文件名]前缀'!。跨簿引用（已含 [）不改。
        /// </summary>
        private static string RewriteCrossSheetPrefixes(string text, string sourcePath)
        {
            // 生成跨簿前缀包裹：'<目录>\[<文件名>]<原前缀>'
            string dir = null, file = null;
            try
            {
                dir = System.IO.Path.GetDirectoryName(sourcePath);
                file = System.IO.Path.GetFileName(sourcePath);
            }
            catch { }
            if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(file)) return text;

            // 扫描所有「!」，改写其后紧跟地址、且前缀不含 [ 的跨表引用前缀。
            string body = text;
            bool leading = body.StartsWith("=", StringComparison.Ordinal);
            if (leading) body = body.Substring(1);

            bool[] literal = Literal(body);
            bool[] quoted = Quoted(body);
            List<int> bangs = new List<int>();
            for (int p = 0; p < body.Length; p++)
                if (body[p] == '!' && !literal[p] && !quoted[p]) bangs.Add(p);

            StringBuilder builder = new StringBuilder(body);
            for (int i = bangs.Count - 1; i >= 0; i--)
            {
                int bang = bangs[i];
                // 跨簿前缀（含 [）不改。
                if (PrefixHasBracket(body, quoted, bang)) continue;
                int start = PrefixStartOf(body, quoted, bang);
                if (start < 0) continue;
                string original = body.Substring(start, bang - start);   // 原前缀（不含 !）
                // 去掉可能已有的外层单引号（PrefixStartOf 已含起始引号，original 里若以 ' 开头则需剥）。
                string core = original;
                if (core.Length >= 2 && core[0] == '\'' && core[core.Length - 1] == '\'')
                    core = core.Substring(1, core.Length - 2);
                string wrapped = "'" + dir + "\\[" + file + "]" + core + "'";
                builder.Remove(start, bang - start);
                builder.Insert(start, wrapped);
            }
            string result = builder.ToString();
            return leading ? "=" + result : result;
        }
    }
}
