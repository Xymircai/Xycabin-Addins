using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel_Iterate
{
    internal sealed class SummaryWorkbookChoice
    {
        internal string Path, Name;
        public override string ToString() { return Name + "  [" + Path + "]"; }
    }
    // =====================================================================
    // 单体汇总表生成 —— 纯逻辑层（不碰 COM，可离线核对）
    //
    // 职责：
    //   · 列字母 ↔ 列号（模板列定位）
    //   · 表头名归一化（"编码 & 简称" / 全角空格 等写法等价）
    //   · 外链文件名提取与换链（本功能最容易出错的一段，单独成类）
    //   · 请求校验（N=0 阻断、列越界、全称重复）
    // =====================================================================

    /// <summary>Lead 表里的一家公司 = 汇总表里的一列。</summary>
    internal sealed class SummaryCompany
    {
        internal int Row;                    // Lead 表中的行号（回写「选用」用）
        internal string Code = "";           // 编码&简称 → 写进该列表头
        internal string Name = "";           // 全称 → 换链目标文件名主体
        internal bool Selected = true;       // 确认窗勾选状态（默认全勾）
        internal string PriorSelection = ""; // 「选用」列原值，仅作展示

        /// <summary>换链目标文件名：全称 + ".xlsx"。含「（合并）」的公司同样只有这一条规则。</summary>
        internal string FileName { get { return Name + ".xlsx"; } }
    }

    /// <summary>一次生成的完整请求（全部来自弹窗，不含 COM）。</summary>
    internal sealed class SummaryRequest
    {
        internal string WorkbookName = "";
        internal string LeadSheet = "";
        internal string TargetSheet = "";
        internal int HeaderRow = 1;        // Lead 表里「编码&简称/全称」所在行
        internal int SheetHeaderRow = 1;   // 汇总表里公司列标题（占位）所在行，可与 Lead 不同
        internal int TemplateColumn;         // 模板列列号（1 起，D=4）
        internal string SourceFile = "";     // 模板公司文件名，如 "示例有限公司.xlsx"
        internal string Placeholder = SummaryHeaders.Placeholder;  // 模板列的表头占位名，默认「大合并」，可改；留空=不校验
        internal readonly List<SummaryCompany> Companies = new List<SummaryCompany>();

        internal string TemplateLetter { get { return SummaryColumns.Letter(TemplateColumn); } }

        /// <summary>本次真正参与生成的公司（勾选且两列都非空），保持 Lead 行序。</summary>
        internal List<SummaryCompany> Chosen()
        {
            List<SummaryCompany> list = new List<SummaryCompany>();
            foreach (SummaryCompany company in Companies)
                if (company.Selected && company.Code.Trim().Length > 0 && company.Name.Trim().Length > 0) list.Add(company);
            return list;
        }

        /// <summary>返回 null 表示通过；否则是给用户看的阻断原因。</summary>
        internal string Validate()
        {
            if (LeadSheet.Trim().Length == 0) return "请先选择 Lead 工作表。";
            if (TargetSheet.Trim().Length == 0) return "找不到汇总工作表，请重新打开窗口。";
            if (string.Equals(LeadSheet.Trim(), TargetSheet.Trim(), StringComparison.OrdinalIgnoreCase))
                return "Lead 表与汇总表不能是同一张工作表。";
            if (HeaderRow < 1) return "Lead 表头行号无效。";
            if (SheetHeaderRow < 1) return "汇总表表头行号无效。";
            if (TemplateColumn < 1 || TemplateColumn > 16384) return "模板列无效，请填列字母（如 D）。";
            if (SourceFile.Trim().Length == 0) return "请填写模板公司全称（换链源文件名，如 示例有限公司.xlsx）。";
            if (!SummaryLinkRewriter.IsLinkFile(SourceFile)) return "模板公司全称需拼成「全称.xlsx」的样式。";
            List<SummaryCompany> chosen = Chosen();
            if (chosen.Count == 0) return "没有勾选任何公司（N=0），已阻断，未改动工作表。";
            if (TemplateColumn + chosen.Count - 1 > 16384)
                return "勾选 " + chosen.Count.ToString(CultureInfo.InvariantCulture)
                    + " 家公司，从 " + TemplateLetter + " 起向右会超出工作表列上限。";
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SummaryCompany company in chosen)
                if (!names.Add(company.Name.Trim()))
                    return "勾选名单中「全称」重复：" + company.Name.Trim()
                        + "。重复会生成两列同名链接，请先回 Lead 表修掉。";
            return null;
        }
    }

    /// <summary>列字母 ↔ 列号。</summary>
    internal static class SummaryColumns
    {
        internal static string Letter(int column)
        {
            if (column < 1 || column > 16384) throw new ArgumentOutOfRangeException("column");
            StringBuilder text = new StringBuilder(3);
            int value = column;
            while (value > 0)
            {
                int digit = (value - 1) % 26;
                text.Insert(0, (char)('A' + digit));
                value = (value - 1) / 26;
            }
            return text.ToString();
        }

        internal static bool TryParse(string text, out int column)
        {
            column = 0;
            if (text == null) return false;
            string value = text.Trim().ToUpperInvariant().Replace("$", "");
            if (value.Length == 0 || value.Length > 3) return false;
            int result = 0;
            foreach (char c in value)
            {
                if (c < 'A' || c > 'Z') return false;
                result = result * 26 + (c - 'A' + 1);
            }
            if (result < 1 || result > 16384) return false;
            column = result;
            return true;
        }
    }

    /// <summary>表头名归属判断：去空白、全角&amp;归一、忽略大小写。</summary>
    internal static class SummaryHeaders
    {
        internal const string CodeHeader = "编码&简称";
        internal const string NameHeader = "全称";
        internal const string PickHeader = "选用";
        internal const string Placeholder = "大合并";

        internal static string Normalize(string text)
        {
            if (text == null) return "";
            StringBuilder builder = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (c == '\u3000' || c == '\u00A0' || char.IsWhiteSpace(c)) continue;
                builder.Append(c == '＆' ? '&' : c);
            }
            return builder.ToString().ToUpperInvariant();
        }

        internal static bool IsCode(string text) { return Normalize(text) == Normalize(CodeHeader); }
        internal static bool IsName(string text) { return Normalize(text) == Normalize(NameHeader); }
        internal static bool IsPick(string text) { return Normalize(text) == Normalize(PickHeader); }
        internal static bool IsPlaceholder(string text) { return Normalize(text) == Normalize(Placeholder); }

        /// <summary>
        /// 模板列表头是否与占位名一致。占位名可配：留空表示不校验（直接按列写表头），
        /// 非空时按「去空白 + 全角＆归一」后忽略大小写比较。
        /// </summary>
        internal static bool MatchesPlaceholder(string actual, string expected)
        {
            string want = expected == null ? "" : expected.Trim();
            if (want.Length == 0) return true;
            return string.Equals(Normalize(actual), Normalize(want), StringComparison.OrdinalIgnoreCase);
        }
    }

    internal sealed class SummaryRewrite
    {
        internal string Text = "";
        internal int Count;
    }

    // =====================================================================
    // 换链（核心）
    //
    // 规则（与需求逐条对应）：
    //   1. 只处理公式串；不含 [xxx.xlsx] 的本簿引用（=BS!L14-BS!G14）一律不碰。
    //   2. 只把方括号里的 "[模板公司全称.xlsx]" 换成 "[当前列公司全称.xlsx]"；
    //      方括号外的路径（C:\目录\）、"]BS"、"'!$AB$9" 一律原样保留。
    //   3. 公式中出现的其他公司文件名（不是模板公司）保持不动 —— 因为只与
    //      模板公司文件名做整名比较，不做包含/前缀匹配。
    //   4. 双引号字符串常量里的方括号不算链接（如 "见[附件.xlsx]"），不改。
    //   5. 含「（合并）」的公司名没有例外，走同一比较逻辑。
    // =====================================================================
    internal static class SummaryLinkRewriter
    {
        internal static SummaryRewrite Replace(string formula, string sourceFile, string targetFile)
        {
            SummaryRewrite result = new SummaryRewrite();
            result.Text = formula == null ? "" : formula;
            string source = sourceFile == null ? "" : sourceFile.Trim();
            string target = targetFile == null ? "" : targetFile.Trim();
            if (formula == null || source.Length == 0 || target.Length == 0) return result;
            if (string.Equals(source, target, StringComparison.OrdinalIgnoreCase)) return result;

            StringBuilder text = new StringBuilder(formula.Length + 32);
            bool quoted = false;
            int i = 0;
            while (i < formula.Length)
            {
                char c = formula[i];
                if (c == '"')
                {
                    // 双引号里的 "" 是转义的单引号，状态不变
                    if (quoted && i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        text.Append("\"\"");
                        i += 2;
                        continue;
                    }
                    quoted = !quoted;
                    text.Append(c);
                    i++;
                    continue;
                }
                if (!quoted && c == '[')
                {
                    int close = formula.IndexOf(']', i + 1);
                    if (close > i)
                    {
                        string inner = formula.Substring(i + 1, close - i - 1);
                        if (IsLinkFile(inner) && string.Equals(inner.Trim(), source, StringComparison.OrdinalIgnoreCase))
                        {
                            text.Append('[').Append(target).Append(']');
                            result.Count++;
                            i = close + 1;
                            continue;
                        }
                    }
                }
                text.Append(c);
                i++;
            }
            result.Text = text.ToString();
            return result;
        }

        /// <summary>方括号内容是否像一个工作簿文件名（只认 .xlsx，与本功能口径一致）。</summary>
        internal static bool IsLinkFile(string inner)
        {
            string name = inner == null ? "" : inner.Trim();
            if (name.Length < 6) return false;
            return name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>取公式里出现的全部外链文件名（去重、保持出现顺序），供预检提示用。</summary>
        internal static List<string> Tokens(string formula, List<string> into)
        {
            if (into == null) into = new List<string>();
            if (string.IsNullOrEmpty(formula)) return into;
            bool quoted = false;
            int i = 0;
            while (i < formula.Length)
            {
                char c = formula[i];
                if (c == '"')
                {
                    if (quoted && i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; }
                    quoted = !quoted;
                    i++;
                    continue;
                }
                if (!quoted && c == '[')
                {
                    int close = formula.IndexOf(']', i + 1);
                    if (close > i)
                    {
                        string inner = formula.Substring(i + 1, close - i - 1).Trim();
                        if (IsLinkFile(inner) && !into.Exists(delegate(string item) { return string.Equals(item, inner, StringComparison.OrdinalIgnoreCase); }))
                            into.Add(inner);
                        i = close + 1;
                        continue;
                    }
                }
                i++;
            }
            return into;
        }

        // -----------------------------------------------------------------
        // WPS 实测坑（2026-10-31，11.8.2 / 12.0 内核同现象）：
        //   给 Range.Formula 赋值时，只要外链文件名含括号（半角/全角），
        //   且引用**未加单引号**，WPS 直接抛 0x8FE3000B；
        //   ='C:\data\[甲公司（合并）.xlsx]BS'!$C$3 这种带引号的写法可以写入。
        //   所以：仅当目标文件名需要引号、且该引用原本没引号时，才补一对单引号，
        //   引号范围覆盖「路径 + [文件名] + 工作表名」，`!` 与地址部分保持不动。
        // -----------------------------------------------------------------

        /// <summary>文件名里出现这些字符时，Excel/WPS 要求整段引用加单引号。</summary>
        internal static bool NeedsQuote(string fileName)
        {
            string name = fileName == null ? "" : fileName;
            char[] specials = { ' ', '(', ')', '（', '）', '#', '&', '%', '~', '\'', ',', ';', '@', '+', '!', '^', '=', '<', '>' };
            foreach (char c in specials)
                if (name.IndexOf(c) >= 0) return true;
            return false;
        }

        /// <summary>
        /// 把公式里「未加引号、且文件名需要引号」的外链引用补上单引号。
        /// 其他引用、字符串常量、地址部分一律不动；已经加引号的引用保持原样。
        /// </summary>
        internal static string QuoteLinks(string formula) { return QuoteLinks(formula, null); }

        internal static string QuoteLinks(string formula, string onlyFileName)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            List<int[]> spans = new List<int[]>();
            bool quoted = false;
            int i = 0;
            while (i < formula.Length)
            {
                char c = formula[i];
                if (c == '"')
                {
                    if (quoted && i + 1 < formula.Length && formula[i + 1] == '"') { i += 2; continue; }
                    quoted = !quoted;
                    i++;
                    continue;
                }
                if (!quoted && c == '[')
                {
                    int close = formula.IndexOf(']', i + 1);
                    string inner = close > i ? formula.Substring(i + 1, close - i - 1) : "";
                    bool target = (onlyFileName == null || string.Equals(inner.Trim(), onlyFileName.Trim(), StringComparison.OrdinalIgnoreCase)) && NeedsQuote(inner);
                    if (close > i && target && IsLinkFile(inner))
                    {
                        int bang = formula.IndexOf('!', close + 1);
                        if (bang > close)
                        {
                            int start = i;
                            while (start > 0 && !IsReferenceDelimiter(formula[start - 1])) start--;
                            bool alreadyQuoted = start > 0 && formula[start - 1] == '\'' && formula[bang - 1] == '\'';
                            if (!alreadyQuoted) spans.Add(new int[] { start, bang });
                            i = bang + 1;
                            continue;
                        }
                    }
                }
                i++;
            }
            if (spans.Count == 0) return formula;
            StringBuilder text = new StringBuilder(formula);
            for (int s = spans.Count - 1; s >= 0; s--)
                text.Insert(spans[s][1], '\'').Insert(spans[s][0], '\'');
            return text.ToString();
        }

        /// <summary>引用令牌的边界字符：遇到它们说明到了下一个表达式片段。</summary>
        internal static bool IsReferenceDelimiter(char c)
        {
            // 不含 : 与反斜杠：Windows 路径（C:\data\...）里的它们不能当引用边界
            const string delimiters = "=+-*/()<>&^%,;{}[]! \t\r\n~\"'";
            return delimiters.IndexOf(c) >= 0;
        }
    }

    /// <summary>一处换链样例（界面里给人看，避免"以为换了"）。</summary>
    internal sealed class SummaryLinkSample
    {
        internal string Column = "";
        internal string Cell = "";
        internal string Before = "";
        internal string After = "";
    }
}
