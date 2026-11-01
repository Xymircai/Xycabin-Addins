using System;
using System.Collections.Generic;
using System.Text;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // =====================================================================
    // SingleSummaryProbe — 单体汇总表生成离线自测（不连 WPS，不需要 COM）
    //
    // 覆盖「换链」这条核心规则的边界：
    //   1. 只替换方括号里的模板公司文件名，路径 / ]BS / '!$AB$9 不动
    //   2. 公式里出现的其他公司文件名保持不动
    //   3. 不含 [xxx.xlsx] 的本簿引用一律不碰
    //   4. 双引号字符串常量里的方括号不改
    //   5. 含「（合并）」的公司走同一规则，大小写不敏感
    //   6. 列字母、表头归一化、N=0 / 全称重复 / 列越界的阻断
    //
    // 编译：project.json tools 段（/r:Excel_Iterate.dll）
    // 运行：SingleSummaryProbe.exe  退出码 0=全 PASS，1=有 FAIL
    // =====================================================================

    internal static class SingleSummaryProbe
    {
        static int _pass = 0, _fail = 0;

        static void Assert(string name, bool cond, string detail)
        {
            if (cond) { _pass++; Console.WriteLine("  [PASS] " + name); }
            else { _fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
        }

        static void AssertRewrite(string formula, string source, string target, string expected, int expectedCount)
        {
            SummaryRewrite rewrite = SummaryLinkRewriter.Replace(formula, source, target);
            Assert("换链 " + formula,
                string.Equals(rewrite.Text, expected, StringComparison.Ordinal) && rewrite.Count == expectedCount,
                "得到 " + rewrite.Text + "（" + rewrite.Count + " 处），应为 " + expected + "（" + expectedCount + " 处）");
        }

        static SummaryCompany Company(int row, string code, string name, bool selected)
        {
            SummaryCompany company = new SummaryCompany();
            company.Row = row; company.Code = code; company.Name = name; company.Selected = selected;
            return company;
        }

        static SummaryRequest Request(params SummaryCompany[] companies)
        {
            SummaryRequest request = new SummaryRequest();
            request.LeadSheet = "lead";
            request.TargetSheet = "汇总";
            request.HeaderRow = 1;
            request.TemplateColumn = 4;
            request.SourceFile = "甲公司.xlsx";
            request.Companies.AddRange(companies);
            return request;
        }

        static int Main()
        {
            Console.WriteLine("=== 单体汇总表生成：换链与规划离线自测 ===");
            Console.WriteLine();
            Console.WriteLine("--- 1. 精确换链：只动方括号里的文件名 ---");
            AssertRewrite(
                "='C:\\报表\\2026\\[甲公司.xlsx]BS'!$AB$9",
                "甲公司.xlsx", "乙公司.xlsx",
                "='C:\\报表\\2026\\[乙公司.xlsx]BS'!$AB$9", 1);
            AssertRewrite(
                "=[甲公司.xlsx]BS!$AB$9+[丙公司.xlsx]BS!$B$1",
                "甲公司.xlsx", "乙公司.xlsx",
                "=[乙公司.xlsx]BS!$AB$9+[丙公司.xlsx]BS!$B$1", 1);
            AssertRewrite(
                "=BS!L14-BS!G14",
                "甲公司.xlsx", "乙公司.xlsx",
                "=BS!L14-BS!G14", 0);
            AssertRewrite(
                "=\"见[甲公司.xlsx]\"&A1",
                "甲公司.xlsx", "乙公司.xlsx",
                "=\"见[甲公司.xlsx]\"&A1", 0);
            AssertRewrite(
                "=[甲公司（合并）.xlsx]BS!A1",
                "甲公司（合并）.xlsx", "乙公司（合并）.xlsx",
                "=[乙公司（合并）.xlsx]BS!A1", 1);
            AssertRewrite(
                "=[abc.xlsx]BS!A1",
                "ABC.XLSX", "乙公司.xlsx",
                "=[乙公司.xlsx]BS!A1", 1);
            AssertRewrite(
                "=[甲公司.xlsx]BS!A1",
                "甲公司.xlsx", "甲公司.xlsx",
                "=[甲公司.xlsx]BS!A1", 0);
            AssertRewrite(
                "='[甲公司.xlsx]BS'!A1+'[丁公司.xlsx]BS'!A1",
                "甲公司.xlsx", "乙公司.xlsx",
                "='[乙公司.xlsx]BS'!A1+'[丁公司.xlsx]BS'!A1", 1);

            Console.WriteLine();
            Console.WriteLine("--- 1b. WPS 括号引号兼容（0x8FE3000B 实测坑） ---");
            string quotedPath = SummaryLinkRewriter.QuoteLinks("='C:\\data\\[甲公司（合并）.xlsx]BS'!$AB$9", "甲公司（合并）.xlsx");
            Assert("已加引号的引用保持原样",
                quotedPath == "='C:\\data\\[甲公司（合并）.xlsx]BS'!$AB$9", quotedPath);
            string quotedBare = SummaryLinkRewriter.QuoteLinks("=[甲公司（合并）.xlsx]BS!$A$1", "甲公司（合并）.xlsx");
            Assert("裸引用补引号（含工作表名）",
                quotedBare == "='[甲公司（合并）.xlsx]BS'!$A$1", quotedBare);
            string quotedWithPath = SummaryLinkRewriter.QuoteLinks("=C:\\data\\[甲公司（合并）.xlsx]BS!$A$1", "甲公司（合并）.xlsx");
            Assert("带路径时引号覆盖整段",
                quotedWithPath == "='C:\\data\\[甲公司（合并）.xlsx]BS'!$A$1", quotedWithPath);
            string untouched = SummaryLinkRewriter.QuoteLinks("=[乙公司.xlsx]BS!$A$1", "乙公司.xlsx");
            Assert("文件名无需引号时不加引号", untouched == "=[乙公司.xlsx]BS!$A$1", untouched);
            string other = SummaryLinkRewriter.QuoteLinks("=[甲公司（合并）.xlsx]BS!$A$1+[丁公司.xlsx]BS!$A$1", "甲公司（合并）.xlsx");
            Assert("只给目标文件加引号", other == "='[甲公司（合并）.xlsx]BS'!$A$1+[丁公司.xlsx]BS!$A$1", other);
            Assert("字符串常量里的括号链接不补引号",
                SummaryLinkRewriter.QuoteLinks("=\"见[甲公司（合并）.xlsx]\"", "甲公司（合并）.xlsx") == "=\"见[甲公司（合并）.xlsx]\"",
                SummaryLinkRewriter.QuoteLinks("=\"见[甲公司（合并）.xlsx]\"", "甲公司（合并）.xlsx"));
            Assert("NeedsQuote 判定括号", SummaryLinkRewriter.NeedsQuote("甲公司（合并）.xlsx") && !SummaryLinkRewriter.NeedsQuote("甲公司.xlsx"), "判定错");

            Console.WriteLine();
            Console.WriteLine("--- 1c. 占位名可配（默认「大合并」，可改可空） ---");
            Assert("默认占位名是大合并", new SummaryRequest().Placeholder == "大合并", new SummaryRequest().Placeholder);
            Assert("占位一致", SummaryHeaders.MatchesPlaceholder("大合并", "大合并"), "未匹配");
            Assert("占位允许空格差异", SummaryHeaders.MatchesPlaceholder(" 大 合并 ", "大合并"), "未匹配");
            Assert("占位可换成公司名", SummaryHeaders.MatchesPlaceholder("甲公司（合并）", "甲公司（合并）")
                && !SummaryHeaders.MatchesPlaceholder("甲公司（合并）", "大合并"), "判定错");
            Assert("占位名留空=不校验", SummaryHeaders.MatchesPlaceholder("任意表头", ""), "留空仍在校验");
            Assert("空表头对非空占位不通过", !SummaryHeaders.MatchesPlaceholder("", "大合并"), "误通过");

            Console.WriteLine();
            Console.WriteLine("--- 2. 外链文件名提取 ---");
            List<string> tokens = SummaryLinkRewriter.Tokens("=[甲公司.xlsx]BS!A1+[丙公司.xlsx]BS!B2", null);
            Assert("提取两个外链文件名", tokens.Count == 2 && tokens[0] == "甲公司.xlsx" && tokens[1] == "丙公司.xlsx",
                string.Join("、", tokens.ToArray()));
            tokens = SummaryLinkRewriter.Tokens("=BS!L14-BS!G14", null);
            Assert("本簿引用不产生外链文件名", tokens.Count == 0, tokens.Count.ToString());

            Console.WriteLine();
            Console.WriteLine("--- 3. 列字母与表头归一化 ---");
            Assert("Letter(4)=D", SummaryColumns.Letter(4) == "D", SummaryColumns.Letter(4));
            Assert("Letter(27)=AA", SummaryColumns.Letter(27) == "AA", SummaryColumns.Letter(27));
            int column;
            Assert("解析 d=4", SummaryColumns.TryParse(" d ", out column) && column == 4, column.ToString());
            Assert("解析 AA=27", SummaryColumns.TryParse("AA", out column) && column == 27, column.ToString());
            Assert("解析 1 被拒", !SummaryColumns.TryParse("1", out column), column.ToString());
            Assert("表头 & 归一化", SummaryHeaders.IsCode(" 编码 & 简称 ") && SummaryHeaders.IsCode("编码＆简称"), "未识别");
            Assert("表头 全称 归一化", SummaryHeaders.IsName("全 称"), "未识别");
            Assert("占位 大合并 识别", SummaryHeaders.IsPlaceholder("大合并 "), "未识别");

            Console.WriteLine();
            Console.WriteLine("--- 4. 请求校验（N=0 / 重复 / 越界） ---");
            SummaryRequest empty = Request(Company(2, "A", "甲公司", false));
            Assert("N=0 阻断", empty.Validate() != null && empty.Validate().Contains("没有勾选"), empty.Validate());
            SummaryRequest duplicate = Request(Company(2, "A", "甲公司", true), Company(3, "B", "甲 公司", true));
            Assert("全称重复放行（不同名）", duplicate.Validate() == null, duplicate.Validate());
            SummaryRequest duplicated = Request(Company(2, "A", "甲公司", true), Company(3, "B", "甲公司", true));
            Assert("全称重复阻断", duplicated.Validate() != null && duplicated.Validate().Contains("重复"), duplicated.Validate());
            SummaryRequest overflow = Request(Company(2, "A", "甲", true), Company(3, "B", "乙", true));
            overflow.TemplateColumn = 16384;
            Assert("列越界阻断", overflow.Validate() != null && overflow.Validate().Contains("上限"), overflow.Validate());
            SummaryRequest edge = Request(Company(2, "A", "甲", true));
            edge.TemplateColumn = 16384;
            Assert("恰好在最后一列不误伤", edge.Validate() == null, edge.Validate());
            SummaryRequest ok = Request(Company(2, "A", "甲公司", true), Company(3, "B", "乙公司", true));
            Assert("正常请求通过", ok.Validate() == null, ok.Validate());
            Assert("Chosen 保持行序", ok.Chosen().Count == 2 && ok.Chosen()[0].Row == 2 && ok.Chosen()[1].Row == 3, "顺序错");
            Assert("目标文件名 = 全称.xlsx", ok.Chosen()[1].FileName == "乙公司.xlsx", ok.Chosen()[1].FileName);

            Console.WriteLine();
            Console.WriteLine("--- 5. 弹窗可离线构造（尺寸登记 summary.single） ---");
            try
            {
                using (SingleSummaryForm form = new SingleSummaryForm())
                {
                    form.SetContext("C:\\样例.xlsx", new List<string> { "lead", "汇总" }, "lead", "汇总", "D");
                    SummaryRequest built = form.BuildRequest();
                    Assert("弹窗默认 Lead=lead", built.LeadSheet == "lead", built.LeadSheet);
                    Assert("弹窗默认模板列 D", built.TemplateColumn == 4, built.TemplateColumn.ToString());
                    Assert("弹窗默认表头行 1", built.HeaderRow == 1, built.HeaderRow.ToString());
                    form.SetCompanies(new List<SummaryCompany> { Company(2, "A", "甲公司", true) });
                    Assert("名单默认勾选", form.BuildRequest().Chosen().Count == 1, "未勾选");
                }
            }
            catch (Exception ex)
            {
                Assert("弹窗可构造", false, ex.Message);
            }

            Console.WriteLine();
            Console.WriteLine("PASS " + _pass + "，FAIL " + _fail);
            return _fail == 0 ? 0 : 1;
        }
    }
}
