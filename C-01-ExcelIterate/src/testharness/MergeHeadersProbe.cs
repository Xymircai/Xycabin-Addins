using System;
using System.Collections.Generic;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // =====================================================================
    // MergeHeadersProbe — 跨工作簿表头匹配合并离线自测（不连 WPS，不需要 COM）
    //
    // 覆盖三块纯逻辑：
    //   1. FormulaRefRewriter：公式引用的扫描与重写（含应跳过的四类误判）
    //   2. MergeMatcher：表头归一化匹配、缺列收集
    //   3. MergePlanner：写入分块与总量保护
    //
    // 编译：build.py 的 tools 段（/r:Excel_Iterate.dll）
    // 运行：build/MergeHeadersProbe.exe  退出码 0=全 PASS，1=有 FAIL
    // =====================================================================

    internal static class MergeHeadersProbe
    {
        static int _pass = 0, _fail = 0;

        static void Assert(string name, bool cond, string detail)
        {
            if (cond) { _pass++; Console.WriteLine("  [PASS] " + name); }
            else { _fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
        }

        static MergeColumnMapper Mapper(params int[] pairs)
        {
            MergeColumnMapper mapper = new MergeColumnMapper();
            for (int i = 0; i + 1 < pairs.Length; i += 2) mapper.Add(pairs[i], pairs[i + 1]);
            return mapper;
        }

        static void AssertRewrite(string formula, MergeColumnMapper mapper, int delta, string expected)
        {
            RewriteResult result = FormulaRefRewriter.Rewrite(formula, mapper, delta, null);
            Assert("Rewrite " + formula + " → " + expected,
                string.Equals(result.Text, expected, StringComparison.Ordinal),
                "实际 " + result.Text);
        }

        static void AssertUnchanged(string formula, MergeColumnMapper mapper, int delta)
        {
            RewriteResult result = FormulaRefRewriter.Rewrite(formula, mapper, delta, null);
            Assert("保持不变 " + formula,
                !result.Changed && string.Equals(result.Text, formula, StringComparison.Ordinal),
                "实际 " + result.Text);
        }

        static void AssertRewriteCtx(string formula, MergeColumnMapper mapper, int delta, int columnDelta, string sourcePath, string expected)
        {
            RewriteContext ctx = new RewriteContext { ColumnDelta = columnDelta, SourcePath = sourcePath };
            RewriteResult result = FormulaRefRewriter.Rewrite(formula, mapper, delta, null, ctx);
            Assert("RewriteCtx " + formula + " → " + expected,
                string.Equals(result.Text, expected, StringComparison.Ordinal),
                "实际 " + result.Text);
        }

        static MergeHeaderSet Set(int headerRow, params string[] headers)
        {
            MergeHeaderSet set = new MergeHeaderSet();
            set.HeaderRow = headerRow;
            set.LastUsedRow = headerRow + 2;
            set.FirstColumn = 1;
            set.LastColumn = headers.Length;
            for (int i = 0; i < headers.Length; i++)
            {
                MergeColumn column = new MergeColumn();
                column.Number = i + 1;
                column.Letter = CellAddress.ColumnLetters(i + 1);
                column.Text = headers[i];
                column.Key = MergeKey.Normalize(headers[i]);
                set.Columns.Add(column);
            }
            set.Seal();
            return set;
        }

        static void Main()
        {
            Console.WriteLine("=== 1. 列号与列字母 ===");
            Assert("ColumnLetters(1)=A", CellAddress.ColumnLetters(1) == "A", CellAddress.ColumnLetters(1));
            Assert("ColumnLetters(4)=D", CellAddress.ColumnLetters(4) == "D", CellAddress.ColumnLetters(4));
            Assert("ColumnLetters(26)=Z", CellAddress.ColumnLetters(26) == "Z", CellAddress.ColumnLetters(26));
            Assert("ColumnLetters(27)=AA", CellAddress.ColumnLetters(27) == "AA", CellAddress.ColumnLetters(27));
            Assert("ColumnLetters(16384)=XFD", CellAddress.ColumnLetters(16384) == "XFD", CellAddress.ColumnLetters(16384));
            Assert("ColumnNumber(A)=1", CellAddress.ColumnNumber("A") == 1, "x");
            Assert("ColumnNumber($D)=4", CellAddress.ColumnNumber("$D") == 4, "x");
            Assert("ColumnNumber(XFD)=16384", CellAddress.ColumnNumber("XFD") == 16384, "x");
            Assert("往返 702", CellAddress.ColumnNumber(CellAddress.ColumnLetters(702)) == 702, "x");

            Console.WriteLine("=== 2. 相对/绝对四种组合 ===");
            // 用户场景：源第 1 列（表头 D）→ 目标第 4 列，delta=+2
            MergeColumnMapper d4 = Mapper(1, 4);
            AssertRewrite("=A1+A2", d4, 2, "=D3+D4");
            AssertRewrite("=$A$1", d4, 2, "=$A$1");
            AssertRewrite("=$A1", d4, 2, "=$A3");
            AssertRewrite("=A$1", d4, 2, "=D$1");
            AssertRewrite("=A1", d4, 0, "=D1");
            AssertRewrite("=A5", d4, -2, "=D3");

            Console.WriteLine("=== 2b. 乱序/交叉映射（用户场景：源 C→目标 B、源 D→目标 A）===");
            MergeColumnMapper crossed = Mapper(3, 2, 4, 1);
            AssertRewrite("=C2+D2", crossed, 0, "=B2+A2");
            AssertRewrite("=C2+D2", crossed, 5, "=B7+A7");   // 行偏移 +5
            AssertRewrite("=C3+D3", crossed, 5, "=B8+A8");
            AssertRewrite("=SUM(C2:D4)", crossed, 3, "=SUM(B5:A7)");

            Console.WriteLine("=== 2c. ReadColumn 一维/二维数组取值（At 兼容）===");
            // WPS 单列 Range 返回一维数组；Excel 返回二维。At 必须两者都取对。
            object[] oneDim = new object[] { "=C2+D2", "=C3+D3", "=C4+D4" };
            Assert("一维数组 At[0]", string.Equals(MergeSheetReader.At(oneDim, 0, 3, 0) as string, "=C2+D2", StringComparison.Ordinal), Convert.ToString(MergeSheetReader.At(oneDim, 0, 3, 0)));
            Assert("一维数组 At[2]", string.Equals(MergeSheetReader.At(oneDim, 2, 3, 0) as string, "=C4+D4", StringComparison.Ordinal), Convert.ToString(MergeSheetReader.At(oneDim, 2, 3, 0)));
            Assert("一维数组越界返回 null", MergeSheetReader.At(oneDim, 3, 3, 0) == null, "非 null");
            object[,] twoDim = new object[3, 1] { { "v1" }, { "v2" }, { "v3" } };
            Assert("二维数组 At[1,0]", string.Equals(MergeSheetReader.At(twoDim, 1, 3, 0) as string, "v2", StringComparison.Ordinal), Convert.ToString(MergeSheetReader.At(twoDim, 1, 3, 0)));
            Assert("二维数组越界返回 null", MergeSheetReader.At(twoDim, 3, 3, 0) == null, "非 null");
            // 单值（height==1 时非数组）原样返回
            Assert("单值原样返回", string.Equals(MergeSheetReader.At("=A1", 0, 1, 0) as string, "=A1", StringComparison.Ordinal), "x");

            Console.WriteLine("=== 2d. MergeSnapshot.Pick 一维/二维数组取值（快照与读回核对兼容）===");
            // 单列写入块读回 Value2/Formula 是一维数组，Pick 必须取对，否则报「数值写入核对失败」。
            object[] pickOne = new object[] { "b1", "b2", "b3" };
            Assert("Pick 一维[0]", string.Equals(MergeSnapshot.Pick(pickOne, null, 0, 0, 3, 1) as string, "b1", StringComparison.Ordinal), Convert.ToString(MergeSnapshot.Pick(pickOne, null, 0, 0, 3, 1)));
            Assert("Pick 一维[2]", string.Equals(MergeSnapshot.Pick(pickOne, null, 2, 0, 3, 1) as string, "b3", StringComparison.Ordinal), Convert.ToString(MergeSnapshot.Pick(pickOne, null, 2, 0, 3, 1)));
            Assert("Pick 一维越界 null", MergeSnapshot.Pick(pickOne, null, 3, 0, 3, 1) == null, "非 null");
            object[,] pickTwo = new object[3, 2] { { "a", "b" }, { "c", "d" }, { "e", "f" } };
            Assert("Pick 二维[1,0]", string.Equals(MergeSnapshot.Pick(pickTwo, null, 1, 0, 3, 2) as string, "c", StringComparison.Ordinal), Convert.ToString(MergeSnapshot.Pick(pickTwo, null, 1, 0, 3, 2)));
            Assert("Pick 二维[2,1]", string.Equals(MergeSnapshot.Pick(pickTwo, null, 2, 1, 3, 2) as string, "f", StringComparison.Ordinal), Convert.ToString(MergeSnapshot.Pick(pickTwo, null, 2, 1, 3, 2)));
            // 单格（rows==1 && columns==1）返回标量
            Assert("Pick 单格返回标量", string.Equals(MergeSnapshot.Pick(null, "x", 0, 0, 1, 1) as string, "x", StringComparison.Ordinal), "x");

            Console.WriteLine("=== 3. 跨表 / 跨簿引用：地址行偏移、列保持不变 ===");
            AssertRewrite("=Sheet2!A1", d4, 2, "=Sheet2!A3");
            AssertRewrite("='My Sheet'!A1", d4, 2, "='My Sheet'!A3");
            AssertRewrite("=[Book1.xlsx]Sheet1!A1", d4, 2, "=[Book1.xlsx]Sheet1!A3");
            AssertUnchanged("='D:\\p\\[工作簿1]Sheet1'!$J$8", d4, 2);   // 全绝对 → 仍不变
            AssertRewrite("=Sheet1:Sheet3!A1", d4, 2, "=Sheet1:Sheet3!A3");
            AssertRewrite("=SUM(Sheet2!A1:B2)", d4, 2, "=SUM(Sheet2!A3:B4)");

            Console.WriteLine("=== 3b. 跨簿带路径 / 整列 / 混合绝对 ===");
            AssertRewrite("=A2+'C:\\fixtures\\中文 目录\\[工32作簿1.xlsx]Sheet1'!A1",
                Mapper(1, 1), 5, "=A7+'C:\\fixtures\\中文 目录\\[工32作簿1.xlsx]Sheet1'!A6");
            AssertRewrite("=A6+'C:\\fixtures\\中文 目录\\[工32作簿1.xlsx]Sheet1'!A5",
                Mapper(1, 1), 5, "=A11+'C:\\fixtures\\中文 目录\\[工32作簿1.xlsx]Sheet1'!A10");
            AssertUnchanged("='C:\\fixtures\\中文 目录\\[工32作簿1.xlsx]Sheet1'!$A$1", Mapper(1, 1), 5);
            AssertRewrite("=Sheet1!$A1", Mapper(1, 4), 2, "=Sheet1!$A3");   // 列绝对、行相对 → 行偏移
            AssertRewrite("=Sheet1!A$1", Mapper(1, 4), 2, "=Sheet1!A$1");   // 列相对、行绝对 → 列不变
            AssertUnchanged("=SUM(Sheet1!A:A)", Mapper(1, 4), 2);           // 整列：列不变、无行 → 不变
            AssertRewrite("=SUM(Sheet1!$A1:B$2)", Mapper(1, 4), 2, "=SUM(Sheet1!$A3:B$2)"); // 区域各半绝对

            Console.WriteLine("=== 3c. 跨表/跨簿列物理偏移 + 跨表前缀加路径（RewriteContext）===");
            string srcPath = "D:\\fixtures\\中文 目录\\工作簿2.xlsx";
            // 本表引用 B2：列映射 B→F（源列2→目标列6），行 +4 → F6
            // 跨表引用 Sheet2!E18：物理列差 +4（E→I），行 +4（18→22），前缀加路径 → '路径[工作簿2.xlsx]Sheet2'!I22
            AssertRewriteCtx("=B2+Sheet2!E18", Mapper(2, 6), 4, 4, srcPath,
                "=F6+'D:\\fixtures\\中文 目录\\[工作簿2.xlsx]Sheet2'!I22");
            // 跨表列物理偏移：Sheet2!A1 + 物理列差 4 → Sheet2!E1，行 +2 → E3
            AssertRewriteCtx("=Sheet2!A1", Mapper(1, 1), 2, 4, srcPath,
                "='D:\\fixtures\\中文 目录\\[工作簿2.xlsx]Sheet2'!E3");
            // 跨簿引用（已含 [）：列物理偏移、行偏移，前缀不变
            AssertRewriteCtx("=[Book1.xlsx]Sheet1!A1", Mapper(1, 1), 2, 3, srcPath,
                "=[Book1.xlsx]Sheet1!D3");
            // 绝对引用跨表：列 $A 不动，行照偏移，前缀仍加路径
            AssertRewriteCtx("=Sheet2!$A1", Mapper(1, 1), 2, 4, srcPath,
                "='D:\\fixtures\\中文 目录\\[工作簿2.xlsx]Sheet2'!$A3");
            // 跨簿全绝对：完全不变
            AssertUnchanged("='D:\\p\\[工作簿1]Sheet1'!$J$8", d4, 2);

            Console.WriteLine("=== 4. 不该改写的区域 ===");
            AssertUnchanged("=\"A1\"", d4, 2);                 // 双引号字面量
            // 字面量里的 A1 不动，但 &B2 是真实引用，仍要按规则重写（B 未映射 → 保持列、行偏移）
            AssertRewrite("=\"A1\"&B2", d4, 2, "=\"A1\"&B4");
            AssertUnchanged("=1.5E10", d4, 2);                 // 科学计数法
            AssertUnchanged("=Total1*2", d4, 2);               // 标识符内数字
            AssertUnchanged("=#REF!", d4, 2);                  // 错误值
            AssertUnchanged("=NA()", d4, 2);
            AssertRewrite("=LOG10(A1)", d4, 2, "=LOG10(D3)");  // 函数名不被误判
            AssertRewrite("=ATAN2(A1,B1)", Mapper(1, 4, 2, 5), 2, "=ATAN2(D3,E3)");

            Console.WriteLine("=== 5. 区域与整列 ===");
            AssertRewrite("=SUM(A1:B2)", Mapper(1, 4, 2, 5), 2, "=SUM(D3:E4)");
            AssertRewrite("=SUM(A:A)", d4, 2, "=SUM(D:D)");
            AssertRewrite("=SUM($A:$A)", d4, 2, "=SUM($A:$A)");
            AssertRewrite("=$A$1:$B$2", d4, 2, "=$A$1:$B$2");

            Console.WriteLine("=== 6. 边界与降级 ===");
            // 列越界：保持原样并报告
            RewriteResult oob = FormulaRefRewriter.Rewrite("=A1", Mapper(1, 20000), 0, null);
            Assert("列越界保持原样", oob.Text == "=A1", oob.Text);
            Assert("列越界产生 ColumnOutOfRange", oob.Issues.Count == 1
                && oob.Issues[0].Kind == RefIssueKind.ColumnOutOfRange, "实际 " + oob.Issues.Count);
            // 行越界
            RewriteResult rowOob = FormulaRefRewriter.Rewrite("=A1", d4, -5, null);
            Assert("行越界保持原行号", rowOob.Text == "=D1", rowOob.Text);
            Assert("行越界产生 RowOutOfRange", rowOob.Issues.Count == 1
                && rowOob.Issues[0].Kind == RefIssueKind.RowOutOfRange, "实际 " + rowOob.Issues.Count);
            // 未映射列：保持原列字母，但行仍然偏移
            RewriteResult unmapped = FormulaRefRewriter.Rewrite("=A1+B1", d4, 2, null);
            Assert("未映射列保持列字母、行仍偏移", unmapped.Text == "=D3+B3", unmapped.Text);
            Assert("未映射列产生 UnmappedColumn", unmapped.Issues.Count == 1
                && unmapped.Issues[0].Kind == RefIssueKind.UnmappedColumn, "实际 " + unmapped.Issues.Count);
            // 空公式与无等号
            Assert("空公式", FormulaRefRewriter.Rewrite("", d4, 2, null).Text == "", "x");
            Assert("null 公式", FormulaRefRewriter.Rewrite(null, d4, 2, null).Text == "", "x");
            Assert("无等号也处理", FormulaRefRewriter.Rewrite("A1", d4, 2, null).Text == "D3", "x");
            // 定义名保护
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            names.Add("AB2");
            Assert("定义名保护", FormulaRefRewriter.Rewrite("=AB2", d4, 2, names).Text == "=AB2",
                FormulaRefRewriter.Rewrite("=AB2", d4, 2, names).Text);

            Console.WriteLine("=== 7. 掩码单测 ===");
            bool[] literal = FormulaRefRewriter.Literal("=\"A1\"&B2");
            Assert("Literal 覆盖 A1", literal[2] && literal[3], "x");
            bool[] quoted = FormulaRefRewriter.Quoted("'My Sheet'!A1");
            Assert("Quoted 覆盖表名", quoted[0] && quoted[8] && !quoted[10], "x");
            // "=Sheet2!A1" 长度 10（索引 0..9）：掩码只覆盖前缀 Sheet2!（索引 1..7），不含地址 A1（索引 8..9）
            bool[] qualified = FormulaRefRewriter.Qualified("=Sheet2!A1",
                FormulaRefRewriter.Literal("=Sheet2!A1"), FormulaRefRewriter.Quoted("=Sheet2!A1"));
            Assert("Qualified 覆盖前缀 Sheet2! 但不含地址 A1", qualified[1] && qualified[6] && qualified[7] && !qualified[8] && !qualified[9], "x");

            Console.WriteLine("=== 8. 表头匹配 ===");
            MergeHeaderSet target = Set(1, "A", "B", "C", "D", "E", "F", "G");
            MergeHeaderSet x2 = Set(1, "A", "B", "C");
            MergeHeaderSet x3 = Set(1, "D", "E");
            MergeMatchResult match = MergeMatcher.Match(target, new List<MergeHeaderSet> { x2, x3 });
            Assert("无阻断项", match.Blocking.Count == 0, string.Join(";", match.Blocking.ToArray()));
            Assert("两个来源", match.Sources.Count == 2, match.Sources.Count.ToString());
            Assert("X2 匹配 3 列", match.Sources[0].Maps.Count == 3, match.Sources[0].Maps.Count.ToString());
            Assert("X3 匹配 2 列", match.Sources[1].Maps.Count == 2, match.Sources[1].Maps.Count.ToString());
            Assert("X3 的 D 映射到目标第 4 列", match.Sources[1].Maps[0].TargetColumn == 4,
                match.Sources[1].Maps[0].TargetColumn.ToString());
            Assert("X3 的 E 映射到目标第 5 列", match.Sources[1].Maps[1].TargetColumn == 5,
                match.Sources[1].Maps[1].TargetColumn.ToString());
            Assert("共 5 列", match.MappedColumns == 5, match.MappedColumns.ToString());
            Assert("无缺失", !match.HasMissing, "有缺失");

            Console.WriteLine("=== 9. 缺列与异常表头 ===");
            MergeHeaderSet withMissing = Set(1, "A", "ZZ");
            MergeMatchResult missing = MergeMatcher.Match(target, new List<MergeHeaderSet> { withMissing });
            Assert("识别出 1 列缺失", missing.Sources[0].Missing.Count == 1, missing.Sources[0].Missing.Count.ToString());
            Assert("缺失列为 ZZ", missing.Sources[0].Missing[0] == "ZZ", missing.Sources[0].Missing[0]);
            Assert("HasMissing=true", missing.HasMissing, "false");
            Assert("缺列文本非空", MergePlanText.MissingText(missing).Length > 0, "空");

            MergeHeaderSet blank = Set(1, "A", "  ");
            MergeMatchResult blankResult = MergeMatcher.Match(target, new List<MergeHeaderSet> { blank });
            Assert("空列名被阻断", blankResult.Blocking.Count == 1, blankResult.Blocking.Count.ToString());

            MergeHeaderSet dup = Set(1, "A", "A");
            MergeMatchResult dupResult = MergeMatcher.Match(target, new List<MergeHeaderSet> { dup });
            Assert("重复列名被阻断", dupResult.Blocking.Count == 1, dupResult.Blocking.Count.ToString());

            MergeHeaderSet spaced = Set(1, " A ", "B");
            MergeMatchResult spacedResult = MergeMatcher.Match(target, new List<MergeHeaderSet> { spaced });
            Assert("忽略首尾空格后仍能匹配", spacedResult.Sources[0].Maps.Count == 2,
                spacedResult.Sources[0].Maps.Count.ToString());

            MergeHeaderSet cased = Set(1, "a", "b");
            MergeMatchResult casedResult = MergeMatcher.Match(target, new List<MergeHeaderSet> { cased });
            Assert("忽略大小写后仍能匹配", casedResult.Sources[0].Maps.Count == 2,
                casedResult.Sources[0].Maps.Count.ToString());

            Console.WriteLine("=== 10. 写入分块与保护 ===");
            List<MergeColumnMap> maps = new List<MergeColumnMap>();
            for (int i = 0; i < 3; i++)
            {
                MergeColumnMap map = new MergeColumnMap();
                map.SourceColumn = i + 1;
                map.TargetColumn = i + 1;
                map.RowCount = 10;
                map.TargetStartRow = 4;
                map.SourceFirstRow = 2;
                map.Cells = new object[10];
                for (int r = 0; r < 10; r++) map.Cells[r] = r;
                maps.Add(map);
            }
            MergeWritePlan plan = MergePlanner.Build(maps, 1);
            Assert("相邻同起始行合并为 1 块", plan.Blocks.Count == 1, plan.Blocks.Count.ToString());
            Assert("总格数 30", plan.TotalCells == 30, plan.TotalCells.ToString());
            Assert("块宽 3", plan.Blocks[0].ColumnCount == 3, plan.Blocks[0].ColumnCount.ToString());
            Assert("Guard 通过", MergePlanner.Guard(plan) == null, "被拒");

            maps[1].TargetStartRow = 9;   // 起始行不同 → 不能合并
            MergeWritePlan split = MergePlanner.Build(maps, 1);
            Assert("起始行不同则拆块", split.Blocks.Count == 3, split.Blocks.Count.ToString());

            List<MergeColumnMap> huge = new List<MergeColumnMap>();
            MergeColumnMap big = new MergeColumnMap();
            big.TargetColumn = 1; big.RowCount = MergeLimits.MaxCells + 1; big.TargetStartRow = 2;
            big.Cells = new object[big.RowCount];
            huge.Add(big);
            Assert("超大批量被 Guard 拒绝", MergePlanner.Guard(MergePlanner.Build(huge, 1)) != null, "未被拒");

            MergeColumnMap overflow = new MergeColumnMap();
            overflow.TargetColumn = 1; overflow.RowCount = 5; overflow.TargetStartRow = MergeLimits.MaxRow - 1;
            overflow.Cells = new object[5];
            List<MergeColumnMap> of = new List<MergeColumnMap>(); of.Add(overflow);
            Assert("超末行被 Guard 拒绝", MergePlanner.Guard(MergePlanner.Build(of, 1)) != null, "未被拒");

            Console.WriteLine("=== 11. 与 FormulaRefParser 交叉校验 ===");
            // 两套实现互为回归网：本地引用 ⇔ RefType.SameSheet。
            // 整列引用 A:A 是本次新增的能力，现成的 FormulaRefParser 不识别，
            // 故不参与交叉校验（其重写行为已在第 5 节单独覆盖）。
            string[] local = { "A1", "$A$1", "A1:B2" };
            foreach (string text in local)
            {
                PrecedentItem item = FormulaRefParser.ParseToken(text);
                Assert("ParseToken 判定 " + text + " 为同表引用",
                    item != null && item.RefType == RefType.SameSheet,
                    item == null ? "null" : item.RefType.ToString());
            }
            string[] foreign = { "Sheet2!A1", "'My Sheet'!A1" };
            foreach (string text in foreign)
            {
                PrecedentItem item = FormulaRefParser.ParseToken(text);
                Assert("ParseToken 判定 " + text + " 为跨表引用",
                    item != null && item.RefType != RefType.SameSheet,
                    item == null ? "null" : item.RefType.ToString());
                Assert("重写器对 " + text + " 行偏移列不变",
                    FormulaRefRewriter.Rewrite("=" + text, d4, 2, null).Text == "=" + text.Replace("A1", "A3"), "被改写了");
            }

            Console.WriteLine();
            Console.WriteLine("PASS " + _pass + "，FAIL " + _fail);
            Environment.Exit(_fail == 0 ? 0 : 1);
        }
    }
}
