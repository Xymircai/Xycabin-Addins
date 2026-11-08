using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Excel_Iterate;

namespace Excel_Iterate.TestHarness
{
    // =====================================================================
    // EliminationProbe — 合并抵消（往来抵消 / 收入成本抵消）离线自测
    // （不连 WPS 表格，不需要 COM）
    //
    // 覆盖：
    //   1. EliminationEngine：金额/比例解析、角色猜测、映射校验、
    //      往来抵消（聚合/对冲/坏账按比例冲回/跳过）、收入成本抵消
    //      （U=(R−C)×比例 的三种极端：比例 0 / 100% / 中间值）
    //   2. EliminationGateway 静态方法：分录 → 6 列写入计划、起始单元格解析
    //
    // 关键期望值（与用户确认的会计模型一致，改模型必须同步改这里）：
    //   R=100000, C=60000, p=100% → U=40000，贷营业成本 60000、贷存货 40000
    //   R=100000, C=60000, p=0   → U=0，贷营业成本 100000、无存货行
    //   坏账冲回 = 坏账准备 × 对冲额 ÷ 应收
    //
    // 编译：build.py 的 tools 段（/r:Excel_Iterate.dll）
    // 运行：build/EliminationProbe.exe  退出码 0=全 PASS，1=有 FAIL
    // =====================================================================

    internal static class EliminationProbe
    {
        static int _pass = 0, _fail = 0;

        static void Assert(string name, bool cond, string detail)
        {
            if (cond) { _pass++; Console.WriteLine("  [PASS] " + name); }
            else { _fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
        }

        /// <summary>按 (角色, 文本, 角色, 文本, …) 构造一行源数据。</summary>
        static EliminationRow Row(int sourceRow, params object[] roleText)
        {
            EliminationRow row = new EliminationRow();
            row.SourceRow = sourceRow;
            for (int i = 0; i + 1 < roleText.Length; i += 2)
                row.Values[(EliminationRole)roleText[i]] = Convert.ToString(roleText[i + 1]);
            return row;
        }

        static bool HasLine(EliminationResult result, bool debit, string account, decimal amount)
        {
            foreach (EliminationEntry entry in result.Entries)
                foreach (EliminationLine line in entry.Lines)
                    if (line.IsDebit == debit && line.Account == account && line.Amount == amount) return true;
            return false;
        }

        static int CountAccount(EliminationResult result, string account)
        {
            int count = 0;
            foreach (EliminationEntry entry in result.Entries)
                foreach (EliminationLine line in entry.Lines)
                    if (line.Account == account) count++;
            return count;
        }

        static bool WarningContains(EliminationResult result, string token)
        {
            foreach (string warning in result.Warnings)
                if (warning.Contains(token)) return true;
            return false;
        }

        static void Main()
        {
            Console.WriteLine("=== EliminationProbe — 合并抵消离线自测 ===");
            Console.WriteLine();

            // ===== 1. 金额解析 =====
            Console.WriteLine("--- 1. TryParseAmount ---");
            decimal amount;
            Assert("1234.56", EliminationEngine.TryParseAmount("1234.56", out amount) && amount == 1234.56m, "实际 " + amount);
            Assert("千分位 1,234.56", EliminationEngine.TryParseAmount("1,234.56", out amount) && amount == 1234.56m, "实际 " + amount);
            Assert("货币符号 ¥1,234.56 元", EliminationEngine.TryParseAmount("¥1,234.56 元", out amount) && amount == 1234.56m, "实际 " + amount);
            Assert("括号负数 (500)", EliminationEngine.TryParseAmount("(500)", out amount) && amount == -500m, "实际 " + amount);
            Assert("负号 -500", EliminationEngine.TryParseAmount("-500", out amount) && amount == -500m, "实际 " + amount);
            Assert("拒绝 abc", !EliminationEngine.TryParseAmount("abc", out amount), "实际 " + amount);
            Assert("拒绝空串", !EliminationEngine.TryParseAmount("", out amount), "实际 " + amount);
            Assert("拒绝 12/34", !EliminationEngine.TryParseAmount("12/34", out amount), "实际 " + amount);
            Console.WriteLine();

            // ===== 2. 比例解析 =====
            Console.WriteLine("--- 2. TryParseRatio ---");
            decimal ratio;
            Assert("30% → 0.3", EliminationEngine.TryParseRatio("30%", out ratio) && ratio == 0.3m, "实际 " + ratio);
            Assert("0.25 → 0.25", EliminationEngine.TryParseRatio("0.25", out ratio) && ratio == 0.25m, "实际 " + ratio);
            Assert("100% → 1", EliminationEngine.TryParseRatio("100%", out ratio) && ratio == 1m, "实际 " + ratio);
            Assert("裸数字 30 拒绝（歧义）", !EliminationEngine.TryParseRatio("30", out ratio), "实际 " + ratio);
            Assert("150% 越界拒绝", !EliminationEngine.TryParseRatio("150%", out ratio), "实际 " + ratio);
            Assert("-10% 越界拒绝", !EliminationEngine.TryParseRatio("-10%", out ratio), "实际 " + ratio);
            Console.WriteLine();

            // ===== 3. 角色猜测 =====
            Console.WriteLine("--- 3. GuessRole ---");
            Assert("坏账准备 → BadDebt", EliminationEngine.GuessRole("坏账准备", EliminationKind.Receivable) == EliminationRole.BadDebt, "");
            Assert("应收账款 → Receivable", EliminationEngine.GuessRole("应收账款", EliminationKind.Receivable) == EliminationRole.Receivable, "");
            Assert("应付账款 → Payable", EliminationEngine.GuessRole("应付账款", EliminationKind.Receivable) == EliminationRole.Payable, "");
            Assert("未对外销售比例 → UnsoldRatio", EliminationEngine.GuessRole("未对外销售比例", EliminationKind.RevenueCost) == EliminationRole.UnsoldRatio, "");
            Assert("营业收入 不再自动匹配（收紧为单体收入）", EliminationEngine.GuessRole("营业收入", EliminationKind.RevenueCost) == EliminationRole.Ignore, "");
            Assert("营业成本 不再自动匹配（收紧为单体成本）", EliminationEngine.GuessRole("营业成本", EliminationKind.RevenueCost) == EliminationRole.Ignore, "");
            Assert("单位名称 → Unit", EliminationEngine.GuessRole("单位名称", EliminationKind.Receivable) == EliminationRole.Unit, "");
            Assert("备注 → Ignore", EliminationEngine.GuessRole("备注", EliminationKind.Receivable) == EliminationRole.Ignore, "");
            Assert("空表头 → Ignore", EliminationEngine.GuessRole("  ", EliminationKind.Receivable) == EliminationRole.Ignore, "");
            Console.WriteLine();

            // ===== 4. 映射校验 =====
            Console.WriteLine("--- 4. ValidateMapping ---");
            // 往来抵消改走矩阵式后，必填角色为 核算主体/对方单位/科目/审定数/截止日
            List<EliminationColumn> missing = new List<EliminationColumn>();
            missing.Add(new EliminationColumn { Number = 2, Letter = "B", Header = "核算主体", Role = EliminationRole.Entity });
            missing.Add(new EliminationColumn { Number = 4, Letter = "D", Header = "对方单位", Role = EliminationRole.Counterparty });
            missing.Add(new EliminationColumn { Number = 3, Letter = "C", Header = "科目", Role = EliminationRole.Account });
            missing.Add(new EliminationColumn { Number = 21, Letter = "U", Header = "审定数", Role = EliminationRole.Amount });
            string mapError = EliminationEngine.ValidateMapping(EliminationKind.Receivable, missing);
            Assert("往来缺必填「截止日」被拒绝", mapError != null && mapError.Contains("截止日"), mapError ?? "null");

            List<EliminationColumn> duplicate = new List<EliminationColumn>();
            duplicate.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "单位", Role = EliminationRole.Unit });
            duplicate.Add(new EliminationColumn { Number = 2, Letter = "B", Header = "应收账款", Role = EliminationRole.Receivable });
            duplicate.Add(new EliminationColumn { Number = 3, Letter = "C", Header = "应收余额", Role = EliminationRole.Receivable });
            duplicate.Add(new EliminationColumn { Number = 4, Letter = "D", Header = "应付账款", Role = EliminationRole.Payable });
            mapError = EliminationEngine.ValidateMapping(EliminationKind.Receivable, duplicate);
            Assert("一角色两列被拒绝", mapError != null && mapError.Contains("都指定为"), mapError ?? "null");

            List<EliminationColumn> legal = new List<EliminationColumn>();
            legal.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "截止日", Role = EliminationRole.Deadline });
            legal.Add(new EliminationColumn { Number = 2, Letter = "B", Header = "核算主体", Role = EliminationRole.Entity });
            legal.Add(new EliminationColumn { Number = 3, Letter = "C", Header = "科目", Role = EliminationRole.Account });
            legal.Add(new EliminationColumn { Number = 4, Letter = "D", Header = "对方单位", Role = EliminationRole.Counterparty });
            legal.Add(new EliminationColumn { Number = 6, Letter = "F", Header = "是否合并内关联方", Role = EliminationRole.Condition });
            legal.Add(new EliminationColumn { Number = 21, Letter = "U", Header = "审定数", Role = EliminationRole.Amount });
            Assert("往来合法映射通过", EliminationEngine.ValidateMapping(EliminationKind.Receivable, legal) == null, "");
            Assert("往来：ValidateMatrixMapping 直接放行（不涉及损益来源）",
                EliminationEngine.ValidateMatrixMapping(EliminationKind.Receivable, legal) == null, "");

            // 收入成本抵消改走矩阵式底稿后，必填角色为 核算主体/客商/年份/收入/成本；
            // 「未对外销售比例」降为软性，改由 ValidateMatrixMapping 与损益列二选一校验。
            List<EliminationColumn> revenueMissing = new List<EliminationColumn>();
            revenueMissing.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "收入", Role = EliminationRole.Revenue });
            revenueMissing.Add(new EliminationColumn { Number = 2, Letter = "B", Header = "成本", Role = EliminationRole.Cost });
            mapError = EliminationEngine.ValidateMapping(EliminationKind.RevenueCost, revenueMissing);
            Assert("收入成本缺「核算主体」被拒绝", mapError != null && mapError.Contains("核算主体"), mapError ?? "null");

            List<EliminationColumn> matrixFull = new List<EliminationColumn>();
            matrixFull.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "核算主体", Role = EliminationRole.Entity });
            matrixFull.Add(new EliminationColumn { Number = 10, Letter = "J", Header = "客户名称", Role = EliminationRole.Counterparty });
            matrixFull.Add(new EliminationColumn { Number = 5, Letter = "E", Header = "年份", Role = EliminationRole.Year });
            matrixFull.Add(new EliminationColumn { Number = 14, Letter = "N", Header = "是否为合并内关联方", Role = EliminationRole.Condition });
            matrixFull.Add(new EliminationColumn { Number = 15, Letter = "O", Header = "单体收入（未审）", Role = EliminationRole.Revenue });
            matrixFull.Add(new EliminationColumn { Number = 16, Letter = "P", Header = "单体成本（未审）", Role = EliminationRole.Cost });
            matrixFull.Add(new EliminationColumn { Number = 22, Letter = "V", Header = "未对外销售比例", Role = EliminationRole.UnsoldRatio });
            Assert("矩阵式合法映射通过", EliminationEngine.ValidateMapping(EliminationKind.RevenueCost, matrixFull) == null, "");
            Assert("矩阵式：有比例列即可算损益", EliminationEngine.ValidateMatrixMapping(EliminationKind.RevenueCost, matrixFull) == null, "");

            List<EliminationColumn> noProfitSource = new List<EliminationColumn>();
            noProfitSource.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "核算主体", Role = EliminationRole.Entity });
            noProfitSource.Add(new EliminationColumn { Number = 10, Letter = "J", Header = "客户名称", Role = EliminationRole.Counterparty });
            noProfitSource.Add(new EliminationColumn { Number = 5, Letter = "E", Header = "年份", Role = EliminationRole.Year });
            noProfitSource.Add(new EliminationColumn { Number = 15, Letter = "O", Header = "单体收入（未审）", Role = EliminationRole.Revenue });
            noProfitSource.Add(new EliminationColumn { Number = 16, Letter = "P", Header = "单体成本（未审）", Role = EliminationRole.Cost });
            Assert("损益列与比例列全缺被拒绝", EliminationEngine.ValidateMatrixMapping(EliminationKind.RevenueCost, noProfitSource) != null, "");
            Console.WriteLine();

            // ===== 5. 往来抵消 =====
            Console.WriteLine("--- 5. BuildReceivable ---");

            // 场景A：应收 100000 / 应付 60000 → 对冲 60000，一条分录
            List<EliminationRow> sceneA = new List<EliminationRow>();
            sceneA.Add(Row(2, EliminationRole.Unit, "甲公司", EliminationRole.Receivable, "100000",
                EliminationRole.Payable, "60000", EliminationRole.BadDebt, "0"));
            EliminationResult resultA = EliminationEngine.BuildReceivable(sceneA, false);
            Assert("A：1 条分录", resultA.Entries.Count == 1, "实际 " + resultA.Entries.Count);
            Assert("A：借 应付账款 60000", HasLine(resultA, true, EliminationEngine.AccPayable, 60000m), "");
            Assert("A：贷 应收账款 60000", HasLine(resultA, false, EliminationEngine.AccReceivable, 60000m), "");
            Assert("A：借贷平衡", resultA.Balanced, "总借 " + resultA.TotalDebit + " / 总贷 " + resultA.TotalCredit);
            Assert("A：无坏账准备行（未含坏账列）", CountAccount(resultA, EliminationEngine.AccBadDebt) == 0, "");

            // 场景B：应收 100000 / 应付 40000 / 坏账准备 5000
            // 对冲 40000；坏账冲回 = 5000 × 40000 ÷ 100000 = 2000（第二条分录）
            List<EliminationRow> sceneB = new List<EliminationRow>();
            sceneB.Add(Row(2, EliminationRole.Unit, "乙公司", EliminationRole.Receivable, "100000",
                EliminationRole.Payable, "40000", EliminationRole.BadDebt, "5000"));
            EliminationResult resultB = EliminationEngine.BuildReceivable(sceneB, true);
            Assert("B：2 条分录（对冲 + 坏账冲回）", resultB.Entries.Count == 2, "实际 " + resultB.Entries.Count);
            Assert("B：借 应付账款 40000", HasLine(resultB, true, EliminationEngine.AccPayable, 40000m), "");
            Assert("B：借 坏账准备 2000", HasLine(resultB, true, EliminationEngine.AccBadDebt, 2000m), "");
            Assert("B：贷 信用减值损失 2000", HasLine(resultB, false, EliminationEngine.AccCreditLoss, 2000m), "");
            Assert("B：借贷平衡", resultB.Balanced, "总借 " + resultB.TotalDebit + " / 总贷 " + resultB.TotalCredit);

            // 场景C：应收 30000 / 应付 80000（应收<应付）→ 无坏账冲回
            List<EliminationRow> sceneC = new List<EliminationRow>();
            sceneC.Add(Row(2, EliminationRole.Unit, "丙公司", EliminationRole.Receivable, "30000",
                EliminationRole.Payable, "80000", EliminationRole.BadDebt, "1000"));
            EliminationResult resultC = EliminationEngine.BuildReceivable(sceneC, true);
            Assert("C：1 条分录（无坏账冲回）", resultC.Entries.Count == 1, "实际 " + resultC.Entries.Count);
            Assert("C：借 应付账款 30000", HasLine(resultC, true, EliminationEngine.AccPayable, 30000m), "");
            Assert("C：无坏账准备行", CountAccount(resultC, EliminationEngine.AccBadDebt) == 0, "");

            // 场景D：多行聚合 + 无对冲 + 无法解析 三类跳过
            List<EliminationRow> sceneD = new List<EliminationRow>();
            sceneD.Add(Row(2, EliminationRole.Unit, "丁公司", EliminationRole.Receivable, "50000", EliminationRole.Payable, "0"));
            sceneD.Add(Row(3, EliminationRole.Unit, "丁公司", EliminationRole.Receivable, "0", EliminationRole.Payable, "30000"));
            sceneD.Add(Row(4, EliminationRole.Unit, "戊公司", EliminationRole.Receivable, "0", EliminationRole.Payable, "0"));
            sceneD.Add(Row(5, EliminationRole.Unit, "己公司", EliminationRole.Receivable, "abc", EliminationRole.Payable, "100"));
            sceneD.Add(Row(6, EliminationRole.Unit, "", EliminationRole.Receivable, "100", EliminationRole.Payable, "100"));
            EliminationResult resultD = EliminationEngine.BuildReceivable(sceneD, false);
            Assert("D：丁公司两行聚合成对冲 30000", HasLine(resultD, true, EliminationEngine.AccPayable, 30000m), "");
            Assert("D：仅 1 条分录", resultD.Entries.Count == 1, "实际 " + resultD.Entries.Count);
            Assert("D：跳过 3 行", resultD.SkippedRows == 3, "实际 " + resultD.SkippedRows);
            Assert("D：提示含「无内部对冲额」", WarningContains(resultD, "无内部对冲额"), "");
            Assert("D：提示含「无法解析」", WarningContains(resultD, "无法解析"), "");
            Assert("D：提示含「缺少单位名称」", WarningContains(resultD, "缺少单位名称"), "");
            Assert("D：借贷平衡", resultD.Balanced, "");
            Console.WriteLine();

            // ===== 6. 收入成本抵消 =====
            Console.WriteLine("--- 6. BuildRevenueCost ---");

            // 场景E：R=100000 C=60000 比例100% → U=40000；贷成本 60000、贷存货 40000
            List<EliminationRow> sceneE = new List<EliminationRow>();
            sceneE.Add(Row(2, EliminationRole.Unit, "庚公司", EliminationRole.Revenue, "100000",
                EliminationRole.Cost, "60000", EliminationRole.UnsoldRatio, "100%"));
            EliminationResult resultE = EliminationEngine.BuildRevenueCost(sceneE);
            Assert("E：1 条分录", resultE.Entries.Count == 1, "实际 " + resultE.Entries.Count);
            Assert("E：借 营业收入 100000", HasLine(resultE, true, EliminationEngine.AccRevenue, 100000m), "");
            Assert("E：贷 营业成本 60000", HasLine(resultE, false, EliminationEngine.AccCost, 60000m), "");
            Assert("E：贷 存货 40000", HasLine(resultE, false, EliminationEngine.AccInventory, 40000m), "");
            Assert("E：借贷平衡", resultE.Balanced, "总借 " + resultE.TotalDebit + " / 总贷 " + resultE.TotalCredit);

            // 场景F：比例 0 → U=0，全额反转收入成本，无存货行
            List<EliminationRow> sceneF = new List<EliminationRow>();
            sceneF.Add(Row(2, EliminationRole.Unit, "辛公司", EliminationRole.Revenue, "100000",
                EliminationRole.Cost, "60000", EliminationRole.UnsoldRatio, "0"));
            EliminationResult resultF = EliminationEngine.BuildRevenueCost(sceneF);
            Assert("F：贷 营业成本 100000", HasLine(resultF, false, EliminationEngine.AccCost, 100000m), "");
            Assert("F：无存货行", CountAccount(resultF, EliminationEngine.AccInventory) == 0, "");
            Assert("F：分录 2 行（借收入 + 贷成本）", resultF.Entries.Count == 1 && resultF.Entries[0].Lines.Count == 2,
                "实际 " + resultF.Entries[0].Lines.Count + " 行");
            Assert("F：借贷平衡", resultF.Balanced, "");

            // 场景G：R=200000 C=150000 比例30% → U=15000；贷成本 185000、贷存货 15000
            List<EliminationRow> sceneG = new List<EliminationRow>();
            sceneG.Add(Row(2, EliminationRole.Unit, "壬公司", EliminationRole.Revenue, "200000",
                EliminationRole.Cost, "150000", EliminationRole.UnsoldRatio, "30%"));
            EliminationResult resultG = EliminationEngine.BuildRevenueCost(sceneG);
            Assert("G：贷 营业成本 185000", HasLine(resultG, false, EliminationEngine.AccCost, 185000m), "");
            Assert("G：贷 存货 15000", HasLine(resultG, false, EliminationEngine.AccInventory, 15000m), "");
            Assert("G：借贷平衡", resultG.Balanced, "");

            // 场景H：R=C=50000 比例50% → U=0（金额 0 的行不入结果）
            List<EliminationRow> sceneH = new List<EliminationRow>();
            sceneH.Add(Row(2, EliminationRole.Unit, "癸公司", EliminationRole.Revenue, "50000",
                EliminationRole.Cost, "50000", EliminationRole.UnsoldRatio, "50%"));
            EliminationResult resultH = EliminationEngine.BuildRevenueCost(sceneH);
            Assert("H：无存货行（U=0）", CountAccount(resultH, EliminationEngine.AccInventory) == 0, "");
            Assert("H：贷 营业成本 50000", HasLine(resultH, false, EliminationEngine.AccCost, 50000m), "");
            Assert("H：借贷平衡", resultH.Balanced, "");

            // 场景I：收入≤0 与无法解析 → 跳过并警告
            List<EliminationRow> sceneI = new List<EliminationRow>();
            sceneI.Add(Row(2, EliminationRole.Unit, "甲公司", EliminationRole.Revenue, "-5",
                EliminationRole.Cost, "10", EliminationRole.UnsoldRatio, "50%"));
            sceneI.Add(Row(3, EliminationRole.Unit, "乙公司", EliminationRole.Revenue, "100",
                EliminationRole.Cost, "abc", EliminationRole.UnsoldRatio, "50%"));
            EliminationResult resultI = EliminationEngine.BuildRevenueCost(sceneI);
            Assert("I：无分录", resultI.Entries.Count == 0, "实际 " + resultI.Entries.Count);
            Assert("I：跳过 2 行", resultI.SkippedRows == 2, "实际 " + resultI.SkippedRows);
            Assert("I：提示含「无抵消意义」", WarningContains(resultI, "无抵消意义"), "");
            Assert("I：提示含「无法解析」", WarningContains(resultI, "无法解析"), "");

            // 空数据 → 空结果不抛
            EliminationResult empty = EliminationEngine.BuildReceivable(new List<EliminationRow>(), true);
            Assert("空数据：无分录无跳过", empty.Entries.Count == 0 && empty.SkippedRows == 0 && empty.Balanced, "");
            Console.WriteLine();

            // ===== 7. 写入计划 =====
            Console.WriteLine("--- 7. BuildWritePlan（6 列：分录号/摘要/方向/科目/金额/计算说明）---");
            // resultB：2 条分录各 2 行 → 表头 1 行 + 借贷 4 行 = 5 行
            MergeWritePlan plan = EliminationGateway.BuildWritePlan(resultB, 3, 5);
            Assert("计划 1 个块", plan.Blocks.Count == 1, "实际 " + plan.Blocks.Count);
            MergeWriteBlock block = plan.Blocks.Count > 0 ? plan.Blocks[0] : null;
            Assert("块：5 行（表头 + 4 借贷行）× 6 列", block != null && block.RowCount == 5 && block.ColumnCount == 6,
                block == null ? "无块" : block.RowCount + "×" + block.ColumnCount);
            Assert("块：起始 (3,5)", block != null && block.FirstColumn == 3 && block.StartRow == 5, "");
            if (block != null)
            {
                Assert("表头行：分录号/摘要/方向/科目/金额/计算说明",
                    Convert.ToString(block.Cells[0, 0]) == "分录号" && Convert.ToString(block.Cells[0, 1]) == "摘要"
                    && Convert.ToString(block.Cells[0, 2]) == "方向" && Convert.ToString(block.Cells[0, 3]) == "科目"
                    && Convert.ToString(block.Cells[0, 4]) == "金额" && Convert.ToString(block.Cells[0, 5]) == "计算说明", "");
                Assert("行1：分录号 1 / 借 / 应付账款 / 40000",
                    Convert.ToInt32(block.Cells[1, 0]) == 1 && Convert.ToString(block.Cells[1, 2]) == "借"
                    && Convert.ToString(block.Cells[1, 3]) == EliminationEngine.AccPayable
                    && Convert.ToDouble(block.Cells[1, 4]) == 40000d, "");
                Assert("行2：贷 / 应收账款 / 40000",
                    Convert.ToString(block.Cells[2, 2]) == "贷"
                    && Convert.ToString(block.Cells[2, 3]) == EliminationEngine.AccReceivable
                    && Convert.ToDouble(block.Cells[2, 4]) == 40000d, "");
                Assert("行3：分录号 2 / 借 / 坏账准备 / 2000",
                    Convert.ToInt32(block.Cells[3, 0]) == 2 && Convert.ToString(block.Cells[3, 2]) == "借"
                    && Convert.ToString(block.Cells[3, 3]) == EliminationEngine.AccBadDebt
                    && Convert.ToDouble(block.Cells[3, 4]) == 2000d, "");
                Assert("行4：贷 / 信用减值损失 / 2000",
                    Convert.ToString(block.Cells[4, 2]) == "贷"
                    && Convert.ToString(block.Cells[4, 3]) == EliminationEngine.AccCreditLoss
                    && Convert.ToDouble(block.Cells[4, 4]) == 2000d, "");
                Assert("摘要列含单位名", Convert.ToString(block.Cells[1, 1]).Contains("乙公司"), "");
            }
            Assert("计划总量 30 格", plan.TotalCells == 30, "实际 " + plan.TotalCells);
            Assert("空结果 → 空计划", EliminationGateway.BuildWritePlan(empty, 1, 1).Blocks.Count == 0, "");
            Console.WriteLine();

            // ===== 8. 起始单元格解析 =====
            Console.WriteLine("--- 8. TryParseCell ---");
            int row, column;
            Assert("A1 → (1,1)", EliminationGateway.TryParseCell("A1", out row, out column) && row == 1 && column == 1, "");
            Assert("B3 → (3,2)", EliminationGateway.TryParseCell("B3", out row, out column) && row == 3 && column == 2, "");
            Assert("AA10 → (10,27)", EliminationGateway.TryParseCell("AA10", out row, out column) && row == 10 && column == 27, "");
            Assert("小写 xfd1048576 → 末行末列", EliminationGateway.TryParseCell("xfd1048576", out row, out column)
                && row == 1048576 && column == 16384, "");
            Assert("拒绝 1A", !EliminationGateway.TryParseCell("1A", out row, out column), "");
            Assert("拒绝 A", !EliminationGateway.TryParseCell("A", out row, out column), "");
            Assert("拒绝空", !EliminationGateway.TryParseCell("", out row, out column), "");
            Assert("拒绝 A0", !EliminationGateway.TryParseCell("A0", out row, out column), "");
            Assert("拒绝 A1B", !EliminationGateway.TryParseCell("A1B", out row, out column), "");
            Assert("拒绝超宽列 XFE1", !EliminationGateway.TryParseCell("XFE1", out row, out column), "");
            Console.WriteLine();

            // ===== 9. 表头常量 =====
            Console.WriteLine("--- 9. EntryHeaders 常量 ---");
            Assert("6 列表头", EliminationGateway.EntryHeaders.Length == 6, "实际 " + EliminationGateway.EntryHeaders.Length);
            Assert("金额列在第 5 位", EliminationGateway.EntryHeaders[4] == "金额", EliminationGateway.EntryHeaders[4]);
            Assert("NewSheetItem 非空", EliminationGateway.NewSheetItem.Length > 0, "");
            Assert("DefaultSheetName=抵消分录", EliminationGateway.DefaultSheetName == "抵消分录", EliminationGateway.DefaultSheetName);

            // ===== 10. 矩阵式底稿（收入成本抵消）=====
            Console.WriteLine("--- 10. 矩阵式底稿：三维聚合与恒等式 ---");

            // 源列布局对齐真实「内部交易明细」：
            //   A 核算主体 / E 年份 / J 客户名称 / N 是否为合并内关联方
            //   O 单体收入（未审）/ P 单体成本（未审）/ V 未对外销售比例 / W 未实现内部交易损益（存货）
            // 口径铁律：W 的公式是 (O−P)×V，故 Σ收入/Σ成本 也取未审列 O/P。
            List<EliminationColumn> matrixMap = MatrixMap(true);
            List<EliminationRow> matrixRows = new List<EliminationRow>();

            // 组1 · 销售主体A→客户A · 2024：无损益列值 → 现算 (O−P)×V = 73714.038
            matrixRows.Add(Row(3, EliminationRole.Entity, "销售主体A", EliminationRole.Year, "2024",
                EliminationRole.Counterparty, "客户A", EliminationRole.Condition, "是",
                EliminationRole.Revenue, "2217860.59", EliminationRole.Cost, "1849290.40",
                EliminationRole.UnsoldRatio, "20%"));
            // 组1 · 2025：有损益列 → 直接 ΣW（负值，验证 U<0 时恒等式仍成立）
            matrixRows.Add(Row(4, EliminationRole.Entity, "销售主体A", EliminationRole.Year, "2025",
                EliminationRole.Counterparty, "客户A", EliminationRole.Condition, "是",
                EliminationRole.Revenue, "1998483.64", EliminationRole.Cost, "2258286.52",
                EliminationRole.UnsoldRatio, "0%", EliminationRole.UnrealizedProfit, "-46492.5225"));
            // 组2 · 销售主体B→客户B · 2024：V=0% → 现算 U=0
            matrixRows.Add(Row(5, EliminationRole.Entity, "销售主体B", EliminationRole.Year, "2024",
                EliminationRole.Counterparty, "客户B", EliminationRole.Condition, "是",
                EliminationRole.Revenue, "151503.00", EliminationRole.Cost, "147503.00",
                EliminationRole.UnsoldRatio, "0%"));
            // 干扰行：条件列为「否」，必须被筛掉
            matrixRows.Add(Row(6, EliminationRole.Entity, "销售主体C", EliminationRole.Year, "2025",
                EliminationRole.Counterparty, "客户C", EliminationRole.Condition, "否",
                EliminationRole.Revenue, "999999", EliminationRole.Cost, "1",
                EliminationRole.UnsoldRatio, "50%"));

            EliminationMatrix matrix = EliminationMatrixBuilder.Build(matrixRows, matrixMap, MatrixOpt(true));

            Assert("年度降序 [2025,2024]",
                matrix.Years.Count == 2 && matrix.Years[0] == "2025" && matrix.Years[1] == "2024",
                string.Join(",", matrix.Years.ToArray()));
            Assert("2 组交易对", matrix.GroupCount == 2, "实际 " + matrix.GroupCount);
            Assert("每组 4 行 → 共 8 行", matrix.Rows.Count == 8, "实际 " + matrix.Rows.Count);
            Assert("列数 = 5 + 3×2 = 11", matrix.ColumnCount == 11, "实际 " + matrix.ColumnCount);
            Assert("2025 借/贷列在 G/H",
                matrix.DebitColumn("2025") != null && matrix.DebitColumn("2025").Letter == "G"
                && matrix.CreditColumn("2025").Letter == "H", "");
            Assert("2024 借/贷列在 J/K",
                matrix.DebitColumn("2024").Letter == "J" && matrix.CreditColumn("2024").Letter == "K", "");
            Assert("条件列筛掉 1 行", matrix.FilteredRows == 1, "实际 " + matrix.FilteredRows);
            Assert("无跳过行", matrix.SkippedRows == 0, "实际 " + matrix.SkippedRows);

            Assert("组1·2024 收入 R = 2217860.59",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Revenue", "2024").Value, 2217860.59m),
                MatrixCell(matrix, "销售主体A", "客户A", "Revenue", "2024").Value.ToString());
            Assert("组1·2024 成本 C = 1849290.40",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Cost", "2024").Value, 1849290.40m), "");
            Assert("组1·2024 存货 U = 73714.04（现算分支）",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Unrealized", "2024").Value, 73714.04m),
                MatrixCell(matrix, "销售主体A", "客户A", "Unrealized", "2024").Value.ToString());
            Assert("组1·2024 成本还原 = 294856.15",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Restore", "2024").Value, 294856.15m),
                MatrixCell(matrix, "销售主体A", "客户A", "Restore", "2024").Value.ToString());

            Assert("组1·2025 收入 R = 1998483.64",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Revenue", "2025").Value, 1998483.64m), "");
            Assert("组1·2025 成本 C = 2258286.52",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Cost", "2025").Value, 2258286.52m), "");
            Assert("组1·2025 存货 U = −46492.52（损益列分支，负值）",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Unrealized", "2025").Value, -46492.52m),
                MatrixCell(matrix, "销售主体A", "客户A", "Unrealized", "2025").Value.ToString());
            Assert("组1·2025 成本还原 = −213310.36",
                Near(MatrixCell(matrix, "销售主体A", "客户A", "Restore", "2025").Value, -213310.36m),
                MatrixCell(matrix, "销售主体A", "客户A", "Restore", "2025").Value.ToString());

            Assert("组2·2024 收入 R = 151503.00",
                Near(MatrixCell(matrix, "销售主体B", "客户B", "Revenue", "2024").Value, 151503.00m), "");
            Assert("组2·2024 成本还原 = 4000.00（U=0）",
                Near(MatrixCell(matrix, "销售主体B", "客户B", "Restore", "2024").Value, 4000.00m), "");

            Assert("恒等式成立（每组每年借贷平衡）", matrix.Balanced, "");
            Assert("无阻断错误", matrix.Errors.Count == 0, string.Join("；", matrix.Errors.ToArray()));

            // U 的两个来源：不映射损益列时，2025 也应改由现算得出（V=0% → 0）
            EliminationMatrix noProfit = EliminationMatrixBuilder.Build(matrixRows, MatrixMap(false), MatrixOpt(true));

            // ---- 公式文本（整列引用，锚点：主体/客商列绝对行相对，年度表头列相对行绝对）----
            string revenue2025 = MatrixCell(matrix, "销售主体A", "客户A", "Revenue", "2025").Formula;
            Assert("收入行公式 = SUMIFS(整列 O，锚点 $A4/$B4/G$2/\"是\")",
                revenue2025 == "=SUMIFS(内部交易明细!$O:$O,内部交易明细!$A:$A,$A4,"
                    + "内部交易明细!$J:$J,$B4,内部交易明细!$E:$E,G$2,内部交易明细!$N:$N,\"是\")",
                revenue2025 ?? "null");
            string cost2025 = MatrixCell(matrix, "销售主体A", "客户A", "Cost", "2025").Formula;
            Assert("成本行公式引用整列 P 且锚点为自身行 $A5",
                cost2025 != null && cost2025.Contains("$P:$P") && cost2025.Contains("$A5"),
                cost2025 ?? "null");
            string restore2025 = MatrixCell(matrix, "销售主体A", "客户A", "Restore", "2025").Formula;
            Assert("成本还原行引用本组单元格 =(G4-H5)-H7", restore2025 == "=(G4-H5)-H7", restore2025 ?? "null");
            string unrealized2025 = MatrixCell(matrix, "销售主体A", "客户A", "Unrealized", "2025").Formula;
            Assert("存货行有损益列时走 SUMIFS(W)",
                unrealized2025 != null && unrealized2025.StartsWith("=SUMIFS(内部交易明细!$W:$W"),
                unrealized2025 ?? "null");
            string revenue2024 = MatrixCell(matrix, "销售主体A", "客户A", "Revenue", "2024").Formula;
            Assert("2024 年度锚点切到 J$2", revenue2024 != null && revenue2024.Contains(",J$2,"), revenue2024 ?? "null");
            string revenueNoProfit2024 = MatrixCell(noProfit, "销售主体A", "客户A", "Unrealized", "2024").Formula;
            // 实机验收教训：①整列引用会把源表头/汇总行文本卷入乘法，且 WPS 对 SUMPRODUCT 内
            // IFERROR 的数组语义支持不完整拦不住传染（#VALUE!）；②更要命的是**维度必须一致**——
            // 条件数组用整列（1048576 行）而因子用区间（2693 行）会直接返回 #N/A。
            // 因此现算分支的条件数组与金额因子必须同区间，且不得残留任何整列引用（detects ":$"）。
            Assert("未实现损益现算：SUMPRODUCT 条件数组与因子同区间、无整列引用",
                revenueNoProfit2024 != null && revenueNoProfit2024.StartsWith("=SUMPRODUCT(")
                && revenueNoProfit2024.Contains("IFERROR(")
                && revenueNoProfit2024.Contains("$A$3:") && revenueNoProfit2024.Contains("$J$3:")
                && revenueNoProfit2024.Contains("$E$3:") && revenueNoProfit2024.Contains("$N$3:")
                && revenueNoProfit2024.Contains("$O$3:") && revenueNoProfit2024.Contains("$P$3:")
                && revenueNoProfit2024.Contains("$V$3:")
                // 整列引用形如 $A:$A；区间引用 $A$3:$A$1048576 不匹配（冒号前是数字）
                && !Regex.IsMatch(revenueNoProfit2024, @"\$[A-Z]{1,3}:\$[A-Z]{1,3}"),
                revenueNoProfit2024 ?? "null");
            // 区间须到第 3 行起、工作表末尾止：第 1/2 行的汇总行与表头文本会 #VALUE!，
            // 末行必须是工作表最大行才能随源表向下扩展自动纳入新增行。
            Assert("现算区间 = 数据首行 → 工作表末尾（新增行自动纳入）",
                revenueNoProfit2024 != null && revenueNoProfit2024.Contains("$A$3:$A$1048576")
                && revenueNoProfit2024.Contains("$O$3:$O$1048576")
                && revenueNoProfit2024.Contains("$V$3:$V$1048576"),
                revenueNoProfit2024 ?? "null");
            MatrixOptions limited = MatrixOpt(true);
            limited.LimitComputedRange = true;
            EliminationMatrix limitedMatrix = EliminationMatrixBuilder.Build(matrixRows, MatrixMap(false), limited);
            string limitedFormula = MatrixCell(limitedMatrix, "销售主体A", "客户A", "Unrealized", "2024").Formula;
            Assert("限定当前数据区时退回 $O$3:$O$200（性能优先）",
                limitedFormula != null && limitedFormula.Contains("$O$3:$O$200"), limitedFormula ?? "null");

            Assert("无损益列：组1·2025 U 改由现算 = 0（V=0%）",
                Near(MatrixCell(noProfit, "销售主体A", "客户A", "Unrealized", "2025").Value, 0m),
                MatrixCell(noProfit, "销售主体A", "客户A", "Unrealized", "2025").Value.ToString());
            Assert("无损益列：组1·2024 U 仍为 73714.04",
                Near(MatrixCell(noProfit, "销售主体A", "客户A", "Unrealized", "2024").Value, 73714.04m), "");
            Assert("无损益列仍然平衡", noProfit.Balanced, "");

            // ===== 11. 矩阵写入计划 =====
            Console.WriteLine("--- 11. BuildMatrixPlan ---");
            MergeWritePlan matrixPlan = EliminationGateway.BuildMatrixPlan(matrix, MatrixOpt(true));
            Assert("1 个写入块", matrixPlan.Blocks.Count == 1, "实际 " + matrixPlan.Blocks.Count);
            MergeWriteBlock mb = matrixPlan.Blocks.Count > 0 ? matrixPlan.Blocks[0] : null;
            // 1 行 SUBTOTAL + 2 行表头 + 8 行数据 = 11 行；5 固定列 + 2 年度×3 = 11 列
            Assert("11 行 × 11 列",
                mb != null && mb.RowCount == 11 && mb.ColumnCount == 11,
                mb == null ? "无块" : mb.RowCount + "×" + mb.ColumnCount);
            if (mb != null)
            {
                object[,] c = mb.Cells;
                Assert("G1 = SUBTOTAL(9,G4:G1048576)",
                    Convert.ToString(c[0, 6]) == "=SUBTOTAL(9,G4:G1048576)", Convert.ToString(c[0, 6]));
                Assert("分隔列 F1 留空", c[0, 5] == null, Convert.ToString(c[0, 5]));
                Assert("G2 = 2025 / J2 = 2024",
                    Convert.ToString(c[1, 6]) == "2025" && Convert.ToString(c[1, 9]) == "2024",
                    Convert.ToString(c[1, 6]) + "/" + Convert.ToString(c[1, 9]));
                Assert("G3 = 借方 / H3 = 贷方",
                    Convert.ToString(c[2, 6]) == "借方" && Convert.ToString(c[2, 7]) == "贷方", "");
                Assert("A2 = 核算主体 且 A3 留空（纵向合并）",
                    Convert.ToString(c[1, 0]) == "核算主体" && c[2, 0] == null, Convert.ToString(c[2, 0]));
                Assert("数据首行：主体/客商/备注/调整事项/科目",
                    Convert.ToString(c[3, 0]) == "销售主体A" && Convert.ToString(c[3, 1]) == "客户A"
                    && Convert.ToString(c[3, 2]) == "销售主体A → 客户A：收入抵消"
                    && Convert.ToString(c[3, 3]) == "收入成本抵消"
                    && Convert.ToString(c[3, 4]) == "营业收入", Convert.ToString(c[3, 2]));
                Assert("收入行写借列 G（年度 2025）",
                    Convert.ToString(c[3, 6]).StartsWith("=SUMIFS(内部交易明细!$O:$O"), Convert.ToString(c[3, 6]));
                Assert("成本行写贷列 H",
                    Convert.ToString(c[4, 7]).StartsWith("=SUMIFS(内部交易明细!$P:$P"), Convert.ToString(c[4, 7]));
                Assert("成本还原行 =(G4-H5)-H7",
                    Convert.ToString(c[5, 7]) == "=(G4-H5)-H7", Convert.ToString(c[5, 7]));
                Assert("存货行（有损益列）走 SUMIFS(W)",
                    Convert.ToString(c[6, 7]).StartsWith("=SUMIFS(内部交易明细!$W:$W"), Convert.ToString(c[6, 7]));
                Assert("第二组从第 8 行起且 2025 列留空",
                    Convert.ToString(c[7, 0]) == "销售主体B" && c[7, 6] == null, Convert.ToString(c[7, 6]));
                Assert("第二组 2024 收入写在 J 列",
                    Convert.ToString(c[7, 9]).StartsWith("=SUMIFS(内部交易明细!$O:$O"), Convert.ToString(c[7, 9]));
            }
            Assert("计划总量 121 格", matrixPlan.TotalCells == 121, "实际 " + matrixPlan.TotalCells);

            // ---- 写入空行（GroupGap）：每个交易对之间空 N 行 ----
            // gap=1 时：组1 行 4-7、空行 8、组2 行 9-12（与用户手动插入第 8 行的效果一致）
            MatrixOptions gapOpt = MatrixOpt(true);
            gapOpt.GroupGap = 1;
            EliminationMatrix gapMatrix = EliminationMatrixBuilder.Build(matrixRows, matrixMap, gapOpt);
            Assert("空行：组间距 = 模板高 + gap，第二组从第 9 行起",
                gapMatrix.Rows.Count == 8 && gapMatrix.GroupStride == 5
                && gapMatrix.Rows[4].RowNumber == 9,
                "行数 " + gapMatrix.Rows.Count + " 跨行 " + gapMatrix.GroupStride
                + " 第二组首行 " + (gapMatrix.Rows.Count > 4 ? gapMatrix.Rows[4].RowNumber.ToString() : "无"));
            Assert("空行：组内引用随行号平移 =(J9-K10)-K12",
                MatrixCell(gapMatrix, "销售主体B", "客户B", "Restore", "2024").Formula == "=(J9-K10)-K12",
                MatrixCell(gapMatrix, "销售主体B", "客户B", "Restore", "2024").Formula ?? "null");
            Assert("空行：仍满足借贷恒等式", gapMatrix.Balanced, "");
            MergeWritePlan gapPlan = EliminationGateway.BuildMatrixPlan(gapMatrix, gapOpt);
            MergeWriteBlock gb = gapPlan.Blocks.Count > 0 ? gapPlan.Blocks[0] : null;
            Assert("空行：块高 12（8 数据行 + 1 空行）",
                gb != null && gb.RowCount == 12, gb == null ? "无块" : gb.RowCount.ToString());
            if (gb != null)
            {
                Assert("空行：第 8 行（索引 7）整行留空", gb.Cells[7, 0] == null && gb.Cells[7, 10] == null,
                    Convert.ToString(gb.Cells[7, 0]));
                Assert("空行：索引 8 对应工作第 9 行 = 销售主体B",
                    Convert.ToString(gb.Cells[8, 0]) == "销售主体B", Convert.ToString(gb.Cells[8, 0]));
                Assert("空行：组内引用写入 =(J9-K10)-K12",
                    Convert.ToString(gb.Cells[10, 10]) == "=(J9-K10)-K12", Convert.ToString(gb.Cells[10, 10]));
            }
            Assert("空行：Guard 仍通过", MergePlanner.Guard(gapPlan) == null, MergePlanner.Guard(gapPlan) ?? "");
            Assert("含公式格", matrixPlan.FormulaCells > 0, "实际 " + matrixPlan.FormulaCells);
            Assert("Guard 通过", MergePlanner.Guard(matrixPlan) == null, MergePlanner.Guard(matrixPlan) ?? "");
            Assert("空矩阵 → 空计划",
                EliminationGateway.BuildMatrixPlan(new EliminationMatrix(), MatrixOpt(true)).Blocks.Count == 0, "");

            // 偏移到 B5 起始：起始位置由 options 决定，重新构建后所有锚点须同步平移
            MatrixOptions shiftedOpt = MatrixOpt(true);
            shiftedOpt.StartColumn = 2;
            shiftedOpt.StartRow = 5;
            EliminationMatrix shiftedMatrix = EliminationMatrixBuilder.Build(matrixRows, matrixMap, shiftedOpt);
            MergeWritePlan shifted = EliminationGateway.BuildMatrixPlan(shiftedMatrix, shiftedOpt);
            MergeWriteBlock sb2 = shifted.Blocks.Count > 0 ? shifted.Blocks[0] : null;
            if (sb2 != null)
            {
                Assert("偏移后 SUBTOTAL 从第 8 行起",
                    Convert.ToString(sb2.Cells[0, 6]).Contains("H8:H1048576"), Convert.ToString(sb2.Cells[0, 6]));
                Assert("偏移后成本还原 =(H8-I9)-I11",
                    Convert.ToString(sb2.Cells[5, 7]) == "=(H8-I9)-I11", Convert.ToString(sb2.Cells[5, 7]));
            }

            // ===== 12. 往来抵消：差异表 + 结果表 =====
            Console.WriteLine("--- 12. 往来抵消（净额法、以小抵大）---");
            string entityA = "主体A有限公司";
            string entityB = "主体B有限公司";
            List<EliminationRow> wlRows = new List<EliminationRow>();
            // 主体A账上：其他应收款—对方=主体B 65,950,000
            wlRows.Add(Row(4, EliminationRole.Entity, entityA, EliminationRole.Counterparty, entityB,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "65950000", EliminationRole.Condition, "是"));
            // 主体B账上：其他应收款—对方=主体A 50,000
            wlRows.Add(Row(5, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "50000", EliminationRole.Condition, "是"));
            // 主体B账上：其他应付款—对方=主体A 66,000,000
            wlRows.Add(Row(6, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应付款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "66000000", EliminationRole.Condition, "是"));
            // 干扰：非合并内关联方，须被筛掉
            wlRows.Add(Row(7, EliminationRole.Entity, entityA, EliminationRole.Counterparty, "外部客户",
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "999", EliminationRole.Condition, "否"));

            MatrixOptions diffOpt = DiffOpt();
            EliminationDiffTable diff = null;
            EliminationMatrix wlResult = null;
            EliminationDiffBuilder.Build(wlRows, DiffMap(), diffOpt, out diff, out wlResult);

            Assert("差异表：1 对主体 × 1 年度 × 3 科目对 = 3 行",
                diff != null && diff.Rows.Count == 3 && diff.GroupCount == 1 && diff.Years.Count == 1,
                diff == null ? "null" : diff.Rows.Count + " 行/" + diff.GroupCount + " 对/" + diff.Years.Count + " 年");
            Assert("差异表：条件列筛掉 1 行", diff != null && diff.FilteredRows == 1, "实际 " + diff.FilteredRows);
            Assert("差异表：年度从截止日前 4 位取 2026",
                diff != null && diff.Years.Count == 1 && diff.Years[0] == "2026", "");

            EliminationDiffRow other = diff == null ? null : FindDiff(diff, "其他应收款");
            // 净额法：netA = 主体A应收(65,950,000) − 应付(0) = 65,950,000；
            //         netB = 主体B应收(50,000) − 应付(66,000,000) = −65,950,000
            Assert("资产端账面净额 = 65,950,000（净额法：资产−负债）",
                other != null && Near(other.Asset.Value, 65950000m),
                other == null ? "无行" : other.Asset.Value.ToString());
            Assert("负债端账面净额 = 65,950,000（净额法：负债−资产）",
                other != null && Near(other.Liability.Value, 65950000m),
                other == null ? "无行" : other.Liability.Value.ToString());
            Assert("抵消金额 = MIN = 65,950,000",
                other != null && Near(other.Offset.Value, 65950000m),
                other == null ? "无行" : other.Offset.Value.ToString());
            // 净额计算过程四列：债权方资产/负债余额、债务方资产/负债余额
            Assert("债权方资产余额 = 65,950,000（主体A对其他应收款）",
                other != null && Near(other.CreditorAsset.Value, 65950000m),
                other == null ? "无行" : other.CreditorAsset.Value.ToString());
            Assert("债权方负债余额 = 0（主体A对其他应付款）",
                other != null && Near(other.CreditorLiability.Value, 0m),
                other == null ? "无行" : other.CreditorLiability.Value.ToString());
            Assert("债务方资产余额 = 50,000（主体B对其他应收款）",
                other != null && Near(other.DebtorAsset.Value, 50000m),
                other == null ? "无行" : other.DebtorAsset.Value.ToString());
            Assert("债务方负债余额 = 66,000,000（主体B对其他应付款）",
                other != null && Near(other.DebtorLiability.Value, 66000000m),
                other == null ? "无行" : other.DebtorLiability.Value.ToString());
            // 动态方向：净额较大方（主体A）为债权方、主体B为债务方
            Assert("动态债权方 = 主体A、债务方 = 主体B",
                other != null && other.Creditor == entityA && other.Debtor == entityB,
                other == null ? "无行" : other.Creditor + "/" + other.Debtor);
            // 差额 = 资产端账面净额 − 负债端账面净额（不是 资产−抵消）
            Assert("差额 = 资产 − 负债 = 0（相等情形）",
                other != null && Near(other.Difference.Value, 0m),
                other == null ? "无行" : other.Difference.Value.ToString());
            Assert("备注 = 以小抵大：双方余额相等，全额抵消",
                other != null && other.Note == "以小抵大：双方余额相等，全额抵消",
                other == null ? "无行" : other.Note);
            // 差异表第 1 行是表头，数据自第 2 行起；三个科目对按顺序排列 →「其他应收款」是第 2 行（第 3 行）
            Assert("抵消/差额写本行引用 =MIN(I3,L3) / =I3-L3",
                other != null && other.RowNumber == 3
                && other.Offset.Formula == "=MIN(I3,L3)" && other.Difference.Formula == "=I3-L3",
                other == null ? "无行" : other.RowNumber + " " + other.Offset.Formula + " / " + other.Difference.Formula);
            Assert("净额列引用本行四列做减法：资产端 =G3-H3、负债端 =K3-J3",
                other != null && other.Asset.Formula == "=G3-H3" && other.Liability.Formula == "=K3-J3",
                other == null ? "无行" : (other.Asset.Formula ?? "null") + " / " + (other.Liability.Formula ?? "null"));
            Assert("债权方资产余额走 SUMPRODUCT 精确匹配（方向合计），区间自动扩展到工作表末尾",
                other != null && other.CreditorAsset.Formula != null
                && other.CreditorAsset.Formula.StartsWith("=SUMPRODUCT((")
                && other.CreditorAsset.Formula.Contains("\"" + entityA + "\"") && other.CreditorAsset.Formula.Contains("\"" + entityB + "\"")
                && other.CreditorAsset.Formula.Contains("其他应收款")
                && other.CreditorAsset.Formula.Contains("$U$4:$U$1048576")
                // 括号必须平衡（曾出现提前闭合导致语法错误）
                && Balanced(other.CreditorAsset.Formula),
                other == null ? "无行" : other.CreditorAsset.Formula ?? "null");
            Assert("债权方负债余额方向合计取「其他应付款」",
                other != null && other.CreditorLiability.Formula != null
                && other.CreditorLiability.Formula.Contains("其他应付款")
                && Balanced(other.CreditorLiability.Formula),
                other == null ? "无行" : other.CreditorLiability.Formula ?? "null");

            EliminationDiffRow ar = diff == null ? null : FindDiff(diff, "应收账款");
            Assert("无往来的科目对也列行且填 0",
                ar != null && ar.Asset.Value == 0m && ar.Liability.Value == 0m && ar.Offset.Value == 0m,
                ar == null ? "无行" : ar.Asset.Value.ToString());
            Assert("无往来备注 = 未查询到相关账户，余额为0",
                ar != null && ar.Note == "未查询到相关账户，余额为0", ar == null ? "无行" : ar.Note);

            // 结果表：只输出有金额的科目对 → 2 行（债权方借 / 债务方贷）
            Assert("结果表：仅输出有金额的科目对 → 2 行",
                wlResult != null && wlResult.Rows.Count == 2,
                wlResult == null ? "null" : wlResult.Rows.Count + " 行");
            if (wlResult != null && wlResult.Rows.Count == 2)
            {
                Assert("结果表：第 1 行 债权方借 其他应收款 65,950,000",
                    wlResult.Rows[0].Entity == entityA && wlResult.Rows[0].Account == "其他应收款"
                    && Near(wlResult.Rows[0].Debit.ContainsKey("2026") ? wlResult.Rows[0].Debit["2026"].Value : 0m, 65950000m),
                    wlResult.Rows[0].Entity + "/" + wlResult.Rows[0].Account);
                Assert("结果表：第 2 行 债务方贷 其他应付款 65,950,000",
                    wlResult.Rows[1].Entity == entityB && wlResult.Rows[1].Account == "其他应付款"
                    && Near(wlResult.Rows[1].Credit.ContainsKey("2026") ? wlResult.Rows[1].Credit["2026"].Value : 0m, 65950000m),
                    wlResult.Rows[1].Entity + "/" + wlResult.Rows[1].Account);
                Assert("结果表：调整事项 = 往来抵消", wlResult.Rows[0].Adjustment == "往来抵消", wlResult.Rows[0].Adjustment);
            }

            // ===== 12b. 差额口径与备注分支（资产方大 / 负债方大）=====
            Console.WriteLine("--- 12b. 差额(资产-负债) 与备注建议 ---");
            string m1 = "甲公司", m2 = "乙公司";
            List<EliminationRow> gapRows = new List<EliminationRow>();
            // 甲应收乙 80 万（资产），乙应付甲 30 万（负债）→ 资产方大，差额 50 万
            gapRows.Add(Row(4, EliminationRole.Entity, m1, EliminationRole.Counterparty, m2,
                EliminationRole.Account, "应收账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "800000", EliminationRole.Condition, "是"));
            gapRows.Add(Row(5, EliminationRole.Entity, m2, EliminationRole.Counterparty, m1,
                EliminationRole.Account, "应付账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "300000", EliminationRole.Condition, "是"));
            // 乙预付甲 10 万（资产），甲合同负债乙 25 万（负债）→ 负债方大，差额 15 万
            gapRows.Add(Row(6, EliminationRole.Entity, m2, EliminationRole.Counterparty, m1,
                EliminationRole.Account, "预付款项", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "100000", EliminationRole.Condition, "是"));
            gapRows.Add(Row(7, EliminationRole.Entity, m1, EliminationRole.Counterparty, m2,
                EliminationRole.Account, "预收款项", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "250000", EliminationRole.Condition, "是"));

            EliminationDiffTable gapDiff = null;
            EliminationMatrix gapResult = null;
            EliminationDiffBuilder.Build(gapRows, DiffMap(), DiffOpt(), out gapDiff, out gapResult);
            EliminationDiffRow gAr = gapDiff == null ? null : FindDiff(gapDiff, "应收账款");
            EliminationDiffRow gPre = gapDiff == null ? null : FindDiff(gapDiff, "预付账款");

            Assert("表头第 7 列 = 债权方资产余额",
                EliminationDiffTable.Headers[6] == "债权方资产余额", EliminationDiffTable.Headers[6]);
            Assert("表头第 8 列 = 债权方负债余额",
                EliminationDiffTable.Headers[7] == "债权方负债余额", EliminationDiffTable.Headers[7]);
            Assert("表头第 9 列 = 资产端账面净额(审定数)",
                EliminationDiffTable.Headers[8] == "资产端账面净额(审定数)", EliminationDiffTable.Headers[8]);
            Assert("表头第 10 列 = 债务方资产余额",
                EliminationDiffTable.Headers[9] == "债务方资产余额", EliminationDiffTable.Headers[9]);
            Assert("表头第 11 列 = 债务方负债余额",
                EliminationDiffTable.Headers[10] == "债务方负债余额", EliminationDiffTable.Headers[10]);
            Assert("表头第 12 列 = 负债端账面净额(审定数)",
                EliminationDiffTable.Headers[11] == "负债端账面净额(审定数)", EliminationDiffTable.Headers[11]);
            Assert("表头第 14 列 = 差额(资产-负债)",
                EliminationDiffTable.Headers[13] == "差额(资产-负债)", EliminationDiffTable.Headers[13]);
            Assert("动态方向：应收账款 债权方 = 甲公司、债务方 = 乙公司",
                gAr != null && gAr.Creditor == m1 && gAr.Debtor == m2,
                gAr == null ? "无行" : gAr.Creditor + "/" + gAr.Debtor);
            Assert("动态方向：预付账款 债权方 = 乙公司（净额较大方）、债务方 = 甲公司",
                gPre != null && gPre.Creditor == m2 && gPre.Debtor == m1,
                gPre == null ? "无行" : gPre.Creditor + "/" + gPre.Debtor);
            Assert("资产方大：抵消 = MIN = 300,000",
                gAr != null && Near(gAr.Offset.Value, 300000m), gAr == null ? "无行" : gAr.Offset.Value.ToString());
            Assert("资产方大：差额 = 资产 800,000 − 负债 300,000 = 500,000（非 800,000−300,000 之外的旧口径）",
                gAr != null && Near(gAr.Difference.Value, 500000m), gAr == null ? "无行" : gAr.Difference.Value.ToString());
            Assert("资产方大：备注按方向给建议（核查对方）",
                gAr != null && gAr.Note ==
                "资产方大于负债方，差额500,000.00为对方应收超额，建议核查对方未达账项（单据在途、入账时间差）、对方漏记入账及预收账款未重分类",
                gAr == null ? "无行" : gAr.Note);
            Assert("负债方大：差额 = 100,000 − 250,000 = -150,000",
                gPre != null && Near(gPre.Difference.Value, -150000m), gPre == null ? "无行" : gPre.Difference.Value.ToString());
            Assert("负债方大：备注按方向给建议（核查本单位）",
                gPre != null && gPre.Note ==
                "负债方大于资产方，差额150,000.00为对方应付超额，建议核查本单位未达账项（单据在途、入账时间差）、暂估未入账及预付账款未重分类",
                gPre == null ? "无行" : gPre.Note);
            Assert("差额公式 = 资产端 − 负债端（=I?-L?）",
                gAr != null && gAr.Difference.Formula == "=I" + gAr.RowNumber + "-L" + gAr.RowNumber,
                gAr == null ? "无行" : gAr.Difference.Formula ?? "null");

            // ===== 12b2. 双方净负债（净额为负）：净额保留真实符号、不钳制为 0 =====
            Console.WriteLine("--- 12b2. 双方净负债：净额保留负号、抵消为 0 ---");
            List<EliminationRow> negRows = new List<EliminationRow>();
            // 甲应收乙 100、应付乙 300 → 甲净额 −200；乙应收甲 50、应付甲 400 → 乙净额 −350（双方均净负债）
            negRows.Add(Row(4, EliminationRole.Entity, m1, EliminationRole.Counterparty, m2,
                EliminationRole.Account, "应收账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "100", EliminationRole.Condition, "是"));
            negRows.Add(Row(5, EliminationRole.Entity, m1, EliminationRole.Counterparty, m2,
                EliminationRole.Account, "应付账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "300", EliminationRole.Condition, "是"));
            negRows.Add(Row(6, EliminationRole.Entity, m2, EliminationRole.Counterparty, m1,
                EliminationRole.Account, "应收账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "50", EliminationRole.Condition, "是"));
            negRows.Add(Row(7, EliminationRole.Entity, m2, EliminationRole.Counterparty, m1,
                EliminationRole.Account, "应付账款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "400", EliminationRole.Condition, "是"));

            EliminationDiffTable negDiff = null;
            EliminationMatrix negResult = null;
            EliminationDiffBuilder.Build(negRows, DiffMap(), DiffOpt(), out negDiff, out negResult);
            EliminationDiffRow negAr = negDiff == null ? null : FindDiff(negDiff, "应收账款");
            Assert("净负债：资产端净额保留负号 = −200（不再钳制为 0）",
                negAr != null && Near(negAr.Asset.Value, -200m), negAr == null ? "无行" : negAr.Asset.Value.ToString());
            Assert("净负债：负债端净额 = 债务方负债 − 债务方资产 = 350",
                negAr != null && Near(negAr.Liability.Value, 350m), negAr == null ? "无行" : negAr.Liability.Value.ToString());
            Assert("净负债：抵消 = MAX(0, MIN(资产端,负债端)) = 0",
                negAr != null && Near(negAr.Offset.Value, 0m), negAr == null ? "无行" : negAr.Offset.Value.ToString());
            Assert("净负债：四列映射 债权方资产=100/债权方负债=300/债务方资产=50/债务方负债=400",
                negAr != null && Near(negAr.CreditorAsset.Value, 100m) && Near(negAr.CreditorLiability.Value, 300m)
                && Near(negAr.DebtorAsset.Value, 50m) && Near(negAr.DebtorLiability.Value, 400m),
                negAr == null ? "无行" : negAr.CreditorAsset.Value + "/" + negAr.CreditorLiability.Value
                    + "/" + negAr.DebtorAsset.Value + "/" + negAr.DebtorLiability.Value);

            // ===== 12c. 结果表合并同类行（同标签多年度合并为一行）=====
            Console.WriteLine("--- 12c. 结果表合并同类行 ---");
            List<EliminationRow> mergeRows = new List<EliminationRow>();
            // 其他应收款 2026：主体A应收主体B 65,950,000；主体B应付主体A 66,000,000、应收 50,000
            mergeRows.Add(Row(4, EliminationRole.Entity, entityA, EliminationRole.Counterparty, entityB,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "65950000", EliminationRole.Condition, "是"));
            mergeRows.Add(Row(5, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应付款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "66000000", EliminationRole.Condition, "是"));
            mergeRows.Add(Row(6, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20260630",
                EliminationRole.Amount, "50000", EliminationRole.Condition, "是"));
            // 其他应收款 2025：主体A应收主体B 30,000,000；主体B应付主体A 30,000,000
            mergeRows.Add(Row(7, EliminationRole.Entity, entityA, EliminationRole.Counterparty, entityB,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20251231",
                EliminationRole.Amount, "30000000", EliminationRole.Condition, "是"));
            mergeRows.Add(Row(8, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应付款", EliminationRole.Deadline, "20251231",
                EliminationRole.Amount, "30000000", EliminationRole.Condition, "是"));

            EliminationDiffTable mergeDiff = null;
            EliminationMatrix mergeResult = null;
            EliminationDiffBuilder.Build(mergeRows, DiffMap(), DiffOpt(), out mergeDiff, out mergeResult);

            Assert("合并：2 年度 × 1 科目对 → 结果表合并为 2 行（不再是 4 行）",
                mergeResult != null && mergeResult.Rows.Count == 2,
                mergeResult == null ? "null" : mergeResult.Rows.Count + " 行");
            if (mergeResult != null && mergeResult.Rows.Count == 2)
            {
                bool creditOk = mergeResult.Rows[0].Entity == entityA
                    && mergeResult.Rows[0].Account == "其他应收款"
                    && mergeResult.Rows[0].Debit.ContainsKey("2026") && mergeResult.Rows[0].Debit.ContainsKey("2025")
                    && Near(mergeResult.Rows[0].Debit["2026"].Value, 65950000m)
                    && Near(mergeResult.Rows[0].Debit["2025"].Value, 30000000m);
                Assert("合并：债权方借其他应收款 2026=65,950,000、2025=30,000,000 同在一行",
                    creditOk, mergeResult.Rows[0].Entity + "/" + mergeResult.Rows[0].Account);
                bool debtOk = mergeResult.Rows[1].Entity == entityB
                    && mergeResult.Rows[1].Account == "其他应付款"
                    && mergeResult.Rows[1].Credit.ContainsKey("2026") && mergeResult.Rows[1].Credit.ContainsKey("2025")
                    && Near(mergeResult.Rows[1].Credit["2026"].Value, 65950000m)
                    && Near(mergeResult.Rows[1].Credit["2025"].Value, 30000000m);
                Assert("合并：债务方贷其他应付款 2026=65,950,000、2025=30,000,000 同在一行",
                    debtOk, mergeResult.Rows[1].Entity + "/" + mergeResult.Rows[1].Account);
            }

            // ===== 12d. 坏账准备抵消（借坏账准备 / 贷管理费用，全额、联动）=====
            Console.WriteLine("--- 12d. 坏账准备抵消 ---");
            List<EliminationRow> bdRows = new List<EliminationRow>();
            // 主体A→主体B 其他应收款 2025：审定数 30000000、坏账准备 150000
            bdRows.Add(Row(4, EliminationRole.Entity, entityA, EliminationRole.Counterparty, entityB,
                EliminationRole.Account, "其他应收款", EliminationRole.Deadline, "20251231",
                EliminationRole.Amount, "30000000", EliminationRole.BadDebt, "150000", EliminationRole.Condition, "是"));
            bdRows.Add(Row(5, EliminationRole.Entity, entityB, EliminationRole.Counterparty, entityA,
                EliminationRole.Account, "其他应付款", EliminationRole.Deadline, "20251231",
                EliminationRole.Amount, "30000000", EliminationRole.Condition, "是"));
            // 干扰：非合并内，坏账准备不参与
            bdRows.Add(Row(6, EliminationRole.Entity, entityA, EliminationRole.Counterparty, "外部客户",
                EliminationRole.Account, "应收账款", EliminationRole.Deadline, "20251231",
                EliminationRole.Amount, "999", EliminationRole.BadDebt, "888", EliminationRole.Condition, "否"));

            EliminationDiffTable bdDiff = null;
            EliminationMatrix bdResult = null;
            EliminationDiffBuilder.Build(bdRows, DiffMapWithBadDebt(), DiffOpt(), out bdDiff, out bdResult);

            Assert("坏账准备抵消：结果表 4 行（2 往来抵消 + 2 坏账准备抵消）",
                bdResult != null && bdResult.Rows.Count == 4,
                bdResult == null ? "null" : bdResult.Rows.Count + " 行");
            if (bdResult != null && bdResult.Rows.Count == 4)
            {
                MatrixRow bdCredit = bdResult.Rows[0];
                MatrixRow bdBad = bdResult.Rows[1];
                MatrixRow bdMgmt = bdResult.Rows[2];
                MatrixRow bdDebt = bdResult.Rows[3];
                Assert("坏账准备抵消：第 1 行 债权方借其他应收款",
                    bdCredit.Entity == entityA && bdCredit.Account == "其他应收款" && bdCredit.Adjustment == "往来抵消",
                    bdCredit.Entity + "/" + bdCredit.Account);
                Assert("坏账准备抵消：第 2 行 借坏账准备（紧跟债权方行、金额全额、D列区分）",
                    bdBad.Entity == entityA && bdBad.Customer == entityB && bdBad.Account == "坏账准备"
                    && bdBad.Adjustment == "坏账准备抵消"
                    && bdBad.Debit.ContainsKey("2025") && Near(bdBad.Debit["2025"].Value, 150000m),
                    bdBad.Entity + "/" + bdBad.Account + "/" + bdBad.Adjustment);
                Assert("坏账准备抵消：第 3 行 贷管理费用（B列=—、金额全额）",
                    bdMgmt.Entity == entityA && bdMgmt.Customer == "—" && bdMgmt.Account == "管理费用"
                    && bdMgmt.Adjustment == "坏账准备抵消"
                    && bdMgmt.Credit.ContainsKey("2025") && Near(bdMgmt.Credit["2025"].Value, 150000m),
                    bdMgmt.Entity + "/" + bdMgmt.Customer + "/" + bdMgmt.Account);
                Assert("坏账准备抵消：第 4 行 债务方贷其他应付款",
                    bdDebt.Entity == entityB && bdDebt.Account == "其他应付款" && bdDebt.Adjustment == "往来抵消",
                    bdDebt.Entity + "/" + bdDebt.Account);
            }

            // ===== 12e. 角色自动匹配收紧 + 列映射临时记忆 =====
            Console.WriteLine("--- 12e. GuessRole 收紧与记忆配置 ---");
            Assert("收入成本：单体收入 → 营业收入",
                EliminationEngine.GuessRole("单体收入", EliminationKind.RevenueCost) == EliminationRole.Revenue, "");
            Assert("收入成本：单体收入（未审）→ 营业收入",
                EliminationEngine.GuessRole("单体收入（未审）", EliminationKind.RevenueCost) == EliminationRole.Revenue, "");
            Assert("收入成本：收入调整 不匹配营业收入",
                EliminationEngine.GuessRole("收入调整", EliminationKind.RevenueCost) != EliminationRole.Revenue, "");
            Assert("收入成本：单体成本 → 营业成本",
                EliminationEngine.GuessRole("单体成本", EliminationKind.RevenueCost) == EliminationRole.Cost, "");
            Assert("收入成本：成本调整 不匹配营业成本",
                EliminationEngine.GuessRole("成本调整", EliminationKind.RevenueCost) != EliminationRole.Cost, "");
            Assert("收入成本：是否为合并内关联方 → 条件列",
                EliminationEngine.GuessRole("是否为合并内关联方", EliminationKind.RevenueCost) == EliminationRole.Condition, "");
            Assert("收入成本：是否为合并外关联方 不匹配条件列",
                EliminationEngine.GuessRole("是否为合并外关联方", EliminationKind.RevenueCost) != EliminationRole.Condition, "");
            Assert("往来抵消：是否为合并内关联方 → 条件列",
                EliminationEngine.GuessRole("是否为合并内关联方", EliminationKind.Receivable) == EliminationRole.Condition, "");
            Assert("往来抵消：是否为合并外关联方 不匹配条件列",
                EliminationEngine.GuessRole("是否为合并外关联方", EliminationKind.Receivable) != EliminationRole.Condition, "");

            List<EliminationColumn> memCols = new List<EliminationColumn>();
            memCols.Add(new EliminationColumn { Letter = "B", Role = EliminationRole.Entity });
            memCols.Add(new EliminationColumn { Letter = "U", Role = EliminationRole.Amount });
            memCols.Add(new EliminationColumn { Letter = "Z", Role = EliminationRole.Ignore });
            string memPath = @"C:\测试\记忆.xlsx";
            string memKey = EliminationForm.EliminationMapStore.Key(EliminationKind.Receivable, memPath);
            EliminationForm.EliminationMapStore.Save(memKey, memCols);
            Dictionary<string, string> memLoaded = EliminationForm.EliminationMapStore.Load(memKey);
            Assert("记忆配置：Save 后 Load 命中且角色一致",
                memLoaded != null && memLoaded.ContainsKey("B") && memLoaded["B"] == "核算主体"
                && memLoaded.ContainsKey("U") && memLoaded["U"] == "审定数",
                memLoaded == null ? "null" : (memLoaded.ContainsKey("B") ? memLoaded["B"] : "-") + "/" + (memLoaded.ContainsKey("U") ? memLoaded["U"] : "-"));
            Assert("记忆配置：忽略角色也被记忆并回填",
                memLoaded != null && memLoaded.ContainsKey("Z") && memLoaded["Z"] == "忽略",
                memLoaded == null ? "null" : (memLoaded.ContainsKey("Z") ? memLoaded["Z"] : "-"));
            EliminationForm.EliminationMapStore.ClearFile(memPath);
            Assert("记忆配置：ClearFile 后 Load 为空",
                EliminationForm.EliminationMapStore.Load(memKey) == null, "");

            // ===== 13. 写入计划：组内行数不均（往来抵消场景）=====
            Console.WriteLine("--- 13. BuildMatrixPlan 组内行数不均 ---");
            // 模拟往来抵消结果表：组1 有 2 行（1 个有金额的科目对）、组2 也 2 行，
            // 但组间空 1 行——旧实现按「每 TemplateCount(=2) 行加一个 gap」推算，
            // 与构建器的 RowNumber 错位，会越界（实机验收：执行写入报「索引超出了数组界限」）。
            EliminationMatrix uneven = new EliminationMatrix();
            uneven.Years = new List<string>(new string[] { "2026" });
            uneven.Columns = new List<MatrixColumn>();
            for (int i = 0; i < 5; i++)
            {
                MatrixColumn col = new MatrixColumn();
                col.Index = i; col.Number = i + 1; col.Letter = EliminationMatrixBuilder.Letter(i + 1);
                col.Head2 = new string[] { "核算主体", "明细【客商/二级明细】", "备注", "调整事项", "会计科目" }[i];
                uneven.Columns.Add(col);
            }
            MatrixColumn spacer = new MatrixColumn { Index = 5, Number = 6, IsSpacer = true };
            uneven.Columns.Add(spacer);
            MatrixColumn colDebit = new MatrixColumn { Index = 6, Number = 7, Year = "2026", IsDebit = true, Head2 = "2026", Head3 = "借方" };
            colDebit.Letter = EliminationMatrixBuilder.Letter(7);
            uneven.Columns.Add(colDebit);
            MatrixColumn colCredit = new MatrixColumn { Index = 7, Number = 8, Year = "2026", IsDebit = false, Head2 = "2026", Head3 = "贷方" };
            colCredit.Letter = EliminationMatrixBuilder.Letter(8);
            uneven.Columns.Add(colCredit);

            MatrixRow u1 = new MatrixRow { RowNumber = 4, Entity = "甲公司", Customer = "乙公司", Account = "其他应收款", Adjustment = "往来抵消" };
            u1.Debit["2026"] = new MatrixAmount { Value = 100m };
            MatrixRow u2 = new MatrixRow { RowNumber = 5, Entity = "乙公司", Customer = "甲公司", Account = "其他应付款", Adjustment = "往来抵消" };
            u2.Credit["2026"] = new MatrixAmount { Value = 100m };
            MatrixRow u3 = new MatrixRow { RowNumber = 7, Entity = "丙公司", Customer = "丁公司", Account = "其他应收款", Adjustment = "往来抵消" };
            u3.Debit["2026"] = new MatrixAmount { Value = 200m };
            MatrixRow u4 = new MatrixRow { RowNumber = 8, Entity = "丁公司", Customer = "丙公司", Account = "其他应付款", Adjustment = "往来抵消" };
            u4.Credit["2026"] = new MatrixAmount { Value = 200m };
            uneven.Rows.Add(u1); uneven.Rows.Add(u2); uneven.Rows.Add(u3); uneven.Rows.Add(u4);
            uneven.GroupCount = 2;
            uneven.TemplateCount = 2;
            uneven.GroupGap = 1;
            uneven.GroupStride = 3;

            MatrixOptions unevenOpt = new MatrixOptions
            {
                SourceSheetName = "往来明细表",
                HasConditionColumn = true,
                TotalRow = true,
                StaticValues = false,
                SourceFirstRow = 4,
                SourceLastRow = 100,
                StartRow = 1,
                StartColumn = 1,
                GroupGap = 1
            };
            MergeWritePlan unevenPlan = EliminationGateway.BuildMatrixPlan(uneven, unevenOpt);
            MergeWriteBlock ub = unevenPlan.Blocks.Count > 0 ? unevenPlan.Blocks[0] : null;
            // dataTop = 1+1+2 = 4；最后一行 RowNumber=8 → 数据 5 行 → 块高 = 1+2+5 = 8
            Assert("组内行数不均：块高 8（不再越界）",
                ub != null && ub.RowCount == 8, ub == null ? "无块" : ub.RowCount.ToString());
            if (ub != null)
            {
                Assert("组内行数不均：组1 首行在块内偏移 3（工作第 4 行）",
                    Convert.ToString(ub.Cells[3, 0]) == "甲公司", Convert.ToString(ub.Cells[3, 0]));
                Assert("组内行数不均：空行（块内偏移 5，工作第 6 行）留空",
                    ub.Cells[5, 0] == null, Convert.ToString(ub.Cells[5, 0]));
                Assert("组内行数不均：组2 首行在块内偏移 6（工作第 7 行）",
                    Convert.ToString(ub.Cells[6, 0]) == "丙公司", Convert.ToString(ub.Cells[6, 0]));
                Assert("组内行数不均：Guard 通过", MergePlanner.Guard(unevenPlan) == null,
                    MergePlanner.Guard(unevenPlan) ?? "");
            }

            Console.WriteLine();
            Console.WriteLine("============================================");
            Console.WriteLine("=== 结果：" + _pass + " PASS / " + _fail + " FAIL ===");
            Console.WriteLine("============================================");
        }

        // ---- 往来抵消用辅助 ----

        static MatrixOptions DiffOpt()
        {
            return new MatrixOptions
            {
                SourceSheetName = "往来明细表",
                SourceBookName = "",
                ConditionValue = "是",
                HasConditionColumn = true,
                TotalRow = true,
                StaticValues = false,
                FullColumnRefs = true,
                LimitComputedRange = false,
                SourceFirstRow = 4,
                SourceLastRow = 2864,
                StartRow = 1,
                StartColumn = 1
            };
        }

        static List<EliminationColumn> DiffMap()
        {
            List<EliminationColumn> list = new List<EliminationColumn>();
            list.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "截止日", Role = EliminationRole.Deadline });
            list.Add(new EliminationColumn { Number = 2, Letter = "B", Header = "核算主体", Role = EliminationRole.Entity });
            list.Add(new EliminationColumn { Number = 3, Letter = "C", Header = "科目", Role = EliminationRole.Account });
            list.Add(new EliminationColumn { Number = 4, Letter = "D", Header = "对方单位", Role = EliminationRole.Counterparty });
            list.Add(new EliminationColumn { Number = 6, Letter = "F", Header = "是否合并内关联方", Role = EliminationRole.Condition });
            list.Add(new EliminationColumn { Number = 21, Letter = "U", Header = "审定数", Role = EliminationRole.Amount });
            return list;
        }

        static List<EliminationColumn> DiffMapWithBadDebt()
        {
            List<EliminationColumn> list = DiffMap();
            list.Add(new EliminationColumn { Number = 39, Letter = "AM", Header = "坏账准备", Role = EliminationRole.BadDebt });
            return list;
        }

        /// <summary>公式括号是否配对（防止生成语法错误的公式）。</summary>
        static bool Balanced(string formula)
        {
            if (formula == null) return false;
            int depth = 0;
            bool inQuote = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char c = formula[i];
                if (c == '"') { inQuote = !inQuote; continue; }
                if (inQuote) continue;
                if (c == '(') depth++;
                else if (c == ')') { depth--; if (depth < 0) return false; }
            }
            return depth == 0;
        }

        static EliminationDiffRow FindDiff(EliminationDiffTable diff, string assetAccount)
        {
            foreach (EliminationDiffRow row in diff.Rows)
                if (row.AssetAccount == assetAccount) return row;
            return null;
        }

        // ---- 矩阵式底稿用辅助 ----

        static MatrixOptions MatrixOpt(bool fullColumn)
        {
            return new MatrixOptions
            {
                SourceSheetName = "内部交易明细",
                SourceBookName = "",
                ConditionValue = "是",
                HasConditionColumn = true,
                TotalRow = true,
                StaticValues = false,
                FullColumnRefs = fullColumn,
                SourceFirstRow = 3,
                SourceLastRow = 200,
                StartRow = 1,
                StartColumn = 1
            };
        }

        static List<EliminationColumn> MatrixMap(bool withProfitColumn)
        {
            List<EliminationColumn> list = new List<EliminationColumn>();
            list.Add(new EliminationColumn { Number = 1, Letter = "A", Header = "核算主体", Role = EliminationRole.Entity });
            list.Add(new EliminationColumn { Number = 5, Letter = "E", Header = "年份", Role = EliminationRole.Year });
            list.Add(new EliminationColumn { Number = 10, Letter = "J", Header = "客户名称", Role = EliminationRole.Counterparty });
            list.Add(new EliminationColumn { Number = 14, Letter = "N", Header = "是否为合并内关联方", Role = EliminationRole.Condition });
            list.Add(new EliminationColumn { Number = 15, Letter = "O", Header = "单体收入（未审）", Role = EliminationRole.Revenue });
            list.Add(new EliminationColumn { Number = 16, Letter = "P", Header = "单体成本（未审）", Role = EliminationRole.Cost });
            list.Add(new EliminationColumn { Number = 22, Letter = "V", Header = "未对外销售比例", Role = EliminationRole.UnsoldRatio });
            if (withProfitColumn)
                list.Add(new EliminationColumn { Number = 23, Letter = "W", Header = "未实现内部交易损益（存货）", Role = EliminationRole.UnrealizedProfit });
            return list;
        }

        static bool Near(decimal actual, decimal expect)
        { return Math.Abs(actual - expect) < 0.005m; }

        /// <summary>按（主体, 客商, 模板行 Key, 年度）取出金额格。</summary>
        static MatrixAmount MatrixCell(EliminationMatrix matrix, string entity, string customer,
            string templateKey, string year)
        {
            foreach (MatrixRow row in matrix.Rows)
            {
                if (row.Entity != entity || row.Customer != customer || row.TemplateKey != templateKey) continue;
                MatrixAmount amount;
                if (row.Debit.TryGetValue(year, out amount)) return amount;
                if (row.Credit.TryGetValue(year, out amount)) return amount;
            }
            return new MatrixAmount();
        }
    }
}
