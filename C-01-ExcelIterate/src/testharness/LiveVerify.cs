using System;
using System.Collections.Generic;
using System.Globalization;
using Excel_Iterate;

// Standalone WPS regression probe. Samples and source sheets are created in an
// owned workbook; no personal path, user workbook save or Application.Quit.
internal static class LiveVerify
{
    static int pass = 0, fail = 0;

    static void Check(string name, bool ok, string detail)
    {
        if (ok) { pass++; Console.WriteLine("  [PASS] " + name); }
        else { fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
    }

    static EliminationRow R(int row, params object[] kv)
    {
        EliminationRow r = new EliminationRow();
        r.SourceRow = row;
        for (int i = 0; i + 1 < kv.Length; i += 2)
            r.Values[(EliminationRole)kv[i]] = Convert.ToString(kv[i + 1], CultureInfo.InvariantCulture);
        return r;
    }

    static EliminationColumn C(int number, string letter, string header, EliminationRole role)
    { return new EliminationColumn { Number = number, Letter = letter, Header = header, Role = role }; }

    private sealed class Fixture
    {
        private readonly ProbeSession session;
        private readonly ProbeReferences sheets;
        private readonly object book;
        internal Fixture(ProbeSession owner, ProbeReferences references, object workbook)
        { session = owner; sheets = references; book = workbook; }
        internal object Sheet(string name)
        {
            object collection = sheets.Take((object)((dynamic)book).Worksheets);
            object sheet = sheets.Take((object)((dynamic)collection).Add());
            ((dynamic)sheet).Name = name; return sheet;
        }
        internal void Put(object sheet, string address, object value, bool expression = false)
        {
            using (var refs = session.References())
            {
                object range = refs.Take((object)((dynamic)sheet).Range[address]);
                if (expression && value is string) ((dynamic)range).Formula = (string)value;
                else ((dynamic)range).Value2 = value;
            }
        }
        internal object Read(object sheet, string address, bool display = false)
        {
            using (var refs = session.References())
            {
                object range = refs.Take((object)((dynamic)sheet).Range[address]);
                return display ? (object)((dynamic)range).Text : (object)((dynamic)range).Value2;
            }
        }
        internal void Seed(string name, IList<EliminationRow> rows, IList<EliminationColumn> columns)
        {
            object sheet = Sheet(name);
            foreach (var row in rows) foreach (var column in columns)
            {
                string value; if (!row.Values.TryGetValue(column.Role, out value)) continue;
                double number; object payload = value;
                if (double.TryParse(value.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                    payload = value.EndsWith("%", StringComparison.Ordinal) ? number / 100 : number;
                Put(sheet, column.Letter + row.SourceRow, payload);
            }
        }
    }
    [STAThread] static int Main(string[] args)
    {
        using (var session = new ProbeSession("elimination-live"))
        {
            try
            {
                string[] ids = Array.IndexOf(args, "--standalone") >= 0 ? new[] { "Excel.Application" } : new[] { "KET.Application", "et.Application", "Excel.Application" };
                object application = session.Take(Array.IndexOf(args, "--new-instance") >= 0
                    ? ProbeSession.Connect(ids, false, id => null, id => { Type type = Type.GetTypeFromProgID(id); return type == null ? null : Activator.CreateInstance(type); }, Console.WriteLine)
                    : ProbeSession.Connect(ids, false));
                if (application == null) return session.Finish(2);
                dynamic app = application;
                object previousSheet = null, previousSelection = null;
                try { previousSheet = session.Take((object)app.ActiveSheet); previousSelection = session.Take((object)app.Selection); }
                catch (System.Runtime.InteropServices.COMException) { }
                session.OnClose(delegate
                {
                    if (previousSheet != null) ((dynamic)previousSheet).Activate();
                    if (previousSelection != null) ((dynamic)previousSelection).Select();
                });
                session.Override(() => (object)app.DisplayAlerts, value => app.DisplayAlerts = Convert.ToBoolean(value), false, "DisplayAlerts");
                session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = Convert.ToBoolean(value), false, "ScreenUpdating");
                session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = Convert.ToBoolean(value), false, "EnableEvents");
                session.Override(() => (object)app.Calculation, value => app.Calculation = Convert.ToInt32(value), -4135, "Calculation");
                session.Override(() => (object)app.Visible, value => app.Visible = Convert.ToBoolean(value), (object)app.Visible, "Visible (unchanged)");
                using (var refs = session.References())
                {
                    object books = refs.Take((object)app.Workbooks), book = null;
                    book = session.Own((object)((dynamic)books).Add(), delegate { ((dynamic)book).Close(false); });
                    var fixture = new Fixture(session, refs, book);
                    VerifyRevenueCost(fixture); VerifyReceivable(fixture);
                }
                Console.WriteLine("LiveVerify: " + pass + " PASS / " + fail + " FAIL; owned generic samples only");
                return session.Finish(fail == 0 ? 0 : 1);
            }
            catch (Exception error) { Console.WriteLine(error); return session.Finish(1); }
        }
    }

    // ---------------- 收入成本抵消 ----------------
    static void VerifyRevenueCost(Fixture fixture)
    {
        Console.WriteLine();
        Console.WriteLine("--- 收入成本抵消（矩阵式底稿）---");
        List<EliminationColumn> cols = new List<EliminationColumn>();
        cols.Add(C(1, "A", "核算主体", EliminationRole.Entity));
        cols.Add(C(5, "E", "年份", EliminationRole.Year));
        cols.Add(C(10, "J", "客户名称", EliminationRole.Counterparty));
        cols.Add(C(14, "N", "是否为合并内关联方", EliminationRole.Condition));
        cols.Add(C(15, "O", "单体收入（未审）", EliminationRole.Revenue));
        cols.Add(C(16, "P", "单体成本（未审）", EliminationRole.Cost));
        cols.Add(C(22, "V", "未对外销售比例", EliminationRole.UnsoldRatio));
        cols.Add(C(23, "W", "未实现内部交易损益（存货）", EliminationRole.UnrealizedProfit));

        List<EliminationRow> rows = new List<EliminationRow>();
        rows.Add(R(3, EliminationRole.Entity, "销售主体A", EliminationRole.Year, "2024",
            EliminationRole.Counterparty, "客户A", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "2217860.59", EliminationRole.Cost, "1849290.40",
            EliminationRole.UnsoldRatio, "20%"));
        rows.Add(R(4, EliminationRole.Entity, "销售主体A", EliminationRole.Year, "2025",
            EliminationRole.Counterparty, "客户A", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "1998483.64", EliminationRole.Cost, "2258286.52",
            EliminationRole.UnsoldRatio, "20%", EliminationRole.UnrealizedProfit, "-46492.5225"));
        rows.Add(R(5, EliminationRole.Entity, "销售主体B", EliminationRole.Year, "2024",
            EliminationRole.Counterparty, "客户B", EliminationRole.Condition, "是",
            EliminationRole.Revenue, "151503.00", EliminationRole.Cost, "147503.00",
            EliminationRole.UnsoldRatio, "10%"));
        rows.Add(R(6, EliminationRole.Entity, "销售主体C", EliminationRole.Year, "2025",
            EliminationRole.Counterparty, "客户C", EliminationRole.Condition, "否",
            EliminationRole.Revenue, "999999", EliminationRole.Cost, "1",
            EliminationRole.UnsoldRatio, "50%"));

        MatrixOptions opt = new MatrixOptions();
        opt.SourceSheetName = "内部交易明细";
        opt.ConditionValue = "是";
        opt.HasConditionColumn = true;
        opt.TotalRow = false;      // 验证表从第 1 行写表头，简化
        opt.FullColumnRefs = true;
        opt.SourceFirstRow = 3;
        opt.SourceLastRow = 1000;
        opt.StartRow = 1;
        opt.StartColumn = 1;

        fixture.Seed(opt.SourceSheetName, rows, cols);
        EliminationMatrix m = EliminationMatrixBuilder.Build(rows, cols, opt);
        Check("构建：2 组 / 2 年度 / 8 行", m.GroupCount == 2 && m.Years.Count == 2 && m.Rows.Count == 8,
            m.GroupCount + "/" + m.Years.Count + "/" + m.Rows.Count);

        // 写入验证表（手工复刻 BuildMatrixPlan 布局：表头 2 行 + 数据）
        object sheet = fixture.Sheet("验证-收入抵消");
        // 第 1 层表头：年度（公式锚点 G$2 / J$2 用第 2 行，故这里数据从第 4 行起需保持一致）
        // 为与插件一致，仍按 TotalRow=false → 表头在第 1、2 行，数据从第 3 行起
        fixture.Put(sheet, "G1", "2025");
        fixture.Put(sheet, "J1", "2024");
        foreach (MatrixRow row in m.Rows)
        {
            fixture.Put(sheet, "A" + row.RowNumber, row.Entity);
            fixture.Put(sheet, "B" + row.RowNumber, row.Customer);
            foreach (string year in m.Years)
            {
                MatrixColumn d = m.DebitColumn(year), c = m.CreditColumn(year);
                MatrixAmount a;
                if (d != null && row.Debit.TryGetValue(year, out a)) fixture.Put(sheet, d.Letter + row.RowNumber, a.Formula, true);
                if (c != null && row.Credit.TryGetValue(year, out a)) fixture.Put(sheet, c.Letter + row.RowNumber, a.Formula, true);
            }
        }

        ((dynamic)sheet).Calculate();

        Func<string, double> val = address => Number(fixture, sheet, address);
        Func<string, string> txt = address => Text(fixture, sheet, address);

        // 组1 起始行：TotalRow=false → 表头 2 行 → 数据从第 3 行起；年度锚点为 G$1 / J$1
        // 但插件公式里的年度锚点用的是 firstHeaderRow（= StartRow + TotalRow?1:0 = 1），
        // 且数据 dataTop = 1 + 2 = 3 ✓ 与上面写入一致
        Console.WriteLine("  组1首行 RowNumber=" + m.Rows[0].RowNumber);
        int t1 = m.Rows[0].RowNumber;   // 组1 首行
        int t2 = t1 + 4;                 // 组2 首行（gap=0）
        Check("销售主体A·2025 收入 = 1,998,483.64", Math.Abs(val("G" + t1) - 1998483.64) < 0.01,
            txt("G" + t1));
        Check("销售主体A·2025 成本 = 2,258,286.52", Math.Abs(val("H" + (t1 + 1)) - 2258286.52) < 0.01,
            txt("H" + (t1 + 1)));
        Check("销售主体A·2024 收入 = 2,217,860.59", Math.Abs(val("J" + t1) - 2217860.59) < 0.01,
            txt("J" + t1));
        Check("销售主体A·2024 成本 = 1,849,290.40", Math.Abs(val("K" + (t1 + 1)) - 1849290.40) < 0.01,
            txt("K" + (t1 + 1)));
        // 存货行：2025 走 SUMIFS(W) = -46492.5225；2024 走 SUMPRODUCT 现算 = (R-C)*20% = 73714.038
        Check("销售主体A·2025 存货 = -46,492.52（SUMIFS 损益列）",
            Math.Abs(val("H" + (t1 + 3)) - (-46492.5225)) < 0.01, txt("H" + (t1 + 3)));
        Check("销售主体A·2024 存货 = 73,714.04（SUMPRODUCT 现算，关键）",
            Math.Abs(val("K" + (t1 + 3)) - 73714.038) < 0.05, txt("K" + (t1 + 3)));
        // 成本还原 =(G-H)-K
        Check("销售主体A·2024 成本还原 = 294,856.15",
            Math.Abs(val("K" + (t1 + 2)) - 294856.15) < 0.05, txt("K" + (t1 + 2)));
        Check("销售主体B·2024 收入 = 151,503.00", Math.Abs(val("J" + t2) - 151503.00) < 0.01,
            txt("J" + t2));
        Check("销售主体B·2024 存货 = 400.00（(151503-147503)*10%）",
            Math.Abs(val("K" + (t2 + 3)) - 400.0) < 0.01, txt("K" + (t2 + 3)));
    }

    // ---------------- 往来抵消 ----------------
    static void VerifyReceivable(Fixture fixture)
    {
        Console.WriteLine();
        Console.WriteLine("--- 往来抵消（差异表，净额法 精确匹配）---");
        List<EliminationColumn> cols = new List<EliminationColumn>();
        cols.Add(C(1, "A", "截止日", EliminationRole.Deadline));
        cols.Add(C(2, "B", "核算主体", EliminationRole.Entity));
        cols.Add(C(3, "C", "科目", EliminationRole.Account));
        cols.Add(C(4, "D", "对方单位", EliminationRole.Counterparty));
        cols.Add(C(6, "F", "是否合并内关联方", EliminationRole.Condition));
        cols.Add(C(21, "U", "审定数", EliminationRole.Amount));

        string A = "主体A有限公司";
        string B = "主体B有限公司";
        List<EliminationRow> rows = new List<EliminationRow>();
        rows.Add(R(4, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "其他应收款", EliminationRole.Counterparty, B,
            EliminationRole.Condition, "是", EliminationRole.Amount, "65950000"));
        rows.Add(R(5, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "其他应收款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "50000"));
        rows.Add(R(6, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "其他应付款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "66000000"));
        rows.Add(R(7, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "应收账款", EliminationRole.Counterparty, "外部客户甲",
            EliminationRole.Condition, "否", EliminationRole.Amount, "999"));
        rows.Add(R(8, EliminationRole.Deadline, "20260630", EliminationRole.Entity, A,
            EliminationRole.Account, "应收账款", EliminationRole.Counterparty, B,
            EliminationRole.Condition, "是", EliminationRole.Amount, "800000"));
        rows.Add(R(9, EliminationRole.Deadline, "20260630", EliminationRole.Entity, B,
            EliminationRole.Account, "应付账款", EliminationRole.Counterparty, A,
            EliminationRole.Condition, "是", EliminationRole.Amount, "300000"));

        MatrixOptions opt = new MatrixOptions();
        opt.SourceSheetName = "往来明细表";
        opt.ConditionValue = "是";
        opt.HasConditionColumn = true;
        opt.FullColumnRefs = true;
        opt.SourceFirstRow = 4;
        opt.SourceLastRow = 1000;
        opt.StartRow = 1;
        opt.StartColumn = 1;
        // 结果表金额跨表引用差异表，需先指定差异表名（与下方 fixture.Sheet 名一致）
        opt.DiffSheetName = "验证-往来差异";

        EliminationDiffTable diff = null;
        EliminationMatrix result = null;
        fixture.Seed(opt.SourceSheetName, rows, cols);
        EliminationDiffBuilder.Build(rows, cols, opt, out diff, out result);
        Check("构建：1 主体对 / 1 年度 / 3 行",
            diff != null && diff.GroupCount == 1 && diff.Years.Count == 1 && diff.Rows.Count == 3,
            diff == null ? "null" : diff.GroupCount + "/" + diff.Years.Count + "/" + diff.Rows.Count);

        object sheet = fixture.Sheet("验证-往来差异");
        for (int c = 0; c < EliminationDiffTable.Headers.Length; c++)
            fixture.Put(sheet, ((char)('A' + c)).ToString() + "1", EliminationDiffTable.Headers[c]);
        foreach (EliminationDiffRow row in diff.Rows)
        {
            int r = row.RowNumber;   // 表头占第 1 行，数据自第 2 行起
            fixture.Put(sheet, "A" + r, row.Index);
            fixture.Put(sheet, "B" + r, 2026);
            fixture.Put(sheet, "C" + r, row.Creditor);
            fixture.Put(sheet, "D" + r, row.Debtor);
            fixture.Put(sheet, "E" + r, row.AssetAccount);
            fixture.Put(sheet, "F" + r, row.LiabilityAccount);
            // 与插件 BuildDiffPlan 一致：公式为空（科目不存在）时写数值 0
            fixture.Put(sheet, "G" + r, MatrixAmount.Cell(row.CreditorAsset, false), true);
            fixture.Put(sheet, "H" + r, MatrixAmount.Cell(row.CreditorLiability, false), true);
            fixture.Put(sheet, "I" + r, MatrixAmount.Cell(row.Asset, false), true);
            fixture.Put(sheet, "J" + r, MatrixAmount.Cell(row.DebtorAsset, false), true);
            fixture.Put(sheet, "K" + r, MatrixAmount.Cell(row.DebtorLiability, false), true);
            fixture.Put(sheet, "L" + r, MatrixAmount.Cell(row.Liability, false), true);
            fixture.Put(sheet, "M" + r, MatrixAmount.Cell(row.Offset, false), true);
            fixture.Put(sheet, "N" + r, MatrixAmount.Cell(row.Difference, false), true);
            fixture.Put(sheet, "O" + r, row.Note);
        }
        ((dynamic)sheet).Calculate();

        Func<string, double> val = address => Number(fixture, sheet, address);
        Func<string, string> txt = address => Text(fixture, sheet, address);

        Check("应收账款对：资产端 = 800,000（净额法）", Math.Abs(val("I2") - 800000.0) < 0.01, txt("I2"));
        Check("应收账款对：负债端 = 300,000", Math.Abs(val("L2") - 300000.0) < 0.01, txt("L2"));
        Check("应收账款对：抵消 = MIN = 300,000", Math.Abs(val("M2") - 300000.0) < 0.01, txt("M2"));
        Check("应收账款对：差额 = 500,000", Math.Abs(val("N2") - 500000.0) < 0.01, txt("N2"));
        Check("其他应收/应付：资产端 = 65,950,000（关键：净额法）",
            Math.Abs(val("I3") - 65950000.0) < 0.01, txt("I3"));
        Check("其他应收/应付：负债端 = 65,950,000", Math.Abs(val("L3") - 65950000.0) < 0.01, txt("L3"));
        Check("其他应收/应付：抵消 = 65,950,000", Math.Abs(val("M3") - 65950000.0) < 0.01, txt("M3"));
        Check("其他应收/应付：差额 = 0", Math.Abs(val("N3") - 0.0) < 0.01, txt("N3"));
        Check("预付/合同负债：无往来 → 0", Math.Abs(val("I4")) < 0.01 && Math.Abs(val("L4")) < 0.01,
            txt("I4") + "/" + txt("L4"));

        // ---- 结果表：跨表引用差异表的抵消金额格 ----
        object resultSheet = fixture.Sheet("验证-往来结果");
        // 简化布局：第 1 行表头 + 第 2 行起数据（不写 SUBTOTAL/双层表头，只验证公式计算）
        fixture.Put(resultSheet, "A1", "核算主体");
        fixture.Put(resultSheet, "B1", "对方");
        fixture.Put(resultSheet, "E1", "会计科目");
        fixture.Put(resultSheet, "G1", "2026借");
        fixture.Put(resultSheet, "H1", "2026贷");
        foreach (MatrixRow row in result.Rows)
        {
            int r = row.RowNumber;   // BuildResult 给的行号从 dataTop=4 起；这里平移到第 2 行起
            int rr = r - 2;          // 简化：dataTop=4 → 第 2 行
            fixture.Put(resultSheet, "A" + rr, row.Entity);
            fixture.Put(resultSheet, "B" + rr, row.Customer);
            fixture.Put(resultSheet, "E" + rr, row.Account);
            MatrixAmount debitAmt, creditAmt;
            if (row.Debit.TryGetValue("2026", out debitAmt))
                fixture.Put(resultSheet, "G" + rr, MatrixAmount.Cell(debitAmt, false), true);
            if (row.Credit.TryGetValue("2026", out creditAmt))
                fixture.Put(resultSheet, "H" + rr, MatrixAmount.Cell(creditAmt, false), true);
        }
        ((dynamic)resultSheet).Calculate();
        // 结果表金额应 = 差异表抵消金额（跨表引用生效，不再错乱）
        // 行序（dataTop=4，平移 -2 → 第 2 行起）：应收账款借 G2 / 贷 H3；其他应收借 G4 / 贷 H5
        Check("结果表·应收账款对 借 = 300,000（跨表引用差异表 M2）",
            Math.Abs(Number(fixture, resultSheet, "G2") - 300000.0) < 0.01, Text(fixture, resultSheet, "G2"));
        Check("结果表·应收账款对 贷 = 300,000",
            Math.Abs(Number(fixture, resultSheet, "H3") - 300000.0) < 0.01, Text(fixture, resultSheet, "H3"));
        Check("结果表·其他应收对 借 = 65,950,000（跨表引用差异表 M3，关键）",
            Math.Abs(Number(fixture, resultSheet, "G4") - 65950000.0) < 0.01, Text(fixture, resultSheet, "G4"));
        Check("结果表·其他应收对 贷 = 65,950,000",
            Math.Abs(Number(fixture, resultSheet, "H5") - 65950000.0) < 0.01, Text(fixture, resultSheet, "H5"));
    }

    private static double Number(Fixture fixture, object sheet, string addr)
    {
        try { object v = fixture.Read(sheet, addr); return v == null ? double.NaN : Convert.ToDouble(v, CultureInfo.InvariantCulture); }
        catch { return double.NaN; }
    }
    private static string Text(Fixture fixture, object sheet, string addr)
    {
        try { return Convert.ToString(fixture.Read(sheet, addr, true)); } catch { return ""; }
    }
}
