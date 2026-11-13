using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel_Iterate;

// 显式实机探针；0=PASS，1=FAIL，2=SKIP。Application 所有权不明，永不Quit。
// 仅关闭自己创建/打开的隔离样本，成功清理已登记文件，失败保留证据。
class MergeLiveProbe
{
    private static readonly string[] ExcelIds = { "KET.Application", "et.Application", "Excel.Application" };
    static int checks;
    static void Check(bool value, string name)
    {
        checks++;
        if (!value) throw new InvalidOperationException("校验失败: " + name);
        Console.WriteLine("PASS: " + name);
        Console.Out.Flush();
    }

    [STAThread]
    static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            string target = Environment.GetEnvironmentVariable("MERGE_REPRO_TARGET");
            if (!string.IsNullOrEmpty(target))
                return Repro(target, Environment.GetEnvironmentVariable("MERGE_REPRO_SOURCE"));
            if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("MERGE_INSPECT"))) return Inspect();
            return Run();
        }
        catch (Exception error) { Console.WriteLine("FAIL: " + error); return 1; }
    }

    static int Run()
    {
        using (ProbeSession session = new ProbeSession("merge-live"))
        {
            int result = 1;
            try
            {
                object app = session.Take(ProbeSession.Connect(ExcelIds, false));
                if (app == null)
                {
                    Console.WriteLine("SKIP: WPS表格COM宿主不可用。");
                    return session.Finish(2);
                }
                dynamic a = app;
                session.Override(() => (object)a.DisplayAlerts, value => a.DisplayAlerts = value, false, "Excel.DisplayAlerts");
                Console.WriteLine("HOST=" + a.Name + " VERSION=" + Convert.ToString(a.Version));
            // 1. 用 COM 造三个 xlsx（造完就 Close，让 Resolve 自己开）
            string target = session.Output("工作簿1.xlsx");
            string src2 = session.Output("工作簿2.xlsx");
            string src3 = session.Output("工作簿3.xlsx");
            MakeBook(session, app, target, new[] { "A", "B", "C", "D", "E" }, new string[][] {
                new[] { "a1", "b1", "c1", "", "" }, new[] { "a2", "b2", "c2", "", "" },
                new[] { "a3", "b3", "c3", "", "" }, new[] { "a4", "b4", "c4", "", "" },
                new[] { "a5", "b5", "c5", "", "" } });
            MakeBook(session, app, src2, new[] { "A", "B", "C" }, new string[][] {
                new[] { "x1", "y1", "z1" }, new[] { "x2", "y2", "z2" }, new[] { "x3", "y3", "z3" },
                new[] { "x4", "y4", "z4" }, new[] { "x5", "y5", "z5" } });
            MakeBook(session, app, src3, new[] { "D", "E" }, new string[][] {
                new[] { "d1", "e1" }, new[] { "d2", "e2" }, new[] { "d3", "e3" }, new[] { "d4", "e4" } });

            // 1b. 造一个「含公式」的源簿：X7 列（A 列）公式 =B2+C2 形式，验证 ReadColumn 能读出公式。
            //     这是本次修复的目标场景——WPS 对单列 Range 返回一维数组，ReadColumn 需兼容。
            string srcFormula = session.Output("工作簿4.xlsx");
            MakeFormulaBook(session, app, srcFormula);

            // 2. 真实合并链路
            using (MergeWorkbookGateway gateway = new MergeWorkbookGateway(app))
            {
                MergeWorkbookHandle targetHandle = gateway.Resolve(target, "", false);
                Check(targetHandle != null && targetHandle.Book != null && targetHandle.Sheet != null, "Resolve target (Book/Sheet alive)");
                MergeHeaderSet targetSet = MergeSheetReader.ReadHeaders(targetHandle.Sheet, 1, target);
                Check(targetSet.Columns.Count == 5, "target has 5 headers");
                Check(targetSet.LastUsedRow == 6, "target last used row is 6 (header + 5)");

                MergeWorkbookHandle src2Handle = gateway.Resolve(src2, "", true);
                List<string> src2Sheets = MergeWorkbookGateway.SheetNames(src2Handle.Book);
                Check(src2Sheets.Count >= 1, "source2 sheet names read (RCW alive after Resolve)");
                MergeHeaderSet src2Set = MergeSheetReader.ReadHeaders(src2Handle.Sheet, 1, src2);

                MergeWorkbookHandle src3Handle = gateway.Resolve(src3, "", true);
                List<string> src3Sheets = MergeWorkbookGateway.SheetNames(src3Handle.Book);
                Check(src3Sheets.Count >= 1, "source3 sheet names read");
                MergeHeaderSet src3Set = MergeSheetReader.ReadHeaders(src3Handle.Sheet, 1, src3);

                MergeMatchResult match = MergeMatcher.Match(targetSet, new List<MergeHeaderSet> { src2Set, src3Set });
                Check(match.Blocking.Count == 0, "no blocking: " + string.Join("; ", match.Blocking.ToArray()));
                Check(match.MappedColumns == 5, "5 columns matched across two sources");

                // 各列独立追加：先算每列 TargetStartRow = 目标该列末行 + 1，再 ReadColumn 灌 Cells
                List<MergeWorkbookHandle> sourceHandles = new List<MergeWorkbookHandle> { src2Handle, src3Handle };
                for (int i = 0; i < match.Sources.Count; i++)
                {
                    MergeSourcePlan plan = match.Sources[i];
                    MergeWorkbookHandle handle = sourceHandles[i];
                    foreach (MergeColumnMap map in plan.Maps)
                    {
                        map.TargetStartRow = MergeSheetReader.LastDataRow(targetHandle.Sheet, map.TargetColumn, targetSet.HeaderRow, targetSet.LastUsedRow) + 1;
                        MergeSheetReader.ReadColumn(handle.Sheet, map, 1);
                        Check(map.Cells != null && map.Cells.Length == map.RowCount, "source " + (i + 1) + " col " + map.SourceColumn + " read " + (map.Cells == null ? 0 : map.Cells.Length) + " cells");
                    }
                }

                List<MergeColumnMap> allMaps = new List<MergeColumnMap>();
                foreach (MergeSourcePlan s in match.Sources) foreach (MergeColumnMap m in s.Maps) allMaps.Add(m);
                MergeWritePlan plan2 = MergePlanner.Build(allMaps, 1);
                string guard = MergePlanner.Guard(plan2);
                Check(guard == null, "plan guard ok: " + (guard ?? ""));
                Check(plan2.TotalCells == 23, "total 23 cells (5+5+5 + 4+4)");

                // 快照 + 备份 + 写入 + 保存（顺序对齐 RunExecute：先存盘→备份→快照→写→存盘）
                string savedPath = WorkbookSaving.Save(targetHandle.Book, true);
                string backupPath = MergeBackup.Create(savedPath);
                session.TrackFile(backupPath);
                Check(WpsShared.SourcePaths.Exists(backupPath), "backup file created at " + backupPath);
                MergeSnapshot snap = MergeSnapshot.Capture(targetHandle.Sheet, plan2);
                Check(snap != null, "snapshot captured");
                MergeTargetWriter.Write(app, targetHandle.Sheet, plan2, null);
                WorkbookSaving.Save(targetHandle.Book, true);
                Check(WpsShared.SourcePaths.Exists(savedPath), "target saved after write");

                // 3. 读回断言。注意「各列独立追加」：A/B/C 原有 5 行数据→从第 7 行起写；
                // D/E 原本只有表头（空列）→从第 2 行起写。
                dynamic tSheet = targetHandle.Sheet;
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 7, 1, "Value2")) == "x1", "A7=x1 (source2 appended at row 7)");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 11, 1, "Value2")) == "x5", "A11=x5");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 7, 2, "Value2")) == "y1", "B7=y1");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 7, 3, "Value2")) == "z1", "C7=z1");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 2, 4, "Value2")) == "d1", "D2=d1 (source3 appended at row 2, independent column)");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 5, 4, "Value2")) == "d4", "D5=d4");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 2, 5, "Value2")) == "e1", "E2=e1");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 5, 5, "Value2")) == "e4", "E5=e4");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 2, 1, "Value2")) == "a1", "A2=a1 untouched (original data preserved)");

                // 3b. 公式读取回归：真实 WPS 下 ReadColumn 必须读出「公式字符串」，而非数值。
                //     单列 Range 在 WPS 下返回一维数组，At 需兼容（本次修复点）。
                MergeWorkbookHandle fHandle = gateway.Resolve(srcFormula, "", true);
                MergeHeaderSet fSet = MergeSheetReader.ReadHeaders(fHandle.Sheet, 1, srcFormula);
                MergeColumnMap fMap = new MergeColumnMap();
                fMap.SourceColumn = 1;                       // A 列（X7）
                fMap.TargetColumn = 1;
                fMap.RowCount = 3;
                fMap.SourceFirstRow = 2;                     // 数据从第 2 行起
                fMap.Header = "X7";
                MergeSheetReader.ReadColumn(fHandle.Sheet, fMap, 1);
                Check(fMap.FormulaCount == 3, "formula col read 3 formulas (got " + fMap.FormulaCount + ")");
                object f0 = fMap.Cells.Length > 0 ? fMap.Cells[0] : null;
                Check(f0 is string && ((string)f0).StartsWith("=", StringComparison.Ordinal),
                    "formula cell[0] is a formula string (got " + (f0 == null ? "null" : f0.GetType().Name + ":" + f0) + ")");

                // 3c. 单列块 + 公式写入核对回归：用户真实场景是「孤立单列块」（TargetStartRow 与
                //     相邻列不同 → 不合并成宽块）。单列块写入后读回 Value2/Formula 是一维数组，
                //     MergeSnapshot.Pick 必须兼容，否则报「数值写入核对失败」（本次修复点）。
                {
                    MergeColumnMap single = new MergeColumnMap();
                    single.SourceColumn = 1; single.TargetColumn = 1;
                    single.RowCount = 3; single.SourceFirstRow = 2; single.TargetStartRow = 20;  // 孤立的第 20 行起，单列块
                    single.Header = "X7";
                    single.Cells = new object[] { "=B2+C2", "=B3+C3", "=B4+C4" };  // 公式
                    List<MergeColumnMap> one = new List<MergeColumnMap>(); one.Add(single);
                    MergeWritePlan singlePlan = MergePlanner.Build(one, 1);
                    Check(singlePlan.Blocks.Count == 1 && singlePlan.Blocks[0].ColumnCount == 1,
                        "single-column block is isolated (width=1)");
                    MergeSnapshot singleSnap = MergeSnapshot.Capture(targetHandle.Sheet, singlePlan);
                    MergeTargetWriter.Write(app, targetHandle.Sheet, singlePlan, null);
                    // 读回：单列块第 20 行公式应写为 =B20+C20（未映射列 B/C 保持列字母，行偏移随 TargetStartRow）
                    dynamic sSheet = targetHandle.Sheet;
                    string wf0 = Convert.ToString(CellProperty(session, (object)sSheet, 20, 1, "Formula"));
                    Check(wf0.StartsWith("=", StringComparison.Ordinal), "single-col formula written (got " + wf0 + ")");
                    MergeTargetWriter.Restore(app, targetHandle.Sheet, singleSnap);
                }

                // 4. 还原（写入区域回到写入前：A7-A11、D2-D5 全空；A2-A5 原数据保留）
                MergeTargetWriter.Restore(app, targetHandle.Sheet, snap);
                WorkbookSaving.Save(targetHandle.Book, true);
                object a7 = CellProperty(session, (object)tSheet, 7, 1, "Value2");
                Check(a7 == null || Convert.ToString(a7) == "", "restored A7 empty (write reverted)");
                Check(Convert.ToString(CellProperty(session, (object)tSheet, 2, 1, "Value2")) == "a1", "restored A2=a1 (untouched data still there)");
                object d2 = CellProperty(session, (object)tSheet, 2, 4, "Value2");
                Check(d2 == null || Convert.ToString(d2) == "", "restored D2 empty (write reverted)");
            }

                Console.WriteLine("MergeLiveProbe: " + checks + " checks completed");
                result = 0;
            }
            catch (Exception error) { session.Fail("MergeLiveProbe (" + checks + " passed): " + error); result = 1; }
            return session.Finish(result);
        }
    }

    // 仅附加现有宿主；不启动、不写入/保存/关闭用户工作簿，所有读到的引用逐项释放。
    static int Inspect()
    {
        using (ProbeSession session = new ProbeSession("merge-inspect"))
        {
            int result = 1;
            try
            {
                object app = session.Take(ProbeSession.Connect(ExcelIds, true));
                if (app == null)
                {
                    Console.WriteLine("INSPECT SKIP: 没有可附加的现有表格宿主。");
                    return session.Finish(2);
                }
                dynamic a = app;
                Console.WriteLine("HOST=" + a.Name + " VERSION=" + Convert.ToString(a.Version));
                dynamic books = session.Take((object)a.Workbooks);
                int count = Convert.ToInt32(books.Count);
                Console.WriteLine("打开工作簿数: " + count);
                for (int i = 1; i <= count; i++)
                    using (ProbeReferences refs = session.References())
                    {
                        dynamic book = refs.Take((object)MergeWorkbookGateway.GetItem((object)books, i));
                        Console.WriteLine("=== [" + i + "] " + Convert.ToString(book.FullName) + " saved=" + book.Saved);
                        dynamic sheets = refs.Take((object)book.Worksheets);
                        for (int k = 1; k <= Convert.ToInt32(sheets.Count); k++)
                            using (ProbeReferences sheetRefs = session.References())
                            {
                                dynamic sheet = sheetRefs.Take((object)MergeWorkbookGateway.GetItem((object)sheets, k));
                                dynamic used = sheetRefs.Take((object)sheet.UsedRange);
                                dynamic rows = sheetRefs.Take((object)used.Rows);
                                dynamic columns = sheetRefs.Take((object)used.Columns);
                                dynamic cells = sheetRefs.Take((object)sheet.Cells);
                                int top = Convert.ToInt32(used.Row), left = Convert.ToInt32(used.Column);
                                int height = Convert.ToInt32(rows.Count), width = Convert.ToInt32(columns.Count);
                                Console.WriteLine("  表[" + k + "] " + Convert.ToString(sheet.Name)
                                    + " UsedRange=" + CellAddress.Name(top, left) + ":"
                                    + CellAddress.Name(top + height - 1, left + width - 1));
                                int maxRow = Math.Min(top + height - 1, 10);
                                for (int row = top; row <= maxRow; row++)
                                    for (int col = left; col <= left + width - 1; col++)
                                        using (ProbeReferences cellRefs = session.References())
                                        {
                                            dynamic cell = cellRefs.Take((object)cells[row, col]);
                                            object value = cell.Value2, formula = cell.Formula;
                                            bool merged = Convert.ToBoolean(cell.MergeCells);
                                            if (value == null && (formula == null || Equals(formula, "")) && !merged) continue;
                                            Console.WriteLine("    " + CellAddress.Name(row, col) + " V=" + Describe(value)
                                                + " F=" + Describe(formula) + (merged ? " [合并]" : ""));
                                        }
                            }
                    }
                result = 0;
            }
            catch (Exception error) { session.Fail("INSPECT: " + error); result = 1; }
            return session.Finish(result);
        }
    }

    // 用户原文件只读复制；写入/回滚/诊断只在本轮独立目录中操作副本。
    static int Repro(string targetArg, string sourceArg)
    {
        using (ProbeSession session = new ProbeSession("merge-repro"))
        {
            int result = 1;
            try
            {
                string tPath = session.CopyInput(targetArg, "T-" + Path.GetFileName(targetArg));
                string sPath = session.CopyInput(sourceArg, "S-" + Path.GetFileName(sourceArg));
                object app = session.Take(ProbeSession.Connect(ExcelIds, false));
                if (app == null)
                {
                    Console.WriteLine("REPRO SKIP: WPS表格COM宿主不可用。");
                    return session.Finish(2);
                }
                dynamic a = app;
                session.Override(() => (object)a.DisplayAlerts, value => a.DisplayAlerts = value, false, "Excel.DisplayAlerts");
                Console.WriteLine("HOST=" + a.Name + " VERSION=" + Convert.ToString(a.Version));
            using (MergeWorkbookGateway gateway = new MergeWorkbookGateway(app))
            {
                MergeWorkbookHandle th = gateway.Resolve(tPath, "", false);
                MergeWorkbookHandle sh = gateway.Resolve(sPath, "", true);
                MergeHeaderSet tSet = MergeSheetReader.ReadHeaders(th.Sheet, 1, tPath);
                MergeHeaderSet sSet = MergeSheetReader.ReadHeaders(sh.Sheet, 1, sPath);
                MergeMatchResult match = MergeMatcher.Match(tSet, new List<MergeHeaderSet> { sSet });
                Console.WriteLine("匹配 " + match.MappedColumns + " 列，阻断 " + match.Blocking.Count);
                Check(match.Blocking.Count == 0 && match.MappedColumns > 0, "Repro匹配可执行");
                MergeColumnMapper mapper = new MergeColumnMapper();
                foreach (MergeSourcePlan plan in match.Sources)
                    foreach (MergeColumnMap map in plan.Maps)
                    {
                        map.TargetStartRow = MergeSheetReader.LastDataRow(th.Sheet, map.TargetColumn, tSet.HeaderRow, tSet.LastUsedRow) + 1;
                        MergeSheetReader.ReadColumn(sh.Sheet, map, 1);
                        mapper.Add(map.SourceColumn, map.TargetColumn);
                        Console.WriteLine("列 " + map.Header + " 源" + map.SourceLetter + "->目标" + map.TargetLetter
                            + " 写入起始行 " + map.TargetStartRow + " 行数 " + map.RowCount + " 公式数 " + map.FormulaCount);
                        for (int r = 0; r < map.RowCount; r++)
                            Console.WriteLine("   Cells[" + r + "] = " + (map.Cells[r] == null ? "null" : map.Cells[r] + " (" + map.Cells[r].GetType().Name + ")"));
                    }
                foreach (MergeSourcePlan plan in match.Sources)
                    foreach (MergeColumnMap map in plan.Maps)
                        for (int r = 0; r < map.RowCount; r++)
                        {
                            object v = map.Cells[r];
                            if (!MergePlanner.IsFormula(v)) continue;
                            string f = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
                            RewriteContext ctx = new RewriteContext { ColumnDelta = map.TargetColumn - map.SourceColumn, SourcePath = sPath };
                            RewriteResult rr = FormulaRefRewriter.Rewrite(f, mapper, map.RowDelta, null, ctx);
                            Console.WriteLine("重写 " + map.SourceLetter + (map.SourceFirstRow + r) + ": " + f + "  ->  " + rr.Text);
                            map.Cells[r] = rr.Text;
                        }
                List<MergeColumnMap> all = new List<MergeColumnMap>();
                foreach (MergeSourcePlan s2 in match.Sources) all.AddRange(s2.Maps);
                MergeWritePlan wp = MergePlanner.Build(all, 1);
                foreach (MergeWriteBlock b in wp.Blocks)
                    Console.WriteLine("块: 目标列" + b.FirstColumn + " 宽" + b.ColumnCount + " 起始行" + b.StartRow + " 高" + b.RowCount);
                WorkbookSaving.Save(th.Book, true);
                MergeSnapshot snap = MergeSnapshot.Capture(th.Sheet, wp);
                // 两轮执行：模拟用户「执行→回滚→再执行」的确切操作序列。
                for (int round = 1; round <= 2; round++)
                {
                    Console.WriteLine("==== 第 " + round + " 轮执行 ====");
                    try
                    {
                        MergeTargetWriter.Write(app, th.Sheet, wp, null);
                        Console.WriteLine("第 " + round + " 轮: 写入成功");
                        WorkbookSaving.Save(th.Book, true);
                    }
                    catch (Exception ex)
                    {
                        session.Fail("第 " + round + " 轮写入失败: " + ex.Message);
                        // 诊断会话状态（逐项读，避免一个属性失败全挂）
                        foreach (string prop in new string[] { "CalculationState", "Calculation", "CutCopyMode", "Ready", "Interactive", "EditMode", "CellDragAndDrop" })
                            try { Console.WriteLine("  app." + prop + " = " + ((dynamic)a).GetType().InvokeMember(prop, System.Reflection.BindingFlags.GetProperty, null, a, null)); }
                            catch (Exception se) { Console.WriteLine("  app." + prop + " 读取失败: " + se.GetBaseException().Message); }
                        dynamic tSheet0 = th.Sheet;
                        // 合并单元格检查
                        try
                        {
                            dynamic rngCheck = session.Take((object)tSheet0.Range["A3:C7"]);
                            object mc = rngCheck.MergeCells;
                            Console.WriteLine("  A3:C7 MergeCells=" + Describe(mc));
                            dynamic rngCheck2 = session.Take((object)tSheet0.Range["A3"]);
                            Console.WriteLine("  A3 MergeCells=" + Describe(rngCheck2.MergeCells) + " B3 MergeCells=" + Describe(((dynamic)session.Take((object)tSheet0.Range["B3"])).MergeCells));
                        }
                        catch (Exception se) { Console.WriteLine("  合并检查失败: " + se.Message); }
                        // 整块读前单格读 B3
                        Console.WriteLine("  整块读前 B3 单格=" + Describe(((dynamic)session.Take((object)tSheet0.Range["B3"])).Value2));
                        // 失效链接检查与断开实验
                        try
                        {
                            dynamic tBook = th.Book;
                            object links = tBook.LinkSources(1);  // xlExcelLinks=1
                            Array la = links as Array;
                            Console.WriteLine("  LinkSources: " + (la == null ? "null(无链接)" : la.Length + " 个"));
                            if (la != null)
                                foreach (object l in la) Console.WriteLine("    链接: " + l);
                        }
                        catch (Exception se) { Console.WriteLine("  LinkSources 读取失败: " + se.Message); }
                        // 断开链接后整块读是否恢复
                        try
                        {
                            dynamic tBook = th.Book;
                            dynamic rngWhole = session.Take((object)tSheet0.Range["A3:C7"]);
                            object before = rngWhole.Value2;
                            Array ba = before as Array;
                            Console.WriteLine("  断链前整块读 B3 位置=" + (ba == null ? "非数组" : Describe(ba.GetValue(0 + ba.GetLowerBound(0), 1 + ba.GetLowerBound(1)))));
                            object links2 = tBook.LinkSources(1);
                            if (links2 is Array)
                                foreach (object l in (Array)links2)
                                    try { tBook.BreakLink((string)l, 1); Console.WriteLine("  已断开链接: " + l); }
                                    catch (Exception be) { Console.WriteLine("  断链失败 " + l + ": " + be.Message); }
                            object after = ((dynamic)session.Take((object)tSheet0.Range["A3:C7"])).Value2;
                            Array aa = after as Array;
                            Console.WriteLine("  断链后整块读 B3 位置=" + (aa == null ? "非数组" : Describe(aa.GetValue(0 + aa.GetLowerBound(0), 1 + aa.GetLowerBound(1)))));
                        }
                        catch (Exception se) { Console.WriteLine("  断链实验失败: " + se.Message); }
                        // 单格写入测试
                        try
                        {
                            ((dynamic)session.Take((object)tSheet0.Range["B3"])).Value2 = 999;
                            object probe = ((dynamic)session.Take((object)tSheet0.Range["B3"])).Value2;
                            Console.WriteLine("单格写 B3=999 读回: " + Describe(probe));
                        }
                        catch (Exception se) { Console.WriteLine("单格写入失败: " + se.Message); }
                        // 二次批量写入测试
                        try
                        {
                            object[,] retry = new object[1, 2] { { 777, 888 } };
                            ((dynamic)session.Take((object)tSheet0.Range["B3:C3"])).Value2 = retry;
                            object p2 = ((dynamic)session.Take((object)tSheet0.Range["B3"])).Value2;
                            object p3 = ((dynamic)session.Take((object)tSheet0.Range["C3"])).Value2;
                            Console.WriteLine("批量写 B3:C3=777/888 读回: B3=" + Describe(p2) + " C3=" + Describe(p3));
                        }
                        catch (Exception se) { Console.WriteLine("批量写入失败: " + se.Message); }
                        DumpBlocks(session, th.Sheet, wp);
                    }
                    // 第 1 轮成功后回滚，模拟用户失败/取消后的快照还原，再跑第 2 轮。
                    if (round == 1)
                    {
                        MergeTargetWriter.Restore(app, th.Sheet, snap);
                        WorkbookSaving.Save(th.Book, true);
                        Console.WriteLine("第 1 轮后已回滚（模拟用户操作序列）");
                        // 回滚后重新计算起始行（对齐 RunPreview 逻辑）
                        foreach (MergeSourcePlan plan in match.Sources)
                            foreach (MergeColumnMap map in plan.Maps)
                                map.TargetStartRow = MergeSheetReader.LastDataRow(th.Sheet, map.TargetColumn, tSet.HeaderRow, tSet.LastUsedRow) + 1;
                        List<MergeColumnMap> all2 = new List<MergeColumnMap>();
                        foreach (MergeSourcePlan s3 in match.Sources) all2.AddRange(s3.Maps);
                        wp = MergePlanner.Build(all2, 1);
                        foreach (MergeWriteBlock b in wp.Blocks)
                            Console.WriteLine("  第 2 轮块: 目标列" + b.FirstColumn + " 宽" + b.ColumnCount + " 起始行" + b.StartRow + " 高" + b.RowCount);
                        snap = MergeSnapshot.Capture(th.Sheet, wp);
                    }
                }
            }

                result = 0; // Finish保留任何一轮已记录的失败，诊断完成不能把FAIL改成PASS。
            }
            catch (Exception error) { session.Fail("REPRO: " + error); result = 1; }
            return session.Finish(result);
        }
    }

    static void DumpBlocks(ProbeSession session, object sheet, MergeWritePlan wp)
    {
        dynamic tSheet = sheet;
        foreach (MergeWriteBlock b in wp.Blocks)
        {
            dynamic rng = session.Take((object)tSheet.Range[
                CellAddress.Name(b.StartRow, b.FirstColumn) + ":" +
                CellAddress.Name(b.StartRow + b.RowCount - 1, b.FirstColumn + b.ColumnCount - 1)]);
            object fml = MergeSheetReader.ReadFormulas(rng);
            object val = null; try { val = rng.Value2; } catch { }
            Console.WriteLine(" 块 目标列" + b.FirstColumn + " 宽" + b.ColumnCount + " 读回现场:");
            Array fa = fml as Array, va = val as Array;
            Console.WriteLine("   [数组形态] fml=" + (fml == null ? "null" : fml.GetType().Name + (fa != null ? " Rank=" + fa.Rank + " len0=" + fa.GetLength(0) + " len1=" + fa.GetLength(1) : "(非数组)"))
                + "  val=" + (val == null ? "null" : val.GetType().Name + (va != null ? " Rank=" + va.Rank + " len0=" + va.GetLength(0) + " len1=" + va.GetLength(1) : "(非数组)")));
            if (va != null && va.Rank == 2)
            {
                Console.WriteLine("   LowerBound0=" + va.GetLowerBound(0) + " LowerBound1=" + va.GetLowerBound(1));
                for (int r = 0; r < va.GetLength(0); r++)
                {
                    string row = "";
                    for (int c = 0; c < va.GetLength(1); c++)
                    {
                        object x = va.GetValue(r + va.GetLowerBound(0), c + va.GetLowerBound(1));
                        row += "[" + r + "," + c + "]=" + (x == null ? "null" : x.ToString()) + " ";
                    }
                    Console.WriteLine("   row" + r + ": " + row);
                }
            }
            for (int r = 0; r < b.RowCount; r++)
                for (int c = 0; c < b.ColumnCount; c++)
                {
                    object f = MergeSnapshot.Pick(fml as Array, fml, r, c, b.RowCount, b.ColumnCount);
                    object v = MergeSnapshot.Pick(val as Array, val, r, c, b.RowCount, b.ColumnCount);
                    Console.WriteLine("   [" + r + "," + c + "] " + CellAddress.Name(b.StartRow + r, b.FirstColumn + c)
                        + "  F=" + Describe(f) + "  V=" + Describe(v)
                        + "  期望=" + Describe(b.Cells[r, c]));
                }
        }
    }

    static string Describe(object v)
    {
        if (v == null) return "null";
        if (v is System.Double) return ((double)v).ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "(double)";
        return "[" + v.GetType().Name + "] " + v;
    }


    static object CellProperty(ProbeSession session, object sheet, int row, int col, string property)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cells = refs.Take((object)((dynamic)sheet).Cells);
            dynamic cell = refs.Take((object)cells[row, col]);
            return property == "Formula" ? (object)cell.Formula : (object)cell.Value2;
        }
    }

    static void SetCell(ProbeSession session, object sheet, int row, int col, object value, bool formula)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic cells = refs.Take((object)((dynamic)sheet).Cells);
            dynamic cell = refs.Take((object)cells[row, col]);
            if (formula) cell.Formula2 = value;
            else cell.Value2 = value;
        }
    }

    static void MakeBook(ProbeSession session, object app, string path, string[] headers, string[][] rows)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic books = refs.Take((object)((dynamic)app).Workbooks);
            object created = null;
            try
            {
                created = refs.Take((object)books.Add());
                dynamic book = created;
                dynamic sheets = refs.Take((object)book.Worksheets);
                object sheet = refs.Take((object)MergeWorkbookGateway.GetItem((object)sheets, 1));
                for (int col = 0; col < headers.Length; col++) SetCell(session, sheet, 1, col + 1, headers[col], false);
                for (int row = 0; row < rows.Length; row++)
                    for (int col = 0; col < rows[row].Length; col++) SetCell(session, sheet, row + 2, col + 1, rows[row][col], false);
                book.SaveAs(path, 51);
            }
            finally
            {
                if (created != null)
                    try { ((dynamic)created).Close(false); }
                    catch (Exception error) { session.Fail("关闭自建工作簿: " + error.Message); }
            }
        }
    }

    static void MakeFormulaBook(ProbeSession session, object app, string path)
    {
        using (ProbeReferences refs = session.References())
        {
            dynamic books = refs.Take((object)((dynamic)app).Workbooks);
            object created = null;
            try
            {
                created = refs.Take((object)books.Add());
                dynamic book = created;
                dynamic sheets = refs.Take((object)book.Worksheets);
                object sheet = refs.Take((object)MergeWorkbookGateway.GetItem((object)sheets, 1));
                SetCell(session, sheet, 1, 1, "X7", false);
                SetCell(session, sheet, 1, 2, "X2", false);
                SetCell(session, sheet, 1, 3, "X1", false);
                for (int row = 2; row <= 4; row++)
                {
                    SetCell(session, sheet, row, 2, (row - 1) * 10, false);
                    SetCell(session, sheet, row, 3, row - 1, false);
                    SetCell(session, sheet, row, 1, "=B" + row + "+C" + row, true);
                }
                book.SaveAs(path, 51);
            }
            finally
            {
                if (created != null)
                    try { ((dynamic)created).Close(false); }
                    catch (Exception error) { session.Fail("关闭自建公式工作簿: " + error.Message); }
            }
        }
    }
}
