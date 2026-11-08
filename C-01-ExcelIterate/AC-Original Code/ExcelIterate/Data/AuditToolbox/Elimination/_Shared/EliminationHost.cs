using System;
using System.Collections.Generic;
using System.Globalization;

namespace Excel_Iterate
{
    // =====================================================================
    // 合并抵消（往来抵消 / 收入成本抵消）—— COM 宿主层
    //
    // 直接复用 MergeHeadersHost.cs / WorkbookFeatures.cs 的既有组件，不新造：
    //   MergeWorkbookGateway（定位/只读开簿/Hold/CloseOpened/Dispose）
    //   MergeSheetReader    （表头快照；一维/二维数组兼容的 At）
    //   MergeSnapshot       （写入前快照，「取消操作」的主还原路径）
    //   MergeTargetWriter   （HostQuiet 内批量写 + 隐藏行补写 + 读回核对 + Restore）
    //   MergeBackup         （同目录「ExcelIterate合并备份」+ SHA256 核对）
    //   WorkbookUi          （NewSheet 重名编号 / DeleteSheet 删表 / StyleHeader）
    //
    // 本文件只新增抵消特有的三件事：
    //   ① 源数据按角色读成文本行（合并单元格按锚点值向下填充——
    //      往来余额表的「单位」列常见合并且必须填充后聚合才有意义）；
    //   ② 目标工作表定位/新建（「（新建工作表）」项 → 抵消分录表，重名自动编号，
    //      登记实际表名供「取消操作」整表删除）；
    //   ③ 分录 → 6 列写入计划（分录号｜摘要｜方向｜科目｜金额｜计算说明），
    //      表头行 + 分录行一个块写走，快照/核对/还原全部沿用 Merge 组件。
    //
    // 线程纪律同 Connect.MergeHeaders.cs：COM 只在 Ribbon 宿主线程
    // （窗体事件与回调同线程，不开后台 STA）。
    // =====================================================================

    /// <summary>抵消写入目标：一次「定位/新建」的结果。</summary>
    internal sealed class EliminationTargetHandle
    {
        internal object Book;
        internal object Sheet;
        internal bool OpenedByUs;
        internal string Path = "";
        internal string SheetName = "";
        internal bool IsNewSheet;     // 本次新建的工作表（还原时整表删除）
    }

    /// <summary>
    /// 抵消功能的 COM 编排：源读取 + 目标定位/新建 + 备份 + 写入 + 还原。
    /// 生命周期跟随窗体：FormClosed 时 Dispose（关掉自己开的簿、释放持有引用）。
    /// </summary>
    internal sealed class EliminationGateway : IDisposable
    {
        /// <summary>工作表下拉里的「新建工作表」项。</summary>
        internal const string NewSheetItem = "（新建工作表）";
        /// <summary>新建表的默认名（重名时由 WorkbookUi.UniqueName 编号）。</summary>
        internal const string DefaultSheetName = "抵消分录";
        /// <summary>分录区列数：分录号｜摘要｜方向｜科目｜金额｜计算说明。</summary>
        internal const int EntryColumns = 6;

        /// <summary>矩阵式底稿新建表的默认名（收入成本抵消专用）。</summary>
        internal const string MatrixDefaultSheetName = "收入抵消结果";
        /// <summary>矩阵式底稿的表头层数：第 1 层年度、第 2 层借/贷。</summary>
        internal const int MatrixHeaderRows = 2;
        /// <summary>往来抵消「差异表」的默认名（固定新建，不提供下拉选择）。</summary>
        internal const string DiffDefaultSheetName = "往来抵消差异";
        /// <summary>往来抵消「结果表」的默认名。</summary>
        internal const string ReceivableDefaultSheetName = "往来抵消结果";

        /// <summary>矩阵式底稿默认表名按业务取。</summary>
        internal static string MatrixSheetNameFor(EliminationKind kind)
        { return kind == EliminationKind.RevenueCost ? MatrixDefaultSheetName : ReceivableDefaultSheetName; }

        /// <summary>分录区表头（预览网格与写入表头共用，避免两处不一致）。</summary>
        internal static readonly string[] EntryHeaders = new[]
        {
            "分录号", "摘要", "方向", "科目", "金额", "计算说明"
        };

        private const int ChunkRows = 4096;

        private readonly object app;
        private MergeWorkbookGateway gateway;
        private readonly List<object> createdSheets = new List<object>();
        private bool disposed;

        internal EliminationGateway(object application) { app = application; }

        private MergeWorkbookGateway Ensure()
        {
            if (gateway == null) gateway = new MergeWorkbookGateway(app);
            return gateway;
        }

        // ---------------------------- 源：读取 ----------------------------

        /// <summary>列出工作簿的工作表名（供界面下拉；排除隐藏表）。</summary>
        internal List<string> SheetNames(string path)
        {
            MergeWorkbookGateway g = Ensure();
            MergeWorkbookHandle handle = g.Resolve(path, "", true);
            return MergeWorkbookGateway.SheetNames(handle.Book);
        }

        /// <summary>读取源表头行快照（含 LastUsedRow，供数据区定位）。</summary>
        internal MergeHeaderSet ReadHeaders(string path, string sheetName, int headerRow)
        {
            MergeWorkbookGateway g = Ensure();
            MergeWorkbookHandle handle = g.Resolve(path, sheetName, true);
            return MergeSheetReader.ReadHeaders(handle.Sheet, headerRow, path);
        }

        internal EliminationSnapshotReader BeginReadRows(EliminationRequest request,MergeHeaderSet headers)
        {
            var handle=Ensure().Resolve(request.Path,request.SheetName,true);
            return new EliminationSnapshotReader(app,handle.Book,handle.Sheet,request.Columns,request.HeaderRow,headers.LastUsedRow);
        }

        /// <summary>
        /// 按列映射读取数据行：只读值不搬公式（抵消要的是账面余额的数值口径）。
        /// fillMerged=true 时，数据区含合并单元格按锚点值向下填充（往来余额表的单位列依赖这个）；
        /// 矩阵式底稿传 false——年列/主体列若有合并，向下填充会虚构出不存在的数据行。
        /// 返回行的 SourceRow 为源工作表真实行号（供提示定位）。
        /// </summary>
        internal List<EliminationRow> ReadRows(string path, string sheetName, int headerRow,
            List<EliminationColumn> columns, int lastUsedRow, bool fillMerged = true)
        {
            List<EliminationRow> rows = new List<EliminationRow>();
            if (columns == null) return rows;
            List<EliminationColumn> mapped = new List<EliminationColumn>();
            foreach (EliminationColumn column in columns)
                if (column.Role != EliminationRole.Ignore) mapped.Add(column);
            if (mapped.Count == 0) return rows;

            int rowCount = lastUsedRow - headerRow;
            if (rowCount <= 0) return rows;

            MergeWorkbookGateway g = Ensure();
            MergeWorkbookHandle handle = g.Resolve(path, sheetName, true);
            for (int i = 0; i < rowCount; i++)
                rows.Add(new EliminationRow { SourceRow = headerRow + 1 + i });

            foreach (EliminationColumn column in mapped)
                ReadOneColumn(handle.Sheet, column, headerRow, rowCount, rows, fillMerged);
            return rows;
        }

        private static void ReadOneColumn(object sheet, EliminationColumn column,
            int headerRow, int rowCount, List<EliminationRow> rows, bool fillMerged = true)
        {
            dynamic s = sheet;
            string lastText = "";   // 合并单元格锚点值（向下填充用）
            for (int offset = 0; offset < rowCount; offset += ChunkRows)
            {
                int height = Math.Min(ChunkRows, rowCount - offset);
                using (ComBag chunk = new ComBag())
                {
                    dynamic range = chunk.Own(s.Range[
                        CellAddress.Name(headerRow + 1 + offset, column.Number) + ":"
                        + CellAddress.Name(headerRow + offset + height, column.Number)]);
                    bool merged = false;
                    try { merged = ComFormulaHost.Boolean(range.MergeCells, "MergeCells"); }
                    catch { }
                    object values = null;
                    try { values = range.Value2; }
                    catch { }
                    for (int r = 0; r < height; r++)
                    {
                        // WPS 单列 Range 的 Value2 是一维数组、Excel 是二维，At 两种都兼容。
                        object raw = MergeSheetReader.At(values, r, height, 0);
                        string text = raw == null ? "" : Convert.ToString(raw, CultureInfo.InvariantCulture);
                        if (fillMerged && merged && text.Length == 0) text = lastText;
                        else if (text.Length > 0) lastText = text;
                        rows[offset + r].Values[column.Role] = text;
                    }
                }
            }
        }

        // ---------------------------- 目标：定位/新建 ----------------------------

        /// <summary>
        /// 定位或新建目标工作表。sheetItem 为 NewSheetItem 时在目标工作簿末尾
        /// 新建「抵消分录」表（重名自动编号），并登记实际表名供还原删表。
        /// </summary>
        internal EliminationTargetHandle CreateTarget(string path, string sheetItem,
            string defaultName = DefaultSheetName)
        {
            MergeWorkbookGateway g = Ensure();
            EliminationTargetHandle target = new EliminationTargetHandle();
            target.Path = path;

            if (!string.Equals(sheetItem, NewSheetItem, StringComparison.Ordinal))
            {
                MergeWorkbookHandle handle = g.Resolve(path, sheetItem, false);
                target.Book = handle.Book;
                target.Sheet = handle.Sheet;
                target.SheetName = handle.SheetName;
                target.OpenedByUs = handle.OpenedByUs;
                target.IsNewSheet = false;
                return target;
            }

            MergeWorkbookHandle book = g.Resolve(path, "", false);
            target.Book = book.Book;
            target.OpenedByUs = book.OpenedByUs;
            using (ComBag bag = new ComBag())
            {
                dynamic added = WorkbookUi.NewSheet(book.Book, defaultName, bag, false);
                target.Sheet = g.Hold(bag.Detach(added));
                target.SheetName = Convert.ToString(((dynamic)target.Sheet).Name);
            }
            target.IsNewSheet = true;
            createdSheets.Add(target.Sheet);
            return target;
        }

        /// <summary>本次是否新建过工作表（窗体关闭时用于日志提示）。</summary>
        internal bool HasCreatedSheet { get { return createdSheets.Count > 0; } }

        // ---------------------------- 写入计划 ----------------------------

        /// <summary>
        /// 分录 → 6 列写入计划：第 1 行表头，其后每个借贷行一行。
        /// 金额写 double（COM Value2 数值口径，读回核对的 SameValue 才能对齐）。
        /// </summary>
        internal static MergeWritePlan BuildWritePlan(EliminationResult result, int startColumn, int startRow)
        {
            MergeWritePlan plan = new MergeWritePlan();
            if (result == null) return plan;
            int lineCount = 0;
            foreach (EliminationEntry entry in result.Entries) lineCount += entry.Lines.Count;
            if (lineCount == 0) return plan;

            MergeWriteBlock block = new MergeWriteBlock();
            block.FirstColumn = startColumn;
            block.ColumnCount = EntryColumns;
            block.StartRow = startRow;
            block.RowCount = lineCount + 1;   // 表头行 + 分录行
            block.Cells = new object[block.RowCount, EntryColumns];
            for (int c = 0; c < EntryColumns; c++) block.Cells[0, c] = EntryHeaders[c];
            int r = 1;
            foreach (EliminationEntry entry in result.Entries)
            {
                foreach (EliminationLine line in entry.Lines)
                {
                    block.Cells[r, 0] = entry.Index;
                    block.Cells[r, 1] = entry.Summary;
                    block.Cells[r, 2] = line.IsDebit ? "借" : "贷";
                    block.Cells[r, 3] = line.Account;
                    block.Cells[r, 4] = Convert.ToDouble(line.Amount);
                    block.Cells[r, 5] = line.Note;
                    r++;
                }
            }
            block.AnyFormula = false;
            plan.Blocks.Add(block);
            plan.TotalCells = block.RowCount * EntryColumns;
            plan.FormulaCells = 0;
            return plan;
        }

        // ---------------------------- 矩阵式底稿写入计划 ----------------------------

        /// <summary>
        /// 矩阵底稿 → 单个写入块：第 1 行 SUBTOTAL 合计 + 双层表头 + 数据行。
        /// 金额格填 "=..." 字符串（MergePlanner.IsFormula 据此判定公式），
        /// 静态降级模式下填 double。
        ///
        /// 起始行/列一律取自 options.StartRow / options.StartColumn——
        /// 公式里的单元格锚点在 Build 阶段就已按这两个值生成，
        /// 此处若另行传入不同值，写入位置与公式引用会整体错位。
        /// </summary>
        internal static MergeWritePlan BuildMatrixPlan(EliminationMatrix matrix, MatrixOptions options)
        {
            MergeWritePlan plan = new MergeWritePlan();
            if (matrix == null || matrix.Rows.Count == 0 || matrix.Columns.Count == 0) return plan;
            if (options == null) options = new MatrixOptions();

            int startRow = options.StartRow;
            int startColumn = options.StartColumn;
            int firstHeaderRow = startRow + (options.TotalRow ? 1 : 0);
            int headerRows = MatrixHeaderRows;
            int dataTop = firstHeaderRow + headerRows;
            // 行数由最后一行的 RowNumber 决定：RowNumber 由构建器算好，组内行数与组间空行都已包含。
            // 不能再用「每组固定 TemplateCount 行、每 height 行加一个 gap」推算——
            // 往来抵消每个主体对的行数 = 有金额的科目对数 × 2，并不固定，
            // 旧假设会让空行插入位置错位、行指针越界（实机验收：执行写入报「索引超出了数组界限」）。
            int lastRow = dataTop - 1;
            foreach (MatrixRow rowScan in matrix.Rows)
                if (rowScan.RowNumber > lastRow) lastRow = rowScan.RowNumber;
            int dataCount = matrix.Rows.Count > 0 ? lastRow - dataTop + 1 : 0;
            int rowCount = (options.TotalRow ? 1 : 0) + headerRows + dataCount;

            MergeWriteBlock block = new MergeWriteBlock();
            block.FirstColumn = startColumn;
            block.ColumnCount = matrix.Columns.Count;
            block.StartRow = startRow;
            block.RowCount = rowCount;
            block.Cells = new object[rowCount, matrix.Columns.Count];

            int r = 0;
            // ---- 第 1 行：金额列写 SUBTOTAL 合计 ----
            if (options.TotalRow)
            {
                foreach (MatrixColumn column in matrix.Columns)
                {
                    if (column.IsSpacer || column.Year == null) continue;
                    block.Cells[r, column.Index] = "=SUBTOTAL(9," + column.Letter + dataTop
                        + ":" + column.Letter + MergeLimits.MaxRow + ")";
                }
                r++;
            }
            // ---- 双层表头 ----
            foreach (MatrixColumn column in matrix.Columns)
            {
                if (column.IsSpacer) continue;
                block.Cells[r, column.Index] = column.Head2;
                block.Cells[r + 1, column.Index] = column.Head3;
            }
            r += headerRows;

            // ---- 数据行：按构建器给出的 RowNumber 直接定位（块内偏移 = RowNumber - startRow，
            //      组间空行自然留 null；不再按固定模板行数推算）----
            foreach (MatrixRow row in matrix.Rows)
            {
                int cellRow = row.RowNumber - startRow;
                if (cellRow < 0 || cellRow >= rowCount) continue;   // 越界保护，正常不应触发
                block.Cells[cellRow, 0] = row.Entity;
                block.Cells[cellRow, 1] = row.Customer;
                block.Cells[cellRow, 2] = row.Note;
                block.Cells[cellRow, 3] = row.Adjustment;
                block.Cells[cellRow, 4] = row.Account;
                foreach (string year in matrix.Years)
                {
                    MatrixColumn debit = matrix.DebitColumn(year);
                    MatrixColumn credit = matrix.CreditColumn(year);
                    MatrixAmount amount;
                    if (debit != null && row.Debit.TryGetValue(year, out amount))
                        block.Cells[cellRow, debit.Index] = MatrixAmount.Cell(amount, options.StaticValues);
                    if (credit != null && row.Credit.TryGetValue(year, out amount))
                        block.Cells[cellRow, credit.Index] = MatrixAmount.Cell(amount, options.StaticValues);
                }
            }

            block.AnyFormula = !options.StaticValues;
            plan.Blocks.Add(block);
            int total = rowCount * matrix.Columns.Count;
            plan.TotalCells = total;
            plan.FormulaCells = options.StaticValues ? 0 : CountFormulas(block);
            return plan;
        }

        private static int CountFormulas(MergeWriteBlock block)
        {
            int count = 0;
            for (int r = 0; r < block.RowCount; r++)
                for (int c = 0; c < block.ColumnCount; c++)
                    if (MergePlanner.IsFormula(block.Cells[r, c])) count++;
            return count;
        }

        /// <summary>
        /// 「往来抵消差异」表 → 写入计划：第 1 行表头 + 数据行（每个主体对之间空 GroupGap 行）。
        /// 与矩阵底稿不同，差异表只有一层表头、没有 SUBTOTAL 行。
        /// </summary>
        internal static MergeWritePlan BuildDiffPlan(EliminationDiffTable diff, MatrixOptions options)
        {
            MergeWritePlan plan = new MergeWritePlan();
            if (diff == null || diff.Rows.Count == 0) return plan;
            if (options == null) options = new MatrixOptions();

            int startRow = options.StartRow;
            int startColumn = options.StartColumn;
            int gap = options.GroupGap > 0 ? options.GroupGap : 0;
            int gaps = gap > 0 && diff.GroupCount > 1 ? (diff.GroupCount - 1) * gap : 0;
            int rowCount = 1 + diff.Rows.Count + gaps;

            MergeWriteBlock block = new MergeWriteBlock();
            block.FirstColumn = startColumn;
            block.ColumnCount = diff.ColumnCount;
            block.StartRow = startRow;
            block.RowCount = rowCount;
            block.Cells = new object[rowCount, diff.ColumnCount];

            int r = 0;
            for (int c = 0; c < diff.ColumnCount; c++) block.Cells[r, c] = EliminationDiffTable.Headers[c];
            r++;

            int lastGroup = -1;
            for (int i = 0; i < diff.Rows.Count; i++)
            {
                EliminationDiffRow row = diff.Rows[i];
                if (lastGroup >= 0 && row.GroupIndex != lastGroup) r += gap;
                lastGroup = row.GroupIndex;
                block.Cells[r, 0] = row.Index;
                block.Cells[r, 1] = YearCell(row.Year);
                block.Cells[r, 2] = row.Creditor;
                block.Cells[r, 3] = row.Debtor;
                block.Cells[r, 4] = row.AssetAccount;
                block.Cells[r, 5] = row.LiabilityAccount;
                block.Cells[r, 6] = MatrixAmount.Cell(row.CreditorAsset, options.StaticValues);
                block.Cells[r, 7] = MatrixAmount.Cell(row.CreditorLiability, options.StaticValues);
                block.Cells[r, 8] = MatrixAmount.Cell(row.Asset, options.StaticValues);
                block.Cells[r, 9] = MatrixAmount.Cell(row.DebtorAsset, options.StaticValues);
                block.Cells[r, 10] = MatrixAmount.Cell(row.DebtorLiability, options.StaticValues);
                block.Cells[r, 11] = MatrixAmount.Cell(row.Liability, options.StaticValues);
                block.Cells[r, 12] = MatrixAmount.Cell(row.Offset, options.StaticValues);
                block.Cells[r, 13] = MatrixAmount.Cell(row.Difference, options.StaticValues);
                block.Cells[r, 14] = row.Note;
                r++;
            }

            block.AnyFormula = !options.StaticValues;
            plan.Blocks.Add(block);
            plan.TotalCells = rowCount * diff.ColumnCount;
            plan.FormulaCells = options.StaticValues ? 0 : CountFormulas(block);
            return plan;
        }

        /// <summary>年度写数值（与样张一致）；公式里再用 TEXT(…,"0") 转文本比较。</summary>
        private static object YearCell(string year)
        {
            int value;
            if (int.TryParse(year, NumberStyles.Integer, CultureInfo.InvariantCulture, out value)) return value;
            return year;
        }

        /// <summary>差异表表头样式（单层表头）：与矩阵底稿一致的深蓝底白字加粗 + 细框线 + 居中。</summary>
        internal void StyleDiffHeader(EliminationTargetHandle target, int headerRow,
            int startColumn, int columns)
        {
            if (target == null || target.Sheet == null) return;
            dynamic s = target.Sheet;
            string address = CellAddress.Name(headerRow, startColumn) + ":"
                + CellAddress.Name(headerRow, startColumn + columns - 1);
            using (ComBag bag = new ComBag())
            {
                dynamic range = bag.Own(s.Range[address]);
                dynamic font = bag.Own(range.Font), fill = bag.Own(range.Interior);
                fill.Color = 0x784E1F;
                font.Color = 0xFFFFFF;
                font.Bold = true;
                font.Size = 10;
                dynamic borders = bag.Own(range.Borders);
                borders.LineStyle = 1;
                borders.Weight = 2;
                range.HorizontalAlignment = -4108;
                range.VerticalAlignment = -4108;
                range.WrapText = false;
                range.RowHeight = 26;   // 规范 (2)i：表头行按字体适配、保留 26 磅下限
            }
        }

        /// <summary>
        /// 接入开发日志第四章 (2) 规范的其余要素：金额列会计数字格式（含千位分隔）、
        /// 数据区垂直居中、指定列水平居中、列宽 AutoFit（含表头+数据）、表头行高 26 下限。
        /// 不重设表头颜色——颜色已由 StyleDiffHeader/StyleMatrixHeader 设为深蓝白字。
        /// </summary>
        internal void FinishMatrixOutput(EliminationTargetHandle target,
            int headerRow, int headerRows, int dataTop, int startColumn, int columnCount, int rowCount,
            int[] moneyColumns, int[] centerColumns)
        {
            FinishMatrixOutput(target, headerRow, headerRows, dataTop, startColumn, columnCount, rowCount,
                moneyColumns, centerColumns, false);
        }

        /// <summary>
        /// 带「格式参考」样张变体：referenceStyle=true 时按用户样张追加——
        /// 全表宋体 11、边框 left=medium + 其余 slantDashDot（斜点划线网格、左边缘粗线）、
        /// 会计数字格式从第 1 行（SUBTOTAL 行）开始铺（样张 G1 即此格式）。
        /// </summary>
        internal void FinishMatrixOutput(EliminationTargetHandle target,
            int headerRow, int headerRows, int dataTop, int startColumn, int columnCount, int rowCount,
            int[] moneyColumns, int[] centerColumns, bool referenceStyle)
        {
            if (target == null || target.Sheet == null || rowCount <= 0) return;
            dynamic s = target.Sheet;
            // 会计格式：正数千位分隔、负数带减号、零显短横；与样张一致
            string moneyFmt = "_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * \"-\"??_ ;_ @_ ";
            using (ComBag bag = new ComBag())
            {
                // 「格式参考」样张：表头与数据区统一宋体 11（表头 Bold 由 Style 方法保留）
                if (referenceStyle)
                {
                    // 「格式维护」样张：全表（含 R1 SUBTOTAL 行）统一 Times New Roman 11；
                    // 表头 Bold 由 StyleMatrixHeader 保留
                    int fontTop = headerRow > 1 ? headerRow - 1 : headerRow;
                    dynamic whole = bag.Own(s.Range[CellAddress.Name(fontTop, startColumn) + ":"
                        + CellAddress.Name(dataTop + rowCount - 1, startColumn + columnCount - 1)]);
                    dynamic font = bag.Own(whole.Font);
                    font.Name = "Times New Roman";
                    font.Size = 11;
                }
                // 金额列会计格式：referenceStyle 时只铺 R1（SUBTOTAL）+ 数据区，**跳过 R2/R3 表头行**
                // ——表头年度格有 General"年""度" 等专用格式，被会计格式覆盖会变成「2,026.00」（用户实机反馈）
                if (moneyColumns != null)
                {
                    int subtotalRow = referenceStyle ? (headerRow - 1) : 0;
                    for (int i = 0; i < moneyColumns.Length; i++)
                    {
                        int col = startColumn + moneyColumns[i];
                        string addr = CellAddress.Name(dataTop, col) + ":"
                            + CellAddress.Name(dataTop + rowCount - 1, col);
                        dynamic range = bag.Own(s.Range[addr]);
                        range.NumberFormat = moneyFmt;
                        if (subtotalRow >= 1)
                        {
                            string subAddr = CellAddress.Name(subtotalRow, col) + ":"
                                + CellAddress.Name(subtotalRow, col);
                            dynamic subRange = bag.Own(s.Range[subAddr]);
                            subRange.NumberFormat = moneyFmt;
                        }
                    }
                }
                // 数据区垂直居中
                string bodyAddr = CellAddress.Name(dataTop, startColumn) + ":" + CellAddress.Name(dataTop + rowCount - 1, startColumn + columnCount - 1);
                dynamic body = bag.Own(s.Range[bodyAddr]);
                body.VerticalAlignment = -4108;
                // 指定列水平居中
                if (centerColumns != null)
                {
                    for (int i = 0; i < centerColumns.Length; i++)
                    {
                        int col = startColumn + centerColumns[i];
                        string addr = CellAddress.Name(dataTop, col) + ":" + CellAddress.Name(dataTop + rowCount - 1, col);
                        dynamic range = bag.Own(s.Range[addr]);
                        range.HorizontalAlignment = -4108;
                    }
                }
                // 「格式参考」样张边框：内部 slantDashDot（斜点划线）、左边缘与顶边 medium
                if (referenceStyle)
                {
                    string allAddr = CellAddress.Name(headerRow, startColumn) + ":"
                        + CellAddress.Name(dataTop + rowCount - 1, startColumn + columnCount - 1);
                    dynamic all = bag.Own(s.Range[allAddr]);
                    dynamic borders = bag.Own(all.Borders);
                    borders.LineStyle = 13;      // xlSlantDashDot
                    dynamic edgeLeft = bag.Own(all.Borders[7]);    // xlEdgeLeft
                    edgeLeft.LineStyle = 1;      // xlContinuous
                    edgeLeft.Weight = -4138;     // xlMedium
                    dynamic edgeTop = bag.Own(all.Borders[9]);     // xlEdgeTop
                    edgeTop.LineStyle = 1;
                    edgeTop.Weight = -4138;
                }
                // 列宽 AutoFit（含表头+数据，规范 ii）；表头不换行，AutoFit 保证单行显示
                string fitAddr = CellAddress.Name(headerRow, startColumn) + ":"
                    + CellAddress.Name(dataTop + rowCount - 1, startColumn + columnCount - 1);
                dynamic fit = bag.Own(s.Range[fitAddr]);
                fit.WrapText = false;
                dynamic cols = bag.Own(fit.Columns); cols.AutoFit();
                // 表头行高 26 下限（双层表头每行都设）
                for (int r = 0; r < headerRows; r++)
                {
                    dynamic header = bag.Own(s.Range[CellAddress.Name(headerRow + r, startColumn) + ":" + CellAddress.Name(headerRow + r, startColumn + columnCount - 1)]);
                    object h = header.RowHeight;
                    double hd = h == null ? 0 : Convert.ToDouble(h, CultureInfo.InvariantCulture);
                    if (hd < 26) header.RowHeight = 26;
                }
            }
        }

        /// <summary>从矩阵列定义收集金额列的相对索引（非分隔、有年度的借/贷列）。</summary>
        internal static int[] MoneyColumnIndices(EliminationMatrix matrix)
        {
            if (matrix == null || matrix.Columns == null) return new int[0];
            List<int> list = new List<int>();
            foreach (MatrixColumn col in matrix.Columns)
                if (!col.IsSpacer && col.Year != null && col.Year.Length > 0) list.Add(col.Index);
            return list.ToArray();
        }

        /// <summary>
        /// 写入前把目标矩形清干净：先拆旧合并再清内容。
        /// 目标若选了已存在的表且之前写过矩阵（残留 G2:H2 合并），
        /// 不拆合并直接赋值会抛 COM 异常或只写进锚点格。
        /// </summary>
        internal void PrepareTargetRegion(EliminationTargetHandle target, MergeWritePlan plan)
        {
            if (target == null || target.Sheet == null || plan == null) return;
            dynamic s = target.Sheet;
            foreach (MergeWriteBlock block in plan.Blocks)
            {
                string address = CellAddress.Name(block.StartRow, block.FirstColumn) + ":"
                    + CellAddress.Name(block.StartRow + block.RowCount - 1,
                                             block.FirstColumn + block.ColumnCount - 1);
                using (ComBag bag = new ComBag())
                {
                    dynamic range = bag.Own(s.Range[address]);
                    try { range.UnMerge(); } catch { }
                    try { range.ClearContents(); } catch { }
                }
            }
        }

        /// <summary>
        /// 写入之后做表头合并：固定五列纵向合并 X2:X3、每个年度的借/贷列横向合并、
        /// 分隔列纵向合并。单个合并失败不影响整批。
        /// </summary>
        internal void MergeMatrixHeader(EliminationTargetHandle target, EliminationMatrix matrix,
            int startRow, int startColumn, bool totalRow)
        {
            if (target == null || target.Sheet == null || matrix == null) return;
            int first = startRow + (totalRow ? 1 : 0);
            dynamic s = target.Sheet;

            for (int i = 0; i < EliminationMatrixBuilder.FixedColumns; i++)
                Merge(s, first, startColumn + i, first + 1, startColumn + i);

            foreach (string year in matrix.Years)
            {
                MatrixColumn debit = matrix.DebitColumn(year);
                MatrixColumn credit = matrix.CreditColumn(year);
                if (debit == null || credit == null) continue;
                Merge(s, first, debit.Number, first, credit.Number);
            }
            foreach (MatrixColumn column in matrix.Columns)
            {
                if (!column.IsSpacer) continue;
                Merge(s, first, column.Number, first + 1, column.Number);
            }
        }

        private static void Merge(dynamic sheet, int row1, int column1, int row2, int column2)
        {
            if (row1 == row2 && column1 == column2) return;
            using (ComBag bag = new ComBag())
            {
                try
                {
                    dynamic range = bag.Own(sheet.Range[
                        CellAddress.Name(row1, column1) + ":"
                        + CellAddress.Name(row2, column2)]);
                    range.Merge();
                    range.HorizontalAlignment = -4108;
                    range.VerticalAlignment = -4108;
                }
                catch { }
            }
        }

        /// <summary>
        /// 矩阵底稿表头样式：复刻用户样张「收入抵消结果」——
        /// 深蓝底（RGB 1F4E78）、白色加粗 10 号字、四边细框线、水平垂直居中；
        /// 年度合并格数字格式 General"年""度"（显示如「2024年度」）。
        /// 不用 WorkbookUi.StyleHeader（那是插件统一的浅紫表头，与样张不符）。
        /// </summary>
        internal void StyleMatrixHeader(EliminationTargetHandle target, EliminationMatrix matrix,
            int firstHeaderRow, int startColumn, bool totalRow)
        {
            if (target == null || target.Sheet == null || matrix == null) return;
            dynamic s = target.Sheet;
            int rows = totalRow ? 2 : 2;   // 第 1 层年度 + 第 2 层借/贷
            string address = CellAddress.Name(firstHeaderRow, startColumn) + ":"
                + CellAddress.Name(firstHeaderRow + rows - 1, startColumn + matrix.Columns.Count - 1);
            using (ComBag bag = new ComBag())
            {
                dynamic range = bag.Own(s.Range[address]);
                dynamic font = bag.Own(range.Font), fill = bag.Own(range.Interior);
                fill.Color = 0xCB7448;              // OLE(BGR) = RGB(0x48,0x74,0xCB) = #4874CB（「格式参考」主题色 accent1）
                font.Color = 0xFFFFFF;              // 白
                font.Bold = true;
                font.Name = "Times New Roman";   // 「格式维护」样张字体
                font.Size = 11;
                dynamic borders = bag.Own(range.Borders);
                borders.LineStyle = 1;              // xlContinuous（细线）
                borders.Weight = 2;                 // xlThin
                range.HorizontalAlignment = -4108;  // 居中
                range.VerticalAlignment = -4108;
                range.WrapText = false;
                // 规范 (2)i：表头行按字体适配、保留 26 磅下限（双层表头每行都设）
                for (int r = 0; r < rows; r++)
                {
                    dynamic hr = bag.Own(s.Range[CellAddress.Name(firstHeaderRow + r, startColumn) + ":" + CellAddress.Name(firstHeaderRow + r, startColumn + matrix.Columns.Count - 1)]);
                    object h = hr.RowHeight;
                    double hd = h == null ? 0 : Convert.ToDouble(h, CultureInfo.InvariantCulture);
                    if (hd < 26) hr.RowHeight = 26;
                }
            }
            // 年度格显示「2024年度」：样张数字格式 General"年""度"
            foreach (string year in matrix.Years)
            {
                MatrixColumn debit = matrix.DebitColumn(year);
                if (debit == null) continue;
                using (ComBag bag = new ComBag())
                {
                    try
                    {
                        dynamic cell = bag.Own(s.Range[
                            CellAddress.Name(firstHeaderRow, debit.Number) + ":"
                            + CellAddress.Name(firstHeaderRow, debit.Number + 1)]);
                        cell.NumberFormat = "General\"年\"\"度\"";
                    }
                    catch { }
                }
            }
        }

        /// <summary>解析「B3」形态的起始单元格；越界或非法返回 false。</summary>
        internal static bool TryParseCell(string text, out int row, out int column)
        {
            row = 0; column = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string s = text.Trim().ToUpperInvariant();
            int i = 0;
            while (i < s.Length && s[i] >= 'A' && s[i] <= 'Z') i++;
            if (i == 0 || i >= s.Length) return false;
            int letters = CellAddress.ColumnNumber(s.Substring(0, i), MergeLimits.MaxColumn);
            if (letters == 0) return false;
            int digits = 0;
            for (int k = i; k < s.Length; k++)
            {
                if (s[k] < '0' || s[k] > '9') return false;
                digits = digits * 10 + (s[k] - '0');
                if (digits > MergeLimits.MaxRow) return false;
            }
            column = letters;
            row = digits;
            return row >= 1;
        }

        // ---------------------------- 快照/写入/还原/备份 ----------------------------

        internal MergeSnapshot Capture(EliminationTargetHandle target, MergeWritePlan plan)
        {
            return MergeSnapshot.Capture(target.Sheet, plan);
        }

        internal void Write(EliminationTargetHandle target, MergeWritePlan plan)
        {
            MergeTargetWriter.Write(app, target.Sheet, plan, null);
        }

        /// <summary>
        /// 还原：新建表整表删除（快照还原对空表没有意义）；
        /// 已存在的表用写入前快照回写。
        /// </summary>
        internal void Restore(EliminationTargetHandle target, MergeSnapshot snapshot)
        {
            if (target == null || target.Sheet == null) return;
            if (target.IsNewSheet)
            {
                WorkbookUi.DeleteSheet(app, target.Sheet);
                createdSheets.Remove(target.Sheet);
                return;
            }
            MergeTargetWriter.Restore(app, target.Sheet, snapshot);
        }

        /// <summary>
        /// 给写入的表头行套用与插件其它输出一致的表头样式。
        /// 必须走 StyleHeader 的四参重载 (sheet, row, startColumn, columns)——
        /// 三参重载会把行号硬写成 1，起始行不是第 1 行时样式会错位。
        /// </summary>
        internal void StyleEntryHeader(EliminationTargetHandle target, int startRow, int startColumn)
        {
            WorkbookUi.StyleHeader(target.Sheet, startRow, startColumn, EntryColumns);
        }

        internal string CreateBackup(string savedPath)
        {
            return MergeBackup.Create(savedPath);
        }

        // ---------------------------- 保存 ----------------------------

        /// <summary>带 DisplayAlerts / AskToUpdateLinks 保护保存目标工作簿。</summary>
        internal void SaveTargetBook(object book)
        {
            dynamic a = app;
            bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts");
            bool ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
            try
            {
                a.DisplayAlerts = false;
                a.AskToUpdateLinks = false;
                WorkbookSaving.Save(book, true);
            }
            finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; }
        }

        /// <summary>保存并取回目标工作簿落盘路径（备份用）。</summary>
        internal string SaveTargetBookPath(object book)
        {
            dynamic a = app;
            bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts");
            bool ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
            try
            {
                a.DisplayAlerts = false;
                a.AskToUpdateLinks = false;
                return WorkbookSaving.Save(book, true);
            }
            finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; }
        }

        // ---------------------------- 生命周期 ----------------------------

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { if (gateway != null) gateway.Dispose(); }
            catch { }
            gateway = null;
            createdSheets.Clear();
        }
    }
}
