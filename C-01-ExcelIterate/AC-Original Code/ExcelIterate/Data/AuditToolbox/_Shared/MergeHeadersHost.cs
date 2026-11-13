using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    // =====================================================================
    // 跨工作簿表头匹配合并 —— COM 宿主层
    //
    // 纪律（与项目既有实现一致）：
    //   · 全程 dynamic 晚期绑定，不使用 Microsoft.Office.Interop.Excel
    //   · 借用的 Application 永不 Release；自有的 Range/集合一律 ComBag.Own
    //   · 写入在 HostQuiet 内完成，结束后读回校验再恢复事件与刷新
    //   · WPS 实测：批量 Value2 赋值会跳过被筛选隐藏的行，必须逐格补写
    // =====================================================================







    internal static class Marshal2
    {
        internal static bool IsCom(object value)
        { return value != null && System.Runtime.InteropServices.Marshal.IsComObject(value); }
        internal static void Release(object value)
        { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); }
    }

    internal static class MergeSheetReader
    {
        private const int ChunkRows = 4096;

        /// <summary>读取一张表的表头行快照（合并单元格取左上角锚点）。</summary>
        internal static MergeHeaderSet ReadHeaders(object sheet, int headerRow, string path)
        {
            MergeHeaderSet set = new MergeHeaderSet();
            set.Path = path;
            set.HeaderRow = headerRow;
            set.SheetName = Convert.ToString(((dynamic)sheet).Name);
            if (ComFormulaHost.Boolean(((dynamic)sheet).ProtectContents, "ProtectContents"))
                throw new InvalidOperationException("工作表受保护，无法读取：" + set.SheetName);

            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                dynamic used = bag.Own(s.UsedRange);
                int top = Convert.ToInt32(used.Row);
                int left = Convert.ToInt32(used.Column);
                int height = Convert.ToInt32(bag.Own(used.Rows).Count);
                int width = Convert.ToInt32(bag.Own(used.Columns).Count);
                set.LastUsedRow = top + height - 1;
                set.FirstColumn = left;
                set.LastColumn = left + width - 1;
                if (headerRow < top || headerRow > set.LastUsedRow)
                    throw new InvalidOperationException("表头第 " + headerRow + " 行不在数据区（"
                        + top + "～" + set.LastUsedRow + " 行）内。");

                for (int column = left; column <= set.LastColumn; column++)
                {
                    MergeColumn item = new MergeColumn();
                    item.Number = column;
                    item.Letter = CellAddress.ColumnLetters(column);
                    using (ComBag cellBag = new ComBag())
                    {
                        dynamic cell = cellBag.Own(s.Range[CellAddress.Name(headerRow, column)]);
                        dynamic source = cell;
                        bool merged = false;
                        try { merged = ComFormulaHost.Boolean(cell.MergeCells, "MergeCells"); }
                        catch { }
                        if (merged)
                        {
                            dynamic area = cellBag.Own(cell.MergeArea);
                            int r = Convert.ToInt32(area.Row), c = Convert.ToInt32(area.Column);
                            source = cellBag.Own(s.Range[CellAddress.Name(r, c)]);
                        }
                        object value = null;
                        try { value = source.Value2; }
                        catch { }
                        string display = null;
                        try { display = source.Text as string; }
                        catch { }
                        // 列宽过窄时 Text 会返回 "###"，回退到值。
                        if (string.IsNullOrEmpty(display) || display.Trim() == "###")
                            display = value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture);
                        item.Text = display;
                        item.Key = MergeKey.Normalize(display);
                        item.Signature = (value == null ? "null" : value.GetType().FullName + ":" + Convert.ToString(value, CultureInfo.InvariantCulture))
                            + "|" + item.Key;
                    }
                    set.Columns.Add(item);
                }
            }
            set.Seal();
            return set;
        }

        /// <summary>
        /// 目标列已有数据的最后一行；没有数据则返回 headerRow（写入从 headerRow+1 开始）。
        /// 判据：Formula 非空或 Value2 非空。自底向上分块扫描，避免整列读取。
        /// </summary>
        internal static int LastDataRow(object sheet, int column, int headerRow, int lastUsedRow)
        {
            int top = headerRow + 1;
            if (lastUsedRow < top) return headerRow;
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                for (int end = lastUsedRow; end >= top; end -= ChunkRows)
                {
                    int start = Math.Max(top, end - ChunkRows + 1);
                    int height = end - start + 1;
                    using (ComBag chunk = new ComBag())
                    {
                        dynamic range = chunk.Own(s.Range[
                            CellAddress.Name(start, column) + ":" + CellAddress.Name(end, column)]);
                        object formulas = ReadFormulas(range);
                        object values = null;
                        try { values = range.Value2; }
                        catch { }
                        for (int r = height - 1; r >= 0; r--)
                        {
                            if (!IsBlank(At(formulas, r, height, 0)) || !IsBlank(At(values, r, height, 0)))
                                return start + r;
                        }
                    }
                }
            }
            return headerRow;
        }

        /// <summary>读取源列的数据（公式优先，其次值），填进 map.Cells。</summary>
        internal static void ReadColumn(object sheet, MergeColumnMap map, int headerRow)
        {
            map.Cells = new object[map.RowCount];
            map.FormulaCount = 0;
            if (map.RowCount == 0) return;
            int firstRow = map.SourceFirstRow;
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                for (int offset = 0; offset < map.RowCount; offset += ChunkRows)
                {
                    int height = Math.Min(ChunkRows, map.RowCount - offset);
                    using (ComBag chunk = new ComBag())
                    {
                        dynamic range = chunk.Own(s.Range[
                            CellAddress.Name(firstRow + offset, map.SourceColumn) + ":"
                            + CellAddress.Name(firstRow + offset + height - 1, map.SourceColumn)]);
                        // 源区预检：合并单元格与数组公式会让逐格搬运失去意义。
                        object merged = null;
                        try { merged = range.MergeCells; }
                        catch { }
                        if (merged is bool && (bool)merged && height * 1 > 1)
                            throw new InvalidOperationException(map.Header + " 列的数据区含合并单元格，无法逐格搬运。");
                        object arrayFlag = null;
                        try { arrayFlag = range.HasArray; }
                        catch { }
                        if (arrayFlag is bool && (bool)arrayFlag)
                            throw new InvalidOperationException(map.Header + " 列的数据区含数组公式，无法逐格搬运。");

                        object formulas = ReadFormulas(range);
                        object values = null;
                        try { values = range.Value2; }
                        catch { }

                        for (int r = 0; r < height; r++)
                        {
                            object formula = At(formulas, r, height, 0);
                            object value = At(values, r, height, 0);
                            // 判公式不靠 HasFormula（Excel 对它返回标量 bool / DBNull，不可靠），
                            // 直接用 Formula2/Formula 数组元素本身：公式格以 '=' 开头，非公式格是数值/文本/空。
                            string formulaText = formula as string;
                            bool hasFormula = formulaText != null && formulaText.StartsWith("=", StringComparison.Ordinal);
                            if (hasFormula)
                            {
                                map.Cells[offset + r] = formulaText;
                                map.FormulaCount++;
                            }
                            else map.Cells[offset + r] = value;
                        }
                    }
                }
            }
        }

        internal static object ReadFormulas(object range)
        {
            dynamic r = range;
            try { return r.Formula2; }
            catch (Exception ex)
            {
                if (!ComFormulaHost.Missing(ex)) throw;
                try { return r.Formula; }
                catch (Exception inner) { if (!ComFormulaHost.Missing(inner)) throw; return null; }
            }
        }

        // WPS（KET）对单列 Range 的 HasFormula/Formula2/Value2 返回一维数组（object[]），
        // 而 Excel 惯用二维数组（object[,]）。此处两种维度都要兼容，否则公式会被误判为数值。
        internal static object At(object values, int row, int height, int column)
        {
            Array array = values as Array;
            if (array == null) return height == 1 ? values : null;
            if (array.Rank == 1)
            {
                int lo = array.GetLowerBound(0);
                int len = array.GetLength(0);
                if (row < 0 || row >= len) return null;
                return array.GetValue(lo + row);
            }
            if (array.Rank != 2) return null;
            if (row < 0 || row >= array.GetLength(0) || column < 0 || column >= array.GetLength(1)) return null;
            return array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }

        private static bool IsBlank(object value)
        {
            if (value == null) return true;
            string text = value as string;
            return text != null && text.Length == 0;
        }
    }

    /// <summary>写入前的内存快照 —— 「取消操作」的主还原路径。</summary>
    internal sealed class MergeSnapshot
    {
        internal sealed class Region
        {
            internal int FirstColumn;
            internal int ColumnCount;
            internal int StartRow;
            internal int RowCount;
            internal object[,] Values;
            internal object[,] Formulas;
            internal bool[,] HasFormula;
        }
        internal readonly List<Region> Regions = new List<Region>();

        internal static MergeSnapshot Capture(object sheet, MergeWritePlan plan)
        {
            MergeSnapshot snapshot = new MergeSnapshot();
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                foreach (MergeWriteBlock block in plan.Blocks)
                {
                    Region region = new Region();
                    region.FirstColumn = block.FirstColumn;
                    region.ColumnCount = block.ColumnCount;
                    region.StartRow = block.StartRow;
                    region.RowCount = block.RowCount;
                    region.Values = new object[block.RowCount, block.ColumnCount];
                    region.Formulas = new object[block.RowCount, block.ColumnCount];
                    region.HasFormula = new bool[block.RowCount, block.ColumnCount];
                    if (block.RowCount > 0)
                    {
                        dynamic range = bag.Own(s.Range[
                            CellAddress.Name(block.StartRow, block.FirstColumn) + ":"
                            + CellAddress.Name(block.StartRow + block.RowCount - 1,
                                                     block.FirstColumn + block.ColumnCount - 1)]);
                        object formulas = MergeSheetReader.ReadFormulas(range);
                        object values = null;
                        try { values = range.Value2; }
                        catch { }
                        Array fa = formulas as Array, va = values as Array;
                        for (int r = 0; r < block.RowCount; r++)
                            for (int c = 0; c < block.ColumnCount; c++)
                            {
                                region.Values[r, c] = Pick(va, values, r, c, block.RowCount, block.ColumnCount);
                                region.Formulas[r, c] = Pick(fa, formulas, r, c, block.RowCount, block.ColumnCount);
                                // 判公式不靠 HasFormula（Excel 对它返回标量/DBNull，不可靠），
                                // 直接用 Formula 数组元素是否以 '=' 开头判断。
                                object fcell = region.Formulas[r, c];
                                region.HasFormula[r, c] = fcell is string
                                    && ((string)fcell).StartsWith("=", StringComparison.Ordinal);
                            }
                    }
                    snapshot.Regions.Add(region);
                }
            }
            return snapshot;
        }

        internal static object Pick(Array array, object scalar, int row, int column, int rows, int columns)
        {
            if (rows == 1 && columns == 1) return scalar;
            if (array == null) return null;
            // WPS 单列 Range 返回一维数组（object[]），多列才返回二维 object[,]。两种维度都要兼容。
            if (array.Rank == 1)
            {
                int lo = array.GetLowerBound(0);
                int len = array.GetLength(0);
                if (row < 0 || row >= len) return null;
                return array.GetValue(lo + row);
            }
            if (array.Rank != 2) return null;
            if (row < 0 || row >= array.GetLength(0) || column < 0 || column >= array.GetLength(1)) return null;
            return array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }
    }

    internal static class MergeTargetWriter
    {
        /// <summary>
        /// 写入。任一步失败抛异常，由调用方用快照逆序还原。
        /// </summary>
        internal static void Write(object app, object sheet, MergeWritePlan plan, Func<bool> cancelled)
        {
            using (HostQuiet quiet = new HostQuiet(app))
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                bool filtered = false;
                try { filtered = ComFormulaHost.Boolean(s.AutoFilterMode, "AutoFilterMode"); }
                catch { }

                foreach (MergeWriteBlock block in plan.Blocks)
                {
                    if (cancelled != null && cancelled()) throw new OperationCanceledException("已取消，本批未继续写入。");
                    if (block.RowCount == 0) continue;
                    using (ComBag chunk = new ComBag())
                    {
                        dynamic range = chunk.Own(s.Range[
                            CellAddress.Name(block.StartRow, block.FirstColumn) + ":"
                            + CellAddress.Name(block.StartRow + block.RowCount - 1,
                                                     block.FirstColumn + block.ColumnCount - 1)]);

                        // ① 整批写值：公式格先留空，随后逐格写公式。
                        object[,] values = new object[block.RowCount, block.ColumnCount];
                        for (int r = 0; r < block.RowCount; r++)
                            for (int c = 0; c < block.ColumnCount; c++)
                                values[r, c] = MergePlanner.IsFormula(block.Cells[r, c]) ? null : block.Cells[r, c];
                        range.Value2 = values;

                        // ② 隐藏行逐格补写（WPS 批量赋值会跳过筛选隐藏的行）。
                        if (filtered)
                        {
                            for (int r = 0; r < block.RowCount; r++)
                            {
                                int row = block.StartRow + r;
                                bool hidden = false;
                                using (ComBag one = new ComBag())
                                {
                                    dynamic rowRange = one.Own(s.Rows[row]);
                                    try { hidden = ComFormulaHost.Boolean(rowRange.Hidden, "Hidden"); }
                                    catch { }
                                }
                                if (!hidden) continue;
                                for (int c = 0; c < block.ColumnCount; c++)
                                    using (ComBag one = new ComBag())
                                    {
                                        dynamic cell = one.Own(s.Range[
                                            CellAddress.Name(row, block.FirstColumn + c)]);
                                        cell.Value2 = values[r, c];
                                    }
                            }
                        }

                        // ③ 逐格写公式。
                        for (int r = 0; r < block.RowCount; r++)
                            for (int c = 0; c < block.ColumnCount; c++)
                            {
                                object value = block.Cells[r, c];
                                if (!MergePlanner.IsFormula(value)) continue;
                                using (ComBag one = new ComBag())
                                {
                                    dynamic cell = one.Own(s.Range[
                                        CellAddress.Name(block.StartRow + r, block.FirstColumn + c)]);
                                    cell.Formula = Convert.ToString(value, CultureInfo.InvariantCulture);
                                }
                            }

                        // ④ 读回逐格核对。
                        // 注意：工作簿若残留失效外部链接（如引用过不存在的文件），Excel 的
                        // 多格批量读（Value2/Formula 数组）可能返回损坏数据（单格读正常）。
                        // 因此批量核对不一致时，先用单格读复核——单格一致即视为写入成功。
                        object writtenFormulas = MergeSheetReader.ReadFormulas(range);
                        object writtenValues = null;
                        try { writtenValues = range.Value2; }
                        catch { }
                        Array wf = writtenFormulas as Array, wv = writtenValues as Array;
                        for (int r = 0; r < block.RowCount; r++)
                            for (int c = 0; c < block.ColumnCount; c++)
                            {
                                object expected = block.Cells[r, c];
                                if (MergePlanner.IsFormula(expected))
                                {
                                    object actual = MergeSnapshot.Pick(wf, writtenFormulas, r, c, block.RowCount, block.ColumnCount);
                                    if (!SameFormula(actual, expected) && !SingleCellMatches(s, block.StartRow + r, block.FirstColumn + c, expected, true))
                                        throw new InvalidOperationException("公式写入核对失败："
                                            + CellAddress.Name(block.StartRow + r, block.FirstColumn + c)
                                            + "（期望[" + NormalizeFormula(expected) + "] 实际["
                                            + NormalizeFormula(actual) + "]，可据此判断是宿主规范化差异还是写入失败）");
                                }
                                else if (!IsEmpty(expected))
                                {
                                    object actual = MergeSnapshot.Pick(wv, writtenValues, r, c, block.RowCount, block.ColumnCount);
                                    if (!SameValue(actual, expected) && !SingleCellMatches(s, block.StartRow + r, block.FirstColumn + c, expected, false))
                                        throw new InvalidOperationException("数值写入核对失败："
                                            + CellAddress.Name(block.StartRow + r, block.FirstColumn + c));
                                }
                            }
                    }
                }
            }
        }

        /// <summary>用快照整块回写，还原到写入前状态。</summary>
        internal static void Restore(object app, object sheet, MergeSnapshot snapshot)
        {
            using (HostQuiet quiet = new HostQuiet(app))
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                foreach (MergeSnapshot.Region region in snapshot.Regions)
                {
                    if (region.RowCount == 0) continue;
                    using (ComBag chunk = new ComBag())
                    {
                        dynamic range = chunk.Own(s.Range[
                            CellAddress.Name(region.StartRow, region.FirstColumn) + ":"
                            + CellAddress.Name(region.StartRow + region.RowCount - 1,
                                                     region.FirstColumn + region.ColumnCount - 1)]);
                        object[,] values = new object[region.RowCount, region.ColumnCount];
                        for (int r = 0; r < region.RowCount; r++)
                            for (int c = 0; c < region.ColumnCount; c++)
                                values[r, c] = region.HasFormula[r, c] ? null : region.Values[r, c];
                        range.Value2 = values;
                        for (int r = 0; r < region.RowCount; r++)
                            for (int c = 0; c < region.ColumnCount; c++)
                            {
                                if (!region.HasFormula[r, c]) continue;
                                using (ComBag one = new ComBag())
                                {
                                    dynamic cell = one.Own(s.Range[
                                        CellAddress.Name(region.StartRow + r, region.FirstColumn + c)]);
                                    cell.Formula = Convert.ToString(region.Formulas[r, c], CultureInfo.InvariantCulture);
                                }
                            }
                    }
                }
            }
        }

        /// <summary>
        /// 单格读复核。带失效外部链接的工作簿上，Excel 的多格批量读可能损坏，
        /// 单格读可靠——批量核对不一致时用单格读仲裁，一致则视为写入成功。
        /// </summary>
        private static bool SingleCellMatches(dynamic s, int row, int column, object expected, bool isFormula)
        {
            using (ComBag one = new ComBag())
            {
                try
                {
                    dynamic cell = one.Own(s.Range[CellAddress.Name(row, column)]);
                    object actual = isFormula ? MergeSheetReader.ReadFormulas(cell) : cell.Value2;
                    return isFormula ? SameFormula(actual, expected) : SameValue(actual, expected);
                }
                catch { return false; }
            }
        }

        private static bool IsEmpty(object value)
        {
            if (value == null) return true;
            string text = value as string;
            return text != null && text.Length == 0;
        }

        private static bool SameFormula(object actual, object expected)
        {
            string a = NormalizeFormula(actual);
            string b = NormalizeFormula(expected);
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeFormula(object value)
        {
            string text = value == null ? "" : Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
            if (text.StartsWith("=", StringComparison.Ordinal)) text = text.Substring(1);
            text = text.Trim();
            // 归一化跨簿引用：Excel 会把「'完整路径\[文件名]表名'!」简化为「[文件名]表名!」
            // （当引用文件能解析时）。核对时把带路径形态归一化成相对形态，避免路径差异误判。
            text = System.Text.RegularExpressions.Regex.Replace(text, @"'[^']*\[([^\]]+)\]", "[$1]");
            text = text.Replace("'!", "!");
            // 宿主会把相对引用规范化为绝对引用（I3 → $I$3），核对时剥掉 $ 避免误判
            // （实机验收：往来抵消结果表跨表引用 ='差异表'!I3 报「公式写入核对失败：G4」）。
            text = text.Replace("$", "");
            return text.Trim();
        }

        private static bool SameValue(object actual, object expected)
        {
            if (actual == null && expected == null) return true;
            if (actual == null || expected == null) return false;
            if (actual is double && expected is double)
                return Math.Abs((double)actual - (double)expected) < 0.0000001;
            return string.Equals(Convert.ToString(actual, CultureInfo.InvariantCulture),
                                 Convert.ToString(expected, CultureInfo.InvariantCulture),
                                 StringComparison.Ordinal);
        }
    }

    internal static class MergeBackup
    {
        /// <summary>文件级备份（兜底）。写入前目标须已保存落盘。</summary>
        internal static string Create(string targetPath)
        {
            string directory = Path.Combine(Path.GetDirectoryName(targetPath), "ExcelIterate合并备份");
            Directory.CreateDirectory(directory);
            string name = Path.GetFileNameWithoutExtension(targetPath)
                + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff", CultureInfo.InvariantCulture)
                + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".xlsx";
            string destination = Path.Combine(directory, name);
            File.Copy(targetPath, destination, false);
            string original = Hash(targetPath);
            string copy = Hash(destination);
            if (original != copy) throw new IOException("备份与目标文件核对不一致，已中止：" + destination);
            return destination;
        }

        internal static string Hash(string path)
        {
            // 目标工作簿此时已被 WPS/Excel 打开并持有写锁；必须用 ReadWrite 共享才能再开一份读。
            // WPS 刚 b.Save() 完句柄释放也有过渡期；这里短暂重试，避免假失败把整批写入带进回滚。
            Exception last = null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (SHA256 sha = SHA256.Create())
                        return BitConverter.ToString(sha.ComputeHash(stream));
                }
                catch (IOException ex) { last = ex; if (attempt < 2) System.Threading.Thread.Sleep(50); }
            }
            throw new IOException("读取文件校验失败：" + path + "（" + (last == null ? "" : last.Message) + "）", last);
        }
    }
}
