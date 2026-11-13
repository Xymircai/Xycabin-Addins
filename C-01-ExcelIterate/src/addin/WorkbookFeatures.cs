using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class ComBag : IDisposable
    {
        private readonly List<object> references = new List<object>();
        internal object Own(object value) { if (value != null) references.Add(value); return value; }

        /// <summary>
        /// 把引用移出本袋：调用方将长期持有它，本袋 Dispose 时不得 Release。
        /// 用于「方法内取到、但要返回给上层继续用」的 COM 引用（Workbook / Worksheet）。
        /// 少了这一步，返回的引用会在 using 结束时被 Release，之后再碰就是
        /// “COM 对象与其基础 RCW 分开后就不能再使用”。
        /// </summary>
        internal object Detach(object value)
        {
            if (value != null)
                for (int i = references.Count - 1; i >= 0; i--)
                    if (ReferenceEquals(references[i], value)) { references.RemoveAt(i); break; }
            return value;
        }

        public void Dispose()
        {
            List<string> errors = new List<string>();
            for (int i = references.Count - 1; i >= 0; i--)
                try { if (Marshal.IsComObject(references[i])) Marshal.ReleaseComObject(references[i]); } catch (Exception ex) { errors.Add(ex.Message); }
            references.Clear();
            if (errors.Count != 0) throw new InvalidOperationException("引用释放失败：" + string.Join("；", errors));
        }
    }
    internal sealed class HostQuiet : IDisposable
    {
        private readonly object app;
        private readonly bool events, screen;
        private bool disposed;
        internal HostQuiet(object application)
        {
            app = application;
            events = ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents");
            screen = ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating");
            try { ((dynamic)app).EnableEvents = false; ((dynamic)app).ScreenUpdating = false;
                if (ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents") || ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating")) throw new InvalidOperationException("宿主未接受暂停事件或刷新。"); }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            List<string> errors = new List<string>();
            try { ((dynamic)app).ScreenUpdating = screen; if (ComFormulaHost.Boolean(((dynamic)app).ScreenUpdating, "ScreenUpdating") != screen) throw new Exception("读回不一致"); } catch (Exception ex) { errors.Add("刷新恢复：" + ex.Message); }
            try { ((dynamic)app).EnableEvents = events; if (ComFormulaHost.Boolean(((dynamic)app).EnableEvents, "EnableEvents") != events) throw new Exception("读回不一致"); } catch (Exception ex) { errors.Add("事件恢复：" + ex.Message); }
            if (errors.Count != 0) throw new InvalidOperationException(string.Join("；", errors));
        }
    }
    internal static class WorkbookUi
    {
        internal static string Quote(string name) { return "'" + name.Replace("'", "''") + "'"; }
        internal static string UniqueName(string basis, IEnumerable<string> existing)
        {
            HashSet<string> names = new HashSet<string>(existing, StringComparer.OrdinalIgnoreCase);
            if (!names.Contains(basis)) return basis;
            for (int i = 2; ; i++) { string suffix = "(" + i + ")"; string name = basis.Substring(0, Math.Min(basis.Length, 31 - suffix.Length)) + suffix; if (!names.Contains(name)) return name; }
        }
        internal static object NewSheet(object workbook, string basis, ComBag bag, bool first)
        {
            if (ComFormulaHost.Boolean(((dynamic)workbook).ProtectStructure, "ProtectStructure")) throw new InvalidOperationException("工作簿结构受到保护，不能新增工作表。");
            dynamic sheets = bag.Own(((dynamic)workbook).Sheets); List<string> names = new List<string>();
            for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) { using (ComBag item = new ComBag()) { dynamic s = item.Own(sheets.Item(i)); names.Add((string)s.Name); } }
            dynamic worksheets = bag.Own(((dynamic)workbook).Worksheets);
            object anchor = bag.Own(sheets.Item(first ? 1 : Convert.ToInt32(sheets.Count)));
            dynamic result = bag.Own(first ? worksheets.Add(anchor, Type.Missing, 1, Type.Missing) : worksheets.Add(Type.Missing, anchor, 1, Type.Missing));
            try { result.Name = UniqueName(basis, names); return result; }
            catch { object app = bag.Own(((dynamic)workbook).Application); DeleteSheet(app, (object)result); throw; }
        }
        internal static void DeleteSheet(object app, object sheet)
        {
            bool alerts = ComFormulaHost.Boolean(((dynamic)app).DisplayAlerts, "DisplayAlerts");
            try { ((dynamic)app).DisplayAlerts = false; ((dynamic)sheet).Delete(); }
            finally { ((dynamic)app).DisplayAlerts = alerts; }
        }
        internal static void StyleHeader(object sheet, int startColumn, int columns)
        { StyleHeader(sheet, 1, startColumn, columns); }

        /// <summary>表头行样式（row 为表头所在行，抵消分录可从任意起始单元格写起）。</summary>
        internal static void StyleHeader(object sheet, int row, int startColumn, int columns)
        {
            using (ComBag bag = new ComBag())
            {
                string first = CellAddress.Name(row, startColumn);
                string last = CellAddress.Name(row, startColumn + columns - 1);
                dynamic header = bag.Own(((dynamic)sheet).Range[first + ":" + last]);
                dynamic headFont = bag.Own(header.Font), fill = bag.Own(header.Interior);
                headFont.Bold = true; headFont.Color = 0x473224; fill.Color = 0xFBEFF0; header.RowHeight = 26;
                header.HorizontalAlignment = -4108; header.VerticalAlignment = -4108; header.WrapText = false;
            }
        }
        internal static void Center(object range)
        {
            ((dynamic)range).HorizontalAlignment = -4108;
            ((dynamic)range).VerticalAlignment = -4108;
        }
        private static bool PositiveHeight(object value)
        {
            if (value == null) return false;
            double height = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return !double.IsNaN(height) && !double.IsInfinity(height) && height > 0;
        }
        // Hidden rows can report zero rather than their latent stored height.
        // The caller owns the row reference and has already captured Hidden.
        internal static object ReadRowHeight(object row, bool hidden)
        {
            object height = ((dynamic)row).RowHeight;
            if (PositiveHeight(height)) return height;
            if (!hidden) throw new InvalidOperationException("无法读取有效行高，未开始格式适配。");
            Exception failure = null;
            try
            {
                ((dynamic)row).Hidden = false;
                if (ComFormulaHost.Boolean(((dynamic)row).Hidden, "Hidden")) throw new InvalidOperationException("宿主未接受临时显露输出行。");
                height = ((dynamic)row).RowHeight;
                if (!PositiveHeight(height)) throw new InvalidOperationException("临时显露后仍无法读取有效行高。");
            }
            catch (Exception ex) { failure = ex; }
            finally
            {
                try
                {
                    ((dynamic)row).Hidden = true;
                    if (!ComFormulaHost.Boolean(((dynamic)row).Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                }
                catch (Exception ex) { throw new InvalidOperationException("读取隐藏行高后未能恢复原隐藏状态：" + ex.Message + (failure == null ? "" : "；原始失败：" + failure.Message), failure ?? ex); }
            }
            if (failure != null) throw new InvalidOperationException("读取隐藏行高失败：" + failure.Message, failure);
            return height;
        }
        internal static void ClearSheet(object sheet)
        {
            using (ComBag bag = new ComBag())
            {
                try { ((dynamic)sheet).AutoFilterMode = false; } catch {}
                dynamic cells = bag.Own(((dynamic)sheet).Cells); cells.Clear();
            }
        }
        internal static object EnsureSheet(object workbook, string basis, ComBag bag, bool first)
        {
            dynamic worksheets = bag.Own(((dynamic)workbook).Worksheets);
            for (int i = 1; i <= Convert.ToInt32(worksheets.Count); i++)
            {
                dynamic existing = bag.Own(worksheets.Item(i));
                if (string.Equals(Convert.ToString(existing.Name), basis, StringComparison.OrdinalIgnoreCase))
                {
                    ClearSheet((object)existing);
                    return existing;
                }
            }
            return NewSheet(workbook, basis, bag, first);
        }
        internal static void StyleTable(object sheet, int columns, int rows)
        {
            using (ComBag bag = new ComBag())
            {
                string last = CellAddress.Name(1, columns).TrimEnd('1');
                dynamic data = bag.Own(((dynamic)sheet).Range["A1:" + last + Math.Max(1, rows)]);
                dynamic font = bag.Own(data.Font); font.Name = "Microsoft YaHei UI"; font.Size = 10;
            }
            StyleHeader(sheet, 1, columns);
        }
        // Fits only the product-owned output rectangle, after values, number
        // formats, hyperlinks and fonts are final. Hidden rows are temporarily
        // exposed for measurement without changing the filter or its criteria.
        internal static void FitOutput(object sheet, int startColumn, int columns, int rows)
        {
            if (startColumn < 1 || columns < 1 || rows < 1) throw new ArgumentOutOfRangeException("输出区域必须包含表头。");
            using (ComBag bag = new ComBag())
            {
                string address = CellAddress.Name(1, startColumn) + ":" + CellAddress.Name(rows, startColumn + columns - 1);
                dynamic output = bag.Own(((dynamic)sheet).Range[address]);
                // Complete all snapshots before changing visibility or formats.
                var before = new OutputFormatSnapshot(sheet, (object)output, startColumn, columns, rows);
                Exception failure = null;
                var restoration = new List<string>();
                try
                {
                    try
                    {
                    foreach (int r in before.HiddenRows) using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r, startColumn)]);
                        dynamic row = one.Own(cell.EntireRow); row.Hidden = false;
                        if (ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("宿主未接受临时显露输出行。");
                    }
                    StyleHeader(sheet, startColumn, columns);
                    output.WrapText = false;
                    dynamic widths = bag.Own(output.Columns); widths.AutoFit();
                    bool wrapped = false;
                    for (int c = startColumn; c < startColumn + columns; c++) using (ComBag one = new ComBag())
                    {
                        dynamic column = one.Own(((dynamic)sheet).Range[CellAddress.Name(1, c) + ":" + CellAddress.Name(rows, c)]);
                        double width = Convert.ToDouble((object)column.ColumnWidth, CultureInfo.InvariantCulture);
                        if (double.IsNaN(width) || double.IsInfinity(width) || width <= 0) throw new InvalidOperationException("无法读回有效输出列宽。");
                        // Excel-compatible hosts cap ColumnWidth at 255. Keep the
                        // header on one line; wrap only the data at that limit.
                        if (rows > 1 && width >= 255)
                        {
                            dynamic body = one.Own(((dynamic)sheet).Range[CellAddress.Name(2, c) + ":" + CellAddress.Name(rows, c)]);
                            body.WrapText = true; wrapped = true;
                        }
                    }
                    if (wrapped)
                    {
                        dynamic body = bag.Own(((dynamic)sheet).Range[CellAddress.Name(2, startColumn) + ":" + CellAddress.Name(rows, startColumn + columns - 1)]);
                        dynamic heights = bag.Own(body.Rows); heights.AutoFit();
                    }
                    dynamic header = bag.Own(((dynamic)sheet).Range[CellAddress.Name(1, startColumn) + ":" + CellAddress.Name(1, startColumn + columns - 1)]);
                    dynamic headerRows = bag.Own(header.Rows); headerRows.AutoFit();
                    object headerHeight = header.RowHeight;
                    if (!PositiveHeight(headerHeight)) throw new InvalidOperationException("无法读回有效表头行高。");
                    if (Convert.ToDouble(headerHeight, CultureInfo.InvariantCulture) < 26) header.RowHeight = 26;
                    }
                    catch (Exception ex) { failure = ex; }
                }
                finally
                {
                    if (failure != null) before.Restore(sheet, (object)output, restoration);
                    int previousErrors = restoration.Count;
                    foreach (int r in before.HiddenRows)
                    {
                        try { using (ComBag one = new ComBag())
                        {
                            dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r, startColumn)]);
                            dynamic row = one.Own(cell.EntireRow); row.Hidden = true;
                            if (!ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                        } }
                        catch (Exception ex) { restoration.Add("第" + r + "行：" + ex.Message); }
                    }
                    // Visibility is part of the same format transaction. A failed
                    // finally restore must undo successful measurement as well.
                    // Retain the initial row diagnostic even if the retry works.
                    if (failure == null && restoration.Count > previousErrors)
                    {
                        failure = new InvalidOperationException("隐藏状态恢复失败，已尝试回滚输出格式。");
                        before.Restore(sheet, (object)output, restoration);
                        foreach (int r in before.HiddenRows)
                        {
                            try { using (ComBag one = new ComBag())
                            {
                                dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r, startColumn)]);
                                dynamic row = one.Own(cell.EntireRow); row.Hidden = true;
                                if (!ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) throw new InvalidOperationException("隐藏状态读回不一致");
                            } }
                            catch (Exception ex) { restoration.Add("第" + r + "行再次恢复：" + ex.Message); }
                        }
                    }
                }
                if (restoration.Count != 0) throw new InvalidOperationException("输出格式测量后原状态未完整恢复：" + string.Join("；", restoration) + (failure == null ? "" : "；原始失败：" + failure.Message), failure);
                if (failure != null) throw new InvalidOperationException("输出格式适配失败：" + failure.Message, failure);
            }
        }
        private sealed class OutputFormatSnapshot
        {
            private readonly int first, columns, rows;
            private readonly object uniformWrap;
            private readonly bool[,] cellWrap;
            private readonly object[] widths, heights, bold, fontColor, fillColor, horizontal, vertical;
            internal readonly List<int> HiddenRows = new List<int>();
            internal OutputFormatSnapshot(object sheet, object output, int first, int columns, int rows)
            {
                this.first = first; this.columns = columns; this.rows = rows;
                widths = new object[columns]; heights = new object[rows]; bold = new object[columns];
                fontColor = new object[columns]; fillColor = new object[columns]; horizontal = new object[columns]; vertical = new object[columns];
                uniformWrap = ((dynamic)output).WrapText;
                if (!(uniformWrap is bool)) cellWrap = new bool[rows, columns];
                for (int r = 1; r <= rows; r++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r, first)]);
                    dynamic row = one.Own(cell.EntireRow);
                    if (ComFormulaHost.Boolean((object)row.Hidden, "Hidden")) HiddenRows.Add(r);
                }
                var hiddenSet = new HashSet<int>(HiddenRows);
                for (int r = 1; r <= rows; r++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r, first)]);
                    object row = one.Own(cell.EntireRow); heights[r - 1] = ReadRowHeight(row, hiddenSet.Contains(r));
                }
                for (int c = 0; c < columns; c++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(1, first + c)]);
                    widths[c] = cell.ColumnWidth;
                    dynamic font = one.Own(cell.Font), fill = one.Own(cell.Interior);
                    bold[c] = font.Bold; fontColor[c] = font.Color; fillColor[c] = fill.Color;
                    horizontal[c] = cell.HorizontalAlignment; vertical[c] = cell.VerticalAlignment;
                }
                if (cellWrap != null) for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++) using (ComBag one = new ComBag())
                {
                    dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(r + 1, first + c)]);
                    cellWrap[r, c] = ComFormulaHost.Boolean((object)cell.WrapText, "WrapText");
                }
            }
            private static void Attempt(Action action, string label, List<string> errors)
            { try { action(); } catch (Exception ex) { errors.Add(label + "：" + ex.Message); } }
            internal void Restore(object sheet, object output, List<string> errors)
            {
                if (cellWrap == null) Attempt(delegate { ((dynamic)output).WrapText = uniformWrap; }, "换行", errors);
                else for (int r = 0; r < rows; r++) for (int c = 0; c < columns; c++)
                {
                    int rr = r, cc = c;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    { dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(rr + 1, first + cc)]); cell.WrapText = cellWrap[rr, cc]; } }, "换行" + CellAddress.Name(r + 1, first + c), errors);
                }
                for (int c = 0; c < columns; c++)
                {
                    int cc = c;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(1, first + cc)]);
                        cell.ColumnWidth = widths[cc];
                    } }, "列宽" + (first + c), errors);
                    Attempt(delegate { using (ComBag one = new ComBag())
                    {
                        dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(1, first + cc)]);
                        Attempt(delegate { cell.HorizontalAlignment = horizontal[cc]; }, "表头水平对齐" + (first + cc), errors);
                        Attempt(delegate { cell.VerticalAlignment = vertical[cc]; }, "表头垂直对齐" + (first + cc), errors);
                        Attempt(delegate { dynamic font = one.Own(cell.Font);
                            Attempt(delegate { font.Bold = bold[cc]; }, "表头粗体" + (first + cc), errors);
                            Attempt(delegate { font.Color = fontColor[cc]; }, "表头字色" + (first + cc), errors);
                        }, "表头字体引用" + (first + cc), errors);
                        Attempt(delegate { dynamic fill = one.Own(cell.Interior);fill.Color = fillColor[cc]; }, "表头底色" + (first + cc), errors);
                    } }, "列及表头" + (first + c), errors);
                }
                for (int r = 0; r < rows; r++)
                {
                    int rr = r;
                    Attempt(delegate { using (ComBag one = new ComBag())
                    { dynamic cell = one.Own(((dynamic)sheet).Range[CellAddress.Name(rr + 1, first)]); dynamic row = one.Own(cell.EntireRow); row.RowHeight = heights[rr]; } }, "行高" + (r + 1), errors);
                }
            }
        }
    }

}
