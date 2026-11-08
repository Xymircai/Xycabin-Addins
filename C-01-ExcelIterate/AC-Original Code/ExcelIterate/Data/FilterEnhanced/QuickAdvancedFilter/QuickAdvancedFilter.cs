using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class FilterCriterion
    {
        internal int Row, Column;
        internal object Value;
    }

    internal static class QuickFilterCriteria
    {
        // Physical row identity defines AND, distinct rows define OR. A union
        // across columns must not create a Cartesian product of selected values.
        internal static object[,] Build(string[] headers, IList<FilterCriterion> cells)
        {
            SortedDictionary<int, Dictionary<int, object>> rows = new SortedDictionary<int, Dictionary<int, object>>();
            SortedSet<int> columns = new SortedSet<int>();
            foreach (FilterCriterion cell in cells)
            {
                if (cell.Value is ErrorWrapper || (cell.Value is int && (((uint)(int)cell.Value & 0xFFFF0000u) == 0x800A0000u)) || (cell.Value is uint && (((uint)cell.Value & 0xFFFF0000u) == 0x800A0000u))) throw new InvalidOperationException("条件不能包含错误值。");
                if (cell.Column < 0 || cell.Column >= headers.Length) throw new InvalidOperationException("条件列不在数据区域内。");
                Dictionary<int, object> row;
                if (!rows.TryGetValue(cell.Row, out row)) rows.Add(cell.Row, row = new Dictionary<int, object>());
                row[cell.Column] = cell.Value; columns.Add(cell.Column);
            }
            List<Dictionary<int, object>> effective = new List<Dictionary<int, object>>();
            foreach (Dictionary<int, object> row in rows.Values)
            {
                bool any = false; foreach (object value in row.Values) if (value != null && Convert.ToString(value, CultureInfo.InvariantCulture).Length != 0) any = true;
                if (any) effective.Add(row); // An entirely blank OR row would match everything.
            }
            if (effective.Count == 0) throw new InvalidOperationException("所选条件均为空，请选择有内容的数据单元格。");
            object[,] result = new object[effective.Count + 1, columns.Count]; int c = 0;
            foreach (int col in columns)
            {
                result[0, c] = headers[col];
                for (int r = 0; r < effective.Count; r++) { object value; if (effective[r].TryGetValue(col, out value)) result[r + 1, c] = value; }
                c++;
            }
            return result;
        }
    }

    internal static class QuickAdvancedFilter
    {
        internal static int Run(object app)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic window = bag.Own(((dynamic)app).ActiveWindow), groups = bag.Own(window.SelectedSheets);
                if (Convert.ToInt32(groups.Count) != 1) throw new InvalidOperationException("请先取消工作表组合选择。");
                dynamic sheet = bag.Own(((dynamic)app).ActiveSheet), selection = bag.Own(((dynamic)app).Selection);
                if (sheet == null || selection == null) throw new InvalidOperationException("请选中数据行中的条件单元格。");
                using (WorksheetIdentity identity = WorksheetIdentity.Capture((object)sheet))
                {
                    object selectedSheet = bag.Own(selection.Worksheet);
                    if (!identity.Matches(selectedSheet)) throw new InvalidOperationException("请在当前工作表中选择条件。");
                    if (ComFormulaHost.Boolean((object)sheet.ProtectContents, "ProtectContents")) throw new InvalidOperationException("请先取消工作表保护，再取得连续数据区域。");
                    dynamic areas = bag.Own(selection.Areas), first = bag.Own(areas.Item(1));
                    dynamic firstCell = bag.Own(sheet.Range[CellAddress.Name(Convert.ToInt32(first.Row), Convert.ToInt32(first.Column))]);
                    dynamic source = bag.Own(firstCell.CurrentRegion);
                    int top = Convert.ToInt32(source.Row), left = Convert.ToInt32(source.Column);
                    dynamic sourceRows = bag.Own(source.Rows), sourceColumns = bag.Own(source.Columns);
                    int height = Convert.ToInt32(sourceRows.Count), width = Convert.ToInt32(sourceColumns.Count);
                    if (height < 2) throw new InvalidOperationException("条件应位于带表头的连续数据表内。");
                    string[] headers = new string[width]; HashSet<string> unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int c = 0; c < width; c++)
                    {
                        using (ComBag cellBag = new ComBag())
                        {
                            dynamic cell = cellBag.Own(sheet.Range[CellAddress.Name(top, left + c)]);
                            if (ComFormulaHost.Boolean((object)cell.MergeCells, "MergeCells")) throw new InvalidOperationException("表头不能合并。");
                            headers[c] = Convert.ToString((object)cell.Value2, CultureInfo.InvariantCulture);
                            if (string.IsNullOrWhiteSpace(headers[c]) || !unique.Add(headers[c])) throw new InvalidOperationException("数据表第一行需要非空且不重复的列名。");
                        }
                    }
                    List<FilterCriterion> conditions = new List<FilterCriterion>();
                    for (int a = 1; a <= Convert.ToInt32(areas.Count); a++)
                    {
                        using (ComBag areaBag = new ComBag())
                        {
                            dynamic area = areaBag.Own(areas.Item(a)), rr = areaBag.Own(area.Rows), cc = areaBag.Own(area.Columns);
                            int ar = Convert.ToInt32(area.Row), ac = Convert.ToInt32(area.Column), rh = Convert.ToInt32(rr.Count), cw = Convert.ToInt32(cc.Count);
                            if (ar <= top || ar + rh > top + height || ac < left || ac + cw > left + width) throw new InvalidOperationException("条件必须位于同一连续数据表的数据行内，不能含表头或跨表格。");
                            if ((long)conditions.Count + (long)rh * cw > 100000) throw new InvalidOperationException("请选择需要的条件行，条件选区不能超过十万个单元格。");
                            for (int r = ar; r < ar + rh; r++) for (int c = ac; c < ac + cw; c++)
                                using (ComBag one = new ComBag()) { dynamic cell = one.Own(sheet.Range[CellAddress.Name(r, c)]); if (ComFormulaHost.Boolean((object)cell.MergeCells, "MergeCells")) throw new InvalidOperationException("条件单元格不能合并。"); object v = cell.Value2; if (v is ErrorWrapper) throw new InvalidOperationException("条件不能包含错误值。"); conditions.Add(new FilterCriterion { Row = r, Column = c - left, Value = v }); }
                        }
                    }
                    object[,] criteria = QuickFilterCriteria.Build(headers, conditions);
                    dynamic book = bag.Own(sheet.Parent); object output = null, temp = null; bool completed = false;
                    using (HostQuiet quiet = new HostQuiet(app))
                    {
                        try
                        {
                            output = WorkbookUi.NewSheet((object)book, "高筛结果", bag, false);
                            temp = WorkbookUi.NewSheet((object)book, "_EI高筛条件", bag, false);
                            dynamic criterionRange = bag.Own(((dynamic)temp).Range["A1:" + CellAddress.Name(criteria.GetLength(0), criteria.GetLength(1))]);
                            criterionRange.NumberFormat = "@"; criterionRange.Value2 = criteria;
                            dynamic destination = bag.Own(((dynamic)output).Range["A1:" + CellAddress.Name(1, width)]);
                            object[,] head = new object[1, width]; for (int c = 0; c < width; c++) head[0, c] = headers[c]; destination.Value2 = head;
                            ((dynamic)output).Activate();
                            // WPS/Excel's native engine defines comparisons, dates,
                            // wildcard escaping and text criteria semantics.
                            source.AdvancedFilter(2, criterionRange, destination, false);
                            dynamic used = bag.Own(((dynamic)output).UsedRange), rows = bag.Own(used.Rows);
                            int count = Math.Max(0, Convert.ToInt32(rows.Count) - 1);
                            object values = used.Value2; used.NumberFormat = "@"; used.Value2 = values;
                            for (int c = 0; c < width; c++) using (ComBag one = new ComBag())
                            {
                                dynamic original = one.Own(sheet.Range[CellAddress.Name(top + 1, left + c)]);
                                dynamic target = one.Own(((dynamic)output).Range[CellAddress.Name(2, c + 1) + ":" + CellAddress.Name(Math.Max(2, count + 1), c + 1)]);
                                target.NumberFormat = original.NumberFormat;
                            }
                            WorkbookUi.StyleTable(output, width, count + 1);
                            WorkbookUi.FitOutput(output, 1, width, count + 1);
                            WorkbookUi.DeleteSheet(app, temp); temp = null; completed = true; return count;
                        }
                        finally
                        {
                            List<string> failures = new List<string>();
                            if (temp != null) try { WorkbookUi.DeleteSheet(app, temp); } catch (Exception ex) { failures.Add("临时条件表清理失败：" + ex.Message); }
                            if (!completed && output != null) try { WorkbookUi.DeleteSheet(app, output); } catch (Exception ex) { failures.Add("未完成结果表清理失败：" + ex.Message); }
                            if (!completed) try { sheet.Activate(); selection.Select(); } catch (Exception ex) { failures.Add("选区恢复失败：" + ex.Message); }
                            if (failures.Count > 0) throw new InvalidOperationException(string.Join("；", failures));
                        }
                    }
                }
            }
        }
    }
}
