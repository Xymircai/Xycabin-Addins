using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Excel_Iterate
{
    // Own a reference to the explicitly chosen workbook, never the active workbook.
    // Browsed workbooks are left open: this feature changes them without auto-saving.
    internal sealed class SummaryWorkbookBinding : IDisposable
    {
        private readonly object application;
        private ComBag held = new ComBag();
        private object selected;
        internal object Application { get { return application; } }
        internal string Path { get; private set; }
        internal long Version { get; private set; }
        internal bool OpenedByBrowse { get; private set; }
        internal SummaryWorkbookBinding(object app) { application = app; Path = ""; }
        internal List<SummaryWorkbookChoice> Choices()
        {
            var choices = new List<SummaryWorkbookChoice>();
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)application).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic book = bag.Own(books.Item(i));
                    choices.Add(new SummaryWorkbookChoice { Path = Convert.ToString(book.FullName), Name = Convert.ToString(book.Name) });
                }
            }
            return choices;
        }
        private object Find(string path, ComBag bag)
        {
            dynamic books = bag.Own(((dynamic)application).Workbooks);
            for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
            {
                object book = bag.Own(books.Item(i));
                if (string.Equals(Convert.ToString(((dynamic)book).FullName), path, StringComparison.OrdinalIgnoreCase)) return book;
            }
            return null;
        }
        private void Bind(object book, ComBag bag)
        {
            if (book == null) throw new InvalidOperationException("所选工作簿已关闭，请刷新列表后重新选择。");
            if (Convert.ToBoolean(((dynamic)book).ReadOnly)) throw new InvalidOperationException("所选工作簿为只读，无法生成汇总表；请选择可编辑的工作簿。");
            string path = Convert.ToString(((dynamic)book).FullName);
            Clear(); selected = held.Own(bag.Detach(book)); Path = path;
        }
        internal void Choose(string path)
        {
            using (var bag = new ComBag()) Bind(Find(path, bag), bag);
        }
        internal void UseCurrent()
        {
            using (var bag = new ComBag()) Bind(bag.Own(((dynamic)application).ActiveWorkbook), bag);
        }
        internal bool Browse(string path)
        {
            string full = System.IO.Path.GetFullPath(path);
            using (var bag = new ComBag())
            {
                object existing = Find(full, bag);
                if (existing != null) { Bind(existing, bag); return false; }
                dynamic app = application;
                bool alerts = Convert.ToBoolean(app.DisplayAlerts), ask = Convert.ToBoolean(app.AskToUpdateLinks);
                object security = app.AutomationSecurity;
                object opened = null;
                var restore = new List<Action>();
                using (var quiet = new HostQuiet(application))
                {
                    try
                    {
                        restore.Add(delegate { app.DisplayAlerts = alerts; }); app.DisplayAlerts = false;
                        restore.Add(delegate { app.AskToUpdateLinks = ask; }); app.AskToUpdateLinks = false;
                        restore.Add(delegate { app.AutomationSecurity = (dynamic)security; }); app.AutomationSecurity = 3;
                        dynamic books = bag.Own(app.Workbooks);
                        opened = bag.Own(books.Open(full, 0, false));
                        try { Bind(opened, bag); OpenedByBrowse = true; }
                        catch { if (opened != null) ((dynamic)opened).Close(false); throw; }
                    }
                    finally
                    {
                        var failures = new List<string>();
                        for (int i = restore.Count - 1; i >= 0; i--) try { restore[i](); } catch (Exception error) { failures.Add(error.Message); }
                        if (failures.Count > 0) throw new InvalidOperationException("打开工作簿后恢复宿主状态失败：" + string.Join("；", failures.ToArray()));
                    }
                }
                return true;
            }
        }
        internal object RequireBook()
        {
            if (selected == null) throw new InvalidOperationException("请先选择工作簿，或点击“使用当前工作簿”。");
            using (var bag = new ComBag())
            {
                object current = Find(Path, bag);
                if (!WorksheetIdentity.SameObject(selected, current))
                    throw new InvalidOperationException("所选工作簿已关闭、重新打开或另存为其他路径。请明确重新选择并预检，不能改用当前工作簿。");
                if (Convert.ToBoolean(((dynamic)current).ReadOnly)) throw new InvalidOperationException("所选工作簿已变为只读。");
            }
            return selected; // Borrowed from this binding; caller must not release it.
        }
        internal void Clear()
        {
            selected = null; Path = ""; Version++; OpenedByBrowse = false;
            held.Dispose(); held = new ComBag();
        }
        public void Dispose() { Clear(); }
    }
    // =====================================================================
    // 单体汇总批量链接 —— COM 宿主层
    //
    // 纪律（与项目既有实现一致）：
    //   · 全程 dynamic 晚期绑定，不引用 Microsoft.Office.Interop.Excel
    //   · 自建的 Range/集合一律 ComBag.Own，借用的 Application 永不 Release
    //   · 写入包在 HostQuiet（关事件 + 关刷新）与手动重算里，结束读回校验再恢复
    //   · WPS 对单列 Range 可能返回一维 object[]，多列才是 object[,]，两种都要兼容
    // =====================================================================

    /// <summary>Lead 表读取结果（表头列位置、候选公司、选用列位置）。</summary>
    internal sealed class SummaryLeadSnapshot
    {
        internal int HeaderRow;
        internal int CodeColumn;      // 0 = 没有这一列
        internal int NameColumn;
        internal int SourceFileColumn;
        internal int PickColumn;      // 0 = 没有这一列
        internal int LastColumn;      // 表头行最后一个非空列
        internal int PickWriteColumn; // 回写「选用」的列（可能是新建的列）
        internal bool PickCreated;
        internal int LastRow;
        internal readonly List<SummaryCompany> Companies = new List<SummaryCompany>();
        internal readonly List<string> Blocking = new List<string>();
    }

    /// <summary>模板列体检结果。</summary>
    internal sealed class SummaryTemplateProbe
    {
        internal readonly List<string> Formulas = new List<string>();
        internal int HeaderRow;
        internal int LastRow;                 // 模板列最后一个非空行
        internal int LastColumn;              // 目标区间最后一列
        internal int FormulaCells;
        internal bool Placeholder;            // 表头行是该列的表头占位
        internal string ExpectedPlaceholder = "";  // 本次要求匹配的占位名（空=不校验）
        internal string PlaceholderText = "";
        internal readonly List<string> Tokens = new List<string>();
        internal readonly List<string> CrossColumnMerges = new List<string>();
        internal readonly List<string> Blocking = new List<string>();
        internal readonly List<string> Warnings = new List<string>();
    }

    internal sealed class SummaryRunReport
    {
        internal int Columns;
        internal int CopiedColumns;
        internal int FormulaCells;
        internal int RewrittenCells;
        internal int RewrittenLinks;
        internal int SkippedColumns;
        internal string BackupSheet = "";
        internal string LogSheet = "";
        internal readonly List<string> Warnings = new List<string>();
        internal readonly List<SummaryLinkSample> Samples = new List<SummaryLinkSample>();
    }

    /// <summary>数组取值兼容：WPS 单列返回一维、多列返回二维。</summary>
    internal static class SummaryArrays
    {
        internal static object Pick(object values, int row, int rows, int column, int columns)
        {
            Array array = values as Array;
            if (array == null) return rows == 1 && columns == 1 ? values : null;
            if (array.Rank == 1)
            {
                int low = array.GetLowerBound(0);
                if (row < 0 || row >= array.GetLength(0)) return null;
                return array.GetValue(low + row);
            }
            if (array.Rank != 2) return null;
            if (row < 0 || row >= array.GetLength(0) || column < 0 || column >= array.GetLength(1)) return null;
            return array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1));
        }

        internal static string Text(object value)
        {
            if (value == null || value is DBNull) return "";
            string text = value as string;
            if (text != null) return text.Trim();
            if (value is double || value is int || value is long || value is decimal || value is float)
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            if (value is DateTime) return ((DateTime)value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
        }

        internal static bool HasContent(object value)
        {
            if (value == null || value is DBNull) return false;
            string text = value as string;
            return text == null || text.Trim().Length > 0;
        }
    }

    /// <summary>Lead 表读取：动态定位「编码&amp;简称」「全称」「选用」三列，不做列号硬编码。</summary>
    internal static class SummaryLeadReader
    {
        internal static SummaryLeadSnapshot Read(object sheet, int headerRow)
        {
            SummaryLeadSnapshot snapshot = new SummaryLeadSnapshot();
            snapshot.HeaderRow = headerRow;
            dynamic s = sheet;
            if (ComFormulaHost.Boolean(s.ProtectContents, "ProtectContents"))
            {
                snapshot.Blocking.Add("Lead 工作表受保护，无法读取或回写「选用」列。");
                return snapshot;
            }

            using (ComBag bag = new ComBag())
            {
                dynamic used = bag.Own(s.UsedRange);
                int top = Convert.ToInt32(used.Row, CultureInfo.InvariantCulture);
                int left = Convert.ToInt32(used.Column, CultureInfo.InvariantCulture);
                int height = Convert.ToInt32(bag.Own(used.Rows).Count, CultureInfo.InvariantCulture);
                int width = Convert.ToInt32(bag.Own(used.Columns).Count, CultureInfo.InvariantCulture);
                int bottom = top + height - 1;
                int lastColumn = left + width - 1;
                if (bottom < headerRow)
                {
                    snapshot.Blocking.Add("Lead 表在第 " + headerRow.ToString(CultureInfo.InvariantCulture) + " 行没有表头。");
                    return snapshot;
                }

                // 表头行：整行读一次，按归一化后的名字定位列
                dynamic headerRange = bag.Own(s.Range[
                    CellAddress.Name(headerRow, left) + ":"
                    + CellAddress.Name(headerRow, lastColumn)]);
                object headerValues = ReadValues(headerRange);
                int columns = lastColumn - left + 1;
                int headerLast = 0;
                for (int c = 0; c < columns; c++)
                {
                    string text = SummaryArrays.Text(SummaryArrays.Pick(headerValues, 0, 1, c, columns));
                    if (SummaryHeaders.Normalize(text) == SummaryHeaders.Normalize(SummaryFileMapping.Header))
                    {
                        if (snapshot.SourceFileColumn != 0) snapshot.Blocking.Add("Lead 表「来源文件名」表头重复，请仅保留一列。");
                        snapshot.SourceFileColumn = left + c;
                    }
                    if (text.Length > 0) headerLast = left + c;
                    if (snapshot.CodeColumn == 0 && SummaryHeaders.IsCode(text)) snapshot.CodeColumn = left + c;
                    else if (snapshot.NameColumn == 0 && SummaryHeaders.IsName(text)) snapshot.NameColumn = left + c;
                    else if (snapshot.PickColumn == 0 && SummaryHeaders.IsPick(text)) snapshot.PickColumn = left + c;
                }
                snapshot.LastColumn = headerLast;
                snapshot.LastRow = bottom;

                if (snapshot.CodeColumn == 0) snapshot.Blocking.Add("Lead 表第 " + headerRow.ToString(CultureInfo.InvariantCulture) + " 行找不到表头「" + SummaryHeaders.CodeHeader + "」。");
                if (snapshot.NameColumn == 0) snapshot.Blocking.Add("Lead 表第 " + headerRow.ToString(CultureInfo.InvariantCulture) + " 行找不到表头「" + SummaryHeaders.NameHeader + "」。");
                if (snapshot.Blocking.Count > 0) return snapshot;

                // 「选用」列不存在时，在表头行最右一列的右边补一列（不硬编码列号）
                snapshot.PickWriteColumn = snapshot.PickColumn > 0 ? snapshot.PickColumn : snapshot.LastColumn + 1;
                snapshot.PickCreated = snapshot.PickColumn == 0;
                if (snapshot.PickWriteColumn > 16384)
                {
                    snapshot.Blocking.Add("Lead 表已到列上限，无法新建「选用」列。");
                    return snapshot;
                }

                int rows = bottom - headerRow;
                if (rows <= 0) return snapshot;
                object codeValues = ReadColumn(s, headerRow + 1, bottom, snapshot.CodeColumn, bag);
                object nameValues = ReadColumn(s, headerRow + 1, bottom, snapshot.NameColumn, bag);
                object fileValues = snapshot.SourceFileColumn > 0 ? ReadColumn(s, headerRow + 1, bottom, snapshot.SourceFileColumn, bag) : null;
                object pickValues = snapshot.PickColumn > 0
                    ? ReadColumn(s, headerRow + 1, bottom, snapshot.PickColumn, bag)
                    : null;

                for (int r = 0; r < rows; r++)
                {
                    string code = SummaryArrays.Text(SummaryArrays.Pick(codeValues, r, rows, 0, 1));
                    string name = SummaryArrays.Text(SummaryArrays.Pick(nameValues, r, rows, 0, 1));
                    string prior = pickValues == null ? "" : SummaryArrays.Text(SummaryArrays.Pick(pickValues, r, rows, 0, 1));
                    if (code.Length == 0 || name.Length == 0) continue;   // 候选行 = 两列均非空
                    SummaryCompany company = new SummaryCompany();
                    company.Row = headerRow + 1 + r;
                    company.Code = code;
                    company.Name = name;
                    company.SourceFileName = fileValues == null ? "" : Convert.ToString(SummaryArrays.Pick(fileValues, r, rows, 0, 1), CultureInfo.InvariantCulture);
                    company.PriorSelection = prior;
                    company.Selected = true;                              // 候选行预填勾
                    snapshot.Companies.Add(company);
                }
                if (snapshot.Companies.Count == 0)
                    snapshot.Blocking.Add("Lead 表没有候选行（「" + SummaryHeaders.CodeHeader + "」与「"
                        + SummaryHeaders.NameHeader + "」同时非空的行）。");
            }
            return snapshot;
        }

        private static object ReadValues(object range)
        {
            try { return ((dynamic)range).Value2; }
            catch (Exception ex) { if (!ComFormulaHost.Missing(ex)) throw; return null; }
        }

        private static object ReadColumn(object sheet, int firstRow, int lastRow, int column, ComBag bag)
        {
            if (lastRow < firstRow || column < 1) return null;
            dynamic range = bag.Own(((dynamic)sheet).Range[
                CellAddress.Name(firstRow, column) + ":" + CellAddress.Name(lastRow, column)]);
            return ReadValues(range);
        }

        internal static void ApplySelection(SummaryLeadSnapshot snapshot, IList<SummaryCompany> selected)
        {
            var rows = new HashSet<int>(); foreach (var company in selected) rows.Add(company.Row);
            foreach (var company in snapshot.Companies) company.Selected = rows.Contains(company.Row);
        }
        /// <summary>按实际来源行回写，非候选行不写，年度重复列不重复回写。</summary>
        internal static void WriteSelection(object sheet, SummaryLeadSnapshot snapshot)
        {
            if (snapshot.Companies.Count == 0) return;
            dynamic s = sheet;
            using (ComBag bag = new ComBag())
            {
                if (snapshot.PickCreated)
                {
                    dynamic header = bag.Own(s.Cells[snapshot.HeaderRow, snapshot.PickWriteColumn]);
                    header.Value2 = SummaryHeaders.PickHeader;
                }
                foreach (var company in snapshot.Companies)
                    using (var one = new ComBag())
                    {
                        dynamic cell = one.Own(s.Cells[company.Row, snapshot.PickWriteColumn]);
                        cell.Value2 = company.Selected;
                    }
            }
        }
    }

    /// <summary>模板列体检：占位表头、末行、外链文件名、跨列合并。</summary>
    internal static class SummaryTemplateReader
    {
        internal static void ValidateRelativeFormulas(SummaryRequest request, SummaryTemplateProbe probe)
        {
            // Check the farthest destination before touching the workbook. Intermediate
            // column offsets are nonnegative and cannot overflow if this one is valid.
            var columns = request.Columns();
            int delta = columns.Count == 0 ? 0 : columns[columns.Count - 1].Column - request.TemplateColumn;
            foreach (string formula in probe.Formulas) SummaryFormulaShift.Shift(formula, delta);
        }
        internal static List<string> AnnualPreview(object book, object sheet, SummaryRequest request, SummaryTemplateProbe probe)
        {
            var lines = new List<string>();
            if (!request.Annual) return lines;
            SummaryAnnualLinks.Prepare(request, probe.Formulas);
            using (var bag = new ComBag())
            {
                dynamic s = sheet;
                dynamic rows = bag.Own(s.Rows);
                if (request.YearHeaderRow > Convert.ToInt32(rows.Count)) throw new InvalidOperationException("年度表头行超出当前工作表范围。");
                var columns = request.Columns();
                for (int column = request.TemplateColumn; column < request.TemplateColumn + columns.Count; column++)
                    foreach (int row in new[] { request.SheetHeaderRow, request.YearHeaderRow })
                    {
                        dynamic cell = bag.Own(s.Cells[row, column]);
                        if (Convert.ToBoolean(cell.MergeCells) || Convert.ToBoolean(cell.HasArray))
                            throw new InvalidOperationException("公司和年度表头须为独立单元格，不能合并或含数组公式：" + SummaryColumns.Letter(column) + row);
                    }
                probe.LastRow = Math.Max(probe.LastRow, request.YearHeaderRow);
                lines.Add("年度目录：" + request.YearPathPattern + "；模板年度：" + request.TemplateYear + "；年度表头：第 " + request.YearHeaderRow + " 行。");
                var checkedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var column in columns)
                    foreach (string path in SummaryAnnualLinks.Paths(probe.Formulas, request, column))
                    {
                        string full = path;
                        if (request.CheckYearFiles)
                        {
                            if (!System.IO.Path.IsPathRooted(full))
                            {
                                string folder = System.IO.Path.GetDirectoryName(Convert.ToString(((dynamic)book).FullName));
                                if (string.IsNullOrEmpty(folder)) throw new InvalidOperationException("相对来源路径需要已保存的目标工作簿，或取消来源文件检查。");
                                full = System.IO.Path.GetFullPath(System.IO.Path.Combine(folder, full));
                            }
                            if (checkedPaths.Add(full) && !System.IO.File.Exists(full)) throw new InvalidOperationException("年度来源文件不存在或当前无法访问：" + full);
                        }
                        lines.Add(SummaryColumns.Letter(column.Column) + "｜" + column.Company.Code + "｜" + column.YearLabel + "｜" + full);
                    }
                if (!request.CheckYearFiles) lines.Add("已关闭年度来源文件检查：仅生成链接，不代表来源文件存在或可读取。");
            }
            return lines;
        }
        internal static SummaryTemplateProbe Read(object sheet, int headerRow, int firstColumn, int lastColumn, string placeholder)
        {
            SummaryTemplateProbe probe = new SummaryTemplateProbe();
            probe.HeaderRow = headerRow;
            probe.ExpectedPlaceholder = placeholder == null ? "" : placeholder.Trim();
            probe.LastColumn = lastColumn;
            dynamic s = sheet;
            if (ComFormulaHost.Boolean(s.ProtectContents, "ProtectContents"))
            {
                probe.Blocking.Add("汇总工作表受保护，无法写入。");
                return probe;
            }

            using (ComBag bag = new ComBag())
            {
                dynamic used = bag.Own(s.UsedRange);
                int top = Convert.ToInt32(used.Row, CultureInfo.InvariantCulture);
                int height = Convert.ToInt32(bag.Own(used.Rows).Count, CultureInfo.InvariantCulture);
                int bottom = top + height - 1;
                if (bottom < headerRow) bottom = headerRow;

                // 整列从第 1 行读起：表头行可能不在第 1 行（真实底稿常见第 2 行），
                // 但它影响的是「表头写哪一行」，不该影响复制范围。
                string letter = SummaryColumns.Letter(firstColumn);
                dynamic columnRange = bag.Own(s.Range[letter + "1:" + letter + bottom.ToString(CultureInfo.InvariantCulture)]);
                object values = ReadValues(columnRange);
                object formulas = ReadFormulas(columnRange);
                int rows = bottom;

                int lastRow = headerRow - 1;
                for (int r = rows - 1; r >= 0; r--)
                {
                    if (SummaryArrays.HasContent(SummaryArrays.Pick(values, r, rows, 0, 1)) ||
                        SummaryArrays.HasContent(SummaryArrays.Pick(formulas, r, rows, 0, 1)))
                    {
                        lastRow = 1 + r;
                        break;
                    }
                }
                probe.LastRow = lastRow;

                for (int r = 0; r < rows; r++)
                {
                    object formula = SummaryArrays.Pick(formulas, r, rows, 0, 1);
                    string text = formula as string;
                    if (text == null || text.Length == 0 || text[0] != '=') continue;
                    probe.FormulaCells++;
                    probe.Formulas.Add(text);
                    SummaryLinkRewriter.Tokens(text, probe.Tokens);
                }

                dynamic headerCell = bag.Own(s.Cells[headerRow, firstColumn]);
                probe.PlaceholderText = SummaryArrays.Text(headerCell.Value2);
                // 占位名可配：留空表示不校验；否则按去空白/全角归一后的文字比较
                probe.Placeholder = SummaryHeaders.MatchesPlaceholder(probe.PlaceholderText, probe.ExpectedPlaceholder);

                // 第一行必须是逐列独立单元格：任何跨列合并都终止
                for (int c = firstColumn; c <= lastColumn; c++)
                {
                    dynamic cell = bag.Own(s.Cells[headerRow, c]);
                    if (!ComFormulaHost.Boolean(cell.MergeCells, "MergeCells")) continue;
                    dynamic area = bag.Own(cell.MergeArea);
                    int span = Convert.ToInt32(bag.Own(area.Columns).Count, CultureInfo.InvariantCulture);
                    if (span > 1)
                        probe.CrossColumnMerges.Add(SummaryColumns.Letter(c) + headerRow.ToString(CultureInfo.InvariantCulture)
                            + "（跨 " + span.ToString(CultureInfo.InvariantCulture) + " 列）");
                }
            }

            if (!probe.Placeholder)
            {
                probe.Blocking.Add("模板列 " + SummaryColumns.Letter(firstColumn) + " 第 "
                    + headerRow.ToString(CultureInfo.InvariantCulture) + " 行的值是「" + probe.PlaceholderText
                    + "」，与占位名「" + probe.ExpectedPlaceholder + "」不一致。"
                    + (probe.PlaceholderText.Length > 0
                        ? "若这一格就是本表模板列的表头，再点一次「读取 Lead 并预检」会自动把占位名改成它；也可以在「占位名」里自己改。"
                        : "这一格是空的：把占位名清空即可不校验，直接写入表头。"));
            }
            if (probe.CrossColumnMerges.Count > 0)
            {
                probe.Blocking.Add("表头行检测到跨列合并单元格：" + string.Join("、", probe.CrossColumnMerges.ToArray())
                    + "。表头必须是逐列独立单元格，已终止。");
            }
            if (probe.LastRow <= headerRow)
                probe.Blocking.Add("模板列 " + SummaryColumns.Letter(firstColumn) + " 在表头行以下没有内容，没有可复制的区块。");
            return probe;
        }

        private static object ReadValues(object range)
        {
            try { return ((dynamic)range).Value2; }
            catch (Exception ex) { if (!ComFormulaHost.Missing(ex)) throw; return null; }
        }

        private static object ReadFormulas(object range)
        {
            try { return ((dynamic)range).Formula; }
            catch (Exception ex) { if (!ComFormulaHost.Missing(ex)) throw; return null; }
        }
    }

    /// <summary>手动重算 + 恢复（需求：批量前 xlCalculationManual）。</summary>
    internal sealed class HostCalculation : IDisposable
    {
        private readonly object app;
        private readonly int previous;
        private bool disposed;

        internal HostCalculation(object application)
        {
            app = application;
            previous = Convert.ToInt32(((dynamic)app).Calculation, CultureInfo.InvariantCulture);
            try
            {
                ((dynamic)app).Calculation = -4135;   // xlCalculationManual
                if (Convert.ToInt32(((dynamic)app).Calculation, CultureInfo.InvariantCulture) != -4135)
                    throw new InvalidOperationException("宿主未接受手动重算设置。");
            }
            catch { try { ((dynamic)app).Calculation = previous; } catch { } throw; }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { ((dynamic)app).Calculation = previous; }
            catch (Exception ex) { Connect.Log("单体汇总批量链接：恢复重算方式失败 " + ex.Message); }
        }
    }

    /// <summary>
    /// 关掉「更新值 / 更新链接」询问。
    ///
    /// 为什么需要：这一版写进去的是指向各公司工作簿的外链，测试机上那些文件往往
    /// 不在，WPS 恢复自动重算时会逐条弹「更新值」，一条接一条，看起来像"反复弹出"。
    /// 本类在批处理与还原期间把 AskToUpdateLinks / DisplayAlerts 关掉，
    /// 由外到内先恢复重算方式、最后才恢复询问，保证恢复那一刻也不再问。
    /// </summary>
    internal sealed class HostLinkQuiet : IDisposable
    {
        private readonly object app;
        private readonly bool ask, alerts;
        private bool disposed;

        internal HostLinkQuiet(object application)
        {
            app = application;
            ask = ComFormulaHost.Boolean(((dynamic)app).AskToUpdateLinks, "AskToUpdateLinks");
            alerts = ComFormulaHost.Boolean(((dynamic)app).DisplayAlerts, "DisplayAlerts");
            try
            {
                ((dynamic)app).AskToUpdateLinks = false;
                ((dynamic)app).DisplayAlerts = false;
            }
            catch { try { ((dynamic)app).AskToUpdateLinks = ask; ((dynamic)app).DisplayAlerts = alerts; } catch { } throw; }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            try { ((dynamic)app).DisplayAlerts = alerts; }
            catch (Exception ex) { Connect.Log("单体汇总批量链接：恢复提醒设置失败 " + ex.Message); }
            try { ((dynamic)app).AskToUpdateLinks = ask; }
            catch (Exception ex) { Connect.Log("单体汇总批量链接：恢复链接询问失败 " + ex.Message); }
        }
    }

    /// <summary>汇总表写入：备份 → 铺列 → 换链 → 表头 → 日志。</summary>
    internal static class SummaryWriter
    {
        private const string LogSheetName = "日志_单体汇总";
        private const string BackupPrefix = "备份_单体汇总_";

        internal static SummaryRunReport Run(object app, object book, object sheet, SummaryRequest request,
            SummaryTemplateProbe probe, string backupSheet, Action<string> status, Action<int, string> progress = null)
        {
            List<SummaryYearColumn> columns = request.Columns();
            SummaryRunReport report = new SummaryRunReport();
            report.Columns = columns.Count;
            int firstColumn = request.TemplateColumn;
            int lastColumn = firstColumn + columns.Count - 1;

            // 1) 备份由调用方在此之前完成（失败时也要能还原，所以备份名先落到会话里）
            report.BackupSheet = backupSheet;
            if (status != null) status("已备份工作表为「" + report.BackupSheet + "」，开始铺列…");

            using (HostQuiet quiet = new HostQuiet(app))
            using (HostLinkQuiet links = new HostLinkQuiet(app))
            using (HostCalculation calculation = new HostCalculation(app))
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                int headerRow = request.SheetHeaderRow;   // 公司列标题（占位）所在行
                int startRow = 1;                          // 复制从表顶开始，表头行可以不在第 1 行
                int lastRow = probe.LastRow;
                bool quoteWarned = false;

                // 2) 源列公式快照（换链与「原样带出」都以此为准，先读后写）
                List<string> sourceFormulas = new List<string>();
                object formulaValues = ReadFormulas(bag, s, firstColumn, startRow, lastRow);
                int rows = lastRow - startRow + 1;
                for (int r = 0; r < rows; r++)
                {
                    object value = SummaryArrays.Pick(formulaValues, r, rows, 0, 1);
                    string text = value as string;
                    sourceFormulas.Add(text != null && text.Length > 0 && text[0] == '=' ? text : null);
                }

                double width = ReadColumnWidth(bag, s, firstColumn);

                dynamic sourceRange = bag.Own(s.Range[
                    CellAddress.Name(startRow, firstColumn) + ":"
                    + CellAddress.Name(lastRow, firstColumn)]);

                // 3) 逐列铺开：List[0] 就是模板列本身，从第 2 家公司起复制到右侧
                for (int i = 1; i < columns.Count; i++)
                {
                    if (status != null) status("正在生成第 " + (i + 1).ToString(CultureInfo.InvariantCulture)
                        + "/" + columns.Count.ToString(CultureInfo.InvariantCulture) + " 列…");
                    int column = firstColumn + i;
                    dynamic targetRange = bag.Own(s.Range[
                        CellAddress.Name(startRow, column) + ":"
                        + CellAddress.Name(lastRow, column)]);
                    ClearBlock(targetRange);
                    sourceRange.Copy(targetRange);
                    report.CopiedColumns++;
                    if (width > 0) WriteColumnWidth(bag, s, column, width);
                }

                // 4) From the immutable snapshot, shift internal relative columns;
                //    freeze external addresses and replace only the source company/year.
                for (int i = 0; i < columns.Count; i++)
                {
                    int column = firstColumn + i;
                    SummaryCompany company = columns[i].Company;
                    if (status != null) status("正在处理公司列 " + SummaryColumns.Letter(column) + "：" + company.Code);
                    string letter = SummaryColumns.Letter(column);
                    int touched = 0;
                    for (int r = 0; r < rows; r++)
                    {
                        string original = sourceFormulas[r];
                        if (original == null) continue;
                        SummaryRewrite rewrite = request.Rewrite(original, columns[i]);
                        bool mustWriteBack = i > 0 || rewrite.Count > 0;
                        if (!mustWriteBack) continue;
                        // WPS 拒绝「文件名含括号且引用没加引号」的外链公式（实测 0x8FE3000B），
                        // 这种公司名只补引号、不动引用里的任何内容。
                        string text = rewrite.Text;
                        if (rewrite.Count > 0)
                        {
                            string quotedText = SummaryLinkRewriter.QuoteLinks(text, company.FileName);
                            if (!string.Equals(quotedText, text, StringComparison.Ordinal))
                            {
                                text = quotedText;
                                if (!quoteWarned)
                                {
                                    quoteWarned = true;
                                    report.Warnings.Add("公司名含括号等特殊字符：WPS 只接受带单引号的外链写法，已按「只加引号、不改文件名与地址」处理。");
                                }
                            }
                        }
                        dynamic cell = bag.Own(s.Cells[startRow + r, column]);
                        cell.Formula = text;
                        touched++;
                        if (rewrite.Count > 0) report.RewrittenLinks += rewrite.Count;
                        report.FormulaCells++;
                        if (report.Samples.Count < 20)
                        {
                            SummaryLinkSample sample = new SummaryLinkSample();
                            sample.Column = letter;
                            sample.Cell = letter + (startRow + r).ToString(CultureInfo.InvariantCulture);
                            sample.Before = original;
                            sample.After = text;
                            report.Samples.Add(sample);
                        }
                    }
                    if (touched > 0) report.RewrittenCells += touched;
                    else if (probe.FormulaCells == 0) report.SkippedColumns++;
                }

                // 5) 表头：只改该列表头行的「大合并」占位，逐列写各自「编码&简称」
                for (int i = 0; i < columns.Count; i++)
                {
                    dynamic cell = bag.Own(s.Cells[headerRow, firstColumn + i]);
                    cell.Value2 = columns[i].Company.Code;
                    if (request.Annual)
                    {
                        dynamic yearCell = bag.Own(s.Cells[request.YearHeaderRow, firstColumn + i]);
                        yearCell.Value2 = columns[i].YearLabel;
                        yearCell.HorizontalAlignment = -4108; // xlCenter
                        yearCell.VerticalAlignment = -4108;   // xlCenter
                    }
                    if (progress != null) progress(i + 1, "已写入列 " + SummaryColumns.Letter(firstColumn + i) + "，等待最终核验");
                }

                // 6) 读回校验：表头与首列公式必须与预期一致
                if (status != null) status("正在逐列核验表头和公式…");
                string verify = Verify(s, columns, request, headerRow, startRow, lastRow, sourceFormulas);
                if (verify != null) throw new InvalidOperationException(verify);

                // 6b) 趁「不询问链接」仍然生效时先把目标表算完：
                //     否则稍后自动重算会逐条弹「更新值」，用户看到的就是反复弹窗。
                try { s.Calculate(); }
                catch (Exception ex) { Connect.Log("单体汇总批量链接：目标表重算失败 " + ex.Message); }

                if (probe.Tokens.Count == 0)
                    report.Warnings.Add("模板列里没有 [xxx.xlsx] 形式的外链，换链已跳过（按需求记日志、不阻断）。");
                else if (probe.Tokens.Count > 0 && report.RewrittenLinks == 0)
                    report.Warnings.Add("模板列存在外链但没有一处匹配「" + request.SourceFile
                        + "」，请核对模板公司全称是否与公式里的文件名一致。");
            }

            // 7) 日志表（不自动保存工作簿，日志与备份都留在未保存状态）
            report.LogSheet = AppendLog(book, request, columns, probe, report);
            return report;
        }

        // ---------------------------- 备份 / 还原 ----------------------------

        internal static string Backup(object app, object book, object sheet)
        {
            if (ComFormulaHost.Boolean(((dynamic)book).ProtectStructure, "ProtectStructure"))
                throw new InvalidOperationException("工作簿结构受到保护，无法复制备份工作表。");
            using (ComBag bag = new ComBag())
            {
                string name = BackupPrefix + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
                ((dynamic)sheet).Copy(Type.Missing, sheet);
                dynamic created = bag.Own(((dynamic)app).ActiveSheet);
                created.Name = WorkbookUi.UniqueName(name, SheetNames(book, bag));
                created.Visible = 0;   // xlSheetHidden：静默备份，不打扰用户
                return Convert.ToString(created.Name);
            }
        }

        internal static void Restore(object app, object book, object sheet, string backupName,
            int firstColumn, int columnCount, int lastRow, int headerRow)
        {
            using (HostQuiet quiet = new HostQuiet(app))
            using (HostLinkQuiet links = new HostLinkQuiet(app))
            using (ComBag bag = new ComBag())
            {
                dynamic backup = bag.Own(((dynamic)book).Worksheets[backupName]);
                dynamic s = sheet;
                dynamic b = backup;
                for (int i = 0; i < columnCount; i++)
                {
                    int column = firstColumn + i;
                    // 还原范围与生成范围一致：第 1 行到末行（表头行只是其中一行）
                    dynamic target = bag.Own(s.Range[
                        CellAddress.Name(1, column) + ":"
                        + CellAddress.Name(lastRow, column)]);
                    dynamic original = bag.Own(b.Range[
                        CellAddress.Name(1, column) + ":"
                        + CellAddress.Name(lastRow, column)]);
                    ClearBlock(target);
                    original.Copy(target);
                    double width = ReadColumnWidth(bag, b, column);
                    if (width > 0) WriteColumnWidth(bag, s, column, width);
                }
            }
        }

        // ---------------------------- 日志表 ----------------------------

        private static string AppendLog(object book, SummaryRequest request, List<SummaryYearColumn> columns,
            SummaryTemplateProbe probe, SummaryRunReport report)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                dynamic sheet = null;
                int count = Convert.ToInt32(sheets.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic candidate = bag.Own(sheets.Item(i));
                    if (string.Equals(Convert.ToString(candidate.Name), LogSheetName, StringComparison.Ordinal))
                    {
                        sheet = candidate;
                        break;
                    }
                }
                bool created = false;
                if (sheet == null)
                {
                    sheet = WorkbookUi.NewSheet(book, LogSheetName, bag, false);
                    created = true;
                }

                dynamic s = sheet;
                int startRow;
                if (created)
                {
                    string[] headers = { "时间", "编码&简称", "全称", "目标列", "结果", "说明" };
                    object[,] head = new object[1, headers.Length];
                    for (int c = 0; c < headers.Length; c++) head[0, c] = headers[c];
                    dynamic headRange = bag.Own(s.Range[
                        "A1:" + SummaryColumns.Letter(headers.Length) + "1"]);
                    headRange.Value2 = head;
                    WorkbookUi.StyleHeader(s, 1, 1, headers.Length);
                    startRow = 2;
                }
                else
                {
                    dynamic used = bag.Own(s.UsedRange);
                    startRow = Convert.ToInt32(used.Row, CultureInfo.InvariantCulture)
                        + Convert.ToInt32(bag.Own(used.Rows).Count, CultureInfo.InvariantCulture);
                }

                string stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                object[,] rows = new object[columns.Count, 6];
                for (int i = 0; i < columns.Count; i++)
                {
                    SummaryCompany company = columns[i].Company;
                    string column = SummaryColumns.Letter(columns[i].Column);
                    rows[i, 0] = stamp;
                    rows[i, 1] = company.Code;
                    rows[i, 2] = company.Name;
                    rows[i, 3] = column;
                    rows[i, 4] = "已生成";
                    rows[i, 5] = "[" + request.SourceFile + "] → [" + company.FileName + "]"
                        + (request.Annual ? "；" + columns[i].YearLabel + "；目录 " + request.YearPathPattern : "")
                        + "；第 1～" + probe.LastRow.ToString(CultureInfo.InvariantCulture) + " 行；Lead 表头第 " + request.HeaderRow.ToString(CultureInfo.InvariantCulture) + " 行";
                }
                dynamic body = bag.Own(s.Range[
                    "A" + startRow.ToString(CultureInfo.InvariantCulture) + ":F"
                    + (startRow + columns.Count - 1).ToString(CultureInfo.InvariantCulture)]);
                body.Value2 = rows;
                try { ((dynamic)bag.Own(s.Columns)).AutoFit(); } catch (Exception ex) { Connect.Log("单体汇总批量链接：日志列宽自适应失败 " + ex.Message); }
                return Convert.ToString(sheet.Name);
            }
        }

        // ---------------------------- 小工具 ----------------------------

        private static void ClearBlock(object range)
        {
            dynamic r = range;
            // 合并格必须先拆再清；否则 Clear 后仍然保留合并，粘贴会被拒
            try { if (ComFormulaHost.Boolean(r.MergeCells, "MergeCells")) r.UnMerge(); } catch (Exception ex) { if (!ComFormulaHost.Missing(ex)) throw; }
            r.Clear();
        }

        private static object ReadFormulas(ComBag bag, object sheet, int column, int firstRow, int lastRow)
        {
            dynamic range = bag.Own(((dynamic)sheet).Range[
                CellAddress.Name(firstRow, column) + ":" + CellAddress.Name(lastRow, column)]);
            try { return range.Formula; }
            catch (Exception ex) { if (!ComFormulaHost.Missing(ex)) throw; return null; }
        }

        private static double ReadColumnWidth(ComBag bag, object sheet, int column)
        {
            try
            {
                string letter = SummaryColumns.Letter(column);
                dynamic range = bag.Own(((dynamic)sheet).Range[letter + ":" + letter]);
                return Convert.ToDouble(range.ColumnWidth, CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                if (!ComFormulaHost.Missing(ex)) Connect.Log("单体汇总批量链接：读取列宽失败 " + ex.Message);
                return 0;
            }
        }

        private static void WriteColumnWidth(ComBag bag, object sheet, int column, double width)
        {
            try
            {
                string letter = SummaryColumns.Letter(column);
                dynamic range = bag.Own(((dynamic)sheet).Range[letter + ":" + letter]);
                range.ColumnWidth = width;
            }
            catch (Exception ex) { Connect.Log("单体汇总批量链接：设置列宽失败 " + ex.Message); }
        }

        private static List<string> SheetNames(object book, ComBag bag)
        {
            List<string> names = new List<string>();
            dynamic sheets = bag.Own(((dynamic)book).Worksheets);
            int count = Convert.ToInt32(sheets.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
                using (ComBag one = new ComBag())
                    names.Add(Convert.ToString(((dynamic)one.Own(sheets.Item(i))).Name));
            return names;
        }

        /// <summary>读回校验：表头值、首列公式是否与预期一致；不一致只记警告，不假装成功。</summary>
        private static string Verify(object sheet, List<SummaryYearColumn> columns, SummaryRequest request,
            int headerRow, int startRow, int lastRow, List<string> sourceFormulas)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic s = sheet;
                for (int i = 0; i < columns.Count; i++)
                {
                    dynamic cell = bag.Own(s.Cells[headerRow, columns[i].Column]);
                    if (SummaryArrays.Text(cell.Value2) != columns[i].Company.Code)
                        return "公司表头读回不一致：" + SummaryColumns.Letter(columns[i].Column);
                    if (request.Annual)
                    {
                        dynamic year = bag.Own(s.Cells[request.YearHeaderRow, columns[i].Column]);
                        if (SummaryArrays.Text(year.Value2) != columns[i].YearLabel) return "年度表头读回不一致：" + SummaryColumns.Letter(columns[i].Column);
                    }
                    object formulas = ReadFormulas(bag, s, columns[i].Column, startRow, lastRow);
                    for (int r = 0; r < sourceFormulas.Count; r++)
                    {
                        if (sourceFormulas[r] == null || startRow + r == headerRow || (request.Annual && startRow + r == request.YearHeaderRow)) continue;
                        SummaryRewrite rewrite = request.Rewrite(sourceFormulas[r], columns[i]);
                        string expected = rewrite.Count > 0 ? SummaryLinkRewriter.QuoteLinks(rewrite.Text, columns[i].Company.FileName) : rewrite.Text;
                        string actual = SummaryArrays.Text(SummaryArrays.Pick(formulas, r, lastRow - startRow + 1, 0, 1));
                        if (actual != expected) return "公式读回不一致：" + SummaryColumns.Letter(columns[i].Column) + (startRow + r);
                    }
                }
            }
            return null;
        }
    }
}

