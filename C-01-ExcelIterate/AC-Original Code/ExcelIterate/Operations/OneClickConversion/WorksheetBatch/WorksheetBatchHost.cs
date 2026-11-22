using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace Excel_Iterate
{
    internal sealed class WorksheetBatchRow
    {
        internal string Workbook, Detail, Fingerprint;
        internal int Changed, Skipped, ConvertedToText, SkippedDates;
    }

    internal sealed class WorksheetBatchReport
    {
        internal readonly List<WorksheetBatchRow> Rows = new List<WorksheetBatchRow>();
        internal WorksheetEditBatch Undo;
        internal string CleanupWarning;
        internal int Cells { get { return Rows.Sum(row => row.Changed); } }
        internal int Workbooks { get { return Rows.Count(row => row.Changed > 0); } }
        internal bool Matches(WorksheetBatchReport other)
        {
            return Rows.Count == other.Rows.Count && Rows.Zip(other.Rows, (a, b) =>
                string.Equals(a.Workbook, b.Workbook, StringComparison.OrdinalIgnoreCase) && a.Detail == b.Detail &&
                a.Fingerprint == b.Fingerprint && a.Changed == b.Changed && a.Skipped == b.Skipped).All(same => same);
        }
    }

    internal sealed class WorksheetBatchHost
    {
        // The first entry restores geometry after cells on undo. The last
        // captures geometry after all writes, or restores it on undo rollback.
        private sealed class RowLayout
        {
            private sealed class RowState
            {
                private readonly double height;
                private readonly bool hidden;
                private readonly bool? standardHeight;

                private RowState(double height, bool hidden, bool? standardHeight)
                { this.height = height; this.hidden = hidden; this.standardHeight = standardHeight; }

                internal static RowState Read(dynamic row)
                {
                    double height = Height(row);
                    if (double.IsNaN(height) || double.IsInfinity(height) || height <= 0)
                        throw new InvalidOperationException("无法可靠读取行高，当前工作簿未加入批次。");
                    return new RowState(height, ComConversionHost.Boolean((object)row.Hidden, "Hidden"), StandardHeight(row));
                }

                private static double Height(dynamic row)
                { return Convert.ToDouble((object)row.RowHeight, CultureInfo.InvariantCulture); }

                private static bool? StandardHeight(dynamic row)
                { return WorksheetBatchFormat.ReadOptional(() => (object)row.UseStandardHeight) as bool?; }

                private bool SameHeight(double current) { return Math.Abs(height - current) < 0.01; }

                internal bool Matches(RowState current)
                { return SameHeight(current.height) && hidden == current.hidden && standardHeight == current.standardHeight; }

                internal void Restore(dynamic row)
                {
                    if (!SameHeight(Height(row)) || standardHeight != StandardHeight(row))
                    {
                        if (standardHeight == true) row.UseStandardHeight = true;
                        else row.RowHeight = height;
                    }
                    if (hidden != ComConversionHost.Boolean((object)row.Hidden, "Hidden")) row.Hidden = hidden;
                }

                public override string ToString()
                { return string.Format(CultureInfo.InvariantCulture, "行高={0:R}，隐藏={1}，标准行高={2}", height, hidden, standardHeight); }
            }

            private sealed class Entry : IWorksheetChange
            {
                private readonly RowLayout owner;
                private readonly bool end;

                internal Entry(RowLayout owner, bool end) { this.owner = owner; this.end = end; }
                public string Address { get { return owner.cellAddress + " 所在行"; } }
                public string DescribeBefore() { return owner.before.ToString(); }

                public void Check(object sheet, bool after)
                {
                    if (after != end) return;
                    using (var bag = new ComBag())
                    {
                        var expected = after ? owner.after : owner.before;
                        RowState current = RowState.Read(owner.ResolveRow(sheet, bag));
                        if (expected == null || !expected.Matches(current))
                            throw new InvalidOperationException("行高或隐藏状态已变化。");
                    }
                }

                public void Write(object sheet, bool after)
                {
                    if (after != end) return;
                    using (var bag = new ComBag())
                    {
                        dynamic row = owner.ResolveRow(sheet, bag);
                        if (after && owner.after == null) owner.after = RowState.Read(row);
                        else (after ? owner.after : owner.before).Restore(row);
                    }
                }
            }

            private readonly string cellAddress;
            private readonly RowState before;
            private RowState after;
            internal readonly IWorksheetChange Start, End;

            internal RowLayout(string cellAddress, object cell, ComBag bag)
            {
                this.cellAddress = cellAddress;
                before = RowState.Read(bag.Own(((dynamic)cell).EntireRow));
                Start = new Entry(this, false);
                End = new Entry(this, true);
            }

            private object ResolveRow(object sheet, ComBag bag)
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[cellAddress]);
                return bag.Own(cell.EntireRow);
            }
        }

        private sealed class BookItem
        {
            internal object Book;
            internal string FullName, Name, SkipReason;
            internal bool Hidden;
        }

        private readonly object application;
        internal WorksheetBatchHost(object app) { application = app; }

        internal WorksheetBatchReport Run(WorksheetBatchRequest request, WorksheetBatchReport preview, string recoveryDirectory)
        {
            IList<string> cells = request.Validate();
            IList<string> paths = request.FilePaths();
            var report = new WorksheetBatchReport();
            WorksheetEditBatch batch = preview == null ? null : new WorksheetEditBatch("工作表批量操作");
            var opened = new List<object>();
            bool committed = false;
            using (var bag = new ComBag())
            {
                dynamic app = application;
                object originalWindow = bag.Own((object)app.ActiveWindow);
                try
                {
                    using (new HostQuiet(application))
                    {
                        dynamic books = bag.Own((object)app.Workbooks);
                        List<BookItem> openBooks = ReadOpenBooks((object)books, bag);
                        var targets = request.Source == WorksheetBatchSource.OpenWorkbooks
                            ? openBooks.Where(item => !item.Hidden).OrderBy(item => item.FullName, StringComparer.OrdinalIgnoreCase).ToList()
                            : MatchFiles(paths, openBooks);
                        if (targets.Count == 0) throw new InvalidOperationException("当前 WPS 中没有打开的工作簿。");
                        if (targets.Count > 100 || (long)targets.Count * cells.Count > ConversionRangeAddress.MaximumCells)
                            throw new InvalidOperationException("本批最多扫描 100 个工作簿、合计 100000 个单元格，请缩小范围。");
                        foreach (BookItem item in targets)
                            report.Rows.Add(Scan(item, (object)books, request, cells, batch, bag, opened));
                    }
                    if (preview != null)
                    {
                        if (!report.Matches(preview))
                            throw new InvalidOperationException("工作簿、预览结果或目标单元格已变化。请重新预览。");
                        if (report.Cells > 0)
                        {
                            try { batch.Execute(application, true, recoveryDirectory); }
                            catch (Exception error)
                            {
                                if (!batch.LastExecutionCommitted) throw;
                                report.CleanupWarning = error.Message;
                                UiDiagnostics.Record("worksheet-batch-finish", error);
                            }
                            committed = true; report.Undo = batch; batch = null;
                        }
                    }
                }
                finally
                {
                    if (batch != null) batch.Dispose();
                    if (!committed)
                        for (int i = opened.Count - 1; i >= 0; i--)
                            try { ((dynamic)opened[i]).Close(false); }
                            catch (Exception error) { UiDiagnostics.Record("worksheet-batch-close", error); }
                    if (originalWindow != null && paths.Count > 0)
                        try { ((dynamic)originalWindow).Activate(); }
                        catch (Exception error) { UiDiagnostics.Record("worksheet-batch-window", error); }
                }
            }
            return report;
        }

        private static List<BookItem> ReadOpenBooks(dynamic books, ComBag bag)
        {
            var result = new List<BookItem>();
            for (int i = 1; i <= (int)books.Count; i++)
            {
                dynamic book = bag.Own(books.Item(i));
                using (var refs = new ComBag())
                {
                    dynamic windows = refs.Own(book.Windows);
                    bool visible = false;
                    for (int j = 1; j <= (int)windows.Count; j++)
                    {
                        dynamic window = refs.Own(windows.Item(j));
                        visible |= (bool)window.Visible;
                    }
                    result.Add(new BookItem { Book = (object)book, FullName = (string)book.FullName,
                        Name = (string)book.Name, Hidden = !visible });
                }
            }
            return result;
        }

        private static List<BookItem> MatchFiles(IList<string> paths, List<BookItem> openBooks)
        {
            var duplicateNames = new HashSet<string>(paths.GroupBy(Path.GetFileName, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1).Select(group => group.Key), StringComparer.OrdinalIgnoreCase);
            var result = new List<BookItem>();
            foreach (string path in paths)
            {
                var existing = openBooks.FirstOrDefault(item => string.Equals(item.FullName, path, StringComparison.OrdinalIgnoreCase));
                string name = Path.GetFileName(path);
                result.Add(existing ?? new BookItem { FullName = path, Name = name,
                    SkipReason = duplicateNames.Contains(name) ? "所选文件中有同名工作簿，请分批处理。" : null });
            }
            return result;
        }

        private WorksheetBatchRow Scan(BookItem item, dynamic books, WorksheetBatchRequest request, IList<string> cells,
            WorksheetEditBatch batch, ComBag keep, List<object> opened)
        {
            var row = new WorksheetBatchRow { Workbook = item.FullName };
            if (item.SkipReason != null) { row.Detail = item.SkipReason; return row; }
            using (var bag = new ComBag())
            {
                object book = item.Book;
                bool keepOpen = false;
                try
                {
                    if (item.Hidden) throw new InvalidOperationException("工作簿窗口当前处于隐藏状态。");
                    if (book == null)
                    {
                        if (!File.Exists(item.FullName)) throw new FileNotFoundException("工作簿文件已移动或无法访问。");
                        if ((File.GetAttributes(item.FullName) & FileAttributes.ReadOnly) != 0)
                            throw new InvalidOperationException("工作簿文件为只读。");
                        for (int i = 1; i <= (int)books.Count; i++)
                        {
                            using (var refs = new ComBag())
                            {
                                dynamic existing = refs.Own(books.Item(i));
                                if (string.Equals((string)existing.Name, item.Name, StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("同名工作簿已打开，请关闭后重试。");
                            }
                        }
                        book = bag.Own(OpenBook((object)books, item.FullName, batch == null));
                        if (!string.Equals((string)((dynamic)book).FullName, item.FullName, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("宿主打开的工作簿路径与所选文件不一致。");
                    }
                    if ((item.Book != null || batch != null) && (bool)((dynamic)book).ReadOnly)
                        throw new InvalidOperationException("工作簿以只读方式打开。");
                    dynamic worksheets = bag.Own(((dynamic)book).Worksheets);
                    object sheet = null;
                    for (int i = 1; i <= (int)worksheets.Count; i++)
                    {
                        object candidate = worksheets.Item(i);
                        if (string.Equals((string)((dynamic)candidate).Name, request.SheetName, StringComparison.OrdinalIgnoreCase))
                        { sheet = bag.Own(candidate); break; }
                        ComConversionHost.Release(candidate);
                    }
                    if (sheet == null) throw new InvalidOperationException("缺少工作表「" + request.SheetName + "」。");
                    if ((bool)((dynamic)sheet).ProtectContents) throw new InvalidOperationException("工作表受保护。");
                    WorksheetChangeSet.CheckGroups(book);
                    var defaults = request.ClearFormatting ? WorksheetBatchFormat.ClearDefaults.Read(book) : null;
                    List<IWorksheetChange> changes = PlanCells(sheet, request, cells, row, defaults);
                    row.Fingerprint = Fingerprint(changes);
                    row.Detail = row.Changed > 0 ? "可处理" : "没有匹配的改动";
                    if (row.Skipped > 0) row.Detail += "，跳过 " + row.Skipped + " 个合并格或受保护的公式区域";
                    if (row.ConvertedToText > 0) row.Detail += "，" + row.ConvertedToText + " 个数值转文本";
                    if (row.SkippedDates > 0) row.Detail += "，跳过 " + row.SkippedDates + " 个日期时间值的替换";
                    if (batch != null && changes.Count > 0)
                    {
                        batch.AddSheet(sheet).Changes.AddRange(changes);
                        if (item.Book == null)
                        {
                            opened.Add(keep.Own(bag.Detach(book)));
                            keepOpen = true;
                        }
                    }
                }
                catch (Exception error)
                {
                    row.Changed = 0; row.Skipped = 0; row.ConvertedToText = 0; row.SkippedDates = 0; row.Fingerprint = null;
                    row.Detail = error.Message;
                }
                finally
                {
                    if (item.Book == null && book != null && !keepOpen) ((dynamic)book).Close(false);
                }
            }
            return row;
        }

        private static List<IWorksheetChange> PlanCells(object sheet, WorksheetBatchRequest request, IList<string> cells,
            WorksheetBatchRow row, WorksheetBatchFormat.ClearDefaults defaults)
        {
            var changes = new List<IWorksheetChange>();
            var layouts = new SortedDictionary<int, RowLayout>();
            foreach (string address in cells)
            {
                try
                {
                    using (var bag = new ComBag())
                    {
                        dynamic cell = bag.Own(((dynamic)sheet).Range[address]);
                        if ((bool)cell.MergeCells) { row.Skipped++; continue; }
                        int beforeCount = changes.Count;
                        if (request.ReplacesText && !(bool)cell.HasFormula)
                        {
                            object raw = cell.Value2;
                            string text = WorksheetBatchRequest.SearchText(raw);
                            if (text != null)
                            {
                                bool numeric = !(raw is string);
                                if (numeric && (WorksheetBatchRequest.IsTemporalFormat((string)cell.NumberFormat) || ((object)cell.Value) is DateTime))
                                    row.SkippedDates++;
                                else
                                {
                                    string replacement = request.Replace(text);
                                    if (replacement != text)
                                    {
                                        if ((bool)cell.HasArray || (bool)cell.HasSpill) { row.Skipped++; continue; }
                                        var value = new ComConversionHost.ComCell((object)cell, sheet, address, capture: false).Read(false);
                                        var after = value.Copy(); after.Value = WorksheetBatchRequest.ReplacementValue(value.Value, replacement);
                                        if (numeric && after.Value is string) row.ConvertedToText++;
                                        changes.Add(new CellContentChange(address, value, after, false));
                                    }
                                }
                            }
                        }
                        if (request.FormatsCells)
                        {
                            WorksheetBatchFormat format = WorksheetBatchFormat.Plan((object)cell, address, request, defaults);
                            if (format != null)
                            {
                                if (format.AffectsRowLayout)
                                {
                                    int rowNumber = (int)cell.Row;
                                    if (!layouts.ContainsKey(rowNumber)) layouts.Add(rowNumber, new RowLayout(address, (object)cell, bag));
                                }
                                changes.Add(format);
                            }
                        }
                        if (beforeCount != changes.Count) row.Changed++;
                    }
                }
                catch (Exception error) { throw new InvalidOperationException(address + "：" + error.Message, error); }
            }
            if (layouts.Count == 0) return changes;
            var ordered = new List<IWorksheetChange>(changes.Count + 2 * layouts.Count);
            ordered.AddRange(layouts.Values.Select(layout => layout.Start));
            ordered.AddRange(changes);
            ordered.AddRange(layouts.Values.Select(layout => layout.End));
            return ordered;
        }

        private static string Fingerprint(List<IWorksheetChange> changes)
        {
            using (var sha = SHA256.Create())
            {
                using (var stream = new CryptoStream(Stream.Null, sha, CryptoStreamMode.Write, true))
                using (var writer = new BinaryWriter(stream, Encoding.UTF8, true))
                {
                    foreach (var change in changes)
                    {
                        writer.Write(change.Address);
                        writer.Write(change.DescribeBefore());
                        var format = change as WorksheetBatchFormat;
                        var content = change as CellContentChange;
                        writer.Write(format != null ? format.DescribePlan() : content != null ? content.DescribePlan() : "");
                    }
                }
                return Convert.ToBase64String(sha.Hash);
            }
        }

        private object OpenBook(dynamic books, string path, bool readOnly)
        {
            dynamic app = application;
            int security = (int)app.AutomationSecurity;
            bool alerts = (bool)app.DisplayAlerts;
            object book = null;
            try
            {
                try
                {
                    app.AutomationSecurity = 3;
                    if ((int)app.AutomationSecurity != 3) throw new InvalidOperationException("无法为打开文件禁用宏。");
                    app.DisplayAlerts = false;
                    book = books.Open(Filename: path, UpdateLinks: 0, ReadOnly: readOnly, Password: "", WriteResPassword: "",
                        IgnoreReadOnlyRecommended: true, AddToMru: false);
                }
                finally
                {
                    try { app.DisplayAlerts = alerts; }
                    finally { app.AutomationSecurity = security; }
                }
                return book;
            }
            catch
            {
                if (book != null)
                {
                    try { ((dynamic)book).Close(false); }
                    finally { ComConversionHost.Release(book); }
                }
                throw;
            }
        }
    }
}
