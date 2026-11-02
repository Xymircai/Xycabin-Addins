using System;
using System.Collections.Generic;
using System.IO;

namespace Excel_Iterate
{
    internal sealed class CompanyWorkbookStateException : InvalidOperationException
    {
        internal CompanyWorkbookStateException(string message) : base(message) { }
    }
    // All COM operations stay on the host STA. Only disk snapshots are opened.
    internal sealed class CompanyWorkbookHost : IDisposable
    {
        private readonly object app;
        private readonly ComBag held = new ComBag();
        private object book;
        private readonly string folder;
        internal readonly string Snapshot;
        internal readonly List<string> Sheets = new List<string>();
        internal CompanyWorkbookHost(object application, string template)
        {
            app = application;
            string extension = Path.GetExtension(template).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls") throw new InvalidOperationException("模板仅支持 .xlsx 和 .xls。");
            folder = Path.Combine(Path.GetTempPath(), "ExcelIterate-Company-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder);
            Snapshot = Path.Combine(folder, Path.GetFileName(template));
            try
            {
                // Copy disk bytes, never Save/SaveCopyAs the user's workbook.
                using (var source = new FileStream(template, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var target = new FileStream(Snapshot, FileMode.CreateNew, FileAccess.Write)) source.CopyTo(target);
                string reading = Path.Combine(folder, "read-" + Guid.NewGuid().ToString("N") + extension);
                File.Copy(Snapshot, reading, false);
                using (var quiet = new CompanyWorkbookQuiet(app))
                {
                    dynamic books = held.Own(((dynamic)app).Workbooks);
                    book = held.Own(books.Open(reading, 0, true));
                    if (book == null) throw new InvalidOperationException("无法打开模板读取副本。");
                    using (var bag = new ComBag())
                    {
                        dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                        for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++)
                            using (var one = new ComBag()) Sheets.Add(Convert.ToString(((dynamic)one.Own(sheets.Item(i))).Name));
                    }
                }
            }
            catch { Dispose(); throw; }
        }
        private object Sheet(string name, ComBag bag)
        {
            if (!Sheets.Contains(name)) throw new InvalidOperationException("工作表不存在：" + name);
            dynamic sheets = bag.Own(((dynamic)book).Worksheets);
            return bag.Own(sheets.Item(name));
        }
        internal Dictionary<string, int> Headers(string name, int row)
        {
            using (var bag = new ComBag()) return HeadersOn(Sheet(name, bag), row);
        }
        internal static Dictionary<string, int> HeadersOn(object sheet, int row)
        {
            using (var bag = new ComBag())
            {
                dynamic s = sheet;
                dynamic allRows = bag.Own(s.Rows);
                if (row < 1 || row > Convert.ToInt32(allRows.Count)) throw new InvalidOperationException("表头行超出工作表范围。");
                dynamic used = bag.Own(s.UsedRange);
                dynamic columns = bag.Own(used.Columns);
                int last = Convert.ToInt32(used.Column) + Convert.ToInt32(columns.Count) - 1;
                dynamic cells = bag.Own(s.Cells);
                var result = new Dictionary<string, int>(StringComparer.Ordinal);
                for (int c = 1; c <= last; c++)
                    using (var one = new ComBag())
                    {
                        dynamic cell = one.Own(cells.Item(row, c));
                        string raw = CompanyWorkbookRules.Text(cell.Value2), value = raw.Trim();
                        if (value.Length == 0) continue;
                        if (result.ContainsKey(value)) throw new InvalidOperationException("工作表「" + Convert.ToString(s.Name) + "」第 " + row + " 行表头重复：" + value);
                        result.Add(value, c);
                    }
                if (result.Count == 0) throw new InvalidOperationException("所选表头行没有字段。");
                return result;
            }
        }
        private static string Read(object cells, int row, int column)
        {
            using (var bag = new ComBag())
                return CompanyWorkbookRules.Text(((dynamic)bag.Own(((dynamic)cells).Item(row, column))).Value2);
        }
        internal List<CompanyWorkbookRecord> Preview(CompanyWorkbookRequest request, Action<string> log)
        {
            request.Validate();
            var headers = Headers(request.SourceSheet, request.HeaderRow);
            int name = CompanyWorkbookRules.Column(headers, request.NameField);
            int code = request.UpdateLead ? CompanyWorkbookRules.Column(headers, request.CodeField) : 0;
            int full = request.UpdateLead ? CompanyWorkbookRules.Column(headers, request.FullNameField) : 0;
            if (request.UpdateLead)
            {
                using (var bag = new ComBag()) CheckTarget(Sheet(request.TargetSheet, bag), request);
            }
            var result = new List<CompanyWorkbookRecord>();
            using (var bag = new ComBag())
            {
                dynamic sheet = Sheet(request.SourceSheet, bag);
                dynamic allRows = bag.Own(sheet.Rows);
                if (request.LastRow > Convert.ToInt32(allRows.Count)) throw new InvalidOperationException("来源结束行超出工作表范围。");
                if (request.LastRow - request.FirstRow > 100000) throw new InvalidOperationException("单批最多读取 100001 行，请缩小范围。");
                object cells = bag.Own(sheet.Cells);
                for (int row = request.FirstRow; row <= request.LastRow; row++)
                {
                    try
                    {
                        string n = Read(cells, row, name);
                        string c = code == 0 ? "" : Read(cells, row, code), f = full == 0 ? "" : Read(cells, row, full);
                        if (string.IsNullOrWhiteSpace(n) && string.IsNullOrWhiteSpace(c) && string.IsNullOrWhiteSpace(f))
                        { log("跳过来源第 " + row + " 行：所选字段全空。"); continue; }
                        if (request.UpdateLead && (string.IsNullOrWhiteSpace(c) || string.IsNullOrWhiteSpace(f)))
                            throw new InvalidOperationException("来源编码或全称为空。");
                        if (c != c.Trim() || f != f.Trim()) log("提示：第 " + row + " 行编码/全称含前后空格，写入时保留原值。");
                        result.Add(new CompanyWorkbookRecord { Row = row, Name = n, Code = c, FullName = f });
                    }
                    catch (Exception error) { throw new InvalidOperationException("来源表「" + request.SourceSheet + "」第 " + row + " 行：" + error.Message, error); }
                }
            }
            CompanyWorkbookRules.ValidateRecords(result, request.Output, Path.GetExtension(Snapshot), request.StripExcelSuffix, request.RemoveTailCharacters);
            // Check write access without touching existing outputs.
            string probe = Path.Combine(request.Output, ".company-write-" + Guid.NewGuid().ToString("N"));
            using (var file = new FileStream(probe, FileMode.CreateNew, FileAccess.Write)) { }
            File.Delete(probe);
            return result;
        }
        private static void CheckTarget(object sheet, CompanyWorkbookRequest request)
        {
            var headers = HeadersOn(sheet, request.TargetHeaderRow);
            int code = CompanyWorkbookRules.Column(headers, request.TargetCodeField), full = CompanyWorkbookRules.Column(headers, request.TargetNameField);
            if (code == full) throw new InvalidOperationException("两个目标字段不能指向同一列。");
            using (var bag = new ComBag())
            {
                dynamic s = sheet;
                dynamic rows = bag.Own(s.Rows);
                if (request.TargetRow > Convert.ToInt32(rows.Count)) throw new InvalidOperationException("Lead 写入行超出工作表范围。");
                if (Convert.ToBoolean(s.ProtectContents)) throw new InvalidOperationException("目标工作表受保护，请先在模板中解除保护。");
                dynamic cells = bag.Own(s.Cells);
                foreach (int column in new[] { code, full })
                {
                    dynamic cell = bag.Own(cells.Item(request.TargetRow, column));
                    if (Convert.ToBoolean(cell.MergeCells) || Convert.ToBoolean(cell.HasArray))
                        throw new InvalidOperationException("目标第 " + request.TargetRow + " 行包含合并单元格或数组公式，请选择独立数据单元格。");
                }
            }
        }
        // The update delegate is separate from the copy/publish transaction for future rules.
        internal static void UpdateLead(object outputBook, CompanyWorkbookRequest request, CompanyWorkbookRecord record)
        {
            using (var bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)outputBook).Worksheets);
                dynamic sheet = bag.Own(sheets.Item(request.TargetSheet));
                CheckTarget(sheet, request);
                var headers = HeadersOn(sheet, request.TargetHeaderRow);
                dynamic cells = bag.Own(sheet.Cells);
                WriteText(bag.Own(cells.Item(request.TargetRow, CompanyWorkbookRules.Column(headers, request.TargetCodeField))), record.Code);
                WriteText(bag.Own(cells.Item(request.TargetRow, CompanyWorkbookRules.Column(headers, request.TargetNameField))), record.FullName);
            }
        }
        private static void WriteText(object cell, string value)
        {
            dynamic c = cell;
            object format = c.NumberFormat;
            try { c.NumberFormat = "@"; c.Value2 = value; }
            finally { c.NumberFormat = format; }
            if (Convert.ToBoolean(c.HasFormula) || CompanyWorkbookRules.Text(c.Value2) != value)
                throw new InvalidOperationException("目标字段写入后核验不一致。");
        }
        internal void Generate(CompanyWorkbookRequest request, CompanyWorkbookRecord record, Action<object, CompanyWorkbookRequest, CompanyWorkbookRecord> update)
        {
            string stage = Path.Combine(request.Output, ".company-" + Guid.NewGuid().ToString("N") + Path.GetExtension(Snapshot));
            object outputBook = null;
            bool closed = false;
            using (var bag = new ComBag())
            {
                try
                {
                    File.Copy(Snapshot, stage, false);
                    File.SetAttributes(stage, FileAttributes.Normal);
                    if (request.UpdateLead)
                    {
                        using (var quiet = new CompanyWorkbookQuiet(app))
                        {
                            dynamic books = bag.Own(((dynamic)app).Workbooks);
                            outputBook = bag.Own(books.Open(stage, 0, false));
                            if (outputBook == null || Convert.ToBoolean(((dynamic)outputBook).ReadOnly)) throw new InvalidOperationException("输出副本无法以可写方式打开。");
                            update(outputBook, request, record);
                            ((dynamic)outputBook).Save();
                            ((dynamic)outputBook).Close(false); closed = true;
                            outputBook = bag.Own(books.Open(stage, 0, true)); closed = false;
                            using (var verify = new ComBag())
                            {
                                dynamic sheets = verify.Own(((dynamic)outputBook).Worksheets);
                                if (Convert.ToInt32(sheets.Count) != Sheets.Count) throw new InvalidOperationException("保存后工作表数量发生变化。");
                                for (int i = 1; i <= Sheets.Count; i++)
                                    if (Convert.ToString(((dynamic)verify.Own(sheets.Item(i))).Name) != Sheets[i - 1])
                                        throw new InvalidOperationException("保存后工作表名称或顺序发生变化。");
                                dynamic sheet = verify.Own(sheets.Item(request.TargetSheet));
                                var headers = HeadersOn(sheet, request.TargetHeaderRow);
                                object cells = verify.Own(sheet.Cells);
                                if (Read(cells, request.TargetRow, CompanyWorkbookRules.Column(headers, request.TargetCodeField)) != record.Code
                                    || Read(cells, request.TargetRow, CompanyWorkbookRules.Column(headers, request.TargetNameField)) != record.FullName)
                                    throw new InvalidOperationException("保存后重开核验 Lead 字段失败。");
                            }
                            ((dynamic)outputBook).Close(false); closed = true;
                        }
                    }
                    File.Move(stage, record.OutputPath); // Atomic no-overwrite publication.
                }
                finally
                {
                    if (outputBook != null && !closed)
                    {
                        using (var quiet = new CompanyWorkbookQuiet(app))
                        { quiet.ExcludeClosingBook(outputBook); ((dynamic)outputBook).Close(false); }
                    }
                    if (File.Exists(stage)) File.Delete(stage);
                }
            }
        }
        public void Dispose()
        {
            try
            {
                if (book != null) using (var quiet = new CompanyWorkbookQuiet(app))
                { quiet.ExcludeClosingBook(book); ((dynamic)book).Close(false); book = null; }
            }
            finally
            {
                held.Dispose();
                if (book == null && Directory.Exists(folder)) Directory.Delete(folder, true);
            }
        }
    }
    internal sealed class CompanyWorkbookQuiet : IDisposable
    {
        private readonly dynamic app;
        private readonly List<Action> restore = new List<Action>();
        private readonly ComBag held = new ComBag();
        private object activeBook, activeSheet;
        internal void ExcludeClosingBook(object closingBook)
        {
            // A session-owned active book may be the only open workbook.
            // It must not be activated after we have just closed it.
            if (ReferenceEquals(activeBook, closingBook)) activeBook = activeSheet = null;
        }
        internal CompanyWorkbookQuiet(object application)
        {
            app = application;
            // Capture all required settings before changing any setting.
            bool alerts = Convert.ToBoolean(app.DisplayAlerts), events = Convert.ToBoolean(app.EnableEvents), screen = Convert.ToBoolean(app.ScreenUpdating), ask = Convert.ToBoolean(app.AskToUpdateLinks);
            object security = app.AutomationSecurity;
            try
            {
                activeBook = held.Own(app.ActiveWorkbook);
                activeSheet = held.Own(app.ActiveSheet);
                restore.Add(delegate { app.DisplayAlerts = alerts; }); app.DisplayAlerts = false;
                restore.Add(delegate { app.EnableEvents = events; }); app.EnableEvents = false;
                restore.Add(delegate { app.ScreenUpdating = screen; }); app.ScreenUpdating = false;
                restore.Add(delegate { app.AskToUpdateLinks = ask; }); app.AskToUpdateLinks = false;
                restore.Add(delegate { app.AutomationSecurity = (dynamic)security; }); app.AutomationSecurity = 3;
            }
            catch { Dispose(); throw; }
        }
        public void Dispose()
        {
            var failures = new List<string>();
            // Restore the user's context while our event suppression is still active.
            try { if (activeBook != null) ((dynamic)activeBook).Activate(); if (activeSheet != null) ((dynamic)activeSheet).Activate(); }
            catch (Exception error) { failures.Add("活动窗口：" + error.Message); }
            for (int i = restore.Count - 1; i >= 0; i--) try { restore[i](); } catch (Exception error) { failures.Add(error.Message); }
            restore.Clear();
            held.Dispose(); activeBook = activeSheet = null;
            if (failures.Count > 0) throw new CompanyWorkbookStateException("恢复宿主状态失败：" + string.Join("；", failures.ToArray()));
        }
    }
}
