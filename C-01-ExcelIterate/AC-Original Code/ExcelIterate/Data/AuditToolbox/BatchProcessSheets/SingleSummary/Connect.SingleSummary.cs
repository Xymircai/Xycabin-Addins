using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Text;

namespace Excel_Iterate
{
    // =====================================================================
    // 单体汇总表生成 —— 功能区接线与业务编排
    //
    // 全部 COM 调用都在 Ribbon 宿主线程执行；回调必须写在 Connect*.cs 内，
    // 否则配置器（configure_core.py）的回调存在性校验会失败。
    //
    // 编排顺序与需求一致：
    //   读 Lead（动态定位列）→ 候选行预填勾 → 确认窗 → 回写「选用」
    //   → 备份工作表 → 铺列（固定覆盖）→ 换链 → 表头 → 日志表 → 不保存
    // =====================================================================

    public partial class Connect
    {
        private static SingleSummaryForm _summaryForm;
        private static SummarySession _summarySession;
        private static SummaryWorkbookBinding _summaryWorkbook;

        private sealed class SummarySession
        {
            internal long WorkbookVersion;
            internal string TargetSheet = "";
            internal string LeadSheet = "";
            internal string BackupSheet = "";
            internal int HeaderRow = 1;
            internal int FirstColumn;
            internal int ColumnCount;
            internal int LastRow;
            internal SummaryTemplateProbe Probe;
            internal bool CanUndo;
        }

        internal static void CloseSummaryForm()
        {
            SingleSummaryForm form = _summaryForm;
            _summaryForm = null;
            if (form == null) return;
            try
            {
                if (!form.IsDisposed && form.IsHandleCreated)
                    form.BeginInvoke((MethodInvoker)delegate { try { form.Close(); } catch { } });
            }
            catch { try { form.Close(); } catch { } }
        }

        private static object RequireSummaryBook()
        {
            if (_summaryWorkbook == null) throw new InvalidOperationException("请先选择工作簿。");
            return _summaryWorkbook.RequireBook();
        }

        private static void ApplySelectedSummaryContext(SingleSummaryForm form)
        {
            SummaryRequest prior = form.BuildRequest();
            using (var bag = new ComBag())
            {
                object book = RequireSummaryBook();
                dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                var names = new List<string>();
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++)
                    using (var one = new ComBag()) names.Add(Convert.ToString(((dynamic)one.Own(sheets.Item(i))).Name));
                string lead = names.Contains(prior.LeadSheet) ? prior.LeadSheet : "";
                if (lead.Length == 0)
                    foreach (string name in names)
                        if (string.Equals(name, "lead", StringComparison.OrdinalIgnoreCase)) { lead = name; break; }
                string target = names.Contains(prior.TargetSheet) ? prior.TargetSheet : "";
                form.SetCompanies(null);
                form.SetContext(_summaryWorkbook.Path, names, lead, target, "");
                form.AppendLog("已载入所选工作簿：" + _summaryWorkbook.Path + "，共 " + names.Count + " 张表。请明确选择汇总工作表后预检。");
            }
        }

        private static void RunSummaryWorkbookList(SingleSummaryForm form)
        {
            if (_busy || (_summarySession != null && _summarySession.CanUndo)) return;
            _busy = true; form.SetBusy(true);
            try
            {
                var choices = _summaryWorkbook.Choices();
                if (_summaryWorkbook.Path.Length > 0)
                {
                    try { _summaryWorkbook.RequireBook(); }
                    catch (Exception error)
                    {
                        _summaryWorkbook.Clear(); _summarySession = null; form.ClearWorkbookContext();
                        form.AppendLog(error.Message);
                    }
                }
                form.SetWorkbooks(choices, _summaryWorkbook.Path);
                if (choices.Count == 0) form.AppendLog("当前宿主没有已打开的工作簿，可点击“浏览文件…”选择。");
            }
            catch (Exception error) { form.AppendLog("刷新工作簿列表失败：" + error.Message); }
            finally { form.SetBusy(false); _busy = false; }
        }

        private static void RunSummaryWorkbookChoice(SingleSummaryForm form, Func<bool> choose)
        {
            if (_busy || (_summarySession != null && _summarySession.CanUndo)) return;
            _busy = true; form.SetBusy(true);
            long version = _summaryWorkbook.Version;
            try
            {
                if (!choose()) return;
                _summarySession = null; form.ClearWorkbookContext();
                ApplySelectedSummaryContext(form);
                form.SetWorkbooks(_summaryWorkbook.Choices(), _summaryWorkbook.Path);
            }
            catch (Exception error)
            {
                if (_summaryWorkbook.Version != version)
                { _summaryWorkbook.Clear(); _summarySession = null; form.ClearWorkbookContext(); }
                // Restore the selector to the actual bound book after a rejected choice.
                try { form.SetWorkbooks(_summaryWorkbook.Choices(), _summaryWorkbook.Path); } catch { }
                Log("单体汇总表生成：选择工作簿失败 " + error);
                form.AppendLog("选择工作簿失败：" + error.Message);
            }
            finally { form.SetBusy(false); _busy = false; }
        }

        private static void RunSummaryBrowse(SingleSummaryForm form)
        {
            RunSummaryWorkbookChoice(form, delegate
            {
                string[] paths = HostDialogService.OpenFiles(form, "表格工作簿 (*.xlsx;*.xls;*.xlsm;*.et)|*.xlsx;*.xls;*.xlsm;*.et");
                if (paths.Length == 0) return false;
                if (paths.Length != 1) throw new InvalidOperationException("请每次只选择一个待处理工作簿。");
                bool opened = _summaryWorkbook.Browse(paths[0]);
                if (opened) form.AppendLog("已打开所选文件；生成后不自动保存或关闭，请自行核对并保存。");
                return true;
            });
        }

        private static void RunSummaryRefresh(SingleSummaryForm form)
        {
            if (_busy || (_summarySession != null && _summarySession.CanUndo)) return;
            _busy = true; form.SetBusy(true);
            try { _summarySession = null; ApplySelectedSummaryContext(form); }
            catch (Exception error)
            {
                _summaryWorkbook.Clear(); form.ClearWorkbookContext();
                form.AppendLog("刷新所选工作簿失败：" + error.Message);
            }
            finally { form.SetBusy(false); _busy = false; }
        }

        private static object FindSheet(object book, string name, ComBag bag)
        {
            if (book == null || string.IsNullOrEmpty(name)) return null;
            dynamic sheets = bag.Own(((dynamic)book).Worksheets);
            int count = Convert.ToInt32(sheets.Count, CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                dynamic sheet = bag.Own(sheets.Item(i));
                if (string.Equals(Convert.ToString(sheet.Name), name.Trim(), StringComparison.Ordinal)) return sheet;
            }
            return null;
        }

        // ---------------------------- 预检 ----------------------------

        private static void RunSummaryPreview(SingleSummaryForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                SummaryRequest request = form.BuildRequest();
                if (request.TemplateColumn < 1) throw new InvalidOperationException("模板列请填列字母，例如 D。");
                if (request.HeaderRow < 1) throw new InvalidOperationException("表头行必须是正整数。");
                string yearError = request.ValidateYears(); if (yearError != null) throw new InvalidOperationException(yearError);

                using (ComBag bag = new ComBag())
                {
                    object book = RequireSummaryBook();
                    object lead = FindSheet(book, request.LeadSheet, bag);
                    object target = FindSheet(book, request.TargetSheet, bag);
                    if (lead == null) throw new InvalidOperationException("找不到 Lead 工作表：" + request.LeadSheet);
                    if (target == null) throw new InvalidOperationException("找不到汇总工作表：" + request.TargetSheet);

                    SummaryLeadSnapshot snapshot = SummaryLeadReader.Read(lead, request.HeaderRow);
                    foreach (var company in snapshot.Companies)
                        foreach (var prior in request.Companies)
                            if (company.Row == prior.Row && company.Code == prior.Code && company.Name == prior.Name) { company.Selected = prior.Selected; break; }
                    form.SetCompanies(snapshot.Companies);
                    request.Companies.Clear(); request.Companies.AddRange(snapshot.Companies);

                    StringBuilder text = new StringBuilder();
                    foreach (string blocking in snapshot.Blocking) text.AppendLine("阻断：" + blocking);
                    if (snapshot.Blocking.Count > 0)
                    {
                        form.SetPreview("", false, text.ToString());
                        return;
                    }

                    int count = request.Chosen().Count;
                    if (count == 0) throw new InvalidOperationException("没有勾选任何公司，请先选择公司再预检。");
                    var columns = request.Columns();
                    int lastColumn = request.TemplateColumn + columns.Count - 1;
                    SummaryTemplateProbe probe = SummaryTemplateReader.Read(target, request.SheetHeaderRow, request.TemplateColumn, lastColumn, request.Placeholder);
                    // 占位名可配：模板列实际表头与默认值不一致时，先按实际值改一次再复检
                    if (!probe.Placeholder && probe.PlaceholderText.Trim().Length > 0
                        && !string.Equals(SummaryHeaders.Normalize(probe.PlaceholderText), SummaryHeaders.Normalize(request.Placeholder), StringComparison.OrdinalIgnoreCase))
                    {
                        form.SetPlaceholder(probe.PlaceholderText.Trim());
                        request.Placeholder = probe.PlaceholderText.Trim();
                        probe = SummaryTemplateReader.Read(target, request.SheetHeaderRow, request.TemplateColumn, lastColumn, request.Placeholder);
                        text.AppendLine("占位名已按模板列表头改为「" + request.Placeholder + "」（可再改）。");
                    }
                    foreach (string blocking in probe.Blocking) text.AppendLine("阻断：" + blocking);
                    foreach (string warning in probe.Warnings) text.AppendLine("提示：" + warning);

                    text.AppendLine("Lead：" + request.LeadSheet + "（表头第 "
                        + request.HeaderRow.ToString(CultureInfo.InvariantCulture) + " 行，"
                        + SummaryHeaders.CodeHeader + "=" + SummaryColumns.Letter(snapshot.CodeColumn)
                        + "，" + SummaryHeaders.NameHeader + "=" + SummaryColumns.Letter(snapshot.NameColumn)
                        + "，" + SummaryHeaders.PickHeader + "="
                        + (snapshot.PickColumn > 0 ? SummaryColumns.Letter(snapshot.PickColumn) : "（不存在，执行时在第 "
                            + snapshot.LastColumn.ToString(CultureInfo.InvariantCulture) + " 列右侧新增）") + "）");
                    text.AppendLine("候选公司 " + count.ToString(CultureInfo.InvariantCulture) + " 家，将生成 "
                        + SummaryColumns.Letter(request.TemplateColumn) + "～" + SummaryColumns.Letter(Math.Max(1, lastColumn))
                        + " 共 " + columns.Count.ToString(CultureInfo.InvariantCulture) + " 列（每家公司 " + request.YearCount + " 列，模板列也参与生成）。");
                    text.AppendLine("汇总表第 " + request.SheetHeaderRow.ToString(CultureInfo.InvariantCulture)
                        + " 行占位「" + (request.Placeholder.Length == 0 ? "（空=不校验）" : request.Placeholder) + "」："
                        + (probe.Placeholder ? "已确认" : "不一致，已阻断") + "；复制范围 "
                        + request.HeaderRow.ToString(CultureInfo.InvariantCulture) + "～"
                        + probe.LastRow.ToString(CultureInfo.InvariantCulture) + " 行，含公式 "
                        + probe.FormulaCells.ToString(CultureInfo.InvariantCulture) + " 个。");
                    if (probe.Tokens.Count > 0)
                        text.AppendLine("模板列外链文件名：" + string.Join("、", probe.Tokens.ToArray()));

                    string suggested = request.SourceFile;
                    if (suggested.Length == 0) suggested = GuessSourceFile(probe, snapshot.Companies);
                    if (suggested.Length == 0)
                        text.AppendLine("未能从公式推断模板公司全称，请在「模板公司全称」里填写（如 示例有限公司.xlsx）。");
                    else if (request.SourceFile.Length == 0)
                        text.AppendLine("已按公式推断模板公司全称：" + suggested + "（可改）。");

                    request.SourceFile = suggested;
                    if (request.Annual && suggested.Length > 0 && probe.Blocking.Count == 0)
                    {
                        foreach (string line in SummaryTemplateReader.AnnualPreview(book, target, request, probe)) text.AppendLine(line);
                        text.AppendLine("公司名在每列重复写入，年度按" + (request.DescendingYears ? "倒序" : "升序") + "排列；公司表头不合并。");
                    }

                    bool ok = probe.Blocking.Count == 0 && suggested.Length > 0;
                    form.SetPreview(suggested, ok, text.ToString());

                    _summarySession = new SummarySession();
                    _summarySession.WorkbookVersion = _summaryWorkbook.Version;
                    _summarySession.TargetSheet = request.TargetSheet;
                    _summarySession.LeadSheet = request.LeadSheet;
                    _summarySession.BackupSheet = "";
                    _summarySession.HeaderRow = request.SheetHeaderRow;
                    _summarySession.FirstColumn = request.TemplateColumn;
                    _summarySession.ColumnCount = columns.Count;
                    _summarySession.LastRow = probe.LastRow;
                    _summarySession.Probe = probe;
                    _summarySession.CanUndo = false;

                    Log("单体汇总表生成预检：" + request.TargetSheet + "，候选 " + count + " 家，模板列 "
                        + request.TemplateColumn + "，末行 " + probe.LastRow);
                }
            }
            catch (Exception ex)
            {
                Log("单体汇总表生成预检：" + ex);
                form.SetPreview("", false, "预检失败：" + ex.Message);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        /// <summary>从模板列的外链文件名推断「模板公司全称」：优先与候选全称同名的那一个。</summary>
        private static string GuessSourceFile(SummaryTemplateProbe probe, List<SummaryCompany> companies)
        {
            if (probe == null || probe.Tokens.Count == 0) return "";
            foreach (SummaryCompany company in companies)
                foreach (string token in probe.Tokens)
                    if (string.Equals(token, company.FileName, StringComparison.OrdinalIgnoreCase)) return token;
            return probe.Tokens.Count == 1 ? probe.Tokens[0] : "";
        }

        // ---------------------------- 执行 ----------------------------

        private static void RunSummaryExecute(SingleSummaryForm form)
        {
            if (_busy) return;
            SummaryRequest request = form.BuildRequest();
            if (_summaryWorkbook == null || _summarySession == null || _summarySession.WorkbookVersion != _summaryWorkbook.Version)
            { form.SetLog("所选工作簿尚未预检或已更换，请重新预检。"); return; }
            string invalid = request.Validate();
            if (invalid != null) { form.SetLog(invalid); return; }

            _busy = true;
            form.SetBusy(true);
            SummaryRunReport report = new SummaryRunReport();
            try
            {
                using (ComBag bag = new ComBag())
                {
                    object book = RequireSummaryBook();
                    object lead = FindSheet(book, request.LeadSheet, bag);
                    object target = FindSheet(book, request.TargetSheet, bag);
                    if (lead == null || target == null) throw new InvalidOperationException("Lead 或汇总工作表已不存在，请重新预检。");

                    // 二次核对：Lead 名单与模板列体检必须仍然成立（防止预检后被人改动）
                    SummaryLeadSnapshot snapshot = SummaryLeadReader.Read(lead, request.HeaderRow);
                    if (snapshot.Blocking.Count > 0) throw new InvalidOperationException(string.Join("；", snapshot.Blocking.ToArray()));
                    string drift = CompareSelection(snapshot, request.Chosen());
                    if (drift != null) throw new InvalidOperationException(drift);

                    List<SummaryCompany> chosen = request.Chosen();
                    var columns = request.Columns();
                    int lastColumn = request.TemplateColumn + columns.Count - 1;
                    SummaryTemplateProbe probe = SummaryTemplateReader.Read(target, request.SheetHeaderRow, request.TemplateColumn, lastColumn, request.Placeholder);
                    if (probe.Blocking.Count > 0) throw new InvalidOperationException(string.Join("；", probe.Blocking.ToArray()));
                    foreach (string line in SummaryTemplateReader.AnnualPreview(book, target, request, probe)) form.AppendLog(line);

                    string confirm = "工作簿：" + request.WorkbookName + "\r\n将覆盖 "
                        + SummaryColumns.Letter(request.TemplateColumn) + "～" + SummaryColumns.Letter(lastColumn)
                        + " 共 " + columns.Count.ToString(CultureInfo.InvariantCulture) + " 列（" + chosen.Count + " 家公司 × " + request.YearCount + "），第 1～"
                        + probe.LastRow.ToString(CultureInfo.InvariantCulture) + " 行。"
                        + "\r\n换链：" + request.SourceFile + " → 各列公司全称.xlsx" + (request.Annual ? "（另按列年度替换指定目录年份）" : "（只改方括号里的文件名）")
                        + "\r\n执行前自动备份工作表，勾选状态回写 Lead 的「" + SummaryHeaders.PickHeader + "」列；本功能不自动保存工作簿。"
                        + (request.Annual ? "\r\n年度范围：" + request.FirstYear + "—" + request.LastYear + "，" + (request.DescendingYears ? "倒序" : "升序") + "；年度表头第 " + request.YearHeaderRow + " 行，目录 " + request.YearPathPattern : "")
                        + "\r\n继续吗？";
                    if (UiDialog.Show(form, confirm, "确认生成单体汇总表",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK)
                    {
                        form.AppendLog("已取消，未写入任何内容。");
                        return;
                    }

                    RequireSummaryBook(); // Recheck identity after the isolated confirmation.
                    // 回写「选用」列（勾选 TRUE / 取消 FALSE），再生成
                    SummaryLeadReader.ApplySelection(snapshot, chosen);
                    using (HostQuiet quiet = new HostQuiet(_summaryWorkbook.Application))
                        SummaryLeadReader.WriteSelection(lead, snapshot);
                    form.AppendLog("已回写「" + SummaryHeaders.PickHeader + "」列：" + chosen.Count.ToString(CultureInfo.InvariantCulture) + " 家勾选。");

                    // 备份先做、且先把「可还原区间」落到会话：即使中途失败，catch 里也能按备份回滚
                    string backupSheet = SummaryWriter.Backup(_summaryWorkbook.Application, book, target);
                    if (_summarySession == null) _summarySession = new SummarySession();
                    _summarySession.TargetSheet = request.TargetSheet;
                    _summarySession.HeaderRow = request.SheetHeaderRow;
                    _summarySession.FirstColumn = request.TemplateColumn;
                    _summarySession.ColumnCount = columns.Count;
                    _summarySession.LastRow = probe.LastRow;
                    _summarySession.BackupSheet = backupSheet;
                    _summarySession.WorkbookVersion = _summaryWorkbook.Version;
                    _summarySession.CanUndo = true;

                    report = SummaryWriter.Run(_summaryWorkbook.Application, book, target, request, probe, backupSheet,
                        delegate(string status) { form.AppendLog(status); });

                    _summarySession.CanUndo = true;
                    form.SetUndoAvailable(true);

                    StringBuilder text = new StringBuilder();
                    text.AppendLine("生成完成：" + report.Columns.ToString(CultureInfo.InvariantCulture) + " 列（复制 "
                        + report.CopiedColumns.ToString(CultureInfo.InvariantCulture) + " 列），公式 "
                        + report.FormulaCells.ToString(CultureInfo.InvariantCulture) + " 个，换链 "
                        + report.RewrittenLinks.ToString(CultureInfo.InvariantCulture) + " 处。");
                    text.AppendLine("备份工作表：" + report.BackupSheet + "（隐藏）");
                    text.AppendLine("日志工作表：" + report.LogSheet);
                    text.AppendLine("工作簿未自动保存；确认无误请点「确认操作」，需要退回请点「取消操作」。");
                    foreach (string warning in report.Warnings) text.AppendLine("提示：" + warning);
                    foreach (SummaryLinkSample sample in report.Samples)
                        text.AppendLine("样例 " + sample.Cell + "：" + sample.Before + "  →  " + sample.After);
                    form.SetLog(text.ToString());
                    form.AppendLog("完成。");
                    Log("单体汇总表生成：" + request.TargetSheet + "，" + report.Columns + " 列，备份 " + report.BackupSheet);
                }
            }
            catch (Exception ex)
            {
                Log("单体汇总表生成执行：" + ex);
                form.AppendLog("执行失败：" + ex.Message + "；正在尝试还原，结果见下方日志。");
                TryRestoreSummary(form, false);
                UiDialog.Show(form, "执行失败：" + ex.Message + "\r\n（已尽量按备份表还原目标区间；详见日志。）",
                    "单体汇总表生成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        /// <summary>预检与执行之间，Lead 的勾选行内容必须没变（只核对编码/全称）。</summary>
        private static string CompareSelection(SummaryLeadSnapshot snapshot, List<SummaryCompany> chosen)
        {
            Dictionary<int, SummaryCompany> byRow = new Dictionary<int, SummaryCompany>();
            foreach (SummaryCompany company in snapshot.Companies) byRow[company.Row] = company;
            foreach (SummaryCompany company in chosen)
            {
                SummaryCompany actual;
                if (!byRow.TryGetValue(company.Row, out actual))
                    return "Lead 第 " + company.Row.ToString(CultureInfo.InvariantCulture) + " 行已不再是候选行，请重新预检。";
                if (!string.Equals(actual.Code, company.Code, StringComparison.Ordinal) ||
                    !string.Equals(actual.Name, company.Name, StringComparison.Ordinal))
                    return "Lead 第 " + company.Row.ToString(CultureInfo.InvariantCulture) + " 行的编码或全称已变化，请重新预检。";
            }
            return null;
        }

        // ---------------------------- 取消 / 确认 ----------------------------

        private static void RunSummaryUndo(SingleSummaryForm form)
        {
            if (_busy) return;
            _busy = true;
            form.SetBusy(true);
            try
            {
                TryRestoreSummary(form, true);
            }
            finally
            {
                form.SetBusy(false);
                _busy = false;
            }
        }

        private static void TryRestoreSummary(SingleSummaryForm form, bool ask)
        {
            SummarySession session = _summarySession;
            if (session == null || session.BackupSheet.Length == 0)
            {
                form.AppendLog("没有可还原的批次。");
                return;
            }
            if (ask && UiDialog.Show(form, "将按备份工作表还原 " + SummaryColumns.Letter(session.FirstColumn) + "～"
                    + SummaryColumns.Letter(session.FirstColumn + session.ColumnCount - 1)
                    + " 列（不删除日志表，不自动保存）。是否继续？",
                "取消操作", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button2) != DialogResult.OK)
            {
                form.AppendLog("已取消还原。");
                return;
            }
            try
            {
                using (ComBag bag = new ComBag())
                {
                    if (_summaryWorkbook == null || session.WorkbookVersion != _summaryWorkbook.Version)
                        throw new InvalidOperationException("还原批次不属于当前所选工作簿。");
                    object book = RequireSummaryBook();
                    object target = FindSheet(book, session.TargetSheet, bag);
                    if (target == null) throw new InvalidOperationException("汇总工作表已不存在。");
                    SummaryWriter.Restore(_summaryWorkbook.Application, book, target, session.BackupSheet,
                        session.FirstColumn, session.ColumnCount, session.LastRow, session.HeaderRow);
                }
                session.CanUndo = false;
                session.BackupSheet = "";
                form.SetUndoAvailable(false);
                form.AppendLog("已按备份表还原目标区间。");
                Log("单体汇总表生成：已还原 " + session.TargetSheet);
            }
            catch (Exception ex)
            {
                Log("单体汇总表生成还原失败：" + ex);
                session.CanUndo = true; form.SetUndoAvailable(true);
                form.AppendLog("还原失败：" + ex.Message + "；备份工作表仍在工作簿里，可手动恢复。");
            }
        }

        private static void RunSummaryConfirm(SingleSummaryForm form)
        {
            if (_summarySession != null)
            {
                _summarySession.CanUndo = false;
                form.SetUndoAvailable(false);
            }
            Log("单体汇总表生成：已确认，本批结束。");
        }
    }
}
