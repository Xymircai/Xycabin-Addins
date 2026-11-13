using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 批量生成过程表 → 单体汇总批量链接
// 控件：btnSingleSummary
namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnSingleSummary(object control)
        {
            if (_busy) return;
            if (_summaryForm != null && !_summaryForm.IsDisposed) { _summaryForm.Activate(); return; }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("无法连接当前表格宿主。");
                SingleSummaryForm form = new SingleSummaryForm();
                _summaryWorkbook = new SummaryWorkbookBinding(_app);
                form.ClearWorkbookContext();
                RunSummaryWorkbookList(form);
                form.WorkbooksRequested += delegate { RunSummaryWorkbookList(form); };
                form.WorkbookRequested += delegate(string path) { RunSummaryWorkbookChoice(form, delegate { _summaryWorkbook.Choose(path); return true; }); };
                form.PresetRequested += delegate(BatchPreset value) { RunSummaryPreset(form, value); };
                form.CurrentRequested += delegate { RunSummaryWorkbookChoice(form, delegate { _summaryWorkbook.UseCurrent(); return true; }); };
                form.BrowseRequested += delegate { RunSummaryBrowse(form); };
                form.SheetsRequested += delegate { RunSummaryRefresh(form); };
                form.SelectionRequested += delegate(string kind) { RunSummarySelection(form, kind); };
                form.RangeRequested += delegate(string kind) { RunSummarySelection(form, kind, true); };
                form.PreviewRequested += delegate { RunSummaryPreview(form); };
                form.ExecuteRequested += delegate { RunSummaryExecute(form); };
                form.UndoRequested += delegate { RunSummaryUndo(form); };
                form.ConfirmRequested += delegate { RunSummaryConfirm(form); };
                form.FormClosed += delegate
                {
                    _summaryForm = null;
                    SummarySession session = _summarySession;
                    _summarySession = null;
                    if (session != null && session.CanUndo)
                        Log("单体汇总批量链接：窗口关闭时仍有可还原批次，备份表 " + session.BackupSheet + " 保留在工作簿里");
                    try { if (_summaryWorkbook != null) _summaryWorkbook.Dispose(); }
                    catch (Exception error) { Log("单体汇总批量链接：释放所选工作簿引用失败 " + error.Message); }
                    _summaryWorkbook = null;
                    form.Dispose();
                };
                _summaryForm = form;
                // 与表头匹配/链接助手一致：取宿主主窗口作 owner，不设 TopMost（R-25）
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception ex)
            {
                Log("单体汇总批量链接：" + ex);
                UiDialog.Show(ex.Message, "单体汇总批量链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public IPictureDisp GetImage_btnSingleSummary(object control)
        { return GetButtonImageCore("btnSingleSummary"); }
    }
}
