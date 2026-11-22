using System;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static WorksheetBatchForm worksheetBatchForm;
        private const string WorksheetBatchUndoSource = "worksheet-batch";

        public IPictureDisp GetImage_btnWorksheetBatch(object control)
        { return GetButtonImageCore("btnWorksheetBatch"); }

        public void OnWorksheetBatch(object control)
        {
            if (_busy) return;
            if (worksheetBatchForm != null && !worksheetBatchForm.IsDisposed) { worksheetBatchForm.Activate(); return; }
            if (!TryBeginConversion()) return;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                string name = null, address = null;
                using (var bag = new ComBag())
                {
                    try
                    {
                        dynamic app = _app;
                        object sheet = bag.Own((object)app.ActiveSheet);
                        if (sheet != null) name = (string)((dynamic)sheet).Name;
                        object selected = bag.Own((object)app.Selection);
                        if (selected != null) address = ComConversionHost.LocalRangeAddress(selected);
                    }
                    catch (Exception error) { UiDiagnostics.Record("worksheet-batch-selection", error); }
                }
                var form = new WorksheetBatchForm(name, address);
                worksheetBatchForm = form;
                object application = _app;
                Action refreshUndo = delegate { form.Undo.SetState(UndoState(WorksheetBatchUndoSource)); };
                WorksheetUndoChanged += refreshUndo;
                refreshUndo();
                form.FormClosed += delegate
                {
                    WorksheetUndoChanged -= refreshUndo;
                    if (worksheetBatchForm == form) worksheetBatchForm = null;
                };
                form.Undo.Click += delegate
                {
                    var undo = form.Undo.State;
                    if (undo == null) return;
                    try { form.ShowUndoResult(UndoWorksheetEdit(WorksheetBatchUndoSource, undo.Id, application, ConversionRecoveryDirectory)); }
                    catch (Exception error) { form.ShowError(error.Message); }
                };
                form.PickRange += delegate { PickWorksheetBatchRange(form, application); };
                form.Run += delegate(WorksheetBatchRequest request, WorksheetBatchReport preview)
                {
                    if (!TryBeginConversion()) throw new InvalidOperationException("另一项加载项操作正在执行，请稍后重试。");
                    try
                    {
                        var report = new WorksheetBatchHost(application).Run(request, preview, ConversionRecoveryDirectory);
                        RememberWorksheetEdit(WorksheetBatchUndoSource, report.Undo);
                        return report;
                    }
                    finally { EndConversion(); }
                };
                ShowFileForm(form);
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("worksheet-batch-open", error);
                CloseWorksheetBatchForm();
                UiDialog.Show(error.Message, "工作表批量操作", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { EndConversion(); }
        }

        private static void PickWorksheetBatchRange(WorksheetBatchForm form, object application)
        {
            if (!TryBeginConversion()) { form.ShowError("另一项加载项操作正在执行，请稍后重试。"); return; }
            try
            {
                form.Hide();
                using (var bag = new ComBag())
                {
                    object range = bag.Own((object)((dynamic)application).InputBox("请选择一个样例区域。将对同名工作表使用相同地址。",
                        "工作表批量操作 · 选择区域", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8));
                    if (range == null || range is bool) return;
                    dynamic sheet = bag.Own(((dynamic)range).Worksheet);
                    form.SetRange((string)sheet.Name, ComConversionHost.LocalRangeAddress(range));
                }
            }
            catch (Exception error) { form.ShowError(error.Message); }
            finally
            {
                EndConversion();
                if (!form.IsDisposed) ShowFileForm(form);
            }
        }

        private static void CloseWorksheetBatchForm()
        {
            WorksheetBatchForm form = worksheetBatchForm;
            worksheetBatchForm = null;
            if (form != null && !form.IsDisposed) form.Close();
        }
    }
}
