using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 筛选增强 → 按列筛选
// 控件：btnFilterColumns
namespace Excel_Iterate
{
    public partial class Connect
    {
        private const string ColumnUndoSource = "column-filter";

        public void OnFilterColumns(object control)
        {
            if (_busy) return;
            if (_columnForm != null && !_columnForm.IsDisposed) { _columnForm.Activate(); return; }
            _busy = true;
            ComColumnVisibilityHost host = null;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格并选中一行。");
                host = new ComColumnVisibilityHost(_app);
                System.Collections.Generic.IList<ColumnChoice> columns;
                string description;
                try { columns = host.Collect(); description = host.RangeDescription; }
                catch (Exception error)
                {
                    if (UndoState(ColumnUndoSource) == null) throw;
                    host.Dispose(); host = null;
                    columns = new ColumnChoice[0]; description = error.Message;
                }
                ColumnFilterForm form = new ColumnFilterForm(columns, description);
                ComColumnVisibilityHost bound = host;
                object application = _app;
                string context = form.Hint.Text;
                string recovery = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "ColumnRecovery");
                Action refreshUndo = delegate { form.Undo.SetState(UndoState(ColumnUndoSource)); };
                WorksheetUndoChanged += refreshUndo;
                refreshUndo();
                form.Undo.Click += delegate
                {
                    var undo = form.Undo.State;
                    if (undo == null) return;
                    try
                    {
                        form.Hint.Text = context + "\r\n" + UndoWorksheetEdit(ColumnUndoSource, undo.Id, application, recovery);
                        if (bound != null) form.RefreshVisibility(bound.RefreshVisibility());
                    }
                    catch (Exception error) { form.Hint.Text = context + "\r\n" + error.Message; }
                };
                form.ApplyRequested += delegate(System.Collections.Generic.IList<ColumnChoice> choices)
                {
                    if (_busy || bound == null) return;
                    _busy = true;
                    try
                    {
                        FormulaCommandResult result = ColumnVisibilityEngine.Apply(bound, choices, recovery, recordUndo: true);
                        RememberWorksheetEdit(ColumnUndoSource, result.Undo);
                        if (result.Changed > 0) form.RefreshVisibility(bound.RefreshVisibility());
                        form.Hint.Text = context + "\r\n" + result.Details;
                    }
                    catch (Exception error) { form.Hint.Text = context + "\r\n" + error.Message; Log(error.Message); }
                    finally { _busy = false; }
                };
                form.FormClosed += delegate
                {
                    WorksheetUndoChanged -= refreshUndo;
                    _columnForm = null;
                    try { if (bound != null) bound.Dispose(); } catch (Exception ex) { Log("列筛选引用释放：" + ex.Message); }
                    form.Dispose();
                };
                _columnForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
                host = null; // The non-modal form owns the snapshot until it closes.
            }
            catch (Exception ex) { Log("按列筛选：" + ex); UiDialog.Show(ex.Message, "按列筛选", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { if (host != null) host.Dispose(); _busy = false; }
        }

        public IPictureDisp GetImage_btnFilterColumns(object control) { return GetButtonImageCore("btnFilterColumns"); }
    }
}
