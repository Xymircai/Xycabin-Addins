using System;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static CompanyWorkbookForm _companyWorkbookForm;
        internal static void CloseCompanyWorkbookForm()
        {
            if (_companyWorkbookForm != null && !_companyWorkbookForm.IsDisposed) _companyWorkbookForm.Shutdown();
        }
        public void OnCompanyWorkbooks(object control)
        {
            if (_busy) return;
            if (_companyWorkbookForm != null && !_companyWorkbookForm.IsDisposed) { _companyWorkbookForm.Activate(); return; }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("无法连接当前表格宿主。");
                var form = new CompanyWorkbookForm();
                form.CanRun = delegate { return !_busy; };
                form.OpenTemplate = delegate(string path) { return new CompanyWorkbookHost(_app, path); };
                form.BusyChanged += delegate(bool busy) { _busy = busy; };
                form.FormClosed += delegate { _companyWorkbookForm = null; form.Dispose(); };
                _companyWorkbookForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception error) { Log("单体模板生成：" + error); UiDialog.Show(error.Message, "单体模板生成", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public IPictureDisp GetImage_btnCompanyWorkbooks(object control) { return GetButtonImageCore("btnCompanyWorkbooks"); }
    }
}
