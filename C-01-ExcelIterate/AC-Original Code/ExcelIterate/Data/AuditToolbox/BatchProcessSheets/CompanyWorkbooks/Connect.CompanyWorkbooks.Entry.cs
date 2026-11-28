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
            if (_companyWorkbookForm != null && !_companyWorkbookForm.IsDisposed)
            {
                if (_companyWorkbookForm.WindowState == FormWindowState.Minimized)
                    _companyWorkbookForm.WindowState = FormWindowState.Normal;
                _companyWorkbookForm.Activate();
                return;
            }

            CompanyWorkbookForm form = null;
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("无法连接当前表格宿主。");
                form = new CompanyWorkbookForm
                {
                    CanRun = delegate { return !_busy; },
                    OpenTemplate = path => CompanyWorkbookHost.FromOpenOrDisk(_app, path),
                    OpenCurrentTemplate = () => CompanyWorkbookHost.FromCurrent(_app)
                };
                form.Structure.OpenSource = path => CompanyWorkbookHost.FromOpenOrDisk(_app, path);
                form.BusyChanged += delegate(bool busy) { _busy = busy; };
                form.FormClosed += delegate { _companyWorkbookForm = null; form.Dispose(); };
                _companyWorkbookForm = form;
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show();
                else form.Show(new WindowWrapper(owner));
            }
            catch (Exception error)
            {
                _companyWorkbookForm = null;
                if (form != null) form.Dispose();
                Log("单体模板生成：" + error);
                UiDialog.Show(error.Message, "单体模板生成", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public IPictureDisp GetImage_btnCompanyWorkbooks(object control) { return GetButtonImageCore("btnCompanyWorkbooks"); }
    }
}
