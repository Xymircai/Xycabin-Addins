using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 批量生成过程表 → 合并批量链接
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
                SingleSummaryForm form = new SingleSummaryForm(true);
                var view = (SummaryAreaView)form.Controls[0];
                var controller = new SummaryAreaHost(_app, form, view, delegate(bool value) { _busy = value; });
                form.FormClosed += delegate { controller.Dispose(); _summaryForm = null; form.Dispose(); };

                _summaryForm = form;
                // 与表头匹配/链接助手一致：取宿主主窗口作 owner，不设 TopMost（R-25）
                IntPtr owner = GetHostWindowHandle();
                if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception ex)
            {
                Log("合并批量链接：" + ex);
                UiDialog.Show(ex.Message, "合并批量链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public IPictureDisp GetImage_btnSingleSummary(object control)
        { return GetButtonImageCore("btnSingleSummary"); }
    }
}
