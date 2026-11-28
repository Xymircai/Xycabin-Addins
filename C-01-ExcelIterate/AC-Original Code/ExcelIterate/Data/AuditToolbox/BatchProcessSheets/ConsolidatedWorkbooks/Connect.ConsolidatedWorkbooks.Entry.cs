using System;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static ConsolidatedWorkbookForm _consolidatedWorkbookForm;
        internal static void CloseConsolidatedWorkbookForm(){if(_consolidatedWorkbookForm!=null&&!_consolidatedWorkbookForm.IsDisposed)_consolidatedWorkbookForm.Shutdown();}
        public void OnConsolidatedWorkbooks(object control)
        {
            if(_busy)return;if(_consolidatedWorkbookForm!=null&&!_consolidatedWorkbookForm.IsDisposed){_consolidatedWorkbookForm.Activate();return;}
            try
            {
                if(!EnsureApplication())throw new InvalidOperationException("无法连接当前表格宿主。");
                var form=new ConsolidatedWorkbookForm();form.CanRun=delegate{return !_busy;};form.OpenTemplate=delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(_app,path);};form.OpenCurrentTemplate=delegate{return CompanyWorkbookHost.FromCurrent(_app);};form.Structure.OpenSource=delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(_app,path);};
                form.RememberConfigurationEnabled=true;form.BusyChanged+=delegate(bool busy){_busy=busy;};form.FormClosed+=delegate{_consolidatedWorkbookForm=null;form.Dispose();};_consolidatedWorkbookForm=form;form.RestoreLastConfiguration();
                IntPtr owner=GetHostWindowHandle();if(owner==IntPtr.Zero)form.Show();else form.Show(new WindowWrapper(owner));
            }
            catch(Exception error){Log("合并模板生成："+error);UiDialog.Show(error.Message,"合并模板生成",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
        public IPictureDisp GetImage_btnConsolidatedWorkbooks(object control){return GetButtonImageCore("btnConsolidatedWorkbooks");}
    }
}
