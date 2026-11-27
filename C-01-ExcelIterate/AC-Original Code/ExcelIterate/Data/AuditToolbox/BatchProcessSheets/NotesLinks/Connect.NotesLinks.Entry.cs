using System;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static NotesLinksForm _notesLinksForm;
        internal static void CloseNotesLinksForm(){if(_notesLinksForm!=null&&!_notesLinksForm.IsDisposed)_notesLinksForm.Shutdown();}
        public void OnNotesLinks(object control)
        {
            if(_busy)return;if(_notesLinksForm!=null&&!_notesLinksForm.IsDisposed){_notesLinksForm.Activate();return;}
            try
            {
                if(!EnsureApplication())throw new InvalidOperationException("无法连接当前表格宿主。");
                var form=new NotesLinksForm{Application=_app,CanRun=delegate{return !_busy;}};form.BusyChanged+=delegate(bool busy){_busy=busy;};form.FormClosed+=delegate{_notesLinksForm=null;form.Dispose();};_notesLinksForm=form;
                IntPtr owner=GetHostWindowHandle();if(owner==IntPtr.Zero)form.Show();else form.Show(new WindowWrapper(owner));
            }
            catch(Exception error){Log("Notes批量链接："+error);UiDialog.Show(error.Message,"Notes 批量链接",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
        }
        public IPictureDisp GetImage_btnNotesLinks(object control){return GetButtonImageCore("btnNotesLinks");}
    }
}
