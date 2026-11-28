using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    // Owns a newly-created, independent automation process, never a user host.
    internal sealed class HiddenWorkbookApplication : IDisposable
    {
        [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        private readonly ComBag references=new ComBag();
        internal object Application {get;private set;}
        private bool owns;
        internal static string HostProgId(string name,string path,string processName)
        {
            if(string.Equals(processName,"EXCEL",StringComparison.OrdinalIgnoreCase))return "Excel.Application";
            if(string.Equals(processName,"et",StringComparison.OrdinalIgnoreCase)||string.Equals(processName,"ket",StringComparison.OrdinalIgnoreCase))return "ket.Application";
            string identity=(name??"")+"|"+(path??"");
            return identity.IndexOf("wps",StringComparison.OrdinalIgnoreCase)>=0||identity.IndexOf("kingsoft",StringComparison.OrdinalIgnoreCase)>=0||identity.IndexOf("office6",StringComparison.OrdinalIgnoreCase)>=0?"ket.Application":(name??"").IndexOf("Excel",StringComparison.OrdinalIgnoreCase)>=0?"Excel.Application":"ket.Application";
        }
        internal static bool IsWps(object application)
        {
            string path="",process="";try{path=Convert.ToString(((dynamic)application).Path);}catch{}
            try{uint pid;GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd)),out pid);process=System.Diagnostics.Process.GetProcessById((int)pid).ProcessName;}catch{}
            return HostProgId(Convert.ToString(((dynamic)application).Name),path,process)=="ket.Application";
        }
        internal HiddenWorkbookApplication(object source)
        {
            object candidate=null;
            try
            {
                string progId=IsWps(source)?"ket.Application":"Excel.Application";
                Type type=Type.GetTypeFromProgID(progId);if(type==null)throw new InvalidOperationException("无法创建独立后台表格进程。");
                candidate=references.Own(Activator.CreateInstance(type));uint original,worker;
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)source).Hwnd)),out original);
                GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(((dynamic)candidate).Hwnd)),out worker);
                if(original==0||worker==0||original==worker)throw new InvalidOperationException("宿主未提供独立后台进程，已停止，避免隐藏或关闭你的操作窗口。");
                using(var bag=new ComBag()){dynamic books=bag.Own(((dynamic)candidate).Workbooks);if(Convert.ToBoolean(((dynamic)candidate).Visible)||Convert.ToInt32(books.Count)!=0)throw new InvalidOperationException("后台进程已含工作簿，未接管或关闭该进程。");}
                owns=true;Application=candidate;
                dynamic app=candidate;app.Visible=false;app.DisplayAlerts=false;app.EnableEvents=false;app.ScreenUpdating=false;app.AskToUpdateLinks=false;app.AutomationSecurity=3;
            }
            catch{try{if(owns&&candidate!=null)((dynamic)candidate).Quit();}finally{owns=false;Application=null;references.Dispose();}throw;}
        }
        public void Dispose()
        {try{if(owns&&Application!=null)((dynamic)Application).Quit();}finally{owns=false;Application=null;references.Dispose();}}
    }
}
