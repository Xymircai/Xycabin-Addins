using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif

internal static class HostDialogProbe
{
    [DllImport("user32.dll")] private static extern bool EnableWindow(IntPtr handle,bool enabled);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr handle);
    [DllImport("user32.dll")] private static extern IntPtr GetFocus();
    private static int checks;
    private static int activatedDialogs;
    private static void Check(bool value,string name) { if(!value)throw new InvalidOperationException(name);checks++;Console.WriteLine("PASS "+name); }
    // A modal Form must become Visible or WinForms cancels it immediately.
    // Keep this fixture transparent, off-screen and non-activating instead.
    private sealed class HiddenResultDialog : Form
    {
        private readonly System.Windows.Forms.Timer finish = new System.Windows.Forms.Timer { Interval = 15 };
        internal HiddenResultDialog(DialogResult answer, Action disposed)
        {
            Text = "{\"preset\":\"工作簿\"}";
            Opacity=0; ShowInTaskbar=false; StartPosition=FormStartPosition.Manual;
            Left=Top=-32000; Width=Height=1; FormBorderStyle=FormBorderStyle.None;
            finish.Tick += delegate { finish.Stop(); DialogResult = answer; };
            Disposed += delegate { finish.Dispose(); if(disposed!=null)disposed(); };
            Shown += delegate { finish.Start(); };
        }
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams { get { var value=base.CreateParams;value.ExStyle|=0x08000000;return value; } }
        protected override void OnActivated(EventArgs e) { Interlocked.Increment(ref activatedDialogs);base.OnActivated(e); }
    }
    private static void ResultChecks(IWin32Window owner, int host)
    {
        int created=0,read=0,disposed=0;bool sta=false;IntPtr focus=GetFocus();activatedDialogs=0;
        string result=HostDialogService.ShowIsolatedResult<string>(owner,delegate {
            created=Thread.CurrentThread.ManagedThreadId;sta=Thread.CurrentThread.GetApartmentState()==ApartmentState.STA;
            return new HiddenResultDialog(DialogResult.OK,delegate {disposed=Thread.CurrentThread.ManagedThreadId;});
        },delegate(Form dialog) {read=Thread.CurrentThread.ManagedThreadId;return dialog.Text;},null);
        Check(result=="{\"preset\":\"工作簿\"}","isolated chooser returns plain preset data");
        Check(sta&&created!=host&&read==created&&disposed==created,"chooser creation, result extraction and disposal stay on one worker STA");
        int reads=0;
        result=HostDialogService.ShowIsolatedResult<string>(owner,delegate {return new HiddenResultDialog(DialogResult.Cancel,null);},delegate(Form dialog) {reads++;return dialog.Text;},null);
        Check(result==null&&reads==0,"cancelled chooser never reads or applies a result");
        result=HostDialogService.ShowIsolatedResult<string>(owner,delegate {return new HiddenResultDialog(DialogResult.OK,null);},delegate(Form dialog) {throw new InvalidOperationException("private chooser data");},null);
        Check(result==null&&IsWindowEnabled(owner.Handle),"result extraction failure cancels and restores the owner");
        int nestedThread=0;
        result=HostDialogService.ShowIsolatedResult<string>(owner,delegate {return new HiddenResultDialog(DialogResult.OK,null);},delegate(Form dialog) {
            int worker=Thread.CurrentThread.ManagedThreadId;
            return HostDialogService.ShowIsolatedResult<string>(dialog,delegate {nestedThread=Thread.CurrentThread.ManagedThreadId;return new HiddenResultDialog(DialogResult.OK,null);},delegate(Form nested) {return nestedThread==worker?nested.Text:null;},null);
        },null);
        Check(result=="{\"preset\":\"工作簿\"}"&&nestedThread!=host,"nested chooser stays on the already isolated UI thread");
        Check(GetFocus()==focus&&activatedDialogs==0&&IsWindowEnabled(owner.Handle),"chooser paths preserve host focus and release their owner");
    }
    [STAThread] public static int Main()
    {
        Console.OutputEncoding=System.Text.Encoding.UTF8;
        string sandbox=Path.Combine(Path.GetTempPath(),"HostDialogProbe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(sandbox);UiDiagnostics.RootOverride=sandbox;
        int main=Thread.CurrentThread.ManagedThreadId;IntPtr focus=GetFocus();bool success=false;
        try
        {
            HostDialogService.Register();
            using(var parking=new Panel())using(var owner=new Form {TopLevel=false})using(var dispatcher=new Control())
            {
                parking.Controls.Add(owner);var parent=parking.Handle;var handle=owner.Handle;var dispatch=dispatcher.Handle;
                SynchronizationContext.SetSynchronizationContext(null);
                int invoked=0,worker=0;bool disabled=false;int nested=0;IntPtr owned=IntPtr.Zero;
                int value=HostDialogService.Run<int>(owner,uiOwner=>{
                    worker=Thread.CurrentThread.ManagedThreadId;
                    Check(worker!=main&&Thread.CurrentThread.GetApartmentState()==ApartmentState.STA,"confirmation uses independent STA");
                    Check(!ReferenceEquals(uiOwner,owner),"UI receives HWND wrapper rather than host Form");
                    owned=uiOwner.Handle;disabled=!IsWindowEnabled(owned);
                    dispatcher.BeginInvoke((MethodInvoker)delegate {invoked++;nested=HostDialogService.Run<int>(null,unused=>99,-1);});
                    var until=DateTime.UtcNow.AddSeconds(2);while(Volatile.Read(ref invoked)==0&&DateTime.UtcNow<until)Thread.Sleep(2);
                    return 42;
                },-1);
                Check(value==42&&worker!=main,"synchronous business result returns to original STA");
                Check(invoked==1&&nested==-1,"native COM wait pumps host messages and blocks nested confirmations");
                Check(disabled&&IsWindowEnabled(owned),"only owned HWND is disabled and restored");
                Check(SynchronizationContext.Current==null&&GetFocus()==focus,"host context and desktop focus are preserved");
                EnableWindow(owned,false);
                HostDialogService.Run<int>(owner,unused=>1,0);
                Check(!IsWindowEnabled(owned),"pre-disabled owner remains disabled");EnableWindow(owned,true);
                value=HostDialogService.Run<int>(owner,unused=>{throw new InvalidOperationException("private input");},-1);
                Check(value==-1&&IsWindowEnabled(owned),"UI failure cancels safely and restores owner");
                Check(!UiDiagnostics.Text.Contains("private input"),"shared dialog diagnostics exclude exception content");
                int loops=0;
                HostDialogService.Run<int>(owner,unused=>{
                    using(var hidden=new Form {TopLevel=false})using(var container=new Panel())using(var timer=new System.Windows.Forms.Timer {Interval=15})
                    {container.Controls.Add(hidden);using(var context=new ApplicationContext(hidden)) {hidden.Show();timer.Tick+=delegate{loops++;hidden.Close();};timer.Start();Application.Run(context);}}
                    return 1;
                },0);
                Check(loops==1&&IsWindowEnabled(owned),"real hidden UI loop exits without leaving owner disabled");
                ResultChecks(owner,main);
                using(var entered=new ManualResetEvent(false))
                {
                    HostDialogService.WaitForProbe=delegate {entered.WaitOne(2000);return unchecked((int)0x80004005);};
                    try
                    {
                        value=HostDialogService.Run<int>(owner,unused=>{
                            using(var hidden=new Form {TopLevel=false})using(var container=new Panel())
                            {container.Controls.Add(hidden);using(var context=new ApplicationContext(hidden)){hidden.Show();entered.Set();Application.Run(context);}}
                            return 7;
                        },-1);
                    }
                    finally {HostDialogService.WaitForProbe=null;}
                    var deadline=DateTime.UtcNow.AddSeconds(5);while(HostDialogService.ActiveWorkers!=0&&DateTime.UtcNow<deadline){Application.DoEvents();Thread.Sleep(5);}
                    Check(value==-1&&HostDialogService.ActiveWorkers==0&&IsWindowEnabled(owned),"native wait failure cancels worker UI and releases owner");
                }
            }
            success=true;Console.WriteLine("HostDialogProbe: "+checks+" PASS");return 0;
        }
        catch(Exception error){Console.WriteLine("FAIL "+error);return 1;}
        finally{if(success&&Path.GetDirectoryName(Path.GetFullPath(sandbox))==Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar))Directory.Delete(sandbox,true);}
    }
}
