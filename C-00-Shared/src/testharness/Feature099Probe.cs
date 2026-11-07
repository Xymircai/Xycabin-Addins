using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif

partial class Feature099Probe
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]static extern bool IsWindowVisible(IntPtr handle);
    [System.Runtime.InteropServices.DllImport("user32.dll")]static extern IntPtr GetFocus();
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]static extern IntPtr GetStdHandle(int channel);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]static extern bool SetStdHandle(int channel,IntPtr value);
    sealed class PixelSampler {
        readonly byte[] data;readonly int width;
        internal PixelSampler(Bitmap bitmap){width=bitmap.Width;data=new byte[width*bitmap.Height*4];var bits=bitmap.LockBits(new Rectangle(0,0,width,bitmap.Height),ImageLockMode.ReadOnly,PixelFormat.Format32bppArgb);try{for(int y=0;y<bitmap.Height;y++)Marshal.Copy(IntPtr.Add(bits.Scan0,y*bits.Stride),data,y*width*4,width*4);}finally{bitmap.UnlockBits(bits);}}
        internal Color Get(int x,int y){int i=(y*width+x)*4;return Color.FromArgb(data[i+3],data[i+2],data[i+1],data[i]);}
        internal static void Verify(){using(var bitmap=new Bitmap(13,7)){for(int y=0;y<bitmap.Height;y++)for(int x=0;x<bitmap.Width;x++)bitmap.SetPixel(x,y,Color.FromArgb(255,x*19,y*31,(x+y)*11));var sample=new PixelSampler(bitmap);for(int y=0;y<bitmap.Height;y++)for(int x=0;x<bitmap.Width;x++)Check(sample.Get(x,y).ToArgb()==bitmap.GetPixel(x,y).ToArgb(),"pixel buffer matches reference");}}
    }
    static int checks;static string sandbox;static bool showUi;
    static void Check(bool ok,string message){if(!ok)throw new Exception(message);checks++;}
    static void Reject(Action action,string message){bool failed=false;try {action();}catch{failed=true;}Check(failed,message);}

    static void SessionAndThread() {
        string path=Path.Combine(sandbox,"format-session.xml");var store=new LocalCustomization(path);var initial=store.Data.Copy();initial.Names["btnAgentNames"]="initial name";initial.Orders["menuAgentCustomization"]=new List<string>{"btnAgentSort","btnAgentNames"};initial.Formats["probe::window"]=new FormatRule {Text="initial",Width96=180,Height96=40};store.Save(initial);FormatSession.SourceOverride=path;
        var existing=new Form();var frozen=FormatSession.Capture(existing);Check(frozen.Formats["probe::window"].Text=="initial","form snapshot initialized from disk");
        var next=store.Data.Copy();next.Formats["probe::window"].Text="saved";next.Formats["probe::window"].Width96=240;store.Save(next);next.Names["btnAgentNames"]="name change";store.Save(next);next.Orders["menuAgentCustomization"]=new List<string>{"btnAgentNames","btnAgentSort"};store.Save(next);
        Check(FormatSession.Data.Formats["probe::window"].Text=="saved"&&FormatSession.Capture(existing).Formats["probe::window"].Width96==180,"saved configuration updates latest data without mutating existing form");using(var reopened=new Form())Check(FormatSession.Capture(reopened).Formats["probe::window"].Width96==240,"reopened form reads saved dimensions in same process");existing.Dispose();
        var domain=AppDomain.CreateDomain("format-session-probe");try{var reader=(FormatDomainReader)domain.CreateInstanceAndUnwrap(Assembly.GetExecutingAssembly().FullName,typeof(FormatDomainReader).FullName);Check(reader.Read(path)=="saved|240","second AppDomain reads latest saved data");}finally{AppDomain.Unload(domain);}
        using(var child=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,"--format-read \""+path+"\""){UseShellExecute=false,CreateNoWindow=true,RedirectStandardOutput=true,RedirectStandardError=true})) {
            Check(child.WaitForExit(10000),"new process format check exits");Check(child.ExitCode==0&&child.StandardOutput.ReadToEnd().Trim()=="saved|240","new process reads newly saved format");
        }
        Check(FormatSession.Data.Names["btnAgentNames"]=="name change"&&FormatSession.Data.Orders["menuAgentCustomization"][0]=="btnAgentNames","name and order update in current host process");
        var sizesBefore=FormatSession.Data.WindowSizes;Check(sizesBefore!=null,"window sizes included in process snapshot");
        FormatEditorThread.HiddenForProbe=true;var focus=GetFocus();
        try {
            for(int cycle=0;cycle<2;cycle++){
                FormatEditorThread.Open(IntPtr.Zero);var until=DateTime.UtcNow.AddSeconds(10);while(FormatEditorThread.EditorHandle==IntPtr.Zero&&FormatEditorThread.ThreadId!=0&&DateTime.UtcNow<until)Thread.Sleep(20);
                var handle=FormatEditorThread.EditorHandle;int thread=FormatEditorThread.ThreadId;Check(handle!=IntPtr.Zero&&thread!=Thread.CurrentThread.ManagedThreadId,"editor created on independent STA loop "+cycle);Check(!IsWindowVisible(handle)&&GetFocus()==focus,"background editor check has no visible top-level form or focus request "+cycle);
                FormatEditorThread.Open(IntPtr.Zero);Check(FormatEditorThread.ThreadId==thread,"repeat open retains single editor thread "+cycle);
                FormatEditorThread.Close();until=DateTime.UtcNow.AddSeconds(10);while(FormatEditorThread.ThreadId!=0&&DateTime.UtcNow<until)Thread.Sleep(20);Check(FormatEditorThread.ThreadId==0&&FormatEditorThread.EditorHandle==IntPtr.Zero,"editor thread closes and releases resources "+cycle);
            }
            FormatEditorThread.CreatedForProbe=()=>FormatEditorThread.Close();
            FormatEditorThread.Open(IntPtr.Zero);var closedBy=DateTime.UtcNow.AddSeconds(10);while(FormatEditorThread.ThreadId!=0&&DateTime.UtcNow<closedBy)Thread.Sleep(20);
            Check(FormatEditorThread.ThreadId==0&&FormatEditorThread.EditorHandle==IntPtr.Zero,"close before editor handle creation releases STA loop");
        }finally{FormatEditorThread.CreatedForProbe=null;FormatEditorThread.Close();FormatEditorThread.HiddenForProbe=false;}
    }
    static void Pump(Task task) {var until=DateTime.UtcNow.AddSeconds(30);while(!task.IsCompleted&&DateTime.UtcNow<until){Application.DoEvents();Thread.Sleep(2);}Check(task.IsCompleted,"workflow finishes without a blocked UI");task.GetAwaiter().GetResult();}
    static void LiveRefresh(){
        var ribbon=new FakeRibbon();LiveCustomization.Stop();RibbonCustomization.Register(ribbon);
        var store=new LocalCustomization(FormatSession.SourceOverride);var data=store.Data.Copy();data.Names["btnAgentNames"]="live saved name";store.Save(data);
        var until=DateTime.UtcNow.AddSeconds(3);while(ribbon.Count==0&&DateTime.UtcNow<until){Application.DoEvents();Thread.Sleep(10);}Check(ribbon.Count>0,"atomic save refreshes a registered Ribbon on the UI thread");
        var appearance=FormatSession.Data;int before=ribbon.Count;store=new LocalCustomization(FormatSession.SourceOverride);data=store.Data.Copy();data.Formats["probe::window"].Text="appearance only";store.Save(data);
        until=DateTime.UtcNow.AddMilliseconds(450);while(DateTime.UtcNow<until){Application.DoEvents();Thread.Sleep(10);}Check(ribbon.Count==before,"format changes do not invalidate menus");
        RibbonCustomization.Clear();
        using(var form=new CommandCustomizationForm("name",true)){int saved=0;CustomizationSaving.SavedForProbe=()=>saved++;try{CustomizationSaving.AfterSave(form);}finally{CustomizationSaving.SavedForProbe=null;}Check(saved==1&&form.IsDisposed,"saving closes only the editor and requires no restart adapter");}
        Check(!typeof(Connect).Assembly.GetManifestResourceNames().Contains("WpsShared.RestartHelper.exe"),"new addin contains no restart helper");
    }
#if !WPS_WORD
    sealed class DelayedOcr:IFileOcr {public async Task<List<OcrSegment>> RecognizeAsync(string path,int page,CancellationToken token){await Task.Delay(400,token);return new List<OcrSegment>{new OcrSegment {Text="磅单：001",NoLocation=true,Page=page}};}public void Dispose(){}}
    static void LiveCloudWorkflow(string folder){
        FormatSession.SourceOverride=Path.Combine(sandbox,"live.xml");DialogSizeStore.RootOverride=sandbox;
        string key=Environment.GetEnvironmentVariable("XY_BAIDU_API"),secret=Environment.GetEnvironmentVariable("XY_BAIDU_SECRET");Check(!string.IsNullOrWhiteSpace(key)&&!string.IsNullOrWhiteSpace(secret),"live credentials supplied in process environment only");
        Control.CheckForIllegalCrossThreadCalls=true;
        foreach(int mode in new[]{0,1,2,3})using(var form=new FileNamesForm(new FileNameWorksheet(new FakeApplication(new FakeSheet(new FakeWorkbook())))))using(var parent=new Panel()){
            form.TopLevel=false;parent.Controls.Add(form);form.Show();var all=DialogFormatting.Walk(form).ToList();
            all.OfType<TextBox>().Single(c=>c.Name=="folderInput").Text=folder;form.SetPrefixRules(new List<OcrPrefixRule>{new OcrPrefixRule {Text="磅单号："}});all.OfType<TextBox>().Single(c=>c.Name=="apiKey").Text=key;all.OfType<TextBox>().Single(c=>c.Name=="secretKey").Text=secret;all.OfType<CheckBox>().Single(c=>c.Name=="useOcr").Checked=true;all.OfType<ComboBox>().Single(c=>c.Name=="cloudPrecision").SelectedIndex=mode;form.ConfirmUpload=_=>true;
            var context=SynchronizationContext.Current;Task task;try{SynchronizationContext.SetSynchronizationContext(null);task=(Task)typeof(FileNamesForm).GetMethod("Run",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{false});}finally{SynchronizationContext.SetSynchronizationContext(context);}
            Pump(task);var grid=all.OfType<ResultGrid>().Single();int matches=0;foreach(DataGridViewRow row in grid.Rows){var item=(FileNameEntry)row.Tag;string expected=item.Name.Contains("1-2")?"PD202512020015":"PD202512020009";if(item.Feature==expected&&item.Status=="已提取")matches++;}
            all.OfType<TextBox>().Single(c=>c.Name=="apiKey").Clear();all.OfType<TextBox>().Single(c=>c.Name=="secretKey").Clear();using(var bitmap=new Bitmap(grid.Width,grid.Height))grid.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));Check(matches==grid.Rows.Count&&matches>0&&!grid.PaintFailed,"real cloud workflow returns expected rows without painting error "+mode);
            form.Close();Check(form.IsDisposed,"real cloud workflow closes without exception "+mode);Console.WriteLine("LIVE_UI mode="+mode+" rows="+matches+" expected=True paint=True close=True");
        }
        Console.WriteLine("Live cloud pipeline PASS; fake worksheet and hidden WinForms, not WPS process");
    }
    static void Workflow(){
        string folder=Path.Combine(sandbox,"workflow");Directory.CreateDirectory(folder);using(var image=new Bitmap(100,60))image.Save(Path.Combine(folder,"test.png"));
        bool prior=Control.CheckForIllegalCrossThreadCalls;Control.CheckForIllegalCrossThreadCalls=true;
        try{foreach(bool custom in new[]{false,true})foreach(int mode in new[]{0,1,2,3}){
            var sheet=new FakeSheet(new FakeWorkbook());var app=new FakeApplication(sheet);
            using(var form=new FileNamesForm(new FileNameWorksheet(app)))using(var parent=new Panel()){
                form.TopLevel=false;parent.Controls.Add(form);form.Show();var all=DialogFormatting.Walk(form).ToList();
                all.OfType<TextBox>().Single(c=>c.Name=="folderInput").Text=folder;form.SetPrefixRules(new List<OcrPrefixRule>{new OcrPrefixRule {Text="磅单"}});all.OfType<TextBox>().Single(c=>c.Name=="apiKey").Text="mock";all.OfType<TextBox>().Single(c=>c.Name=="secretKey").Text="mock";all.OfType<CheckBox>().Single(c=>c.Name=="useOcr").Checked=true;all.OfType<ComboBox>().Single(c=>c.Name=="cloudPrecision").SelectedIndex=mode;
                form.ConfirmUpload=_=>true;form.OcrFactory=(key,secret,variant)=>new BaiduFileOcr(key,secret,variant,new MockHttp {Basic=!OcrModes.Positioned(variant)});
                var context=SynchronizationContext.Current;Task task;try{SynchronizationContext.SetSynchronizationContext(custom?new SynchronizationContext():null);task=(Task)typeof(FileNamesForm).GetMethod("Run",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{false});}finally{SynchronizationContext.SetSynchronizationContext(context);}
                Pump(task);var grid=all.OfType<ResultGrid>().Single();Check(grid.Rows.Count==1&&(string)grid.Rows[0].Cells["feature"].Value=="001","complete OCR workflow returns to STA without host synchronization context "+mode+" "+custom);
                using(var bitmap=new Bitmap(grid.Width,grid.Height))grid.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));Check(!grid.PaintFailed,"successful workflow has no paint failure");
                form.OcrFactory=(key,secret,variant)=>new DelayedOcr();task=(Task)typeof(FileNamesForm).GetMethod("Run",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{false});form.Close();Pump(task);var until=DateTime.UtcNow.AddSeconds(2);while(!form.IsDisposed&&DateTime.UtcNow<until)Application.DoEvents();Check(form.IsDisposed,"close safely cancels and disposes once");
            }
        }}finally{Control.CheckForIllegalCrossThreadCalls=prior;}
        using(var grid=new ResultGrid {Size=new Size(400,200)}){
            grid.PaintForProbe=()=>{throw new ArgumentException("sensitive-message-must-not-be-logged");};grid.Columns.Add("sample","sample");grid.Rows.Add();using(var bitmap=new Bitmap(400,200))grid.DrawToBitmap(bitmap,new Rectangle(0,0,400,200));Check(grid.PaintFailed&&UiDiagnostics.Text.Contains("file.names-paint")&&!UiDiagnostics.Text.Contains("sensitive-message"),"painting failure is contained and diagnostics redact exception messages");
        }
        string previous=UiDiagnostics.RootOverride;UiDiagnostics.RootOverride=Path.Combine(sandbox,"not-directory");File.WriteAllText(UiDiagnostics.RootOverride,"fixture");UiDiagnostics.Record("write-failure",new ArgumentException());Check(UiDiagnostics.Text.Contains("diagnostic-write"),"failed disk logging remains available in memory");UiDiagnostics.RootOverride=previous;
    }
#endif
    static void Notes() {
        RequirementsForm.RootOverride=sandbox;
        try {
            using(var first=new RequirementsForm(false))using(var stale=new RequirementsForm(false)) {
                var body=DialogFormatting.Walk(first).OfType<TextBox>().Single();var other=DialogFormatting.Walk(stale).OfType<TextBox>().Single();
                Check(body.Text=="","notes help is not saved user data");body.Text="测试需求\r\n第二行";
                var save=typeof(RequirementsForm).GetMethod("Save",BindingFlags.NonPublic|BindingFlags.Instance);save.Invoke(first,null);
                string path=Path.Combine(sandbox,typeof(Connect).Assembly.GetName().Name,"Requirements.txt");
                Check(WpsShared.SourcePaths.ReadAllText(path,Encoding.UTF8)==body.Text,"notes UTF-8 round trip");other.Text="旧窗口";Reject(()=>save.Invoke(stale,null),"stale notes rejected");
                typeof(RequirementsForm).GetMethod("LoadNotes",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(stale,null);Check(other.Text==body.Text,"notes reload latest content");
                other.Text="删除后的旧草稿";File.Delete(path);Reject(()=>save.Invoke(stale,null),"deleted notes conflict detected");
                stale.CloseForShutdown();string recovery=WpsShared.SourcePaths.GetFiles(Path.GetDirectoryName(path),"Requirements.txt.unsaved-*.txt").Single();Check(WpsShared.SourcePaths.ReadAllText(recovery,Encoding.UTF8)=="删除后的旧草稿"&&!WpsShared.SourcePaths.Exists(path),"shutdown preserves unsaved draft without overwriting saved notes");Check(stale.IsDisposed,"shutdown releases notes window");File.Delete(recovery);
            }
        }finally {RequirementsForm.RootOverride=null;}
    }
    static void Settings() {
        var path=Path.Combine(sandbox,"custom.xml");var first=new LocalCustomization(path);var other=new LocalCustomization(path);var data=first.Data.Copy();data.Names["test"]="示例";data.Orders["menu"]=new List<string>{"b","a"};first.Save(data);
        Check(new LocalCustomization(path).Data.Names["test"]=="示例","XML settings round trip");Reject(()=>other.Save(other.Data),"stale settings cannot overwrite");
        File.WriteAllText(path,"<customization version=\"99\">future</customization>");var future=new LocalCustomization(path);Reject(()=>future.Save(future.Data),"future settings remain intact");
        Check(WpsShared.SourcePaths.ReadAllText(path).Contains("99"),"future data preserved");File.WriteAllText(path,"<!DOCTYPE a [<!ENTITY x SYSTEM 'file:///C:/Windows/win.ini'>]><customization version=\"1\">&x;</customization>");var corrupt=new LocalCustomization(path);Reject(()=>corrupt.Save(corrupt.Data),"DTD cannot be read or overwritten");
        var copy=data.Copy();copy.Names["test"]="changed";Check(data.Names["test"]=="示例","draft isolated from saved settings");
    }
    static void Ribbon() {
        var saved=RibbonCustomization.Store.Data.Copy();try {
            RibbonCustomization.Store.Data=new CustomizationData();RibbonCustomization.DataForProbe=RibbonCustomization.Store.Data;var command=RibbonCustomization.Find("menuAgentCustomization");Check(command!=null,"customization menu keeps ID");
            var children=command.ChildNodes.OfType<XmlElement>().Select(e=>e.GetAttribute("id")).ToList();Check(children.SequenceEqual(new[]{"btnAgentSort","btnAgentNames","btnAgentFormats","btnAgentDialogSizes"}),"default child order");
            RibbonCustomization.Store.Data.Orders["menuAgentCustomization"]=children.AsEnumerable().Reverse().ToList();RibbonCustomization.Store.Data.Names["btnAgentNames"]="改名示例";
            var content=new XmlDocument();content.LoadXml(RibbonCustomization.Content("menuAgentCustomization"));Check(content.DocumentElement.FirstChild.Attributes["id"].Value=="btnAgentDialogSizes","dynamic content uses saved order");
            Check(content.OuterXml.Contains("改名示例"),"dynamic content uses current labels");var dynamic=RibbonCustomization.Effective(true);Check(dynamic.SelectNodes("//*[local-name()='dynamicMenu']").Count>0,"root menus are dynamic");
            Check(dynamic.SelectNodes("//*[local-name()='button' and @label]").Count==0,"root buttons use dynamic labels");Check(RibbonCustomization.Find("btnAgentOcrPlaceholder").GetAttribute("enabled")=="false","experimental OCR remains disabled");
            foreach(XmlElement node in dynamic.SelectNodes("//*[@getLabel]"))Check(typeof(Connect).GetMethod(node.GetAttribute("getLabel"))!=null,"dedicated label callback "+node.GetAttribute("id"));
            foreach(XmlElement node in dynamic.SelectNodes("//*[@getContent]"))Check(typeof(Connect).GetMethod(node.GetAttribute("getContent"))!=null,"dedicated content callback "+node.GetAttribute("id"));
            var ribbon=new FakeRibbon();RibbonCustomization.Register(ribbon);RibbonCustomization.Refresh();Check(ribbon.Count==1,"saving refreshes registered ribbon");RibbonCustomization.Clear();
            Check(DialogTree.All(DialogTree.Build()).Any(n=>n.DialogId=="agent.notes"),"notes stays in dialog tree after dynamic menu conversion");
        }finally {RibbonCustomization.Store.Data=saved;RibbonCustomization.DataForProbe=null;RibbonCustomization.Clear();}
    }
    static IEnumerable<TreeNode> Nodes(TreeNodeCollection nodes){foreach(TreeNode node in nodes){yield return node;foreach(var child in Nodes(node.Nodes))yield return child;}}
    static void DragControl(PreviewSelectionOverlay surface,Size desired,bool width,bool height,bool cancel=false) {
        var handle=surface.ResizeHandle(width,height);var start=new Point(handle.Left+handle.Width/2,handle.Top+handle.Height/2);float scale;using(var g=surface.Selected.CreateGraphics())scale=g.DpiX/96f;
        var finish=new Point(start.X+(width?(int)Math.Round(desired.Width*scale)-surface.Selected.Width:0),start.Y+(height?(int)Math.Round(desired.Height*scale)-surface.Selected.Height:0));
        var flags=BindingFlags.Instance|BindingFlags.NonPublic;typeof(PreviewSelectionOverlay).GetMethod("OnMouseDown",flags).Invoke(surface,new object[]{new MouseEventArgs(MouseButtons.Left,1,start.X,start.Y,0)});typeof(PreviewSelectionOverlay).GetMethod("OnMouseMove",flags).Invoke(surface,new object[]{new MouseEventArgs(MouseButtons.Left,0,finish.X,finish.Y,0)});if(cancel)surface.CancelDrag();else typeof(PreviewSelectionOverlay).GetMethod("OnMouseUp",flags).Invoke(surface,new object[]{new MouseEventArgs(MouseButtons.Left,1,finish.X,finish.Y,0)});
    }
    static void ControlDimensions() {
        Console.WriteLine("Control dimensions start");
        var old=LocalCustomization.Decode(Encoding.UTF8.GetBytes("{\"Version\":1,\"Names\":[],\"Orders\":[],\"Formats\":[{\"Key\":\"old\",\"Value\":{\"Text\":\"legacy\",\"Size\":11}}]}"));
        Check(old.Formats["old"].Width96==0&&old.Formats["old"].Height96==0,"old format JSON defaults to automatic geometry");
        using(var form=new Form {ClientSize=new Size(640,420)})using(var parking=new Panel()) {
            var root=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=2};root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            var flow=new FlowLayoutPanel {AutoSize=true,Dock=DockStyle.Fill,WrapContents=true};var box=new TextBox {Name="probeText",Text="用户输入0001"};var choice=new ComboBox {Name="probeChoice",DropDownStyle=ComboBoxStyle.DropDownList};choice.Items.Add("全部字符");choice.SelectedIndex=0;var button=new Button {Name="probeButton",Text="按钮",AutoSize=true};flow.Controls.AddRange(new Control[]{box,choice,button});root.Controls.Add(flow);var body=new TextBox {Name="probeBody",Multiline=true,Dock=DockStyle.Fill};root.Controls.Add(body);form.Controls.Add(root);form.TopLevel=false;parking.Controls.Add(form);form.Show();
            var targets=DialogFormatting.Register(form,"dimensions");var data=new CustomizationData();foreach(var t in targets.Where(t=>t.ResizeWidth))data.Formats["dimensions::"+t.Id]=new FormatRule {Width96=t.Control==box?210:t.Control==choice?180:140,Height96=t.Control==choice?80:60,Size=t.OriginalFont.Size,Family=t.OriginalFont.Name};
            int native=choice.Height;DialogFormatting.Apply(form,"dimensions",data);float scale=DialogSizing.Scale(form);
            Check(box.Width==(int)Math.Round(210*scale)&&box.Height==(int)Math.Round(60*scale)&&!box.Multiline,"single-line textbox accepts both dimensions without changing input semantics");
            Check(choice.Width==(int)Math.Round(180*scale)&&choice.Height==native,"dropdown supports width and rejects custom height");
            Check(button.Height==(int)Math.Round(60*scale)&&!button.Bounds.IntersectsWith(choice.Bounds),"flow neighbours reflow around explicit button height");
            Check(body.Width==(int)Math.Round(140*scale)&&body.Height==(int)Math.Round(60*scale),"fill-docked table editor accepts explicit dimensions");
            Check(box.Text=="用户输入0001","resizing preserves business input");
            foreach(var rule in data.Formats.Values)rule.Size=14;DialogFormatting.Apply(form,"dimensions",data);Check(box.Width==(int)Math.Round(210*scale)&&Math.Abs(box.Font.Size-14)<.01,"font changes retain width and height");
            var copy=data.Copy();Check(copy.Formats.Values.All(r=>r.Width96>0),"optional dimensions survive serialization");
            data.Formats.Clear();DialogFormatting.Apply(form,"dimensions",data);Check(box.Height<60*scale&&body.Dock==DockStyle.Fill&&button.AutoSize,"restoring defaults also restores native layout properties");
        }
        foreach(var definition in DialogCatalog.All)using(var form=DialogPreviewFactory.CreateFormatRender(definition.Id,definition.Default))using(var parking=new Panel()) {
            Console.WriteLine("Control dimensions: "+definition.Id);
            form.TopLevel=false;parking.Controls.Add(form);form.Show();var targets=DialogFormatting.Register(form,definition.Id);Check(targets.Where(t=>t.Id=="window"||t.Id=="self.adjust"||t.Id.Contains(".header.")).All(t=>!t.ResizeWidth&&!t.ResizeHeight),"synthetic targets do not resize shared controls "+definition.Id);
            var target=targets.FirstOrDefault(t=>t.ResizeWidth&&(t.Control is TextBoxBase||t.Control is ComboBox));if(target==null)continue;
            var data=new CustomizationData();data.Formats[definition.Id+"::"+target.Id]=new FormatRule {Width96=240,Height96=target.ResizeHeight?52:0,Size=target.OriginalFont.Size,Family=target.OriginalFont.Name};DialogFormatting.Apply(form,definition.Id,data);form.PerformLayout();
            Check(target.Control.Width==(int)Math.Round(240*DialogSizing.Scale(form)),"registered window keeps explicit width after layout "+definition.Id);
            using(var surface=new PreviewSelectionOverlay(form)){surface.RefreshSnapshot();var bounds=surface.BoundsInPreview(target.Control);Check(bounds.Width==target.Control.Width,"snapshot selection matches resized geometry "+definition.Id);}
            data.Formats.Clear();DialogFormatting.Apply(form,definition.Id,data);
        }
        using(var form=new Form())using(var flow=new FlowLayoutPanel {Dock=DockStyle.Fill})using(var parking=new Panel()) {
            var button=Ui.Button("默认按钮",112);flow.Controls.Add(button);form.Controls.Add(flow);form.TopLevel=false;parking.Controls.Add(form);form.Show();var target=DialogFormatting.Register(form,"reset").Single(t=>t.Control==button);var data=new CustomizationData();var rule=new FormatRule {Width96=220,Height96=70,Size=target.OriginalFont.Size,Family=target.OriginalFont.Name,Style=(target.OriginalFont.Bold?1:0)|(target.OriginalFont.Italic?2:0)};data.Formats["reset::"+target.Id]=rule;DialogFormatting.Apply(form,"reset",data);rule.Width96=rule.Height96=0;DialogFormatting.Apply(form,"reset",data);Check(button.Width==112&&button.Height==38&&!button.AutoSize,"resetting only dimensions restores the real default button size");
        }
        var disk=RibbonCustomization.Store.Path;byte[] saved=WpsShared.SourcePaths.Exists(disk)?File.ReadAllBytes(disk):null;
        Console.WriteLine("Control dimensions editor drag");
        using(var editor=new DialogFormatEditor(false))using(var parking=new Panel()) {
            editor.TopLevel=false;editor.FormBorderStyle=FormBorderStyle.None;parking.Controls.Add(editor);editor.Show();
            var selector=DialogFormatting.Walk(editor).OfType<ComboBox>().Single(c=>c.Name=="dialogChoice");var tree=DialogFormatting.Walk(editor).OfType<TreeView>().Single();
            selector.SelectedIndex=Array.FindIndex(DialogCatalog.All,d=>d.Id=="agent.notes");var node=Nodes(tree.Nodes).First(n=>{var t=n.Tag as FormatTarget;return t!=null&&t.Control is TextBoxBase&&t.Id!="window";});tree.SelectedNode=node;
            var resize=DialogFormatting.Walk(editor).OfType<Button>().Single(c=>c.Name=="resizeControl");resize.PerformClick();var surface=DialogFormatting.Walk(editor).OfType<PreviewSelectionOverlay>().Single();Check(!DialogFormatting.Walk(editor).Any(c=>c.Name=="controlWidth"||c.Name=="controlHeight"),"resize has no numeric entry fields");var originalDraft=LocalCustomization.Encode((CustomizationData)typeof(DialogFormatEditor).GetField("draft",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(editor));DragControl(surface,new Size(230,130),true,true,true);Check(originalDraft.SequenceEqual(LocalCustomization.Encode((CustomizationData)typeof(DialogFormatEditor).GetField("draft",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(editor))),"cancel drag restores absent rule and entire prior draft");DragControl(surface,new Size(260,150),true,true);
            var draft=(CustomizationData)typeof(DialogFormatEditor).GetField("draft",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(editor);var target=(FormatTarget)node.Tag;var rule=draft.Formats["agent.notes::"+target.Id];Check(rule.Width96==260&&rule.Height96==150,"mouse drag edits isolated draft");
            var fontSize=DialogFormatting.Walk(editor).OfType<NumericUpDown>().Single(c=>c.Name=="fontSize");fontSize.Value=12;Check(rule.Width96==260&&rule.Height96==150,"style editor cannot discard geometry");DragControl(surface,new Size(280,150),true,false);Check(rule.Size==12,"geometry editor cannot discard style");
            var actions=DialogFormatting.Walk(editor).OfType<FlowLayoutPanel>().Single(c=>c.Name=="styleActions");Check(actions.Controls.Count==3&&actions.Controls.Cast<Control>().All(c=>c.Top<actions.Height),"bold italic colour share one row");
            DialogFormatting.Walk(editor).OfType<Button>().Single(c=>c.Text=="恢复默认尺寸").PerformClick();Check(rule.Width96==0&&rule.Height96==0&&rule.Size==12,"dimension reset preserves style");
            tree.SelectedNode=tree.Nodes[0];Check(!resize.Enabled&&!surface.CanResizeWidth&&!surface.CanResizeHeight,"window title cannot accidentally resize the entire form");
            editor.Dispose();
        }
        Check(saved==null?!WpsShared.SourcePaths.Exists(disk):File.ReadAllBytes(disk).SequenceEqual(saved),"cancelled dimension editor writes no user configuration");
#if !WPS_WORD
        using(var formula=(FormulaOptionsForm)DialogPreviewFactory.CreateFormatRender("formula.add",DialogCatalog.Find("formula.add").Default))using(var parking=new Panel()) {
            formula.TopLevel=false;parking.Controls.Add(formula);formula.Show();var targets=DialogFormatting.Register(formula,"formula.add");var t=targets.Single(x=>x.Control==formula.ValueInput);var data=new CustomizationData();data.Formats["formula.add::"+t.Id]=new FormatRule {Width96=460,Height96=68,Size=t.OriginalFont.Size,Family=t.OriginalFont.Name};DialogFormatting.Apply(formula,"formula.add",data);formula.PerformLayout();Check(formula.ValueInput.Width==460&&formula.ValueInput.Height==68,"fixed-coordinate formula layout respects explicit dimensions");Check(formula.Hint.Top>=formula.ValueInput.Bottom&&!formula.ValueInput.Multiline,"formula content below resized input remains separate");formula.ClientSize=new Size(540,380);formula.PerformLayout();Check(formula.ValueInput.Width==460&&formula.AutoScrollMinSize.Width>=540,"narrow formula window retains dimensions with scrollable content");
        }
        using(var names=(FileNamesForm)DialogPreviewFactory.CreateFormatRender("file.names",DialogCatalog.Find("file.names").Default))using(var parking=new Panel()) {
            names.TopLevel=false;parking.Controls.Add(names);names.Show();var choice=DialogFormatting.Walk(names).OfType<ComboBox>().Single(c=>c.Name=="nameExtraction");int textWidth=TextRenderer.MeasureText("全部字符",choice.Font,Size.Empty,TextFormatFlags.NoPadding).Width;Check(choice.Width>=textWidth+Ui.Px(choice,30),"default character choice fits text and arrow");var t=DialogFormatting.Register(names,"file.names").Single(x=>x.Control==choice);var data=new CustomizationData();data.Formats["file.names::"+t.Id]=new FormatRule {Width96=230,Size=14,Style=1};DialogFormatting.Apply(names,"file.names",data);Check(choice.Width==230,"explicit width wins over font-driven choice measurement");
        }
#endif
        Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory,"build/link-restart-drag"));
        foreach(float factor in new[]{1f,1.25f,1.5f,2f})using(var editor=new DialogFormatEditor(false))using(var parking=new Panel())using(var font=Ui.MakeFont(10*factor,FontStyle.Regular,editor.Font)) {
            Console.WriteLine("Control dimensions font scale: "+factor);
            editor.Font=font;editor.TopLevel=false;editor.FormBorderStyle=FormBorderStyle.None;editor.ClientSize=new Size((int)(1000*factor),(int)(760*factor));parking.Controls.Add(editor);editor.Show();Console.WriteLine("Scale editor shown");var selector=DialogFormatting.Walk(editor).OfType<ComboBox>().Single(c=>c.Name=="dialogChoice");selector.SelectedIndex=Array.FindIndex(DialogCatalog.All,d=>d.Id=="agent.notes");var tree=DialogFormatting.Walk(editor).OfType<TreeView>().Single();tree.SelectedNode=Nodes(tree.Nodes).First(n=>{var t=n.Tag as FormatTarget;return t!=null&&t.Control is TextBoxBase;});DialogFormatting.Walk(editor).OfType<Button>().Single(c=>c.Name=="resizeControl").PerformClick();Console.WriteLine("Scale resize panel shown");var row=DialogFormatting.Walk(editor).OfType<FlowLayoutPanel>().Single(c=>c.Name=="styleActions");Check(row.Controls.Cast<Control>().All(c=>c.Right<=row.Width),"style row remains uncut at font scale "+factor);var viewport=DialogFormatting.Walk(editor).OfType<Panel>().Single(c=>c.Name=="propertyViewport");Check(viewport.AutoScrollMinSize.Height>=viewport.Controls[0].Height,"property scroll range includes the full dimension panel "+factor);using(var bitmap=new Bitmap(editor.Width,editor.Height)){editor.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(Environment.CurrentDirectory,"build/link-restart-drag/editor-"+((int)(factor*100))+".png"));}editor.Dispose();
        }
    }
    static void PreviewButtonText(PreviewSelectionOverlay surface, Form actual, RoundedButton button, int index) {
        // Visible does not imply exposed: a scrollable parameter panel can
        // clip this button. Test its real pixels after scrolling it into view.
        for(Control parent=button.Parent;parent!=null&&parent!=actual;parent=parent.Parent) {
            var scroll=parent as ScrollableControl;
            if(scroll!=null&&scroll.AutoScroll)scroll.ScrollControlIntoView(button);
        }
        actual.PerformLayout();surface.RefreshSnapshot();
        var bounds=surface.BoundsInPreview(button);bounds.Inflate(-4,-4);
        var exposed=new Rectangle(Point.Empty,actual.ClientSize);
        for(Control parent=button.Parent;parent!=null&&parent!=actual;parent=parent.Parent) {
            var client=surface.BoundsInPreview(parent);client.Size=parent.ClientSize;
            exposed=Rectangle.Intersect(exposed,client);
        }
        Check(bounds.Width>0&&bounds.Height>0&&exposed.Contains(bounds),"preview button scrolls fully into view "+index+" "+button.Text);
        using(var bitmap=new Bitmap(surface.Width,surface.Height)) {
            surface.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));var pixels=new PixelSampler(bitmap);int ink=0;
            for(int y=Math.Max(0,bounds.Top);y<Math.Min(bitmap.Height,bounds.Bottom);y+=2)
                for(int x=Math.Max(0,bounds.Left);x<Math.Min(bitmap.Width,bounds.Right);x+=2) {
                    var pixel=pixels.Get(x,y);if(pixel.R<180&&pixel.G<180)ink++;
                }
            Check(ink>3,"actual preview button has text "+index+" "+button.Text);
        }
    }
    static void Formatting() {
        DialogSizeStore.RootOverride=sandbox;DialogSizing.Initialize();
        using(var form=new Form {ClientSize=new Size(320,280)})using(var parking=new Panel()) {
            var background=new Panel {Bounds=new Rectangle(0,90,300,180),BackColor=Color.LightGreen};
            var viewport=new Panel {Bounds=new Rectangle(0,0,300,80),BackColor=Color.White};
            viewport.Controls.Add(new TextBox {Bounds=new Rectangle(10,120,80,24),Text="clipped text"});
            viewport.Controls.Add(new RichTextBox {Bounds=new Rectangle(110,120,80,40),Text="clipped rich text"});
            viewport.Controls.Add(new ComboBox {Bounds=new Rectangle(210,120,80,24),Text="clipped choice"});
            form.Controls.Add(background);form.Controls.Add(viewport);form.TopLevel=false;form.FormBorderStyle=FormBorderStyle.None;parking.Controls.Add(form);form.Show();
            using(var surface=new PreviewSelectionOverlay(form)) {surface.RefreshSnapshot();using(var image=new Bitmap(surface.Width,surface.Height)) {
                surface.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));
                foreach(int x in new[]{30,130,230})Check(image.GetPixel(x,130).ToArgb()==Color.LightGreen.ToArgb(),"native preview text respects parent clipping "+x);
            }}
        }
        foreach(var id in new[]{"agent.sort","agent.name","agent.format","agent.notes","agent.dialogsizes"}) {
            Console.WriteLine("Formatting preview: "+id);
            using(var form=DialogPreviewFactory.Create(id,DialogCatalog.Find(id).Default)) {
                var targets=DialogFormatting.Register(form,id);Check(targets.Count>1,"registered format targets "+id);
                var data=new CustomizationData();data.Formats[id+"::window"]=new FormatRule {Text="固定标题示例",Family="Microsoft YaHei UI",Size=12,Style=1,Color=Color.DarkGreen.ToArgb(),HasColor=true};
                var box=targets.FirstOrDefault(t=>t.Control is TextBoxBase);string before=box==null?null:box.Control.Text;
                if(box!=null)data.Formats[id+"::"+box.Id]=new FormatRule {Text="不得替换业务数据",Family="Missing-Font-For-Probe",Size=11,Style=0};
                DialogFormatting.Apply(form,id,data);Check(form.Text=="固定标题示例","window caption applied");if(box!=null){Check(!box.Caption,"text box caption is not editable");Check(box.Control.Text==before,"input contents preserved");}
                if(id=="agent.dialogsizes"){data.Formats[id+"::self.adjust"]=new FormatRule {Text="调整测试",Family="Microsoft YaHei UI",Size=12,Style=1};DialogFormatting.Apply(form,id,data);Check(DialogFormatting.AdjustmentCaption(form)=="调整测试","owner-drawn text follows format rule");}
                foreach(float factor in new[]{1f,1.25f,1.5f,2f}) {
                    form.SuspendLayout();form.ClientSize=new Size((int)(DialogCatalog.Find(id).Default.Width*factor),(int)(DialogCatalog.Find(id).Default.Height*factor));form.ResumeLayout(true);form.PerformLayout();
                    using(var bitmap=new Bitmap(form.Width,form.Height))form.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));Check(form.ClientSize.Width>0,"hidden render "+id+" "+factor);
                }
                foreach(var button in DialogFormatting.Walk(form).OfType<Button>())if(button.Text!="关闭"&&button.Text!="取消")Check(!button.Enabled,"preview business button disabled");
            }
        }
        Console.WriteLine("Formatting full editor");
        using(var editor=new DialogFormatEditor(false))using(var parking=new Panel {Size=editor.Size}) {
            editor.TopLevel=false;editor.FormBorderStyle=FormBorderStyle.None;parking.Controls.Add(editor);var initialFocus=GetFocus();editor.Show();Check(!IsWindowVisible(editor.Handle)&&GetFocus()==initialFocus,"complete editor hidden render neither shows nor takes focus");
            var choices=DialogFormatting.Walk(editor).OfType<ComboBox>().Single(c=>c.Name=="dialogChoice");
            var tree=DialogFormatting.Walk(editor).OfType<TreeView>().Single();
            Form previousSource=null;
            for(int i=0;i<choices.Items.Count;i++){Console.WriteLine("Formatting selection: "+i);var beforeSwitch=GetFocus();choices.SelectedIndex=i;Check(GetFocus()==beforeSwitch,"window switching preserves UI thread focus "+i);if(previousSource!=null)Check(previousSource.IsDisposed,"window switching releases previous source "+i);Check(tree.Nodes.Count==1&&tree.Nodes[0].Nodes.Count>0,"real format control tree "+i);
                var surface=DialogFormatting.Walk(editor).OfType<PreviewSelectionOverlay>().Single();var actual=(Form)typeof(DialogFormatEditor).GetField("preview",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(editor);
                previousSource=actual;
                Check(surface.Parent!=actual.Parent&&!DialogFormatting.Walk(editor).Contains(actual)&&!actual.TopLevel,"single visible surface; source in isolated render host "+i);
                Check(DialogPreviewFactory.IsFormatRendering(actual)&&!IsWindowVisible(actual.Handle),"format source never becomes a visible native window "+i);
                var sourceParent=actual.Parent;var focus=GetFocus();surface.RefreshSnapshot();surface.RefreshSnapshot();Check(actual.Parent==sourceParent&&GetFocus()==focus,"refresh preserves source parent and UI thread focus "+i);
                using(var bitmap=new Bitmap(surface.Width,surface.Height)){surface.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));var pixels=new PixelSampler(bitmap);int ink=0;for(int y=0;y<bitmap.Height;y+=7)for(int x=0;x<bitmap.Width;x+=7){var pixel=pixels.Get(x,y);if(pixel.R<180&&pixel.G<180)ink++;}Check(ink>20,"selection surface contains real dialog text "+i);
                    foreach(var rich in DialogFormatting.Walk(actual).OfType<RichTextBox>().Where(r=>r.TextLength>0)){var bounds=surface.BoundsInPreview(rich);bounds.Inflate(-4,-4);int bodyInk=0;for(int y=Math.Max(0,bounds.Top);y<Math.Min(bitmap.Height,bounds.Bottom);y+=5)for(int x=Math.Max(0,bounds.Left);x<Math.Min(bitmap.Width,bounds.Right);x+=5){var pixel=pixels.Get(x,y);if(pixel.R<180&&pixel.G<180)bodyInk++;}Check(bodyInk>20,"actual rich text body rendered "+i);Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory,"build/issue-fixes"));bitmap.Save(Path.Combine(Environment.CurrentDirectory,"build/issue-fixes/rich-preview-"+i+".png"));}
                    foreach(var button in DialogFormatting.Walk(actual).OfType<RoundedButton>().Where(b=>b.Text.Length>0&&b.Visible))PreviewButtonText(surface,actual,button,i);
                }
            }
            Console.WriteLine("Formatting long text");
            var caption=DialogFormatting.Walk(editor).OfType<TextBox>().Single(c=>c.Name=="captionText");var fontSize=DialogFormatting.Walk(editor).OfType<NumericUpDown>().Single(c=>c.Name=="fontSize");
            choices.SelectedIndex=Array.FindIndex(DialogCatalog.All,d=>d.Id=="agent.notes");caption.Text="加长窗口标题用于检查目录选择及预览的独立渲染";Check(choices.SelectedItem.ToString()=="需求记录","selector caption stays concise");var selectedSurface=DialogFormatting.Walk(editor).OfType<PreviewSelectionOverlay>().Single();Check(selectedSurface.Hit(new Point(-1,-1))==null,"outside snapshot cannot select a control");
            var split=DialogFormatting.Walk(editor).OfType<SplitContainer>().Single();int left=split.SplitterDistance;split.SplitterDistance=left+60;Check(split.Panel1.Width>left,"property column can widen for long labels");
#if !WPS_WORD
            choices.SelectedIndex=Array.FindIndex(DialogCatalog.All,d=>d.Id=="formula.add");tree.SelectedNode=tree.Nodes[0].Nodes[0];fontSize.Value=22;caption.Text="加 · 自定义计算（长标题布局验证）";editor.PerformLayout();
            using(var bitmap=new Bitmap(editor.Width,editor.Height)){editor.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(Environment.CurrentDirectory,"build/issue-fixes/formula-format-fixed.png"));}
#endif
            choices.SelectedIndex=Array.FindIndex(DialogCatalog.All,d=>d.Id=="agent.notes");editor.CreateControl();editor.PerformLayout();
            if(showUi){foreach(var button in DialogFormatting.Walk(editor).OfType<Button>())button.Enabled=false;editor.Show();Application.DoEvents();Thread.Sleep(250);}
            using(var bitmap=new Bitmap(editor.Width,editor.Height)){editor.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(Environment.CurrentDirectory,"build/feature099-format-preview.png"));}
        }
        Console.WriteLine("Formatting command trees");
        using(var sort=new CommandCustomizationForm("sort",true)) {var tree=DialogFormatting.Walk(sort).OfType<SortableCommandTree>().Single();Check(tree.DrawMode==TreeViewDrawMode.OwnerDrawAll&&!tree.ShowPlusMinus,"blue expanders use custom tree renderer");Check(tree.Nodes.Cast<TreeNode>().Any(n=>n.Nodes.Count>0),"sortable child handles have real menu children");Check(DialogFormatting.Walk(sort).OfType<RoundedButton>().All(b=>b.Radius==14),"sort actions have rounded corners");}
        using(var tree=new SortableCommandTree {Size=new Size(800,400)})using(var parent=new Panel()) {
            parent.Controls.Add(tree);var root=tree.Nodes.Add("一键公式");var branch=root.Nodes.Add("进行公式封装");branch.Nodes.Add("IFERROR封装");var leaf=root.Nodes.Add("移除公式封装");var other=root.Nodes.Add("数值计算");other.Nodes.Add("加：自定义计算");tree.ExpandAll();tree.CreateControl();
            foreach(float factor in new[]{1f,1.25f,1.5f,2f})using(var font=Ui.MakeFont(10*factor,FontStyle.Regular,tree.Font)) {
                tree.Font=font;Check(tree.TextArea(branch).Left==tree.TextArea(leaf).Left&&tree.TextArea(leaf).Left==tree.TextArea(other).Left,"branches and leaves align at same depth "+factor);
                int gap=Ui.Px(tree,4);Check(tree.Grip(root).IsEmpty&&tree.TextArea(root).Left-tree.Expander(root).Right==gap,"top menu has no invisible drag slot "+factor);
                Check(tree.Grip(branch).Left-tree.Expander(branch).Right==gap&&tree.TextArea(branch).Left-tree.Grip(branch).Right==gap,"visible expander, grip and caption use compact gaps "+factor);
                foreach(var node in new[]{branch,leaf,other}) {Check(!tree.Expander(node).IntersectsWith(tree.Grip(node))&&!tree.Grip(node).IntersectsWith(tree.TextArea(node)),"expand/grip/caption are disjoint "+factor);var bounds=tree.Grip(node);Check(tree.RowAt(new Point(bounds.Left,bounds.Top+1))==node,"drag handle resolves row "+factor);}
                int drags=0;ItemDragEventHandler onDrag=delegate(object sender,ItemDragEventArgs e){if(e.Item==leaf)drags++;};tree.ItemDrag+=onDrag;var handle=tree.Grip(leaf);typeof(SortableCommandTree).GetMethod("OnMouseDown",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(tree,new object[]{new MouseEventArgs(MouseButtons.Left,1,handle.Left+1,handle.Top+1,0)});typeof(SortableCommandTree).GetMethod("OnMouseMove",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(tree,new object[]{new MouseEventArgs(MouseButtons.Left,0,handle.Left+SystemInformation.DragSize.Width+10,handle.Top+1,0)});tree.ItemDrag-=onDrag;Check(drags==1,"six-dot handle starts drag after threshold "+factor);
                var box=tree.Expander(branch);bool expanded=branch.IsExpanded;typeof(SortableCommandTree).GetMethod("OnMouseDown",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(tree,new object[]{new MouseEventArgs(MouseButtons.Left,1,box.Left+1,box.Top+1,0)});Check(branch.IsExpanded!=expanded,"blue expander toggles actual branch; row="+(tree.RowAt(new Point(box.Left+1,box.Top+1))==branch)+" before="+expanded+" after="+branch.IsExpanded);branch.Expand();
                SaveCompactTree(tree,"sort",factor);
            }
        }
        using(var tree=new SortableCommandTree(false) {Size=new Size(800,400)})using(var parent=new Panel()) {
            parent.Controls.Add(tree);var root=tree.Nodes.Add("一键公式");var branch=root.Nodes.Add("进行公式封装");branch.Nodes.Add("IFERROR封装");var leaf=root.Nodes.Add("移除公式封装（长名称显示与全文提示检查）");tree.ExpandAll();tree.CreateControl();
            foreach(float factor in new[]{1f,1.25f,1.5f,2f})using(var font=Ui.MakeFont(10*factor,FontStyle.Regular,tree.Font)) {
                tree.Font=font;int gap=Ui.Px(tree,4);
                Check(tree.Grip(root).IsEmpty&&tree.Grip(branch).IsEmpty&&tree.Grip(leaf).IsEmpty,"names tree has no drag slots "+factor);
                Check(tree.TextArea(root).Left-tree.Expander(root).Right==gap&&tree.TextArea(branch).Left-tree.Expander(branch).Right==gap,"names captions follow expander slot directly "+factor);
                Check(tree.TextArea(branch).Left==tree.TextArea(leaf).Left,"names branches and leaves align "+factor);
                var box=tree.Expander(branch);bool expanded=branch.IsExpanded;typeof(SortableCommandTree).GetMethod("OnMouseDown",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(tree,new object[]{new MouseEventArgs(MouseButtons.Left,1,box.Left+1,box.Top+1,0)});Check(branch.IsExpanded!=expanded,"compact names expander remains clickable "+factor);branch.Expand();
                tree.Width=220;Check(tree.TextArea(leaf).Width>0,"narrow names tree keeps text area "+factor);SaveCompactTree(tree,"name-narrow",factor);Check(leaf.ToolTipText==leaf.Text,"long name retains full tooltip "+factor);tree.Width=800;
            }
        }
        Console.WriteLine("Formatting size directory");
        using(var sizes=new DialogSizesForm(true)) {sizes.ExpandAll();sizes.ClientSize=new Size(1500,640);sizes.FitColumns();int width=sizes.Grid.Columns[2].Width;sizes.ClientSize=new Size(1900,700);sizes.FitColumns();Check(sizes.Grid.Columns[2].Width==width&&sizes.Grid.Columns[2].AutoSizeMode==DataGridViewAutoSizeColumnMode.None,"command width follows content rather than spare space");Check(sizes.Grid.Columns[2].DefaultCellStyle.Alignment==DataGridViewContentAlignment.MiddleLeft,"command text aligns with hierarchy glyphs");}
    }
    static void SaveCompactTree(SortableCommandTree tree,string mode,float factor) {
        string directory=Path.Combine(Environment.CurrentDirectory,"build/compact-ui");Directory.CreateDirectory(directory);
        using(var bitmap=new Bitmap(tree.Width,tree.Height)){tree.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(directory,mode+"-"+((int)(factor*100))+".png"));}
    }
#if !WPS_WORD
    static void Files() {
        string root=Path.Combine(sandbox,"files");Directory.CreateDirectory(root);Directory.CreateDirectory(Path.Combine(root,"sub"));File.WriteAllText(Path.Combine(root,"磅单001Ab.txt"),"x");File.WriteAllText(Path.Combine(root,"sub/child.png"),"x");
        Check(FileNameEngine.Scan(root,1,0,CancellationToken.None).Entries.Count==1,"depth one excludes subdirectories");Check(FileNameEngine.Scan(root,2,0,CancellationToken.None).Entries.Count==2,"depth two includes child files");
        Check(FileNameEngine.ExtractName("磅单001Ab",1)=="磅单","Chinese extraction");Check(FileNameEngine.ExtractName("磅单001Ab",2)=="001Ab","non-Chinese extraction");Check(FileNameEngine.ExtractName("磅单001Ab",3)=="001","digits preserve zeros");Check(FileNameEngine.ExtractName("磅单001Ab",4)=="Ab","letters extraction");
        string supplementary=char.ConvertFromUtf32(0x20000);Check(FileNameEngine.ExtractName(supplementary+"123",1)==supplementary,"supplementary CJK extraction");
        var token=new CancellationTokenSource();token.Cancel();Reject(()=>FileNameEngine.Scan(root,1,0,token.Token),"scan cancellation");Reject(()=>FileNameEngine.Scan(root,0,0,CancellationToken.None),"invalid depth");
        var entries=FileNameEngine.Scan(root,1,3,CancellationToken.None).Entries;Check(entries[0].Name=="001"&&entries[0].Extension=="txt","name excludes extension");entries[0].Feature="000123";var values=FileNameEngine.Values(entries,true);Check(values.GetLength(1)==4&&(string)values[1,3]=="000123","four-column text output");
        var segments=new List<OcrSegment>{new OcrSegment {Text="磅单：",Left=0,Top=0,Width=60,Height=20,Page=1},new OcrSegment {Text="00123",Left=70,Top=1,Width=90,Height=18,Page=1},new OcrSegment {Text="磅单: 00999",Left=0,Top=50,Width=200,Height=20,Page=1}};
        Check(OcrFeatures.Extract(segments,"磅 单：").SequenceEqual(new[]{"00123","00999"}),"position merge and fullwidth prefix normalization");segments.Add(new OcrSegment {Text="磅单：00123",Left=0,Top=0,Width=100,Height=20,Page=2});Check(OcrFeatures.Extract(segments,"磅单:").Count==2,"duplicate candidates removed across pages");
        Check(OcrFeatures.Extract(new[]{new OcrSegment {Text="磅单：",Height=20,Page=1},new OcrSegment {Text="009",Height=20,Page=2}},"磅单:").Count==0,"pages never merged");Reject(()=>OcrFeatures.Extract(segments," "),"empty prefix rejected");
        Check(OcrFeatures.Normalize(" 磅单：Ｏ０①O0 ")=="磅单:Ｏ０①O0","normalization preserves identifier glyphs");Check(OcrFeatures.Extract(new[]{new OcrSegment {Text="磅单：;001,",Height=20,Page=1}},"磅单:").Single()==";001,","identifier punctuation is not guessed away");
        Check(OcrFeatures.Extract(new[]{new OcrSegment {Text="磅单号：PD202512020015",Left=0,Width=400,Height=30,Page=1},new OcrSegment {Text="计量单位吨",Left=480,Width=200,Height=30,Page=1}},"磅单号：").Single()=="PD202512020015","neighboring field is excluded by horizontal gap");
        Check(OcrFeatures.Extract(new[]{new OcrSegment {Text="磅单号：PD202512020015",Height=20,Page=1}},"磅单号").Single()=="PD202512020015","optional colon excluded when prefix omits separator");
        Check(OcrFeatures.Extract(new[]{new OcrSegment {Text="磅单号：:001-23",Height=20,Page=1}},"磅单号").Single()==":001-23","only one optional colon removed; identifier punctuation retained");
        using(var names=(FileNamesForm)DialogPreviewFactory.Create("file.names",DialogCatalog.Find("file.names").Default))using(var parent=new Panel()){names.TopLevel=false;names.FormBorderStyle=FormBorderStyle.None;parent.Controls.Add(names);names.Show();var options=DialogFormatting.Walk(names).OfType<TableLayoutPanel>().Single(c=>c.Name=="fileOptions");Check(options.RowCount==6,"file names has six semantic rows");var title=DialogFormatting.Walk(names).Single(c=>c.Name=="configurationHeading");Check(title.Font.Bold&&title.ForeColor==Ui.Blue,"configuration heading bold blue");var help=DialogFormatting.Walk(names).Single(c=>c.Name=="targetSummary");Check(help.Font.Bold&&help.ForeColor==Color.Firebrick,"explanation and target bold red");Check(DialogFormatting.Walk(names).OfType<ComboBox>().Any(c=>c.Name=="scanDepth")&&DialogFormatting.Walk(names).OfType<ComboBox>().Any(c=>c.Name=="startColumn"),"numeric selections use validated editable dropdowns");
            Directory.CreateDirectory(Path.Combine(Environment.CurrentDirectory,"build/issue-fixes"));foreach(var dimensions in new[]{new Size(1200,760),new Size(900,620)}){names.ClientSize=dimensions;names.CreateControl();names.PerformLayout();options.PerformLayout();var heights=options.GetRowHeights();Check(heights.Length==6&&heights.All(h=>h>0),"six rows retain height "+dimensions.Width);var table=DialogFormatting.Walk(names).OfType<DataGridView>().Single();Check(table.Height>=180,"preview grid retains usable height "+dimensions.Width);using(var bitmap=new Bitmap(names.Width,names.Height)){names.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size));bitmap.Save(Path.Combine(Environment.CurrentDirectory,"build/issue-fixes/file-names-"+dimensions.Width+".png"));}}
        }
        var workbook=new FakeWorkbook();var sheet=new FakeSheet(workbook);var app=new FakeApplication(sheet);
        app.ActiveCell=sheet.Cells[1,8];using(var names=new FileNamesForm(new FileNameWorksheet(app)))Check(DialogFormatting.Walk(names).OfType<ComboBox>().Single(c=>c.Name=="startColumn").Text=="1","default column is one even when H column is active");
        using(var session=new FileNameWorksheet(app)) {
            Check(session.Write(entries,2,true,()=>true),"write succeeds");Check((string)sheet.Cells[2,5].Value2=="000123","identifier is stored as string");Check(sheet.Cells[2,5].NumberFormat.Equals("@"),"text format applied");Check(sheet.AutoFilterMode&&sheet.FilterCalls==1,"filter created only when absent");
            sheet.Cells[1,2].Value2="original";Check(!session.Write(entries,2,true,()=>false),"overwrite cancelled");Check((string)sheet.Cells[1,2].Value2=="original","cancel preserves original value");
            for(int c=2;c<=5;c++)sheet.Cells[2,c].EntireRow.Hidden=true;
            Check(session.Write(entries,2,true,()=>true)&&sheet.FilterCalls==1,"existing filter preserved during write");Check((string)sheet.Cells[2,5].Value2=="000123","filtered hidden row receives text via single cells");
            sheet.Cells[1,2].HasArray=true;Reject(()=>session.Write(entries,2,true,()=>true),"array formula destination rejected");sheet.Cells[1,2].HasArray=false;
            sheet.Cells[1,2].MergeCells=true;Reject(()=>session.Write(entries,2,true,()=>true),"merged destination rejected");sheet.Cells[1,2].MergeCells=false;
            sheet.ProtectContents=true;Reject(()=>session.Write(entries,2,true,()=>true),"protected sheet rejected");sheet.ProtectContents=false;
            sheet.Cells[1,2].Value2="=literal text";sheet.Cells[2,2].Formula="=1+2";sheet.FailWrite=true;Reject(()=>session.Write(entries,2,true,()=>true),"partial write fails");Check((string)sheet.Cells[1,2].Value2=="=literal text"&&!sheet.Cells[1,2].HasFormula,"literal equals text restored");Check(Convert.ToString(sheet.Cells[2,2].Formula)=="=1+2","formula restored");
            sheet.Name="renamed";Reject(()=>session.Write(entries,2,true,()=>true),"changed target identity rejected");
        }
        workbook.Worksheets.Items.Add(sheet);
        using(var session=new FileNameWorksheet(app)) {
            string created=session.EnsureResultSheet(FileNameWorksheet.ResultSheetName);
            Check(created=="OCR提取结果"&&workbook.Worksheets.Count==2,"result sheet appended instead of replacing the entry sheet");
            var result=workbook.Worksheets.Item(2);
            Check(result.Name=="OCR提取结果","result sheet carries the requested name");
            Check(session.Write(entries,1,true,()=>true),"write lands on the dedicated result sheet");
            Check((string)result.Cells[1,1].Value2=="文件提取名称"&&(string)result.Cells[1,4].Value2=="文件OCR识别","result sheet carries the latest header text");
            Check(sheet.Cells[1,1].Value2==null,"entry sheet is never overwritten");
            result.Cells[6,6].Value2="stale";bool reused=session.EnsureResultSheet(FileNameWorksheet.ResultSheetName)=="OCR提取结果"&&workbook.Worksheets.Count==2;
            Check(reused,"repeat write reuses the single result sheet");Check(result.Cells[6,6].Value2==null,"existing result sheet is cleared before rewriting");
        }
        var lockedBook=new FakeWorkbook();var lockedSheet=new FakeSheet(lockedBook);lockedBook.Worksheets.Items.Add(lockedSheet);lockedBook.Worksheets.Protected=true;
        using(var session=new FileNameWorksheet(new FakeApplication(lockedSheet)))Reject(()=>session.EnsureResultSheet(FileNameWorksheet.ResultSheetName),"protected workbook cannot receive a result sheet");
        string tif=Path.Combine(sandbox,"pages.tiff");
        using(var original=new Bitmap(240,100))using(var first=new Bitmap(original))using(var second=new Bitmap(original)) {
            first.Save(Path.Combine(sandbox,"offline.png"),ImageFormat.Png);
            var encoder=ImageCodecInfo.GetImageEncoders().Single(c=>c.FormatID==ImageFormat.Tiff.Guid);using(var parameters=new EncoderParameters(1)) {
                parameters.Param[0]=new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag,(long)EncoderValue.MultiFrame);first.Save(tif,encoder,parameters);
                parameters.Param[0].Dispose();parameters.Param[0]=new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag,(long)EncoderValue.FrameDimensionPage);first.SaveAdd(second,parameters);
                parameters.Param[0].Dispose();parameters.Param[0]=new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag,(long)EncoderValue.Flush);first.SaveAdd(parameters);
            }
        }
        Check(OcrImages.Pages(tif)==2,"multipage TIFF detected");
    }
    static async Task Cloud() {
        var mock=new MockHttp();using(var ocr=new BaiduFileOcr("fake-api","fake-secret",false,mock)) {
            string path=Path.Combine(sandbox,"image.png");File.WriteAllBytes(path,new byte[]{251,255,0,1});var result=await ocr.RecognizeAsync(path,1,CancellationToken.None);Check(result.Count==1&&result[0].Text=="磅单：001","Baidu result normalized");Check(result[0].Score==.95,"optional confidence parsed");
            await ocr.RecognizeAsync(path,2,CancellationToken.None);Check(mock.Auth==1&&mock.Ocr==2,"token cached in memory");Check(mock.Bodies.Any(b=>b.Contains("image=%2B%2F8AAQ%3D%3D")),"Base64 uses exact form encoding");
            mock.TokenInvalid=true;await ocr.RecognizeAsync(path,1,CancellationToken.None);Check(mock.Auth==2,"invalid token refreshed once");
            mock.Quota=true;bool pause=false;try {await ocr.RecognizeAsync(path,1,CancellationToken.None);}catch(OcrFailure ex){pause=ex.Pause;}Check(pause,"quota error pauses batch");
        }
        string image=Path.Combine(sandbox,"image.png");var precise=new MockHttp();using(var ocr=new BaiduFileOcr("fake","fake",true,precise)){await ocr.RecognizeAsync(image,1,CancellationToken.None);Check(precise.LastPath.EndsWith("/accurate"),"high precision endpoint selected");}
        var invalid=new MockHttp {InvalidAuth=true};using(var ocr=new BaiduFileOcr("fake","fake",false,invalid)){bool failed=false;try {await ocr.RecognizeAsync(image,1,CancellationToken.None);}catch(OcrFailure){failed=true;}Check(failed&&invalid.Ocr==0,"missing token never reaches OCR");}
        foreach(string fault in new[]{"timeout","network"}) {var transport=new MockHttp {Fault=fault};using(var ocr=new BaiduFileOcr("fake","fake",false,transport)){bool failed=false;try {await ocr.RecognizeAsync(image,1,CancellationToken.None);}catch(OcrFailure ex){failed=!ex.Pause;}Check(failed&&transport.Ocr==1,"no automatic retry after "+fault);}}
        var forever=new MockHttp {AlwaysInvalid=true};using(var ocr=new BaiduFileOcr("fake","fake",false,forever)){bool stopped=false;try {await ocr.RecognizeAsync(image,1,CancellationToken.None);}catch(OcrFailure ex){stopped=ex.Pause;}Check(stopped&&forever.Auth==2&&forever.Ocr==2,"invalid token retries at most once");}
        var cancelled=new MockHttp();using(var ocr=new BaiduFileOcr("fake","fake",false,cancelled))using(var cts=new CancellationTokenSource()){cts.Cancel();bool stopped=false;try{await ocr.RecognizeAsync(image,1,cts.Token);}catch(OperationCanceledException){stopped=true;}Check(stopped&&cancelled.Ocr==0,"cloud cancellation stops before upload");}
        Reject(()=>new BaiduFileOcr("","",false),"credentials required");
    }
    static async Task CloudExpanded() {
        string image=Path.Combine(sandbox,"image.png");
        foreach(var mode in new[]{OcrMode.General,OcrMode.Accurate,OcrMode.GeneralBasic,OcrMode.AccurateBasic}) {
            var transport=new MockHttp {Basic=!OcrModes.Positioned(mode)};
            using(var ocr=new BaiduFileOcr("fake","fake",mode,transport)) {
                var segments=await ocr.RecognizeAsync(image,1,CancellationToken.None);
                Check(transport.LastPath.EndsWith("/"+OcrModes.Method(mode)),"exact Baidu variant endpoint "+mode);
                Check(OcrFeatures.Extract(segments,"磅单").Single()=="001","basic response kept without location "+mode);
                Check(segments[0].NoLocation==!OcrModes.Positioned(mode),"explicit location capability "+mode);
            }
        }
        var isolated=new[]{new OcrSegment {Text="磅单号",NoLocation=true,Sequence=0,Page=1},new OcrSegment {Text="PD202512020015",NoLocation=true,Sequence=1,Page=1}};
        Check(OcrFeatures.Extract(isolated,"磅单号").Count==0,"unpositioned lines cannot merge");
        var expected=new[]{"PD202512020015","PD202512020009","PD202512030003"};
        Check(OcrFeatures.Extract(expected.Select((v,i)=>new OcrSegment {Text="磅单号："+v,NoLocation=true,Page=i+1}),"磅单号").SequenceEqual(expected),"three expected identifiers extracted from fixture text, not a live recognition claim");
        using(var request=AliyunFileOcr.Request("test-id","test-secret",new byte[]{251,255,0,1},"2026-10-07T00:00:00Z","fixed-nonce")) {
            Check(request.RequestUri.Host==AliyunFileOcr.Host&&request.RequestUri.Query=="?NeedRotate=true&NeedSortPage=true&Row=true","Aliyun query options and regional endpoint");
            Check((await request.Content.ReadAsByteArrayAsync()).SequenceEqual(new byte[]{251,255,0,1})&&request.Content.Headers.ContentType.MediaType=="application/octet-stream","Aliyun uses unchanged binary body");
            Check(request.Headers.GetValues("Authorization").Single()=="ACS3-HMAC-SHA256 Credential=test-id,SignedHeaders=content-type;host;x-acs-action;x-acs-content-sha256;x-acs-date;x-acs-signature-nonce;x-acs-version,Signature=4c8ecd250e27c006faecadfbb3ed9320f8e8dedbe689a9a1900257902b804125","ACS3 independent fixed signature vector");
        }
        var mock=new AliyunMock();using(var ocr=new AliyunFileOcr("fake-id","fake-secret",mock)) {
            var result=await ocr.RecognizeAsync(image,2,CancellationToken.None);Check(OcrFeatures.Extract(result,"磅单号").Single()==expected[0]&&result[0].Page==2&&result[0].LineId=="7","Aliyun double JSON and explicit row result");
            Check(mock.Calls==1,"Aliyun serial request has no auth or worker call");
        }
        var points=AliyunFileOcr.Parse("{\"prism_wordsInfo\":[{\"word\":\"磅单号：001\",\"prob\":99,\"pos\":[{\"x\":10,\"y\":20},{\"x\":110,\"y\":20},{\"x\":110,\"y\":40},{\"x\":10,\"y\":40}]}]}",1);
        Check(points[0].Left==10&&points[0].Top==20&&points[0].Width==100&&points[0].Height==20&&points[0].Score==.99,"Aliyun points and percent confidence normalized");
        foreach(string fault in new[]{"quota","permission","signature","timeout","network","invalid","disconnect"}) {
            var transport=new AliyunMock {Fault=fault};using(var ocr=new AliyunFileOcr("fake","fake",transport)){bool failed=false,paused=false;try{await ocr.RecognizeAsync(image,1,CancellationToken.None);}catch(OcrFailure error){failed=true;paused=error.Pause;}Check(failed&&transport.Calls==1,"Aliyun failure never repeated "+fault);Check(paused==(fault=="quota"||fault=="permission"||fault=="signature"),"Aliyun batch pause category "+fault);}
        }
        using(var cts=new CancellationTokenSource()){cts.Cancel();var transport=new AliyunMock();using(var ocr=new AliyunFileOcr("fake","fake",transport)){bool stopped=false;try{await ocr.RecognizeAsync(image,1,cts.Token);}catch(OperationCanceledException){stopped=true;}Check(stopped&&transport.Calls==0,"Aliyun cancellation before upload");}}
        string pdf=Path.Combine(sandbox,"document.pdf"),broken=Path.Combine(sandbox,"broken.pdf");
        File.WriteAllBytes(pdf,Encoding.ASCII.GetBytes("%PDF-1.4\n1 0 obj\ntrailer\n%%EOF\n"));File.WriteAllBytes(broken,Encoding.ASCII.GetBytes("not a real document"));
        Check(OcrImages.IsPdf(pdf)&&OcrImages.Supported(pdf)&&!OcrImages.IsImage(pdf),"PDF is a supported OCR source rather than an image");
        Check(!OcrImages.Supported(Path.Combine(sandbox,"notes.txt")),"plain files stay rows without recognition");
        Check(OcrImages.LooksLikePdf(pdf)&&!OcrImages.LooksLikePdf(broken),"PDF header checked locally before upload");
        var pdfHttp=new MockHttp {Basic=true};
        using(var ocr=new BaiduFileOcr("fake","fake",OcrMode.AccurateBasic,pdfHttp)) {
            var result=await ocr.RecognizeAsync(pdf,3,CancellationToken.None);
            Check(pdfHttp.Bodies.Any(b=>b.StartsWith("pdf_file=")&&b.Contains("&pdf_file_num=3")),"Baidu uploads the PDF itself with the requested page");
            Check(pdfHttp.LastPath.EndsWith("/accurate_basic"),"PDF keeps the selected Baidu endpoint");
            Check(result[0].Page==3&&OcrFeatures.Extract(result,"磅单").Single()=="001","PDF page result normalized");
        }
        using(var ocr=new BaiduFileOcr("fake","fake",OcrMode.AccurateBasic,new MockHttp {Basic=true})){bool failed=false,paused=false;try{await ocr.RecognizeAsync(broken,1,CancellationToken.None);}catch(OcrFailure error){failed=true;paused=error.Pause;}Check(failed&&!paused,"renamed file rejected without uploading");}
        var aliyunPdf=new AliyunMock();using(var ocr=new AliyunFileOcr("fake-id","fake-secret",aliyunPdf)){bool failed=false,paused=false;try{await ocr.RecognizeAsync(pdf,1,CancellationToken.None);}catch(OcrFailure error){failed=true;paused=error.Pause;}Check(failed&&!paused&&aliyunPdf.Calls==0,"Aliyun refuses PDF before any upload");}
        OcrCredentials.Save(false,"baidu-key","baidu-secret");OcrCredentials.Save(true,"aliyun-key","aliyun-secret");Check(OcrCredentials.Read().SequenceEqual(new[]{"baidu-key","baidu-secret"}),"legacy Baidu credential file remains compatible");Check(OcrCredentials.Read(true).SequenceEqual(new[]{"aliyun-key","aliyun-secret"}),"Aliyun credentials isolated");OcrCredentials.Clear(true);Check(OcrCredentials.Read(true)[0]==""&&OcrCredentials.Read()[0]=="baidu-key","clear affects current provider only");OcrCredentials.Clear();
        string preparation=Path.Combine(sandbox,"prepared");Directory.CreateDirectory(preparation);
        using(var tiny=new Bitmap(15,20)){string small=Path.Combine(sandbox,"tiny.png");tiny.Save(small);Reject(()=>OcrImages.Prepare(small,1,OcrMode.AliyunAdvanced,preparation),"Aliyun rejects 15-pixel side");Check(WpsShared.SourcePaths.Exists(OcrImages.Prepare(small,1,OcrMode.GeneralBasic,preparation)),"Baidu accepts 15-pixel side");}
        using(var narrow=new Bitmap(1000,20)){string small=Path.Combine(sandbox,"ratio.png");narrow.Save(small);Reject(()=>OcrImages.Prepare(small,1,OcrMode.AliyunAdvanced,preparation),"Aliyun rejects ratio exactly 50");}
        foreach(var mode in new[]{OcrMode.GeneralBasic,OcrMode.AccurateBasic,OcrMode.AliyunAdvanced})for(int page=1;page<=2;page++){string converted=OcrImages.Prepare(Path.Combine(sandbox,"pages.tiff"),page,mode,preparation);Check(OcrImages.Pages(converted)==1,"TIFF converted independently "+mode+" "+page);}
        using(var form=new FileNamesForm(null)) {
            var controls=DialogFormatting.Walk(form).ToArray();var provider=controls.OfType<ComboBox>().Single(c=>c.Name=="ocrProvider");var variants=controls.OfType<ComboBox>().Single(c=>c.Name=="cloudPrecision");
            var key=controls.OfType<TextBox>().Single(c=>c.Name=="apiKey");var secret=controls.OfType<TextBox>().Single(c=>c.Name=="secretKey");
            var grid=controls.OfType<DataGridView>().Single();Check(!grid.RowHeadersVisible,"file names preview hides current-row triangle and row header");
            var entry=new FileNameEntry {Name="000123",Extension="png",Path=Path.Combine(sandbox,"preview.png"),Candidates=new List<string>{"PD001","PD002"}};
            controls.OfType<CheckBox>().Single(c=>c.Name=="useOcr").Checked=true;
            typeof(FileNamesForm).GetField("entries",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(form,new List<FileNameEntry>{entry});typeof(FileNamesForm).GetField("scannedFeatures",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(form,true);typeof(FileNamesForm).GetMethod("ShowEntries",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,null);
            Check(grid.Rows.Count==1&&grid.Rows[0].Tag==entry&&(string)grid.Rows[0].Cells["name"].Value=="000123","hidden row header retains entry identity and leading zeros");
            Check(grid.Columns["candidates"] is DataGridViewButtonColumn&&grid.Columns["candidates"].Visible&&!grid.Rows[0].Cells["feature"].ReadOnly,"candidate button and manual feature editing remain available");
            grid.Rows[0].Cells["feature"].Value="000456";typeof(DataGridView).GetMethod("OnCellEndEdit",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(grid,new object[]{new DataGridViewCellEventArgs(grid.Columns["feature"].Index,0)});
            Check(entry.Feature=="000456"&&entry.Resolved,"manual editing still commits to bound entry after hiding row header");
            Check(provider.Items.Cast<string>().SequenceEqual(new[]{"百度云OCR","阿里云OCR"})&&variants.Items.Cast<string>().SequenceEqual(OcrModes.BaiduNames),"default provider and four Baidu choices");
            key.Text="draft-baidu";secret.Text="draft-secret";provider.SelectedIndex=1;Check(variants.Items.Count==1&&variants.Text=="全文识别高精版"&&key.Text==""&&secret.Text=="","Aliyun shows one choice and separate empty draft");
            key.Text="draft-aliyun";provider.SelectedIndex=0;Check(key.Text=="draft-baidu"&&secret.Text=="draft-secret","provider drafts restored without mixing");
            Check(provider.Width>TextRenderer.MeasureText("百度云OCR",provider.Font).Width+20&&secret.UseSystemPasswordChar,"provider label has arrow padding and secret stays masked");
            Check(controls.OfType<CheckBox>().Single(c=>c.Name=="useOcr").Font.Bold,"configuration text renders bold");
            Check(controls.OfType<TextBox>().Single(c=>c.Name=="ocrPrefix").Width==120,"prefix input keeps the halved default");
            Check(controls.OfType<WidthGrip>().Count()==9,"every editable input carries a width grip");
            var jump=controls.OfType<RoundedButton>().Single(b=>b.Name=="apiPlatformJump");
            Check(jump.Text=="跳转到API平台"&&jump.Font.Bold,"api platform entry point in the OCR row");
            Check(ApiPlatforms.Names.SequenceEqual(new[]{"百度智能云","阿里云平台","未来计划"})&&ApiPlatforms.Urls[0].StartsWith("https://cloud.baidu.com/product-s/")&&ApiPlatforms.Urls[2]=="","platform options and placeholder come from the AA-Program configurator");
            using(var chooser=new ApiPlatformDialog()) {
                var buttons=DialogFormatting.Walk(chooser).OfType<RoundedButton>().ToList();
                Check(buttons.Count==4&&buttons[0].Text=="百度智能云"&&buttons[0].Enabled&&buttons[1].Enabled,"configured platforms open in the chooser");
                Check(buttons[2].Text=="未来计划"&&!buttons[2].Enabled&&buttons[3].Text=="关闭","placeholder platform without a URL stays disabled");
            }
            using(var host=new Panel())using(var box=new TextBox {Width=120})using(var grip=new WidthGrip(box)) {
                host.Controls.Add(box);host.Controls.Add(grip);box.Height=30;
                Check(grip.Height==box.Height&&grip.Cursor==Cursors.SizeWE,"grip tracks the input height and offers the resize cursor");
            }
            bool guarded=true;try{DialogFormatting.SaveControlWidth(null,null,0);}catch{guarded=false;}
            Check(guarded,"invalid width persistence is rejected without touching the store");
        }
        var resources=typeof(Connect).Assembly.GetManifestResourceNames();Check(!resources.Any(r=>r.StartsWith("WpsShared.Ocr.")),"compiled addin contains no offline OCR payload");
        // Deterministic boundary fixture: 1.64MP rather than developer desktop photographs.
        using(var boundary=new Bitmap(8200,200)){string source=Path.Combine(sandbox,"boundary.png");using(var g=Graphics.FromImage(boundary)){g.Clear(Color.White);g.FillRectangle(Brushes.Black,20,20,160,100);}boundary.Save(source,ImageFormat.Png);foreach(var mode in new[]{OcrMode.General,OcrMode.Accurate,OcrMode.AliyunAdvanced}){string converted=OcrImages.Prepare(source,1,mode,preparation);using(var picture=Image.FromFile(converted))Check(Math.Max(picture.Width,picture.Height)<=OcrModes.Longest(mode)&&Math.Min(picture.Width,picture.Height)>=16,"generated boundary dimensions "+mode);}}
        using(var noisy=new Bitmap(128,128)){var random=new Random(99);for(int y=0;y<128;y++)for(int x=0;x<128;x++)noisy.SetPixel(x,y,Color.FromArgb(random.Next(256),random.Next(256),random.Next(256)));string source=Path.Combine(sandbox,"encoded-limit.png");noisy.Save(source,ImageFormat.Png);string converted=OcrImages.Prepare(source,1,4096,preparation,4096);Check(BaiduFileOcr.ImageBody(File.ReadAllBytes(converted)).Length<=4096,"encoded byte limit shrinks deterministic fixture");Reject(()=>OcrImages.Prepare(source,1,4096,preparation,64),"impossible encoded limit stops before minimum dimensions");}

    }

#endif
#if !WPS_WORD
    static string OutputState(FakeSheet sheet,int first,int count,int rows) {
        var parts=new List<string>();for(int r=1;r<=rows;r++){parts.Add("row"+r+":"+sheet.RowAt(r).StoredHeight+":"+sheet.RowAt(r).Hidden);for(int c=first;c<first+count;c++){
            var cell=sheet.Cells[r,c];parts.Add(r+","+c+":"+Convert.ToString(cell.Value2)+":"+cell.HasFormula+":"+cell.NumberFormat+":"+cell.WrapText+":"+cell.HorizontalAlignment+":"+cell.VerticalAlignment+":"+cell.Font.Bold+":"+cell.Font.Color+":"+cell.Interior.Color);
        }}for(int c=first;c<first+count;c++)parts.Add("width"+c+":"+sheet.ColumnAt(c).ColumnWidth);return string.Join("|",parts);
    }
    static FakeSheet FormattingFixture() {
        var sheet=new FakeSheet(new FakeWorkbook());for(int r=1;r<=3;r++){sheet.RowAt(r).RowHeight=18+r;for(int c=2;c<=3;c++){
            var cell=sheet.Cells[r,c];cell.Value2=r==3?new string('x',600):"原"+r+c;cell.WrapText=(r+c)%2==0;cell.HorizontalAlignment=7;cell.VerticalAlignment=9;cell.Font.Bold=false;cell.Font.Color=37;cell.Interior.Color=51;
        }}sheet.RowAt(3).Hidden=true;return sheet;
    }
    // rule=R-28
    static void OutputFormatChecks() {
        var entries=new List<FileNameEntry>{new FileNameEntry{Name="短",Extension="png",Path="短路径",Feature="001"},new FileNameEntry{Name=new string('N',90),Extension="pdf",Path=new string('P',600),Feature="000123"}};
        foreach(bool zero in new[]{false,true})foreach(bool nullHeight in new[]{false,true}){
            var sheet=new FakeSheet(new FakeWorkbook()){AutoFilterMode=true,FilterMode=true,HiddenHeightZero=zero,HiddenHeightNull=nullHeight};sheet.RowAt(3).Hidden=true;sheet.RowAt(3).RowHeight=23;sheet.Cells[1,1].Value2="untouched";sheet.ColumnAt(1).ColumnWidth=17;
            for(int c=2;c<=5;c++)sheet.Cells[1,c].Font.Size=32;
            using(var output=new FileNameWorksheet(new FakeApplication(sheet)))Check(output.Write(entries,2,true,()=>true),"OCR output full rectangle write hidden height zero="+zero+" null="+nullHeight);
            Check(sheet.ColumnFitRanges.Single()=="B1:E3","OCR measures final values including last hidden row and non-A start");
            Check(Convert.ToDouble(sheet.ColumnAt(2).ColumnWidth)>=entries[1].Name.Length&&Convert.ToDouble(sheet.ColumnAt(4).ColumnWidth)==255,"full output data grows beyond 30 chars and caps only at host maximum");
            Check(sheet.RowAt(3).Hidden&&sheet.AutoFilterMode&&sheet.FilterMode&&sheet.FilterCalls==0,"temporary reveal preserves hidden state and existing filter");
            Check((bool)sheet.Cells[3,4].WrapText&&!(bool)sheet.Cells[1,4].WrapText&&!(bool)sheet.Cells[2,2].WrapText,"maximum-width column wraps body only");
            Check(sheet.RowAt(3).StoredHeight>=45&&sheet.RowAt(1).StoredHeight>=44,"wrapped body and large-font header heights fit complete text");
            Check(sheet.Cells[1,1].Value2.Equals("untouched")&&Convert.ToDouble(sheet.ColumnAt(1).ColumnWidth)==17,"non-output column content and width preserved");
            for(int c=2;c<=5;c++)Check((bool)sheet.Cells[1,c].Font.Bold&&Convert.ToInt32(sheet.Cells[1,c].HorizontalAlignment)==-4108&&Convert.ToInt32(sheet.Cells[1,c].VerticalAlignment)==-4108,"OCR header centered bold col="+c);
        }
        foreach(bool ocr in new[]{false,true}){
            var sheet=new FakeSheet(new FakeWorkbook());using(var output=new FileNameWorksheet(new FakeApplication(sheet)))Check(output.Write(new List<FileNameEntry>(),3,ocr,()=>true),"empty output header only "+ocr);
            Check(sheet.ColumnFitRanges.Single()==(ocr?"C1:F1":"C1:E1"),"three/four column output full range when empty "+ocr);
            Check(sheet.RowFitRanges.Single()==(ocr?"C1:F1":"C1:E1")&&sheet.RowAt(1).StoredHeight>=26,"empty output header fit uses minimum row height "+ocr);
        }
        foreach(string fault in new[]{"column","body-row","header-row"}){
            var sheet=FormattingFixture();string before=OutputState(sheet,2,2,3);sheet.FailAutoFit=fault=="column";sheet.FailRowFit=fault=="body-row";sheet.FailHeaderFit=fault=="header-row";
            Reject(()=>WorkbookUi.FitOutput(sheet,2,2,3),"partial AutoFit failure reports "+fault);
            Check(OutputState(sheet,2,2,3)==before,"helper restores mixed wraps widths row heights header and hidden rows "+fault);
        }
        var restore=FormattingFixture();restore.FailAutoFit=true;restore.FailRestoreWidthColumn=2;
        string failure="";try{WorkbookUi.FitOutput(restore,2,2,3);}catch(Exception ex){failure=ex.Message;}
        Check(failure.Contains("列宽2")&&failure.Contains("injected width restore failure"),"helper identifies unrestored width");
        Check(!(bool)restore.Cells[1,2].Font.Bold&&Convert.ToInt32(restore.Cells[1,2].HorizontalAlignment)==7&&restore.RowAt(3).Hidden&&restore.RowAt(3).StoredHeight==21,"width restore failure does not skip same-column header or row hidden recovery");
        foreach(string fault in new[]{"column","body-row","header-row","restore-width"}){
            var sheet=FormattingFixture();sheet.AutoFilterMode=false;sheet.Cells[2,2].Formula="=2+3";sheet.Cells[2,3].Value2="=literal";string before=OutputState(sheet,2,4,3);
            sheet.FailAutoFit=fault=="column"||fault=="restore-width";sheet.FailRowFit=fault=="body-row";sheet.FailHeaderFit=fault=="header-row";sheet.FailRestoreWidthColumn=fault=="restore-width"?2:0;
            using(var output=new FileNameWorksheet(new FakeApplication(sheet)))Reject(()=>output.Write(entries,2,true,()=>true),"OCR write fit failure "+fault);
            if(fault!="restore-width")Check(OutputState(sheet,2,4,3)==before,"OCR outer rollback restores original values formulas formats and dimensions "+fault);
            else Check(Convert.ToString(sheet.Cells[2,2].Formula)=="=2+3"&&sheet.Cells[2,3].Value2.Equals("=literal")&&sheet.RowAt(3).Hidden&&sheet.RowAt(3).StoredHeight==21&&!(bool)sheet.Cells[1,2].Font.Bold,"OCR width failure still restores contents styles heights hidden rows");
            Check(!sheet.AutoFilterMode,"OCR failed write removes only newly created filter "+fault);
        }
        var invalid=FormattingFixture();invalid.HiddenHeightInvalid=true;string invalidBefore=OutputState(invalid,2,2,3);Reject(()=>WorkbookUi.FitOutput(invalid,2,2,3),"invalid latent hidden height refuses format mutation");Check(OutputState(invalid,2,2,3)==invalidBefore&&invalid.ColumnFitRanges.Count==0,"snapshot failure restores hidden state and leaves format untouched");
        var visibility=FormattingFixture();visibility.HiddenHeightZero=true;visibility.FailHideAfterFit=true;string visibilityBefore=OutputState(visibility,2,2,3),visibilityError="";
        try {WorkbookUi.FitOutput(visibility,2,2,3);}catch(Exception ex){visibilityError=ex.Message;}
        Check(visibilityError.Contains("第3行")&&visibilityError.Contains("injected hidden restore failure"),"successful fit with failed finally hiding reports specific original row");
        Check(OutputState(visibility,2,2,3)==visibilityBefore,"finally hide failure rolls back widths mixed wraps headers heights and all hidden states");
        Check(visibility.ColumnFitRanges.Count==1&&visibility.RowFitRanges.Count==2,"hide failure injection occurs only after successful column body and header fit");
        Console.WriteLine("Generated worksheet state and rollback PASS: "+checks);
    }
    static void ImageSamples(string directory){
        if(!Directory.Exists(directory))throw new DirectoryNotFoundException("指定图片目录不存在");
        string output=Path.Combine(sandbox,"samples");Directory.CreateDirectory(output);var originals=WpsShared.SourcePaths.GetFiles(directory,"*.png");if(originals.Length==0)throw new ArgumentException("指定目录没有 PNG 图片");
        foreach(var original in originals)foreach(var mode in new[]{OcrMode.General,OcrMode.Accurate,OcrMode.AliyunAdvanced}){string converted=OcrImages.Prepare(original,1,mode,output);using(var picture=Image.FromFile(converted))Check(Math.Max(picture.Width,picture.Height)<=OcrModes.Longest(mode),"explicit image preprocessing "+mode);}
        Console.WriteLine("ImageSamples: "+originals.Length+" PNG, "+checks+" PASS; local preprocessing only");
    }
#endif
    [STAThread]static int Main(string[] args) {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        if(args.Length==2&&args[0]=="--format-read"){FormatSession.SourceOverride=args[1];Console.WriteLine(FormatSession.Data.Formats["probe::window"].Text+"|"+FormatSession.Data.Formats["probe::window"].Width96);return 0;}
        if(args.Contains("--sleep")){Thread.Sleep(30000);return 0;}
        sandbox=Path.Combine(Path.GetTempPath(),"Feature099Probe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(sandbox);
#if !WPS_WORD
        OcrCredentials.RootOverride=sandbox;OcrDiagnostics.RootOverride=sandbox;OcrPrefixStore.RootOverride=sandbox;
#endif
        UiDiagnostics.RootOverride=sandbox;
        try {
#if !WPS_WORD
            if(args.Length==2&&args[0]=="--ocr-samples"){ImageSamples(args[1]);return 0;}
            if(args.Length==2&&args[0]=="--live-cloud"){LiveCloudWorkflow(args[1]);return 0;}
            if(args.Contains("--output-format")){OutputFormatChecks();return 0;}
#endif
            PixelSampler.Verify();if(args.Contains("--grid-headers")){GridHeaderChecks();return 0;}GridHeaderChecks();showUi=args.Contains("--ui");if(showUi){Formatting();Console.WriteLine("Foreground format preview PASS");return 0;}SessionAndThread();Console.WriteLine("Process format session and editor STA PASS");Settings();Notes();Console.WriteLine("Settings and notes PASS");Ribbon();Console.WriteLine("Ribbon PASS");Formatting();ControlDimensions();Console.WriteLine("Formatting and control dimensions PASS");LiveRefresh();Console.WriteLine("Live configuration PASS");
#if !WPS_WORD
            Files();OutputFormatChecks();MultiplePrefixes();FileNamesLayoutChecks();Workflow();Console.WriteLine("Files, multiple prefixes, result layout and asynchronous UI workflow PASS");if(args.Contains("--layout")){Console.WriteLine("Layout fixes PASS: "+checks);return 0;}Task.Run(()=>Cloud()).GetAwaiter().GetResult();Console.WriteLine("Cloud mock PASS");Task.Run(()=>CloudExpanded()).GetAwaiter().GetResult();Console.WriteLine("Cloud variants and Aliyun PASS");
#endif
            Console.WriteLine("Feature099Probe: "+checks+" PASS, 0 FAIL");return 0;
        }catch(Exception ex){Console.WriteLine("FAILED: "+ex);return 1;}
        finally {LiveCustomization.Stop();if(sandbox.StartsWith(Path.GetTempPath(),StringComparison.OrdinalIgnoreCase)&&Directory.Exists(sandbox))Directory.Delete(sandbox,true);}
    }
}
public sealed class FormatDomainReader:MarshalByRefObject {
    public string Read(string path){FormatSession.SourceOverride=path;return FormatSession.Data.Formats["probe::window"].Text+"|"+FormatSession.Data.Formats["probe::window"].Width96;}
    public override object InitializeLifetimeService(){return null;}
}
public sealed class FakeRibbon {public int Count;public void Invalidate(){Count++;}}
public sealed class RestartMockApp {public long Hwnd;public object ActiveWorkbook;public object ActiveDocument{get{return ActiveWorkbook;}}public RestartMockBooks Workbooks;public RestartMockBooks Documents{get{return Workbooks;}}public RestartMockDialogs Dialogs;}
public sealed class RestartMockBooks {public RestartMockBook[] Items;public int Count{get{return Items.Length;}}public object Item(int index){return Items[index-1];}}
public sealed class RestartMockBook {public string FullName;public string Name{get{return System.IO.Path.GetFileName(FullName);}}public string Path{get{return System.IO.Path.GetDirectoryName(FullName)??"";}}public bool Saved,ReadOnly,CancelSave;public int SaveCount;public void Activate(){}public void Save(){SaveCount++;if(!CancelSave)Saved=true;}}
public sealed class RestartMockDialogs {public int Kind,Count;public bool Result;public Action OnShow;public object Item(int kind){Kind=kind;return this;}public object Show(){Count++;if(Result&&OnShow!=null)OnShow();return Kind==84?(object)(Result?-1:0):(object)Result;}}
#if !WPS_WORD
public sealed class AliyunMock:HttpMessageHandler {
    public int Calls;public string Fault;
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken cancel) {
        cancel.ThrowIfCancellationRequested();Calls++;
        if(Fault=="timeout")throw new TaskCanceledException();if(Fault=="network"||Fault=="disconnect")throw new HttpRequestException();
        string value="{\"RequestId\":\"test-request\",\"Data\":\"{\\\"prism_rowsInfo\\\":[{\\\"rowId\\\":7,\\\"word\\\":\\\"磅单号：PD202512020015\\\"}]}\"}";
        if(Fault=="quota"||Fault=="permission"||Fault=="signature")value="{\"RequestId\":\"test-request\",\"Code\":\""+(Fault=="quota"?"QuotaExhausted":Fault=="permission"?"noPermission":"SignatureDoesNotMatch")+"\"}";
        if(Fault=="invalid")value="invalid-json";
        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK){Content=new StringContent(value,Encoding.UTF8,"application/json")});
    }
}
public sealed class MockHttp:HttpMessageHandler {
    public int Auth,Ocr;public bool TokenInvalid,Quota,InvalidAuth,AlwaysInvalid,Basic;public string LastPath,Fault;public List<string>Bodies=new List<string>();
    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken cancel) {
        cancel.ThrowIfCancellationRequested();Bodies.Add(await request.Content.ReadAsStringAsync());LastPath=request.RequestUri.AbsolutePath;string value;
        if(LastPath.EndsWith("/token")){Auth++;value=InvalidAuth?"{\"error\":\"invalid_client\"}":"{\"access_token\":\"fake-token\",\"expires_in\":3600}";}
        else {Ocr++;if(Fault=="timeout")throw new TaskCanceledException();if(Fault=="network")throw new HttpRequestException();if(TokenInvalid||AlwaysInvalid){TokenInvalid=false;value="{\"error_code\":110}";}else if(Quota)value="{\"error_code\":17}";else value=Basic?"{\"words_result\":[{\"words\":\"磅单：001\",\"probability\":{\"average\":0.95}}]}":"{\"words_result\":[{\"words\":\"磅单：001\",\"location\":{\"left\":1,\"top\":1,\"width\":100,\"height\":20},\"probability\":{\"average\":0.95}}]}";}
        return new HttpResponseMessage(HttpStatusCode.OK){Content=new StringContent(value,Encoding.UTF8,"application/json")};
    }
}
public sealed class FakeWorkbook {public string FullName {get{return "C:\\test\\fixture.xlsx";}}public FakeWorksheets Worksheets;public FakeWorkbook(){Worksheets=new FakeWorksheets(this);}}
public sealed class FakeWorksheets {
    private readonly FakeWorkbook owner;public readonly List<FakeSheet> Items=new List<FakeSheet>();
    public FakeWorksheets(FakeWorkbook owner){this.owner=owner;}
    public int Count{get{return Items.Count;}}
    public FakeSheet Item(int index){return Items[index-1];}
    public FakeSheet Add(object before,object after){if(Protected)throw new System.Runtime.InteropServices.COMException("workbook structure protected");var sheet=new FakeSheet(owner){Name="Sheet"+(Items.Count+1)};Items.Add(sheet);return sheet;}
    public bool Protected;
}
public sealed class FakeApplication {public FakeSheet ActiveSheet;public FakeCell ActiveCell;public FakeApplication(FakeSheet sheet){ActiveSheet=sheet;ActiveCell=sheet.Cells[1,2];}}
public sealed class FakeCount {public int Count;public FakeCount(int count){Count=count;}}
public sealed class FakeFont {
    private readonly Func<IEnumerable<FakeCell>> cells;private object bold=false,color=0,name="Calibri",size=11;
    public FakeFont(){}public FakeFont(Func<IEnumerable<FakeCell>> cells){this.cells=cells;}
    private object Get(Func<FakeFont,object> select,object own){return cells==null?own:select(cells().First().Font);}
    private void Set(Action<FakeFont> update){if(cells!=null)foreach(var cell in cells())update(cell.Font);}
    public object Bold{get{return Get(f=>f.Bold,bold);}set{bold=value;Set(f=>f.Bold=value);}}
    public object Color{get{return Get(f=>f.Color,color);}set{color=value;Set(f=>f.Color=value);}}
    public object Name{get{return Get(f=>f.Name,name);}set{name=value;Set(f=>f.Name=value);}}
    public object Size{get{return Get(f=>f.Size,size);}set{size=value;Set(f=>f.Size=value);}}
}
public sealed class FakeInterior {
    private readonly Func<IEnumerable<FakeCell>> cells;private object color=16777215;
    public FakeInterior(){}public FakeInterior(Func<IEnumerable<FakeCell>> cells){this.cells=cells;}
    public object Color{get{return cells==null?color:cells().First().Interior.Color;}set{color=value;if(cells!=null)foreach(var cell in cells())cell.Interior.Color=value;}}
}
public sealed class FakeVisibility {
    private readonly FakeSheet sheet;private readonly int row;private bool hidden;private object height=15;
    public bool Hidden{get{return hidden;}set{if(value&&row==3&&sheet.FailHideAfterFit&&sheet.ColumnFitRanges.Count>0){sheet.FailHideAfterFit=false;throw new IOException("injected hidden restore failure");}hidden=value;}}
    public FakeVisibility(FakeSheet sheet,int row){this.sheet=sheet;this.row=row;}
    public object RowHeight{get{return row==3&&sheet.HiddenHeightInvalid?(object)0:Hidden&&sheet.HiddenHeightZero?(object)0:Hidden&&sheet.HiddenHeightNull?null:height;}set{height=value;}}
    public double StoredHeight{get{return Convert.ToDouble(height);}}
}
public sealed class FakeColumn {
    private readonly FakeSheet sheet;public readonly int Column;private object width=8.43;
    public FakeColumn(FakeSheet sheet,int column){this.sheet=sheet;Column=column;}
    public object ColumnWidth{get{return width;}set{if(Column==sheet.FailRestoreWidthColumn&&Convert.ToDouble(value)==8.43)throw new IOException("injected width restore failure");width=value;}}
}
public sealed class FakeCell {
    private readonly FakeSheet sheet;public FakeVisibility EntireRow{get{return sheet.RowAt(Row);}}public FakeColumn EntireColumn{get{return sheet.ColumnAt(Column);}}
    public object RowHeight{get{return EntireRow.RowHeight;}set{EntireRow.RowHeight=value;}}public object ColumnWidth{get{return EntireColumn.ColumnWidth;}set{EntireColumn.ColumnWidth=value;}}
    public readonly int Row,Column;public object Value2;public object NumberFormat="General",HorizontalAlignment=0,VerticalAlignment=0,WrapText=false;public FakeFont Font=new FakeFont();public FakeInterior Interior=new FakeInterior();public bool MergeCells,HasArray,HasFormula;private string formula;
    public FakeCell(FakeSheet sheet,int row,int column){this.sheet=sheet;Row=row;Column=column;}
    public object Formula {get{return HasFormula?formula:Convert.ToString(Value2);}set {formula=Convert.ToString(value);Value2=value;HasFormula=value!=null&&Convert.ToString(value).StartsWith("=");}}
}
public sealed class FakeCells {
    private readonly FakeSheet sheet;private Dictionary<string,FakeCell> cells=new Dictionary<string,FakeCell>();public FakeCells(FakeSheet sheet){this.sheet=sheet;}public void Clear(){cells.Clear();}public FakeCell this[int row,int column]{get {string key=row+","+column;FakeCell cell;if(!cells.TryGetValue(key,out cell)){cell=new FakeCell(sheet,row,column);cells[key]=cell;}return cell;}}
    public FakeVisibility Rows(int index) {return sheet.RowAt(index);}
}
public sealed class FakeRanges {
    private FakeSheet sheet;public FakeRanges(FakeSheet s){sheet=s;}
    public FakeRange this[object first,object last]{get{return new FakeRange(sheet,(FakeCell)first,(FakeCell)last);}}
    public FakeRange this[string address]{get {string[] parts=(address??"A1").Split(':');FakeCell a=Parse(sheet,parts[0]);return new FakeRange(sheet,a,parts.Length>1?Parse(sheet,parts[1]):a);}}
    private static FakeCell Parse(FakeSheet sheet,string reference) {
        int column=0,row=0,i=0;string text=reference??"A1";
        while(i<text.Length&&text[i]!='$'&&!char.IsDigit(text[i])){char c=text[i];if(c>='a'&&c<='z')c=(char)(c-'a'+'A');if(c>='A'&&c<='Z')column=column*26+(c-'A'+1);i++;}
        while(i<text.Length&&char.IsDigit(text[i])){row=row*10+(text[i]-'0');i++;}
        return sheet.Cells[row>0?row:1,column>0?column:1];
    }
}
public sealed class FakeRange {
    internal readonly FakeSheet Sheet;internal readonly FakeCell First,Last;private FakeSheet sheet{get{return Sheet;}}private FakeCell first{get{return First;}}private FakeCell last{get{return Last;}}
    public FakeRange(FakeSheet sheet,FakeCell first,FakeCell last){Sheet=sheet;First=first;Last=last;Font=new FakeFont(All);Interior=new FakeInterior(All);}
    public IEnumerable<FakeCell> All(){for(int r=first.Row;r<=last.Row;r++)for(int c=first.Column;c<=last.Column;c++)yield return sheet.Cells[r,c];}
    internal string Address{get{return ComFormulaHost.Address(first.Row,first.Column)+":"+ComFormulaHost.Address(last.Row,last.Column);}}
    public object Formula {get {var a=new object[last.Row-first.Row+1,last.Column-first.Column+1];for(int r=first.Row;r<=last.Row;r++)for(int c=first.Column;c<=last.Column;c++)a[r-first.Row,c-first.Column]=sheet.Cells[r,c].Formula;return a;}}
    public object Value2 {get {var a=new object[last.Row-first.Row+1,last.Column-first.Column+1];for(int r=first.Row;r<=last.Row;r++)for(int c=first.Column;c<=last.Column;c++)a[r-first.Row,c-first.Column]=sheet.Cells[r,c].Value2;return a;}set {var a=(object[,])value;for(int r=first.Row;r<=last.Row;r++)for(int c=first.Column;c<=last.Column;c++){if(sheet.Cells[r,c].EntireRow.Hidden)continue;sheet.Cells[r,c].Value2=a[r-first.Row,c-first.Column];sheet.Cells[r,c].HasFormula=false;if(sheet.FailWrite){sheet.FailWrite=false;throw new IOException("injected partial write");}}}}
    public object NumberFormat {set {for(int r=first.Row;r<=last.Row;r++)for(int c=first.Column;c<=last.Column;c++)sheet.Cells[r,c].NumberFormat=value;}}
    public FakeFont Font;public FakeInterior Interior;public object HorizontalAlignment{get{return first.HorizontalAlignment;}set{foreach(var cell in All())cell.HorizontalAlignment=value;}}public object VerticalAlignment{get{return first.VerticalAlignment;}set{foreach(var cell in All())cell.VerticalAlignment=value;}}
    public object WrapText{get{var kinds=All().Select(c=>c.WrapText).Distinct().ToArray();return kinds.Length==1?kinds[0]:null;}set{foreach(var cell in All())cell.WrapText=value;}}
    public object RowHeight{get{return first.RowHeight;}set{for(int r=first.Row;r<=last.Row;r++)sheet.RowAt(r).RowHeight=value;}}
    public object ColumnWidth{get{return first.ColumnWidth;}set{for(int c=first.Column;c<=last.Column;c++)sheet.ColumnAt(c).ColumnWidth=value;}}
    public FakeVisibility EntireRow{get{return first.EntireRow;}}public FakeColumn EntireColumn{get{return first.EntireColumn;}}
    public FakeDimensions Columns{get{return new FakeDimensions(this,false);}}public FakeDimensions Rows{get{return new FakeDimensions(this,true);}}
    public void AutoFilter(){sheet.FilterCalls++;sheet.AutoFilterMode=true;}
}
public sealed class FakeDimensions {
    private readonly FakeRange range;private readonly bool rows;public FakeDimensions(FakeRange range,bool rows){this.range=range;this.rows=rows;}
    public void AutoFit(){var sheet=range.Sheet;if(rows){sheet.RowFitRanges.Add(range.Address);for(int r=range.First.Row;r<=range.Last.Row;r++){
        int lines=1;double size=11;for(int c=range.First.Column;c<=range.Last.Column;c++){var cell=sheet.Cells[r,c];size=Math.Max(size,Convert.ToDouble(cell.Font.Size));if((bool)cell.WrapText)lines=Math.Max(lines,(int)Math.Ceiling(Convert.ToString(cell.Value2).Length/Math.Max(1d,Convert.ToDouble(cell.ColumnWidth)-2)));}sheet.RowAt(r).RowHeight=lines*Math.Max(15,size*1.4);
        if(range.First.Row>1&&sheet.FailRowFit||range.First.Row==1&&sheet.FailHeaderFit){sheet.FailRowFit=sheet.FailHeaderFit=false;throw new IOException("injected partial row fit failure");}
    }}else{sheet.ColumnFitRanges.Add(range.Address);for(int c=range.First.Column;c<=range.Last.Column;c++){
        double needed=1;for(int r=range.First.Row;r<=range.Last.Row;r++){var cell=sheet.Cells[r,c];if(cell.EntireRow.Hidden)continue;needed=Math.Max(needed,Convert.ToString(cell.Value2).Length+2+((bool)cell.Font.Bold?1:0));}sheet.ColumnAt(c).ColumnWidth=Math.Min(255,needed);
        if(sheet.FailAutoFit){sheet.FailAutoFit=false;throw new IOException("injected partial column fit failure");}
    }}}
}
public sealed class FakeSheet {
    private readonly Dictionary<int,FakeVisibility> rowState=new Dictionary<int,FakeVisibility>();private readonly Dictionary<int,FakeColumn> columnState=new Dictionary<int,FakeColumn>();
    public FakeVisibility RowAt(int row){FakeVisibility value;if(!rowState.TryGetValue(row,out value)){value=new FakeVisibility(this,row);rowState[row]=value;}return value;}
    public FakeColumn ColumnAt(int column){FakeColumn value;if(!columnState.TryGetValue(column,out value)){value=new FakeColumn(this,column);columnState[column]=value;}return value;}
    public FakeWorkbook Parent;public string Name="Fixture";public FakeCount Rows=new FakeCount(1048576),Columns=new FakeCount(16384);public FakeCells Cells;public FakeRanges Range;public bool ProtectContents,AutoFilterMode,FilterMode,FailWrite,FailAutoFit,FailRowFit,FailHeaderFit,HiddenHeightZero,HiddenHeightNull,HiddenHeightInvalid,FailHideAfterFit;public int FilterCalls,FailRestoreWidthColumn;public List<string> ColumnFitRanges=new List<string>(),RowFitRanges=new List<string>();
    public FakeSheet(FakeWorkbook workbook){Parent=workbook;Cells=new FakeCells(this);Range=new FakeRanges(this);}
}
#endif
