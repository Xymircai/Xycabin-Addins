using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

internal static class DialogSizesProbe
{
    const BindingFlags All=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Static|BindingFlags.Instance;
    static int checks;
    [DllImport("user32.dll")]static extern IntPtr SendMessage(IntPtr h,int msg,IntPtr w,IntPtr l);
    [StructLayout(LayoutKind.Sequential)]struct WindowRect {internal int Left,Top,Right,Bottom;}
    static void ResizePreview(Form form,Size client,bool edge)
    {
        SendMessage(form.Handle,0x0231,IntPtr.Zero,IntPtr.Zero);
        var rect=new WindowRect {Left=form.Left,Top=form.Top,Right=form.Right,Bottom=form.Bottom};IntPtr buffer=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(WindowRect)));
        try{Marshal.StructureToPtr(rect,buffer,false);SendMessage(form.Handle,edge?0x0214:0x0216,new IntPtr(8),buffer);form.ClientSize=client;}
        finally{Marshal.FreeHGlobal(buffer);SendMessage(form.Handle,0x0232,IntPtr.Zero,IntPtr.Zero);}
    }
    static void Check(bool ok,string label){if(!ok)throw new Exception(label);checks++;}
    static object Call(Type type,string method,params object[] args)
    {try{return type.GetMethod(method,All).Invoke(null,args);}catch(TargetInvocationException ex){throw ex.InnerException;}}
    static object Field(object instance,string name){return instance.GetType().GetField(name,All).GetValue(instance);}
    static object Make(Type type,params object[] args){return Activator.CreateInstance(type,All,null,args,null);}
    static IEnumerable<Control> Controls(Control root){foreach(Control c in root.Controls){yield return c;foreach(Control n in Controls(c))yield return n;}}
    static Button FindButton(Control root,string text)
    {foreach(Control c in root.Controls){Button b=c as Button;if(b!=null && b.Text==text)return b;Button nested=FindButton(c,text);if(nested!=null)return nested;}return null;}
    static DataGridViewRow WindowRow(DataGridView grid)
    {foreach(DataGridViewRow row in grid.Rows)if(row.Tag.GetType().GetProperty("Definition",All).GetValue(row.Tag,null)!=null)return row;throw new Exception("No registered window row");}
    static string RowId(DataGridViewRow row){return (string)Field(row.Tag,"DialogId");}
    static void Invoke(Form form,string method,params object[] args){form.GetType().GetMethod(method,All).Invoke(form,args);}
    static Dictionary<string,Size> Sizes(Type store){return (Dictionary<string,Size>)Field(Call(store,"Read"),"Values");}
    static bool StageSize(Form editor,string id,Size size)
    {
        bool accepted=(bool)editor.GetType().GetMethod("PersistVisualSize",All).Invoke(editor,new object[]{id,size});
        if(accepted)Invoke(editor,"RebuildRows"); // Same draft-to-grid synchronization as preview close.
        return accepted;
    }
    static Size CapturedSize(Type format,Form form,string id)
    {
        var sizes=(Dictionary<string,int[]>)Field(Call(format,"Capture",form),"WindowSizes");int[] value;
        return sizes!=null&&sizes.TryGetValue(id,out value)?new Size(value[0],value[1]):Size.Empty;
    }
    static void VerifyPasswordPreview(Assembly asm,string ns,Type catalog,Type factory)
    {
        const string id="agent.developerpassword";
        object definition=Call(catalog,"Find",id);
        Check(definition!=null&&(Size)Field(definition,"Default")==new Size(420,230)&&(Size)Field(definition,"Minimum")==new Size(380,210),"password window has registered default and minimum");
        Type tree=asm.GetType(ns+"DialogTree",true);int count=0;
        foreach(object node in (System.Collections.IEnumerable)Call(tree,"All",Call(tree,"Build")))
            if((string)Field(node,"DialogId")==id)
            {
                count++;object parent=Field(node,"Parent");
                Check(parent!=null&&(string)Field(parent,"DialogId")=="agent.developer","password belongs beneath developer preview, without a fake Ribbon command");
            }
        Check(count==1,"password window occurs once in the dialog tree");
        using(var preview=(Form)Call(factory,"Create",id,(Size)Field(definition,"Default")))
        {
            var password=(TextBox)Field(preview,"Password");var hint=(Label)Field(preview,"Hint");
            Check((bool)Call(factory,"IsPreview",preview)&&!preview.Visible,"password preview is hidden and registered");
            Check(password.UseSystemPasswordChar&&password.ReadOnly&&password.Name=="developerPassword","password preview masks and locks input");
            Check(hint.Text=="请输入开发者的微信名称"&&hint.Name=="passwordHint","password preview retains the exact prompt");
            Check(preview.AcceptButton!=null&&!((Control)preview.AcceptButton).Enabled,"password preview disables confirmation");
            int attempts=0;
            preview.GetType().GetField("Verify",All).SetValue(preview,new Func<string,bool>(delegate {attempts++;throw new Exception("Preview must not authorize");}));
            Invoke(preview,"Submit");Invoke(preview,"ProcessCmdKey",new Message(),Keys.Enter);
            Check(attempts==0&&preview.DialogResult!=DialogResult.OK&&!preview.IsDisposed,"submit and Enter in size preview never authorize or accept");
        }
    }
    // No Show, activation, preview or message loop: exercise real forms with isolated settings.
    static void VerifyCore(Assembly asm,string ns,Type store,Type sizing,Type format,Type settings)
    {
        string path=(string)store.GetProperty("PathForUser",All).GetValue(null,null);
        Type reader=asm.GetType(ns+"AgentDocumentForm",true);int openForms=Application.OpenForms.Count;
        Call(store,"Save",Call(store,"Read"),new Dictionary<string,Size>());
        using(var existing=(Form)Make(reader,false))
        {
            Size original=existing.ClientSize;
            Check(!(bool)Call(sizing,"HasOverride",existing),"existing reader starts with program default");
            var values=new Dictionary<string,Size>{{"agent.developer",new Size(1100,730)},{"future.window",new Size(850,500)}};
            Call(store,"Save",Call(store,"Read"),values);
            using(var fresh=(Form)Make(reader,false))
            {
                Check((bool)Call(sizing,"HasOverride",fresh)&&CapturedSize(format,fresh,"agent.developer")==new Size(1100,730),"new reader applies saved size in the same process for both hosts");
                Check(!(bool)Call(sizing,"HasOverride",existing)&&existing.ClientSize==original&&CapturedSize(format,existing,"agent.developer")==Size.Empty,"existing reader keeps its original size and snapshot");
                values["agent.developer"]=new Size(960,620);Call(store,"Save",Call(store,"Read"),values);
                using(var next=(Form)Make(reader,false))
                {
                    Check(CapturedSize(format,next,"agent.developer")==new Size(960,620),"subsequent new reader sees the latest saved size");
                    Check(CapturedSize(format,fresh,"agent.developer")==new Size(1100,730),"previous reader retains its saved snapshot after another save");
                    values.Remove("agent.developer");Call(store,"Save",Call(store,"Read"),values);
                    using(var reset=(Form)Make(reader,false))
                        Check(!(bool)Call(sizing,"HasOverride",reset)&&CapturedSize(format,reset,"agent.developer")==Size.Empty&&!reset.Visible,"removing an override restores defaults for the next reader");
                    Check(CapturedSize(format,next,"agent.developer")==new Size(960,620),"restoring defaults leaves an existing reader unchanged");
                    Check(!existing.Visible&&!fresh.Visible&&!next.Visible,"core reader checks never show a window");
                }
            }
        }
        string before=WpsShared.SourcePaths.ReadAllText(path);
        using(var editor=(Form)Make(settings))
        {
            // A visual adjustment is a draft for both hosts, even while disk writes are locked.
            using(var guard=new FileStream(path+".lock",FileMode.Open,FileAccess.ReadWrite,FileShare.None))
            {
                Check(StageSize(editor,"agent.developer",new Size(900,500)),"visual adjustment stages a draft under a disk lock");
                Check(WpsShared.SourcePaths.ReadAllText(path)==before&&!Sizes(store).ContainsKey("agent.developer"),"staging never writes preferences");
                Check(!(bool)settings.GetMethod("SaveEdits",All).Invoke(editor,null)&&((Label)Field(editor,"Status")).Text.Contains("未保存"),"explicit save reports a disk lock failure");
            }
            Check(WpsShared.SourcePaths.ReadAllText(path)==before,"failed explicit save preserves preferences");
            Check((bool)settings.GetMethod("SaveEdits",All).Invoke(editor,null),"explicit save succeeds after releasing the disk lock");
            Check(Sizes(store)["agent.developer"]==new Size(900,500)&&Sizes(store).ContainsKey("future.window"),"explicit save persists the draft and retains future IDs");
            Check(!editor.Visible&&Field(editor,"ActivePreview")==null,"core editor checks never open a preview");
        }
        using(var reopened=(Form)Make(reader,false))
            Check(CapturedSize(format,reopened,"agent.developer")==new Size(900,500),"new reader uses explicitly saved visual adjustment");
        before=WpsShared.SourcePaths.ReadAllText(path);
        using(var cancelled=(Form)Make(settings))
            Check(StageSize(cancelled,"agent.developer",new Size(980,630)),"cancelled editor can hold an unsaved draft");
        Check(WpsShared.SourcePaths.ReadAllText(path)==before,"disposing an unsaved editor discards its draft");
        using(var stale=(Form)Make(settings))
        {
            StageSize(stale,"agent.developer",new Size(960,620));
            var other=Sizes(store);other["agent.user"]=new Size(1100,720);Call(store,"Save",Call(store,"Read"),other);
            string external=WpsShared.SourcePaths.ReadAllText(path);
            Check(!(bool)settings.GetMethod("SaveEdits",All).Invoke(stale,null)&&((Label)Field(stale,"Status")).Text.Contains("未保存"),"stale editor rejects an explicit save after another writer");
            Check(WpsShared.SourcePaths.ReadAllText(path)==external&&Sizes(store)["agent.user"]==new Size(1100,720),"conflict preserves the other writer's preferences");
        }
        Check(Application.OpenForms.Count==openForms,"core checks leave no open forms");
        Call(store,"Save",Call(store,"Read"),new Dictionary<string,Size>());
    }
    [STAThread]static int Main(string[] args)
    {
        Console.OutputEncoding=Encoding.UTF8;
        string root=Path.Combine(Path.GetTempPath(),"dialog-sizes-"+Guid.NewGuid().ToString("N"));
        Type store=null,format=null;
        try
        {
            Application.EnableVisualStyles();var asm=Assembly.LoadFrom(Path.GetFullPath(args[0]));string ns=asm.GetName().Name+".";
            store=asm.GetType(ns+"DialogSizeStore",true);store.GetField("RootOverride",All).SetValue(null,root);
            bool core=Array.IndexOf(args,"--core")>=0;
            format=asm.GetType(ns+"FormatSession",true);format.GetField("SourceOverride",All).SetValue(null,Path.Combine(root,"customization.xml"));Call(format,"Initialize");
            Type catalog=asm.GetType(ns+"DialogCatalog",true),sizing=asm.GetType(ns+"DialogSizing",true),settings=asm.GetType(ns+"DialogSizesForm",true);
            Call(sizing,"Initialize");string path=(string)store.GetProperty("PathForUser",All).GetValue(null,null);
            Array definitions=(Array)catalog.GetField("All",All).GetValue(null);var identities=new HashSet<string>();
            Type factory=asm.GetType(ns+"DialogPreviewFactory",true);
            Check(!Directory.Exists(root),"default measurement creates no preferences, business files or logs");
            Check((Size)Field(Call(catalog,"Find","agent.dialogsizes"),"Default")==new Size(1520,640),"settings default is 1520 logical client width");
            // 5 个 PDF 工具窗口位于 #endif 之后的公共区段，两个宿主都会加载。
            Check(definitions.Length==(ns.StartsWith("Word")?17:42),"complete host-specific catalog count including migration preset windows");
            if(!ns.StartsWith("Word"))foreach(string id in new[]{"migration.preset.name","migration.preset.manager","migration.preset.relocate"})
                Check(Call(catalog,"Find",id)!=null,"registered preset window: "+id);
            foreach(object d in definitions){Check(identities.Add((string)Field(d,"Id")),"unique stable identity");Size size=(Size)Field(d,"Default"),min=(Size)Field(d,"Minimum");Check(size.Width>=min.Width&&size.Height>=min.Height,"default meets client minimum");}
            object empty=Call(store,"Read");Check(((Dictionary<string,Size>)Field(empty,"Values")).Count==0,"missing preferences use defaults");
            VerifyPasswordPreview(asm,ns,catalog,factory);
            if(core){VerifyCore(asm,ns,store,sizing,format,settings);Console.WriteLine("DialogSizesProbe core: "+checks+" PASS, 0 FAIL");return 0;}
            foreach(object d in definitions)
            using(var normal=(Form)Call(factory,"CreateProgramDefault",(string)Field(d,"Id")))
            using(var visual=(Form)Call(factory,"Create",(string)Field(d,"Id"),(Size)Field(d,"Default")))
            {
                normal.Show();visual.Show();Application.DoEvents();
                Size expected=(Size)Field(d,"Default"),actual=(Size)Call(sizing,"LogicalClient",normal),preview=(Size)Call(sizing,"LogicalClient",visual);
                Check(!(bool)Call(sizing,"HasOverride",normal),"normal default is not a user override "+Field(d,"Id"));
                Check(actual==preview,"normal and visual default client match "+Field(d,"Id"));
                float scale=(float)Call(sizing,"Scale",normal);Size frame=new Size(normal.Width-normal.ClientSize.Width,normal.Height-normal.ClientSize.Height);
                Size minimum=(Size)Field(d,"Minimum");Rectangle work=Screen.FromControl(normal).WorkingArea;int margin=(int)Math.Round(24*scale);
                Size expectedMinimum=new Size(Math.Min((int)Math.Round(minimum.Width*scale)+frame.Width,work.Width-margin),Math.Min((int)Math.Round(minimum.Height*scale)+frame.Height,work.Height-margin));
                Check(normal.MinimumSize==expectedMinimum,"minimum managed from same client size service "+Field(d,"Id"));
                Console.WriteLine("SIZE "+Field(d,"Id")+" catalog="+expected.Width+"x"+expected.Height+" normal="+actual.Width+"x"+actual.Height+" preview="+preview.Width+"x"+preview.Height+" minimum="+minimum.Width+"x"+minimum.Height+" dpi="+(int)Math.Round(scale*96));
                if(!(bool)Call(sizing,"WasClamped",normal))Check(actual==expected,"shown default matches catalog "+Field(d,"Id"));
                else Check(actual.Width<=expected.Width&&actual.Height<=expected.Height,"screen crop explains smaller default "+Field(d,"Id"));
                normal.Close();visual.Close();
            }
            Check(!Directory.Exists(root),"all default previews remain side-effect free");
            var values=new Dictionary<string,Size>{{"agent.developer",new Size(1100,730)},{"future.window",new Size(850,500)}};
            Call(store,"Save",empty,values);Check(WpsShared.SourcePaths.Exists(path)&&File.ReadAllBytes(path)[0]!=239,"atomic UTF8 XML without BOM");
            object saved=Call(store,"Read");Check(((Dictionary<string,Size>)Field(saved,"Values"))["agent.developer"]==new Size(1100,730),"round-trip logical client size");
            using(var form=(Form)Make(asm.GetType(ns+"AgentDocumentForm",true),false))
            {
                Check((bool)Call(sizing,"HasOverride",form),"both hosts apply saved sizes to newly opened windows without restarting");form.Show();Application.DoEvents();
                Check(form.ClientSize.Width>=720&&form.ClientSize.Height>=480,"shown reader has safe client area");
                form.ClientSize=new Size(1000,600);form.Close();
            }
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values"))["agent.developer"]==new Size(1100,730),"temporary drag never persists");
            using(var form=(Form)Make(settings))
            {
                Invoke(form,"ExpandAll");DataGridView grid=(DataGridView)Field(form,"Grid");
                var listed=new HashSet<string>();foreach(DataGridViewRow r in grid.Rows)if(RowId(r)!=null)listed.Add(RowId(r));
                Check(listed.Count==definitions.Length,"tree covers every registered window once");
                Check(grid.Columns.Count==8&&grid.Columns[0].HeaderText=="选项卡归属"&&grid.Columns[7].HeaderText=="可视化调整","eight requested headers");
                Check(grid.DefaultCellStyle.Alignment==DataGridViewContentAlignment.MiddleCenter&&grid.ColumnHeadersDefaultCellStyle.Alignment==DataGridViewContentAlignment.MiddleCenter,"headers and body centered");
                foreach(DataGridViewRow r in grid.Rows)if(RowId(r)!=null)
                    Check(Convert.ToString(r.Cells[7].Value)=="自定调整"&&r.Cells[7].GetType().Name=="RoundedAdjustmentCell","rounded adjustment caption "+RowId(r));
                DataGridViewRow row=WindowRow(grid);
                row.Cells[5].Value="1";row.Cells[6].Value="1";
                Check(!(bool)settings.GetMethod("SaveEdits",All).Invoke(form,null),"reject below-minimum size");
                settings.GetMethod("LoadSettings",All).Invoke(form,null);Invoke(form,"ExpandAll");row=WindowRow(grid);
                row.Cells[5].Value="1000";row.Cells[6].Value="";
                Check(!(bool)settings.GetMethod("SaveEdits",All).Invoke(form,null),"reject one blank dimension");
                settings.GetMethod("LoadSettings",All).Invoke(form,null);Invoke(form,"ExpandAll");row=WindowRow(grid);form.Show();Application.DoEvents();
                row.Cells[5].Value="1000";row.Cells[6].Value="750";
                Check((bool)settings.GetMethod("SaveEdits",All).Invoke(form,null),"valid editor save succeeds");
                Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).ContainsKey("future.window"),"editor preserves future stable IDs");
                grid.ClearSelection();row.Selected=true;FindButton(form,"恢复选中默认").PerformClick();row=WindowRow(grid);
                Check(Convert.ToString(row.Cells[5].Value)=="" && Convert.ToString(row.Cells[6].Value)=="","selected restore clears both dimensions");
                Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).ContainsKey(RowId(row)),"unsaved restore does not persist");
                Check((bool)settings.GetMethod("SaveEdits",All).Invoke(form,null) && !((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).ContainsKey(RowId(row)),"selected restore persists only after save");
                FindButton(form,"恢复全部默认").PerformClick();row=WindowRow(grid);
                Check((bool)settings.GetMethod("SaveEdits",All).Invoke(form,null) && ((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==1,"all registered defaults restored while future IDs survive");
                row.Cells[5].Value="1000";row.Cells[6].Value="750";form.Close();
            }
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==1,"close discards unsaved edits");
            Call(store,"Save",Call(store,"Read"),new Dictionary<string,Size>());Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==0,"reset removes overrides");
            object a=Call(store,"Read"),b=Call(store,"Read");Call(store,"Save",a,values);
            try{Call(store,"Save",b,new Dictionary<string,Size>());throw new Exception("stale write accepted");}catch(InvalidOperationException){checks++;}
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==2,"concurrent writer preserved");
            File.WriteAllText(path,"<bad>");Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==0,"corrupt XML defaults");
            File.WriteAllText(path,"<DialogSizes version=\"1\"><Window id=\"agent.user\" width=\"2\" height=\"3\"/><Window id=\"agent.developer\" width=\"1100\" height=\"730\"/></DialogSizes>");
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==1,"invalid entry does not discard valid entry");
            File.WriteAllText(path,"<!DOCTYPE x [<!ENTITY y SYSTEM 'file:///not-used'>]><DialogSizes version=\"1\"/>");
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==0,"DTD prohibited");
            File.WriteAllText(path,"<DialogSizes version=\"2\"/>");
            try{Call(store,"Save",Call(store,"Read"),values);throw new Exception("future format overwritten");}catch(InvalidOperationException){checks++;}
            File.Delete(path);Call(store,"Save",Call(store,"Read"),values);object before=Call(store,"Read");
            using(var guard=new FileStream(path+".lock",FileMode.Open,FileAccess.ReadWrite,FileShare.None))
            {try{Call(store,"Save",before,new Dictionary<string,Size>());throw new Exception("locked save accepted");}catch(IOException){checks++;}}
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==2,"failed save preserves original file");
            values["agent.user"]=new Size(int.MaxValue,int.MaxValue);Call(store,"Save",Call(store,"Read"),values);
            using(var form=(Form)Make(asm.GetType(ns+"AgentDocumentForm",true),true))
            {form.Show();Application.DoEvents();Rectangle area=Screen.FromControl(form).WorkingArea;Check(form.Width<=area.Width && form.Height<=area.Height,"large size is bounded to screen");form.Close();}
            Check(((Dictionary<string,Size>)Field(Call(store,"Read"),"Values"))["agent.user"].Height==int.MaxValue,"screen clamp never changes saved logical size");
            if(!ns.StartsWith("Word"))
            {
                values["file.organize"]=new Size(980,640);values["formula.round"]=new Size(800,640);Call(store,"Save",Call(store,"Read"),values);
                using(var form=(Form)Make(asm.GetType(ns+"OrganizeFilesForm",true))){form.Show();Application.DoEvents();Check((bool)Call(sizing,"HasOverride",form)&&form.ClientSize.Height>=600,"organizer compact Shown does not override custom height");form.Close();}
                object round=Enum.Parse(asm.GetType(ns+"FormulaOperation",true),"Round");
                using(var form=(Form)Make(asm.GetType(ns+"FormulaOptionsForm",true),round)){form.Show();Application.DoEvents();Check((bool)Call(sizing,"HasOverride",form)&&form.ClientSize.Height>=600,"formula compact Shown does not override custom height");form.Close();}
            }

            // Mouse pill hit area and keyboard activation exercise the actual wired grid events.
            using(var form=(Form)Make(settings))
            {
                form.Show();Application.DoEvents();Invoke(form,"ExpandAll");var grid=(DataGridView)Field(form,"Grid");var row=WindowRow(grid);grid.CurrentCell=row.Cells[7];
                foreach(var key in new[]{Keys.Enter,Keys.Space})
                {
                    var e=new KeyEventArgs(key);typeof(DataGridView).GetMethod("OnKeyDown",All).Invoke(grid,new object[]{e});
                    Check(Field(form,"ActivePreview") is Form&&e.Handled&&e.SuppressKeyPress,"rounded adjustment activates from keyboard "+key);
                    ((Form)Field(form,"ActivePreview")).Close();Application.DoEvents();
                    foreach(DataGridViewRow current in grid.Rows)if(RowId(current)==RowId(row)){row=current;break;}grid.CurrentCell=row.Cells[7];
                }
                var outside=new DataGridViewCellMouseEventArgs(7,row.Index,1,1,new MouseEventArgs(MouseButtons.Left,1,1,1,0));
                typeof(DataGridView).GetMethod("OnCellMouseClick",All).Invoke(grid,new object[]{outside});Check(Field(form,"ActivePreview")==null,"cell padding does not activate adjustment");
                var inside=new DataGridViewCellMouseEventArgs(7,row.Index,grid.Columns[7].Width/2,row.Height/2,new MouseEventArgs(MouseButtons.Left,1,0,0,0));
                typeof(DataGridView).GetMethod("OnCellMouseClick",All).Invoke(grid,new object[]{inside});Check(Field(form,"ActivePreview") is Form,"rounded adjustment mouse center opens preview");
                ((Form)Field(form,"ActivePreview")).Close();form.Close();
            }

            // Safe, real-class previews: all business controls disabled, no host required.
            string fingerprint=WpsShared.SourcePaths.ReadAllText(path);int windowsBefore=Application.OpenForms.Count;
            foreach(object d in definitions)
            using(var preview=(Form)Call(factory,"Create",(string)Field(d,"Id"),(Size)Field(d,"Default")))
            {
                preview.Show();Application.DoEvents();Check((bool)Call(factory,"IsPreview",preview),"registered preview "+Field(d,"Id"));
                foreach(Control c in Controls(preview))if(c is Button && c.Text!="关闭"&&c.Text!="取消")Check(!c.Enabled,"business button disabled "+Field(d,"Id")+" "+c.Text);
                Check(preview.FormBorderStyle==FormBorderStyle.Sizable,"preview is resizable "+Field(d,"Id"));preview.Close();
            }
            Check(WpsShared.SourcePaths.ReadAllText(path)==fingerprint&&Application.OpenForms.Count==windowsBefore,"preview creation/close does not persist or leave forms");
            Call(store,"Save",Call(store,"Read"),new Dictionary<string,Size>());
            using(var form=(Form)Make(settings))
            {
                form.Show();Application.DoEvents();Invoke(form,"ExpandAll");DataGridView grid=(DataGridView)Field(form,"Grid");
                DataGridViewRow user=null,parent=null,developer=null;
                foreach(DataGridViewRow r in grid.Rows){if(RowId(r)=="agent.user")user=r;if(RowId(r)=="agent.developer")developer=r;if((string)Field(r.Tag,"Id")=="menuAgentWorkflow")parent=r;}
                user.Cells[5].Value="1000";user.Cells[6].Value="750";
                object parentNode=parent.Tag;Invoke(form,"Toggle",parentNode);Invoke(form,"Toggle",parentNode);
                foreach(DataGridViewRow r in grid.Rows)if(RowId(r)=="agent.user")user=r;
                Check(Convert.ToString(user.Cells[5].Value)=="1000","collapse retains draft");
                Check((bool)settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.developer"}),"open current-row preview");
                Form preview=(Form)Field(form,"ActivePreview");Application.DoEvents();Check(!grid.Enabled,"editor locked during preview");
                ResizePreview(preview,new Size(900,500),true);preview.Close();Application.DoEvents();
                var visual=Sizes(store);Check(!visual.ContainsKey("agent.developer"),"both hosts stage drag-close adjustments until explicit save");
                Check(!visual.ContainsKey("agent.user"),"visual save does not save other manual drafts");
                foreach(DataGridViewRow r in grid.Rows)if(RowId(r)=="agent.user")user=r;Check(Convert.ToString(user.Cells[5].Value)=="1000","other draft remains editable");
                string beforePreview=WpsShared.SourcePaths.ReadAllText(path);settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.developer"});preview=(Form)Field(form,"ActivePreview");Application.DoEvents();
                Check((Size)Call(sizing,"LogicalClient",preview)==new Size(900,500),"second visual adjustment continues from 900 by 500");
                preview.Close();Check(WpsShared.SourcePaths.ReadAllText(path)==beforePreview,"no user resize means no write");
                using(var reopened=(Form)Make(settings))
                {
                    reopened.Show();Application.DoEvents();settings.GetMethod("OpenPreview",All).Invoke(reopened,new object[]{"agent.developer"});Form again=(Form)Field(reopened,"ActivePreview");Application.DoEvents();
                    Check((Size)Call(sizing,"LogicalClient",again)==(Size)Field(Call(catalog,"Find","agent.developer"),"Default"),"reopening discards unsaved drafts for both hosts");again.Close();reopened.Close();
                }
                settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.developer"});preview=(Form)Field(form,"ActivePreview");Application.DoEvents();
                ResizePreview(preview,new Size(960,620),true);
                Call(store,"Save",Call(store,"Read"),new Dictionary<string,Size>{{"agent.user",new Size(1100,720)}});preview.Close();
                Check(!(bool)settings.GetMethod("SaveEdits",All).Invoke(form,null)&&((Label)Field(form,"Status")).Text.Contains("未保存")&&Sizes(store).ContainsKey("agent.user"),"both hosts detect another writer on explicit save and preserve its preferences");
                Invoke(form,"LoadSettings");Invoke(form,"ExpandAll");
                string beforeMove=WpsShared.SourcePaths.ReadAllText(path);settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.developer"});preview=(Form)Field(form,"ActivePreview");Application.DoEvents();
                ResizePreview(preview,new Size(980,630),false);preview.Close();Check(WpsShared.SourcePaths.ReadAllText(path)==beforeMove,"moving and programmatic client changes never save a preview size");
                foreach(DataGridViewRow r in grid.Rows)if(RowId(r)=="agent.user")user=r;
                user.Cells[5].Value="10";user.Cells[6].Value="20";
                Check(!(bool)settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.user"})&&Field(form,"ActivePreview")==null,"invalid draft blocks preview before construction");
                user.Cells[5].Value=int.MaxValue.ToString();user.Cells[6].Value=int.MaxValue.ToString();
                string unclamped=WpsShared.SourcePaths.ReadAllText(path);settings.GetMethod("OpenPreview",All).Invoke(form,new object[]{"agent.user"});preview=(Form)Field(form,"ActivePreview");Application.DoEvents();
                Check(preview.Width<=Screen.FromControl(preview).WorkingArea.Width,"preview is bounded to the working area");preview.Close();Check(WpsShared.SourcePaths.ReadAllText(path)==unclamped,"screen crop without drag never becomes a saved size");
                Invoke(form,"LoadSettings");Invoke(form,"ExpandAll");
                string locked=WpsShared.SourcePaths.ReadAllText(path);
                using(var guard=new FileStream(path+".lock",FileMode.Open,FileAccess.ReadWrite,FileShare.None))
                    Check(StageSize(form,"agent.user",new Size(950,600))&&!(bool)settings.GetMethod("SaveEdits",All).Invoke(form,null)&&((Label)Field(form,"Status")).Text.Contains("未保存"),"both hosts stage drafts under a disk lock and reject explicit writes");
                Check(WpsShared.SourcePaths.ReadAllText(path)==locked,"visual write failure preserves prior preferences");
                Invoke(form,"LoadSettings");Invoke(form,"ExpandAll");
                foreach(DataGridViewRow r in grid.Rows)if((string)Field(r.Tag,"Id")=="menuAgentWorkflow")parent=r;
                grid.ClearSelection();parent.Selected=true;FindButton(form,"恢复选中默认").PerformClick();Check((bool)settings.GetMethod("SaveEdits",All).Invoke(form,null)&&((Dictionary<string,Size>)Field(Call(store,"Read"),"Values")).Count==0,"parent restore reaches hidden descendants");
                form.Close();
            }
            Type message=asm.GetType(ns+"UiMessageForm",true);
            foreach(var buttons in new[]{MessageBoxButtons.OK,MessageBoxButtons.YesNo,MessageBoxButtons.OKCancel})
            using(var form=(Form)Make(message,"测试","确认",buttons,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2))
            {
                DialogResult dismiss=(DialogResult)message.GetProperty("DismissResult",All).GetValue(form,null);
                Check(dismiss==(buttons==MessageBoxButtons.YesNo?DialogResult.No:buttons==MessageBoxButtons.OKCancel?DialogResult.Cancel:DialogResult.OK),"correct dismissal result "+buttons);
                if(buttons!=MessageBoxButtons.OK){Button positive=(Button)Field(form,"Positive"),negative=(Button)Field(form,"Negative");Check(positive.DialogResult==(buttons==MessageBoxButtons.YesNo?DialogResult.Yes:DialogResult.OK)&&negative.DialogResult==dismiss&&form.AcceptButton==null,"confirmation button semantics "+buttons);}
            }
            VerifyCore(asm,ns,store,sizing,format,settings);
            Console.WriteLine("DialogSizesProbe: "+checks+" PASS, 0 FAIL");return 0;
        }
        catch(Exception ex){Console.WriteLine("FAIL: "+ex);return 1;}
        finally
        {
            if(store!=null)store.GetField("RootOverride",All).SetValue(null,null);
            if(format!=null)format.GetField("SourceOverride",All).SetValue(null,null);
            string temp=Path.GetFullPath(Path.GetTempPath());string full=Path.GetFullPath(root);
            if(Directory.Exists(full)&&Path.GetDirectoryName(full)==temp.TrimEnd(Path.DirectorySeparatorChar)&&Path.GetFileName(full).StartsWith("dialog-sizes-",StringComparison.Ordinal))Directory.Delete(full,true);
        }
    }
}
