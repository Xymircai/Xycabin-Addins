using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    [DataContract] internal sealed class CustomizationData
    {
        [DataMember] internal int Version = 1;
        [DataMember] internal Dictionary<string,string> Names = new Dictionary<string,string>();
        [DataMember] internal Dictionary<string,List<string>> Orders = new Dictionary<string,List<string>>();
        [DataMember] internal Dictionary<string,FormatRule> Formats = new Dictionary<string,FormatRule>();
        [DataMember(EmitDefaultValue=false)] internal Dictionary<string,int[]> WindowSizes;
        internal CustomizationData Copy() { return LocalCustomization.Decode(LocalCustomization.Encode(this)); }
    }
    [DataContract] internal sealed class FormatRule
    {
        [DataMember] internal string Text, Family;
        [DataMember] internal float Size;
        [DataMember] internal int Style, Color;
        [DataMember] internal bool HasColor;
        // Optional logical pixels at 96 DPI; zero keeps the normal layout.
        [DataMember(EmitDefaultValue=false)] internal int Width96, Height96;
    }
    // XML envelope with a versioned JSON body. No credentials enter this file.
    internal sealed class LocalCustomization
    {
        private const string ChangedFileMessage = "配置已被另一窗口修改，请重新打开后再保存。";
        internal readonly string Path;
        private string digest;
        private bool readOnly;
        internal CustomizationData Data;
        internal bool IsReadable {get{return !readOnly;}}
        internal LocalCustomization(string path) { Path=path; Reload(); }
        internal static byte[] Encode(CustomizationData data)
        { using(var m=new MemoryStream()) { new DataContractJsonSerializer(typeof(CustomizationData)).WriteObject(m,data); return m.ToArray(); } }
        internal static CustomizationData Decode(byte[] bytes)
        { using(var m=new MemoryStream(bytes)) return (CustomizationData)new DataContractJsonSerializer(typeof(CustomizationData)).ReadObject(m); }
        private static string Hash(byte[] bytes) { using(var s=SHA256.Create()) return Convert.ToBase64String(s.ComputeHash(bytes)); }
        private string CurrentDigest() { return File.Exists(Path)?Hash(File.ReadAllBytes(Path)):""; }
        internal void Reload()
        {
            Data=new CustomizationData(); readOnly=false;
            try {
                digest=CurrentDigest(); if(digest=="")return;
                var xml=new XmlDocument {XmlResolver=null};
                using(var reader=XmlReader.Create(Path,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))xml.Load(reader);
                if(xml.DocumentElement.Name!="customization"||xml.DocumentElement.GetAttribute("version")!="1") {readOnly=true;return;}
                var value=Decode(Convert.FromBase64String(xml.DocumentElement.InnerText));
                if(value.Version!=1||value.Names==null||value.Orders==null||value.Formats==null||value.Orders.Values.Any(v=>v==null||v.Any(s=>s==null))||value.Formats.Values.Any(v=>v==null))throw new InvalidDataException();
                Data=value;
            } catch {readOnly=true;}
        }
        internal void Save(CustomizationData data)
        {
            if(readOnly)throw new InvalidDataException("配置无法读取或来自较新版本。请先备份并移走配置文件："+Path);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path));
            using(var guard=new FileStream(Path+".lock",FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None)) {
                if(CurrentDigest()!=digest)throw new IOException(ChangedFileMessage);
                string temp=Path+"."+Guid.NewGuid().ToString("N")+".tmp";
                try {
                    using(var writer=XmlWriter.Create(temp,new XmlWriterSettings {Encoding=new UTF8Encoding(false)})) {
                        writer.WriteStartElement("customization");writer.WriteAttributeString("version","1");writer.WriteString(Convert.ToBase64String(Encode(data)));writer.WriteEndElement();
                    }
                    CommitFile(temp);
                    digest=CurrentDigest();Data=data.Copy();
                } finally {if(File.Exists(temp))File.Delete(temp);}
            }
        }

        private void CommitFile(string temporary)
        {
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    if (attempt > 0 && CurrentDigest() != digest) throw new IOException(ChangedFileMessage);
                    if (File.Exists(Path)) File.Replace(temporary, Path, null);
                    else File.Move(temporary, Path);
                    return;
                }
                catch (IOException error)
                {
                    if (attempt >= 4 || !IsTransientFileLock(error)) throw;
                    System.Threading.Thread.Sleep(50 << attempt);
                }
            }
        }

        private static bool IsTransientFileLock(IOException error)
        {
            const int SharingViolation = unchecked((int)0x80070020);
            const int LockViolation = unchecked((int)0x80070021);
            const int UnableToRemoveReplaced = unchecked((int)0x80070497);
            return error.HResult == SharingViolation || error.HResult == LockViolation ||
                error.HResult == UnableToRemoveReplaced;
        }
    }
    internal static class RibbonCustomization
    {
        private static XmlDocument canonical;
        private static readonly List<object> ribbons=new List<object>();
        internal static readonly LocalCustomization Store=new LocalCustomization(System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),typeof(Connect).Assembly.GetName().Name,"Customization.xml"));
        internal static XmlDocument Canonical {
            get {if(canonical==null){canonical=new XmlDocument {XmlResolver=null};canonical.LoadXml(RibbonConfig.RibbonXml);}return canonical;}
        }
        internal static IEnumerable<XmlElement> Commands {get {return Canonical.SelectNodes("//*[local-name()='button' or local-name()='toggleButton' or local-name()='menu']").Cast<XmlElement>();}}
        internal static XmlElement Find(string id) {return Commands.FirstOrDefault(e=>e.GetAttribute("id")==id);}
        internal static CustomizationData DataForProbe;
        internal static CustomizationData SessionData {get {
return DataForProbe??FormatSession.Data;
        }}
        internal static string Label(string id) {
            string name;var node=Find(id);
            return SessionData.Names.TryGetValue(id,out name)&&!string.IsNullOrWhiteSpace(name)?name:node==null?id:node.GetAttribute("label");
        }
        internal static XmlDocument Effective(bool dynamic)
        {
            var doc=(XmlDocument)Canonical.CloneNode(true);
            foreach(XmlElement menu in doc.SelectNodes("//*[local-name()='menu']")) {
                List<string> order;
                if(SessionData.Orders.TryGetValue(menu.GetAttribute("id"),out order)) {
                    var children=menu.ChildNodes.OfType<XmlElement>().ToList();
                    foreach(var child in children.OrderBy(c=>{int n=order.IndexOf(c.GetAttribute("id"));return n<0?int.MaxValue:n;}))menu.AppendChild(child);
                }
            }
            foreach(XmlElement node in doc.SelectNodes("//*[local-name()='button' or local-name()='toggleButton' or local-name()='menu']")) {
                string id=node.GetAttribute("id");node.SetAttribute("label",Label(id));
                if(dynamic) {node.RemoveAttribute("label");node.SetAttribute("getLabel","GetLabel_"+id);}
            }
            if(dynamic)foreach(XmlElement menu in doc.SelectNodes("//*[local-name()='menu' and not(ancestor::*[local-name()='menu'])]").Cast<XmlElement>().ToList()) {
                var root=doc.CreateElement("dynamicMenu",menu.NamespaceURI);
                foreach(XmlAttribute attr in menu.Attributes)if(attr.Name!="itemSize")root.SetAttribute(attr.Name,attr.Value);
                root.SetAttribute("getContent","GetContent_"+menu.GetAttribute("id"));root.SetAttribute("invalidateContentOnDrop","true");
                menu.ParentNode.ReplaceChild(root,menu);
            }
            return doc;
        }
        internal static string Content(string id) {
            var doc=Effective(false);var source=doc.SelectNodes("//*[local-name()='menu']").Cast<XmlElement>().First(e=>e.GetAttribute("id")==id);
            var result=new XmlDocument {XmlResolver=null};var menu=result.CreateElement("menu",source.NamespaceURI);result.AppendChild(menu);
            foreach(XmlNode child in source.ChildNodes) {var copy=result.ImportNode(child,true);RemoveRootSize(copy);menu.AppendChild(copy);}
            return result.OuterXml;
        }
        private static void RemoveRootSize(XmlNode n) {var e=n as XmlElement;if(e!=null)e.RemoveAttribute("size");foreach(XmlNode c in n.ChildNodes)RemoveRootSize(c);}
        internal static void Register(object ribbon) {HostDialogService.Register();if(ribbon!=null&&!ribbons.Contains(ribbon))ribbons.Add(ribbon);if(DataForProbe==null)LiveCustomization.Start();}
        internal static void Refresh() {
            var failures=new List<string>();
            foreach(var ribbon in ribbons)try {var typed=ribbon as Microsoft.Office.Core.IRibbonUI;if(typed!=null)typed.Invalidate();else {dynamic r=ribbon;r.Invalidate();}}catch {failures.Add("刷新失败");}
            if(failures.Count>0)throw new InvalidOperationException("配置已保存，但部分功能区刷新失败；请重新打开文档核对。");
        }
        internal static void Clear() {LiveCustomization.Stop();ribbons.Clear();canonical=null;}
    }
    internal static class CustomizationWindows
    {
        private static readonly Dictionary<string,Form> windows=new Dictionary<string,Form>();
        internal static void Open(string mode) {
            if(mode=="format"){FormatEditorThread.Open(Connect.CustomizationOwner());return;}
            Form current;if(windows.TryGetValue(mode,out current)&&!current.IsDisposed){current.Activate();return;}
            Form form=mode=="format"?(Form)new DialogFormatEditor(false):mode=="notes"?(Form)new RequirementsForm(false):new CommandCustomizationForm(mode,false);
            windows[mode]=form;form.FormClosed+=delegate {windows.Remove(mode);form.Dispose();};
            Connect.ShowCustomizationWindow(form);
        }
        internal static void CloseAll() {FormatEditorThread.Close();foreach(var f in windows.Values.ToArray())if(!f.IsDisposed){var notes=f as RequirementsForm;if(notes!=null)notes.CloseForShutdown();else f.Close();}foreach(var key in windows.Keys.ToArray())if(windows[key].IsDisposed)windows.Remove(key);}
    }
    internal sealed class SortableCommandTree:TreeView
    {
        private TreeNode dragCandidate;private Point dragOrigin;
        [DllImport("user32.dll")]private static extern int GetScrollPos(IntPtr window,int bar);
        internal readonly bool DragEnabled;
        internal SortableCommandTree():this(true){}
        internal SortableCommandTree(bool drag) {DragEnabled=drag;DrawMode=TreeViewDrawMode.OwnerDrawAll;ShowLines=false;ShowPlusMinus=false;Indent=28;ItemHeight=30;FullRowSelect=true;ShowNodeToolTips=true;}
        internal Rectangle Expander(TreeNode node) {int edge=Math.Max(Ui.Px(this,18),Font.Height);int left=Ui.Px(this,4)+node.Level*Indent-(IsHandleCreated?GetScrollPos(Handle,0):0);return new Rectangle(left,node.Bounds.Top+(ItemHeight-edge)/2,edge,edge);}
        internal Rectangle Grip(TreeNode node) {if(!DragEnabled||node.Parent==null)return Rectangle.Empty;var box=Expander(node);return new Rectangle(box.Right+Ui.Px(this,4),box.Top,Ui.Px(this,5)+Ui.Px(this,3),box.Height);}
        internal Rectangle TextArea(TreeNode node) {var grip=Grip(node);int left=(grip.IsEmpty?Expander(node).Right:grip.Right)+Ui.Px(this,4);return new Rectangle(left,node.Bounds.Top,Math.Max(1,ClientSize.Width-left),ItemHeight);}
        internal TreeNode RowAt(Point point) {for(var node=TopNode;node!=null;node=node.NextVisibleNode)if(point.Y>=node.Bounds.Top&&point.Y<node.Bounds.Top+ItemHeight)return node;return null;}
        protected override void OnDrawNode(DrawTreeNodeEventArgs e) {
            bool selected=(e.State&TreeNodeStates.Selected)!=0;var row=new Rectangle(0,e.Bounds.Top,ClientSize.Width,ItemHeight);
            using(var background=new SolidBrush(selected?Ui.SelectBack:BackColor))e.Graphics.FillRectangle(background,row);
            e.Graphics.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            if(e.Node.Nodes.Count>0) {
                var r=Expander(e.Node);using(var shape=Ui.RoundRect(r,5))using(var brush=new SolidBrush(Ui.Blue))e.Graphics.FillPath(brush,shape);
                using(var pen=new Pen(Color.White,2)){int x=r.Left+r.Width/2,y=r.Top+r.Height/2;e.Graphics.DrawLine(pen,x-4,y,x+4,y);if(!e.Node.IsExpanded)e.Graphics.DrawLine(pen,x,y-4,x,y+4);}
            }
            var grip=Grip(e.Node);int step=Ui.Px(this,5),dotSize=Ui.Px(this,3);
            if(DragEnabled&&e.Node.Parent!=null)using(var dot=new SolidBrush(Ui.GridHeadInk))for(int x=0;x<2;x++)for(int y=0;y<3;y++)e.Graphics.FillEllipse(dot,grip.Left+x*step,e.Bounds.Top+ItemHeight/2-step-dotSize/2+y*step,dotSize,dotSize);
            TextRenderer.DrawText(e.Graphics,e.Node.Text,Font,TextArea(e.Node),ForeColor,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.NoPrefix|TextFormatFlags.NoPadding|TextFormatFlags.EndEllipsis);
            e.Node.ToolTipText=e.Node.Text;
        }
        protected override void OnFontChanged(EventArgs e) {base.OnFontChanged(e);ItemHeight=Math.Max(30,(int)(Font.Height*1.7));Indent=Math.Max(28,Font.Height*2);}
        protected override void OnMouseDown(MouseEventArgs e) {dragCandidate=null;var node=RowAt(e.Location);if(node!=null&&node.Nodes.Count>0&&Expander(node).Contains(e.Location)){SelectedNode=node;if(node.IsExpanded)node.Collapse();else node.Expand();Invalidate();return;}base.OnMouseDown(e);if(node!=null){SelectedNode=node;if(DragEnabled&&e.Button==MouseButtons.Left&&node.Parent!=null){dragCandidate=node;dragOrigin=e.Location;}}}
        protected override void OnMouseMove(MouseEventArgs e) {base.OnMouseMove(e);if(dragCandidate!=null&&e.Button==MouseButtons.Left&&!new Rectangle(dragOrigin.X-SystemInformation.DragSize.Width/2,dragOrigin.Y-SystemInformation.DragSize.Height/2,SystemInformation.DragSize.Width,SystemInformation.DragSize.Height).Contains(e.Location)){var node=dragCandidate;dragCandidate=null;OnItemDrag(new ItemDragEventArgs(MouseButtons.Left,node));}}
        protected override void OnMouseUp(MouseEventArgs e) {dragCandidate=null;base.OnMouseUp(e);}
        protected override void OnItemDrag(ItemDragEventArgs e) {dragCandidate=null;base.OnItemDrag(e);}
    }
    internal sealed class CommandCustomizationForm:Form
    {
        private readonly string mode;
        private readonly CustomizationData draft;
        private readonly LocalCustomization session;
        private readonly TreeView tree;
        private readonly TextBox name=new TextBox {Width=230,Name="customName"};
        private readonly NumericUpDown number=new NumericUpDown {Minimum=1,Maximum=1000,Width=70,Name="orderNumber"};
        private readonly Label original=new Label {AutoSize=true,Name="originalName"};
        internal CommandCustomizationForm(string mode,bool preview) {
            this.mode=mode;session=new LocalCustomization(RibbonCustomization.Store.Path);draft=session.Data.Copy();Text=mode=="sort"?"自定排序":"自定名称";
            tree=new SortableCommandTree(mode=="sort");tree.Dock=DockStyle.Fill;tree.HideSelection=false;tree.AllowDrop=mode=="sort";
            DialogSizing.Prepare(this,"agent."+mode,"settings",Ui.Blue);
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(16)};
            layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(tree);var edit=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true};edit.Controls.Add(original);
            if(mode=="name"){edit.Controls.Add(Ui.Grip(name,new Label {Text="新名称（清空恢复原名）",AutoSize=true}));}else {edit.Controls.Add(new Label {Text="同一菜单中的顺序号",AutoSize=true});edit.Controls.Add(number);AddButton(edit,"设置顺序",()=>MoveTo((int)number.Value-1));AddButton(edit,"上移",()=>MoveTo(tree.SelectedNode.Index-1));AddButton(edit,"下移",()=>MoveTo(tree.SelectedNode.Index+1));}
            layout.Controls.Add(edit);var bar=new FlowLayoutPanel {AutoSize=true,Dock=DockStyle.Fill};
            AddButton(bar,"恢复所选",RestoreSelected);AddButton(bar,"恢复全部",()=>{if(mode=="name")draft.Names.Clear();else draft.Orders.Clear();Build();});
            AddButton(bar,"保存",()=>{session.Save(draft);

                CustomizationSaving.AfterSave(this);});AddButton(bar,"取消",()=>Close());layout.Controls.Add(bar);Controls.Add(layout);
            tree.AfterSelect+=delegate {var n=tree.SelectedNode;if(n==null)return;string id=(string)n.Tag;original.Text="原名："+RibbonCustomization.Find(id).GetAttribute("label");name.Text=draft.Names.ContainsKey(id)?draft.Names[id]:"";number.Value=n.Index+1;};
            name.TextChanged+=delegate {if(tree.SelectedNode==null)return;string id=(string)tree.SelectedNode.Tag;draft.Names[id]=name.Text.Trim();tree.SelectedNode.Text=string.IsNullOrWhiteSpace(name.Text)?RibbonCustomization.Find(id).GetAttribute("label"):name.Text.Trim();};
            tree.ItemDrag+=delegate(object sender,ItemDragEventArgs e){if(mode=="sort")tree.DoDragDrop(e.Item,DragDropEffects.Move);};
            tree.DragOver+=delegate(object sender,DragEventArgs e){var target=RowAt(tree.PointToClient(new Point(e.X,e.Y)));var source=e.Data.GetData(typeof(TreeNode)) as TreeNode;e.Effect=source!=null&&target!=null&&source.Parent!=null&&source.Parent==target.Parent?DragDropEffects.Move:DragDropEffects.None;};
            tree.DragDrop+=delegate(object sender,DragEventArgs e){var target=RowAt(tree.PointToClient(new Point(e.X,e.Y)));var source=e.Data.GetData(typeof(TreeNode)) as TreeNode;if(source!=null&&target!=null&&source.Parent!=null&&source.Parent==target.Parent){tree.SelectedNode=source;MoveTo(target.Index);}};
            Build();foreach(var b in DialogFormatting.Walk(this).OfType<RoundedButton>()){b.Radius=14;if(b.Text=="保存"||b.Text=="设置顺序")Ui.StylePrimary(b);}
            if(preview)foreach(Control c in bar.Controls)c.Enabled=false;
        }
        internal static void AddButton(Control parent,string caption,Action action) {var b=Ui.Button(caption,110);b.AutoSize=true;b.Click+=delegate {try {action();}catch(Exception ex){UiDialog.Show(ex.Message,"自定配置",MessageBoxButtons.OK,MessageBoxIcon.Warning);}};parent.Controls.Add(b);}
        internal bool HasUnsavedDraft {get {return !LocalCustomization.Encode(draft).SequenceEqual(LocalCustomization.Encode(session.Data));}}
        private void Build() {
            tree.Nodes.Clear();foreach(var e in RibbonCustomization.Commands)if(!(e.ParentNode is XmlElement)||((XmlElement)e.ParentNode).LocalName!="menu")tree.Nodes.Add(Node(e));
            tree.ExpandAll();if(tree.Nodes.Count>0)tree.SelectedNode=tree.Nodes[0];
        }
        private TreeNode Node(XmlElement e) {
            string id=e.GetAttribute("id"),value;var node=new TreeNode(draft.Names.TryGetValue(id,out value)&&!string.IsNullOrWhiteSpace(value)?value:e.GetAttribute("label")) {Tag=id};
            List<string> order;draft.Orders.TryGetValue(id,out order);
            foreach(var child in e.ChildNodes.OfType<XmlElement>().OrderBy(c=>order==null?0:order.IndexOf(c.GetAttribute("id"))<0?int.MaxValue:order.IndexOf(c.GetAttribute("id"))))node.Nodes.Add(Node(child));return node;
        }
        private void MoveTo(int index) {
            var n=tree.SelectedNode;if(n==null||n.Parent==null)throw new InvalidOperationException("请选择菜单中的子项；不能跨父菜单移动。");
            var parent=n.Parent;index=Math.Max(0,Math.Min(index,parent.Nodes.Count-1));n.Remove();parent.Nodes.Insert(index,n);tree.SelectedNode=n;
            draft.Orders[(string)parent.Tag]=parent.Nodes.Cast<TreeNode>().Select(c=>(string)c.Tag).ToList();
        }
        private void RestoreSelected() {
            var n=tree.SelectedNode;if(n==null)return;string id=(string)n.Tag;
            if(mode=="name")draft.Names.Remove(id);else draft.Orders.Remove(n.Nodes.Count>0?id:n.Parent==null?id:(string)n.Parent.Tag);Build();
        }
        private TreeNode RowAt(Point point) {var sortable=tree as SortableCommandTree;return sortable==null?tree.GetNodeAt(point):sortable.RowAt(point);}
    }
    internal sealed class RequirementsForm:Form
    {
        private readonly TextBox body=new TextBox {Multiline=true,AcceptsReturn=true,AcceptsTab=true,ScrollBars=ScrollBars.Both,Dock=DockStyle.Fill,Name="notesBody"};
        internal static string RootOverride;
        private readonly string path=System.IO.Path.Combine(RootOverride??Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),typeof(Connect).Assembly.GetName().Name,"Requirements.txt");
        private string saved="",recoveredPath;private readonly bool preview;private bool hostClosing;
        internal RequirementsForm(bool preview) {
            this.preview=preview;Text="需求记录";DialogSizing.Prepare(this,"agent.notes","document",Ui.Blue);
            var layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=3,Padding=new Padding(16)};
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label {AutoSize=true,Text="记录需要新增或改善的功能。点击保存后，下次打开可继续编辑。",Name="notesHelp"});layout.Controls.Add(body);
            var bar=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true};CommandCustomizationForm.AddButton(bar,"保存",Save);
            CommandCustomizationForm.AddButton(bar,"重新读取",()=>{if(body.Text!=saved&&UiDialog.Show("放弃未保存内容并重新读取？","需求记录",MessageBoxButtons.YesNo,MessageBoxIcon.Question)!=DialogResult.Yes)return;LoadNotes();});
            CommandCustomizationForm.AddButton(bar,"关闭",()=>Close());layout.Controls.Add(bar);Controls.Add(layout);
            if(!preview){LoadNotes();OfferRecovery();}else body.Text="示例需求（预览不读写记录）";
            FormClosing+=delegate(object sender,FormClosingEventArgs e){if(preview||hostClosing||body.Text==saved)return;var result=UiDialog.Show("是否保存本次修改？","需求记录",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);if(result==DialogResult.Cancel)e.Cancel=true;else if(result==DialogResult.Yes)try {Save();}catch(Exception ex){e.Cancel=true;UiDialog.Show(ex.Message,"保存失败",MessageBoxButtons.OK,MessageBoxIcon.Warning);}else if(recoveredPath!=null&&File.Exists(recoveredPath))File.Delete(recoveredPath);};
        }
        private void OfferRecovery() {
            string directory=System.IO.Path.GetDirectoryName(path);if(!Directory.Exists(directory))return;
            string candidate=Directory.GetFiles(directory,System.IO.Path.GetFileName(path)+".unsaved-*.txt").OrderByDescending(File.GetLastWriteTimeUtc).FirstOrDefault();if(candidate==null)return;
            string contents=File.ReadAllText(candidate,Encoding.UTF8);if(contents==saved){try{File.Delete(candidate);}catch{}return;}
            if(UiDialog.Show("上次退出时保留了未保存的需求草稿，是否恢复？","需求记录",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes){body.Text=contents;recoveredPath=candidate;}else File.Delete(candidate);
        }
        internal void CloseForShutdown() {
            if(!preview&&body.Text!=saved) {
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));string recovery=path+".unsaved-"+Guid.NewGuid().ToString("N")+".txt",temporary=recovery+".tmp";
                try {File.WriteAllText(temporary,body.Text,new UTF8Encoding(false));File.Move(temporary,recovery);}finally {if(File.Exists(temporary))File.Delete(temporary);}
            }
            hostClosing=true;Close();if(!IsDisposed)Dispose();
        }
        private void LoadNotes() {saved=File.Exists(path)?File.ReadAllText(path,Encoding.UTF8):"";body.Text=saved;}
        private void Save() {
            if(preview)return;Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));string temp=path+"."+Guid.NewGuid().ToString("N")+".tmp";
            using(var guard=new FileStream(path+".lock",FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None))
            try {if((File.Exists(path)?File.ReadAllText(path,Encoding.UTF8):"")!=saved)throw new IOException("记录已被其他窗口修改，请重新读取。");File.WriteAllText(temp,body.Text,new UTF8Encoding(false));if(File.Exists(path))File.Replace(temp,path,null);else File.Move(temp,path);saved=body.Text;if(recoveredPath!=null){try{if(File.Exists(recoveredPath))File.Delete(recoveredPath);}catch{}recoveredPath=null;}}finally {if(File.Exists(temp))File.Delete(temp);}
        }
    }
}
