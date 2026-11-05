using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using System.Xml;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class DialogTreeNode
    {
        internal string Id, Tab, Group, Label, DialogId;
        internal int Depth;
        internal bool Expanded;
        internal DialogTreeNode Parent;
        internal readonly List<DialogTreeNode> Children=new List<DialogTreeNode>();
        internal DialogDefinition Definition {get{return DialogCatalog.Find(DialogId);}}
    }
    internal static class DialogTree
    {
        private static readonly Dictionary<string,string> Windows=new Dictionary<string,string> {
            {"btnTracePrecedents","trace.precedents"},{"btnTraceDependents","trace.dependents"},
            {"btnFilterColumns","columns.filter"},{"btnCustomIFError","formula.iferror"},{"btnCustomRound","formula.round"},{"btnCustomFunction","formula.custom"},
            {"btnCustomCalculateAdd","formula.add"},{"btnCustomCalculateSubtract","formula.subtract"},
            {"btnCustomCalculateMultiply","formula.multiply"},{"btnCustomCalculateDivide","formula.divide"},
            {"btnLinkAssistant","links.assistant"},{"btnOrganizeFiles","file.organize"},{"btnRenameFiles","file.rename"},{"btnMergeHeaders","merge.headers"},
            {"btnEliminationReceivable","elimination.receivable"},{"btnEliminationRevenueCost","elimination.revenuecost"},
            {"btnDataMigration","migration.data"},
            {"btnTextExtractFilter","conversion.text"},{"btnReferenceStyleConversion","formula.references"},
            {"btnCompanyWorkbooks","company.workbooks"},
            {"btnCurrentPdf","file.pdf.current"},{"btnConvertPdf","file.pdf.convert"},{"btnWatermarkPdf","file.pdf.watermark"},{"btnSplitPdf","file.pdf.split"},{"btnMergePdf","file.pdf.merge"},
            {"btnViewLinks","word.links"},{"btnAgentReadme","agent.developer"},{"btnAgentUserGuide","agent.user"},
            {"btnAgentSort","agent.sort"},{"btnAgentNames","agent.name"},{"btnAgentFormats","agent.format"},{"btnAgentRequirements","agent.notes"},{"btnFileNames","file.names"},{"btnAgentDialogSizes","agent.dialogsizes"}
        };
        internal static List<DialogTreeNode> Build()
        {
            var xml=new XmlDocument {XmlResolver=null};xml.LoadXml(RibbonCustomization.Effective(false).OuterXml);
            var roots=new List<DialogTreeNode>();
            foreach(XmlElement tab in xml.SelectNodes("//*[local-name()='tab']"))
                foreach(XmlElement group in tab.SelectNodes("*[local-name()='group']"))
                    Add(group,roots,null,tab.GetAttribute("label"),group.GetAttribute("label"));
            var common=new DialogTreeNode {Id="common.messages",Tab="通用",Group="通用提示",Label="提示与确认"};
            foreach(var d in DialogCatalog.All)if(d.Id.StartsWith("message.",StringComparison.Ordinal))
                common.Children.Add(new DialogTreeNode {Id=d.Id,Tab=common.Tab,Group=common.Group,Label=d.Name,DialogId=d.Id,Depth=1,Parent=common});
            roots.Add(common);return roots;
        }
        private static void Add(XmlElement element,List<DialogTreeNode> output,DialogTreeNode parent,string tab,string group)
        {
            foreach(XmlNode child in element.ChildNodes)
            {
                var item=child as XmlElement;if(item==null)continue;
                if(item.LocalName=="button"||item.LocalName=="toggleButton"||item.LocalName=="menu")
                {
                    string id=item.GetAttribute("id"),window;Windows.TryGetValue(id,out window);
                    var node=new DialogTreeNode {Id=id,Tab=tab,Group=group,Label=item.GetAttribute("label"),DialogId=window,Parent=parent,Depth=parent==null?0:parent.Depth+1};
                    output.Add(node);Add(item,node.Children,node,tab,group);
                    if(window=="agent.developer")
                    {
                        var passwordWindow=DialogCatalog.Find("agent.developerpassword");
                        node.Children.Add(new DialogTreeNode {Id=passwordWindow.Id,Tab=tab,Group=group,Label=passwordWindow.Name,DialogId=passwordWindow.Id,Parent=node,Depth=node.Depth+1});
                    }
#if !WPS_WORD
                    if(window=="trace.precedents")
                    {
                        var errorSources=DialogCatalog.Find("trace.errorsources");
                        node.Children.Add(new DialogTreeNode {Id=errorSources.Id,Tab=tab,Group=group,Label=errorSources.Name,DialogId=errorSources.Id,Parent=node,Depth=node.Depth+1});
                    }
                    if(window=="migration.data")foreach(string dialogId in new[]{"migration.preset.name","migration.preset.manager","migration.preset.relocate"})
                    {
                        var presetWindow=DialogCatalog.Find(dialogId);
                        node.Children.Add(new DialogTreeNode {Id=dialogId,Tab=tab,Group=group,Label=presetWindow.Name,DialogId=dialogId,Parent=node,Depth=node.Depth+1});
                    }
                    if(window=="file.names")foreach(string dialogId in new[]{"file.ocrprefixes","file.ocrfields"})
                    {
                        var nestedWindow=DialogCatalog.Find(dialogId);
                        node.Children.Add(new DialogTreeNode {Id=dialogId,Tab=tab,Group=group,Label=nestedWindow.Name,DialogId=dialogId,Parent=node,Depth=node.Depth+1});
                    }
#endif
                }
                else Add(item,output,parent,tab,group);
            }
        }
        internal static IEnumerable<DialogTreeNode> All(IEnumerable<DialogTreeNode> nodes)
        {foreach(var n in nodes){yield return n;foreach(var child in All(n.Children))yield return child;}}
    }
    internal sealed class DialogSizeDraft
    {internal string Width="",Height="";}
    internal sealed class RoundedAdjustmentCell : DataGridViewButtonCell
    {
        private bool pressed;
        internal static Rectangle ButtonBounds(Rectangle cell,Font font,float dpi)
        {
            float scale=dpi/96f*font.SizeInPoints/10f;
            int w=Math.Min((int)Math.Round(120*scale),Math.Max(1,cell.Width-(int)Math.Round(12*scale)));
            int h=Math.Min((int)Math.Round(34*scale),Math.Max(1,cell.Height-(int)Math.Round(10*scale)));
            return new Rectangle(cell.Left+(cell.Width-w)/2,cell.Top+(cell.Height-h)/2,w,h);
        }
        protected override void Paint(Graphics g,Rectangle clip,Rectangle bounds,int row,DataGridViewElementStates state,object value,object formatted,string error,DataGridViewCellStyle style,DataGridViewAdvancedBorderStyle border,DataGridViewPaintParts parts)
        {
            base.Paint(g,clip,bounds,row,state,value,formatted,error,style,border,parts&~(DataGridViewPaintParts.ContentBackground|DataGridViewPaintParts.ContentForeground|DataGridViewPaintParts.Focus));
            Font body=style.Font??DataGridView.Font;Rectangle rect=ButtonBounds(bounds,body,g.DpiX);
            var owner=DataGridView.FindForm() as DialogSizesForm;bool enabled=DataGridView.Enabled&&owner!=null&&owner.CanPreview;
            Point mouse=DataGridView.PointToClient(Control.MousePosition);bool hover=enabled&&rect.Contains(mouse);
            Color fill=!enabled?Color.FromArgb(235,238,244):pressed?Color.FromArgb(198,218,255):hover?Color.FromArgb(219,233,255):Ui.SelectBack;
            var smoothing=g.SmoothingMode;g.SmoothingMode=System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using(var path=Ui.RoundRect(new RectangleF(rect.X+.5f,rect.Y+.5f,rect.Width-1,rect.Height-1),rect.Height/2f))
            using(var brush=new SolidBrush(fill))using(var pen=new Pen(Ui.GridFrame)){g.FillPath(brush,path);g.DrawPath(pen,path);}
            if(enabled&&DataGridView.ContainsFocus&&DataGridView.CurrentCell==this)
                using(var path=Ui.RoundRect(new RectangleF(rect.X+2,rect.Y+2,rect.Width-4,rect.Height-4),Math.Max(2,rect.Height/2f-2)))
                using(var pen=new Pen(Ui.Blue,2)){g.DrawPath(pen,path);}
            using(var font=new Font(body,FontStyle.Bold))TextRenderer.DrawText(g,DialogFormatting.AdjustmentCaption(owner),DialogFormatting.AdjustmentFont(owner,font),rect,enabled?DialogFormatting.AdjustmentColor(owner,Ui.Blue):Color.FromArgb(128,139,158),TextFormatFlags.HorizontalCenter|TextFormatFlags.VerticalCenter|TextFormatFlags.SingleLine|TextFormatFlags.NoPrefix);
            g.SmoothingMode=smoothing;
        }
        protected override void OnMouseDown(DataGridViewCellMouseEventArgs e){pressed=e.Button==MouseButtons.Left;base.OnMouseDown(e);DataGridView.InvalidateCell(this);}
        protected override void OnMouseUp(DataGridViewCellMouseEventArgs e){pressed=false;base.OnMouseUp(e);DataGridView.InvalidateCell(this);}
        protected override void OnMouseLeave(int row){pressed=false;base.OnMouseLeave(row);DataGridView.InvalidateCell(this);}
    }
    internal sealed class PreviewResizeCapture : NativeWindow, IDisposable
    {
        internal bool UserSized;
        private readonly Form form;
        internal PreviewResizeCapture(Form target)
        {form=target;form.HandleCreated+=Attach;form.HandleDestroyed+=Detach;if(form.IsHandleCreated)AssignHandle(form.Handle);}
        private void Attach(object sender,EventArgs e){AssignHandle(form.Handle);}
        private void Detach(object sender,EventArgs e){ReleaseHandle();}
        protected override void WndProc(ref Message message)
        {
            if(message.Msg==0x0231)UserSized=false; // WM_ENTERSIZEMOVE includes moving.
            if(message.Msg==0x0214)UserSized=true;  // WM_SIZING only occurs for edge dragging.
            base.WndProc(ref message);
        }
        public void Dispose(){form.HandleCreated-=Attach;form.HandleDestroyed-=Detach;ReleaseHandle();}
    }
    internal sealed class DialogSizesForm : Form
    {
        private static DialogSizesForm window;
        internal readonly DataGridView Grid=new DataGridView {Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,AllowUserToResizeRows=false,AutoGenerateColumns=false,SelectionMode=DataGridViewSelectionMode.FullRowSelect,MultiSelect=true,ScrollBars=ScrollBars.Both};
        internal readonly Label Status=Ui.Label("",9,false);
        internal readonly List<DialogTreeNode> Nodes=DialogTree.Build();
        private readonly Dictionary<string,DialogSizeDraft> drafts=new Dictionary<string,DialogSizeDraft>(StringComparer.Ordinal);
        private DialogSizeSnapshot snapshot;
        private readonly bool previewOnly;
        private readonly FlowLayoutPanel actions=Ui.ActionBar();
        internal Form ActivePreview;
        private bool rebuilding, fitting;
        internal bool HasUnsavedDraft {get {CommitRows();foreach(var item in DialogCatalog.All){Size previous;bool has=snapshot.Values.TryGetValue(item.Id,out previous);var value=drafts[item.Id];if(value.Width.Length==0&&value.Height.Length==0){if(has)return true;}else if(!has||value.Width!=previous.Width.ToString()||value.Height!=previous.Height.ToString())return true;}return false;}}
        internal bool CanPreview {get {return !previewOnly&&ActivePreview==null;}}
        internal static void Open()
        {if(window!=null&&!window.IsDisposed){if(window.WindowState==FormWindowState.Minimized)window.WindowState=FormWindowState.Normal;window.Activate();return;}window=new DialogSizesForm();window.FormClosed+=delegate {window.Dispose();window=null;};window.Show();}
        internal static void CloseWindow(){if(window!=null&&!window.IsDisposed)window.Close();}
        internal DialogSizesForm():this(false){}
        internal DialogSizesForm(bool readOnlyPreview)
        {
            previewOnly=readOnlyPreview;
            Text="交互弹窗 · "+typeof(Connect).Assembly.GetName().Name;DialogSizing.Prepare(this,"agent.dialogsizes","settings",Ui.Blue);DialogSizing.EnsureAllDefaults(this);KeyPreview=true;
            var root=new TableLayoutPanel {Dock=DockStyle.Fill,Padding=new Padding(16),ColumnCount=1,RowCount=4};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var note=Ui.Label("宽高为96 DPI客户区逻辑像素，不含标题栏；同时留空使用程序默认。手填后保存；“自定调整”续用该行尺寸。"+
"拖动后关闭只更新草稿，请点击保存设置。"+
                "屏幕限制不改写设置；系统选择窗口与安装器不在清单中。",9,false);note.AutoSize=true;note.AutoEllipsis=false;note.Margin=new Padding(0,0,0,12);
            Ui.StyleGrid(Grid);Grid.RowHeadersVisible=false;Grid.DefaultCellStyle.Alignment=DataGridViewContentAlignment.MiddleCenter;Grid.ColumnHeadersDefaultCellStyle.Alignment=DataGridViewContentAlignment.MiddleCenter;
            Grid.ColumnHeadersDefaultCellStyle.WrapMode=DataGridViewTriState.False;Grid.DefaultCellStyle.WrapMode=DataGridViewTriState.False;
            string[] labels={"选项卡归属","分组情况","命令按钮","程序默认尺寸","最小尺寸","自定宽度","自定高度","可视化调整"};
            for(int i=0;i<labels.Length;i++)
            {DataGridViewColumn col=i==7?(DataGridViewColumn)new DataGridViewButtonColumn {CellTemplate=new RoundedAdjustmentCell(),FlatStyle=FlatStyle.Flat} :new DataGridViewTextBoxColumn();col.Name="c"+i;col.HeaderText=labels[i];col.ReadOnly=i!=5&&i!=6;col.SortMode=DataGridViewColumnSortMode.NotSortable;Grid.Columns.Add(col);}
            Grid.Columns[2].DefaultCellStyle.Alignment=DataGridViewContentAlignment.MiddleLeft;
            Grid.CellEndEdit+=delegate {CommitRows();FitColumns();};
            Grid.CellPainting+=PaintTree;
            Grid.CellMouseClick+=delegate(object sender,DataGridViewCellMouseEventArgs e){if(e.RowIndex<0||previewOnly||e.Button!=MouseButtons.Left)return;var n=(DialogTreeNode)Grid.Rows[e.RowIndex].Tag;if(e.ColumnIndex==2&&n.Children.Count>0&&e.X<=GlyphLeft(n)+Ui.Px(this,18))Toggle(n);else if(e.ColumnIndex==7&&n.Definition!=null){using(var g=Grid.CreateGraphics())if(RoundedAdjustmentCell.ButtonBounds(new Rectangle(0,0,Grid.Columns[7].Width,Grid.Rows[e.RowIndex].Height),Grid.DefaultCellStyle.Font??Grid.Font,g.DpiX).Contains(e.X,e.Y))OpenPreview(n.DialogId);}};
            Grid.KeyDown+=TreeKeys;
            Grid.DataError+=delegate(object sender,DataGridViewDataErrorEventArgs e){e.ThrowException=false;Status.Text="请输入正整数宽高，或同时留空。";};
            Grid.Resize+=delegate {FitColumns();};Grid.FontChanged+=delegate {FitColumns();};
            Status.AutoSize=true;Status.AutoEllipsis=false;Status.Margin=new Padding(0,8,0,4);
            var save=Ui.Button("保存设置",110);Ui.StylePrimary(save);save.Click+=delegate {if(SaveEdits()) {
CustomizationSaving.AfterSave(this);
            }};
            var selected=Ui.Button("恢复选中默认",140);selected.Click+=delegate {CommitRows();foreach(DataGridViewRow row in Grid.SelectedRows)RestoreNode((DialogTreeNode)row.Tag);RebuildRows();Status.Text="已恢复所选及其窗口后代，点击保存后生效。";};
            var all=Ui.Button("恢复全部默认",140);all.Click+=delegate {CommitRows();foreach(var value in drafts.Values){value.Width="";value.Height="";}RebuildRows();Status.Text="已恢复全部默认，点击保存后生效。";};
            var reload=Ui.Button("重新加载",110);reload.Click+=delegate {LoadSettings();};var close=Ui.Button("关闭",90);close.Click+=delegate {Close();};CancelButton=close;
            actions.Controls.AddRange(new Control[]{save,selected,all,reload,close});root.Controls.Add(note,0,0);root.Controls.Add(Ui.Card(Grid,null),0,1);root.Controls.Add(Status,0,2);root.Controls.Add(actions,0,3);Controls.Add(root);
            root.Resize+=delegate {note.MaximumSize=new Size(Math.Max(100,root.ClientSize.Width-root.Padding.Horizontal),0);Status.MaximumSize=note.MaximumSize;actions.MaximumSize=note.MaximumSize;};
            FormClosing+=delegate {if(ActivePreview!=null&&!ActivePreview.IsDisposed)ActivePreview.Close();};
            LoadSettings();if(previewOnly){Grid.ReadOnly=true;foreach(Control control in actions.Controls)if(control!=close)control.Enabled=false;}
        }
        internal void LoadSettings()
        {
            snapshot=DialogSizeStore.Read();drafts.Clear();
            foreach(var d in DialogCatalog.All){Size s;bool set=snapshot.Values.TryGetValue(d.Id,out s);drafts[d.Id]=new DialogSizeDraft {Width=set?s.Width.ToString():"",Height=set?s.Height.ToString():""};}
            RebuildRows();Status.Text=snapshot.Warning.Length>0?snapshot.Warning:"共 "+DialogCatalog.All.Length+" 类可自定弹窗；三角展开菜单，手填修改需保存。";
        }
        private void CommitRows()
        {if(rebuilding)return;Grid.EndEdit();foreach(DataGridViewRow row in Grid.Rows){var n=(DialogTreeNode)row.Tag;if(n.Definition!=null){drafts[n.DialogId].Width=Convert.ToString(row.Cells[5].Value).Trim();drafts[n.DialogId].Height=Convert.ToString(row.Cells[6].Value).Trim();}}}
        internal void RebuildRows()
        {
            rebuilding=true;try{Grid.Rows.Clear();foreach(var n in Nodes)AddRow(n);FitColumns();}finally{rebuilding=false;}
        }
        private void AddRow(DialogTreeNode n)
        {
            var d=n.Definition;var value=d==null?null:drafts[d.Id];
            int i=Grid.Rows.Add(n.Tab,n.Group,n.Label,d==null?"—":d.Default.Width+" × "+d.Default.Height,d==null?"—":d.Minimum.Width+" × "+d.Minimum.Height,d==null?"—":value.Width,d==null?"—":value.Height,d==null?"—":"自定调整");
            var row=Grid.Rows[i];row.Tag=n;row.Cells[2].ToolTipText=n.Children.Count>0?"三角展开/折叠；左右方向键浏览":n.Label;
            row.Cells[2].Style.Alignment=DataGridViewContentAlignment.MiddleLeft;row.Cells[2].ToolTipText=n.Label+(n.Children.Count>0?"；三角展开/折叠；左右方向键浏览":"");
            row.Cells[2].Style.Padding=new Padding(GlyphLeft(n)+Ui.Px(this,20),0,Ui.Px(this,8),0);
            if(d==null){row.Cells[5].ReadOnly=row.Cells[6].ReadOnly=true;row.Cells[7]=new DataGridViewTextBoxCell {Value="—"};}
            if(n.Expanded)foreach(var child in n.Children)AddRow(child);
        }
        internal void ExpandAll(){CommitRows();foreach(var n in DialogTree.All(Nodes))n.Expanded=true;RebuildRows();}
        internal void Toggle(DialogTreeNode n){CommitRows();n.Expanded=!n.Expanded;RebuildRows();SelectNode(n);}
        private void SelectNode(DialogTreeNode n){foreach(DataGridViewRow row in Grid.Rows)if(row.Tag==n){Grid.ClearSelection();row.Selected=true;Grid.CurrentCell=row.Cells[2];return;}}
        private void TreeKeys(object sender,KeyEventArgs e)
        {
            if(previewOnly||Grid.CurrentRow==null||Grid.IsCurrentCellInEditMode)return;var n=(DialogTreeNode)Grid.CurrentRow.Tag;
            if(e.KeyCode==Keys.Right){if(n.Children.Count>0){if(!n.Expanded)Toggle(n);else SelectNode(n.Children[0]);}e.Handled=true;e.SuppressKeyPress=true;}
            else if(e.KeyCode==Keys.Left){if(n.Expanded)Toggle(n);else if(n.Parent!=null)SelectNode(n.Parent);e.Handled=true;e.SuppressKeyPress=true;}
            else if((e.KeyCode==Keys.Enter||e.KeyCode==Keys.Space)&&Grid.CurrentCell.ColumnIndex==7){if(n.Definition!=null)OpenPreview(n.DialogId);e.Handled=true;e.SuppressKeyPress=true;}
        }
        private int GlyphLeft(DialogTreeNode n){return Ui.Px(this,8+n.Depth*12);}
        private void PaintTree(object sender,DataGridViewCellPaintingEventArgs e)
        {
            if(e.RowIndex<0||e.ColumnIndex!=2)return;var n=(DialogTreeNode)Grid.Rows[e.RowIndex].Tag;e.Paint(e.ClipBounds,DataGridViewPaintParts.All);e.Handled=true;
            int x=e.CellBounds.Left+GlyphLeft(n),y=e.CellBounds.Top+e.CellBounds.Height/2,k=Ui.Px(this,4);
            using(var pen=new Pen(Ui.GridFrame)){for(int level=0;level<n.Depth;level++){int line=e.CellBounds.Left+Ui.Px(this,12+level*12);e.Graphics.DrawLine(pen,line,e.CellBounds.Top,line,e.CellBounds.Bottom);}}
            if(n.Children.Count>0)using(var brush=new SolidBrush(Ui.Blue)){Point[] triangle=n.Expanded?new[]{new Point(x,y-k),new Point(x+2*k,y-k),new Point(x+k,y+k)}:new[]{new Point(x,y-k),new Point(x,y+k),new Point(x+2*k,y)};e.Graphics.FillPolygon(brush,triangle);}
        }
        internal void FitColumns()
        {
            if(fitting||Grid.Columns.Count!=8)return;fitting=true;try{
                Ui.EnsureGridHeaders(Grid);
                Font body=Grid.DefaultCellStyle.Font??Grid.Font;float factor=DialogSizing.Scale(this)*body.SizeInPoints/10f;
                int headerPadding=Math.Max((int)Math.Round(26*factor),Grid.ColumnHeadersDefaultCellStyle.Padding.Horizontal+10);
                int bodyPadding=Math.Max((int)Math.Round(26*factor),Grid.DefaultCellStyle.Padding.Horizontal+10);
                var widths=new int[8];using(var g=Grid.CreateGraphics())using(var bold=new Font(body,FontStyle.Bold))
                {
                    for(int i=0;i<8;i++)widths[i]=Math.Max(Grid.Columns[i].MinimumWidth,TextRenderer.MeasureText(g,Grid.Columns[i].HeaderText,Grid.Columns[i].HeaderCell.InheritedStyle.Font??Grid.Font,Size.Empty,TextFormatFlags.SingleLine|TextFormatFlags.NoPrefix).Width+headerPadding);
                    foreach(var n in DialogTree.All(Nodes))
                    {var d=n.Definition;DialogSizeDraft draft;drafts.TryGetValue(n.DialogId??"",out draft);string[] texts={n.Tab,n.Group,n.Label,d==null?"—":d.Default.Width+" × "+d.Default.Height,d==null?"—":d.Minimum.Width+" × "+d.Minimum.Height,draft==null?"—":draft.Width,draft==null?"—":draft.Height,DialogFormatting.AdjustmentCaption(this)};for(int i=0;i<8;i++)widths[i]=Math.Max(widths[i],TextRenderer.MeasureText(g,texts[i],i==7?bold:body).Width+(i==2?GlyphLeft(n)+Ui.Px(this,28):bodyPadding));}
                }
                widths[7]=Math.Max(widths[7],(int)Math.Round(136*factor));
                for(int i=0;i<8;i++){Grid.Columns[i].AutoSizeMode=DataGridViewAutoSizeColumnMode.None;Grid.Columns[i].MinimumWidth=widths[i];Grid.Columns[i].Width=widths[i];}
                Grid.RowTemplate.Height=(int)Math.Round(46*factor);foreach(DataGridViewRow row in Grid.Rows){row.Height=Grid.RowTemplate.Height;var n=row.Tag as DialogTreeNode;if(n!=null)row.Cells[2].Style.Padding=new Padding(GlyphLeft(n)+Ui.Px(this,20),0,Ui.Px(this,8),0);}
            }finally{fitting=false;}
        }
        private void RestoreNode(DialogTreeNode n){if(n.Definition!=null){drafts[n.DialogId].Width="";drafts[n.DialogId].Height="";}foreach(var child in n.Children)RestoreNode(child);}
        private Size Requested(string id)
        {
            var value=drafts[id];int w,h;
            if(value.Width.Length==0&&value.Height.Length==0)return DialogCatalog.Find(id).Default;
            if(!int.TryParse(value.Width,NumberStyles.None,CultureInfo.InvariantCulture,out w)||!int.TryParse(value.Height,NumberStyles.None,CultureInfo.InvariantCulture,out h)||!DialogCatalog.Valid(id,new Size(w,h)))throw new InvalidOperationException(DialogCatalog.Find(id).Name+"：请填写不小于最小尺寸的正整数宽高，或同时留空。");
            return new Size(w,h);
        }
        internal bool SaveEdits()
        {
            if(previewOnly||ActivePreview!=null)return false;
            try{CommitRows();var values=new Dictionary<string,Size>(snapshot.Values,StringComparer.Ordinal);foreach(var d in DialogCatalog.All){var value=drafts[d.Id];if(value.Width.Length==0&&value.Height.Length==0)values.Remove(d.Id);else values[d.Id]=Requested(d.Id);}DialogSizeStore.Save(snapshot,values);snapshot=DialogSizeStore.Read();Status.Text=
"已保存，下次打开对应弹窗时生效。";
                return true;}
            catch(Exception ex){Status.Text="未保存："+ex.Message;return false;}
        }
        internal bool OpenPreview(string id)
        {
            if(previewOnly||ActivePreview!=null)return false;
            try
            {
                CommitRows();Size desired=Requested(id);Form form=DialogPreviewFactory.Create(id,desired);ActivePreview=form;Grid.Enabled=false;actions.Enabled=false;
                var capture=new PreviewResizeCapture(form);
                Size start=Size.Empty;Size? adjusted=null;
                form.ResizeBegin+=delegate {start=form.ClientSize;};
                form.ResizeEnd+=delegate {if(capture.UserSized&&form.WindowState==FormWindowState.Normal&&start!=Size.Empty&&form.ClientSize!=start)adjusted=DialogSizing.LogicalClient(form);};
                form.FormClosed+=delegate {
                    try{if(adjusted.HasValue)PersistVisualSize(id,adjusted.Value);else Status.Text="预览已关闭，未拖动尺寸，不改写设置。";}
                    finally {capture.Dispose();ActivePreview=null;Grid.Enabled=true;actions.Enabled=true;RebuildRows();form.Dispose();}
                };
                form.Show(this);Size opened=DialogSizing.LogicalClient(form);Status.Text="自定调整：本次打开 "+opened.Width+" × "+opened.Height+
"（逻辑客户区）；拖动后关闭只更新草稿，请点击保存设置。"
                +(DialogSizing.WasClamped(form)?" 受屏幕工作区限制，保存尺寸未改变。":"");return true;
            }
            catch(Exception ex){if(ActivePreview!=null)ActivePreview.Dispose();ActivePreview=null;Grid.Enabled=true;actions.Enabled=true;Status.Text="无法预览："+ex.Message;return false;}
        }
        internal bool PersistVisualSize(string id,Size size)
        {
            if(!DialogCatalog.Valid(id,size)){Status.Text="未保存：调整尺寸小于登记的最小尺寸。";return false;}
            drafts[id].Width=size.Width.ToString();drafts[id].Height=size.Height.ToString();
Status.Text="已更新尺寸草稿；点击保存设置后，下次打开对应弹窗时生效。";return true;
        }
    }
}
