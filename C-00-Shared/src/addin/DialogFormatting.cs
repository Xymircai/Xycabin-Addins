using System.Threading;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class FormatTarget
    {
        public string DisplayLabel {get {return Label;}}
        internal string Id, Label, OriginalText;
        internal bool Caption;
        internal Control Control;
        internal Func<Font> GetFont;
        internal Action<Font> SetFont;
        internal Func<Color> GetColor;
        internal Action<Color> SetColor;
        internal Action<string> SetText;
        internal Font OriginalFont;
        internal Color OriginalColor;
        internal bool ResizeWidth {get {return Id!="window"&&Id!="self.adjust"&&!Id.Contains(".header.")&&DialogControlSizing.Supports(Control);}}
        internal bool ResizeHeight {get {return ResizeWidth&&!(Control is ComboBox)&&!(Control is UpDownBase);}}
    }
    internal static class DialogFormatting
    {
        private static readonly ThreadLocal<Dictionary<Form,List<FormatTarget>>> targetsState=new ThreadLocal<Dictionary<Form,List<FormatTarget>>>(()=>new Dictionary<Form,List<FormatTarget>>());
        private static Dictionary<Form,List<FormatTarget>> targets {get {return targetsState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,List<Font>>> fontsState=new ThreadLocal<Dictionary<Form,List<Font>>>(()=>new Dictionary<Form,List<Font>>());
        private static Dictionary<Form,List<Font>> fonts {get {return fontsState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,string>> identitiesState=new ThreadLocal<Dictionary<Form,string>>(()=>new Dictionary<Form,string>());
        private static Dictionary<Form,string> identities {get {return identitiesState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,FormatRule>> adjustmentState=new ThreadLocal<Dictionary<Form,FormatRule>>(()=>new Dictionary<Form,FormatRule>());
        private static Dictionary<Form,FormatRule> adjustment {get {return adjustmentState.Value;}}
        private static readonly ThreadLocal<Dictionary<Form,List<Action>>> layoutsState=new ThreadLocal<Dictionary<Form,List<Action>>>(()=>new Dictionary<Form,List<Action>>());
        private static Dictionary<Form,List<Action>> layouts {get {return layoutsState.Value;}}
        internal static void Attach(Form form,string id) {
            if(DialogPreviewFactory.IsConstructing)return;
            form.Load+=delegate {Apply(form,id,FormatSession.Capture(form));};
        }
        // Runtime width drags land in the same store the format editor uses, so a dragged
        // width survives restarts and stays visible/editable in 自定格式. Failures only
        // skip persistence; the session keeps the dragged width either way.
        internal static void SaveControlWidth(Form form,Control control,int width96) {
            try {
                List<FormatTarget> list;string id;
                if(form==null||control==null||form.IsDisposed||width96<32||width96>4096
                        ||!targets.TryGetValue(form,out list)||!identities.TryGetValue(form,out id))return;
                var target=list.FirstOrDefault(t=>t.Control==control);if(target==null||!target.ResizeWidth)return;
                var session=new LocalCustomization(RibbonCustomization.Store.Path);if(!session.IsReadable)return;
                var data=session.Data;string key=id+"::"+target.Id;FormatRule rule;
                if(!data.Formats.TryGetValue(key,out rule)||rule==null){rule=new FormatRule();data.Formats[key]=rule;}
                rule.Width96=width96;session.Save(data);
            }
            catch(Exception error){UiDiagnostics.Record("grip-width",error);}
        }
        // Row heights dragged at runtime live in the same store, but under synthetic keys
        // ("row.xxx") that no control owns: DialogControlSizing.Apply looks up
        // "<dialog>::<control id>" and therefore never touches them, so a dragged height
        // cannot be turned into an undocked control. The dialog loads them back itself.
        internal static void SaveRowHeight(Form form,string key,int height96) {
            try {
                string id;
                if(form==null||form.IsDisposed||string.IsNullOrEmpty(key)||height96<24||height96>4096
                        ||!identities.TryGetValue(form,out id))return;
                var session=new LocalCustomization(RibbonCustomization.Store.Path);if(!session.IsReadable)return;
                var data=session.Data;string store=id+"::"+key;FormatRule rule;
                if(!data.Formats.TryGetValue(store,out rule)||rule==null){rule=new FormatRule();data.Formats[store]=rule;}
                rule.Height96=height96;session.Save(data);
            }
            catch(Exception error){UiDiagnostics.Record("grip-height",error);}
        }
        internal static int LoadRowHeight(Form form,string key) {
            try {
                string id;
                if(form==null||string.IsNullOrEmpty(key)||!identities.TryGetValue(form,out id))return 0;
                var session=new LocalCustomization(RibbonCustomization.Store.Path);if(!session.IsReadable)return 0;
                FormatRule rule;if(!session.Data.Formats.TryGetValue(id+"::"+key,out rule)||rule==null)return 0;
                int value=rule.Height96;return value>=24&&value<=4096?value:0;
            }
            catch(Exception error){UiDiagnostics.Record("grip-height",error);return 0;}
        }
        internal static List<FormatTarget> Register(Form form,string id) {
            List<FormatTarget> existing;if(targets.TryGetValue(form,out existing))return existing;
            var result=new List<FormatTarget>();identities[form]=id;
            Add(result,form,"window","窗口标题",true,form.Text,()=>form.Font,f=>form.Font=f,()=>form.ForeColor,c=>form.ForeColor=c,s=>form.Text=s);
            var counts=new Dictionary<string,int>();var fieldNames=new Dictionary<Control,string>();
            for(Type type=form.GetType();type!=null&&type!=typeof(Form);type=type.BaseType)foreach(var field in type.GetFields(BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.DeclaredOnly))if(typeof(Control).IsAssignableFrom(field.FieldType)){var control=field.GetValue(form) as Control;if(control!=null&&!fieldNames.ContainsKey(control))fieldNames[control]=field.Name;}
            foreach(Control c in Walk(form)) {
                string fieldName;fieldNames.TryGetValue(c,out fieldName);string stem=!string.IsNullOrEmpty(c.Name)?c.Name:!string.IsNullOrEmpty(fieldName)?fieldName:c.GetType().Name+"."+Stable(c.Text);
                int count;counts.TryGetValue(stem,out count);counts[stem]=count+1;string key=stem+"."+count;
                bool caption=(c is Label||c is ButtonBase||c is GroupBox)&&!DynamicLabel(c,stem);
                if(c is Panel||c is SplitContainer||c is TableLayoutPanel||c is FlowLayoutPanel)continue;
                Control target=c;
                var bodyGrid=target as DataGridView;
                Add(result,target,key,caption?c.Text:Friendly(c),caption,c.Text,
                    ()=>bodyGrid==null?target.Font:bodyGrid.DefaultCellStyle.Font??target.Font,
                    f=>{target.Font=f;if(bodyGrid!=null)bodyGrid.DefaultCellStyle.Font=f;},
                    ()=>bodyGrid==null?target.ForeColor:bodyGrid.DefaultCellStyle.ForeColor,
                    color=>{target.ForeColor=color;if(bodyGrid!=null)bodyGrid.DefaultCellStyle.ForeColor=color;},s=>target.Text=s);
                var grid=c as DataGridView;if(grid!=null)foreach(DataGridViewColumn column in grid.Columns) {
                    var col=column;string columnId=key+".header."+(string.IsNullOrEmpty(col.Name)?col.Index.ToString(CultureInfo.InvariantCulture):col.Name);
                    Add(result,grid,columnId,"表头："+col.HeaderText,true,col.HeaderText,()=>col.HeaderCell.Style.Font??grid.ColumnHeadersDefaultCellStyle.Font??grid.Font,
                        f=>col.HeaderCell.Style.Font=f,()=>col.HeaderCell.Style.ForeColor.IsEmpty?grid.ColumnHeadersDefaultCellStyle.ForeColor:col.HeaderCell.Style.ForeColor,
                        color=>col.HeaderCell.Style.ForeColor=color,s=>col.HeaderText=s);
                }
            }
            if(id=="agent.dialogsizes") {
                var grid=Walk(form).OfType<DataGridView>().FirstOrDefault();
                if(grid!=null)Add(result,grid,"self.adjust","自绘按钮：自定调整",true,"自定调整",()=>grid.Font,f=>{},()=>Ui.Blue,c=>{},s=>{});
            }
            var reset=new List<Action>();foreach(Control c in Walk(form).Concat(new[]{form})) {
                var scroll=c as ScrollableControl;if(scroll!=null){bool original=scroll.AutoScroll;reset.Add(()=>scroll.AutoScroll=original);}
                var button=c as Button;if(button!=null){bool auto=button.AutoSize;var mode=button.AutoSizeMode;reset.Add(()=>{button.AutoSize=auto;button.AutoSizeMode=mode;});}
                var rounded=c as RoundedButton;if(rounded!=null){Color disabled=rounded.DisabledForeColor;reset.Add(()=>rounded.DisabledForeColor=disabled);}
                var label=c as Label;if(label!=null){var max=label.MaximumSize;reset.Add(()=>label.MaximumSize=max);}
                var grid=c as DataGridView;if(grid!=null){var wrap=grid.ColumnHeadersDefaultCellStyle.WrapMode;var mode=grid.ColumnHeadersHeightSizeMode;int height=grid.ColumnHeadersHeight;reset.Add(()=>{grid.ColumnHeadersDefaultCellStyle.WrapMode=wrap;grid.ColumnHeadersHeightSizeMode=mode;grid.ColumnHeadersHeight=height;});}
            }
            targets[form]=result;fonts[form]=new List<Font>();layouts[form]=reset;
            DialogControlSizing.Register(form,result);
            form.Disposed+=delegate {List<Font> allocated;if(fonts.TryGetValue(form,out allocated))foreach(var f in allocated)f.Dispose();foreach(var t in result)t.OriginalFont.Dispose();targets.Remove(form);fonts.Remove(form);identities.Remove(form);adjustment.Remove(form);layouts.Remove(form);};
            return result;
        }
        private static string Stable(string value) {
            // Deterministic caption fingerprint: unaffected by layout or sibling insertions.
            uint hash=2166136261;foreach(char c in value??"")hash=unchecked((hash^c)*16777619);return hash.ToString("x8",CultureInfo.InvariantCulture);
        }
        private static bool DynamicLabel(Control c,string identity) {
            if(!(c is Label))return false;
            string n=identity.ToLowerInvariant();return new[]{"status","current","result","summary","original","detail","validation","hint","preview","empty","tip","path","progress"}.Any(n.Contains)||
                new[]{"正在","当前","已选择","文档：","工作簿：","目标：","目录：","路径："}.Any(p=>c.Text.StartsWith(p,StringComparison.Ordinal));
        }
        private static string Friendly(Control c) {
            if(c is Label)return string.IsNullOrEmpty(c.Text)?"状态文字":c.Text;
            string kind=c is TextBoxBase?"文本输入框":c is DataGridView?"表格数据":c is ComboBox?"下拉选择":c is TreeView?"菜单树":c is ListBox?"列表":c is NumericUpDown?"数值输入":"控件样式";
            if(c.Name=="notesBody")return "需求内容";
            if(c.Parent!=null){int index=c.Parent.Controls.IndexOf(c);var label=index>0?c.Parent.Controls[index-1] as Label:null;if(label!=null&&!string.IsNullOrEmpty(label.Text))return label.Text+" · "+kind;}return kind;
        }
        internal static IEnumerable<Control> Walk(Control root) {foreach(Control c in root.Controls){yield return c;foreach(var child in Walk(c))yield return child;}}
        private static void Add(List<FormatTarget> list,Control c,string key,string label,bool caption,string text,Func<Font> font,Action<Font> setFont,Func<Color> color,Action<Color> setColor,Action<string> setText) {
            list.Add(new FormatTarget {Id=key,Label=label,Caption=caption,Control=c,OriginalText=text,GetFont=font,SetFont=setFont,GetColor=color,SetColor=setColor,SetText=setText,OriginalFont=(Font)font().Clone(),OriginalColor=color()});
        }
        internal static void Apply(Form form,string id,CustomizationData data) {
            var list=Register(form,id);var previous=fonts[form];var allocated=new List<Font>();form.SuspendLayout();
            try {
                DialogControlSizing.Reset(form);
                foreach(var reset in layouts[form])reset();
                foreach(var t in list) {
                    FormatRule rule;bool has=data.Formats.TryGetValue(id+"::"+t.Id,out rule);
                    if(t.Caption)t.SetText(has&&rule.Text!=null?rule.Text:t.OriginalText);
                    t.SetFont(t.OriginalFont);t.SetColor(t.OriginalColor);
                    if(has&&t.Id!="window") {
                        Font f=null;try {
                            float size=rule.Size>=6&&rule.Size<=36?rule.Size:t.OriginalFont.Size;
                            using(var family=new FontFamily(string.IsNullOrWhiteSpace(rule.Family)?t.OriginalFont.FontFamily.Name:rule.Family)) {
                                var style=(FontStyle)(rule.Style&3);f=new Font(family,size,family.IsStyleAvailable(style)?style:FontStyle.Regular);
                            }
                        }catch {f=new Font(t.OriginalFont.FontFamily,t.OriginalFont.Size,t.OriginalFont.Style);}
                        allocated.Add(f);t.SetFont(f);if(rule.HasColor){t.SetColor(Color.FromArgb(rule.Color));var rounded=t.Control as RoundedButton;if(rounded!=null)rounded.DisabledForeColor=Color.FromArgb(rule.Color);}
                    }
                }
                FormatRule draw;if(data.Formats.TryGetValue(id+"::self.adjust",out draw))adjustment[form]=draw;else adjustment.Remove(form);
                if(data.Formats.Keys.Any(k=>k.StartsWith(id+"::",StringComparison.Ordinal)&&k!=id+"::window")) {
                    foreach(var label in Walk(form).OfType<Label>())if(label.AutoSize&&label.Parent!=null&&label.Text.Length>30)label.MaximumSize=new Size(Math.Max(100,label.Parent.ClientSize.Width-32),0);
                    foreach(var button in Walk(form).OfType<Button>())if(button.Parent is FlowLayoutPanel||button.Parent is TableLayoutPanel){var target=list.FirstOrDefault(t=>t.Control==button);FormatRule rule;if(target!=null&&data.Formats.TryGetValue(id+"::"+target.Id,out rule)&&AppearanceChanged(target,rule)){button.AutoSize=true;button.AutoSizeMode=AutoSizeMode.GrowAndShrink;}}
                }
                fonts[form]=allocated;
                DialogControlSizing.Apply(form,id,list,data);
            } finally {form.ResumeLayout(true);form.PerformLayout();var sizes=form as DialogSizesForm;if(sizes!=null)sizes.FitColumns();else foreach(var grid in Walk(form).OfType<DataGridView>())Ui.AutoFitColumns(grid);foreach(var old in previous)old.Dispose();form.Invalidate(true);}
        }
        private static bool AppearanceChanged(FormatTarget target,FormatRule rule){return rule.Text!=null&&rule.Text!=target.OriginalText||!string.IsNullOrEmpty(rule.Family)&&rule.Family!=target.OriginalFont.Name||rule.Size>0&&Math.Abs(rule.Size-target.OriginalFont.Size)>.01||(rule.Style&3)!=((target.OriginalFont.Bold?1:0)|(target.OriginalFont.Italic?2:0))||rule.HasColor&&rule.Color!=target.OriginalColor.ToArgb();}
        internal static string AdjustmentCaption(Form owner) {FormatRule rule;return owner!=null&&adjustment.TryGetValue(owner,out rule)&&rule.Text!=null?rule.Text:"自定调整";}
        internal static Font AdjustmentFont(Form owner,Font fallback) {
            List<FormatTarget> list;var t=owner!=null&&targets.TryGetValue(owner,out list)?list.FirstOrDefault(x=>x.Id=="self.adjust"):null;
            if(t==null)return fallback;string id=identities[owner];FormatRule r;if(!adjustment.TryGetValue(owner,out r))return fallback;
            // Drawing reuses a font owned by the form rather than allocating every frame.
            var allocated=fonts[owner];return allocated.LastOrDefault(f=>Math.Abs(f.Size-r.Size)<.01&&f.Name==r.Family)??fallback;
        }
        internal static Color AdjustmentColor(Form owner,Color fallback) {FormatRule rule;return owner!=null&&adjustment.TryGetValue(owner,out rule)&&rule.HasColor?Color.FromArgb(rule.Color):fallback;}
    }
    // Layout constraints, owned by the same STA as the form. No host objects are retained.
    internal static class DialogControlSizing
    {
        private sealed class Original {
            internal Size Size,Minimum;internal DockStyle Dock;internal AnchorStyles Anchor;internal bool Auto;
            internal Original(Control c){Size=c.Size;Minimum=c.MinimumSize;Dock=c.Dock;Anchor=c.Anchor;Auto=c.AutoSize;}
            internal void Restore(Control c){c.Dock=Dock;c.Anchor=Anchor;c.AutoSize=Auto;c.MinimumSize=Minimum;c.Size=Size;}
        }
        private sealed class State {
            internal readonly Dictionary<Control,Original> Originals=new Dictionary<Control,Original>();
            internal readonly Dictionary<Control,Size> Rules=new Dictionary<Control,Size>();
            internal readonly List<Action> ParentReset=new List<Action>();
            internal readonly HashSet<RowStyle> ChangedRows=new HashSet<RowStyle>();
            internal bool Busy;internal bool Scroll;internal Size ScrollMin;
        }
        private static readonly ThreadLocal<Dictionary<Form,State>> states=new ThreadLocal<Dictionary<Form,State>>(()=>new Dictionary<Form,State>());
        internal static bool Supports(Control c){return c.Parent is UpDownBase?false:c is TextBoxBase||c is ComboBox||c is UpDownBase||c is ButtonBase||c is Label||c is ListBox||c is ListView||c is TreeView||c is DataGridView||c is GroupBox;}
        private static bool Valid(int value){return value>0&&value<=4096;}
        internal static void Register(Form form,IEnumerable<FormatTarget> targets){
            if(states.Value.ContainsKey(form))return;var state=new State {Scroll=form.AutoScroll,ScrollMin=form.AutoScrollMinSize};states.Value[form]=state;
            foreach(var target in targets.Where(t=>t.ResizeWidth))if(!state.Originals.ContainsKey(target.Control))state.Originals[target.Control]=new Original(target.Control);
            form.Layout+=delegate {Arrange(form);};form.Disposed+=delegate {states.Value.Remove(form);};
        }
        internal static void Reset(Form form){
            State state;if(!states.Value.TryGetValue(form,out state))return;state.Busy=true;
            try{foreach(var reset in state.ParentReset)reset();state.ParentReset.Clear();state.ChangedRows.Clear();foreach(var c in state.Rules.Keys.ToArray())if(!c.IsDisposed)state.Originals[c].Restore(c);state.Rules.Clear();form.AutoScroll=state.Scroll;form.AutoScrollMinSize=state.ScrollMin;}finally{state.Busy=false;}
        }
        internal static bool HasWidth(Control c){State state;return c.FindForm()!=null&&states.Value.TryGetValue(c.FindForm(),out state)&&state.Rules.ContainsKey(c)&&state.Rules[c].Width>0;}
        internal static bool HasHeight(Control c){State state;return c.FindForm()!=null&&states.Value.TryGetValue(c.FindForm(),out state)&&state.Rules.ContainsKey(c)&&state.Rules[c].Height>0;}
        internal static Size Resolve(Control c,Size normal){
            State state;Size requested;var form=c.FindForm();if(form==null||!states.Value.TryGetValue(form,out state)||!state.Rules.TryGetValue(c,out requested))return normal;
            float scale=DialogSizing.Scale(form);return new Size(requested.Width>0?Math.Max(c.MinimumSize.Width,(int)Math.Round(requested.Width*scale)):normal.Width,requested.Height>0?Math.Max(c.MinimumSize.Height,(int)Math.Round(requested.Height*scale)):normal.Height);
        }
        internal static void Apply(Form form,string id,IEnumerable<FormatTarget> targets,CustomizationData data){
            State state=states.Value[form];state.Busy=true;
            try{
                foreach(var target in targets.Where(t=>t.ResizeWidth)){
                    FormatRule rule;if(!data.Formats.TryGetValue(id+"::"+target.Id,out rule))continue;
                    int width=Valid(rule.Width96)?rule.Width96:0,height=target.ResizeHeight&&Valid(rule.Height96)?rule.Height96:0;if(width==0&&height==0)continue;
                    Control c=target.Control;state.Rules[c]=new Size(width,height);
                    // Auto-sized text retains its natural height when only width is pinned.
                    Size normal=c.Size;var label=c as Label;if(label!=null&&height==0){int wrap=Resolve(c,normal).Width;normal.Height=TextRenderer.MeasureText(label.Text,label.Font,new Size(Math.Max(1,wrap),int.MaxValue),TextFormatFlags.WordBreak|TextFormatFlags.NoPrefix).Height+label.Padding.Vertical;}
                    c.AutoSize=false;c.Size=Resolve(c,normal);
                    if(c.Parent is TableLayoutPanel||c.Dock!=DockStyle.None){
                        c.Dock=DockStyle.None;c.Anchor=AnchorStyles.Left|AnchorStyles.Top;
                        if(width==0)c.Anchor|=AnchorStyles.Right;if(height==0&&!(c is ComboBox)&&!(c is UpDownBase))c.Anchor|=AnchorStyles.Bottom;
                    }
                    if(c.Parent is TableLayoutPanel){
                        var table=(TableLayoutPanel)c.Parent;int row=table.GetRow(c);
                        if(height>0&&row>=0&&row<table.RowStyles.Count&&table.GetRowSpan(c)==1){var style=table.RowStyles[row];if(state.ChangedRows.Add(style)){var type=style.SizeType;float old=style.Height;state.ParentReset.Add(()=>{style.SizeType=type;style.Height=old;});}style.SizeType=SizeType.AutoSize;}
                        // Minimum constraints make percent cells participate in measurement.
                        c.MinimumSize=new Size(width>0?c.Width:c.MinimumSize.Width,height>0?c.Height:c.MinimumSize.Height);
                    }
                }
            }finally{state.Busy=false;}
            form.PerformLayout();Arrange(form);
        }
        private static void Arrange(Form form){
            State state;if(!states.Value.TryGetValue(form,out state)||state.Busy||state.Rules.Count==0)return;state.Busy=true;
            try{
                foreach(var pair in state.Rules){var c=pair.Key;if(c.IsDisposed)continue;Size desired=Resolve(c,c.Size);if(c.Size!=desired)c.Size=desired;c.Parent.PerformLayout();}
                int right=form.ClientSize.Width,bottom=form.ClientSize.Height;
                foreach(var pair in state.Rules){Control c=pair.Key;int x=c.Right+c.Margin.Right,y=c.Bottom+c.Margin.Bottom;for(Control p=c.Parent;p!=null&&p!=form;p=p.Parent){x+=p.Left+p.Padding.Right;y+=p.Top+p.Padding.Bottom;}right=Math.Max(right,x);bottom=Math.Max(bottom,y);}
                // Flow/Table containers already reflow neighbours. Only the outer content
                // scrolls when a pinned control cannot fit, rather than every nested panel.
                if(right>form.ClientSize.Width||bottom>form.ClientSize.Height){form.AutoScroll=true;form.AutoScrollMinSize=new Size(right,bottom);}
            }finally{state.Busy=false;}
        }
    }
    internal sealed class PreviewSelectionOverlay:Control
    {
        [StructLayout(LayoutKind.Sequential)] private struct PrintRect {internal int Left,Top,Right,Bottom;}
        [StructLayout(LayoutKind.Sequential)] private struct CharRange {internal int First,Last;}
        [StructLayout(LayoutKind.Sequential)] private struct FormatRange {internal IntPtr Hdc,Target;internal PrintRect Area,Page;internal CharRange Characters;}
        [DllImport("user32.dll",CharSet=CharSet.Auto)] private static extern IntPtr SendMessage(IntPtr handle,int message,IntPtr wParam,IntPtr lParam);
        private readonly Form preview;
        private Bitmap snapshot;
        private readonly Dictionary<Control,Rectangle> rectangles=new Dictionary<Control,Rectangle>();
        internal Action<Control> Pick;
        internal bool ResizeMode,CanResizeWidth,CanResizeHeight;
        internal Action<Size,bool,bool> ResizeSelection;
        internal Action DragStarted,DragCancelled;
        private bool dragging,dragWidth,dragHeight;private Point dragStart;private Size dragSize;
        internal Rectangle ResizeHandle(bool horizontal,bool vertical) {var r=BoundsInPreview(Selected);int n=Ui.Px(this,9);return new Rectangle(horizontal?r.Right-n:r.Left+r.Width/2-n/2,vertical?r.Bottom-n:r.Top+r.Height/2-n/2,n,n);}
        private void Axes(Point point,out bool width,out bool height) {width=height=false;if(!ResizeMode||Selected==null||!CanResizeWidth)return;if(CanResizeHeight&&ResizeHandle(true,true).Contains(point)){width=height=true;return;}width=ResizeHandle(true,false).Contains(point);height=CanResizeHeight&&ResizeHandle(false,true).Contains(point);}
        protected override void OnMouseDown(MouseEventArgs e) {bool x,y;Axes(e.Location,out x,out y);if(e.Button==MouseButtons.Left&&(x||y)){dragWidth=x;dragHeight=y;dragStart=PointToScreen(e.Location);dragSize=Selected.Size;dragging=true;if(DragStarted!=null)DragStarted();Capture=true;return;}base.OnMouseDown(e);var hit=Hit(e.Location);if(hit!=null&&Pick!=null)Pick(hit);}
        protected override void OnMouseMove(MouseEventArgs e) {base.OnMouseMove(e);if(dragging){Point now=PointToScreen(e.Location);float scale;using(var g=Selected.CreateGraphics())scale=g.DpiX/96f;var value=new Size(Math.Max(12,Math.Min(4096,(int)Math.Round((dragSize.Width+now.X-dragStart.X)/scale))),Math.Max(12,Math.Min(4096,(int)Math.Round((dragSize.Height+now.Y-dragStart.Y)/scale))));if(ResizeSelection!=null)ResizeSelection(value,dragWidth,dragHeight);return;}bool x,y;Axes(e.Location,out x,out y);Cursor=x&&y?Cursors.SizeNWSE:x?Cursors.SizeWE:y?Cursors.SizeNS:Cursors.Default;}
        protected override void OnMouseUp(MouseEventArgs e) {dragging=false;Capture=false;base.OnMouseUp(e);}
        internal void CancelDrag() {if(!dragging)return;dragging=false;Capture=false;if(DragCancelled!=null)DragCancelled();}
        protected override void OnMouseCaptureChanged(EventArgs e) {if(dragging&&!Capture)CancelDrag();base.OnMouseCaptureChanged(e);}
        internal Control Selected;
        internal PreviewSelectionOverlay(Form preview) {
            this.preview=preview;TabStop=false;SetStyle(ControlStyles.UserPaint|ControlStyles.AllPaintingInWmPaint|ControlStyles.OptimizedDoubleBuffer,true);BackColor=Ui.Canvas;
            
        }
        internal Control Hit(Point point) {return DialogFormatting.Walk(preview).Where(c=>rectangles.ContainsKey(c)&&rectangles[c].Contains(point)).LastOrDefault();}
        internal Rectangle BoundsInPreview(Control c) {Rectangle bounds;return rectangles.TryGetValue(c,out bounds)?bounds:ManagedBounds(c);}
        private Rectangle ManagedBounds(Control c) {if(c==preview)return new Rectangle(Point.Empty,preview.ClientSize);var result=c.Bounds;for(Control parent=c.Parent;parent!=null&&parent!=preview;parent=parent.Parent)result.Offset(parent.Left,parent.Top);return result;}
        private Rectangle VisibleBounds(Control c) {
            var visible=new Rectangle(Point.Empty,preview.ClientSize);
            for(Control parent=c.Parent;parent!=null&&parent!=preview;parent=parent.Parent) {
                var client=ManagedBounds(parent);client.Size=parent.ClientSize;visible=Rectangle.Intersect(visible,client);
            }
            return Rectangle.Intersect(visible,ManagedBounds(c));
        }
        internal void RefreshSnapshot() {
            // The source stays in its hidden render host for its entire lifetime.
            // Refresh never reparents, shows, focuses or activates a window.
            preview.PerformLayout();var desired=preview.ClientSize;
            var next=new Bitmap(Math.Max(1,desired.Width),Math.Max(1,desired.Height));
            try {
                preview.PerformLayout();
                rectangles.Clear();foreach(var c in DialogFormatting.Walk(preview).Concat(new[]{preview}))if(c.Visible)rectangles[c]=ManagedBounds(c);
                using(var measurement=preview.CreateGraphics())next.SetResolution(measurement.DpiX,measurement.DpiY);
                preview.DrawToBitmap(next,new Rectangle(Point.Empty,next.Size));RenderNativeText(next);
            }catch {next.Dispose();throw;}
            var old=snapshot;snapshot=next;Size=next.Size;if(old!=null)old.Dispose();Invalidate();
        }
        private void RenderNativeText(Bitmap bitmap) {
            foreach(var c in DialogFormatting.Walk(preview)) {
                var rich=c as RichTextBox;var text=c as TextBox;var choice=c as ComboBox;
                if(!rectangles.ContainsKey(c)||(rich==null&&text==null&&choice==null))continue;
                var bounds=BoundsInPreview(c);if(bounds.Width<1||bounds.Height<1)continue;
                var clip=VisibleBounds(c);if(clip.Width<1||clip.Height<1)continue;
                if(rich!=null) {
                    // DrawToBitmap only renders RichTextBox's border. Rich Edit's
                    // documented printing API preserves its actual styled content.
                    using(var part=new Bitmap(bounds.Width,bounds.Height)) {
                        part.SetResolution(bitmap.HorizontalResolution,bitmap.VerticalResolution);
                        using(var g=Graphics.FromImage(part)) {
                            g.Clear(rich.BackColor);var hdc=g.GetHdc();IntPtr memory=IntPtr.Zero;
                            try {var area=new PrintRect {Right=(int)(part.Width*1440/part.HorizontalResolution),Bottom=(int)(part.Height*1440/part.VerticalResolution)};var range=new FormatRange {Hdc=hdc,Target=hdc,Area=area,Page=area,Characters=new CharRange {First=rich.GetCharIndexFromPosition(Point.Empty),Last=-1}};memory=Marshal.AllocHGlobal(Marshal.SizeOf(typeof(FormatRange)));Marshal.StructureToPtr(range,memory,false);SendMessage(rich.Handle,0x0439,new IntPtr(1),memory);}
                            finally {SendMessage(rich.Handle,0x0439,IntPtr.Zero,IntPtr.Zero);if(memory!=IntPtr.Zero)Marshal.FreeHGlobal(memory);g.ReleaseHdc(hdc);}
                        }
                        using(var g=Graphics.FromImage(bitmap)){g.SetClip(clip);g.DrawImageUnscaled(part,bounds.Location);}
                    }
                }else using(var g=Graphics.FromImage(bitmap)) {
                    g.SetClip(clip);
                    // Hidden child TextBoxes are omitted by DrawToBitmap as well.
                    using(var brush=new SolidBrush(c.BackColor))g.FillRectangle(brush,bounds);
                    if(choice!=null||text.BorderStyle!=BorderStyle.None)ControlPaint.DrawBorder(g,bounds,Ui.GridFrame,ButtonBorderStyle.Solid);var inner=bounds;inner.Inflate(-3,-2);
                    string value=c.Text;if(text!=null&&(text.UseSystemPasswordChar||text.PasswordChar!='\0'))value=new string('\u2022',value.Length);
                    var flags=TextFormatFlags.NoPrefix|TextFormatFlags.Left|TextFormatFlags.EndEllipsis;
                    flags|=text!=null&&text.Multiline?TextFormatFlags.WordBreak:TextFormatFlags.SingleLine|TextFormatFlags.VerticalCenter;
                    if(choice!=null){int arrow=Ui.Px(choice,18);inner.Width=Math.Max(1,inner.Width-arrow);ControlPaint.DrawComboButton(g,new Rectangle(bounds.Right-arrow,bounds.Top,arrow,bounds.Height),ButtonState.Inactive);}
                    TextRenderer.DrawText(g,value,c.Font,inner,c.ForeColor,flags);
                }
            }
        }
        protected override void OnPaint(PaintEventArgs e) {base.OnPaint(e);if(snapshot!=null)e.Graphics.DrawImageUnscaled(snapshot,Point.Empty);if(Selected==null||Selected.IsDisposed||!rectangles.ContainsKey(Selected))return;var rect=Rectangle.Intersect(BoundsInPreview(Selected),ClientRectangle);rect.Inflate(-1,-1);if(rect.Width>0&&rect.Height>0)using(var pen=new Pen(Ui.Blue,1))e.Graphics.DrawRectangle(pen,rect);if(ResizeMode&&CanResizeWidth){using(var brush=new SolidBrush(Ui.Blue)){e.Graphics.FillRectangle(brush,ResizeHandle(true,false));if(CanResizeHeight){e.Graphics.FillRectangle(brush,ResizeHandle(false,true));e.Graphics.FillRectangle(brush,ResizeHandle(true,true));}}}}
        protected override void Dispose(bool disposing) {if(disposing&&snapshot!=null){snapshot.Dispose();snapshot=null;}base.Dispose(disposing);}
    }
    internal sealed class DialogFormatEditor:Form
    {
        private readonly CustomizationData draft;
        private readonly LocalCustomization session;
        private readonly ComboBox dialogs=new ComboBox {DropDownStyle=ComboBoxStyle.DropDownList,Dock=DockStyle.Fill,DrawMode=DrawMode.OwnerDrawFixed,Name="dialogChoice"};
        private readonly TreeView list=new TreeView {Dock=DockStyle.Fill,HideSelection=false,ShowNodeToolTips=true,Scrollable=true};
        private readonly ToolTip tooltips=new ToolTip();
        private readonly Dictionary<FormatTarget,TreeNode> targetNodes=new Dictionary<FormatTarget,TreeNode>();
        private readonly TextBox text=new TextBox {Width=240,Name="captionText"};
        private readonly ComboBox family=new ComboBox {Width=240,Name="fontFamily"};
        private readonly NumericUpDown size=new NumericUpDown {Minimum=6,Maximum=36,DecimalPlaces=1,Value=10,Width=80,Name="fontSize"};
        private readonly CheckBox bold=new CheckBox {Text="粗体",AutoSize=true,Name="fontBold"},italic=new CheckBox {Text="斜体",AutoSize=true,Name="fontItalic"};
        private readonly Label hint=new Label {Text="固定文字（数据控件只可改样式）",AutoSize=true,MaximumSize=new Size(285,0),Name="formatHint"};
        private Button colour;
        private Button resizeControl;
        private readonly Label currentDimensions=new Label {AutoSize=true,Name="currentDimensions"};
        private int dragOriginalWidth,dragOriginalHeight;private bool dragHadRule;private string dragKey;
        internal volatile bool HasUnsavedDraft;
        private readonly FlowLayoutPanel dimensionPanel=new FlowLayoutPanel {AutoSize=true,FlowDirection=FlowDirection.TopDown,WrapContents=false,Visible=false,Name="controlDimensions"};
        private readonly Label dimensionHint=new Label {AutoSize=true,MaximumSize=new Size(285,0),Name="dimensionHint"};
        private readonly Panel canvas=new Panel {Dock=DockStyle.Fill,AutoScroll=true};
        private readonly Label previewTitle=new Label {AutoSize=true,ForeColor=Ui.Blue,Name="previewWindowTitle"};
        private Form preview;private PreviewSelectionOverlay overlay;private List<FormatTarget> items;private string id;private bool loading;private Color color=Ui.Ink;
        private Panel renderHost;
        private sealed class Choice {
            internal readonly DialogDefinition Definition;
            internal Choice(DialogDefinition definition){Definition=definition;}
            internal string Path {get {return Definition.Entry.EndsWith(Definition.Name,StringComparison.Ordinal)?Definition.Entry:Definition.Entry+" / "+Definition.Name;}}
            public override string ToString(){return Definition.Name;}
        }
        internal DialogFormatEditor(bool previewOnly) {
            session=new LocalCustomization(RibbonCustomization.Store.Path);draft=session.Data.Copy();Text="自定格式";DialogSizing.Prepare(this,"agent.format","settings",Ui.Blue);
            var root=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Padding=new Padding(12)};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var split=new SplitContainer {Dock=DockStyle.Fill,Width=Math.Max(800,ClientSize.Width-24),SplitterWidth=8,Panel1MinSize=260,Panel2MinSize=250,SplitterDistance=330,Name="formatSplit"};
            var propertyViewport=new Panel {Dock=DockStyle.Fill,AutoScroll=true,Name="propertyViewport"};
            var editor=new TableLayoutPanel {Dock=DockStyle.Top,ColumnCount=1,RowCount=3};editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));editor.RowStyles.Add(new RowStyle(SizeType.Percent,100));editor.RowStyles.Add(new RowStyle(SizeType.AutoSize));editor.Controls.Add(dialogs);editor.Controls.Add(list);
            var props=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true,FlowDirection=FlowDirection.TopDown,WrapContents=false};props.Controls.Add(hint);props.Controls.Add(Ui.Grip(text,(Control)null));props.Controls.Add(Ui.Grip(family,(Control)null));props.Controls.Add(size);
            var styleRow=new FlowLayoutPanel {AutoSize=true,WrapContents=false,Name="styleActions"};styleRow.Controls.Add(bold);styleRow.Controls.Add(italic);
            CommandCustomizationForm.AddButton(styleRow,"文字颜色",()=>{using(var chooser=new ColorDialog {Color=color})if(chooser.ShowDialog(this)==DialogResult.OK){color=chooser.Color;Change();}});colour=(Button)styleRow.Controls[styleRow.Controls.Count-1];props.Controls.Add(styleRow);
            CommandCustomizationForm.AddButton(props,"调整控件缩放",()=>{dimensionPanel.Visible=!dimensionPanel.Visible;SelectTarget();});resizeControl=(Button)props.Controls[props.Controls.Count-1];resizeControl.Name="resizeControl";
            dimensionPanel.Controls.Add(currentDimensions);dimensionPanel.Controls.Add(dimensionHint);
            CommandCustomizationForm.AddButton(dimensionPanel,"恢复默认尺寸",()=>{var t=Selected;if(t==null)return;FormatRule r;if(draft.Formats.TryGetValue(id+"::"+t.Id,out r)){r.Width96=r.Height96=0;}Render();SelectTarget();});props.Controls.Add(dimensionPanel);
            editor.Controls.Add(props);propertyViewport.Controls.Add(editor);split.Panel1.Controls.Add(propertyViewport);split.Panel2.Controls.Add(canvas);root.Controls.Add(split,0,0);
            bool fitting=false;Action fitProperties=delegate {if(fitting||Disposing||IsDisposed)return;fitting=true;try{int needed=styleRow.GetPreferredSize(Size.Empty).Width+Ui.Px(this,20);needed=Math.Max(needed,resizeControl.GetPreferredSize(Size.Empty).Width+Ui.Px(this,20));int limit=Math.Max(260,split.Width-split.Panel2MinSize-split.SplitterWidth);split.Panel1MinSize=Math.Min(needed,limit);if(split.SplitterDistance<split.Panel1MinSize)split.SplitterDistance=split.Panel1MinSize;editor.Width=propertyViewport.ClientSize.Width;int width=Math.Max(120,editor.ClientSize.Width-12);text.Width=DialogControlSizing.Resolve(text,new Size(width,text.Height)).Width;family.Width=DialogControlSizing.Resolve(family,new Size(width,family.Height)).Width;hint.MaximumSize=dimensionHint.MaximumSize=new Size(width,0);editor.Height=Math.Max(propertyViewport.ClientSize.Height,dialogs.PreferredHeight+props.GetPreferredSize(new Size(width,0)).Height+Ui.Px(this,160));propertyViewport.AutoScrollMinSize=new Size(0,editor.Height);}finally{fitting=false;}};
            propertyViewport.Resize+=delegate {fitProperties();};props.SizeChanged+=delegate {fitProperties();};dimensionPanel.VisibleChanged+=delegate {fitProperties();};styleRow.FontChanged+=delegate {fitProperties();};Load+=delegate {fitProperties();};
            var bar=new FlowLayoutPanel {Dock=DockStyle.Fill,AutoSize=true};CommandCustomizationForm.AddButton(bar,"恢复控件",()=>{var t=Selected;if(t!=null)draft.Formats.Remove(id+"::"+t.Id);Render();SelectTarget();});
            CommandCustomizationForm.AddButton(bar,"恢复窗口",()=>{foreach(var key in draft.Formats.Keys.Where(k=>k.StartsWith(id+"::",StringComparison.Ordinal)).ToArray())draft.Formats.Remove(key);Render();SelectTarget();});
            CommandCustomizationForm.AddButton(bar,"恢复全部",()=>{draft.Formats.Clear();Render();SelectTarget();});CommandCustomizationForm.AddButton(bar,"保存",()=>{session.Save(draft);HasUnsavedDraft=false;CustomizationSaving.AfterSave(this);});CommandCustomizationForm.AddButton(bar,"取消",()=>Close());root.Controls.Add(bar,0,1);Controls.Add(root);
            foreach(var d in DialogCatalog.All)dialogs.Items.Add(new Choice(d));foreach(var f in FontFamily.Families)family.Items.Add(f.Name);
            dialogs.DrawItem+=delegate(object sender,DrawItemEventArgs e){e.DrawBackground();if(e.Index<0)return;var choice=dialogs.Items[e.Index] as Choice;string value=choice==null?dialogs.Items[e.Index].ToString():(e.State&DrawItemState.ComboBoxEdit)!=0?choice.ToString():choice.Path;TextRenderer.DrawText(e.Graphics,value,e.Font,e.Bounds,e.ForeColor,TextFormatFlags.Left|TextFormatFlags.VerticalCenter|TextFormatFlags.NoPrefix|TextFormatFlags.EndEllipsis);e.DrawFocusRectangle();};
            dialogs.DropDown+=delegate {int width=dialogs.Width;using(var g=dialogs.CreateGraphics())foreach(var choice in dialogs.Items.OfType<Choice>())width=Math.Max(width,TextRenderer.MeasureText(g,choice.Path,dialogs.Font).Width+Ui.Px(this,30));dialogs.DropDownWidth=Math.Min(width,Math.Max(dialogs.Width,Screen.FromControl(dialogs).WorkingArea.Width-Ui.Px(this,32)));};
            dialogs.SelectedIndexChanged+=delegate {if(!previewOnly)BuildPreview();};list.AfterSelect+=delegate {SelectTarget();};text.TextChanged+=delegate {Change();};family.TextChanged+=delegate {Change();};size.ValueChanged+=delegate {Change();};bold.CheckedChanged+=delegate {Change();};italic.CheckedChanged+=delegate {Change();};
            KeyPreview=true;KeyDown+=delegate(object sender,KeyEventArgs e){if(e.KeyCode==Keys.Escape&&overlay!=null){overlay.CancelDrag();e.Handled=true;}};
            if(!previewOnly)dialogs.SelectedIndex=0;else {dialogs.Items.Clear();dialogs.Items.Add("选择窗口后可安全编辑固定文字与样式");dialogs.SelectedIndex=0;dialogs.Enabled=false;}
            Shown+=delegate {Render();};
            FormClosed+=delegate {DisposePreview();};
        }
        private FormatTarget Selected {get {return list.SelectedNode==null?null:list.SelectedNode.Tag as FormatTarget;}}
        private void BuildControlTree() {
            list.Nodes.Clear();targetNodes.Clear();var controls=new Dictionary<Control,TreeNode>();
            foreach(var target in items){var node=new TreeNode(target.Label){Tag=target,ToolTipText=target.Label};targetNodes[target]=node;TreeNode parent=null;
                if(!controls.TryGetValue(target.Control,out parent)){for(Control c=target.Control.Parent;c!=null&&parent==null;c=c.Parent)controls.TryGetValue(c,out parent);controls[target.Control]=node;}
                if(parent==null)list.Nodes.Add(node);else parent.Nodes.Add(node);
            }
            list.ExpandAll();if(list.Nodes.Count>0)list.SelectedNode=list.Nodes[0];
        }
        private void BuildPreview() {
            if(dialogs.SelectedIndex<0||dialogs.SelectedIndex>=DialogCatalog.All.Length)return;
            DisposePreview();canvas.Controls.Clear();id=DialogCatalog.All[dialogs.SelectedIndex].Id;
            var choice=dialogs.SelectedItem as Choice;tooltips.SetToolTip(dialogs,choice==null?"":choice.Path);
            preview=DialogPreviewFactory.CreateFormatRender(id,DialogCatalog.Find(id).Default);preview.Text=DialogPreviewFactory.OriginalTitle(preview);preview.TopLevel=false;preview.FormBorderStyle=FormBorderStyle.None;preview.StartPosition=FormStartPosition.Manual;preview.Location=Point.Empty;preview.TabStop=false;preview.AcceptButton=null;preview.CancelButton=null;preview.ShowInTaskbar=false;
            foreach(var control in DialogFormatting.Walk(preview)){control.TabStop=false;if(control is ButtonBase)control.Enabled=false;}
            renderHost=new Panel {Size=preview.Size,TabStop=false};renderHost.Controls.Add(preview);preview.Show();preview.Location=Point.Empty;
            items=DialogFormatting.Register(preview,id);BuildControlTree();
            canvas.AutoScrollPosition=Point.Empty;canvas.Controls.Add(previewTitle);previewTitle.Location=Point.Empty;overlay=new PreviewSelectionOverlay(preview);canvas.Controls.Add(overlay);
            overlay.DragStarted=delegate {dragKey=id+"::"+Selected.Id;FormatRule rule;dragHadRule=draft.Formats.TryGetValue(dragKey,out rule);dragOriginalWidth=rule==null?0:rule.Width96;dragOriginalHeight=rule==null?0:rule.Height96;};
            overlay.DragCancelled=delegate {FormatRule rule;if(!dragHadRule)draft.Formats.Remove(dragKey);else if(draft.Formats.TryGetValue(dragKey,out rule)){rule.Width96=dragOriginalWidth;rule.Height96=dragOriginalHeight;}Render();SelectTarget();};
            overlay.ResizeSelection=ChangeDimensions;
            overlay.Pick=delegate(Control c){var t=items.FirstOrDefault(x=>x.Control==c);if(t!=null)list.SelectedNode=targetNodes[t];};Render();SelectTarget();
        }
        private void Render() {HasUnsavedDraft=!LocalCustomization.Encode(draft).SequenceEqual(LocalCustomization.Encode(session.Data));if(preview!=null&&overlay!=null){DialogFormatting.Apply(preview,id,draft);renderHost.Size=preview.Size;previewTitle.Text="窗口："+preview.Text;overlay.RefreshSnapshot();overlay.Location=new Point(canvas.AutoScrollPosition.X,canvas.AutoScrollPosition.Y+previewTitle.PreferredHeight+8);canvas.AutoScrollMinSize=new Size(overlay.Width,overlay.Height+previewTitle.PreferredHeight+8);}}
        private void DisposePreview() {if(overlay!=null){overlay.CancelDrag();overlay.Dispose();overlay=null;}if(preview!=null){preview.Dispose();preview=null;}if(renderHost!=null){renderHost.Dispose();renderHost=null;}}
        protected override void Dispose(bool disposing) {if(disposing){DisposePreview();tooltips.Dispose();}base.Dispose(disposing);}
        private void SelectTarget() {
            var t=Selected;if(t==null)return;if(overlay!=null&&overlay.Selected!=t.Control)overlay.CancelDrag();loading=true;try {FormatRule r;draft.Formats.TryGetValue(id+"::"+t.Id,out r);text.Enabled=t.Caption;text.Text=r!=null&&r.Text!=null?r.Text:t.OriginalText;family.Text=r!=null&&!string.IsNullOrEmpty(r.Family)?r.Family:t.OriginalFont.Name;size.Value=(decimal)Math.Max(6,Math.Min(36,r!=null&&r.Size>0?r.Size:t.OriginalFont.Size));bold.Checked=r!=null?(r.Style&1)!=0:t.OriginalFont.Bold;italic.Checked=r!=null?(r.Style&2)!=0:t.OriginalFont.Italic;color=r!=null&&r.HasColor?Color.FromArgb(r.Color):t.OriginalColor;
                resizeControl.Enabled=t.ResizeWidth;currentDimensions.Text="当前尺寸："+(int)Math.Round(t.Control.Width/DialogSizing.Scale(preview))+" × "+(int)Math.Round(t.Control.Height/DialogSizing.Scale(preview))+"（96 DPI 逻辑像素）";
                dimensionHint.Text=!t.ResizeWidth?"该项目不单独调整尺寸；整体窗口请使用‘交互弹窗’。":(t.ResizeHeight?"拖动右侧、底部或右下角蓝色方块调整宽高；Esc 撤销本次拖动。":"拖动右侧蓝色方块调整宽度；高度由系统和字号管理。");
                if(overlay!=null){overlay.Selected=t.Control;overlay.ResizeMode=dimensionPanel.Visible;overlay.CanResizeWidth=t.ResizeWidth;overlay.CanResizeHeight=t.ResizeHeight;overlay.Invalidate();}}finally {loading=false;}
            bool native=t.Id=="window";family.Enabled=size.Enabled=bold.Enabled=italic.Enabled=colour.Enabled=!native;hint.Text=native?"窗口标题可改文字；标题栏字体和颜色由系统管理。":"固定文字（数据控件只可改样式）";
        }
        private FormatRule DraftRule(FormatTarget target){FormatRule rule;if(!draft.Formats.TryGetValue(id+"::"+target.Id,out rule)){rule=new FormatRule {Family=target.OriginalFont.Name,Size=target.OriginalFont.Size,Style=(target.OriginalFont.Bold?1:0)|(target.OriginalFont.Italic?2:0)};draft.Formats[id+"::"+target.Id]=rule;}return rule;}
        private void Change() {var t=Selected;if(loading||t==null)return;var rule=DraftRule(t);rule.Text=t.Caption?text.Text:null;rule.Family=family.Text;rule.Size=(float)size.Value;rule.Style=(bold.Checked?1:0)|(italic.Checked?2:0);rule.Color=color.ToArgb();rule.HasColor=true;Render();}
        private void ChangeDimensions(Size dimensions,bool width,bool height){var t=Selected;if(loading||t==null||!t.ResizeWidth)return;var rule=DraftRule(t);if(width)rule.Width96=dimensions.Width;if(height&&t.ResizeHeight)rule.Height96=dimensions.Height;Render();SelectTarget();}
    }
}
