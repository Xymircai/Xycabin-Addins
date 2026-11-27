using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Reused task selection UI; workbook reading remains in CompanyWorkbookHost.
    internal sealed class CompanyStructurePicker : UserControl
    {
        internal Func<string, CompanyWorkbookHost> OpenSource;
        internal event Action Changed;
        internal Action<string> Log;
        internal bool UseLevel = true;
        private CompanyWorkbookHost source;
        private readonly TextBox file = new TextBox { ReadOnly = true, Dock = DockStyle.Fill };
        private readonly ComboBox sheet = new ComboBox { Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox sequence = Column("A"), shortName = Column("C"), fullName = Column("E"), fileName = Column("G"), level = Column("D");
        private readonly NumericUpDown first = new NumericUpDown { Minimum = 1, Maximum = 1048576, Value = 3, Width = 75 };
        internal readonly DataGridView Tasks = new ResultGrid { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
        private readonly Label count = new Label { AutoSize = true };
        private bool loading;
        private string preferredSheet;
        private Control mappingBody;
        private bool simpleMode;
        private bool selectParents;
        private RoundedButton defaultSelection;
        internal bool SelectParents { get { return selectParents; } set {selectParents=value;if(defaultSelection!=null)defaultSelection.Text=value?"选择合并主体":"选择末级公司";if(Tasks.Rows.Count>0)SetSelection(0);} }
        internal bool SimpleMode { get { return simpleMode; } set { simpleMode=value;if(mappingBody!=null)mappingBody.Visible=!value;InvalidateRows(); } }
        internal static readonly string[] ConfigurationKeys = { "structureSheet", "structureFirst", "structureSequence", "structureShort", "structureFull", "structureFile", "structureLevel" };
        private static TextBox Column(string text) { return new TextBox { Text = text, Width = 40 }; }
        internal static List<CompanyStructureTask> Choose(Form owner,Func<string,CompanyWorkbookHost> open,bool useLevel)
        {
            var choice=ChooseSelection(owner,open,useLevel);return choice==null?null:choice.Selected;
        }
        internal static CompanyStructureSelection ChooseSelection(Form owner,Func<string,CompanyWorkbookHost> open,bool useLevel)
        {
            using(var dialog=new Form())using(var picker=new CompanyStructurePicker{Dock=DockStyle.Fill,OpenSource=open,UseLevel=useLevel,SimpleMode=true})
            {
                dialog.Text="选择架构任务";DialogSizing.Prepare(dialog,"summary.single","table",Ui.Blue);
                var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Padding=new Padding(16)};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var confirm=Ui.Button("使用选中任务",140);var cancel=Ui.Button("取消",90);var actions=Ui.ActionBar();actions.Controls.AddRange(new Control[]{confirm,cancel});root.Controls.Add(picker,0,0);root.Controls.Add(actions,0,1);dialog.Controls.Add(root);
                picker.Log=delegate(string text){UiDialog.Show(dialog,text,"架构任务",MessageBoxButtons.OK,MessageBoxIcon.Information);};
                confirm.Click+=delegate{if(picker.Selected().Count==0){picker.Log("请至少勾选一个任务。");return;}dialog.DialogResult=DialogResult.OK;};cancel.Click+=delegate{dialog.DialogResult=DialogResult.Cancel;};
                return dialog.ShowDialog(owner)==DialogResult.OK?new CompanyStructureSelection{All=picker.All(),Selected=picker.Selected()}:null;
            }
        }
        internal CompanyStructurePicker()
        {
            Height = 310; Dock = DockStyle.Top; MinimumSize = new Size(550, 260);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Margin = Padding.Empty };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var path = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            path.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100)); path.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));path.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var browse = Ui.Button("选择架构表", 130);var read = Ui.Button("重新读取",110);path.Controls.Add(file,0,0);path.Controls.Add(browse,1,0);path.Controls.Add(read,2,0);root.Controls.Add(path,0,0);
            var mapping = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true };
            foreach(Control control in new Control[] { new Label{Text="工作表",AutoSize=true}, sheet, new Label{Text="首行",AutoSize=true}, first,
                new Label{Text="序号列",AutoSize=true},sequence,new Label{Text="简称列",AutoSize=true},shortName,new Label{Text="全称列",AutoSize=true},fullName,new Label{Text="文件名列",AutoSize=true},fileName,new Label{Text="级次列",AutoSize=true},level }) mapping.Controls.Add(control);
            mappingBody=mapping;root.Controls.Add(mapping,0,1);root.Controls.Add(Tasks,0,2);
            Tasks.Columns.Add(new DataGridViewCheckBoxColumn{Name="selected",HeaderText="生成 / 处理",Width=90});
            Tasks.Columns.Add(new DataGridViewTextBoxColumn{Name="sequence",HeaderText="序号",Width=90,ReadOnly=true});
            Tasks.Columns.Add(new DataGridViewTextBoxColumn{Name="short",HeaderText="简称",Width=130,ReadOnly=true});
            Tasks.Columns.Add(new DataGridViewTextBoxColumn{Name="company",HeaderText="公司全称",Width=200,ReadOnly=true});
            Tasks.Columns.Add(new DataGridViewTextBoxColumn{Name="file",HeaderText="过程表文件名称",AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill,ReadOnly=true});Ui.StyleGrid(Tasks);
            Tasks.Columns.Add(new DataGridViewTextBoxColumn{Name="status",HeaderText="预检状态",Width=180,ReadOnly=true});
            var actions = new FlowLayoutPanel { Dock=DockStyle.Top,AutoSize=true,WrapContents=false };
            var leaves=Ui.Button("选择末级公司",130);defaultSelection=leaves;var all=Ui.Button("全选",80);var clear=Ui.Button("清空",80);actions.Controls.AddRange(new Control[]{leaves,all,clear,count});root.Controls.Add(actions,0,3);Controls.Add(root);
            browse.Click+=delegate{Attempt(delegate{string[] paths=HostDialogService.OpenFiles(FindForm(),"架构表 (*.xlsx;*.xls)|*.xlsx;*.xls");if(paths.Length==1)LoadSource(paths[0]);});};read.Click+=delegate{Attempt(delegate{if(string.IsNullOrEmpty(file.Text))throw new InvalidOperationException("请先选择架构表。");preferredSheet=sheet.Text;LoadSource(file.Text);});};
            leaves.Click+=delegate{SetSelection(0);};all.Click+=delegate{SetSelection(1);};clear.Click+=delegate{SetSelection(2);};
            foreach(var column in new[]{sequence,shortName,fullName,fileName,level})column.TextChanged+=delegate{InvalidateRows();};first.ValueChanged+=delegate{InvalidateRows();};sheet.SelectedIndexChanged+=delegate{InvalidateRows();};
            Tasks.CurrentCellDirtyStateChanged+=delegate{if(Tasks.IsCurrentCellDirty)Tasks.CommitEdit(DataGridViewDataErrorContexts.Commit);};Tasks.CellValueChanged+=delegate{if(!loading)Notify();};
            Disposed+=delegate{if(source!=null)source.Dispose();};
        }
        private void Attempt(Action action){try{action();}catch(Exception error){if(Log!=null)Log("架构任务："+error.Message);}}
        internal void LoadSource(string path)
        {
            if(OpenSource==null)throw new InvalidOperationException("架构表读取服务尚未连接。");
            InvalidateRows();
            var candidate=OpenSource(path);var previous=source;source=candidate;
            file.Text=path;sheet.Items.Clear();sheet.Items.AddRange(candidate.Sheets.ToArray());if(sheet.Items.Count>0)sheet.SelectedIndex=preferredSheet!=null&&sheet.Items.Contains(preferredSheet)?sheet.Items.IndexOf(preferredSheet):0;
            if(previous!=null)previous.Dispose();ReadTasks();
        }
        private static int ColumnNumber(string text)
        {
            string value=(text??"").Trim().ToUpperInvariant();if(!System.Text.RegularExpressions.Regex.IsMatch(value,@"^[A-Z]{1,3}$"))throw new InvalidOperationException("列请填写字母，例如A、B、F。");
            int n=0;foreach(char c in value)n=n*26+c-'A'+1;if(n>16384)throw new InvalidOperationException("列超出Excel范围。");return n;
        }
        internal Dictionary<string,string> Configuration()
        {
            if(SimpleMode)return new Dictionary<string,string>{{"structureSheet",sheet.Text},{"structureFirst","3"},{"structureSequence","A"},{"structureShort","C"},{"structureFull","E"},{"structureFile","G"},{"structureLevel","D"}};
            return new Dictionary<string,string>{{"structureSheet",sheet.Text},{"structureFirst",first.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)},{"structureSequence",sequence.Text},{"structureShort",shortName.Text},{"structureFull",fullName.Text},{"structureFile",fileName.Text},{"structureLevel",level.Text}};
        }
        internal static void ValidateConfiguration(IDictionary<string,string> values)
        {
            if(values.Count==0)return;
            if(values.Count!=ConfigurationKeys.Length||ConfigurationKeys.Any(k=>!values.ContainsKey(k)))throw new InvalidOperationException("架构表列配置不完整。");
            int start;if(!int.TryParse(values["structureFirst"],out start)||start<1||start>1048576)throw new InvalidOperationException("架构表起始行无效。");
            foreach(string key in new[]{"structureSequence","structureShort","structureFull","structureFile","structureLevel"})ColumnNumber(values[key]);
        }
        internal void ApplyConfiguration(IDictionary<string,string> values)
        {
            ValidateConfiguration(values);if(values.Count==0)return;
            if(SimpleMode){preferredSheet=values["structureSheet"];InvalidateRows();return;}
            loading=true;try{preferredSheet=values["structureSheet"];first.Value=int.Parse(values["structureFirst"]);sequence.Text=values["structureSequence"];shortName.Text=values["structureShort"];fullName.Text=values["structureFull"];fileName.Text=values["structureFile"];level.Text=values["structureLevel"];if(sheet.Items.Contains(preferredSheet))sheet.SelectedItem=preferredSheet;}finally{loading=false;}InvalidateRows();
        }
        private void ReadTasks()
        {
            InvalidateRows();if(source==null||sheet.SelectedIndex<0)throw new InvalidOperationException("请先选择架构表和工作表。");
            if(SimpleMode)
            {
                var headers=source.Headers(sheet.Text,1);
                foreach(var expected in new Dictionary<string,int>{{"序号",1},{"简称",3},{"级次",4},{"公司名称",5},{"过程表文件名称",7}})
                {int column;if(!headers.TryGetValue(expected.Key,out column)||column!=expected.Value)throw new InvalidOperationException("架构表格式不匹配：请使用A列序号、C列简称、D列级次、E列公司名称、G列过程表文件名称的版本。");}
            }
            var rows=source.StructureRows(sheet.Text,SimpleMode?new CompanyStructureColumns():new CompanyStructureColumns{FirstRow=(int)first.Value,Sequence=ColumnNumber(sequence.Text),ShortName=ColumnNumber(shortName.Text),FullName=ColumnNumber(fullName.Text),FileName=ColumnNumber(fileName.Text),Level=ColumnNumber(level.Text)});
            CompanyStructureTasks.Classify(rows,UseLevel);SetTasks(rows);
        }
        internal void SetTasks(IList<CompanyStructureTask> rows)
        {
            loading=true;try{Tasks.Rows.Clear();foreach(var task in rows){int row=Tasks.Rows.Add(SelectParents?!task.Leaf:task.Leaf,task.Sequence,task.ShortName,task.FullName,task.FileName,"待预检");Tasks.Rows[row].Tag=task;}}finally{loading=false;}Notify();
        }
        internal void Report(int sourceRow,string status)
        {
            loading=true;try{foreach(DataGridViewRow row in Tasks.Rows)if(((CompanyStructureTask)row.Tag).Row==sourceRow){row.Cells["status"].Value=status;row.Cells["status"].ToolTipText=status;break;}}finally{loading=false;}
        }
        internal List<CompanyStructureTask> Selected()
        {
            Tasks.EndEdit();return Tasks.Rows.Cast<DataGridViewRow>().Where(r=>object.Equals(r.Cells[0].Value,true)).Select(r=>Copy((CompanyStructureTask)r.Tag)).ToList();
        }
        internal List<CompanyStructureTask> All(){return Tasks.Rows.Cast<DataGridViewRow>().Select(r=>Copy((CompanyStructureTask)r.Tag)).ToList();}
        private static CompanyStructureTask Copy(CompanyStructureTask t){return new CompanyStructureTask{Row=t.Row,Sequence=t.Sequence,ShortName=t.ShortName,FullName=t.FullName,FileName=t.FileName,Leaf=t.Leaf,Level=t.Level,ParentFullName=t.ParentFullName};}
        private void SetSelection(int mode){loading=true;try{foreach(DataGridViewRow row in Tasks.Rows)row.Cells[0].Value=mode==1||(mode==0&&(SelectParents?!((CompanyStructureTask)row.Tag).Leaf:((CompanyStructureTask)row.Tag).Leaf));}finally{loading=false;}Notify();}
        private void InvalidateRows(){if(loading)return;Tasks.Rows.Clear();Notify();}
        private void Notify(){count.Text="共"+Tasks.Rows.Count+"项，已选"+Selected().Count+"项";if(Changed!=null)Changed();}
    }
}
