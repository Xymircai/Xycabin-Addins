using System;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // UI only: no Excel objects cross this boundary. Uses the migration UI kit.
    internal sealed class SummaryAreaView : UserControl
    {
        internal event Action<string, bool> Pick;
        internal event Action<string> GroupRemoved;
        internal readonly List<string> ExtraGroups=new List<string>();
        private readonly Dictionary<string,Control> extraRows=new Dictionary<string,Control>();
        private TableLayoutPanel extraBody;
        private int nextGroup;
        internal void AddGroup(string type,string restored=null)
        {
            string key=restored;if(key==null){do{key=type+"_"+(++nextGroup);}while(extraRows.ContainsKey(key));}if(extraRows.ContainsKey(key))return;
            if(ExtraGroups.Count>=20)throw new InvalidOperationException("最多添加20组，请先删除不需要的配置。");
            var row=new TableLayoutPanel{Dock=DockStyle.Top,AutoSize=true,ColumnCount=1};
            row.Controls.Add(Picker(type=="repeat"?"变量来源":"期间来源",key+"Source"));row.Controls.Add(Picker(type=="repeat"?"变量写入位置":"期间写入起点",key+"Target"));
            var remove=Ui.Button("删除这一组",120);row.Controls.Add(remove);inputs.Add(remove);
            remove.Click+=delegate{RemoveGroup(key);};extraRows.Add(key,row);ExtraGroups.Add(key);extraBody.Controls.Add(row);VariableHeaders.Checked=true;InvalidatePlan();
        }
        internal void RemoveGroup(string key)
        {
            Control row;if(!extraRows.TryGetValue(key,out row))return;
            Func<Control,bool> within=null;within=delegate(Control c){return c==row||(c.Parent!=null&&within(c.Parent));};inputs.RemoveAll(c=>within(c));
            addresses.Remove(key+"Source");addresses.Remove(key+"Target");extraRows.Remove(key);ExtraGroups.Remove(key);extraBody.Controls.Remove(row);row.Dispose();if(GroupRemoved!=null)GroupRemoved(key);InvalidatePlan();
        }
        internal void ClearGroups(){foreach(string key in ExtraGroups.ToArray())RemoveGroup(key);}

        internal event Action PreviewRequested, ExecuteRequested, UndoRequested, ConfirmRequested, Changed;
        internal event Action BatchFilesRequested;
        internal Action<string> LogSink;
        private readonly List<Control> inputs = new List<Control>();
        private readonly Dictionary<string, TextBox> addresses = new Dictionary<string, TextBox>();
        internal readonly TextBox Folder = new TextBox { Dock = DockStyle.Fill, Name = "summaryProjectFolder", ReadOnly = false, ShortcutsEnabled = true, AccessibleDescription = "可直接输入或粘贴项目文件夹路径，也可在 Excel 中选择存放路径文字的单元格。" };
        internal readonly ResultGrid ReferenceSources = new ResultGrid { AutoSize=false,Height=145,Dock=DockStyle.Top,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false };
        private Control referenceBody;
        internal readonly CheckBox VariableHeaders = new CheckBox {Text="Lead 变量与期间关联",AutoSize=true,Checked=true};
        internal readonly CheckBox FormulaHeaders = new CheckBox {Text="使用公式关联 Lead（单列年度自动转置）",AutoSize=true,Checked=true};
        internal readonly CheckBox RepeatNames = new CheckBox {Text="每个年度列都填写公司名称",AutoSize=true,Checked=true};
        internal readonly CheckBox ReplaceHeaderArrays = new CheckBox {Text="允许整体重写既有表头数组（预检列出完整范围）",AutoSize=true};
        internal readonly CheckBox HeadersEnabled = new CheckBox { AutoSize = true, Checked=true };
        internal readonly Label HeaderPreview = new Label { AutoSize = true, ForeColor = Ui.InkDim, MaximumSize = new Size(1000,0) };
        internal string ProjectFolder
        {
            get { string value = Folder.Text.Trim(); return value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"' ? value.Substring(1, value.Length - 2) : value; }
        }
        internal readonly CheckBox CheckFiles = new CheckBox { Text = "检查来源文件是否存在（不打开文件）", AutoSize = true, Checked = false };
        internal readonly ResultGrid Companies = new ResultGrid { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White, BorderStyle = BorderStyle.None };
        private readonly RoundedButton preview = Ui.Button("读取 Lead 并预检", 150), execute = Ui.Button("执行", 90), undo = Ui.Button("撤销", 90), confirm = Ui.Button("确认", 90);
        private readonly Label state = new Label { Dock = DockStyle.Top, AutoSize = true, ForeColor = Ui.InkDim };
        private readonly TableLayoutPanel parameters;
        private BatchPresetActions presets;
        internal SummaryAreaView()
        {
            Dock = DockStyle.Fill;
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16), BackColor = Ui.Canvas };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 60)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 40)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            parameters = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 1 };
            parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            parameters.Controls.Add(Picker("模板区域（可 Ctrl 多选）", "template"));
            parameters.Controls.Add(Picker("粘贴起点（一个单元格）", "paste"));
            parameters.Controls.Add(Picker("Lead 公司名称（单列，不含表头）", "lead"));
            var folder = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3 };
            folder.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); folder.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); folder.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            folder.Controls.Add(new Label { Text = "项目文件夹（可输入）", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0); folder.Controls.Add(Folder, 1, 0);
            var browse = Ui.Button("在 Excel 中选择…", 150); browse.Name = "summaryPickFolderCell"; folder.Controls.Add(browse, 2, 0); inputs.Add(browse); inputs.Add(Folder);
            browse.Click += delegate { if (Pick != null) Pick("folder", true); };
            Folder.TextChanged += delegate { InvalidatePlan(); }; CheckFiles.CheckedChanged += delegate { InvalidatePlan(); };
            parameters.Controls.Add(folder); parameters.Controls.Add(CheckFiles); inputs.Add(CheckFiles);
            ReferenceSources.Columns.Add(new DataGridViewTextBoxColumn{Name="source",HeaderText="来源文件 / 工作表",ReadOnly=true,Width=280});
            ReferenceSources.Columns.Add(new DataGridViewTextBoxColumn{Name="positions",HeaderText="涉及位置",ReadOnly=true,Width=130});
            ReferenceSources.Columns.Add(new DataGridViewTextBoxColumn{Name="reason",HeaderText="识别说明",ReadOnly=true,AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill});
            var decision=new DataGridViewComboBoxColumn{Name="decision",HeaderText="处理方式",Width=145};decision.Items.AddRange("待确认","参与公司替换","保留原引用");ReferenceSources.Columns.Add(decision);Ui.StyleGrid(ReferenceSources);
            referenceBody=ReferenceSources;referenceBody.Visible=false;parameters.Controls.Add(referenceBody);
            ReferenceSources.CurrentCellDirtyStateChanged+=delegate {if(ReferenceSources.IsCurrentCellDirty)ReferenceSources.CommitEdit(DataGridViewDataErrorContexts.Commit);};
            ReferenceSources.CellValueChanged+=delegate {InvalidatePlan();};
            var headerBody = new TableLayoutPanel { AutoSize=true,ColumnCount=1,Dock=DockStyle.Top,Visible=true };
            headerBody.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            var expand=Ui.Button("▼ 表头自动填充",180);parameters.Controls.Add(expand);
            expand.Click+=delegate { headerBody.Visible=!headerBody.Visible;expand.Text=(headerBody.Visible?"▼":"▶")+" 表头自动填充"; };
            inputs.Add(HeadersEnabled);
            inputs.Add(VariableHeaders);VariableHeaders.CheckedChanged+=delegate{InvalidatePlan();};
            headerBody.Controls.Add(Picker("重复变量来源","repeatSource"));
            headerBody.Controls.Add(Picker("变量写入位置","repeatTarget"));
            foreach(var option in new[]{FormulaHeaders,RepeatNames,ReplaceHeaderArrays}){inputs.Add(option);option.CheckedChanged+=delegate{InvalidatePlan();};}

            var legacyCompanyPicker=Picker("公司名称写入位置", "companyHeader");headerBody.Controls.Add(legacyCompanyPicker);legacyCompanyPicker.Visible=false;
            headerBody.Controls.Add(Picker("期间来源区域", "yearSource"));
            headerBody.Controls.Add(Picker("期间写入起点", "yearTarget"));parameters.Controls.Add(headerBody);
            extraBody=new TableLayoutPanel{AutoSize=true,Dock=DockStyle.Top,ColumnCount=1};headerBody.Controls.Add(extraBody);
            var addVariables=Ui.Button("＋ 添加变量",140);var addPeriods=Ui.Button("＋ 添加期间",140);var addBar=Ui.ActionBar();addBar.Controls.Add(addVariables);addBar.Controls.Add(addPeriods);headerBody.Controls.Add(addBar);inputs.Add(addVariables);inputs.Add(addPeriods);
            addVariables.Click+=delegate{if(ExtraGroups.Count<20)AddGroup("repeat");else Message("最多添加20组，请先删除不需要的组。");};addPeriods.Click+=delegate{if(ExtraGroups.Count<20)AddGroup("year");else Message("最多添加20组，请先删除不需要的组。");};
            HeadersEnabled.CheckedChanged+=delegate { InvalidatePlan(); };

            var card = new UiCard { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(12) }; card.Controls.Add(parameters);
            var parameterScroll=new Panel { Dock=DockStyle.Fill,AutoScroll=true,BackColor=Ui.Canvas };parameterScroll.Controls.Add(card);root.Controls.Add(parameterScroll,0,0);
            root.RowStyles[0].SizeType=SizeType.Absolute;
            bool fittingParameters=false;Action fitParameters=delegate { if(fittingParameters)return;fittingParameters=true;try{float height=Math.Min(card.GetPreferredSize(new Size(Math.Max(1,parameterScroll.ClientSize.Width-20),0)).Height+6,Math.Max(180,(int)(ClientSize.Height*.55)));if(Math.Abs(root.RowStyles[0].Height-height)>1)root.RowStyles[0].Height=height;}finally{fittingParameters=false;} };
            extraBody.SizeChanged+=delegate{fitParameters();};SizeChanged+=delegate{fitParameters();}; headerBody.VisibleChanged+=delegate{fitParameters();}; referenceBody.VisibleChanged+=delegate{fitParameters();}; HeaderPreview.TextChanged+=delegate{fitParameters();};
            state.Text = "选区绑定同一工作簿；只复制选中的 Area。未选空隙不变；全程不自动保存。";
            state.Text = "";
            Companies.Columns.Add(new DataGridViewCheckBoxColumn { Name = "selected", HeaderText = "选用", Width = 70 });
            Companies.Columns.Add(new DataGridViewTextBoxColumn { Name = "row", HeaderText = "Lead 行", ReadOnly = true, Width = 95 });
            Companies.Columns.Add(new DataGridViewTextBoxColumn { Name = "code", HeaderText = "公司编码及简称", Visible=false, ReadOnly = true, Width = 230 });
            Companies.Columns.Add(new DataGridViewTextBoxColumn { Name = "company", HeaderText = "公司名称 / 文件名", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            Ui.StyleGrid(Companies); root.Controls.Add(Companies, 0, 2); root.Controls.Add(log, 0, 3);
            Companies.CurrentCellDirtyStateChanged += delegate { if (Companies.IsCurrentCellDirty) Companies.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            Companies.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e) { if (e.ColumnIndex == 0) InvalidatePlan(); };
            var batchFiles = Ui.Button("多文件批量链接", 165); inputs.Add(batchFiles);
            batchFiles.Click += delegate { if (BatchFilesRequested != null) BatchFilesRequested(); };
            var bar = Ui.ActionBar(); bar.Controls.AddRange(new Control[] { preview, execute, confirm, undo, batchFiles }); root.Controls.Add(bar, 0, 4);
            Ui.StylePrimary(preview); preview.Click += delegate { if (PreviewRequested != null) PreviewRequested(); };
            execute.Click += delegate { if (ExecuteRequested != null) ExecuteRequested(); }; undo.Click += delegate { if (UndoRequested != null) UndoRequested(); }; confirm.Click += delegate { if (ConfirmRequested != null) ConfirmRequested(); };
            Controls.Add(root); SetState(false, false, false);
        }
        private Control Picker(string title, string kind)
        {
            var row = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 3, Margin = new Padding(0, 0, 0, 6) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.Controls.Add(new Label { Text = title, AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            var text = new TextBox { ReadOnly = true, Dock = DockStyle.Fill, Name = "summaryArea_" + kind }; addresses.Add(kind, text); row.Controls.Add(text, 1, 0);
            var pick = Ui.Button("在 Excel 中选择…", 150); row.Controls.Add(pick, 2, 0);
            pick.Click += delegate { if (Pick != null) Pick(kind, true); };
            inputs.Add(pick); return row;
        }
        internal Dictionary<string,string> ReviewSources(IList<SummaryReferenceSource> sources)
        {
            var prior=new Dictionary<string,string[]>(StringComparer.OrdinalIgnoreCase);
            foreach(DataGridViewRow row in ReferenceSources.Rows) prior[Convert.ToString(row.Cells["source"].Value)]=new[]{Convert.ToString(row.Tag),Convert.ToString(row.Cells["decision"].Value)};
            ReferenceSources.Rows.Clear();var result=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            foreach(var source in sources)
            {
                string[] saved;string decision=prior.TryGetValue(source.Key,out saved)&&saved[0]==source.Signature?saved[1]:source.Decision;
                int index=ReferenceSources.Rows.Add(source.Key,source.Locations,source.Reason,decision);ReferenceSources.Rows[index].Tag=source.Signature;result[source.Key]=decision;
            }
            referenceBody.Visible=sources.Count>0;return result;
        }
        internal void Address(string kind, string value) { addresses[kind].Text = value; InvalidatePlan(); }
        internal void AttachPresets(BatchPresetActions actions)
        {
            presets = actions; parameters.SuspendLayout();
            var children = new List<Control>(); foreach (Control child in parameters.Controls) children.Add(child);
            parameters.RowCount++;
            foreach (var child in children) parameters.SetRow(child, parameters.GetRow(child) + 1);
            parameters.Controls.Add(actions.Header(new Label { Text = "", AutoSize = true }), 0, 0); parameters.ResumeLayout(true);
        }
        internal void InvalidatePlan() { execute.Enabled = false; if (Changed != null) Changed(); }
        internal void Message(string value) { if(LogSink!=null){LogSink(value);return;} log.AppendText(DateTime.Now.ToString("HH:mm:ss") + "  " + value + Environment.NewLine); log.SelectionStart = log.TextLength; log.ScrollToCaret(); if (LogSink != null) LogSink(value); }
        internal void SetState(bool running, bool pending, bool ready)
        { foreach (var c in inputs) c.Enabled = !running && !pending; if (presets != null) presets.Gear.Enabled = !running && !pending; Companies.Enabled = !running && !pending; ReferenceSources.Enabled = !running && !pending; preview.Enabled = !running && !pending; execute.Enabled = !running && !pending && ready; undo.Enabled = confirm.Enabled = !running && pending; }
    }
}
