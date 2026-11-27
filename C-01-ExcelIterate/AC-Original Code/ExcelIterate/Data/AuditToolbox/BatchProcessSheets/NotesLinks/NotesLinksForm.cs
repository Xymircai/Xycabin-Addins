using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class NotesLinksForm : Form
    {
        internal object Application;
        internal Func<bool> CanRun;
        internal event Action<bool> BusyChanged;
        private readonly TextBox parent=Box("",480),folder=Box("",480),parentCode=Box("K2",65),codeColumn=Box("K",45),nameColumn=Box("M",45),fileColumn=Box("N",45),sourceRange=Box("D1:X1700",155),firstStart=Box("AV1",75),firstColumn=Box("D",45),lastColumn=Box("X",45);
        private readonly ComboBox leadSheet=Choice("Lead"),sourceSheet=Choice("Notes"),targetSheet=Choice("Notes");
        private readonly NumericUpDown firstRow=Number(3,1),lastRow=Number(0,0),sourceFirstRow=Number(1,1,1048576),sourceRows=Number(1700,1),columnSpacing=Number(22,1,16384);
        private readonly CheckBox nativePaste=new CheckBox{Text="优先整块粘贴链接（WPS更快；合并布局一致时使用）",AutoSize=true,Checked=true};
        private readonly Label placementHint=new Label{AutoSize=true};
        private readonly FlowLayoutPanel settings=new FlowLayoutPanel{Dock=DockStyle.Fill,FlowDirection=FlowDirection.TopDown,WrapContents=false,AutoScroll=true};
        private readonly DataGridView grid=new ResultGrid{Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false};
        private readonly RichTextBox log=new RichTextBox{Dock=DockStyle.Fill,ReadOnly=true,BackColor=Color.White};
        private readonly ProgressBar progress=new ProgressBar{Dock=DockStyle.Fill,Style=ProgressBarStyle.Continuous};
        private readonly Label state=new Label{AutoSize=true,Text="选择父公司，读取名单，再预检。结束行0表示自动。"};
        private readonly RoundedButton read=Ui.Button("读取公司",115),preview=Ui.Button("预检查",100),execute=Ui.Button("备份后更新",130),cancel=Ui.Button("停止",90),undo=Ui.Button("还原本次更新",135);
        private readonly Timer timer=new Timer{Interval=25};
        private List<NotesLinkTask> tasks=new List<NotesLinkTask>();
        private readonly CheckBox batchMode=new CheckBox{Text="批量处理多个汇总文件",AutoSize=true};
        private readonly CheckBox batchConfirmed=new CheckBox{Text="已核对所有明细表的特殊格式（仅在预检要求时勾选）",AutoSize=true};
        private readonly DataGridView batchGrid=new ResultGrid{Dock=DockStyle.Fill,AllowUserToAddRows=false,AllowUserToDeleteRows=false,RowHeadersVisible=false,Visible=false};
        private readonly List<NotesLinkBatchItem> batchItems=new List<NotesLinkBatchItem>();
        private NotesLinkBatchRun batchRun;
        private string batchPreparedKey;
        private NotesLinkPlan plan;
        private NotesLinkExecution execution;
        private IEnumerator<string> sequence;
        private string leadKey,leadHash,lastParent,lastBackup,lastHash,sheetsPath;
        private bool loading,busy,ticking,closePending;
        private volatile bool stop;
        private static TextBox Box(string value,int width){return new TextBox{Text=value,Width=width};}
        private static ComboBox Choice(string value){return new ComboBox{Text=value,Width=160,DropDownStyle=ComboBoxStyle.DropDown};}
        private static NumericUpDown Number(int value,int minimum,int maximum=100000){return new NumericUpDown{Minimum=minimum,Maximum=maximum,Value=value,Width=82};}
        private static Label Label(string text){return new Label{Text=text,AutoSize=true,Anchor=AnchorStyles.Left,Margin=new Padding(4,7,5,3)};}
        internal NotesLinksForm()
        {
            Text="Notes 批量链接";DialogSizing.Prepare(this,"notes.links","table",Ui.Blue);
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,ColumnCount=1,RowCount=6,Padding=new Padding(14)};root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute,400));root.RowStyles.Add(new RowStyle(SizeType.Percent,60));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));root.RowStyles.Add(new RowStyle(SizeType.Absolute,18));root.RowStyles.Add(new RowStyle(SizeType.Percent,40));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var chooseParent=Ui.Button("选择汇总文件",125);chooseParent.Click+=delegate{Attempt(delegate{var paths=HostDialogService.OpenFiles(this,"Excel工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");if(paths.Length==0)return;if(paths.Length>1||batchMode.Checked){batchMode.Checked=true;AddBatchFiles(paths);}else{parent.Text=paths[0];if(folder.Text.Length==0)folder.Text=Path.GetDirectoryName(paths[0]);RefreshSheets();}});};
            var addFolder=Ui.Button("从文件夹添加",130);addFolder.Click+=delegate{Attempt(delegate{string path=HostDialogService.Folder(this,"选择汇总文件所在的文件夹（不含子文件夹）",true);if(string.IsNullOrEmpty(path))return;batchMode.Checked=true;AddBatchFiles(Directory.GetFiles(path));});};
            var removeFiles=Ui.Button("移除所选行",115);removeFiles.Enabled=false;removeFiles.Click+=delegate{if(busy)return;var item=SelectedBatchItem();if(item==null)return;batchItems.Remove(item);Changed();RenderBatch();};
            Row("处理方式",batchMode,addFolder,removeFiles);
            var chooseFolder=Ui.Button("选择文件夹",125);chooseFolder.Click+=delegate{string path=HostDialogService.Folder(this,"子公司工作簿所在目录",true);if(!string.IsNullOrEmpty(path))folder.Text=path;};
            Row("汇总文件",parent,chooseParent);Row("明细文件夹",folder,chooseFolder);
            Row("工作表",Label("从明细文件取"),sourceSheet,SheetPick(sourceSheet,SampleFile),Label("放到汇总文件"),targetSheet,SheetPick(targetSheet,ParentFileForPick));
            Row("每家取数范围",sourceRange,Pick("明细取数范围",SampleFile,SampleSheet,delegate(MigrationArea a){sourceRange.Text=a.Address;SyncRange();},()=>sourceRange.Text,"在明细文件中选"));
            Row("第一家放在",firstStart,Pick("汇总表放置起点",ParentFileForPick,delegate{return targetSheet.Text;},delegate(MigrationArea a){TargetAnchor(a);firstStart.Text=CellAddress.Name(a.Row,a.Column);},()=>firstStart.Text,"在汇总文件中选"));
            Row("列间距",columnSpacing,Label("相邻两家公司的起始列相差多少列。例：AV到BR填22，不是中间留空的列数。"));
            Row("放置预览",placementHint);
            var advanced=Ui.Button("高级设置 ▼",150);Row("",advanced);int advancedStart=settings.Controls.Count;
            Row("粘贴方式",nativePaste);nativePaste.CheckedChanged+=delegate{Changed();};
            Row("公司名单表",leadSheet);
            Row("名单信息位置",Label("汇总公司编码在"),parentCode,Pick("汇总公司编码",ParentFileForPick,delegate{return leadSheet.Text;},delegate(MigrationArea a){Single(a);parentCode.Text=CellAddress.Name(a.Row,a.Column);},()=>parentCode.Text,"选汇总公司编码"),Label("公司编码列"),codeColumn,ColumnPick(codeColumn),Label("公司名称列"),nameColumn,ColumnPick(nameColumn),Label("完整文件名列"),fileColumn,ColumnPick(fileColumn));
            Row("名单从哪行读",Label("起始"),firstRow,Label("结束（0自动）"),lastRow,Pick("名单行",ParentFileForPick,delegate{return leadSheet.Text;},delegate(MigrationArea a){firstRow.Value=a.Row;lastRow.Value=a.Row+a.Rows-1;},()=>codeColumn.Text+firstRow.Value+":"+codeColumn.Text+(lastRow.Value==0?firstRow.Value:lastRow.Value),"选名单行"),Label("默认N列为带编码的完整文件名，不使用O列的无编码文件名。"));
            Row("每家取多少行",sourceRows,Label("与上方取数范围联动，例如D1:X1700共1700行。"));
            Row("批量特殊表格",batchConfirmed);
            var advancedRows=settings.Controls.Cast<Control>().Skip(advancedStart).ToArray();foreach(var item in advancedRows)item.Visible=false;advanced.Click+=delegate{bool expand=advanced.Text.EndsWith("▼");foreach(var item in advancedRows)item.Visible=expand;advanced.Text=expand?"高级设置 ▲":"高级设置 ▼";};
            Row("规则",Label("公司按名单顺序从左向右排放。取消一家后保留它的空位，后面的公司不前移。表名或合并格与预期不同时，请先打开文件核对，再勾选“已核对特殊表格”。保留合并格、边框和行列大小，只粘贴链接，保留原数字格式；不更新外链、不重新计算。"));
            grid.Columns.Add(new DataGridViewCheckBoxColumn{Name="selected",HeaderText="执行",Width=48});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="code",HeaderText="编码",Width=85,ReadOnly=true});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="company",HeaderText="公司",Width=160,ReadOnly=true});
            grid.Columns.Add(new DataGridViewTextBoxColumn{Name="file",HeaderText="来源文件（可修改）",Width=220});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="sheet",HeaderText="来源工作表",Width=140});grid.Columns.Add(new DataGridViewCheckBoxColumn{Name="confirmed",HeaderText="已核对特殊表格",Width=75});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="target",HeaderText="汇总文件放置范围",Width=160,ReadOnly=true});grid.Columns.Add(new DataGridViewTextBoxColumn{Name="status",HeaderText="状态",AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill,ReadOnly=true});Ui.StyleGrid(grid);
            var chooseSource=Ui.Button("指定所选来源文件",170);chooseSource.Click+=delegate{Attempt(ChangeFile);};var all=Ui.Button("全选",85);var clear=Ui.Button("取消全选",100);all.Click+=delegate{SelectAll(true);};clear.Click+=delegate{SelectAll(false);};
            var actions=Ui.ActionBar();var close=Ui.Button("关闭",90);actions.Controls.AddRange(new Control[]{read,chooseSource,all,clear,preview,execute,cancel,undo,close});
            batchGrid.Columns.Add(new DataGridViewCheckBoxColumn{Name="batchSelected",HeaderText="处理",Width=50});
            batchGrid.Columns.Add(new DataGridViewTextBoxColumn{Name="batchPath",HeaderText="汇总文件完整路径",Width=340,ReadOnly=true});
            batchGrid.Columns.Add(new DataGridViewTextBoxColumn{Name="batchCount",HeaderText="公司数量",Width=85,ReadOnly=true});
            batchGrid.Columns.Add(new DataGridViewTextBoxColumn{Name="batchStatus",HeaderText="预检及处理结果",Width=260,ReadOnly=true});
            batchGrid.Columns.Add(new DataGridViewTextBoxColumn{Name="batchBackup",HeaderText="恢复备份位置",Width=300,ReadOnly=true});Ui.StyleGrid(batchGrid);
            var data=new Panel{Dock=DockStyle.Fill};data.Controls.Add(grid);data.Controls.Add(batchGrid);
            batchGrid.CurrentCellDirtyStateChanged+=delegate{if(batchGrid.IsCurrentCellDirty)batchGrid.CommitEdit(DataGridViewDataErrorContexts.Commit);};
            batchGrid.CellValueChanged+=delegate(object sender,DataGridViewCellEventArgs e){if(loading||busy||e.RowIndex<0||e.ColumnIndex!=0)return;var item=(NotesLinkBatchItem)batchGrid.Rows[e.RowIndex].Tag;item.Selected=object.Equals(batchGrid.Rows[e.RowIndex].Cells[0].Value,true);Changed();};
            batchGrid.SelectionChanged+=delegate{if(loading||busy)return;var item=SelectedBatchItem();if(item!=null){loading=true;parent.Text=item.Parent;loading=false;}undo.Enabled=batchMode.Checked&&item!=null&&!string.IsNullOrEmpty(item.Backup);};
            batchMode.CheckedChanged+=delegate{if(busy)return;parent.ReadOnly=batchMode.Checked;removeFiles.Enabled=batchMode.Checked;Text=batchMode.Checked?"Notes 批量链接 · 多文件":"Notes 批量链接";grid.Visible=!batchMode.Checked;batchGrid.Visible=batchMode.Checked;chooseSource.Enabled=!batchMode.Checked;read.Text=batchMode.Checked?"读取并预检":"读取公司";execute.Text=batchMode.Checked?"批量粘贴链接":"备份后更新";undo.Text=batchMode.Checked?"还原所选文件":"还原本次更新";Changed();state.Text=batchMode.Checked?"所有汇总文件共用下方设置，各自读取公司名单；先勾选文件并统一预检。":"选择汇总文件，读取公司名单，再预检。";if(batchMode.Checked&&batchItems.Count==0&&File.Exists(parent.Text))AddBatchFiles(new[]{parent.Text});};
            batchConfirmed.CheckedChanged+=delegate{Changed();};
            root.Controls.Add(settings,0,0);root.Controls.Add(data,0,1);root.Controls.Add(state,0,2);root.Controls.Add(progress,0,3);root.Controls.Add(log,0,4);root.Controls.Add(actions,0,5);Controls.Add(root);
            parent.Leave+=delegate{if(!busy&&!batchMode.Checked&&File.Exists(parent.Text)&&parent.Text!=sheetsPath)Attempt(RefreshSheets);};
            read.Click+=delegate{if(batchMode.Checked)Preview();else Attempt(ReadCompanies);};preview.Click+=delegate{Preview();};execute.Click+=delegate{Attempt(Begin);};cancel.Click+=delegate{stop=true;if(execution!=null)execution.Cancelled=true;if(batchRun!=null)batchRun.Cancelled=true;Append("已请求停止；提交前取消将丢弃临时修改。");};undo.Click+=delegate{Attempt(Restore);};close.Click+=delegate{Close();};execute.Enabled=cancel.Enabled=undo.Enabled=false;
            foreach(Control control in new Control[]{parent,folder,leadSheet,sourceSheet,targetSheet,parentCode,codeColumn,nameColumn,fileColumn,firstRow,lastRow,sourceRange,firstStart,columnSpacing})
            {if(control is NumericUpDown)((NumericUpDown)control).ValueChanged+=delegate{Changed();};else control.TextChanged+=delegate{Changed();};}
            sourceSheet.TextChanged+=delegate{if(!loading){foreach(var task in tasks){task.Sheet="";task.NameNeedsConfirmation=false;}Render();}};
            sourceRange.TextChanged+=delegate{if(!loading)try{SyncRange();}catch(InvalidOperationException){};};sourceRange.Leave+=delegate{Attempt(SyncRange);};firstColumn.TextChanged+=delegate{Parameters();};lastColumn.TextChanged+=delegate{Parameters();};sourceFirstRow.ValueChanged+=delegate{Parameters();};sourceRows.ValueChanged+=delegate{Parameters();};
            grid.CurrentCellDirtyStateChanged+=delegate{if(grid.IsCurrentCellDirty)grid.CommitEdit(DataGridViewDataErrorContexts.Commit);};grid.CellValueChanged+=Edited;grid.DataError+=delegate(object sender,DataGridViewDataErrorEventArgs e){e.ThrowException=false;};
            timer.Tick+=delegate{Tick();};FormClosing+=delegate(object sender,FormClosingEventArgs e){if(busy){e.Cancel=true;closePending=true;stop=true;if(execution!=null)execution.Cancelled=true;if(batchRun!=null)batchRun.Cancelled=true;}};
            var help=new ToolTip{AutoPopDelay=20000,InitialDelay=300};
            help.SetToolTip(parent,"接收所有公司数据的Excel文件，更新前会自动备份。");
            help.SetToolTip(folder,"各公司明细Excel文件所在的文件夹，文件名要与名单一致。");
            help.SetToolTip(sourceSheet,"每家公司文件中提供数据的工作表名称，通常是Notes。");
            help.SetToolTip(targetSheet,"汇总文件中接收数据的工作表名称。");
            help.SetToolTip(sourceRange,"D1:X1700表示从D1到X1700，共21列、1700行。所有公司使用同一范围。");
            help.SetToolTip(firstStart,"第一家公司在汇总表中的左上角单元格，例如AV1。");
            help.SetToolTip(columnSpacing,"第二家起始列号减去第一家起始列号。每家占21列时，填21紧挨排放，填22中间留一列。");
            help.SetToolTip(leadSheet,"汇总文件中记录公司名单的工作表，通常是Lead。");
            help.SetToolTip(parentCode,"填单元格地址，不是编码本身。例如K2存放03-08-，程序据此找出直属公司。");
            help.SetToolTip(codeColumn,"公司编码在哪一列，填列字母，例如K。");
            help.SetToolTip(nameColumn,"公司全称在哪一列，例如M。");
            help.SetToolTip(fileColumn,"包含编码和.xlsx后缀的完整文件名在哪一列，例如N；用来找到明细文件。");
            help.SetToolTip(firstRow,"从名单第几行开始读取，不包含标题行。");
            help.SetToolTip(lastRow,"读到第几行结束；填0自动找末行。");
            Disposed+=delegate{help.Dispose();};
            PlacementPreview();
            Disposed+=delegate{timer.Stop();timer.Dispose();if(sequence!=null)sequence.Dispose();if(execution!=null)execution.Dispose();};
        }
        private void Row(string name,params Control[] fields)
        {var line=new FlowLayoutPanel{AutoSize=true,WrapContents=true,Margin=new Padding(0,1,0,1),Width=1120};line.Controls.Add(new Label{Text=name,AutoSize=false,Width=90,Height=28,TextAlign=ContentAlignment.MiddleLeft});line.Controls.AddRange(fields);settings.Controls.Add(line);}
        private void Append(string text){if(IsDisposed)return;log.AppendText(DateTime.Now.ToString("HH:mm:ss")+"  "+text+Environment.NewLine);log.SelectionStart=log.TextLength;log.ScrollToCaret();log.Update();}
        private void Attempt(Action action){try{if(busy||CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");action();}catch(OperationCanceledException error){Append(error.Message);}catch(Exception error){Append("失败："+error.Message);UiDialog.Show(this,error.Message,"Notes 批量链接",MessageBoxButtons.OK,MessageBoxIcon.Warning);}}
        private void Changed(){if(loading)return;PlacementPreview();batchPreparedKey=null;plan=null;execute.Enabled=false;state.Text="参数或名单已改变，请重新预检。";}
        private void PlacementPreview()
        {
            try{var area=MigrationArea.Parse(sourceRange.Text);var starts=Enumerable.Range(0,3).Select(i=>NotesLinkRules.Target(area,firstStart.Text,(int)columnSpacing.Value,i)).Select((a,i)=>"第"+(i+1)+"家："+CellAddress.Name(a.Row,a.Column));placementHint.Text=string.Join("；",starts)+"（每家占"+area.Columns+"列，中间空"+((int)columnSpacing.Value-area.Columns)+"列）";}
            catch(Exception e){placementHint.Text=e.Message;}
        }
        private void SetBusy(bool value)
        {busy=value;settings.Enabled=grid.Enabled=batchGrid.Enabled=read.Enabled=preview.Enabled=!value;cancel.Enabled=value;execute.Enabled=!value&&(batchMode.Checked?batchPreparedKey==BatchKey()&&batchItems.Any(i=>i.Selected&&i.Ready):plan!=null&&plan.Tasks.All(t=>t.Status.StartsWith("通过",StringComparison.Ordinal)));var item=SelectedBatchItem();undo.Enabled=!value&&(batchMode.Checked?item!=null&&!string.IsNullOrEmpty(item.Backup):lastBackup!=null);if(BusyChanged!=null)BusyChanged(value);}
        private NotesLinkConfig Config()
        {return new NotesLinkConfig{Parent=parent.Text.Trim().Trim('"'),Folder=folder.Text.Trim().Trim('"'),LeadSheet=leadSheet.Text,ParentCodeCell=parentCode.Text.Trim(),CodeColumn=codeColumn.Text.Trim(),NameColumn=nameColumn.Text.Trim(),FileColumn=fileColumn.Text.Trim(),FirstRow=(int)firstRow.Value,LastRow=(int)lastRow.Value,SourceSheet=sourceSheet.Text,TargetSheet=targetSheet.Text,SourceRange=sourceRange.Text.Trim(),FirstStart=firstStart.Text.Trim(),ColumnSpacing=(int)columnSpacing.Value,NativePaste=nativePaste.Checked};}
        private void RefreshSheets()
        {
            var names=DataMigrationHost.PackageSheets(parent.Text);loading=true;
            try{SetChoices(leadSheet,names,leadSheet.Text);SetChoices(targetSheet,names,targetSheet.Text);sheetsPath=parent.Text;}finally{loading=false;}Changed();
        }
        private static void SetChoices(ComboBox box,List<string> names,string preferred)
        {box.Items.Clear();box.Items.AddRange(names.Cast<object>().ToArray());var matches=NotesLinkRules.Candidates(names,preferred);box.Text=matches.Count==1&&NotesLinkRules.ExactOrAlias(matches[0],preferred)?matches[0]:preferred;}
        private void ReadCompanies()
        {
            var cfg=Config();NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],cfg.Parent,false);tasks=NotesLinkRules.ReadLead(cfg);leadKey=cfg.LeadKey;leadHash=DataMigrationHost.Digest(cfg.Parent);Changed();
            var sample=tasks.FirstOrDefault(t=>File.Exists(t.File));if(sample!=null){loading=true;try{SetChoices(sourceSheet,DataMigrationHost.PackageSheets(sample.File),sourceSheet.Text);}finally{loading=false;}}
            Render();Append("已读取直系公司 "+tasks.Count+" 家，包含名单中的本部，按Lead顺序固定槽位；父编码来源："+cfg.LeadSheet+"!"+cfg.ParentCodeCell);
        }
        private void Render()
        {
            loading=true;
            try
            {
                grid.Rows.Clear();var cfg=Config();
                foreach(var task in tasks)
                {
                    string target="请检查第一家位置和列间距";try{target=NotesLinkRules.Target(MigrationArea.Parse(cfg.SourceRange),cfg.FirstStart,cfg.ColumnSpacing,task.Slot).Address;}catch{}
                    int n=grid.Rows.Add(task.Selected,task.Code,task.Name,task.File,task.Sheet,task.Confirmed,target,task.Status);grid.Rows[n].Tag=task;
                    var cell=new DataGridViewComboBoxCell{DisplayStyle=DataGridViewComboBoxDisplayStyle.DropDownButton,FlatStyle=FlatStyle.Flat};
                    foreach(string name in task.SheetNames.Concat(new[]{task.Sheet.Length>0?task.Sheet:cfg.SourceSheet}).Distinct())cell.Items.Add(name);
                    cell.Value=task.Sheet.Length>0?task.Sheet:cfg.SourceSheet;grid.Rows[n].Cells["sheet"]=cell;
                }
            }
            finally{loading=false;}
        }
        private void Edited(object sender,DataGridViewCellEventArgs e)
        {
            if(loading||e.RowIndex<0||busy)return;var row=grid.Rows[e.RowIndex];var task=row.Tag as NotesLinkTask;if(task==null)return;
            string field=grid.Columns[e.ColumnIndex].Name;
            if(field=="selected")task.Selected=object.Equals(row.Cells[field].Value,true);
            if(field=="confirmed")task.Confirmed=object.Equals(row.Cells[field].Value,true);
            if(field=="file"){task.File=Convert.ToString(row.Cells[field].Value).Trim().Trim('"');task.Sheet="";task.NameNeedsConfirmation=false;task.Confirmed=false;}
            if(field=="sheet"){task.Sheet=Convert.ToString(row.Cells[field].Value);task.NameNeedsConfirmation=false;}
            Changed();
        }
        private void SelectAll(bool value){if(busy)return;if(batchMode.Checked){foreach(var item in batchItems)item.Selected=value;Changed();RenderBatch();return;}foreach(var task in tasks)task.Selected=value;Changed();Render();}
        private NotesLinkTask CurrentTask(){return grid.CurrentRow==null?tasks.FirstOrDefault(t=>t.Selected):(NotesLinkTask)grid.CurrentRow.Tag;}
        private string SampleFile()
        {
            if(tasks.Count==0||leadKey!=Config().LeadKey)ReadCompanies();
            var task=tasks.FirstOrDefault(t=>File.Exists(t.File));if(task!=null)return task.File;
            var files=HostDialogService.OpenFiles(this,"请选择来源工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");
            if(files.Length==1)return files[0];throw new OperationCanceledException("未选择来源工作簿，原配置保留。");
        }
        private RoundedButton SheetPick(ComboBox field,Func<string> file)
        {
            var button=Ui.Button("选择工作表",130);button.Click+=delegate{Attempt(delegate{
                string path=file();var names=DataMigrationHost.PackageSheets(path);
                using(var dialog=new Form{Text="选择工作表 · "+Path.GetFileName(path),Width=340,Height=350,StartPosition=FormStartPosition.CenterParent})
                {var list=new ListBox{Dock=DockStyle.Fill};list.Items.AddRange(names.Cast<object>().ToArray());var ok=new Button{Text="确定",Dock=DockStyle.Bottom,DialogResult=DialogResult.OK};dialog.Controls.Add(list);dialog.Controls.Add(ok);dialog.AcceptButton=ok;
                 if(dialog.ShowDialog(this)==DialogResult.OK&&list.SelectedItem!=null)field.Text=Convert.ToString(list.SelectedItem);}
            });};return button;
        }
        private string SampleSheet(){return sourceSheet.Text;}
        private void ChangeFile(){var task=CurrentTask();if(task==null)throw new InvalidOperationException("请先选择一家公司。");var files=HostDialogService.OpenFiles(this,"Excel工作簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm");if(files.Length==1){task.File=files[0];task.Sheet="";task.Confirmed=false;task.NameNeedsConfirmation=false;task.SheetNames=DataMigrationHost.PackageSheets(files[0]);Changed();Render();}}
        private async void Preview()
        {
            if(busy)return;
            if(batchMode.Checked){await PreviewBatch();return;}
            try
            {
                if(CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");grid.EndEdit();var cfg=Config();if(leadKey!=cfg.LeadKey)throw new InvalidOperationException("父公司、目录或Lead配置已变化，请先重新读取公司。");
                if(DataMigrationHost.Digest(cfg.Parent)!=leadHash)throw new InvalidOperationException("父公司文件在读取名单后已变化，请重新读取公司。");
                var rows=tasks.Select(t=>t.Copy()).ToList();NotesLinkExecution.CheckOpenWorkbooks(Application,rows.Where(t=>t.Selected&&File.Exists(t.File)).Select(t=>t.File),cfg.Parent,false);
                stop=false;plan=null;SetBusy(true);progress.Maximum=Math.Max(1,rows.Count(t=>t.Selected));progress.Value=0;
                var candidate=await Task.Run(()=>NotesLinkPlanner.Build(cfg,rows,delegate(string message){if(!IsDisposed)BeginInvoke((Action)delegate{Append(message);progress.Value=Math.Min(progress.Maximum,progress.Value+1);});},delegate{return stop;}));
                foreach(var task in candidate.Tasks){int index=tasks.FindIndex(t=>t.LeadRow==task.LeadRow);if(index>=0)tasks[index]=task;}
                plan=candidate;Render();int blocked=plan.Tasks.Count(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal));
                state.Text="公司 "+plan.Tasks.Count+" 家；来源 "+candidate.Source.Address+"；覆盖已有内容 "+candidate.Occupied+" 格；阻断 "+blocked+" 项。";
                Append("预检只读，未修改工作簿。合并格内部 "+plan.Tasks.Sum(t=>t.Skipped)+" 格已检查，原生方式由WPS处理，兼容方式跳过；公司位置不重新排列。");
            }
            catch(Exception error){plan=null;Append(error is OperationCanceledException?error.Message:"预检失败："+error.Message);}
            finally{SetBusy(false);if(closePending)Close();}
        }
        private void Begin()
        {
            if(batchMode.Checked){BeginBatch();return;}
            if(plan==null||plan.Config.Key!=Config().Key||plan.Tasks.Any(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal)))throw new InvalidOperationException("请重新预检，确保所有勾选任务通过。");
            NotesLinkExecution.CheckOpenWorkbooks(Application,plan.Tasks.Select(t=>t.File),plan.Config.Parent,true);
            string text="将为 "+plan.Tasks.Count+" 家公司生成真实外链，覆盖父公司“"+plan.TargetSheet+"”中 "+plan.Occupied+" 个已有内容单元格。\n\n执行前自动备份，链接写入并保存后更新汇总文件。保留原格式及合并格；本次不主动更新外链、不重新计算，也不重开检查公式或数值。显示结果可能需下次在Excel中更新链接后刷新。\n\n是否执行？";
            if(UiDialog.Show(this,text,"Notes 批量链接",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)return;
            stop=false;execution=new NotesLinkExecution(Application,plan,Append);sequence=execution.Run().GetEnumerator();progress.Maximum=plan.Tasks.Count;progress.Value=0;SetBusy(true);timer.Start();
        }
        private void Tick()
        {
            if(batchRun!=null){TickBatch();return;}
            if(ticking||sequence==null)return;ticking=true;timer.Stop();
            try
            {
                if(sequence.MoveNext()){state.Text=sequence.Current;Append(sequence.Current);progress.Value=Math.Min(progress.Maximum,execution.Completed);return;}
                lastParent=plan.Config.Parent;lastBackup=execution.Backup;lastHash=execution.CommittedHash;Append("链接已粘贴并保存。备份："+lastBackup);Finish(null);
            }
            catch(Exception error){Finish(error);}
            finally{ticking=false;if(sequence!=null)timer.Start();}
        }
        private void Finish(Exception error)
        {
            bool committed=execution!=null&&execution.Committed;
            if(committed){lastParent=plan.Config.Parent;lastBackup=execution.Backup;lastHash=execution.CommittedHash;}
            timer.Stop();if(sequence!=null){sequence.Dispose();sequence=null;}if(execution!=null){execution.Dispose();execution=null;}
            if(error!=null){foreach(var task in plan.Tasks)task.Status=committed?"已更新／后处理警告":error is OperationCanceledException?"已取消／未交付":"中止／请检查恢复记录";Append((committed?"父公司已更新，后处理警告：":"本批中止，详细恢复路径见日志：")+error.Message);}
            Render();plan=null;SetBusy(false);if(closePending)Close();
        }
        private void Restore()
        {
            if(batchMode.Checked){RestoreBatch();return;}
            if(string.IsNullOrEmpty(lastBackup))throw new InvalidOperationException("本窗口没有可还原的更新。");NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],lastParent,true);
            if(UiDialog.Show(this,"将用备份还原父公司；当前版本也会保留。如果更新后又被修改，将阻止自动还原。是否继续？","还原Notes更新",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;
            string displaced=MigrationOutputTransaction.Restore(lastParent,lastBackup,lastHash);Append("已还原；还原前版本保留："+displaced);lastBackup=null;undo.Enabled=false;Changed();
        }
        private NotesLinkBatchItem SelectedBatchItem(){return batchGrid.CurrentRow==null?null:batchGrid.CurrentRow.Tag as NotesLinkBatchItem;}
        private string BatchKey(){var cfg=Config();cfg.Parent="";return cfg.Key+"|"+batchConfirmed.Checked+"|"+string.Join("\n",batchItems.Where(i=>i.Selected).Select(i=>i.Parent));}
        private void AddBatchFiles(IEnumerable<string> paths)
        {
            foreach(string path in NotesLinkBatch.Files(paths))if(!batchItems.Any(i=>string.Equals(i.Parent,path,StringComparison.OrdinalIgnoreCase)))batchItems.Add(new NotesLinkBatchItem{Parent=path});
            if(batchItems.Count>0){loading=true;parent.Text=batchItems[0].Parent;if(folder.Text.Length==0)folder.Text=Path.GetDirectoryName(parent.Text);loading=false;}
            Changed();RenderBatch();state.Text="已添加 "+batchItems.Count+" 个汇总文件；请勾选需要处理的文件，再统一预检。";
        }
        private void RenderBatch()
        {
            string selected=SelectedBatchItem()==null?null:SelectedBatchItem().Parent;loading=true;
            try{batchGrid.Rows.Clear();foreach(var item in batchItems){int n=batchGrid.Rows.Add(item.Selected,item.Parent,item.Plan==null?"—":item.Plan.Tasks.Count.ToString(),item.Status,item.Backup??"");batchGrid.Rows[n].Tag=item;if(item.Parent==selected)batchGrid.CurrentCell=batchGrid.Rows[n].Cells[1];}}
            finally{loading=false;}
            var current=SelectedBatchItem();if(batchMode.Checked&&current!=null){loading=true;parent.Text=current.Parent;loading=false;}
        }
        private void RefreshBatchRows()
        {foreach(DataGridViewRow row in batchGrid.Rows){var item=(NotesLinkBatchItem)row.Tag;row.Cells[2].Value=item.Plan==null?"—":item.Plan.Tasks.Count.ToString();row.Cells[3].Value=item.Status;row.Cells[4].Value=item.Backup??"";}}
        private async Task PreviewBatch()
        {
            try
            {
                if(CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");batchGrid.EndEdit();
                var selected=batchItems.Where(i=>i.Selected).ToList();if(selected.Count==0)throw new InvalidOperationException("请先添加并勾选汇总文件。");
                var cfg=Config();bool confirmed=batchConfirmed.Checked;string key=BatchKey();var blockers=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
                foreach(var item in selected)try{NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],item.Parent,false);}catch(Exception e){blockers[item.Parent]=e.Message;}
                stop=false;batchPreparedKey=null;SetBusy(true);progress.Maximum=selected.Count;progress.Value=0;
                var prepared=await Task.Run(()=>NotesLinkBatch.Prepare(cfg,selected,confirmed,message=>{if(!IsDisposed)BeginInvoke((Action)delegate{Append(message);});},()=>stop,blockers));
                foreach(var item in prepared)
                {
                    if(item.Ready)try{NotesLinkExecution.CheckOpenWorkbooks(Application,item.Plan.Tasks.Select(t=>t.File),item.Parent,false);}catch(Exception e){item.Plan=null;item.Status="阻止："+e.Message;}
                    batchItems[batchItems.FindIndex(i=>string.Equals(i.Parent,item.Parent,StringComparison.OrdinalIgnoreCase))]=item;
                }
                NotesLinkBatch.Order(prepared);batchPreparedKey=key;progress.Value=progress.Maximum;
                int ready=prepared.Count(i=>i.Ready);state.Text="预检完成：通过 "+ready+" 个，阻止 "+(prepared.Count-ready)+" 个；未修改任何文件。";Append(state.Text);
            }
            catch(Exception e){batchPreparedKey=null;Append(e is OperationCanceledException?e.Message:"批量预检失败："+e.Message);}
            finally{RenderBatch();SetBusy(false);if(closePending)Close();}
        }
        private void BeginBatch()
        {
            if(batchPreparedKey!=BatchKey())throw new InvalidOperationException("参数或勾选文件已变化，请重新批量预检。");
            var ready=NotesLinkBatch.Order(batchItems);if(ready.Count==0)throw new InvalidOperationException("没有通过预检的汇总文件。");
            int blocked=batchItems.Count(i=>i.Selected&&!i.Ready);long occupied=ready.Sum(i=>(long)i.Plan.Occupied);
            if(UiDialog.Show(this,"将处理通过预检的 "+ready.Count+" 个汇总文件（共 "+ready.Sum(i=>i.Plan.Tasks.Count)+" 家公司），覆盖 "+occupied+" 个已有内容单元格。未通过的 "+blocked+" 个文件将跳过。\n\n每个文件先备份，再只粘贴链接并保存；不更新外链、不重算。单个文件失败后继续，停止时保留已完成文件，当前未完成文件不覆盖。是否继续？","Notes 多文件批量链接",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)return;
            stop=false;batchRun=new NotesLinkBatchRun(Application,ready,Append);sequence=batchRun.Run().GetEnumerator();progress.Maximum=ready.Count;progress.Value=0;SetBusy(true);timer.Start();
        }
        private void TickBatch()
        {
            if(ticking||sequence==null)return;ticking=true;timer.Stop();
            try
            {
                batchRun.Cancelled=stop;if(sequence.MoveNext()){state.Text=sequence.Current;Append(sequence.Current);progress.Value=Math.Min(progress.Maximum,batchRun.Completed);RefreshBatchRows();return;}
                FinishBatch(null);
            }
            catch(Exception error){FinishBatch(error);}
            finally{ticking=false;if(sequence!=null)timer.Start();}
        }
        private void FinishBatch(Exception error)
        {
            timer.Stop();if(sequence!=null){sequence.Dispose();sequence=null;}
            state.Text=(stop?"已停止":"批量任务结束")+"：成功 "+batchRun.Succeeded+" 个，失败 "+batchRun.Failed+" 个。";
            if(error!=null)Append("批量任务中止："+error.Message);Append(state.Text);batchRun=null;batchPreparedKey=null;RefreshBatchRows();SetBusy(false);if(closePending)Close();
        }
        private void RestoreBatch()
        {
            var item=SelectedBatchItem();if(item==null||string.IsNullOrEmpty(item.Backup))throw new InvalidOperationException("请在列表中选中一个已完成、具有备份的汇总文件。");
            NotesLinkExecution.CheckOpenWorkbooks(Application,new string[0],item.Parent,true);
            if(UiDialog.Show(this,"将还原所选汇总文件："+Path.GetFileName(item.Parent)+"。当前版本也会保留。是否继续？","还原所选文件",MessageBoxButtons.YesNo,MessageBoxIcon.Warning,MessageBoxDefaultButton.Button2)!=DialogResult.Yes)return;
            string displaced=MigrationOutputTransaction.Restore(item.Parent,item.Backup,item.CommittedHash);item.Backup=null;item.CommittedHash=null;item.Plan=null;item.Status="已还原";Changed();RefreshBatchRows();undo.Enabled=false;Append("已还原："+item.Parent+"；还原前版本："+displaced);
        }
        private void Parameters()
        {
            if(loading||busy)return;
            try{int first=MigrationPlanner.ColumnIndex(firstColumn.Text),last=MigrationPlanner.ColumnIndex(lastColumn.Text);if(last<first)return;sourceRange.Text=new MigrationArea{Row=(int)sourceFirstRow.Value,Column=first,Rows=(int)sourceRows.Value,Columns=last-first+1}.Address;}catch{}
        }
        private void SyncRange()
        {
            var area=MigrationArea.Parse(sourceRange.Text);loading=true;
            try{firstColumn.Text=CellAddress.ColumnLetters(area.Column);lastColumn.Text=CellAddress.ColumnLetters(area.Column+area.Columns-1);sourceFirstRow.Value=area.Row;sourceRows.Value=area.Rows;}finally{loading=false;}
        }
        private void TargetAnchor(MigrationArea area)
        {
            if(area.Rows==1&&area.Columns==1)return;
            string sheet=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(ParentFileForPick()),targetSheet.Text);
            if(!MigrationSheetSnapshot.Load(ParentFileForPick(),sheet).Merges.Any(m=>NotesLinkRules.Same(m,area)))throw new InvalidOperationException("起点请选择一个单元格或一个完整合并单元格。");
        }
        private static void Single(MigrationArea area){if(area.Rows!=1||area.Columns!=1)throw new InvalidOperationException("请选择一个单元格。");}
        private RoundedButton ColumnPick(TextBox field){return Pick("名单字段列",ParentFileForPick,delegate{return leadSheet.Text;},delegate(MigrationArea area){if(area.Columns!=1)throw new InvalidOperationException("请选择单列中的单元格或区域。");field.Text=CellAddress.ColumnLetters(area.Column);},()=>field.Text+((int)firstRow.Value).ToString(),"选择此列");}
        private string ParentFileForPick()
        {
            string path=parent.Text.Trim().Trim('"');
            if(batchMode.Checked)
            {
                var current=SelectedBatchItem();path=NotesLinkBatch.PickerParent(Config(),batchItems,current==null?path:current.Parent);
                loading=true;try{parent.Text=path;foreach(DataGridViewRow row in batchGrid.Rows)if(string.Equals(((NotesLinkBatchItem)row.Tag).Parent,path,StringComparison.OrdinalIgnoreCase)){batchGrid.CurrentCell=row.Cells[1];break;}}finally{loading=false;}
            }
            NotesLinkRules.RequireWorkbook(path);return path;
        }
        private RoundedButton Pick(string role,Func<string> file,Func<string> sheet,Action<MigrationArea> apply,Func<string> current=null,string caption=null)
        {
            var button=Ui.Button(caption??"在Excel中选择",145);button.Click+=delegate{Attempt(delegate{string path=file(),name=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(path),sheet());string initial=current==null?"A1":current();Append("选取"+role+"｜"+Path.GetFileName(path)+"｜"+name+"!"+initial);string address=null;bool previousTop=TopMost;TopMost=false;Hide();try{using(var host=new DataMigrationHost(Application))address=host.PickSourceRange(path,name,initial,null,role,delegate(Action restore){MigrationForeground.Bring(MigrationForeground.ActiveFrame(Application,IntPtr.Zero));restore();});}finally{TopMost=previousTop;Show();Activate();}if(address!=null)apply(MigrationArea.Parse(address));});};return button;
        }
        private RoundedButton Current(string role,Func<string> file,Func<string> expectedSheet,Action<MigrationArea> apply)
        {
            var button=Ui.Button("读取当前选区",125);button.Click+=delegate{Attempt(delegate{
                using(var bag=new ComBag()){
                    dynamic selection=bag.Own(((dynamic)Application).Selection);if(selection==null)throw new InvalidOperationException("请在Excel中选择单元格区域。");dynamic sheet=bag.Own(selection.Worksheet),book=bag.Own(sheet.Parent);
                    string path=file(),name=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(path),expectedSheet());if(!string.Equals(Convert.ToString(book.FullName),Path.GetFullPath(path),StringComparison.OrdinalIgnoreCase)||Convert.ToString(sheet.Name)!=name)throw new InvalidOperationException("请选择指定工作簿的“"+name+"”表中的"+role+"区域。");
                    apply(DataMigrationHost.ReadCoordinates((object)selection,(object)book,(object)sheet));
                }});};return button;
        }
        internal void Shutdown(){if(busy){closePending=true;stop=true;if(execution!=null)execution.Cancelled=true;if(batchRun!=null)batchRun.Cancelled=true;}else Close();}
    }
}
