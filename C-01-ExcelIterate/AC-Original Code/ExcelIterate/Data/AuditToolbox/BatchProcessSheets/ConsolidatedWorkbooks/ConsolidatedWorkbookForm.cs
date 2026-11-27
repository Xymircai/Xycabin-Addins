using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
namespace Excel_Iterate
{
    // Unified template workflow. Existing snapshot and no-overwrite publication remain in CompanyWorkbookHost.
    internal sealed class ConsolidatedWorkbookForm : Form
    {
        internal Func<string,CompanyWorkbookHost> OpenTemplate;
        internal Func<CompanyWorkbookHost> OpenCurrentTemplate;
        internal Func<bool> CanRun;
        internal bool RememberConfigurationEnabled;
        internal event Action<bool> BusyChanged;
        internal readonly CompanyStructurePicker Structure=new CompanyStructurePicker(); // compatibility with existing host binding
        private readonly TextBox structure=Box("",520), columns=Box("A,C,E,G",200), output=Box("",520);
        private readonly ComboBox sheet = new ComboBox { Width = 230, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox single=Box("",520),singleSheet=Box("Lead",150),singleMap=Box("K2,L2,M2,N2",300);
        private readonly TextBox group=Box("",520),groupSheet=Box("Lead",150),groupMap=Box("K3,L3,M3,N3",300);
        private readonly FlowLayoutPanel settings=new FlowLayoutPanel{Dock=DockStyle.Top,AutoSize=true,FlowDirection=FlowDirection.TopDown,WrapContents=false};
        private readonly DataGridView grid=new ResultGrid{Dock=DockStyle.Fill,ReadOnly=true,AllowUserToAddRows=false,RowHeadersVisible=false};
        private readonly RichTextBox log=new RichTextBox{Dock=DockStyle.Fill,ReadOnly=true};
        private readonly Timer timer=new Timer{Interval=40};
        private readonly ProgressBar progress=new ProgressBar{Dock=DockStyle.Bottom};
        private readonly RoundedButton preview=Ui.Button("预检任务",110),start=Ui.Button("开始生成",110),stop=Ui.Button("停止",90);
        private CompanyWorkbookHost singleHost,groupHost;
        private readonly List<Job> jobs=new List<Job>();
        private int position;private bool running,cancelled,closePending;private string frozen, loadedStructurePath;
        private sealed class Company{internal int Row;internal string Sequence,Short,Full,File,Key;}
        private sealed class Job{internal Company Owner;internal List<Company> Rows;internal bool Group;internal MigrationArea[] Areas,Cleanup,OwnerAreas;internal string Sheet,Path;}
        private static TextBox Box(string value,int width){return new TextBox{Text=value,Width=width};}
        private void Row(string title,params Control[] controls){var line=new FlowLayoutPanel{AutoSize=true,WrapContents=true,Width=1080};line.Controls.Add(new Label{Text=title,Width=155,Height=28,TextAlign=ContentAlignment.MiddleLeft});line.Controls.AddRange(controls);settings.Controls.Add(line);}
        private RoundedButton Browse(TextBox field, Action afterSelection = null)
        {
            var button = Ui.Button("选择工作簿", 120);
            button.Click += delegate
            {
                Attempt(delegate
                {
                    var files = HostDialogService.OpenFiles(this, "Excel工作簿 (*.xlsx)|*.xlsx");
                    if (files.Length != 1) return;
                    field.Text = files[0];
                    if (afterSelection != null) afterSelection();
                });
            };
            return button;
        }
        internal ConsolidatedWorkbookForm()
        {
            Text="报表模板生成";DialogSizing.Prepare(this,"company.consolidated","table",Ui.Blue);
            Row("股权架构工作簿",structure,Browse(structure, RefreshStructureSheets));Row("架构工作表",sheet);
            Row("架构字段列",columns,new Label{AutoSize=true,Text="依次：序号、公司简称、公司全称、完整文件名；填写列字母，用逗号分隔"});
            Row("单体模板",single,Browse(single));Row("单体 Lead 名称",singleSheet);Row("单体写入位置",singleMap,new Label{AutoSize=true,Text="依次：序号、简称、全称、文件名（四个单元格）"});
            Row("合并模板",group,Browse(group));Row("合并 Lead 名称",groupSheet);Row("合并写入位置",groupMap,new Label{AutoSize=true,Text="四个单列区域或起点；按直系子公司数向下填充，第2行同列更新母公司，尾部旧名单清空"});
            var folder=Ui.Button("选择文件夹",120);folder.Click+=delegate{string path=HostDialogService.Folder(this,"报表输出文件夹",true);if(!string.IsNullOrEmpty(path))output.Text=path;};Row("输出文件夹",output,folder);
            Row("格式",new Label{AutoSize=true,Text="完整复制模板，保留千分位、金额精度、字体及合并格式；序号、名称、文件名按文本写入。原文件不变，同名不覆盖。"});
            grid.Columns.Add("kind","类型");grid.Columns.Add("company","公司");grid.Columns.Add("rows","Lead 行数");grid.Columns.Add("path","输出文件");grid.Columns.Add("status","状态");Ui.StyleGrid(grid);
            var split=new SplitContainer{Dock=DockStyle.Fill,Orientation=Orientation.Horizontal};split.Panel1.Controls.Add(grid);split.Panel2.Controls.Add(log);
            var root=new TableLayoutPanel{Dock=DockStyle.Fill,RowCount=3,ColumnCount=1,Padding=new Padding(12)};root.RowStyles.Add(new RowStyle(SizeType.Absolute,440));root.RowStyles.Add(new RowStyle(SizeType.Percent,100));root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var scroll=new Panel{Dock=DockStyle.Fill,AutoScroll=true};scroll.Controls.Add(settings);root.Controls.Add(scroll,0,0);root.Controls.Add(split,0,1);
            var actions=Ui.ActionBar();actions.Controls.AddRange(new Control[]{preview,start,stop});root.Controls.Add(actions,0,2);Controls.Add(root);Controls.Add(progress);start.Enabled=stop.Enabled=false;
            preview.Click+=delegate{Attempt(Preview);};start.Click+=delegate{Attempt(Begin);};stop.Click+=delegate{cancelled=true;};timer.Tick+=delegate{Tick();};
            foreach(var field in Fields())field.TextChanged+=delegate{if(!running)start.Enabled=false;};
            structure.TextChanged += delegate
            {
                loadedStructurePath = null;
                sheet.Items.Clear();
            };
            structure.Leave += delegate
            {
                if (StructurePath.Length > 0 && !string.Equals(StructurePath, loadedStructurePath, StringComparison.OrdinalIgnoreCase))
                    Attempt(RefreshStructureSheets);
            };
            sheet.SelectedIndexChanged += delegate { if (!running) start.Enabled = false; };
            FormClosing+=delegate(object sender,FormClosingEventArgs e){if(running){e.Cancel=true;closePending=true;cancelled=true;}};
            Disposed+=delegate{timer.Dispose();Release();Structure.Dispose();};
        }
        private Control[] Fields(){return new Control[]{structure,sheet,columns,output,single,singleSheet,singleMap,group,groupSheet,groupMap};}
        private string StructurePath { get { return structure.Text.Trim().Trim('"'); } }
        private void RefreshStructureSheets()
        {
            string path = StructurePath;
            string selected = string.Equals(path, loadedStructurePath, StringComparison.OrdinalIgnoreCase)
                ? sheet.SelectedItem as string : null;
            loadedStructurePath = null;
            sheet.BeginUpdate();
            try
            {
                sheet.Items.Clear();
                if (path.Length == 0) throw new InvalidOperationException("请选择股权架构工作簿。");
                var names = DataMigrationHost.PackageSheets(path);
                if (names.Count == 0) throw new InvalidOperationException("架构工作簿中没有可用工作表。");
                sheet.Items.AddRange(names.ToArray());
                sheet.SelectedIndex = names.Count == 1 ? 0 : names.FindIndex(name =>
                    string.Equals(name, selected, StringComparison.OrdinalIgnoreCase));
                loadedStructurePath = path;
            }
            finally { sheet.EndUpdate(); }
        }
        private string Key(){return string.Join("\n",Fields().Select(f=>f.Text));}
        private void Append(string message){log.AppendText(DateTime.Now.ToString("HH:mm:ss")+"  "+message+Environment.NewLine);}
        private void Attempt(Action action){try{if(running||CanRun!=null&&!CanRun())throw new InvalidOperationException("请等待当前任务结束。");action();}catch(Exception e){start.Enabled=false;Append("失败："+e.Message);}}
        private void Release(){if(singleHost!=null)singleHost.Dispose();if(groupHost!=null)groupHost.Dispose();singleHost=groupHost=null;}
        internal static string SequenceKey(string text){string value=(text??"").Trim().TrimStart('！','!');if(!Regex.IsMatch(value,@"^\d+(?:-\d+)*-?$"))throw new InvalidOperationException("序号须为数字分段，可带前导！："+text);return string.Join("-",value.TrimEnd('-').Split('-').Select(v=>int.Parse(v).ToString()));}
        private static MigrationArea[] Areas(string text,int count,bool groupMode)
        {
            var areas=text.Split(',','，').Select(s=>MigrationArea.Parse(s.Trim())).ToArray();if(areas.Length!=4)throw new InvalidOperationException("请填写四个字段位置。");
            for(int i=0;i<4;i++){var a=areas[i];if(a.Columns!=1||!groupMode&&a.Rows!=1||a.Rows!=1&&a.Rows<count)throw new InvalidOperationException("写入区域必须是单列且足够容纳任务行数。");if(a.Row+count-1>1048576)throw new InvalidOperationException("写入超出行上限。");a.Rows=count;for(int j=0;j<i;j++)if(NotesOverlap(a,areas[j]))throw new InvalidOperationException("字段写入区域重叠，请分别指定序号、简称、全称和文件名位置。");}return areas;
        }
        internal static MigrationArea[] ParentAreas(MigrationArea[] children)
        {
            if(children.Any(a=>a.Row<=2))throw new InvalidOperationException("合并名单必须从第3行或之后开始，第2行用于母公司。");
            return children.Select(a=>new MigrationArea{Row=2,Column=a.Column,Rows=1,Columns=1}).ToArray();
        }
        private static bool NotesOverlap(MigrationArea a,MigrationArea b){return a.Column==b.Column&&a.Row<b.Row+b.Rows&&b.Row<a.Row+a.Rows;}
        private void Preview()
        {
            start.Enabled=false;Release();jobs.Clear();grid.Rows.Clear();if(OpenTemplate==null)throw new InvalidOperationException("宿主未连接。");if(!Directory.Exists(output.Text))throw new InvalidOperationException("请选择已存在的输出文件夹。");
            var cols=columns.Text.Split(',','，').Select(c=>MigrationArea.Parse(c.Trim()+"1").Column).ToArray();if(cols.Length!=4||cols.Distinct().Count()!=4)throw new InvalidOperationException("架构表四个字段列必须不同。");
            RefreshStructureSheets();
            string actual = sheet.SelectedItem as string;
            if (actual == null) throw new InvalidOperationException("架构工作簿包含多张工作表，请从下拉列表选择架构表。");
            var companies=new List<Company>();
            using(var sourceHost=OpenTemplate(StructurePath))foreach(var data in sourceHost.ReportRows(actual,cols))
            {int r=int.Parse(data[0]);string seq=data[1].Trim();if(seq.Length==0)continue;if(!Regex.IsMatch(seq.TrimStart('！','!'),@"^\d+(?:-\d+)*-?$")){if(r<=20&&new[]{"序号","公司序号","编号","编码"}.Contains(seq))continue;throw new InvalidOperationException("架构表第"+r+"行序号无效。");}companies.Add(new Company{Row=r,Sequence=seq,Key=SequenceKey(seq),Short=data[2],Full=data[3],File=data[4]});}
            if(companies.Count==0)throw new InvalidOperationException("没有识别到公司任务。");if(companies.Select(c=>c.Key).Distinct().Count()!=companies.Count)throw new InvalidOperationException("架构表序号重复。");
            foreach(var company in companies.Where(c=>c.Key.Contains("-"))){string parentKey=company.Key.Substring(0,company.Key.LastIndexOf('-'));if(!companies.Any(c=>c.Key==parentKey))throw new InvalidOperationException("架构表存在缺失的直属父级："+company.Sequence);}
            foreach(var company in companies.Where(c=>c.Key.Contains("-"))){string parentKey=company.Key.Substring(0,company.Key.LastIndexOf('-'));if(!companies.Any(c=>c.Key==parentKey))throw new InvalidOperationException("架构表存在缺失的直属父级："+company.Sequence);}
            singleHost=OpenTemplate(single.Text);groupHost=OpenTemplate(group.Text);
            if(companies.Any(c=>string.IsNullOrWhiteSpace(c.Short)||string.IsNullOrWhiteSpace(c.Full)||string.IsNullOrWhiteSpace(c.File)))throw new InvalidOperationException("架构表公司简称、全称和文件名称不能为空。");
            foreach(var owner in companies)
            {
                var children=companies.Where(c=>c.Key.StartsWith(owner.Key+"-",StringComparison.Ordinal)&&c.Key.Split('-').Length==owner.Key.Split('-').Length+1||owner.Key=="0"&&c.Key!="0"&&!c.Key.Contains("-")).ToList();bool isGroup=children.Count>0;
                var rows=isGroup?children:new List<Company>{owner};var host=isGroup?groupHost:singleHost;string target=isGroup?groupSheet.Text:singleSheet.Text;
                string targetName=host.Sheets.SingleOrDefault(n=>string.Equals(n,target.Trim(),StringComparison.OrdinalIgnoreCase));if(targetName==null)throw new InvalidOperationException("模板工作表不存在："+target);
                string file=CompanyWorkbookRules.FileName(owner.File);if(!file.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("文件名称须带.xlsx："+file);
                string path=Path.Combine(Path.GetFullPath(output.Text),file);if(File.Exists(path)||jobs.Any(j=>string.Equals(j.Path,path,StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("输出已存在或重名："+path);
                var job=new Job{Owner=owner,Rows=rows,Group=isGroup,Sheet=targetName,Path=path,Areas=Areas(isGroup?groupMap.Text:singleMap.Text,rows.Count,isGroup)};
                job.OwnerAreas=isGroup?ParentAreas(job.Areas):new MigrationArea[0];
                foreach(var area in job.OwnerAreas)host.ValidateReportArea(targetName,area);
                job.Cleanup=job.Areas.Select(a=>new MigrationArea{Row=a.Row,Column=a.Column,Columns=1,Rows=isGroup?Math.Max(a.Rows,host.ReportLastRow(targetName,a.Column)-a.Row+1):a.Rows}).ToArray();
                foreach(var area in job.Cleanup)host.ValidateReportArea(targetName,area);
                jobs.Add(job);grid.Rows.Add(isGroup?"合并":"单体",owner.Full,rows.Count,path,"待生成");
            }
            frozen=Key();start.Enabled=true;Append("预检：合并 "+jobs.Count(j=>j.Group)+" 个，单体 "+jobs.Count(j=>!j.Group)+" 个。请核对任务数量和输出位置；12/59仅为当前样本预期，不硬编码。");
        }
        private void Begin(){if(frozen!=Key()||jobs.Count==0)throw new InvalidOperationException("请重新预检。");if(UiDialog.Show(this,"将生成 "+jobs.Count+" 个工作簿，仅覆盖副本中配置的Lead位置，保留模板数字格式。是否继续？","报表模板生成",MessageBoxButtons.OKCancel,MessageBoxIcon.Question)!=DialogResult.OK)return;position=0;cancelled=false;running=true;settings.Enabled=preview.Enabled=start.Enabled=false;stop.Enabled=true;progress.Maximum=jobs.Count;progress.Value=0;if(BusyChanged!=null)BusyChanged(true);timer.Start();}
        private static string Value(Company row,int column){return column==0?row.Sequence:column==1?row.Short:column==2?row.Full:row.File;}
        private static void Write(object book,Job job,bool verify)
        {
            using(var bag=new ComBag()){dynamic sheets=bag.Own(((dynamic)book).Worksheets),lead=bag.Own(sheets.Item(job.Sheet));if(Convert.ToBoolean(lead.ProtectContents))throw new InvalidOperationException("目标工作表受保护。");
                for(int c=0;c<job.OwnerAreas.Length;c++)
                {
                    var area=job.OwnerAreas[c];dynamic cell=bag.Own(lead.Cells[area.Row,area.Column]);
                    if(!verify)CompanyWorkbookHost.WriteText(cell,Value(job.Owner,c));
                    else if(Convert.ToString(cell.Value2)!=Value(job.Owner,c)||Convert.ToBoolean(cell.HasFormula))throw new InvalidOperationException("母公司字段保存后核验失败。");
                }
                for(int c=0;c<4;c++)
                {
                    var area=job.Cleanup[c];
                    if(!verify){dynamic clear=bag.Own(lead.Range[area.Address]);clear.ClearContents();}
                    else for(int r=job.Rows.Count;r<area.Rows;r++){dynamic tail=bag.Own(lead.Cells[area.Row+r,area.Column]);if(!string.IsNullOrEmpty(Convert.ToString(tail.Value2))||Convert.ToBoolean(tail.HasFormula))throw new InvalidOperationException("保存后旧名单清理核验失败："+CellAddress.Name(area.Row+r,area.Column));}
                }
                for(int c=0;c<4;c++)for(int r=0;r<job.Rows.Count;r++){dynamic cell=bag.Own(lead.Cells[job.Areas[c].Row+r,job.Areas[c].Column]);if(!verify){if(Convert.ToBoolean(cell.MergeCells)||Convert.ToBoolean(cell.HasArray))throw new InvalidOperationException("写入位置包含合并或数组公式。");CompanyWorkbookHost.WriteText(cell,Value(job.Rows[r],c));}else if(Convert.ToString(cell.Value2)!=Value(job.Rows[r],c)||Convert.ToBoolean(cell.HasFormula))throw new InvalidOperationException("保存后字段读回不一致。");}}
        }
        private void Tick(){timer.Stop();if(cancelled||position>=jobs.Count){Finish();return;}var job=jobs[position];try{var host=job.Group?groupHost:singleHost;host.Generate(new CompanyWorkbookRequest{Output=output.Text,UpdateLead=true},new CompanyWorkbookRecord{Row=job.Owner.Row,Name=job.Owner.File,OutputPath=job.Path},(b,q,r)=>Write(b,job,false),Append,(b,q,r)=>Write(b,job,true));grid.Rows[position].Cells[4].Value="已生成";Append("完成："+job.Path);}catch(Exception e){grid.Rows[position].Cells[4].Value="失败";Append("失败："+job.Path+"；"+e.Message);}position++;progress.Value=position;timer.Start();}
        private void Finish(){running=false;settings.Enabled=preview.Enabled=true;stop.Enabled=false;if(BusyChanged!=null)BusyChanged(false);Append(cancelled?"已停止，已生成文件保留。":"任务结束，请检查逐项状态。");if(closePending)Close();}
        internal void RestoreLastConfiguration(){}
        internal void Shutdown(){if(running){cancelled=closePending=true;}else Close();}
    }
}
