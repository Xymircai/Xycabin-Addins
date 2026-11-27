using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Sequential STA controller. Only private disk copies are ever opened or saved.
    internal sealed class SummaryFileBatch : IDisposable
    {
        private object app;
        private readonly object sourceApp;
        private HiddenWorkbookApplication applicationRefs;
        private bool ownsApplication;
        private readonly Form dialog = new Form();
        private readonly TableLayoutPanel settings = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 4, Margin = new Padding(0, 0, 0, 12) };
        private readonly TextBox input = new TextBox { Dock = DockStyle.Fill, ShortcutsEnabled = true, AccessibleDescription = "输入或粘贴文件夹路径，按 Enter 或点击读取文件。" }, output = new TextBox { Dock = DockStyle.Fill }, presetPath = new TextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly TextBox pathSheet = new TextBox { Text = "单体数据汇总", Dock = DockStyle.Fill }, pathCell = new TextBox { Text = "A1", Dock = DockStyle.Fill };
        private readonly DataGridView files = new ResultGrid { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false };
        private readonly RichTextBox log = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true };
        private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Top };
        private readonly RoundedButton preview = Ui.Button("批量预检", 120), execute = Ui.Button("执行并另存", 135), stop = Ui.Button("停止", 90);
        private readonly Timer timer = new Timer { Interval = 40 };
        private readonly TextBox templateName=new TextBox{Text="单体数据汇总"},templateArea=new TextBox{Text="A3:E6;A8:E208"},leadName=new TextBox{Text="Lead"},leadArea=new TextBox{Text="L3:L17"},pasteStart=new TextBox{Text="L3"};
        private readonly TextBox linkFileColumn=new TextBox{Text="N",Width=65};
        private readonly CheckBox strictFiles=new CheckBox{Text="严格检查目录及来源文件存在（虚拟路径暂不可用时请取消）",AutoSize=true};
        private readonly Dictionary<string,TextBox> variableFields=new Dictionary<string,TextBox>();
        private sealed class VariableGroup{internal string Kind;internal TextBox Source=new TextBox{Width=300},Target=new TextBox{Width=300};internal Control Row;}
        private readonly List<VariableGroup> groups=new List<VariableGroup>();
        private static Control Inline(params Control[] controls){var panel=new FlowLayoutPanel{AutoSize=true,WrapContents=true};panel.Controls.AddRange(controls);return panel;}
        private void AddVariable(string kind,string source="",string target="")
        {
            var item=new VariableGroup{Kind=kind};item.Source.Text=source;item.Target.Text=target;
            var remove=Ui.Button("删除",70);var line=Inline(item.Source,new Label{Text="粘贴地址",AutoSize=true},item.Target,remove);
            AddRow(kind=="repeat"?"重复变量（表名!地址）":"年度变量（表名!地址）",line);item.Row=line.Parent;groups.Add(item);
            remove.Click+=delegate{settings.Controls.Remove(item.Row);item.Row.Dispose();groups.Remove(item);};
        }
        private readonly Timer refreshTimer=new Timer{Interval=2000};
        private readonly System.Text.StringBuilder pendingLog=new System.Text.StringBuilder();
        private readonly Label progressText=new Label{AutoSize=true,Dock=DockStyle.Top};
        private readonly HashSet<string> warningMessages=new HashSet<string>(),errorFiles=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string currentProgress="";
        private List<DataGridViewRow> queue;
        private BatchPreset preset;
        private BatchPreset chosenPreset;
        private List<CompanyStructureTask> structureTasks;
        private List<CompanyParentLinkTask> parentGroups;
        private Dictionary<string,List<CompanyStructureTask>> parentChildren;
        internal bool StructureUsesLevel = true;
        private string destination, sourceSheet, sourceCell, scratch, workingFile, loadedInput;
        private int position;
        private bool writing, stopping, inTick, currentStarted, disposed;
        private ComBag refs;
        private object workBook;
        private SingleSummaryForm workerForm;
        private SummaryAreaHost worker;
        internal static void Show(object application, Form launcher)
        {
            var batch = new SummaryFileBatch(application);
            EventHandler closeBatch = delegate { batch.Dispose(); };
            bool launcherWasEnabled = launcher.Enabled;
            launcher.Disposed += closeBatch;
            batch.dialog.Disposed += delegate
            {
                launcher.Disposed -= closeBatch;
                if (!launcher.IsDisposed) launcher.Enabled = launcherWasEnabled;
            };
            batch.dialog.FormClosed += delegate { batch.Dispose(); };
            try
            {
                launcher.Enabled = false;
                // Keep the workbook interactive and let the user restore the batch
                // window from the taskbar independently of the launcher.
                batch.dialog.Show();
            }
            catch
            {
                batch.Dispose();
                throw;
            }
        }
        private SummaryFileBatch(object application)
        {
            sourceApp = app = application; dialog.Text = "合并批量链接 · 多文件";
            DialogSizing.Prepare(dialog, "summary.single", "table", Ui.Blue);
            settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); settings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var chooseInput = Ui.Button("选择文件夹", 155);
            chooseInput.Click += delegate { Attempt(delegate {
                string path = HostDialogService.Folder(dialog, "选择待处理工作簿文件夹", false); if (string.IsNullOrEmpty(path)) return;
                input.Text = path; LoadFiles();
            }); };
            var loadInput = Ui.Button("读取文件", 115); loadInput.Click += delegate { Attempt(LoadFiles); };
            input.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Attempt(LoadFiles); } };
            input.TextChanged += delegate { if (queue == null) { loadedInput = null;parentChildren=null;files.Rows.Clear(); } };
            var choosePreset = Ui.Button("选择已保存预设", 155);
            choosePreset.Click += delegate { Attempt(delegate {
                string json = HostDialogService.ShowIsolatedResult<string>(dialog,
                    delegate { return new BatchPresetManagerForm("summary"); },
                    delegate(Form chooser) { var value = ((BatchPresetManagerForm)chooser).Chosen; return value == null ? null : BatchPresetStore.Encode(value); }, null);
                if (json != null) UsePreset(BatchPresetStore.Decode(json, "summary"), "已保存：");
            }); };
            var importPreset = Ui.Button("导入预设", 115);
            importPreset.Click += delegate { Attempt(delegate { string[] paths = HostDialogService.OpenFiles(dialog, "合并链接预设 (*.xybatch.json)|*.xybatch.json"); if (paths.Length != 1) return; UsePreset(BatchPresetStore.Read(paths[0], "summary"), "已导入："); }); };
            var chooseOutput = Ui.Button("选择输出文件夹", 155);
            chooseOutput.Click += delegate { string path = HostDialogService.Folder(dialog, "选择结果输出文件夹", true); if (!string.IsNullOrEmpty(path)) output.Text = path; };
            AddRow("输入文件夹", input, chooseInput, loadInput); AddRow("使用预设", presetPath, choosePreset, importPreset); AddRow("输出文件夹", output, chooseOutput);
            AddParameterRow("路径来源表", pathSheet, "单元格", pathCell);
            AddParameterRow("模板工作表", templateName, "已编排区域", templateArea, "粘贴起点", pasteStart);
            AddParameterRow("Lead 工作表", leadName, "公司范围", leadArea);
            AddRow("外链文件名称列",Inline(linkFileColumn,new Label{Text="读取同一Lead行的完整文件名（含.xlsx）；列名仍由重复变量取简称。",AutoSize=true}));
            var addRepeat=Ui.Button("添加重复变量",140);addRepeat.Click+=delegate{AddVariable("repeat");};var addYear=Ui.Button("添加年度变量",140);addYear.Click+=delegate{AddVariable("year");};
            AddRow("变量设置",Inline(addRepeat,addYear,new Label{Text="来源与粘贴地址均填写 表名!地址；可留空或删除不需要的组。",AutoSize=true}));
            AddRow("路径校验",strictFiles);
            AddVariable("repeat","Lead!L3:L17","单体数据汇总!L2");AddVariable("year","Lead!J5:J6","单体数据汇总!L7");
            var presetActions=new BatchPresetActions(dialog,"summary",ManualPreset,value=>UsePreset(value,"已应用："),delegate{return queue==null;},Message);
            AddRow("保存／管理预设",presetActions.Header(new Label{Text="可直接填写，也可载入预设后修改；多区域用分号分隔。",AutoSize=true}));

            files.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "处理", Width = 60 });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "文件", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            files.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", ReadOnly = true, Width = 200 }); Ui.StyleGrid(files);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 24));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 18));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var actions = Ui.ActionBar(); actions.Controls.AddRange(new Control[] { preview, execute, stop });
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            scroll.Controls.Add(settings);
            var progressPanel = new Panel { Dock = DockStyle.Top, AutoSize = true };
            progressPanel.Controls.Add(progress);
            progressPanel.Controls.Add(progressText);
            root.Controls.Add(scroll, 0, 0);
            root.Controls.Add(files, 0, 1);
            root.Controls.Add(progressPanel, 0, 2);
            root.Controls.Add(log, 0, 3);
            root.Controls.Add(actions, 0, 4);
            dialog.Controls.Add(root);
            preview.Click += delegate { Attempt(delegate { Begin(false); }); }; execute.Click += delegate { Attempt(delegate { Begin(true); }); };
            stop.Enabled = false; stop.Click += delegate { stopping = true; Message("正在停止；当前临时副本将丢弃，已经输出的成功文件保留。"); };
            timer.Tick += delegate { Advance(); };refreshTimer.Tick+=delegate{FlushUi();};refreshTimer.Start();
            dialog.FormClosing += delegate(object sender, FormClosingEventArgs e) { if (queue != null) { e.Cancel = true; stopping = true; Message("正在停止，请等待当前操作结束后关闭。"); } };
        }
        private void AddRow(string label, Control value, Control action = null, Control secondary = null)
        {
            int row = settings.RowCount++; settings.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            settings.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 12, 7) }, 0, row);
            value.Dock = DockStyle.None; value.Anchor = AnchorStyles.Left | AnchorStyles.Right; value.Margin = new Padding(0, 4, 12, 4); settings.Controls.Add(value, 1, row);
            if (action != null) { action.Anchor = AnchorStyles.Left; action.Margin = new Padding(0, 3, 8, 3); settings.Controls.Add(action, 2, row); }
            if (secondary != null) { secondary.Anchor = AnchorStyles.Left; secondary.Margin = new Padding(0, 3, 0, 3); settings.Controls.Add(secondary, 3, row); }
        }
        private void AddParameterRow(string label, TextBox sheet, string addressLabel, TextBox address,
            string endLabel = null, TextBox end = null)
        {
            var fields = new TableLayoutPanel
            {
                ColumnCount = 5, RowCount = 1, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            int labelWidth = TextRenderer.MeasureText("已编排区域", dialog.Font).Width + 16;
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 3));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, labelWidth));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 3));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, labelWidth));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / 3));
            fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            AddParameterField(fields, sheet, 0);
            AddParameterField(fields, address, 2);
            fields.Controls.Add(ParameterLabel(addressLabel), 1, 0);
            if (end != null)
            {
                AddParameterField(fields, end, 4);
                fields.Controls.Add(ParameterLabel(endLabel), 3, 0);
            }
            AddRow(label, fields);
            settings.SetColumnSpan(fields, 3);
        }

        private static void AddParameterField(TableLayoutPanel fields, TextBox field, int column)
        {
            field.Dock = DockStyle.Fill;
            field.Margin = Padding.Empty;
            fields.Controls.Add(field, column, 0);
        }

        private static Label ParameterLabel(string text)
        {
            return new Label
            {
                Text = text, AutoSize = true, Anchor = AnchorStyles.Right,
                Margin = new Padding(8, 0, 8, 0)
            };
        }
        private void LoadFiles()
        {
            if (queue != null) return;
            string path = CellPath(input.Text);
            var found = structureTasks==null?Directory.GetFiles(path).Where(IsWorkbook).Where(p => !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal)).OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray():CompanyStructureTasks.ResolveFiles(structureTasks,path).ToArray();
            input.Text = path; files.Rows.Clear();
            foreach (string file in found) { int row = files.Rows.Add(structureTasks!=null||DefaultSelected(file), Path.GetFileName(file), File.Exists(file)?"待预检":"未找到文件"); files.Rows[row].Tag = file; }
            parentChildren=null;
            if(parentGroups!=null)
            {
                parentChildren=new Dictionary<string,List<CompanyStructureTask>>(StringComparer.OrdinalIgnoreCase);
                for(int i=0;i<found.Length;i++){parentChildren[found[i]]=new List<CompanyStructureTask>(parentGroups[i].Children);files.Rows[i].Cells[1].ToolTipText=parentGroups[i].Parent.FullName+" ← "+string.Join("、",parentGroups[i].Children.Select(c=>c.FileName).ToArray());}
            }
            loadedInput = path;
            Message(structureTasks==null?"已列出 " + found.Length + " 个工作簿，按文件名识别（合并）或(合并)主体；逐文件处理结果见日志。":"已按架构表顺序匹配"+found.Length+"个任务；未找到的文件会明确报告，不使用近似名称替代。");
        }
        internal static bool IsWorkbook(string file)
        { string extension = Path.GetExtension(file); return string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) || string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase); }
        internal static bool DefaultSelected(string file)
        {
            string name = Path.GetFileNameWithoutExtension(file);
            return name.Contains("（合并）") || name.Contains("(合并)");
        }
        internal static string CellPath(object value){return CellPath(value,true);}
        internal static string CellPath(object value,bool requireExists)
        {
            string text = value as string;
            if (string.IsNullOrWhiteSpace(text)) throw new InvalidOperationException("路径来源单元格为空、错误或不是文字。");
            text = text.Trim(); if (text.Length > 1 && text[0] == '"' && text[text.Length - 1] == '"') text = text.Substring(1, text.Length - 2);
            if (!System.Text.RegularExpressions.Regex.IsMatch(text,@"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+)")) throw new InvalidOperationException("路径来源单元格不是完整绝对路径："+text);
            try{Path.GetFullPath(text);}catch(Exception error){throw new InvalidOperationException("路径格式无效："+text,error);}
            if(requireExists&&!Directory.Exists(text))throw new InvalidOperationException("文件夹不存在："+text);
            return text;
        }
        internal static BatchPreset Rebind(BatchPreset original, string file, string folder)
        { var copy = BatchPresetStore.Decode(BatchPresetStore.Encode(original), "summary"); copy.Context = file; copy.Values["projectFolder"] = folder; return copy; }
        internal void UsePreset(BatchPreset value, string prefix)
        {
            if (queue != null) throw new InvalidOperationException("执行期间不能更换预设。");
            var snapshot = BatchPresetStore.Decode(BatchPresetStore.Encode(value), "summary");
            SummaryAreaSettings.Read(snapshot);
            chosenPreset = snapshot; presetPath.Text = prefix + snapshot.Name;
            var cfg=SummaryAreaSettings.Read(snapshot);templateName.Text=cfg.Template[0].Sheet;templateArea.Text=string.Join(";",cfg.Template.Select(a=>a.Address));pasteStart.Text=cfg.Paste.Address;leadName.Text=cfg.Lead[0].Sheet;leadArea.Text=string.Join(";",cfg.Lead.Select(a=>a.Address));strictFiles.Checked=cfg.CheckFiles;string savedFileColumn;linkFileColumn.Text=snapshot.Values.TryGetValue("linkFileColumn",out savedFileColumn)?savedFileColumn:"";
            foreach(var item in groups.ToArray()){settings.Controls.Remove(item.Row);item.Row.Dispose();}groups.Clear();
            Action<string,string,string> restore=(kind,sourceKey,targetKey)=>{
                var source=SummaryAreaSettings.OptionalHeader(snapshot.Values,sourceKey);var target=SummaryAreaSettings.OptionalHeader(snapshot.Values,targetKey);if(source==null&&target==null)return;
                string text=source==null?"":source.Sheet+"!"+source.Address;string multi;if(source!=null&&snapshot.Values.TryGetValue(sourceKey+"Multi",out multi))text=source.Sheet+"!"+string.Join(";",SummaryAreaSettings.Decode(multi,source.Sheet).Select(p=>p.Address));
                AddVariable(kind,text,target==null?"":target.Sheet+"!"+target.Address);
            };
            restore("repeat","repeatSource","repeatTarget");restore("year","yearSource","yearTarget");
            string extra;if(snapshot.Values.TryGetValue("extraGroups",out extra)&&!string.IsNullOrEmpty(extra))foreach(string key in extra.Split(';'))restore(key.StartsWith("repeat_")?"repeat":"year",key+"Source",key+"Target");
            if(cfg.CompanyHeader!=null&&!groups.Any(g=>g.Kind=="repeat")){
                var part=cfg.Lead[0];var source=new SummaryAreaPlan{Sheet=part.Sheet,Row=part.Row,Column=part.Column,Rows=part.Rows,Columns=1};AddVariable("repeat",source.Sheet+"!"+SummaryAreaEngine.Address(source),cfg.CompanyHeader.Sheet+"!"+cfg.CompanyHeader.Address);
                Message("旧公司表头已迁移为一组重复变量，请核对来源列。");}
            if(groups.Count==0){AddVariable("repeat");AddVariable("year");}
            string saved;if(snapshot.Values.TryGetValue("batchPathSheet",out saved))pathSheet.Text=saved;if(snapshot.Values.TryGetValue("batchPathCell",out saved))pathCell.Text=saved;
            foreach (DataGridViewRow row in files.Rows) row.Cells[2].Value = "待预检";
            Message("已选择预设“" + snapshot.Name + "”；仅使用坐标配置，各文件读取自己的数据，无需原预设工作簿存在。");
        }
        private static string EncodeAddresses(string text,string sheet)
        {var parts=text.Split(';','；').Select(t=>MigrationArea.Parse(t.Trim())).Select(a=>new SummaryAreaPlan{Sheet=sheet,Row=a.Row,Column=a.Column,Rows=a.Rows,Columns=a.Columns});return SummaryAreaSettings.Encode(parts);}
        private BatchPreset ManualPreset()
        {
            var value=chosenPreset==null?new BatchPreset{Feature="summary",Name="多文件自定义",Context="",Values=new Dictionary<string,string>()}:BatchPresetStore.Decode(BatchPresetStore.Encode(chosenPreset),"summary");
            var v=value.Values;v["linkFileColumn"]=linkFileColumn.Text.Trim();v["mode"]="area-mvp";v["templateSheet"]=templateName.Text.Trim();v["templateAreas"]=EncodeAddresses(templateArea.Text,v["templateSheet"]);v["pasteSheet"]=v["templateSheet"];v["pasteArea"]=EncodeAddresses(pasteStart.Text,v["templateSheet"]);v["leadSheet"]=leadName.Text.Trim();v["leadAreas"]=EncodeAddresses(leadArea.Text,v["leadSheet"]);v["projectFolder"]=Path.GetTempPath();v["checkFiles"]=strictFiles.Checked.ToString();v["batchPathSheet"]=pathSheet.Text;v["batchPathCell"]=pathCell.Text;
            if(chosenPreset==null){v["headersFormula"]="True";v["headersRepeat"]="True";v["variableHeaders"]="False";}
            foreach(string key in v.Keys.Where(k=>System.Text.RegularExpressions.Regex.IsMatch(k,@"^(repeat|year)_\d+(Source|Target)(Sheet|Multi)?$")).ToArray())v.Remove(key);
            foreach(string key in new[]{"companyHeader","repeatSource","repeatTarget","yearSource","yearTarget"}){v[key]="";v[key+"Sheet"]="";v.Remove(key+"Multi");}
            var extra=new List<string>();bool repeatUsed=false,yearUsed=false;int id=0;int active=0;
            foreach(var item in groups){string from=item.Source.Text.Trim(),to=item.Target.Text.Trim();if(from.Length==0&&to.Length==0)continue;if(from.Length==0||to.Length==0)throw new InvalidOperationException("每组变量必须同时填写来源和粘贴地址。");
                bool first=item.Kind=="repeat"?!repeatUsed:!yearUsed;string groupKey=first?item.Kind:item.Kind+"_"+(++id);if(item.Kind=="repeat")repeatUsed=true;else yearUsed=true;if(!first)extra.Add(groupKey);
                foreach(var pair in new[]{new KeyValuePair<string,string>("Source",from),new KeyValuePair<string,string>("Target",to)}){
                    string key=groupKey=="repeat"?"repeat"+pair.Key:groupKey=="year"?(pair.Key=="Source"?"yearSource":"yearTarget"):groupKey+pair.Key;
                    int split=pair.Value.LastIndexOf('!');if(split<1)throw new InvalidOperationException("变量请填写 工作表!地址。");string name=pair.Value.Substring(0,split).Trim().Trim('\'');string encoded=EncodeAddresses(pair.Value.Substring(split+1),name);var parts=SummaryAreaSettings.Decode(encoded,name);
                    v[key+"Sheet"]=name;v[key]=encoded;if(item.Kind=="repeat"&&pair.Key=="Source"){v[key+"Multi"]=encoded;v[key]=SummaryAreaSettings.Encode(parts.Take(1));}
                }active++;
            }
            v["extraGroups"]=string.Join(";",extra);v["headersEnabled"]=(active>0).ToString();v["variableHeaders"]="True";v["headersFormula"]="True";
            SummaryAreaSettings.Read(value);return value;
        }
        private void Message(string text)
        {pendingLog.Append(DateTime.Now.ToString("HH:mm:ss")).Append("  ").AppendLine(text);if(queue==null)FlushUi();}
        private void FlushUi()
        {
            if(dialog.IsDisposed)return;
            if(pendingLog.Length>0){log.AppendText(pendingLog.ToString());pendingLog.Clear();log.SelectionStart=log.TextLength;log.ScrollToCaret();}
            progress.Value=Math.Min(progress.Maximum,position);
            progressText.Text=queue==null?currentProgress:"已完成 "+position+" / "+queue.Count+" 个文件；"+currentProgress;
        }
        private void Attempt(Action action) { try { action(); } catch (Exception error) { Message("失败：" + error.Message); } }
        private void Begin(bool save)
        {
            if (queue != null) return; files.EndEdit();
            if (loadedInput == null) throw new InvalidOperationException("输入路径后，请先按 Enter 或点击读取文件，再勾选工作簿。");
            var selected = files.Rows.Cast<DataGridViewRow>().Where(r => object.Equals(r.Cells[0].Value, true)).ToList();
            if (selected.Count == 0) throw new InvalidOperationException("请勾选需要处理的文件。");
            var selectedPreset=ManualPreset();SummaryAreaSettings.Read(selectedPreset);
            string directory = Path.GetFullPath(output.Text.Trim()); if (!Directory.Exists(directory)) throw new InvalidOperationException("请选择已存在的输出文件夹。");
            if (!System.Text.RegularExpressions.Regex.IsMatch(pathCell.Text.Trim(), @"^\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}$") || string.IsNullOrWhiteSpace(pathSheet.Text)) throw new InvalidOperationException("请填写路径来源表和一个单元格地址，例如 A1。");
            if (save && UiDialog.Show(dialog, "将逐文件重新预检，并覆盖各临时副本中预设指定的目标区域及表头；成功后自动保存到输出文件夹，原文件保持不变，同名输出不覆盖。失败文件跳过，其他文件继续。确认执行？", "批量输出确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.OK) return;
            preset = selectedPreset; destination = directory; sourceSheet = pathSheet.Text.Trim(); sourceCell = pathCell.Text.Trim(); writing = save; stopping = false; position = 0; queue = selected;errorFiles.Clear();warningMessages.Clear();currentProgress="准备中";
            foreach (var row in queue) row.Cells[2].Value = "等待";
            settings.Enabled = files.Enabled = preview.Enabled = execute.Enabled = false; stop.Enabled = true; progress.Maximum = queue.Count; progress.Value = 0;
            Message((save ? "开始执行" : "开始只读预检") + "，共 " + queue.Count + " 个文件；每个文件重新读取自己的数据，预检不保存任何工作簿。"); timer.Start();
        }
        private void OpenCurrent()
        {
            currentStarted=false;
            string inputFile = Convert.ToString(queue[position].Tag);
            string outputFile = Path.Combine(destination, Path.GetFileName(inputFile));
            if (string.Equals(Path.GetFullPath(inputFile), Path.GetFullPath(outputFile), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("输出文件夹不能覆盖输入文件，请选择其他文件夹。");
            if (File.Exists(outputFile)) throw new InvalidOperationException("输出文件已存在，已跳过，不覆盖。");
            scratch = Path.Combine(destination, ".xybatch-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(scratch);
            workingFile = Path.Combine(scratch, ".xybatch-" + Guid.NewGuid().ToString("N") + Path.GetExtension(inputFile));
            bool copiedOpen = false;
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)sourceApp).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic existing = bag.Own(books.Item(i));
                    if (!string.Equals(Convert.ToString(existing.FullName), Path.GetFullPath(inputFile), StringComparison.OrdinalIgnoreCase)) continue;
                    using (var quiet = new CompanyWorkbookQuiet(sourceApp)) existing.SaveCopyAs(workingFile);
                    copiedOpen = true;
                    Message("读取已打开工作簿的当前内容（包含未保存修改）：" + inputFile + "；不保存、不关闭原工作簿。");
                    break;
                }
            }
            if (!copiedOpen) CopyInputSnapshot(inputFile,workingFile);
            EnsureHiddenApplication();
            refs = new ComBag();
            using (var quiet = new CompanyWorkbookQuiet(app))
            {
                dynamic books = refs.Own(((dynamic)app).Workbooks); workBook = refs.Own(books.Open(workingFile, 0, false));
                if (workBook == null) throw new InvalidOperationException("无法打开工作副本。");
                // Keep the workbook's own window eligible for Worksheet.Copy.
                // The entire dedicated application stays invisible instead.
            }
            string folder;
            using (var bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)workBook).Worksheets), sheet = bag.Own(sheets.Item(sourceSheet)), cell = bag.Own(sheet.Range[sourceCell]);
                folder = CellPath((object)cell.Value2,false);
                if(!Directory.Exists(folder))Message("提示：链接目录暂不可用："+folder+"；保留原路径，未创建目录。严格检查关闭时继续生成链接，来源恢复前不能保证显示值。");
            }
            workerForm = new SingleSummaryForm(true); var view = (SummaryAreaView)workerForm.Controls[0];
            warningMessages.Clear();view.LogSink = delegate(string text) { if((text.StartsWith("提示：")||text.StartsWith("重新计算失败："))&&warningMessages.Add(text))Message(Path.GetFileName(inputFile)+"｜"+text); };
            worker = new SummaryAreaHost(app, workerForm, view, null, delegate { return true; });
            List<CompanyStructureTask> required=null;
            if(parentChildren!=null&&!parentChildren.TryGetValue(inputFile,out required))throw new InvalidOperationException("该父公司缺少架构任务分组，请重新读取架构表。");
            worker.PrepareBatchFile(Rebind(preset, workingFile, folder),required);
            if (writing) { worker.StartBatchFile();currentStarted=true; }
        }
        internal static void CopyInputSnapshot(string inputFile,string targetFile)
        {
            // Another Excel instance or sync client may hold a writer handle.
            // Permit shared reads, but reject a file that changes during copying.
            var before=new FileInfo(inputFile);long length=before.Length;DateTime modified=before.LastWriteTimeUtc;
            try
            {
                using(var source=new FileStream(inputFile,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))
                using(var target=new FileStream(targetFile,FileMode.CreateNew,FileAccess.Write))source.CopyTo(target);
                var after=new FileInfo(inputFile);
                if(after.Length!=length||after.LastWriteTimeUtc!=modified)throw new IOException("输入文件在读取期间发生变化，请等待保存或同步完成后重试。");
            }
            catch(IOException error){throw new IOException("无法读取稳定的工作簿副本，请等待其他Excel进程或OneDrive同步完成后重试："+inputFile+"；"+error.Message,error);}
        }
        private void EnsureHiddenApplication()
        {
            if(ownsApplication)return;
            applicationRefs=new HiddenWorkbookApplication(sourceApp);app=applicationRefs.Application;ownsApplication=true;
            Message("已连接独立隐藏表格进程；临时工作簿不会显示在操作窗口中。");
        }
        private void ReleaseHiddenApplication()
        {
            if(!ownsApplication)return;
            try{if(applicationRefs!=null)applicationRefs.Dispose();}
            finally{ownsApplication=false;app=sourceApp;applicationRefs=null;}
        }
        private void Advance()
        {
            if (inTick || queue == null) return; inTick = true; timer.Stop();
            try
            {
                if (stopping) { Cleanup(); End(); return; }
                if (position >= queue.Count) { End(); return; }
                try
                {
                    if (worker == null)
                    {
                        currentProgress=Path.GetFileName(Convert.ToString(queue[position].Tag))+"｜预检中";queue[position].Cells[2].Value = "预检中"; OpenCurrent();
                        if (writing) { queue[position].Cells[2].Value = "执行中"; return; }
                    }
                    else if (writing)
                    {
                        bool complete=worker.StepBatchFile();
                        currentProgress=Path.GetFileName(Convert.ToString(queue[position].Tag))+"｜"+worker.BatchProgressText;
                        if(!complete)return;
                    }
                    if (writing)
                    {
                        int workerErrorCount=worker.BatchValueErrorCount;
                        worker.FinishBatchFile();
                        VerifyNoTemporaryLinks();
                        using (var quiet = new CompanyWorkbookQuiet(app)) ((dynamic)workBook).Save();
                        string temporary = workingFile, final = Path.Combine(destination, Path.GetFileName(Convert.ToString(queue[position].Tag)));
                        CloseObjects(); File.Move(temporary, final);
                        if(workerErrorCount>0){errorFiles.Add(final);Message("数值错误文件："+final+"（链接已生成并输出）");}else Message("已输出："+final);
                    }
                    queue[position].Cells[2].Value = writing ? "已输出" : "预检通过";
                }
                catch (Exception error) { errorFiles.Add(Convert.ToString(queue[position].Tag));string state=parentChildren!=null&&!currentStarted?"阻止：":"失败：";queue[position].Cells[2].Value = state + error.Message; Message(Convert.ToString(queue[position].Tag) + "｜"+state+error.Message); }
                Cleanup(); position++;
            }
            catch (Exception error) { Message("批处理停止：" + error.Message); End(); }
            finally { inTick = false; if (queue != null) timer.Start(); }
        }
        private void End()
        { timer.Stop();int total=queue==null?position:queue.Count;queue = null;currentProgress="本轮结束：已处理 "+position+" / "+total+"；错误文件 "+errorFiles.Count+" 个。"; ReleaseHiddenApplication(); settings.Enabled = files.Enabled = preview.Enabled = execute.Enabled = true; stop.Enabled = false; Message("本轮结束；成功输出保留，失败原因见文件列表及日志。原文件未修改。"); }
        private void VerifyNoTemporaryLinks()
        {
            object links = ((dynamic)workBook).LinkSources(1);
            var items = links as Array;
            if (items != null) foreach (object link in items)
                if (Convert.ToString(link).IndexOf(Path.GetFileName(workingFile), StringComparison.OrdinalIgnoreCase) >= 0 || Convert.ToString(link).IndexOf(scratch, StringComparison.OrdinalIgnoreCase) >= 0)
                    throw new InvalidOperationException("输出包含指向临时副本的外链，已阻止保存，请检查模板公式。");
        }
        private void CloseObjects()
        {
            if (worker != null) { worker.Dispose(); worker = null; }
            if (workerForm != null) { workerForm.Dispose(); workerForm = null; }
            if (workBook != null) { using (var quiet = new CompanyWorkbookQuiet(app)) { quiet.ExcludeClosingBook(workBook); ((dynamic)workBook).Close(false); } workBook = null; }
            if (refs != null) { refs.Dispose(); refs = null; }
        }
        private void Cleanup()
        {
            CloseObjects();
            if (workingFile != null && File.Exists(workingFile)) File.Delete(workingFile);
            if (scratch != null && Directory.Exists(scratch)) Directory.Delete(scratch, false);
            workingFile = scratch = null;
        }
        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            refreshTimer.Stop();
            refreshTimer.Dispose();
            timer.Stop();
            timer.Dispose();
            try { Cleanup(); }
            finally
            {
                try { ReleaseHiddenApplication(); }
                finally { dialog.Dispose(); }
            }
        }
    }
}
