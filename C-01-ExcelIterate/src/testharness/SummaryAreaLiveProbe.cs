using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using Excel_Iterate;

// Explicit live test only. All documents are created and owned by ProbeSession.
internal static class SingleSummaryLiveProbe
{
    private static int checks;
    private static void Check(bool value,string name){if(!value)throw new Exception(name);checks++;Console.WriteLine("PASS "+name);}
    private static readonly BindingFlags Flags=BindingFlags.Instance|BindingFlags.NonPublic;
    private static void Set(object host,string name,object value){host.GetType().GetField(name,Flags).SetValue(host,value);}
    private static object Get(object host,string name){return host.GetType().GetField(name,Flags).GetValue(host);}
    private static void Call(object host,string name,params object[] args){try{host.GetType().GetMethod(name,Flags).Invoke(host,args);}catch(TargetInvocationException e){throw e.InnerException;}}
    private static void ConsolidatedTemplates(object application,ProbeSession session)
    {
        string path=session.Output("合并测试模板.xlsx");
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=bag.Own(books.Add());bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item(1));lead.Name="Lead";
            ((dynamic)bag.Own(lead.Range["K1:O1"])).Value2=new object[,]{{"编码","简称","公司全称","文件名称","文件名称"}};
            ((dynamic)bag.Own(lead.Range["K2:O9"])).Value2="旧公司";
            ((dynamic)bag.Own(lead.Range["A5"])).Value2="科目保留";((dynamic)bag.Own(lead.Range["I2"])).Value2=45000;
            ((dynamic)bag.Own(lead.Range["J2"])).Value2="2027年度";((dynamic)bag.Own(lead.Range["A6"])).Formula="=COUNTA(Lead!M3:M20)";
            book.SaveAs(path,51);book.Close(false);closed=true;
        }
        byte[] before=System.IO.File.ReadAllBytes(path);
        var tree=new List<CompanyStructureTask>{
            new CompanyStructureTask{Row=3,Sequence="00-",Level="0级",FullName="总公司",FileName="00-总公司（合并）.xlsx"},
            new CompanyStructureTask{Row=4,Sequence="02-",Level="1级",FullName="子集团",ParentFullName="总公司",FileName="02-子集团（合并）.xlsx"},
            new CompanyStructureTask{Row=5,Sequence="01-",Level="1级",FullName="总公司",ParentFullName="母公司本部",FileName="01-总公司.xlsx"},
            new CompanyStructureTask{Row=6,Sequence="02-01-",Level="2级",FullName="子集团",ParentFullName="母公司本部",FileName="02-01-子集团.xlsx"}
        };
        var direct=CompanyStructureTasks.DirectChildren(tree,tree[0]);Check(direct.Count==2&&direct[0].Sequence=="01-"&&direct[1].Sequence=="02-","body first, all direct companies included, grandchildren excluded");
        tree[0].Level="一级";
        Check(CompanyStructureTasks.DirectChildren(tree,tree[1]).Count==1,"unrecognized root level does not require top-level row to have a parent");
        var noRoot=new List<CompanyStructureTask>{tree[1],tree[3],new CompanyStructureTask{Row=99,Sequence="99-01-",FullName="无关公司"}};
        Check(CompanyStructureTasks.DirectChildren(noRoot,tree[1]).Count==1,"unrelated missing parent cannot block selected group's valid children");
        tree[3].ParentFullName="错误归属";bool conflict=false;try{CompanyStructureTasks.DirectChildren(tree,tree[1]);}catch(InvalidOperationException){conflict=true;}Check(conflict,"selected group's own child parent conflict still blocks");tree[3].ParentFullName="母公司本部";
        tree[0].Level="0级";CompanyStructureTasks.Classify(tree,true);
        using(var picker=new CompanyStructurePicker{SelectParents=true}){picker.SetTasks(tree);Check(picker.Selected().Count==2,"consolidated selector defaults to parent companies");}
        using(var host=CompanyWorkbookHost.FromOpenOrDisk(application,path))
        {
            var profile=ConsolidatedTemplateProfile.Read(host.Snapshot);Check(profile.CapacityLast==20&&profile.OldLast==9,"profile detects template company capacity and old tail");
            var plan=profile.Plan(tree,tree[0],session.DirectoryPath);Check(!System.IO.File.Exists(plan.Record.OutputPath),"preflight creates no output");
            profile.CapacityLast=3;bool blocked=false;try{profile.Plan(tree,tree[0],session.DirectoryPath);}catch(InvalidOperationException){blocked=true;}Check(blocked,"capacity overflow blocked");profile.CapacityLast=20;
            host.Generate(new CompanyWorkbookRequest{Output=session.DirectoryPath,UpdateLead=true},plan.Record,delegate(object book,CompanyWorkbookRequest r,CompanyWorkbookRecord record){ConsolidatedLeadWriter.Write(book,plan);},null,delegate(object book,CompanyWorkbookRequest r,CompanyWorkbookRecord record){ConsolidatedLeadWriter.Verify(book,plan);});
            Check(System.IO.File.Exists(plan.Record.OutputPath),"output published only after native save and reopen verification");
            var failedPlan=profile.Plan(tree,tree[1],session.DirectoryPath);bool injected=false;
            try{host.Generate(new CompanyWorkbookRequest{Output=session.DirectoryPath,UpdateLead=true},failedPlan.Record,delegate(object book,CompanyWorkbookRequest r,CompanyWorkbookRecord record){ConsolidatedLeadWriter.Write(book,failedPlan);throw new InvalidOperationException("test failure after private-copy write");});}catch(InvalidOperationException){injected=true;}
            Check(injected&&!System.IO.File.Exists(failedPlan.Record.OutputPath)&&System.IO.Directory.GetFiles(session.DirectoryPath,".company-*.xlsx").Length==0,"failed generation publishes nothing and removes private staging copy");
            blocked=false;try{profile.Plan(tree,tree[0],session.DirectoryPath);}catch(InvalidOperationException){blocked=true;}Check(blocked,"existing output never overwritten");
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)application).Workbooks),book=bag.Own(books.Open(plan.Record.OutputPath,0,true));bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
                dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item("Lead"));
                Check(Convert.ToString(((dynamic)bag.Own(lead.Range["M2"])).Value2)=="总公司（合并）"&&Convert.ToString(((dynamic)bag.Own(lead.Range["K3"])).Value2)=="01-"&&Convert.ToString(((dynamic)bag.Own(lead.Range["M4"])).Value2)=="子集团（合并）","parent/body/direct-company mapping survives reopening");
                Check(Convert.ToString(((dynamic)bag.Own(lead.Range["L4"])).Formula)=="=K4&M4"&&Convert.ToString(((dynamic)bag.Own(lead.Range["N4"])).Formula)=="=K4&O4","concat formulas retained");
                Check(Convert.ToString(((dynamic)bag.Own(lead.Range["A5"])).Value2)=="科目保留"&&Convert.ToDouble(((dynamic)bag.Own(lead.Range["I2"])).Value2)==45000&&Convert.ToString(((dynamic)bag.Own(lead.Range["J2"])).Value2)=="2027年度","accounting subjects and years untouched");
                Check(((dynamic)bag.Own(lead.Range["M9"])).Value2==null,"stale company tail cleared");book.Close(false);closed=true;
            }
        }
        Check(System.Linq.Enumerable.SequenceEqual(before,System.IO.File.ReadAllBytes(path)),"original template unchanged byte for byte");
        var broken=new List<CompanyStructureTask>{tree[0],tree[1],tree[3]};bool missing=false;try{CompanyStructureTasks.DirectChildren(broken,tree[0]);}catch(InvalidOperationException){missing=true;}Check(missing,"missing body blocks parent");
    }
    private static void StructureTemplates(object application,ProbeSession session)
    {
        string sourcePath=session.Output("架构来源.xlsx"),templatePath=session.Output("仅Lead模板.xlsx");
        dynamic architecture;bool architectureClosed=false;
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),template=books.Add();bool closed=false;session.Own((object)template,delegate{if(!closed)template.Close(false);});
            dynamic sheets=bag.Own(template.Worksheets),lead=bag.Own(sheets.Item(1));lead.Name="Lead";
            dynamic headers=bag.Own(lead.Range["A1:B2"]),sentinel=bag.Own(lead.Range["D5"]);headers.Value2=new object[,]{{"简称","全称"},{"原简称","原全称"}};sentinel.Value2=123;
            template.SaveAs(templatePath,51);template.Close(false);closed=true;
            architecture=books.Add();session.Own((object)architecture,delegate{if(!architectureClosed)architecture.Close(false);});
            dynamic architectureSheets=bag.Own(architecture.Worksheets),sheet=bag.Own(architectureSheets.Item(1));sheet.Name="架构";
            dynamic header=bag.Own(sheet.Range["A1:G1"]);header.Value2=new object[,]{{"序号","合并集团简称","简称","级次","公司名称","并表母公司","过程表文件名称"}};
            dynamic rows=bag.Own(sheet.Range["A3:G6"]);rows.Value2=new object[,]{{"00-","集团展示名","总简称","0级","总全称",null,"00-总（合并）.xlsx"},{"03-","分组展示名","组简称","1级","组全称",null,"03-组（合并）.xlsx"},{"03-01-",null,"末级简称","2级","末级全称",null,"03-01-末级.xlsx"},{"03-02-",null,null,"2级","缺简称全称",null,"03-02-缺简称.xlsx"}};
            architecture.SaveAs(sourcePath,51);dynamic code=bag.Own(sheet.Range["C5"]);code.Value2="未保存简称";
        }
        using(var source=CompanyWorkbookHost.FromOpenOrDisk(application,sourcePath))
        using(var template=CompanyWorkbookHost.FromOpenOrDisk(application,templatePath))
        {
            var tasks=source.StructureRows("架构",new CompanyStructureColumns());CompanyStructureTasks.Classify(tasks,true);
            Check(!tasks[0].Leaf&&!tasks[1].Leaf&&tasks[2].Leaf&&tasks[3].Leaf,"native structure leaf classification includes total-group level");
            Check(tasks[2].ShortName=="未保存简称"&&!Convert.ToBoolean(architecture.Saved),"architecture snapshots include unsaved edits without saving original");
            var selected=CompanyStructureTasks.Records(new[]{tasks[2],tasks[3]});var request=new CompanyWorkbookRequest{Output=session.DirectoryPath,TargetSheet="Lead",TargetHeaderRow=1,TargetRow=2,TargetCodeField="简称",TargetNameField="全称",UpdateLead=true};
            int shortBlocks=0;var planned=template.PreviewStructure(request,selected,false,false,delegate(CompanyWorkbookRecord row,string reason){if(row.Row==6&&reason.Contains("简称为空"))shortBlocks++;});Check(planned.Count==1&&shortBlocks==1,"native preflight blocks only missing-short-name item");Check(!System.IO.File.Exists(planned[0].OutputPath),"architecture template preflight creates no output");
            template.Generate(request,planned[0],CompanyWorkbookHost.UpdateLead);
            Check(!System.IO.File.Exists(System.IO.Path.Combine(session.DirectoryPath,tasks[3].FileName)),"blocked architecture task never produces output");
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)application).Workbooks),book=bag.Own(books.Open(planned[0].OutputPath,0,true));
                try{dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item("Lead")),name=bag.Own(lead.Range["A2"]),full=bag.Own(lead.Range["B2"]),sentinel=bag.Own(lead.Range["D5"]);Check(Convert.ToString(name.Value2)=="未保存简称"&&Convert.ToString(full.Value2)=="末级全称","architecture fields write directly without source summary sheet");Check(Convert.ToDouble(sentinel.Value2)==123,"template unrelated content preserved");}
                finally{book.Close(false);}
            }
            bool blocked=false;try{template.PreviewStructure(request,selected,false,false);}catch(InvalidOperationException){blocked=true;}Check(blocked,"architecture generation never overwrites existing output");
            Check(!Convert.ToBoolean(architecture.Saved),"original architecture remains open and unsaved after generation");
        }
        using(var picker=new CompanyStructurePicker{SimpleMode=true,OpenSource=delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(application,path);}})
        {
            picker.LoadSource(sourcePath);var tasks=picker.Selected();Check(tasks.Count==2&&tasks[0].ShortName=="未保存简称"&&tasks[0].FileName=="03-01-末级.xlsx","simple picker reads C short names and G filenames without using B group labels");
        }
        string previousRoot=DialogSizeStore.RootOverride;DialogSizeStore.RootOverride=session.DirectoryPath;
        try
        {
            using(var form=new CompanyWorkbookForm())
            {
                form.OpenTemplate=delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(application,path);};form.RememberConfigurationEnabled=true;form.StructureMode.Checked=true;Call(form,"LoadTemplate",templatePath);
                ((ComboBox)Get(form,"targetCode")).Text="简称";((ComboBox)Get(form,"targetName")).Text="全称";((NumericUpDown)Get(form,"leadRow")).Value=2;((TextBox)Get(form,"output")).Text=session.DirectoryPath;form.RememberConfiguration();
                Check(System.IO.File.Exists(CompanyWorkbookForm.LastConfigurationFile),"last valid structure settings saved separately from preset library");
                var legacy=form.CapturePreset();legacy.Name="旧模板配置";legacy.Values.Remove("structureMode");legacy.Values["source"]="不存在的公司汇总表";form.ApplyPreset(legacy);
                Check(form.StructureMode.Checked&&((ComboBox)Get(form,"targetCode")).Text=="简称","explicit structure mode reuses legacy target settings without requiring old source summary sheet");
            }
            using(var restored=new CompanyWorkbookForm())
            {
                restored.OpenTemplate=delegate(string path){return CompanyWorkbookHost.FromOpenOrDisk(application,path);};restored.RememberConfigurationEnabled=true;restored.RestoreLastConfiguration();
                Check(restored.StructureMode.Checked&&Get(restored,"host")!=null&&((ComboBox)Get(restored,"targetCode")).Text=="简称"&&((NumericUpDown)Get(restored,"leadRow")).Value==2&&((TextBox)Get(restored,"output")).Text==session.DirectoryPath,"next session restores template output and write locations without reconfiguration");
                Check(restored.Structure.Selected().Count==0,"restoring settings does not select companies or execute generation");
            }
        }
        finally{DialogSizeStore.RootOverride=previousRoot;}
        architecture.Close(false);architectureClosed=true;
    }
    private static void TemplateEntryWindows(object application)
    {
        var connectType = typeof(CompanyWorkbookForm).Assembly.GetType("Excel_Iterate.Connect", true);
        object connect = Activator.CreateInstance(connectType);
        var host = connectType.GetField("_app", BindingFlags.Static | BindingFlags.NonPublic);
        var singleEntry = connectType.GetMethod("OnCompanyWorkbooks");
        var reportEntry = connectType.GetMethod("OnConsolidatedWorkbooks");
        host.SetValue(null, application);
        Form single = null, report = null;
        try
        {
            singleEntry.Invoke(connect, new object[] { null });
            foreach (Form form in Application.OpenForms)
                if (form is CompanyWorkbookForm) single = form;
            Check(single is CompanyWorkbookForm && single.Text == "单体模板生成", "single template callback opens its own form");
            reportEntry.Invoke(connect, new object[] { null });
            foreach (Form form in Application.OpenForms)
                if (form is ConsolidatedWorkbookForm) report = form;
            Check(report is ConsolidatedWorkbookForm && report.Text == "报表模板生成", "report template callback opens a different form");
            single.Close();
            Check(single.IsDisposed && report.Visible && !report.IsDisposed, "closing the single template leaves report template open");
        }
        finally
        {
            if (single != null && !single.IsDisposed) single.Close();
            if (report != null && !report.IsDisposed) report.Close();
            host.SetValue(null, null);
        }
    }

    private static void FileBatchWindow(object application)
    {
        using (var workbookWindow = new Form { Text = "自建工作簿窗口样本" })
        using (var launcher = new Form { Text = "合并批量链接测试入口" })
        {
            workbookWindow.Show();
            launcher.Show(workbookWindow);
            SummaryFileBatch.Show(application, launcher);
            Form batch = null;
            foreach (Form open in Application.OpenForms)
                if (open.Text == "合并批量链接 · 多文件") batch = open;
            Check(batch != null, "batch opens modelessly from the launcher");
            try
            {
                Check(!batch.Modal && batch.Owner == null && batch.ShowInTaskbar
                    && batch.MinimizeBox && workbookWindow.Enabled, "batch is independently restorable and workbook stays enabled");
                batch.WindowState = FormWindowState.Minimized;
                Application.DoEvents();
                Check(batch.WindowState == FormWindowState.Minimized && workbookWindow.Visible
                    && workbookWindow.WindowState == FormWindowState.Normal && workbookWindow.Enabled,
                    "minimizing the batch leaves the workbook window accessible");
                batch.WindowState = FormWindowState.Normal;
                Application.DoEvents();
                Check(batch.Visible, "batch can be restored without reopening");
            }
            finally { batch.Close(); }
            Check(batch.IsDisposed && launcher.Enabled, "closing the batch releases it and enables the launcher again");
        }

        using (var batch = (SummaryFileBatch)Activator.CreateInstance(typeof(SummaryFileBatch), Flags, null, new[] { application }, null))
        {
            var dialog = (Form)Get(batch, "dialog");
            dialog.Show();
            foreach (int width in new[] { dialog.MinimumSize.Width, dialog.Width })
            {
                dialog.Width = width;
                dialog.PerformLayout();
                var pathSheet = (TextBox)Get(batch, "pathSheet");
                var pathCell = (TextBox)Get(batch, "pathCell");
                var templateName = (TextBox)Get(batch, "templateName");
                var templateArea = (TextBox)Get(batch, "templateArea");
                var pasteStart = (TextBox)Get(batch, "pasteStart");
                var leadName = (TextBox)Get(batch, "leadName");
                var leadArea = (TextBox)Get(batch, "leadArea");
                Func<Control, System.Drawing.Point> origin = control => control.PointToScreen(System.Drawing.Point.Empty);
                Check(origin(pathSheet).X == origin(templateName).X && origin(templateName).X == origin(leadName).X
                    && origin(pathCell).X == origin(templateArea).X && origin(templateArea).X == origin(leadArea).X,
                    "table and address fields share alignment at width " + width);
                Check(origin(templateName).Y - origin(pathSheet).Y == origin(leadName).Y - origin(templateName).Y
                    && origin(templateName).Y == origin(templateArea).Y && origin(templateArea).Y == origin(pasteStart).Y
                    && Math.Abs(templateName.Width - pasteStart.Width) <= 2 && Math.Abs(templateArea.Width - pasteStart.Width) <= 2,
                    "parameter fields use even widths and spacing at width " + width);
                var files = (DataGridView)Get(batch, "files");
                Check(files.Visible && files.Height > files.ColumnHeadersHeight + files.RowTemplate.Height,
                    "file checklist is visible with room for at least one result");
            }
        }
    }
    private static void FileBatch(object application, ProbeSession session)
    {
        string input = System.IO.Path.Combine(session.DirectoryPath, "inputs"), output = System.IO.Path.Combine(session.DirectoryPath, "outputs");
        System.IO.Directory.CreateDirectory(input); System.IO.Directory.CreateDirectory(output);
        var selected = new List<string>(); var hashes = new Dictionary<string, byte[]>();
        using (var bag = new ComBag())
        {
            dynamic books = bag.Own(((dynamic)application).Workbooks);
            for (int i = 0; i < 3; i++)
            {
                string file = System.IO.Path.Combine(input, "任务" + i + ".xlsx");
                dynamic book = books.Add(); bool closed = false; session.Own((object)book, delegate { if (!closed) book.Close(false); });
                dynamic sheets = bag.Own(book.Worksheets), sheet = bag.Own(sheets.Item(1)), lead = bag.Own(sheets.Add()); sheet.Name = "汇总"; lead.Name = "Lead";
                dynamic path = bag.Own(sheet.Range["A1"]), source = bag.Own(sheet.Range["A3"]), company = bag.Own(lead.Range["L3"]), variable = bag.Own(lead.Range["M3"]), period = bag.Own(lead.Range["J2"]);
                string ownSources=System.IO.Path.Combine(session.DirectoryPath,"sources"+i);System.IO.Directory.CreateDirectory(ownSources);
                string companyName=i==0?"！01-示例（集团）":"！02-示例乙";
                System.IO.File.Copy(System.IO.Path.Combine(session.DirectoryPath,companyName+".xlsx"),System.IO.Path.Combine(ownSources,companyName+".xlsx"));
                path.Value2 = i == 1 ? "" : ownSources; source.Value2 = "BS!$A$1"; company.Value2 = companyName; variable.Value2 = "变量" + i; period.Value2 = i==0?"2027年度":"2026年度";
                if(i==2)
                {
                    dynamic extraCompany=bag.Own(lead.Range["L4"]),extraVariable=bag.Own(lead.Range["M4"]),untouched=bag.Own(sheet.Range["D3"]);extraCompany.Value2="！01-示例（集团）";extraVariable.Value2="无关公司变量";untouched.Value2="保留未选公司区域";
                    System.IO.File.Copy(System.IO.Path.Combine(session.DirectoryPath,"！01-示例（集团）.xlsx"),System.IO.Path.Combine(ownSources,"！01-示例（集团）.xlsx"));
                }
                book.SaveAs(file, 51); book.Close(false); closed = true; selected.Add(file); hashes[file] = System.IO.File.ReadAllBytes(file);
            }
        }
        var preset = new BatchPreset { Feature = "summary", Context = "旧文件.xlsx", Name = "通用模板", Values = new Dictionary<string,string> {
            {"mode","area-mvp"},{"templateSheet","汇总"},{"templateAreas","3,1,1,1"},{"pasteSheet","汇总"},{"pasteArea","3,3,1,1"},{"leadSheet","Lead"},{"leadAreas","3,12,1,1"},{"projectFolder","旧路径"},{"checkFiles","False"},
            {"headersEnabled","True"},{"headersFormula","True"},{"headersRepeat","True"},{"variableHeaders","True"},{"repeatSource","3,13,1,1"},{"repeatSourceSheet","Lead"},{"repeatTarget","2,3,1,1"},{"repeatTargetSheet","汇总"},{"yearSource","2,10,1,1"},{"yearSourceSheet","Lead"},{"yearTarget","4,3,1,1"},{"yearTargetSheet","汇总"}
        } };
        dynamic openOriginal;
        using(var bag=new ComBag()) {dynamic books=bag.Own(((dynamic)application).Workbooks);openOriginal=books.Open(selected[0],0,false);}
        bool openOriginalClosed=false;session.Own((object)openOriginal,delegate{if(!openOriginalClosed)openOriginal.Close(false);});
        using(var bag=new ComBag()){dynamic sheets=bag.Own(openOriginal.Worksheets),lead=bag.Own(sheets.Item("Lead")),variable=bag.Own(lead.Range["M3"]);variable.Value2="未保存变量";}
        Func<bool, SummaryFileBatch> create = delegate(bool save) {
            var batch = (SummaryFileBatch)Activator.CreateInstance(typeof(SummaryFileBatch), Flags, null, new object[] { application }, null);
            var grid = (DataGridView)Get(batch,"files"); var rows = new List<DataGridViewRow>();
            foreach (string file in selected) { int n = grid.Rows.Add(true, System.IO.Path.GetFileName(file), "等待"); grid.Rows[n].Tag = file; rows.Add(grid.Rows[n]); }
            Set(batch,"preset",preset); Set(batch,"destination",output); Set(batch,"sourceSheet","汇总"); Set(batch,"sourceCell","A1"); Set(batch,"writing",save); Set(batch,"queue",rows);
            ((ProgressBar)Get(batch,"progress")).Maximum = rows.Count;
            return batch;
        };
        using (var batch = create(false))
        {
            int guard = 0; while(Get(batch,"queue") != null && guard++ < 100) Call(batch,"Advance");
            Check(Get(batch,"queue")==null,"batch preflight terminates");
            var grid=(DataGridView)Get(batch,"files");Check(Convert.ToString(grid.Rows[0].Cells[2].Value)=="预检通过"&&Convert.ToString(grid.Rows[1].Cells[2].Value).StartsWith("失败")&&Convert.ToString(grid.Rows[2].Cells[2].Value)=="预检通过","batch preflight isolates invalid A1 and continues");
            Check(System.IO.Directory.GetFiles(output).Length==0,"batch preflight never outputs workbook");
        }
        using (var batch = create(true))
        {
            bool sawWorker=false,stayedHidden=true;
            int guard = 0; while(Get(batch,"queue") != null && guard++ < 100)
            {
                Call(batch,"Advance");
                if(Convert.ToBoolean(Get(batch,"ownsApplication")))
                {sawWorker=true;stayedHidden=stayedHidden&&!Convert.ToBoolean(((dynamic)Get(batch,"app")).Visible);}
            }
            Check(sawWorker&&stayedHidden,"dedicated batch application stays invisible through backup and every area write");
            Check(!Convert.ToBoolean(Get(batch,"ownsApplication"))&&Get(batch,"applicationRefs")==null,"dedicated application quits and releases COM ownership at batch end");
            Check(Get(batch,"queue")==null,"batch execution terminates");
            var grid=(DataGridView)Get(batch,"files");for(int i=0;i<3;i++)Console.WriteLine("BATCH STATUS "+i+": "+grid.Rows[i].Cells[2].Value);
            Console.WriteLine(((RichTextBox)Get(batch,"log")).Text);
            Check(System.IO.File.Exists(System.IO.Path.Combine(output,"任务0.xlsx")) && !System.IO.File.Exists(System.IO.Path.Combine(output,"任务1.xlsx")) && System.IO.File.Exists(System.IO.Path.Combine(output,"任务2.xlsx")),"batch saves successful files and skips invalid file");
        }
        Check(!Convert.ToBoolean(openOriginal.Saved),"open original remains unsaved after batch");
        using(var bag=new ComBag()){dynamic sheets=bag.Own(openOriginal.Worksheets),lead=bag.Own(sheets.Item("Lead")),variable=bag.Own(lead.Range["M3"]);Check(Convert.ToString(variable.Value2)=="未保存变量","open original remains available with unsaved edits");}
        openOriginal.Close(false);openOriginalClosed=true;
        foreach(string file in selected) Check(System.Linq.Enumerable.SequenceEqual(hashes[file],System.IO.File.ReadAllBytes(file)),"original bytes unchanged "+System.IO.Path.GetFileName(file));
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks);
            foreach(int i in new[]{0,2})
            {
                dynamic book=bag.Own(books.Open(System.IO.Path.Combine(output,"任务"+i+".xlsx"),0,true));
                try {dynamic sheets=bag.Own(book.Worksheets),sheet=bag.Own(sheets.Item("汇总")),cell=bag.Own(sheet.Range["C3"]),variable=bag.Own(sheet.Range["C2"]),year=bag.Own(sheet.Range["C4"]);Check(Convert.ToString(cell.Formula).Contains(i==0?"！01-示例（集团）":"！02-示例乙"),"each file uses its own Lead filename");Check(Convert.ToString(cell.Formula).Contains("sources"+i),"each file uses its own A1 folder");Check(!Convert.ToString(cell.Formula).Contains(".xybatch-"),"output links do not point to temporary workbook");Check(Convert.ToString(variable.Value2)==(i==0?"未保存变量":"变量"+i)&&Convert.ToString(year.Value2)==(i==0?"2027年度":"2026年度"),"each file maps its own variable and period including unsaved original");}
                finally{book.Close(false);}
            }
        }
        byte[] prior=System.IO.File.ReadAllBytes(System.IO.Path.Combine(output,"任务0.xlsx"));
        using(var batch=create(true)){int guard=0;while(Get(batch,"queue")!=null&&guard++<100)Call(batch,"Advance");Check(Convert.ToString(((DataGridView)Get(batch,"files")).Rows[0].Cells[2].Value).Contains("已存在"),"existing output blocks overwrite");}
        Check(System.Linq.Enumerable.SequenceEqual(prior,System.IO.File.ReadAllBytes(System.IO.Path.Combine(output,"任务0.xlsx"))),"existing output bytes preserved");
        string cancelled=System.IO.Path.Combine(session.DirectoryPath,"cancelled");System.IO.Directory.CreateDirectory(cancelled);
        using(var batch=create(true))
        {
            Set(batch,"destination",cancelled);Call(batch,"Advance");Set(batch,"stopping",true);Call(batch,"Advance");
            Check(Get(batch,"queue")==null&&System.IO.Directory.GetFileSystemEntries(cancelled).Length==0,"stopping discards in-progress copy without output");
        }
        string parentOutput=System.IO.Path.Combine(session.DirectoryPath,"parents");System.IO.Directory.CreateDirectory(parentOutput);
        using(var batch=create(true))
        {
            var rows=(List<DataGridViewRow>)Get(batch,"queue");rows.RemoveAt(1);
            Set(batch,"destination",parentOutput);
            var parentPreset=BatchPresetStore.Decode(BatchPresetStore.Encode(preset),"summary");parentPreset.Values["leadAreas"]="3,12,2,1";parentPreset.Values["repeatSource"]="3,13,2,1";Set(batch,"preset",parentPreset);
            Set(batch,"parentChildren",new Dictionary<string,List<CompanyStructureTask>>(StringComparer.OrdinalIgnoreCase){
                {selected[0],new List<CompanyStructureTask>{new CompanyStructureTask{Sequence="03-01-",FileName="！01-示例（集团）.xlsx"},new CompanyStructureTask{Sequence="03-02-",FileName="缺少的末级公司.xlsx"}}},
                {selected[2],new List<CompanyStructureTask>{new CompanyStructureTask{Sequence="04-01-",FileName="！02-示例乙.xlsx"}}}
            });
            int guard=0;while(Get(batch,"queue")!=null&&guard++<100)Call(batch,"Advance");
            var grid=(DataGridView)Get(batch,"files");string blocked=Convert.ToString(grid.Rows[0].Cells[2].Value);
            Check(blocked.StartsWith("阻止")&&blocked.Contains("Lead缺少"),"missing one required leaf blocks whole parent before writing");
            Check(!System.IO.File.Exists(System.IO.Path.Combine(parentOutput,"任务0.xlsx"))&&System.IO.File.Exists(System.IO.Path.Combine(parentOutput,"任务2.xlsx")),"other valid parent continues and blocked parent has no output");
        }
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=bag.Own(books.Open(System.IO.Path.Combine(parentOutput,"任务2.xlsx"),0,true));
            try
            {
                dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item("Lead")),company=bag.Own(lead.Range["L3"]),variable=bag.Own(lead.Range["M3"]),period=bag.Own(lead.Range["J2"]),sheet=bag.Own(sheets.Item("汇总")),formula=bag.Own(sheet.Range["C3"]),untouched=bag.Own(sheet.Range["D3"]);
                Check(Convert.ToString(company.Value2)=="！02-示例乙"&&Convert.ToString(variable.Value2)=="变量2"&&Convert.ToString(period.Value2)=="2026年度","parent linking keeps original Lead company variable and period values");
                Check(Convert.ToString(formula.Formula).Contains("[！02-示例乙.xlsx]BS"),"parent output contains real external link for matched leaf");
                Check(Convert.ToString(untouched.Value2)=="保留未选公司区域","unrelated Lead companies are not linked and their target cells remain untouched");
            }
            finally{book.Close(false);}
        }
        foreach(string file in selected)Check(System.Linq.Enumerable.SequenceEqual(hashes[file],System.IO.File.ReadAllBytes(file)),"parent workflow never saves original "+System.IO.Path.GetFileName(file));
    }
    private static void ProvidedHeaders(object application,ProbeSession session,string vectorsPath)
    {
        List<ProvidedHeaderVector> samples;
        using(var stream=System.IO.File.OpenRead(vectorsPath))samples=(List<ProvidedHeaderVector>)new System.Runtime.Serialization.Json.DataContractJsonSerializer(typeof(List<ProvidedHeaderVector>)).ReadObject(stream);
        int number=0;
        foreach(var sample in samples)
        {
            number++;using(var bag=new ComBag())
            {
                dynamic app=application,books=bag.Own(app.Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
                dynamic sheets=bag.Own(book.Worksheets),sheet=bag.Own(sheets.Item(1));sheet.Name="汇总";dynamic source=bag.Own(sheets.Add());source.Name=sample.sheet;
                foreach(var year in sample.years){dynamic cell=bag.Own(source.Range["N"+year.row]);cell.Value2=year.text;}
                foreach(var company in sample.companies){dynamic code=bag.Own(source.Range["O"+company.row]),name=bag.Own(source.Range["P"+company.row]);code.Value2=company.code;name.Value2=company.name;}
                int first=sample.companies[0].row,last=sample.companies[sample.companies.Count-1].row,width=sample.years.Count,oldGroups=Math.Min(3,sample.companies.Count);
                dynamic oldNames=bag.Own(sheet.Range["L2:"+CellAddress.Name(2,12+oldGroups*width-1)]);oldNames.FormulaArray="=INDEX('"+sample.sheet+"'!$O$"+first+":$O$"+(first+oldGroups-1)+",INT((COLUMN("+"L2:"+CellAddress.Name(2,12+oldGroups*width-1)+")-COLUMN(L2))/"+width+")+1)";oldNames.Calculate();string oldFormula=Convert.ToString(((dynamic)bag.Own(sheet.Range["L2"])).Formula);
                using(var form=new SingleSummaryForm(true))
                {
                    var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.VariableHeaders.Checked=false;view.HeadersEnabled.Checked=true;view.FormulaHeaders.Checked=true;view.RepeatNames.Checked=true;view.ReplaceHeaderArrays.Checked=true;
                    using(var host=new SummaryAreaHost(application,form,view,null,delegate{return true;}))
                    {
                        Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet",sample.sheet);
                        ((List<SummaryAreaPlan>)Get(host,"areas")).Add(new SummaryAreaPlan{Sheet="汇总",Row=10,Column=7,Rows=1,Columns=width});Set(host,"paste",new SummaryAreaPlan{Row=10,Column=12,Rows=1,Columns=1});
                        Set(host,"lead",new SummaryAreaPlan{Sheet=sample.sheet,Row=first,Column=15,Rows=last-first+1,Columns=2});Set(host,"companyHeader",new SummaryAreaPlan{Sheet="汇总",Row=2,Column=12,Rows=1,Columns=1});Set(host,"yearSource",new SummaryAreaPlan{Sheet=sample.sheet,Row=sample.years[0].row,Column=14,Rows=width,Columns=1});Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=7,Column=12,Rows=1,Columns=1});
                        Call(host,"Preview");Check((bool)Get(host,"ready"),"provided sample "+number+" actual N/O/P fields preflight");Call(host,"Start");for(int i=0;i<sample.companies.Count;i++)Call(host,"Tick",null,EventArgs.Empty);
                        bool correct=!(bool)Get(host,"failed")&&(bool)Get(host,"pending");for(int i=0;i<sample.companies.Count;i++)for(int c=0;c<width;c++){dynamic name=bag.Own(sheet.Range[CellAddress.Name(2,12+i*width+c)]),year=bag.Own(sheet.Range[CellAddress.Name(7,12+i*width+c)]);correct &= Convert.ToBoolean(name.HasFormula)&&Convert.ToString(name.Value2)==sample.companies[i].code&&Convert.ToBoolean(year.HasFormula)&&Convert.ToString(year.Value2)==sample.years[c].text;}
                        Check(correct,"provided sample "+number+" company order repetition and transposed years");Call(host,"Undo");Check(Convert.ToBoolean(oldNames.HasArray)&&Convert.ToString(((dynamic)bag.Own(sheet.Range["L2"])).Formula)==oldFormula,"provided sample "+number+" original array rollback");Check(Convert.ToString(book.Path)=="","provided sample "+number+" isolated workbook never saved");
                    }
                }
                book.Close(false);closed=true;
            }
        }
    }
    private static void VariableHeaders(object application,ProbeSession session)
    {
        using(var bag=new ComBag()){
            dynamic app=application,books=bag.Own(app.Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            dynamic sheets=bag.Own(book.Worksheets),sheet=bag.Own(sheets.Item(1)),lead=bag.Own(sheets.Add());sheet.Name="变量汇总";lead.Name="Lead";
            dynamic extraData=bag.Own(sheet.Range["G12:H12"]);extraData.Value2=new object[,]{{1234567.89,12345.67}};dynamic percentRow=bag.Own(sheet.Range["G13:H13"]);percentRow.Value2=0.125;percentRow.NumberFormat="0.0%";dynamic untouched=bag.Own(sheet.Range["Q11"]);untouched.Value2=0.123;untouched.NumberFormat="0.0%";
            ((dynamic)bag.Own(sheet.Range["G14:H14"])).Value2=new object[,]{{0.0,-1234.5}};
            ((dynamic)bag.Own(sheet.Range["G15"])).Formula="=0";
            dynamic textZero=bag.Own(sheet.Range["H15"]);textZero.NumberFormat="@";textZero.Value2="0";
                    dynamic priorZero=bag.Own(sheet.Range["Q14"]);priorZero.NumberFormat="0.0000";priorZero.Value2=999.0;priorZero.HorizontalAlignment=-4152;priorZero.VerticalAlignment=-4160;untouched.HorizontalAlignment=-4152;
            dynamic data=bag.Own(lead.Range["O3:Q4"]);data.Value2=new object[,]{{"甲简称","！01-示例（集团）","01"},{"乙简称","！02-示例乙","02"}};
            dynamic years=bag.Own(lead.Range["N2:N3"]);years.Value2=new object[,]{{"2025年度"},{"2027年度"}};
            dynamic outputColumn=bag.Own(sheet.Range["Q:Q"]);outputColumn.ColumnWidth=5;double originalWidth=Convert.ToDouble(outputColumn.ColumnWidth);
            dynamic src=bag.Own(sheet.Range["G10:H10"]);src.Value2="BS!$A$1";dynamic percentageSource=bag.Own(sheet.Range["H10"]);percentageSource.NumberFormat="0.00%";dynamic hiddenOutput=bag.Own(sheet.Range["R:R"]);bool hiddenSupported=true;try{hiddenOutput.Hidden=true;}catch(System.Runtime.InteropServices.COMException){hiddenSupported=false;Console.WriteLine("SKIP host refused hidden-column fixture setup");}dynamic prior=bag.Own(sheet.Range["Q2"]);prior.Formula="=1+2";dynamic gap=bag.Own(sheet.Range["Q6"]);gap.Value2="未选空隙";
            using(var form=new SingleSummaryForm(true)){
                var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.HeadersEnabled.Checked=true;view.VariableHeaders.Checked=true;
                using(var host=new SummaryAreaHost(application,form,view,null,delegate{return true;})){
                    Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","变量汇总");Set(host,"pasteSheet","变量汇总");Set(host,"leadSheet","Lead");
                    ((List<SummaryAreaPlan>)Get(host,"areas")).Add(new SummaryAreaPlan{Sheet="变量汇总",Row=10,Column=7,Rows=1,Columns=2});((List<SummaryAreaPlan>)Get(host,"areas")).Add(new SummaryAreaPlan{Sheet="变量汇总",Row=12,Column=7,Rows=5,Columns=2});Set(host,"paste",new SummaryAreaPlan{Row=10,Column=17,Rows=1,Columns=1});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=3,Column=16,Rows=2,Columns=1});
                    Set(host,"repeatSource",new SummaryAreaPlan{Sheet="Lead",Row=3,Column=15,Rows=2,Columns=3});Set(host,"repeatTarget",new SummaryAreaPlan{Sheet="变量汇总",Row=2,Column=17,Rows=3,Columns=1});Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=2,Column=14,Rows=2,Columns=1});Set(host,"yearTarget",new SummaryAreaPlan{Sheet="变量汇总",Row=7,Column=17,Rows=1,Columns=1});
                    int before=Convert.ToInt32(sheets.Count);Call(host,"Preview");Check((bool)Get(host,"ready")&&Convert.ToInt32(sheets.Count)==before&&Convert.ToString(prior.Formula)=="=1+2","variable preflight is read-only");
                    Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"variable formula execution succeeds");
                    dynamic numericOutput=bag.Own(sheet.Range["Q10"]);Check(Convert.ToString(numericOutput.NumberFormat)==SummaryAreaHost.AmountNumberFormat,"output data uses thousands and two decimals without replacing formulas");Check(Convert.ToDouble(outputColumn.ColumnWidth)>originalWidth,"output column grows to fit complete header values");
                    Check(Convert.ToInt32(numericOutput.HorizontalAlignment)==-4108&&Convert.ToInt32(numericOutput.VerticalAlignment)==-4108&&Convert.ToInt32(prior.HorizontalAlignment)==-4108&&Convert.ToInt32(prior.VerticalAlignment)==-4108,"generated data and header centered horizontally and vertically");
                    Check(Convert.ToInt32(untouched.HorizontalAlignment)==-4152,"unselected gap alignment unchanged");
                    Check(Convert.ToString(((dynamic)bag.Own(sheet.Range["R10"])).NumberFormat)=="0.00%","percentage data format survives amount formatting");if(hiddenSupported)Check(Convert.ToBoolean(hiddenOutput.Hidden),"autofit does not reveal hidden columns");
                    Check(Convert.ToString(((dynamic)bag.Own(sheet.Range["Q12"])).Text)=="1,234,567.89"&&Convert.ToDouble(untouched.Value2)==0.123&&Convert.ToString(untouched.NumberFormat)=="0.0%"&&Convert.ToString(((dynamic)bag.Own(sheet.Range["Q13"])).NumberFormat)=="0.0%","multi-Area money display preserves unselected gap values and formats");
                    dynamic zero=bag.Own(sheet.Range["Q14"]),negative=bag.Own(sheet.Range["R14"]),formulaZero=bag.Own(sheet.Range["Q15"]),copiedText=bag.Own(sheet.Range["R15"]),blank=bag.Own(sheet.Range["Q16"]);
                    Check(Convert.ToDouble(zero.Value2)==0&&Convert.ToString(zero.Text)=="-","numeric zero displays dash while its numeric value remains zero");
                    Check(Convert.ToBoolean(formulaZero.HasFormula)&&Convert.ToString(formulaZero.Formula)=="=0"&&Convert.ToDouble(formulaZero.Value2)==0&&Convert.ToString(formulaZero.Text)=="-","zero-result formula is preserved and displays dash");
                    Check(Convert.ToString(negative.Text)=="-1,234.50","negative amount retains sign grouping and two decimals");
                    Check(copiedText.Value2 is string&&Convert.ToString(copiedText.Value2)=="0"&&Convert.ToString(copiedText.Text)=="0"&&blank.Value2==null&&Convert.ToString(blank.Text)=="","text zero and blank cells remain unchanged");
                    dynamic first=bag.Own(sheet.Range["Q7"]),second=bag.Own(sheet.Range["R7"]),next=bag.Own(sheet.Range["S2"]),serial=bag.Own(sheet.Range["Q4"]);
                    Check(Convert.ToString(first.Formula).Replace("'", "")=="=Lead!$N$3"&&Convert.ToString(second.Formula).Replace("'", "")=="=Lead!$N$2","descending years link original cells");
                    Check(Convert.ToString(next.Formula).Replace("'", "")=="=Lead!$O$4"&&Convert.ToString(serial.Formula).Replace("'", "")=="=Lead!$Q$3","multiple selected fields keep Lead row pairing");
                    dynamic changed=bag.Own(lead.Range["N3"]);changed.Value2="2024年度";sheet.Calculate();Check(Convert.ToString(first.Value2)=="2024年度"&&Convert.ToString(first.Formula).Replace("'", "")=="=Lead!$N$3","Lead period edit updates value without reordering");changed.Value2="2027年度";
                    dynamic code=bag.Own(lead.Range["O3"]);code.Value2="甲更新";sheet.Calculate();Check(Convert.ToString(prior.Value2)=="甲更新","repeat formula updates from Lead");code.Value2="甲简称";
                    Call(host,"Undo");Check(Convert.ToString(prior.Formula)=="=1+2"&&string.IsNullOrEmpty((string)Convert.ToString(first.Value2))&&Convert.ToString(gap.Value2)=="未选空隙","variable headers and data undo together without touching gap");
                    Check(Math.Abs(Convert.ToDouble(outputColumn.ColumnWidth)-originalWidth)<0.01,"undo restores original output column width");
                    Check(Convert.ToDouble(priorZero.Value2)==999&&Convert.ToString(priorZero.NumberFormat)=="0.0000","undo restores amount value and format from before zero-display change");
                    Check(Convert.ToInt32(priorZero.HorizontalAlignment)==-4152&&Convert.ToInt32(priorZero.VerticalAlignment)==-4160,"undo restores original alignment");
                    object savedYearSource=Get(host,"yearSource"),savedYearTarget=Get(host,"yearTarget");Set(host,"yearSource",null);Set(host,"yearTarget",null);
                    Call(host,"Preview");Check((bool)Get(host,"ready")&&!view.HeaderPreview.Text.Contains("期间"),"repeat-only variables preflight without any period configuration");
                    Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);
                    Check(!(bool)Get(host,"failed")&&Convert.ToString(prior.Formula).Replace("'","")=="=Lead!$O$3"&&string.IsNullOrEmpty(Convert.ToString(first.Value2)),"repeat-only execution writes source links without creating period headers");
                    Call(host,"Undo");Check(Convert.ToString(prior.Formula)=="=1+2","repeat-only undo restores original header");Set(host,"yearSource",savedYearSource);Set(host,"yearTarget",savedYearTarget);
                    var preset=host.CapturePreset();host.ApplyPreset(preset);Check(view.VariableHeaders.Checked&&Get(host,"repeatSource")!=null&&Get(host,"repeatTarget")!=null,"variable settings preset roundtrip");
                    view.AddGroup("repeat");string repeatKey=view.ExtraGroups[0];view.AddGroup("year");string yearKey=view.ExtraGroups[1];var extras=(Dictionary<string,SummaryAreaPlan>)Get(host,"extraInputs");
                    extras[repeatKey+"Source"]=new SummaryAreaPlan{Sheet="Lead",Row=3,Column=17,Rows=2,Columns=1};extras[repeatKey+"Target"]=new SummaryAreaPlan{Sheet="变量汇总",Row=5,Column=17,Rows=1,Columns=1};
                    extras[yearKey+"Source"]=new SummaryAreaPlan{Sheet="Lead",Row=2,Column=14,Rows=2,Columns=1};extras[yearKey+"Target"]=new SummaryAreaPlan{Sheet="变量汇总",Row=8,Column=17,Rows=1,Columns=1};
                    Call(host,"Preview");Check(((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==24,"additional variable and period groups plan together");Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);
                    dynamic extraValue=bag.Own(sheet.Range["Q5"]),extraPeriod=bag.Own(sheet.Range["Q8"]);Check(!(bool)Get(host,"failed")&&Convert.ToString(extraValue.Formula).Replace("'","")=="=Lead!$Q$3"&&Convert.ToString(extraPeriod.Formula).Replace("'","")=="=Lead!$N$3","additional groups write correct linked sources");Call(host,"Undo");Check(string.IsNullOrEmpty((string)Convert.ToString(extraValue.Value2))&&string.IsNullOrEmpty((string)Convert.ToString(extraPeriod.Value2)),"all groups undo together");
                    var groupedPreset=host.CapturePreset();host.ApplyPreset(groupedPreset);Check(view.ExtraGroups.Count==2&&((Dictionary<string,SummaryAreaPlan>)Get(host,"extraInputs")).Count==4,"additional groups preset roundtrip");
                    extras=(Dictionary<string,SummaryAreaPlan>)Get(host,"extraInputs");extras[repeatKey+"Target"].Row=2;bool overlap=false;try{Call(host,"Preview");}catch(InvalidOperationException){overlap=true;}Check(overlap,"overlapping extra groups block before writing");view.ClearGroups();
                    view.AddGroup("repeat");bool incomplete=false;try{Call(host,"Preview");}catch(InvalidOperationException){incomplete=true;}Check(incomplete,"incomplete extra group blocks preflight");view.ClearGroups();
                    years.Value2=new object[,]{{"2027.1.1"},{"2027.2.1"}};Call(host,"Preview");Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&Convert.ToString(first.Value2)=="2027.2.1","text dates sort descending");Call(host,"Undo");
                    dynamic n2=bag.Own(lead.Range["N2"]),n3=bag.Own(lead.Range["N3"]);n2.Value2=46000.0;n3.Value2=47000.0;n2.NumberFormat=n3.NumberFormat="yyyy.m.d";Call(host,"Preview");Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&Convert.ToDouble(first.Value2)==47000&&Convert.ToString(first.NumberFormat)==Convert.ToString(n3.NumberFormat),"numeric dates retain value and source format");Call(host,"Undo");
                    n2.Value2=n3.Value2;bool blocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){blocked=true;}Check(blocked,"duplicate periods block before writing");
                    n2.Value2=46000.0;n3.Value2=47000.0;dynamic bad=bag.Own(sheet.Range["G10"]);bad.Formula="=1/0";Call(host,"Preview");Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);string executionLog=((RichTextBox)Get(view,"log")).Text;Check(!(bool)Get(host,"failed")&&executionLog.Contains("变量汇总!Q10")&&executionLog.Contains("#DIV/0!"),"execution log lists company error address and Excel error text");Call(host,"Undo");bad.Value2="BS!$A$1";
                    lead.Activate();dynamic moved=bag.Own(lead.Range["O5:Q5"]);dynamic originalSecond=bag.Own(lead.Range["P4"]);moved.Value2=new object[,]{{"second code",(object)originalSecond.Value2,"02"}};dynamic companiesPick=bag.Own(lead.Range["P3,P5"]);companiesPick.Select();Call(host,"Pick","lead",false);dynamic disjoint=bag.Own(lead.Range["O3,Q3,O5,Q5"]);disjoint.Select();Call(host,"Pick","repeatSource",false);Set(host,"repeatTarget",new SummaryAreaPlan{Sheet="变量汇总",Row=2,Column=17,Rows=2,Columns=1});Call(host,"Preview");Check((bool)Get(host,"ready"),"disjoint repeat fields read native Areas");Call(host,"Start");for(int i=0;i<((List<SummaryAreaPlan>)Get(host,"targets")).Count;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&Convert.ToString(((dynamic)bag.Own(sheet.Range["Q3"])).Formula).Replace("'","")=="=Lead!$Q$3"&&Convert.ToString(((dynamic)bag.Own(sheet.Range["S3"])).Formula).Replace("'","")=="=Lead!$Q$5","disjoint fields map original Lead rows and columns");Call(host,"Undo");var disjointPreset=host.CapturePreset();host.ApplyPreset(disjointPreset);Call(host,"Preview");Check((bool)Get(host,"ready")&&((Dictionary<SummaryAreaPlan,List<SummaryAreaPlan>>)Get(host,"variableAreas")).Count>0,"disjoint variable preset restores complete selection");
                    n2.Value2=46000.0;n3.Value2=47000.0;view.VariableHeaders.Checked=false;Set(host,"companyHeader",new SummaryAreaPlan{Sheet="变量汇总",Row=2,Column=7,Rows=1,Columns=1});Call(host,"Preview");Check((bool)Get(host,"ready"),"ordinary G2 name target independent of data paste start is accepted");
                    dynamic across=bag.Own(sheet.Range["G2:I2"]);across.Merge();bool crossed=false;try{Call(host,"Preview");}catch(InvalidOperationException){crossed=true;}Check(crossed,"actual merge crossing name groups still blocked");across.UnMerge();
                    Check(Convert.ToString(book.Path)=="","variable mode never saves workbook");
                }
            }
            book.Close(false);closed=true;
        }
    }
    [STAThread] private static int Main(string[] args)
    {
        if (Array.IndexOf(args, "--timing") >= 0) return SummaryAreaSpeedChecks.Run();
        using(var session=new ProbeSession("SummaryAreaLive"))
        {
            try
            {
                Type type=Type.GetTypeFromProgID("Excel.Application",false);
                if(type==null){Console.WriteLine("SKIP: Excel.Application not registered");return session.Finish(2);}
                dynamic app=session.Take(Activator.CreateInstance(type));
                Console.WriteLine("HOST: "+Convert.ToString(app.Name)+"; version="+Convert.ToString(app.Version));
                if(!session.Override(()=>app.DisplayAlerts,v=>app.DisplayAlerts=v,false,"DisplayAlerts") ||
                   !session.Override(()=>app.EnableEvents,v=>app.EnableEvents=v,false,"EnableEvents") ||
                   !session.Override(()=>app.ScreenUpdating,v=>app.ScreenUpdating=v,false,"ScreenUpdating"))throw new Exception("Cannot establish isolated host state");
                using(var bag=new ComBag())
                {
                    dynamic books=bag.Own(app.Workbooks);
                    foreach(string name in new[]{"！01-示例（集团）","！02-示例乙","公共参数","第三家公司"})
                    {
                        dynamic source=books.Add(); bool closed=false;
                        session.Own((object)source,delegate{if(!closed)source.Close(false);});
                        using(var refs=new ComBag()) { dynamic sheets=refs.Own(source.Worksheets), sheet=refs.Own(sheets.Item(1));sheet.Name="BS";dynamic cell=refs.Own(sheet.Range["A1"]);cell.Value2=25; }
                        source.SaveAs(session.Output(name+".xlsx"),51);source.Close(false);closed=true;
                    }
                    if(Array.IndexOf(args,"--consolidated-only")>=0){ConsolidatedTemplates((object)app,session);Console.WriteLine("PASS: "+checks+" consolidated checks");return session.Finish(0);}
                    if(Array.IndexOf(args,"--structure-only")>=0){StructureTemplates((object)app,session);Console.WriteLine("PASS: "+checks+" structure template checks");return session.Finish(0);}
                    if(Array.IndexOf(args,"--template-entry-only")>=0){TemplateEntryWindows((object)app);Console.WriteLine("PASS: "+checks+" independent template window checks");return session.Finish(0);}
                    if(Array.IndexOf(args,"--file-batch-window-only")>=0){FileBatchWindow((object)app);Console.WriteLine("PASS: "+checks+" file batch window checks");return session.Finish(0);}
                    if(Array.IndexOf(args,"--file-batch-only")>=0){FileBatch((object)app,session);Console.WriteLine("PASS: "+checks+" file batch checks");return session.Finish(0);}
                    if(Array.IndexOf(args,"--variables-only")>=0){VariableHeaders((object)app,session);Console.WriteLine("PASS: "+checks+" variable header checks");return session.Finish(0);}
                    dynamic book=books.Add();session.Own((object)book,delegate{book.Close(false);});
                    dynamic worksheets=bag.Own(book.Worksheets), target=bag.Own(worksheets.Item(1));target.Name="汇总";
                    dynamic lead=bag.Own(worksheets.Add());lead.Name="Lead";
                    dynamic leadRange=bag.Own(lead.Range["A1:B2"]);leadRange.Value2=new object[,]{{"01","！01-示例（集团）"},{"02","！02-示例乙"}};
                    foreach(string address in new[]{"L3","L6"}){dynamic cell=bag.Own(target.Range[address]);cell.Value2="BS!$A$1";}
                    dynamic native=bag.Own(target.Range["M3"]);native.Formula="=ROUND(L3-L3,2)";
                    dynamic blank=bag.Own(target.Range["M6"]), fill=bag.Own(blank.Interior);fill.Color=65535;
                    dynamic destination=bag.Own(target.Range["Q3:T6"]);destination.Value2="原值";
                    dynamic gap=bag.Own(target.Range["Q4"]);gap.Value2="未选空隙";bag.Own(gap.AddComment("空隙批注"));dynamic gapFill=bag.Own(gap.Interior);gapFill.Color=12345;
                    dynamic old=bag.Own(target.Range["Q3"]);old.Formula="='汇总'!A10+7";string oldFormula=Convert.ToString(old.Formula);
                    dynamic note=bag.Own(target.Range["R6"]);bag.Own(note.AddComment("原目标批注"));
                    int before=Convert.ToInt32(worksheets.Count);
                    using(var form=new SingleSummaryForm(true))
                    {
                        var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;view.HeadersEnabled.Checked=false;view.VariableHeaders.Checked=false;view.FormulaHeaders.Checked=false;view.RepeatNames.Checked=false;
                        using(var host=new SummaryAreaHost((object)app,form,view,null,delegate{return true;}))
                        {
                            Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet","Lead");
                            var areas=(List<SummaryAreaPlan>)Get(host,"areas");areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=3,Column=12,Rows=1,Columns=2,Address="L3:M3"});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=6,Column=12,Rows=1,Columns=2,Address="L6:M6"});
                            Set(host,"paste",new SummaryAreaPlan{Row=3,Column=17,Rows=1,Columns=1});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=1,Rows=2,Columns=2,Address="A1:B2"});
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"read-only preflight succeeds");Check(Convert.ToInt32(worksheets.Count)==before,"preflight creates no backup sheet");Check(Convert.ToString(old.Formula)==oldFormula,"preflight leaves target formula untouched");
                            int calculationBefore=Convert.ToInt32(app.Calculation);Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(Convert.ToInt32(app.Calculation)==calculationBefore,"batch restores original calculation mode");Check(Convert.ToInt32(Get(host,"bulkWrites"))==4,"four company Areas use four matrix formula submissions");
                            Check((bool)Get(host,"pending")&&!(bool)Get(host,"running")&&!(bool)Get(host,"failed"),"execution reaches pending confirmation");
                            dynamic q3=bag.Own(target.Range["Q3"]),s3=bag.Own(target.Range["S3"]),r3=bag.Own(target.Range["R3"]),r6=bag.Own(target.Range["R6"]);
                            Check(Convert.ToBoolean(q3.HasFormula)&&Convert.ToString(q3.Formula).Contains("[！01-示例（集团）.xlsx]BS"),"first company real external formula");
                            Check(Convert.ToString(s3.Formula).Contains("[！02-示例乙.xlsx]BS"),"second company from original template");
                            Check(Convert.ToString(r3.Formula)=="=ROUND(Q3-Q3,2)","native relative formula copied horizontally");
                            Check(r6.Value2==null,"selected blank replaces old target content");Check(Convert.ToInt32(((dynamic)bag.Own(r6.Interior)).Color)==65535,"selected blank retains template formatting");
                            Check(Convert.ToString(gap.Value2)=="未选空隙"&&Convert.ToInt32(gapFill.Color)==12345,"unselected gap value and format unchanged");Check(Convert.ToString(((dynamic)bag.Own(gap.Comment)).Text())=="空隙批注","unselected gap comment unchanged");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["L3"])).Value2)=="BS!$A$1","original template remains text");
                            Call(host,"Undo");Check(Convert.ToString(q3.Formula)==oldFormula,"rollback restores exact explicit self-sheet formula");Check(Convert.ToString(((dynamic)bag.Own(r6.Comment)).Text())=="原目标批注","rollback restores target comment");Check(Convert.ToString(r6.Value2)=="原值","rollback restores selected blank target");
                            Call(host,"Preview");dynamic change=bag.Own(target.Range["L3"]);change.Value2="BS!$A$2";bool rejected=false;try{Call(host,"Start");}catch(InvalidOperationException){rejected=true;}Check(rejected,"template edit invalidates preflight without relying on events");change.Value2="BS!$A$1";
                            Call(host,"Preview");Call(host,"Start");Call(host,"Tick",null,EventArgs.Empty);target.Protect();Call(host,"Tick",null,EventArgs.Empty);
                            Check((bool)Get(host,"failed")&&(bool)Get(host,"pending"),"injected protection failure keeps recovery session");target.Unprotect();Call(host,"Undo");Check(!(bool)Get(host,"pending")&&Convert.ToString(q3.Formula)==oldFormula,"retry undo recovers failed batch");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Call(host,"Confirm");Check(!(bool)Get(host,"pending"),"confirmation ends rollback session");Check(Convert.ToString(book.Path)=="","target workbook was never saved");
                            target.Activate();lead.Name="lEaD";Check(host.ActivateLeadSheet(),"Lead jump recognizes case-insensitive sheet name");
                            dynamic active=bag.Own(app.ActiveSheet);Check(Convert.ToString(active.Name)=="lEaD","Lead jump activates actual Lead worksheet");
                            lead.Name="公司清单";target.Activate();Check(!host.ActivateLeadSheet(),"missing Lead keeps manual worksheet selection available");lead.Name="Lead";
                            foreach(var pair in new[]{new[]{"A7","01"},new[]{"D7","！01-示例（集团）"},new[]{"A4","02"},new[]{"D4","！02-示例乙"},new[]{"A5","未选行"},new[]{"D5","未选行不得读取"}}){dynamic cell=bag.Own(lead.Range[pair[0]]);cell.Value2=pair[1];}
                            lead.Activate();dynamic union=bag.Own(lead.Range["A7,D7,A4,D4"]);union.Select();
                            dynamic unionAreas=bag.Own(union.Areas),firstArea=bag.Own(unionAreas.Item(1));int expectedFirst=Convert.ToInt32(firstArea.Row);
                            Call(host,"Pick","lead",false);Check(((List<SummaryAreaPlan>)Get(host,"leadAreas")).Count==4,"actual Excel Selection.Areas captured for disjoint Lead fields");
                            Call(host,"Preview");Check((bool)Get(host,"ready")&&view.Companies.Rows.Count==2,"disjoint Lead preflight ignores unselected middle rows");
                            Check(Convert.ToInt32(view.Companies.Rows[0].Cells["row"].Value)==expectedFirst,"disjoint Lead retains Excel Areas order");
                            Check(Convert.ToString(view.Companies.Rows[0].Cells["company"].Value)==(expectedFirst==7?"！01-示例（集团）":"！02-示例乙"),"nonadjacent code and name columns paired by same row");
                            view.Folder.Text="\""+session.DirectoryPath+"\"";Check(!(bool)Get(host,"ready"),"manual folder edit invalidates prior plan");Call(host,"Preview");Check((bool)Get(host,"ready"),"manually pasted quoted folder passes preflight");
                            var saved=host.CapturePreset();string presetPath=session.Output("区域预设.xybatch.json");BatchPresetStore.Write(presetPath,saved,false);session.TrackFile(presetPath+".lock");saved=BatchPresetStore.Read(presetPath,"summary");
                            Check(!saved.Values.ContainsKey("companies")&&saved.Values["leadAreas"].Split(';').Length==4,"saved presets contain settings and all Areas, not company data");
                            Check((bool)typeof(SummaryAreaHost).GetMethod("ContextAvailable",Flags).Invoke(host,new object[]{saved.Context}),"unsaved open workbook context can be reused without file relocation");
                            view.Folder.Text="待替换路径";var invalid=BatchPresetStore.Decode(BatchPresetStore.Encode(saved),"summary");invalid.Values["templateSheet"]="不存在的工作表";
                            bool failedApply=false;try{host.ApplyPreset(invalid);}catch{failedApply=true;}Check(failedApply&&view.Folder.Text=="待替换路径","invalid preset retains prior UI and workbook binding");
                            var legacy=BatchPresetStore.Decode(BatchPresetStore.Encode(saved),"summary");legacy.Values.Remove("mode");string legacyPath=session.Output("旧整列预设.xybatch.json");BatchPresetStore.Write(legacyPath,legacy,false);session.TrackFile(legacyPath+".lock");string legacyBytes=Convert.ToBase64String(System.IO.File.ReadAllBytes(legacyPath));
                            failedApply=false;try{host.ApplyPreset(legacy);}catch(InvalidOperationException){failedApply=true;}Check(failedApply&&legacyBytes==Convert.ToBase64String(System.IO.File.ReadAllBytes(legacyPath)),"legacy preset remains unchanged and is not guessed into area mode");
                            int sheetsBefore=Convert.ToInt32(worksheets.Count);host.ApplyPreset(saved);view.HeadersEnabled.Checked=false;view.VariableHeaders.Checked=false;Check(!(bool)Get(host,"ready")&&view.Companies.Rows.Count==0&&view.Folder.Text==saved.Values["projectFolder"],"apply restores parameters and requires a fresh company preflight");
                            Check(Convert.ToInt32(worksheets.Count)==sheetsBefore&&Convert.ToString(book.Path)=="","applying preset creates no backup and never saves workbook");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);failedApply=false;try{host.ApplyPreset(saved);}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"pending rollback blocks preset application");Call(host,"Undo");
                            Call(host,"Preview");dynamic formatCell=bag.Own(target.Range["L3"]),formatFont=bag.Own(formatCell.Font);object priorColor=formatFont.Color;formatFont.Color=12345;failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"native block fingerprint rejects a font-only change before execution");formatFont.Color=priorColor;
                            dynamic priorNote=bag.Own(formatCell.AddComment("状态校验批注"));Call(host,"Preview");priorNote.Text("预检后修改批注");failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"native block fingerprint rejects comment-only change before execution");priorNote.Text("状态校验批注");
                            dynamic computedName=bag.Own(lead.Range["D7"]),nameSource=bag.Own(lead.Range["Z7"]);nameSource.Value2="！01-示例（集团）";computedName.Formula="=Z7";computedName.Calculate();Call(host,"Preview");nameSource.Value2="！02-示例乙";computedName.Calculate();failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"computed Lead company name change invalidates plan even with unchanged formula text");computedName.Value2="！01-示例（集团）";nameSource.Clear();
                            Check(!view.HeadersEnabled.Checked,"legacy copy-only fixture explicitly disables headers");view.FormulaHeaders.Checked=false;view.RepeatNames.Checked=false;
                            dynamic yearSource=bag.Own(lead.Range["G1:H1"]);yearSource.Value2=new object[,]{{46388,"2026年度"}};
                            dynamic dated=bag.Own(lead.Range["G1"]);dated.NumberFormat="yyyy年";string displayedYear=Convert.ToString(dated.Text);
                            dynamic firstName=bag.Own(target.Range["Q2:R2"]),secondName=bag.Own(target.Range["S2:T2"]);firstName.Merge();secondName.Merge();
                            dynamic firstNameCell=bag.Own(target.Range["Q2"]);firstNameCell.Formula="='汇总'!A10+12";string nameBefore=Convert.ToString(firstNameCell.Formula);
                            dynamic oldYears=bag.Own(target.Range["Q5:T5"]);oldYears.Value2="旧年度";oldYears.NumberFormat="0.00";
                            target.Activate();firstName.Select();Call(host,"Pick","companyHeader",false);Check(((SummaryAreaPlan)Get(host,"companyHeader")).Columns==1,"merged selection normalizes to name anchor");
                            lead.Activate();yearSource.Select();Call(host,"Pick","yearSource",false);target.Activate();dynamic yearAnchor=bag.Own(target.Range["Q5"]);yearAnchor.Select();Call(host,"Pick","yearTarget",false);view.HeadersEnabled.Checked=true;
                            Call(host,"Preview");Check((bool)Get(host,"ready")&&((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==6,"merged names and years pass read-only preflight");Check(Convert.ToString(firstNameCell.Formula)==nameBefore,"header preflight leaves formula untouched");
                            string expectedName1=Convert.ToString(view.Companies.Rows[0].Cells["code"].Value),expectedName2=Convert.ToString(view.Companies.Rows[1].Cells["code"].Value);
                            var headerPreset=host.CapturePreset();host.ApplyPreset(headerPreset);view.VariableHeaders.Checked=false;Check(view.HeadersEnabled.Checked&&Get(host,"yearSource")!=null,"header presets round-trip");
                            Call(host,"Preview");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check((bool)Get(host,"pending")&&!(bool)Get(host,"failed"),"header batch reaches pending");
                            Check(Convert.ToString(firstNameCell.Value2)==expectedName1&&Convert.ToString(((dynamic)bag.Own(target.Range["S2"])).Value2)==expectedName2,"Lead first column maps to company groups");
                            Check(Convert.ToString(yearAnchor.Value2)==displayedYear&&Convert.ToString(((dynamic)bag.Own(target.Range["S5"])).Value2)==displayedYear,"date source repeats display text instead of serial");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["R5"])).Value2)=="2026年度"&&Convert.ToString(((dynamic)bag.Own(target.Range["T5"])).Value2)=="2026年度","year source order repeats");
                            Check(Convert.ToBoolean(firstName.MergeCells)&&Convert.ToString(gap.Value2)=="未选空隙","legal merges and unselected gaps unchanged");
                            Call(host,"Undo");Check(Convert.ToString(firstNameCell.Formula)==nameBefore&&Convert.ToString(yearAnchor.Value2)=="旧年度"&&Convert.ToString(yearAnchor.NumberFormat)=="0.00","undo restores header formula values and format");
                            Check(Convert.ToString(q3.Formula).Contains(".xlsx]BS"),"header undo restores prior link batch");
                            Call(host,"Preview");dated.NumberFormat="yyyy";failedApply=false;try{Call(host,"Start");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"year display change invalidates preflight");dated.NumberFormat="yyyy年";
                            Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=7,Rows=1,Columns=1});failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"year count mismatch blocked");Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=7,Rows=1,Columns=2});
                            Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=3,Column=17,Rows=1,Columns=1});failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"header and copy conflict blocked");Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=5,Column=17,Rows=1,Columns=1});
                            firstName.UnMerge();secondName.UnMerge();dynamic illegal=bag.Own(target.Range["Q2:S2"]);illegal.Merge();failedApply=false;try{Call(host,"Preview");}catch(InvalidOperationException){failedApply=true;}Check(failedApply,"cross-company merged name blocked");illegal.UnMerge();
                            host.ApplyPreset(saved);Check(view.HeadersEnabled.Checked&&Get(host,"companyHeader")==null,"restored preset uses current automatic-header default");view.HeadersEnabled.Checked=false;view.VariableHeaders.Checked=false;Call(host,"Preview");Check((bool)Get(host,"ready"),"V1 remains usable after header validation failures");
                            // Three companies, width five, two disjoint Areas; selected gaps stay untouched.
                            areas.Clear();areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=10,Column=12,Rows=1,Columns=5});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=14,Column=12,Rows=1,Columns=5});
                            dynamic fiveTemplate=bag.Own(target.Range["L10:P10,L14:P14"]);fiveTemplate.Value2="BS!$A$1";
                            dynamic threeLead=bag.Own(lead.Range["A20:B22"]);threeLead.Value2=new object[,]{{"甲简称","！01-示例（集团）"},{"乙简称","！02-示例乙"},{"丙简称","第三家公司"}};
                            ((List<SummaryAreaPlan>)Get(host,"leadAreas")).Clear();Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=20,Column=1,Rows=3,Columns=2});view.Companies.Rows.Clear();Set(host,"paste",new SummaryAreaPlan{Row=10,Column=17,Rows=1,Columns=1});
                            dynamic fiveYears=bag.Own(lead.Range["G2:K2"]);fiveYears.Value2=new object[,]{{"2027年度","2026年度","2025年度","2024年度","2023年度"}};
                            Set(host,"companyHeader",new SummaryAreaPlan{Sheet="汇总",Row=9,Column=17,Rows=1,Columns=1});Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=2,Column=7,Rows=1,Columns=5});Set(host,"yearTarget",new SummaryAreaPlan{Sheet="汇总",Row=12,Column=17,Rows=1,Columns=1});view.HeadersEnabled.Checked=true;
                            dynamic untouched=bag.Own(target.Range["R11"]);untouched.Value2="未授权间隙";
                            Call(host,"Preview");Check(((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==18,"three companies width five plan has eighteen header tasks");Call(host,"Start");for(int i=0;i<6;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"three-company five-column multi-Area execution succeeds");
                            bool mapping=true;for(int i=0;i<3;i++){dynamic name=bag.Own(target.Cells[9,17+i*5]);mapping &= Convert.ToString(name.Value2)==new[]{"甲简称","乙简称","丙简称"}[i];for(int c=0;c<5;c++){dynamic year=bag.Own(target.Cells[12,17+i*5+c]);mapping &= Convert.ToString(year.Value2)==(2027-c)+"年度";}}
                            Check(mapping,"Q V AA company names and all fifteen years map correctly");Check(Convert.ToString(untouched.Value2)=="未授权间隙","three-company unselected gap remains untouched");
                            Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q9"])).Value2==null&&((dynamic)bag.Own(target.Range["Q12"])).Value2==null&&((dynamic)bag.Own(target.Range["Q10"])).Value2==null,"three-company rollback restores blank headers and link targets");
                            Call(host,"Preview");Call(host,"Start");Call(host,"Tick",null,EventArgs.Empty);typeof(SummaryAreaHost).GetMethod("WriteHeaders",Flags).Invoke(host,new object[]{target,bag});
                            target.Protect();for(int i=0;i<2;i++)Call(host,"Tick",null,EventArgs.Empty);target.Unprotect();Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q9"])).Value2==null&&((dynamic)bag.Own(target.Range["Q12"])).Value2==null,"failed batch retry undo includes headers already touched");
                            view.Companies.Rows[1].Cells["selected"].Value=false;Call(host,"Preview");Check(((List<SummaryAreaPlan>)Get(host,"headerTargets")).Count==12,"company filtering rebuilds header plan with two groups");Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["V9"])).Value2)=="丙简称"&&((dynamic)bag.Own(target.Range["AA9"])).Value2==null,"unchecked middle company does not shift name and link pairing");Call(host,"Undo");
                            // Existing compiled Area: company external + unrelated public external + constants.
                            dynamic compiledSource=bag.Own(target.Range["L10"]),publicSource=bag.Own(target.Range["M10"]),constant=bag.Own(target.Range["N10"]),literalText=bag.Own(target.Range["O10"]);
                            compiledSource.Formula=SummaryTextReference.External(session.DirectoryPath,"！01-示例（集团）","BS!A1")+"+A1";
                            publicSource.Formula=SummaryTextReference.External(session.DirectoryPath,"公共参数","BS!$A$1");constant.Value2=123;literalText.Value2="说明! 请检查";dynamic formulaText=bag.Own(target.Range["P10"]);formulaText.NumberFormat="@";formulaText.Value2="=1+1";
                            bool reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}
                            Check(reviewBlocked&&view.ReferenceSources.Rows.Count==2,"unrelated sources are reported and block first preflight");
                            foreach(DataGridViewRow reviewRow in view.ReferenceSources.Rows)reviewRow.Cells["decision"].Value=Convert.ToString(reviewRow.Cells["source"].Value).Contains("公共参数")?"保留原引用":"参与公司替换";
                            Call(host,"Preview");Check((bool)Get(host,"ready"),"explicit source decisions permit fresh preflight");string publicBefore=Convert.ToString(publicSource.Formula),compiledBefore=Convert.ToString(compiledSource.Formula);
                            Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed"),"existing compiled Area executes");
                            dynamic secondLink=bag.Own(target.Range["V10"]),publicOutput=bag.Own(target.Range["R10"]);
                            Check(Convert.ToString(secondLink.Formula).Contains("[第三家公司.xlsx]BS'!A1+K1"),"existing external address frozen while local relative reference shifts");
                            Check(Convert.ToString(publicOutput.Formula)==publicBefore,"public external formula remains unchanged");
                            Check(Convert.ToDouble(((dynamic)bag.Own(target.Range["S10"])).Value2)==123&&Convert.ToString(((dynamic)bag.Own(target.Range["T10"])).Value2)=="说明! 请检查","plain numbers and text copied unchanged");
                            dynamic literalOutput=bag.Own(target.Range["U10"]);Check(!Convert.ToBoolean(literalOutput.HasFormula)&&Convert.ToString(literalOutput.Value2)=="=1+1","literal equals text never becomes an Excel formula");
                            Call(host,"Undo");Check(((dynamic)bag.Own(target.Range["Q10"])).Value2==null&&Convert.ToString(compiledSource.Formula)==compiledBefore,"compiled Area undo restores targets and leaves source unchanged");
                            publicSource.Formula=publicBefore+"+1";reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}Check(reviewBlocked,"changed source formula requires renewed review");
                            publicSource.Value2="公共说明";view.ReferenceSources.Rows.Clear();Call(host,"Preview");Check((bool)Get(host,"ready")&&view.ReferenceSources.Rows.Count==1&&Convert.ToString(view.ReferenceSources.Rows[0].Cells["decision"].Value)=="参与公司替换","single Lead-matched source is automatically identified");
                            // Linked O/P company source + vertical N years; independent of any elimination sheet.
                            view.HeadersEnabled.Checked=true;view.FormulaHeaders.Checked=true;view.RepeatNames.Checked=true;view.ReplaceHeaderArrays.Checked=false;
                            dynamic newLead=bag.Own(lead.Range["O30:P32"]);newLead.Value2=new object[,]{{"甲简称","！01-示例（集团）"},{"乙简称","！02-示例乙"},{"丙简称","第三家公司"}};
                            dynamic verticalYears=bag.Own(lead.Range["N30:N34"]);verticalYears.Value2=new object[,]{{"2027年度"},{"2026年度"},{"2025年度"},{"2024年度"},{"2023年度"}};
                            ((List<SummaryAreaPlan>)Get(host,"leadAreas")).Clear();Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=30,Column=15,Rows=3,Columns=2});view.Companies.Rows.Clear();
                            Set(host,"yearSource",new SummaryAreaPlan{Sheet="Lead",Row=30,Column=14,Rows=5,Columns=1});
                            // Existing full-row CSE has five-year repetitions, matching LET/SEQUENCE output semantics.
                            dynamic oldNames=bag.Own(target.Range["Q9:AE9"]);oldNames.FormulaArray="=INDEX(Lead!$O$30:$O$32,INT((COLUMN(Q9:AE9)-COLUMN(Q9))/5)+1)";oldNames.Calculate();string oldNamesFormula=Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Formula);
                            dynamic oldYearArray=bag.Own(target.Range["Q12:U12"]);oldYearArray.FormulaArray="=TRANSPOSE(Lead!$N$30:$N$34)";oldYearArray.Calculate();string priorYearFormula=Convert.ToString(((dynamic)bag.Own(target.Range["Q12"])).Formula);
                            reviewBlocked=false;try{Call(host,"Preview");}catch(InvalidOperationException){reviewBlocked=true;}Check(reviewBlocked,"existing header arrays require explicit full-range rewrite consent");
                            view.Companies.Rows[2].Cells["selected"].Value=false;view.ReplaceHeaderArrays.Checked=true;Call(host,"Preview");Check((bool)Get(host,"ready")&&((List<SummaryHeaderArray>)Get(host,"oldHeaderArrays")).Count==2,"vertical years and existing full company/year arrays pass preflight");
                            Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);Check(!(bool)Get(host,"failed")&&(bool)Get(host,"pending"),"linked headers replace existing arrays without LET dependencies");
                            bool linked=true;for(int i=0;i<2;i++)for(int c=0;c<5;c++){dynamic name=bag.Own(target.Cells[9,17+i*5+c]);linked &= Convert.ToBoolean(name.HasFormula)&&Convert.ToString(name.Value2)==(i==0?"甲简称":"乙简称");}
                            Check(linked,"each selected company name repeats across all year columns using Lead O links");
                            dynamic newYearAnchor=bag.Own(target.Range["Q12"]);Check(Convert.ToBoolean(newYearAnchor.HasArray)&&Convert.ToString(newYearAnchor.Formula).Contains("TRANSPOSE")&&Convert.ToString(newYearAnchor.Value2)=="2027年度","vertical N years written as true TRANSPOSE array formula");
                            Check(Convert.ToString(((dynamic)bag.Own(target.Range["AA9"])).Value2)=="丙简称"&&Convert.ToBoolean(((dynamic)bag.Own(target.Range["AA9"])).HasFormula),"unselected company outside current output keeps its Lead mapping after whole-array rewrite");
                            dynamic changedCode=bag.Own(lead.Range["O30"]),changedYear=bag.Own(lead.Range["N30"]);changedCode.Value2="甲简称更新";changedYear.Value2="2030年度";target.Calculate();Check(Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Value2)=="甲简称更新"&&Convert.ToString(newYearAnchor.Value2)=="2030年度","generated headers update when Lead source values change");changedCode.Value2="甲简称";changedYear.Value2="2027年度";target.Calculate();
                            Call(host,"Undo");Check(Convert.ToBoolean(oldNames.HasArray)&&Convert.ToString(((dynamic)bag.Own(target.Range["Q9"])).Formula)==oldNamesFormula&&Convert.ToString(newYearAnchor.Formula)==priorYearFormula,"undo restores complete original company and year array formulas");
                            var linkedPreset=host.CapturePreset();host.ApplyPreset(linkedPreset);view.VariableHeaders.Checked=false;Check(view.FormulaHeaders.Checked&&view.RepeatNames.Checked&&!view.ReplaceHeaderArrays.Checked,"formula and repeat settings persist but replacement permission resets");
                            object priorYearFormat=newYearAnchor.NumberFormat;changedYear.Value2=46388;changedYear.NumberFormat="yyyy\"年\"";string expectedDateDisplay=Convert.ToString(changedYear.Text);view.ReplaceHeaderArrays.Checked=true;
                            Call(host,"Preview");Call(host,"Start");int dateSteps=((List<SummaryAreaPlan>)Get(host,"targets")).Count;for(int i=0;i<dateSteps;i++)Call(host,"Tick",null,EventArgs.Empty);
                            Check(!(bool)Get(host,"failed")&&Convert.ToBoolean(newYearAnchor.HasFormula)&&Convert.ToDouble(newYearAnchor.Value2)==46388&&Convert.ToString(newYearAnchor.Text)==expectedDateDisplay,"linked TRANSPOSE retains numeric dates with source display format");
                            Call(host,"Undo");Check(Convert.ToString(newYearAnchor.NumberFormat)==Convert.ToString(priorYearFormat)&&Convert.ToString(newYearAnchor.Formula)==priorYearFormula,"date header undo restores original format and array formula");changedYear.Value2="2027年度";changedYear.NumberFormat="General";






                        }
                    }
                }
                if(args.Length==2&&args[0]=="--provided-header-vectors")ProvidedHeaders((object)app,session,args[1]);
                VariableHeaders((object)app,session);
                Console.WriteLine("PASS: "+checks+" Excel self-owned sample checks; no user workbook or real button interaction");return session.Finish(0);
            }
            catch(Exception e){Console.WriteLine(e);return session.Finish(1);}
        }
    }
}

[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderVector
{
    [System.Runtime.Serialization.DataMember] public string sheet;
    [System.Runtime.Serialization.DataMember] public List<ProvidedHeaderYear> years;
    [System.Runtime.Serialization.DataMember] public List<ProvidedHeaderCompany> companies;
}
[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderYear
{
    [System.Runtime.Serialization.DataMember] public int row;
    [System.Runtime.Serialization.DataMember] public string text;
}
[System.Runtime.Serialization.DataContract] public sealed class ProvidedHeaderCompany
{
    [System.Runtime.Serialization.DataMember] public int row;
    [System.Runtime.Serialization.DataMember] public string code;
    [System.Runtime.Serialization.DataMember] public string name;
}
