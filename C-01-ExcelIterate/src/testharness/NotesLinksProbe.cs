using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Excel_Iterate;

internal static class NotesLinksProbe
{
    private static int count;
    private static bool scale;
    private static string liveHost="";
    private static void Check(bool ok,string message){if(!ok)throw new Exception(message);count++;Console.WriteLine("PASS "+message);}
    private static void Reject(Action action,string message){bool rejected=false;try{action();}catch(InvalidOperationException){rejected=true;}catch(FormatException){rejected=true;}Check(rejected,message);}
    private static void Pure()
    {
        Check(ConsolidatedWorkbookForm.SequenceKey("！00-")=="0"&&ConsolidatedWorkbookForm.SequenceKey("!01-02-")=="1-2","report hierarchy accepts exclamation prefix and leading zeros");
        Check(CompanyStructureTasks.Key("！03-08-")=="3-8"&&CompanyStructureTasks.Key("!03-08-")=="3-8","hierarchy ignores fullwidth and ASCII leading business markers");
        Check(NotesLinkRules.IsDirect("！03-08-","!03-08-01-")&&NotesLinkRules.IsDirect("!00-","！03-"),"marked parent and child codes preserve direct hierarchy");
        Check(!NotesLinkRules.IsDirect("!03-08-","！03-08-01-01-"),"marker support still excludes grandchildren");
        Reject(()=>CompanyStructureTasks.Key("03-!08-"),"business marker inside a code remains invalid");
        Reject(()=>CompanyStructureTasks.Key("！"),"marker without a numeric code rejected");
        var marked=new CompanyStructureTask{Sequence="！03-08-01-",FileName="!03-08-01-sample.xlsx"};CompanyStructureTasks.Key(marked.Sequence);
        Check(marked.Sequence=="！03-08-01-"&&marked.FileName=="!03-08-01-sample.xlsx","comparison normalization preserves original code and file name");
        string batchRoot=Path.Combine(Path.GetTempPath(),"notes-batch-unit");
        var files=NotesLinkBatch.Files(new[]{Path.Combine(batchRoot,"parent.xlsx"),Path.Combine(batchRoot,"PARENT.xlsx"),Path.Combine(batchRoot,"~$lock.xlsx"),Path.Combine(batchRoot,".xy-notes-stage.xlsx"),Path.Combine(batchRoot,"source.txt"),Path.Combine(batchRoot,"other.xlsm")});
        Check(files.Count==2&&files[1].EndsWith("other.xlsm"),"batch file selection deduplicates paths and excludes locks staging and non-workbooks");
        Func<string,string,NotesLinkBatchItem> node=(file,sourceFile)=>{var plan=new NotesLinkPlan{Config=new NotesLinkConfig{Parent=Path.Combine(batchRoot,file)}};plan.Tasks.Add(new NotesLinkTask{Status="通过",File=Path.Combine(batchRoot,sourceFile)});return new NotesLinkBatchItem{Parent=plan.Config.Parent,Plan=plan};};
        var lower=node("child.xlsx","leaf.xlsx");var upper=node("parent.xlsx","child.xlsx");
        Check(NotesLinkBatch.Order(new[]{lower,upper}).SequenceEqual(new[]{upper,lower}),"batch processes a selected parent before its source workbook to avoid stale hashes");
        var blocked=new NotesLinkBatchItem{Parent=Path.Combine(batchRoot,"bad.xlsx")};var uncheckedFile=node("unchecked.xlsx","leaf.xlsx");uncheckedFile.Selected=false;
        Check(NotesLinkBatch.Order(new[]{blocked,lower,uncheckedFile}).SequenceEqual(new[]{lower}),"blocked and unchecked batch files do not run");
        Reject(()=>NotesLinkBatch.Order(new[]{node("cycle-a.xlsx","cycle-b.xlsx"),node("cycle-b.xlsx","cycle-a.xlsx")}),"cycles among selected parent files blocked before writes");
        try{NotesLinkRules.LeadCodeKey("编码","Lead","K1",false);throw new Exception("header should be rejected");}catch(InvalidOperationException e){Check(e.Message.Contains("Lead!K1")&&e.Message.Contains("编码")&&e.Message.Contains("标题行"),"invalid child code identifies its cell and header range remedy");}
        try{NotesLinkRules.LeadCodeKey("","Lead","K1",true);throw new Exception("empty parent should be rejected");}catch(InvalidOperationException e){Check(e.Message.Contains("Lead!K1")&&e.Message.Contains("空白")&&e.Message.Contains("汇总公司编码在"),"invalid parent identity identifies its actual configured cell");}
        try{NotesLinkRules.LeadCodeKey("03-\u200b","Lead","K3",false);throw new Exception("hidden marker should be rejected");}catch(InvalidOperationException e){Check(e.Message.Contains("U+200B"),"invisible character is made visible in the diagnostic instead of silently removed");}
        Reject(()=>NotesLinkBatch.PickerParent(new NotesLinkConfig(),new[]{new NotesLinkBatchItem{Parent="unchecked.xlsx",Selected=false}},"unchecked.xlsx"),"parent picker does not use an unchecked file as a reference");
        var cleanupFacts=new NotesSheetFacts{Snapshot=new MigrationSheetSnapshot()};
        foreach(string header in new[]{"AV1:AX1","AZ1:BB1","BD1:BF1"})cleanupFacts.Snapshot.Merges.Add(MigrationArea.Parse(header));
        var cleanupConfig=new NotesLinkConfig{FirstStart="AV1",ColumnSpacing=4};
        var cleanupRows=new[]{new NotesLinkTask{Slot=0},new NotesLinkTask{Slot=1,Selected=false}};
        var cleanup=NotesLinkPlanner.FindUnusedTemplate(cleanupFacts,MigrationArea.Parse("D1:F4"),cleanupConfig,cleanupRows);
        Check(cleanup.Address=="BD1:BF4","unused cleanup starts after all Lead slots including unchecked companies");
        cleanupFacts.Arrays.Add(MigrationArea.Parse("BE3"));
        bool unsafeCleanup=false;try{NotesLinkPlanner.FindUnusedTemplate(cleanupFacts,MigrationArea.Parse("D1:F4"),cleanupConfig,cleanupRows);}catch(InvalidOperationException){unsafeCleanup=true;}
        Check(unsafeCleanup,"array formulas block destructive unused-template cleanup");cleanupFacts.Arrays.Clear();
        cleanupFacts.Snapshot.Merges.RemoveAt(0);
        Check(NotesLinkPlanner.FindUnusedTemplate(cleanupFacts,MigrationArea.Parse("D1:F4"),cleanupConfig,cleanupRows)==null,"missing contiguous header evidence cannot clear arbitrary right columns");
        var mergeTarget=new NotesSheetFacts{Snapshot=new MigrationSheetSnapshot()};mergeTarget.Snapshot.Merges.Add(MigrationArea.Parse("AV2:AX2"));
        var mergeTask=new NotesLinkTask{Target=MigrationArea.Parse("AV1:AX4")};mergeTask.SourceMerges.Add(MigrationArea.Parse("D2:F2"));
        Check(NotesLinkRules.CanPasteWhole(MigrationArea.Parse("D1:F4"),mergeTask,mergeTarget),"native whole paste accepts matching source and target merge layouts");
        mergeTarget.Contents[MigrationSheetSnapshot.Key(2,49)]="value:hidden";
        Check(!NotesLinkRules.CanPasteWhole(MigrationArea.Parse("D1:F4"),mergeTask,mergeTarget),"native paste cannot overwrite preexisting content hidden inside a merge");
        mergeTarget.Contents.Clear();mergeTarget.Snapshot.Merges[0]=MigrationArea.Parse("AV3:AX3");
        Check(!NotesLinkRules.CanPasteWhole(MigrationArea.Parse("D1:F4"),mergeTask,mergeTarget),"different merge layouts use the compatibility writer");
        var heightPlan=new NotesLinkPlan{Original=new NotesSheetFacts{Snapshot=new MigrationSheetSnapshot(),StandardHeight=17}};heightPlan.Original.HiddenRows.Add(4);heightPlan.Tasks.Add(new NotesLinkTask{Target=MigrationArea.Parse("AV1:AX8")});
        using(var heightRun=new NotesLinkExecution(null,heightPlan,null))
        {
            var heightSheet=new NotesHeightSheet();typeof(NotesLinkExecution).GetField("sheet",BindingFlags.NonPublic|BindingFlags.Instance).SetValue(heightRun,heightSheet);
            typeof(NotesLinkExecution).GetMethod("RestoreGeometry",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(heightRun,null);
            Check(heightSheet.Range.Lookups.SequenceEqual(new[]{"1:3","5:8"}),"layout finishing excludes hidden rows and avoids unchanged height writes or Hidden setters");
        }
        Check(HiddenWorkbookApplication.HostProgId("Microsoft Excel",@"D:\WPS Office\office6","et")=="ket.Application","WPS compatibility name does not select the Excel automation backend");
        Check(HiddenWorkbookApplication.HostProgId("Microsoft Excel",@"D:\WPSOFF~1\office6","")=="ket.Application","WPS installation path identifies aliases without a process name");
        Check(HiddenWorkbookApplication.HostProgId("Microsoft Excel",@"C:\Office\Office16","EXCEL")=="Excel.Application","real Excel process retains its own backend");
        var map=typeof(ConsolidatedWorkbookForm).GetMethod("Areas",BindingFlags.NonPublic|BindingFlags.Static);
        var planned=(MigrationArea[])map.Invoke(null,new object[]{"K3,L3,M3,N3",15,true});
        Check(planned[0].Address=="K3:K17"&&planned[3].Address=="N3:N17","report field anchors expand to actual child count");
        var parentFields=ConsolidatedWorkbookForm.ParentAreas(planned);
        Check(parentFields.Select(a=>a.Address).SequenceEqual(new[]{"K2:K2","L2:L2","M2:M2","N2:N2"}),"group owner occupies row two separately from children");
        Reject(()=>ConsolidatedWorkbookForm.ParentAreas(new[]{MigrationArea.Parse("K2")}),"children cannot overwrite owner row");
        Check(ConsolidatedWorkbookForm.ParentAreas(new[]{MigrationArea.Parse("B5"),MigrationArea.Parse("D5")}).Select(a=>a.Column).SequenceEqual(new[]{2,4}),"owner follows mapped columns");
        bool overlap=false;try{map.Invoke(null,new object[]{"L3,L3,M3,N3",15,true});}catch(TargetInvocationException e){overlap=e.InnerException is InvalidOperationException;}Check(overlap,"overlapping report identity fields rejected");
        Check(SummaryTextReference.External(@"N:\EASDATA\SZYSKG_20261231","！01-示例公司.xlsx","BS!$G$10")=="='N:\\EASDATA\\SZYSKG_20261231\\[！01-示例公司.xlsx]BS'!$G$10","full file name preserves prefix and extension in external link");
        Check(SummaryTextReference.LinkFileName("sample.xlsm")=="sample.xlsm"&&SummaryTextReference.LinkFileName("sample")=="sample.xlsx","full Excel extension and legacy company name supported");
        var source=MigrationArea.Parse("D1:X1700");Check(source.Columns==21&&source.Rows==1700,"default bounded source has 21 columns and 1700 rows");
        Check(NotesLinkRules.Target(source,"AV1",22,1).Address=="BR1:CL1700","column spacing is difference between start columns");
        Check(NotesLinkRules.Target(source,"AV1",21,1).Address=="BQ1:CK1700","spacing equal to width leaves no blank column");
        Reject(()=>NotesLinkRules.Target(source,"AV1",20,0),"spacing narrower than company block rejected");
        Reject(()=>NotesLinkRules.Target(source,"AV1",0,0),"zero spacing rejected");
        Reject(()=>NotesLinkRules.Target(source,"XFC1",22,1),"spacing overflow rejected");
        string[] expected={"AV1:BP1700","BR1:CL1700","CN1:DH1700","DJ1:ED1700"};
        for(int i=0;i<expected.Length;i++)Check(NotesLinkRules.Target(source,"AV1","BR1",i).Address==expected[i],"fixed slot "+i);
        Check(NotesLinkRules.Target(source,"AV1","BR1",2).Address==expected[2],"unchecked middle company cannot shift later slot");
        Reject(()=>NotesLinkRules.Target(source,"AV1","AW1",0),"overlapping company blocks rejected");Reject(()=>NotesLinkRules.Target(source,"AV1","BR2",0),"different start rows rejected");Reject(()=>NotesLinkRules.Target(source,"BR1","AV1",0),"backwards stride rejected");Reject(()=>NotesLinkRules.Target(source,"XFC1","XFD1",0),"right boundary rejected");
        Check(NotesLinkRules.IsDirect("04-02-","04-02-01-")&&!NotesLinkRules.IsDirect("04-02-","04-02-01-01-")&&!NotesLinkRules.IsDirect("04-02-","04-020-01-"),"sequence segments select immediate children only");
        Check(NotesLinkRules.IsDirect("00-","04-")&&!NotesLinkRules.IsDirect("00-","04-02-")&&!NotesLinkRules.IsDirect("00-","00-"),"zero-level total uses first-level Lead codes without including self or grandchildren");
        Check(NotesLinkRules.Resolve(new[]{" Note "},"notes")==" Note ","Note alias preserves actual spaces and case");Reject(()=>NotesLinkRules.Resolve(new[]{"Notes 1","Notes 2"},"Notes"),"ambiguous sheet names require selection");
        Check(NotesLinkRules.Resolve(new[]{"Notes"," Notes "}," Notes ")==" Notes ","explicit sheet choice disambiguates names differing by spaces");
        Check(NotesLinkRules.Resolve(new[]{"Notes 2027"},"Notes")=="Notes 2027"&&!NotesLinkRules.ExactOrAlias("Notes 2027","Notes"),"unique fuzzy candidate still needs confirmation");
        Check(NotesLinkRules.ZeroFormat("#,##0.00")=="#,##0.00;-#,##0.00;\"-\";@","amount zero format preserves positive precision");Check(NotesLinkRules.ZeroFormat("0.00%")=="0.00%"&&NotesLinkRules.ZeroFormat("yyyy.m.d")=="yyyy.m.d","date and percentage formats preserved");
        Reject(()=>NotesLinkRules.ZeroFormat("[>1]0;0"),"conditional numeric formats require correction");
        Check(NotesLinkRules.ErrorText(unchecked((int)0x800A07D7))=="#DIV/0!"&&NotesLinkRules.ErrorText(new System.Runtime.InteropServices.ErrorWrapper(2023))=="#REF!","bulk error reporting decodes native error variants without per-cell COM reads");
        Check(DataMigrationEngine.SourceLink(@"C:\O'Brien\甲'公司.xlsx","Note's",1,4)=="='C:\\O''Brien\\[甲''公司.xlsx]Note''s'!$D$1","shared link generator escapes directory file and sheet apostrophes");
        var area=MigrationArea.Parse("A1:D3");var mask=new bool[3,4];mask[1,1]=mask[1,2]=true;var blocks=NotesLinkRules.Rectangles(area,mask);var visited=new HashSet<string>();
        foreach(var b in blocks)for(int r=b.Row;r<b.Row+b.Rows;r++)for(int c=b.Column;c<b.Column+b.Columns;c++)Check(!mask[r-1,c-1]&&visited.Add(r+":"+c),"rectangle covers each writable cell exactly once");Check(visited.Count==10,"rectangle decomposition excludes only blocked cells");
        var facts=new NotesSheetFacts{Snapshot=new MigrationSheetSnapshot()};var task=new NotesLinkTask{Target=MigrationArea.Parse("AV1:BP3")};task.SourceMerges.Add(MigrationArea.Parse("D2:X2"));facts.Snapshot.Merges.Add(MigrationArea.Parse("AV2:BP2"));NotesLinkRules.MapMerges(task,MigrationArea.Parse("D1:X3"),facts);Check(task.Anchors.Count==1&&task.Skipped==20,"aligned merged row uses one anchor without changing merge structure");
        var bad=new NotesLinkTask{Target=MigrationArea.Parse("AV1:BP3")};Reject(()=>NotesLinkRules.MapMerges(bad,MigrationArea.Parse("D1:X3"),facts),"target merge cannot swallow unrelated source cells");
        using(var form=new NotesLinksForm())Check(form.Text=="Notes 批量链接","Notes form constructs independently without an Excel connection");
        {
            var src=MigrationArea.Parse("D1:G1");var dst=MigrationArea.Parse("AV1:AY1");
            var parentFacts=new NotesSheetFacts{Snapshot=new MigrationSheetSnapshot()};parentFacts.Snapshot.Merges.Add(dst);
            var emptyFacts=new NotesSheetFacts();var emptyTask=new NotesLinkTask{Target=dst};
            NotesLinkRules.MapMerges(emptyTask,src,parentFacts,emptyFacts);
            Check(emptyTask.Blocks.Count==0&&emptyTask.Anchors.Count==1&&emptyTask.Skipped==3,"empty unmerged source maps only target merge anchor");
            emptyFacts.Contents[MigrationSheetSnapshot.Key(1,5)]="formula:1";
            Reject(()=>NotesLinkRules.MapMerges(new NotesLinkTask{Target=dst},src,parentFacts,emptyFacts),"source formula prevents merge mismatch data loss");
        }
        {
            string sample=Path.Combine(Path.GetTempPath(),"LeadMapping-"+Guid.NewGuid().ToString("N")+".xlsx");
            try
            {
                using(var stream=File.Create(sample))using(var zip=new System.IO.Compression.ZipArchive(stream,System.IO.Compression.ZipArchiveMode.Create))
                {
                    Action<string,string> put=(name,xml)=>{using(var writer=new StreamWriter(zip.CreateEntry(name).Open()))writer.Write(xml);};
                    put("xl/workbook.xml","<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='Lead' r:id='rId1'/></sheets></workbook>");
                    put("xl/_rels/workbook.xml.rels","<Relationships><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>");
                    put("xl/worksheets/sheet1.xml","<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData><row r='1'><c r='B1' t='inlineStr'><is><t>公司全称</t></is></c><c r='D1' t='inlineStr'><is><t>编码</t></is></c><c r='F1' t='inlineStr'><is><t>简称</t></is></c><c r='H1' t='inlineStr'><is><t>完整文件名</t></is></c><c r='J1' t='inlineStr'><is><t>不含编码文件名</t></is></c></row></sheetData></worksheet>");
                }
                Check(ConsolidatedTemplateProfile.Read(sample).Columns.SequenceEqual(new[]{4,6,2,8,10}),"consolidated template identifies relocated columns by header");
                Check(ConsolidatedTemplateProfile.Read(sample,new[]{1,3,5,7,9}).Columns.SequenceEqual(new[]{1,3,5,7,9}),"explicit consolidated column mapping is respected");
                Reject(()=>ConsolidatedTemplateProfile.Read(sample,new[]{1,1,2,3,4}),"duplicate consolidated column mapping rejected");
            }
            finally{if(File.Exists(sample))File.Delete(sample);}
        }
        using(var company=new CompanyWorkbookForm())
        {
            var inputs=(System.Windows.Forms.FlowLayoutPanel)typeof(CompanyWorkbookForm).GetField("inputs",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(company);
            int outputIndex=-1,confirmIndex=-1,structureIndex=-1;
            for(int i=0;i<inputs.Controls.Count;i++){var control=inputs.Controls[i];if(Convert.ToString(control.Tag)=="输出文件夹")outputIndex=i;if(Convert.ToString(control.Tag)=="核对")confirmIndex=i;if(control is CompanyStructurePicker)structureIndex=i;}
            Check(outputIndex>=0&&confirmIndex>outputIndex&&structureIndex>confirmIndex,"output directory and confirmation precede structure task grid");
        }
        using(var form=new NotesLinksForm())
        {
            Func<string,object> field=n=>typeof(NotesLinksForm).GetField(n,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
            var range=(System.Windows.Forms.TextBox)field("sourceRange");var rows=(System.Windows.Forms.NumericUpDown)field("sourceRows");
            range.Text="E3:Y102";Check(rows.Value==100,"manual source range synchronizes row count");rows.Value=200;Check(range.Text=="E3:Y202","advanced row count preserves source columns and first row");
            ((System.Windows.Forms.CheckBox)field("batchMode")).Checked=true;
            typeof(NotesLinksForm).GetMethod("AddBatchFiles",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{new[]{Path.Combine(Path.GetTempPath(),"sample-parent.xlsx"),Path.Combine(Path.GetTempPath(),"sample-other.xlsx")}});
            var batchGrid=(System.Windows.Forms.DataGridView)field("batchGrid");
            Check(batchGrid.Rows.Count==2&&batchGrid.Columns.Count==5,"batch form lists selected files with count status and backup columns");
            Check(((System.Windows.Forms.TextBox)field("parent")).ReadOnly,"batch shared parameters cannot accidentally rename the selected parent path");
            var items=(List<NotesLinkBatchItem>)field("batchItems");foreach(var item in items){item.Plan=new NotesLinkPlan{Config=new NotesLinkConfig{Parent=item.Parent}};item.Plan.Tasks.Add(new NotesLinkTask{Status="通过",File=Path.Combine(Path.GetTempPath(),"leaf.xlsx")});}
            var batchKey=typeof(NotesLinksForm).GetMethod("BatchKey",BindingFlags.Instance|BindingFlags.NonPublic);
            typeof(NotesLinksForm).GetField("batchPreparedKey",BindingFlags.Instance|BindingFlags.NonPublic).SetValue(form,batchKey.Invoke(form,null));
            typeof(NotesLinksForm).GetMethod("SetBusy",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,new object[]{false});
            Check(((System.Windows.Forms.Button)field("execute")).Enabled,"batch execution enables only with matching preflight settings");
            ((System.Windows.Forms.NumericUpDown)field("columnSpacing")).Value=23;
            Check(!((System.Windows.Forms.Button)field("execute")).Enabled,"changing shared spacing invalidates the whole prepared batch");


        }

    }
    private static void SaveFallbackChecks()
    {
        using(var sample=new ProbeSession("NotesSaveFallback"))
        {
            var normal=new NotesSaveFixture();Check(!NotesLinkExecution.SaveWithFallback(normal,sample.Output("unused.xlsx"),null)&&normal.SaveCalls==1&&normal.CopyCalls==0,"normal save does not perform a second save");
            var failed=new NotesSaveFixture{ErrorCode=unchecked((int)0x80004005)};string copy=sample.Output("fallback.xlsx");
            Check(NotesLinkExecution.SaveWithFallback(failed,copy,null)&&failed.CopyCalls==1&&new FileInfo(copy).Length>0,"E_FAIL alone triggers one SaveCopyAs fallback");
            var denied=new NotesSaveFixture{ErrorCode=unchecked((int)0x80070005)};bool rejected=false;
            try{NotesLinkExecution.SaveWithFallback(denied,sample.Output("denied.xlsx"),null);}catch(System.Runtime.InteropServices.COMException){rejected=true;}
            Check(rejected&&denied.CopyCalls==0,"access denied is reported without retrying or overwriting");
            var twice=new NotesSaveFixture{ErrorCode=unchecked((int)0x80004005),CopyFails=true};rejected=false;
            try{NotesLinkExecution.SaveWithFallback(twice,sample.Output("both-failed.xlsx"),null);}catch(System.Runtime.InteropServices.COMException){rejected=true;}
            Check(rejected&&twice.CopyCalls==1,"failed alternate save is not treated as success and does not loop");
            var absent=new NotesSaveFixture{ErrorCode=unchecked((int)0x80004005),NoOutput=true};rejected=false;
            try{NotesLinkExecution.SaveWithFallback(absent,sample.Output("missing-copy.xlsx"),null);}catch(IOException){rejected=true;}
            Check(rejected,"alternate save must produce a nonempty output before use");sample.Finish(0);
        }
    }
    private static void MarkFixtureHiddenRow(string path)
    {
        // WPS itself cannot set Hidden on an invisible automation workbook.
        // Create the own fixture's existing hidden row in its package instead.
        using(var stream=File.Open(path,FileMode.Open,FileAccess.ReadWrite))using(var zip=new System.IO.Compression.ZipArchive(stream,System.IO.Compression.ZipArchiveMode.Update))
        {
            Func<string,System.Xml.XmlDocument> read=name=>{var doc=new System.Xml.XmlDocument{XmlResolver=null};using(var input=zip.GetEntry(name).Open())doc.Load(input);return doc;};
            var workbook=read("xl/workbook.xml");var item=(System.Xml.XmlElement)workbook.SelectSingleNode("//*[local-name()='sheet'][@name='Notes']");string id=item.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            var rels=read("xl/_rels/workbook.xml.rels");var link=(System.Xml.XmlElement)rels.SelectSingleNode("//*[local-name()='Relationship'][@Id='"+id+"']");string part=link.GetAttribute("Target");part=part.StartsWith("/")?part.TrimStart('/'):"xl/"+part;
            var sheet=read(part);var row=(System.Xml.XmlElement)sheet.SelectSingleNode("//*[local-name()='row'][@r='4']");if(row==null)throw new Exception("Own fixture row 4 missing");row.SetAttribute("hidden","1");zip.GetEntry(part).Delete();using(var output=zip.CreateEntry(part).Open())sheet.Save(output);
        }
    }
    private static void WpsNativeLinks(ProbeSession session)
    {
        object source=session.Take(ProbeSession.Connect(new[]{"KET.Application","et.Application"},false));if(source==null)throw new Exception("WPS unavailable");
        string src=session.Output("03-01-native-source.xlsx"),parent=session.Output("03-native-parent.xlsx");
        using(var creator=new HiddenWorkbookApplication(source))using(var bag=new ComBag())
        {
            dynamic app=creator.Application,books=bag.Own(app.Workbooks);app.Calculation=-4135;app.CalculateBeforeSave=false;
            foreach(bool isSource in new[]{true,false})
            {
                dynamic book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
                try
                {
                    dynamic tabs=bag.Own(book.Worksheets),tab=bag.Own(tabs.Item(1));tab.Name="Notes";
                    string range=isSource?"D1:X1700":"AV1:BP1700";((dynamic)bag.Own(tab.Range[range])).Value2=isSource?7:99;((dynamic)bag.Own(tab.Range[range])).NumberFormat=isSource?"0.0%":"0.00";
                    ((dynamic)bag.Own(tab.Range[isSource?"D5:X5":"AV5:BP5"])).Merge();
                    ((dynamic)bag.Own(tab.Range[isSource?"D1:X1":"AV1:BP1"])).Merge();
                    if(isSource){((dynamic)bag.Own(tab.Range["G10"])).ClearContents();((dynamic)bag.Own(tab.Range["H10"])).Value2=0;}
                    else{((dynamic)bag.Own(tab.Range["AY10:AZ10"])).Formula="='[OldTemplate.xlsx]Notes'!G10";((dynamic)bag.Own(tab.Range["BR1:CL1"])).Merge();((dynamic)bag.Own(tab.Range["BR1:CL1700"])).NumberFormat="0.000";((dynamic)bag.Own(tab.Range["BU10"])).Formula="=123";((dynamic)bag.Own(tab.Range["A10"])).Formula="=SUM(AY10,BU10)";((dynamic)bag.Own(tab.Range["CN10"])).Value2="keep outside identified slots";}
                    if(!isSource){dynamic lead=bag.Own(tabs.Add());lead.Name="Lead";((dynamic)bag.Own(lead.Range["K2:N3"])).Value2=new object[,]{{"03-","parent","parent",Path.GetFileName(parent)},{"03-01-","child","child",Path.GetFileName(src)}};}
                    MigrationWindowScope.SetOwnedVisible((object)book,true);book.SaveAs(isSource?src:parent,51);MigrationWindowScope.SetOwnedVisible((object)book,false);app.Visible=false;
                }
                finally{book.Close(false);closed=true;}
            }
        }
        MarkFixtureHiddenRow(parent);
        string original=DataMigrationHost.Digest(parent),sourceHash=DataMigrationHost.Digest(src);
        var cfg=new NotesLinkConfig{Parent=parent,Folder=session.DirectoryPath};var plan=NotesLinkPlanner.Build(cfg,NotesLinkRules.ReadLead(cfg),Console.WriteLine,null);
        Check(plan.Tasks.Count==1&&plan.Tasks[0].NativeCompatible,"WPS production preflight selects whole-range paste for matching merges");
        var watch=System.Diagnostics.Stopwatch.StartNew();
        using(var run=new NotesLinkExecution(source,plan,Console.WriteLine))
        {foreach(string message in run.Run())Console.WriteLine(message);Check(run.Committed&&run.NativePastes==1&&run.FormulaWrites==35660,"WPS production execution commits whole native paste with effective link count");Check(File.Exists(run.Backup)&&DataMigrationHost.Digest(run.Backup)==original,"native whole paste retains a byte-identical parent backup");}
        Console.WriteLine("WPS_NATIVE_LINK_TOTAL_MS: "+watch.ElapsedMilliseconds);
        var after=NotesSheetFacts.Load(parent,"Notes");
        Check(plan.Cleanup!=null&&plan.Cleanup.Address=="BR1:CL1700","unused company slot boundary is explicit");
        Check(!after.Occupied(10,73)&&!after.Snapshot.Merges.Any(m=>m.Address=="BR1:CL1"),"unused right template formulas and merges cleared");
        Check(after.Content(10,1)=="formula:SUM(AY10,BU10)"&&after.Occupied(10,92),"left total and unrelated right content preserved");
        Check(after.HiddenRows.Contains(4),"native whole paste preserves original hidden row without setting Hidden during finishing");Check(after.Snapshot.Merges.Any(m=>m.Address=="AV5:BP5"),"native whole paste preserves target merge structure");
        Check(DataMigrationHost.Digest(src)==sourceHash,"native source is closed without saving and original source bytes stay unchanged");
        using(var reader=new HiddenWorkbookApplication(source))using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)reader.Application).Workbooks),book=books.Open(parent,0,true);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            try{dynamic tabs=bag.Own(book.Worksheets),tab=bag.Own(tabs.Item("Notes")),cell=bag.Own(tab.Range["AV1"]);Check(Convert.ToBoolean(cell.HasFormula)&&Convert.ToString(cell.NumberFormat)=="0.00","saved native result contains real formulas and preserves target number format");foreach(string address in new[]{"AY10","AZ10","BP1700"}){dynamic sample=bag.Own(tab.Range[address]);Check(Convert.ToBoolean(sample.HasFormula)&&Convert.ToString(sample.Formula).Contains("["+Path.GetFileName(src)+"]"),"body link replaced including blank/zero source: "+address);}
                Check(Convert.ToString(((dynamic)bag.Own(tab.Range["BU10"])).NumberFormat)=="General","unused template formatting cleared");
                var links=book.LinkSources(1) as Array;Check(links!=null&&links.Cast<object>().Any(x=>string.Equals(Convert.ToString(x),src,StringComparison.OrdinalIgnoreCase)),"native links point to the real source file rather than a temporary copy");}
            finally{book.Close(false);closed=true;}
        }
    }
    private static void WpsSave(ProbeSession session)
    {
        object source=session.Take(ProbeSession.Connect(new[]{"KET.Application","et.Application"},false));if(source==null)throw new Exception("WPS automation unavailable");
        using(var hidden=new HiddenWorkbookApplication(source))using(var bag=new ComBag())
        {
            dynamic app=hidden.Application,books=bag.Own(app.Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            try
            {
                app.Calculation=-4135;app.CalculateBeforeSave=false;
                dynamic sheets=bag.Own(book.Worksheets),tab=bag.Own(sheets.Item(1));tab.Name="Notes";
                ((dynamic)bag.Own(tab.Range["AV1:CL1700"])).Formula="=1+1";
                string path=session.Output("Wps-save.xlsx"),copy=session.Output("Wps-save-copy.xlsx");
                MigrationWindowScope.SetOwnedVisible((object)book,true);book.SaveAs(path,51);MigrationWindowScope.SetOwnedVisible((object)book,false);app.Visible=false;
                ((dynamic)bag.Own(tab.Range["A1"])).Value2="save-check";
                MigrationWindowScope.SetOwnedVisible((object)book,true);
                try{Check(!NotesLinkExecution.SaveWithFallback((object)book,copy,Console.WriteLine),"WPS saves 73100 formulas with its owned workbook view prepared");}
                finally{MigrationWindowScope.SetOwnedVisible((object)book,false);app.Visible=false;}
                Check(DataMigrationHost.PackageSheets(path).Contains("Notes"),"WPS direct saved package is readable");
                ((dynamic)bag.Own(tab.Range["A1"])).Value2="copy-check";
                var proxy=new NotesSaveFixture{NativeBook=(object)book,ErrorCode=unchecked((int)0x80004005)};
                MigrationWindowScope.SetOwnedVisible((object)book,true);
                try{Check(NotesLinkExecution.SaveWithFallback(proxy,copy,Console.WriteLine),"WPS native SaveCopyAs produces fallback after injected E_FAIL");}
                finally{MigrationWindowScope.SetOwnedVisible((object)book,false);app.Visible=false;}
                Check(DataMigrationHost.PackageSheets(copy).Contains("Notes"),"WPS fallback saved package is readable");
            }
            finally{book.Close(false);closed=true;}
            Check(!Convert.ToBoolean(app.Visible),"WPS save compatibility does not leave the owned background process visible");
        }
    }
    private static void MakeSource(object application,ProbeSession session,string path,string sheetName,int mergeRow,double value)
    {
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            dynamic sheets=bag.Own(book.Worksheets),sheet=bag.Own(sheets.Item(1));sheet.Name=sheetName;
            ((dynamic)bag.Own(sheet.Range["D1:F4"])).Value2=new object[,]{{value,value+1,value+2},{1.0,2.0,3.0},{4.0,5.0,6.0},{0.0,-1234.5,0.125}};
            dynamic merge=bag.Own(sheet.Range["D"+mergeRow+":F"+mergeRow]);merge.ClearContents();merge.Merge();((dynamic)bag.Own(sheet.Range["D"+mergeRow])).Value2="合并标题";
            if(mergeRow==2)((dynamic)bag.Own(sheet.Range["F3"])).Formula="=1/0";
            dynamic fullWidth=bag.Own(sheet.Range["D"+(10+mergeRow)+":X"+(10+mergeRow)]);fullWidth.Merge();((dynamic)bag.Own(sheet.Range["D"+(10+mergeRow)])).Value2="横跨D至X的标题";
            book.SaveAs(path,51);book.Close(false);closed=true;
        }
    }
    private static void BatchLive(object application,ProbeSession session,string template,NotesLinkConfig shared)
    {
        Func<string,bool,string> make=(name,two)=>
        {
            string path=session.Output(name+".xlsx");File.Copy(template,path);
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Open(path,0,false);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
                try{dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item("Lead"));((dynamic)bag.Own(lead.Range["K5:O5"])).ClearContents();if(!two)((dynamic)bag.Own(lead.Range["K6:O6"])).ClearContents();book.Save();}
                finally{book.Close(false);closed=true;}
            }
            return path;
        };
        string first=make("批量两家公司",true),second=make("批量一家公司",false),fail=make("批量运行失败",false),missing=session.Output("批量缺失.xlsx");
        string firstHash=DataMigrationHost.Digest(first),secondHash=DataMigrationHost.Digest(second);
        var input=new[]{new NotesLinkBatchItem{Parent=missing},new NotesLinkBatchItem{Parent=fail},new NotesLinkBatchItem{Parent=first},new NotesLinkBatchItem{Parent=second}};
        string childReference=Path.Combine(shared.Folder,"04-02-01-甲'公司.xlsx");
        var referenceFiles=new[]{new NotesLinkBatchItem{Parent=childReference},new NotesLinkBatchItem{Parent=first},new NotesLinkBatchItem{Parent=second}};
        Check(NotesLinkBatch.PickerParent(shared,referenceFiles,childReference)==first,"batch owner picker skips a selected child and resolves an actual summary workbook");
        var mistakenHeader=shared.Copy();mistakenHeader.ParentCodeCell="K1";
        Check(NotesLinkBatch.PickerParent(mistakenHeader,referenceFiles,childReference)==first&&mistakenHeader.ParentCodeCell=="K1","reference identification can recover from a wrong owner cell without changing execution settings");
        var prepared=NotesLinkBatch.Prepare(shared,input,false,Console.WriteLine,null);
        Check(prepared.Count==4&&prepared.Count(i=>i.Ready)==3&&!prepared[0].Ready,"batch preflight isolates missing parent without blocking valid files");
        Check(prepared.Single(i=>i.Parent==first).Plan.Tasks.Count==2&&prepared.Single(i=>i.Parent==second).Plan.Tasks.Count==1,"batch independently reads each parent company count under shared settings");
        Check(DataMigrationHost.Digest(first)==firstHash&&DataMigrationHost.Digest(second)==secondHash,"unified batch preflight does not modify parent workbooks");
        File.AppendAllText(fail,"changed-after-preflight");string failedHash=DataMigrationHost.Digest(fail);
        var run=new NotesLinkBatchRun(application,prepared,Console.WriteLine);foreach(string message in run.Run())Console.WriteLine(message);
        Check(run.Succeeded==2&&run.Failed==1&&run.Completed==3,"runtime failure continues to the next valid parent");
        Check(DataMigrationHost.Digest(fail)==failedHash,"failed batch parent is not overwritten");
        Check(prepared.Where(i=>i.Parent==first||i.Parent==second).All(i=>File.Exists(i.Backup)&&!string.IsNullOrEmpty(i.CommittedHash)),"each successful parent retains an independent recovery backup");
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Open(second,0,true);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            try{dynamic sheets=bag.Own(book.Worksheets),tab=bag.Own(sheets.Item("Notes"));Check(Convert.ToBoolean(((dynamic)bag.Own(tab.Range["AV1"])).HasFormula),"single-company batch output contains real link formula");Check(Convert.ToString(((dynamic)bag.Own(tab.Range["AZ1"])).Value2)=="未勾选公司保留","smaller company count cannot inherit the preceding parent write range");}
            finally{book.Close(false);closed=true;}
        }
        var saved=prepared.Single(i=>i.Parent==first);MigrationOutputTransaction.Restore(first,saved.Backup,saved.CommittedHash);Check(DataMigrationHost.Digest(first)==firstHash,"one batch parent can be restored independently");
        string cancelFirst=make("批量停止前完成",false),cancelNext=make("批量停止未执行",true),nextHash=DataMigrationHost.Digest(cancelNext);
        var cancelPlans=NotesLinkBatch.Prepare(shared,new[]{new NotesLinkBatchItem{Parent=cancelFirst},new NotesLinkBatchItem{Parent=cancelNext}},false,Console.WriteLine,null);
        var stopping=new NotesLinkBatchRun(application,cancelPlans,Console.WriteLine);foreach(string message in stopping.Run()){if(stopping.Succeeded==1)stopping.Cancelled=true;}
        Check(stopping.Succeeded==1&&stopping.Failed==0&&DataMigrationHost.Digest(cancelNext)==nextHash,"stop between parents retains prior commit and leaves remaining file untouched");
        Check(cancelPlans[1].Status.StartsWith("未执行"),"unstarted batch file has an explicit stopped status");
    }
    private static bool Hidden(NotesLinkExecution execution)
    {
        var hidden=(HiddenWorkbookApplication)typeof(NotesLinkExecution).GetField("hidden",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(execution);
        return hidden==null||!Convert.ToBoolean(((dynamic)hidden.Application).Visible);
    }
    private static void Live(object application,ProbeSession session)
    {
        string sources=Path.Combine(session.DirectoryPath,"O'Brien");Directory.CreateDirectory(sources);
        string a=Path.Combine(sources,"04-02-01-甲'公司.xlsx"),b=Path.Combine(sources,"04-02-03-乙.xlsx"),parent=session.Output("04-02-父公司（合并）.xlsx");
        MakeSource(application,session,a,"Note",2,10);MakeSource(application,session,b,"nOtEs",3,100);
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Add();bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});
            dynamic sheets=bag.Own(book.Worksheets),notes=bag.Own(sheets.Item(1)),lead=bag.Own(sheets.Add());notes.Name="Notes";lead.Name="Lead";
            ((dynamic)bag.Own(lead.Range["K1:N1"])).Value2=new object[,]{{"编码","简称","全称","文件名称"}};
            ((dynamic)bag.Own(lead.Range["K2:O6"])).Value2=new object[,]{{"04-02-","父公司","父公司","04-02-父公司（合并）.xlsx","父公司（合并）.xlsx"},{"04-02-01-","本部","本部",Path.GetFileName(a),"甲'公司.xlsx"},{"04-02-01-01-","孙公司","孙公司","孙公司.xlsx","孙公司.xlsx"},{"04-02-02-","缺少文件","缺少文件","04-02-02-缺少文件.xlsx","缺少文件.xlsx"},{"04-02-03-","乙公司","乙公司",Path.GetFileName(b),"乙.xlsx"}};
            ((dynamic)bag.Own(notes.Range["A1"])).Value2="范围外保留";((dynamic)bag.Own(notes.Range["AY1"])).Value2="公司间隔保留";((dynamic)bag.Own(notes.Range["AZ1"])).Value2="未勾选公司保留";
            ((dynamic)bag.Own(notes.Range["AV1"])).Value2=999;dynamic merge=bag.Own(notes.Range["AV2:AX2"]);merge.Merge();((dynamic)bag.Own(notes.Range["AV2"])).Value2="旧合并标题";
            ((dynamic)bag.Own(notes.Range["AV4:AW4"])).NumberFormat="#,##0.00";((dynamic)bag.Own(notes.Range["AX4"])).NumberFormat="0.00%";
            ((dynamic)bag.Own(notes.Range["AV:AW"])).ColumnWidth=18;((dynamic)bag.Own(notes.Range["4:4"])).RowHeight=24;
            book.SaveAs(parent,51);book.Close(false);closed=true;
        }
        string original=DataMigrationHost.Digest(parent),sourceA=DataMigrationHost.Digest(a),sourceB=DataMigrationHost.Digest(b);
        var cfg=new NotesLinkConfig{Parent=parent,Folder=sources,SourceRange="D1:F4",FirstStart="AV1",ColumnSpacing=4,NativePaste=false};
        var headerParent=cfg.Copy();headerParent.ParentCodeCell="K1";
        try{NotesLinkRules.ReadLead(headerParent);throw new Exception("parent header must fail");}catch(InvalidOperationException e){Check(e.Message.Contains("Lead!K1")&&e.Message.Contains("汇总公司编码"),"production Lead reader locates wrong parent cell");}
        var headerRange=cfg.Copy();headerRange.FirstRow=1;
        try{NotesLinkRules.ReadLead(headerRange);throw new Exception("header row must fail");}catch(InvalidOperationException e){Check(e.Message.Contains("Lead!K1")&&e.Message.Contains("标题行"),"production Lead reader locates header included in company range");}
        var wrongColumn=cfg.Copy();wrongColumn.CodeColumn="M";
        try{NotesLinkRules.ReadLead(wrongColumn);throw new Exception("name column must fail");}catch(InvalidOperationException e){Check(e.Message.Contains("Lead!M3")&&e.Message.Contains("本部"),"production Lead reader locates company name mistaken for code");}

        var rows=NotesLinkRules.ReadLead(cfg);Check(rows.Count==3&&rows[0].Slot==0&&rows[2].Slot==2&&rows.All(t=>t.Code!="04-02-01-01-"),"Lead reader keeps direct order including body and excludes grandchild");
        var bad=NotesLinkPlanner.Build(cfg,rows,Console.WriteLine,null);Check(bad.Tasks.Count(t=>t.Status.StartsWith("阻止"))==1,"missing source reported in preflight");Check(DataMigrationHost.Digest(parent)==original,"preflight does not modify parent");
        rows[1].Selected=false;Func<NotesLinkPlan> prepare=()=>NotesLinkPlanner.Build(cfg,rows,Console.WriteLine,null);
        var plan=prepare();Check(plan.Tasks.Count==2&&plan.Tasks[1].Target.Address=="BD1:BF4","unchecking missing middle source leaves original slot intact");Check(plan.Occupied>0&&plan.Tasks.Sum(t=>t.Skipped)==4,"overwrite and structural merge skips explicitly counted");
        var defaultConfig=cfg.Copy();defaultConfig.SourceRange="D1:X1700";defaultConfig.ColumnSpacing=22;
        Check(NotesLinkPlanner.Build(defaultConfig,rows,Console.WriteLine,null).Tasks.All(t=>t.Status.StartsWith("通过")),"D:X merge markers on different row numbers are detected automatically");
        var noMarker=defaultConfig.Copy();noMarker.SourceRange="D1:Y1700";
        var structureCheck=NotesLinkPlanner.Build(noMarker,rows,Console.WriteLine,null);Check(structureCheck.Tasks.All(t=>t.Status.StartsWith("阻止")),"missing full-width structural marker requires explicit confirmation");
        var confirmed=rows.Select(t=>t.Copy()).ToList();foreach(var item in confirmed)item.Confirmed=true;
        Check(NotesLinkPlanner.Build(noMarker,confirmed,Console.WriteLine,null).Tasks.All(t=>t.Status.StartsWith("通过")),"manual structure confirmation allows a customized source range");
        string backup,committedHash;string originalAmountFormat;
        using(var sampleRefs=new ComBag()){dynamic books=sampleRefs.Own(((dynamic)application).Workbooks),sampleBook=sampleRefs.Own(books.Open(parent,0,true));try{dynamic sheets=sampleRefs.Own(sampleBook.Worksheets),tab=sampleRefs.Own(sheets.Item("Notes")),cell=sampleRefs.Own(tab.Range["AV4"]);originalAmountFormat=Convert.ToString(cell.NumberFormat);}finally{sampleBook.Close(false);}}
        object openSource;bool sourceClosed=false;
        using(var bag=new ComBag()){dynamic books=bag.Own(((dynamic)application).Workbooks);openSource=books.Open(a,0,false);session.Own(openSource,delegate{if(!sourceClosed)((dynamic)openSource).Close(false);});}
        using(var pickerHost=new DataMigrationHost(application))
        {
            string selected=pickerHost.PickSourceRange(parent,"Lead","K2",target=>
            {
                using(var chosenRefs=new ComBag())
                {
                    dynamic activeSheet=chosenRefs.Own(((dynamic)application).ActiveSheet),selection=chosenRefs.Own(((dynamic)application).Selection);
                    Check(WorksheetIdentity.SameObject((object)activeSheet,target),"picker reactivates exact parent worksheet after foreground callback activates child");
                    Check(Convert.ToInt32(selection.Row)==2&&Convert.ToInt32(selection.Column)==11,"owner picker starts on K2 instead of header A1");
                }
                return ((dynamic)target).Range["K2"];
            },"汇总公司编码",restore=>{((dynamic)openSource).Activate();using(var one=new ComBag()){dynamic sheets=one.Own(((dynamic)openSource).Worksheets),tab=one.Own(sheets.Item(1));tab.Activate();}});
            Check(selected=="K2:K2","parent range selection returns owner cell coordinates");
        }
        Check(DataMigrationHost.Digest(parent)==original&&DataMigrationHost.Digest(a)==sourceA,"owner parameter picking does not modify parent or child files");
        using(var execution=new NotesLinkExecution(application,plan,Console.WriteLine))
        {
            foreach(var message in execution.Run()){Console.WriteLine(message);Check(Hidden(execution),"worker remains invisible");}
            Check(execution.Committed&&execution.FormulaWrites>0,"complete verified result committed");backup=execution.Backup;committedHash=execution.CommittedHash;
            Check(execution.ValueErrors==0&&execution.LinkWarnings==0,"paste-only execution skips link refresh and value inspection");
            Check(File.Exists(execution.PreBackup)&&File.Exists(backup)&&DataMigrationHost.Digest(backup)==original,"pre-execution and atomic-commit backups preserve original bytes");
        }
        Check(Convert.ToBoolean(((dynamic)openSource).Saved)&&Convert.ToString(((dynamic)openSource).FullName)==a,"an already-open saved source remains open and unmodified");((dynamic)openSource).Close(false);sourceClosed=true;
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Open(parent,0,true);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});dynamic sheets=bag.Own(book.Worksheets),notes=bag.Own(sheets.Item("Notes"));
            Check(Convert.ToString(((dynamic)bag.Own(notes.Range["AV1"])).Formula)==DataMigrationEngine.SourceLink(a,"Note",1,4),"actual formula uses shared escaped full-path link");
            Check(Convert.ToString(((dynamic)bag.Own(notes.Range["BD1"])).Formula)==DataMigrationEngine.SourceLink(b,"nOtEs",1,4),"last company written under original fixed slot");
            Check(Convert.ToBoolean(((dynamic)bag.Own(notes.Range["AV2"])).MergeCells)&&!Convert.ToBoolean(((dynamic)bag.Own(notes.Range["AW2"])).HasFormula),"existing merge kept and follower not assigned a formula");
            Check(Convert.ToString(((dynamic)bag.Own(notes.Range["AY1"])).Value2)=="公司间隔保留"&&Convert.ToString(((dynamic)bag.Own(notes.Range["AZ1"])).Value2)=="未勾选公司保留"&&Convert.ToString(((dynamic)bag.Own(notes.Range["A1"])).Value2)=="范围外保留","gap unchecked company and outside cells untouched");
            Check(Convert.ToString(((dynamic)bag.Own(notes.Range["AV4"])).NumberFormat)==originalAmountFormat,"paste-only execution keeps original numeric format");
            Check(Convert.ToString(((dynamic)bag.Own(notes.Range["AX4"])).NumberFormat)=="0.00%","negative amount and percent display preserved");
            Check(Convert.ToDouble(((dynamic)bag.Own(notes.Range["AV:AV"])).ColumnWidth)==18&&Convert.ToDouble(((dynamic)bag.Own(notes.Range["4:4"])).RowHeight)==24,"row height and column width unchanged");book.Close(false);closed=true;
        }
        MigrationOutputTransaction.Restore(parent,backup,committedHash);Check(DataMigrationHost.Digest(parent)==original,"restore recovers original parent byte for byte");
        string alternate=session.Output("另一种Lead布局.xlsx");File.Copy(parent,alternate);
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Open(alternate,0,false);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});dynamic sheets=bag.Own(book.Worksheets),lead=bag.Own(sheets.Item("Lead"));
            dynamic from=bag.Own(lead.Range["K2:O6"]),to=bag.Own(lead.Range["B2"]);from.Copy(to);book.Save();book.Close(false);closed=true;
        }
        var moved=cfg.Copy();moved.Parent=alternate;moved.ParentCodeCell="B2";moved.CodeColumn="B";moved.NameColumn="D";moved.FileColumn="E";moved.LastRow=6;
        Check(NotesLinkRules.ReadLead(moved).Count==3,"Lead code name filename columns and parent identity are user-configurable");
        string arrays=session.Output("数组冲突.xlsx");File.Copy(parent,arrays);
        using(var bag=new ComBag())
        {
            dynamic books=bag.Own(((dynamic)application).Workbooks),book=books.Open(arrays,0,false);bool closed=false;session.Own((object)book,delegate{if(!closed)book.Close(false);});dynamic sheets=bag.Own(book.Worksheets),notes=bag.Own(sheets.Item("Notes"));
            ((dynamic)bag.Own(notes.Range["AV1:AW1"])).FormulaArray="=1";book.Save();book.Close(false);closed=true;
        }
        var arrayCfg=cfg.Copy();arrayCfg.Parent=arrays;var blockedArray=NotesLinkPlanner.Build(arrayCfg,rows,Console.WriteLine,null);
        Check(blockedArray.Tasks.Any(t=>t.Status.Contains("数组")),"legacy array destination is blocked by read-only package preflight");
        var boundaryCfg=cfg.Copy();boundaryCfg.FirstStart="AW2";boundaryCfg.ColumnSpacing=4;
        Check(NotesLinkPlanner.Build(boundaryCfg,rows,Console.WriteLine,null).Tasks.Any(t=>t.Status.StartsWith("阻止")),"partial intersection with an existing merge is blocked before writing");
        bool failed=false;
        using(var execution=new NotesLinkExecution(application,prepare(),Console.WriteLine){FailureForProbe=n=>n>0})
        {try{foreach(var ignored in execution.Run()){} }catch(InvalidOperationException){failed=true;}}
        Check(failed&&DataMigrationHost.Digest(parent)==original,"injected write failure never modifies original parent");
        bool cancelled=false;
        using(var execution=new NotesLinkExecution(application,prepare(),Console.WriteLine))
        {try{foreach(var ignored in execution.Run())if(execution.FormulaWrites>0)execution.Cancelled=true;}catch(OperationCanceledException){cancelled=true;}}
        Check(cancelled&&DataMigrationHost.Digest(parent)==original,"cancellation after first block discards all partial writes");
        Check(DataMigrationHost.Digest(a)==sourceA&&DataMigrationHost.Digest(b)==sourceB,"all source workbooks unchanged");Check(Directory.GetFiles(session.DirectoryPath,".xy-notes-*.xlsx").Length==0,"normal success failure and cancellation release staging files");
        BatchLive(application,session,parent,cfg);
        if(scale)
        {
            foreach(var row in confirmed)row.Selected=row.Slot==0;
            var large=NotesLinkPlanner.Build(defaultConfig,confirmed,Console.WriteLine,null);var watch=System.Diagnostics.Stopwatch.StartNew();
            using(var execution=new NotesLinkExecution(application,large,Console.WriteLine))
            {foreach(var ignored in execution.Run()){}Check(execution.Committed&&execution.FormulaWrites==35700-22,"default 1700 by 21 source executes in bulk while excluding merge followers");Console.WriteLine("DEFAULT_RANGE_MS: "+watch.ElapsedMilliseconds);}
        }
    }
    [STAThread]private static int Main(string[] args)
    {
        using(var session=new ProbeSession("NotesLinksProbe"))
        {
            string previous=DialogSizeStore.RootOverride;DialogSizeStore.RootOverride=session.DirectoryPath;
            try
            {
                Pure();SaveFallbackChecks();scale=Array.IndexOf(args,"--scale")>=0;
                if(Array.IndexOf(args,"--wps-copy-benchmark")>=0){NotesCopySpeedProbe.Run(session);return session.Finish(0);}
                if(Array.IndexOf(args,"--wps-native-links")>=0){WpsNativeLinks(session);Console.WriteLine("PASS: "+count+" checks including WPS native link execution");return session.Finish(0);}
                if(Array.IndexOf(args,"--wps-save")>=0){WpsSave(session);Console.WriteLine("PASS: "+count+" checks including WPS save compatibility");return session.Finish(0);}
                if(Array.IndexOf(args,"--render")>=0)using(var form=new NotesLinksForm())
                {
                    if(Array.IndexOf(args,"--batch-render")>=0)
                    {
                        typeof(NotesLinksForm).GetField("batchMode",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form).GetType().GetProperty("Checked").SetValue(typeof(NotesLinksForm).GetField("batchMode",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form),true,null);
                        typeof(NotesLinksForm).GetMethod("AddBatchFiles",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,new object[]{new[]{Path.Combine(session.DirectoryPath,"汇总文件A.xlsx"),Path.Combine(session.DirectoryPath,"汇总文件B.xlsx"),Path.Combine(session.DirectoryPath,"汇总文件C.xlsx")}});
                        var items=(List<NotesLinkBatchItem>)typeof(NotesLinksForm).GetField("batchItems",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(form);
                        for(int i=0;i<items.Count;i++){items[i].Status=i==2?"阻止：缺少一个明细文件":"通过｜公司 "+(i+1)+" 家";if(i<2){items[i].Plan=new NotesLinkPlan();for(int n=0;n<=i;n++)items[i].Plan.Tasks.Add(new NotesLinkTask{Status="通过"});}}
                        typeof(NotesLinksForm).GetMethod("RefreshBatchRows",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,null);
                    }
                    form.ShowInTaskbar=false;form.Opacity=0;form.Size=new System.Drawing.Size(1280,860);form.Show();form.PerformLayout();System.Windows.Forms.Application.DoEvents();
                    var source = (System.Windows.Forms.ComboBox)typeof(NotesLinksForm).GetField("sourceSheet", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    var target = (System.Windows.Forms.ComboBox)typeof(NotesLinksForm).GetField("targetSheet", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    var row = (System.Windows.Forms.TableLayoutPanel)source.Parent.Parent;
                    double center = source.Parent.Top + source.Top + source.Height / 2.0;
                    Check(target.Parent.Parent == row && row.Controls.Count == 7 && source.Width == target.Width,
                        "worksheet selectors share one row and equal input widths");
                    var positions = row.Controls.Cast<System.Windows.Forms.Control>()
                        .Select(c => c.Text + "=" + (c.Top + c.Height / 2.0)).ToArray();
                    Check(row.Controls.Cast<System.Windows.Forms.Control>().All(c => Math.Abs(c.Top + c.Height / 2.0 - center) <= 1),
                        "worksheet labels, selectors and buttons are vertically centered: " + string.Join(", ", positions));
                    using(var bitmap=new System.Drawing.Bitmap(form.Width,form.Height)){form.DrawToBitmap(bitmap,new System.Drawing.Rectangle(0,0,form.Width,form.Height));string image=Path.Combine(session.DirectoryPath,"NotesLinksForm.png");bitmap.Save(image,System.Drawing.Imaging.ImageFormat.Png);Console.WriteLine("PREVIEW: "+image);}form.Hide();
                }
                if(Array.IndexOf(args,"--live")>=0)
                {
                    Type type=Type.GetTypeFromProgID("Excel.Application",false);if(type==null){Console.WriteLine("SKIP Excel.Application unavailable");return session.Finish(2);}
                    dynamic app=session.Take(Activator.CreateInstance(type));liveHost=HiddenWorkbookApplication.IsWps((object)app)?"WPS (Excel-compatible ProgID)":"Microsoft Excel";
                    if(!session.Override(()=>app.DisplayAlerts,v=>app.DisplayAlerts=v,false,"DisplayAlerts")||!session.Override(()=>app.EnableEvents,v=>app.EnableEvents=v,false,"EnableEvents"))throw new Exception("Cannot establish safe test host state");
                    Live((object)app,session);
                }
                Console.WriteLine("PASS: "+count+" Notes checks"+(Array.IndexOf(args,"--live")<0?" (offline only)":" including "+liveHost));return session.Finish(0);
            }
            catch(Exception error){Console.WriteLine(error);return session.Finish(1);}
            finally{DialogSizeStore.RootOverride=previous;}
        }
    }
}

public sealed class NotesSaveFixture
{
    public int ErrorCode,SaveCalls,CopyCalls;
    public bool CopyFails,NoOutput;
    public object NativeBook;
    public void Save(){SaveCalls++;if(ErrorCode!=0)throw new System.Runtime.InteropServices.COMException("injected save error",ErrorCode);if(NativeBook!=null)((dynamic)NativeBook).Save();}
    public void SaveCopyAs(string path){CopyCalls++;if(CopyFails)throw new System.Runtime.InteropServices.COMException("injected alternate save error",unchecked((int)0x80004005));if(NoOutput)return;if(NativeBook!=null)((dynamic)NativeBook).SaveCopyAs(path);else File.WriteAllBytes(path,new byte[]{1});}
}

internal static class NotesCopySpeedProbe
{
    internal static void Run(ProbeSession session)
    {
        object source=session.Take(ProbeSession.Connect(new[]{"KET.Application","et.Application"},false));if(source==null)throw new Exception("WPS unavailable");
        var previousClipboard=System.Windows.Forms.Clipboard.GetDataObject();
        using(var hidden=new HiddenWorkbookApplication(source))using(var refs=new ComBag())
        {
            var owned=new List<object>();var closed=new HashSet<object>();
            Action<object> own=b=>{owned.Add(b);session.Own(b,delegate{if(!closed.Contains(b))((dynamic)b).Close(false);});};
            dynamic app=hidden.Application,books=refs.Own(app.Workbooks);app.Calculation=-4135;app.CalculateBeforeSave=false;
            try
            {
                string src=session.Output("BenchSource.xlsx"),original=session.Output("BenchTarget.xlsx"),native=session.Output("BenchNative.xlsx"),direct=session.Output("BenchDirect.xlsx");
                dynamic sourceBook=books.Add();own((object)sourceBook);dynamic sheets=refs.Own(sourceBook.Worksheets),sourceSheet=refs.Own(sheets.Item(1));sourceSheet.Name="Notes";
                ((dynamic)refs.Own(sourceSheet.Range["D1:X1700"])).Value2=7;((dynamic)refs.Own(sourceSheet.Range["D1:X1700"])).NumberFormat="0.0%";
                for(int r=2;r<=1200;r+=4)((dynamic)refs.Own(sourceSheet.Range["D"+r+":E"+r])).Merge();
                MigrationWindowScope.SetOwnedVisible((object)sourceBook,true);sourceBook.SaveAs(src,51);MigrationWindowScope.SetOwnedVisible((object)sourceBook,false);app.Visible=false;
                dynamic targetBook=books.Add();own((object)targetBook);dynamic targetSheets=refs.Own(targetBook.Worksheets),target=refs.Own(targetSheets.Item(1));target.Name="Notes";
                ((dynamic)refs.Own(target.Range["AV1:BP1700"])).Value2=3;((dynamic)refs.Own(target.Range["AV1:BP1700"])).NumberFormat="0.00";
                for(int r=2;r<=1200;r+=4)((dynamic)refs.Own(target.Range["AV"+r+":AW"+r])).Merge();
                MigrationWindowScope.SetOwnedVisible((object)targetBook,true);targetBook.SaveAs(original,51);MigrationWindowScope.SetOwnedVisible((object)targetBook,false);app.Visible=false;
                File.Copy(original,native);File.Copy(original,direct);
                var from=MigrationArea.Parse("D1:X1700");var task=new NotesLinkTask{File=src,Sheet="Notes",Target=MigrationArea.Parse("AV1:BP1700")};
                var sourceFacts=NotesSheetFacts.Load(src,"Notes");task.SourceMerges=sourceFacts.Snapshot.Merges;NotesLinkRules.MapMerges(task,from,NotesSheetFacts.Load(original,"Notes"),sourceFacts);
                Console.WriteLine("BENCH_SHAPE: blocks="+task.Blocks.Count+" anchors="+task.Anchors.Count);
                var watch=System.Diagnostics.Stopwatch.StartNew();int calls=0;
                foreach(var part in task.Blocks.Concat(task.Anchors))using(var one=new ComBag()){dynamic area=one.Own(target.Range[part.Address]);DataMigrationEngine.AssignFormula((object)area,DataMigrationEngine.LinkMatrix(src,"Notes",part.Row-task.Target.Row+1,part.Column-task.Target.Column+4,part.Rows,part.Columns));calls++;}
                Console.WriteLine("BENCH_BASELINE_MS: "+watch.ElapsedMilliseconds+" calls="+calls);
                dynamic directBook=books.Open(direct,0,false);own((object)directBook);dynamic directSheets=refs.Own(directBook.Worksheets),directSheet=refs.Own(directSheets.Item("Notes"));
                watch.Restart();calls=0;
                foreach(var part in task.Blocks.Concat(task.Anchors))using(var one=new ComBag()){dynamic area=one.Own(directSheet.Range[part.Address]);area.Formula=DataMigrationEngine.LinkMatrix(src,"Notes",part.Row-task.Target.Row+1,part.Column-task.Target.Column+4,part.Rows,part.Columns);calls++;}
                Console.WriteLine("BENCH_DIRECT_MS: "+watch.ElapsedMilliseconds+" calls="+calls);
                dynamic nativeBook=books.Open(native,0,false);own((object)nativeBook);dynamic nativeSheets=refs.Own(nativeBook.Worksheets),nativeSheet=refs.Own(nativeSheets.Item("Notes"));
                try
                {
                    watch.Restart();sourceBook.Activate();sourceSheet.Activate();((dynamic)refs.Own(sourceSheet.Range["D1:X1700"])).Copy();nativeBook.Activate();nativeSheet.Activate();((dynamic)refs.Own(nativeSheet.Range["AV1"])).Select();nativeSheet.Paste(Type.Missing,true);
                    Console.WriteLine("BENCH_NATIVE_MS: "+watch.ElapsedMilliseconds);
                    dynamic start=refs.Own(nativeSheet.Range["AV1"]),last=refs.Own(nativeSheet.Range["BP1700"]),merge=refs.Own(nativeSheet.Range["AV2:AW2"]),all=refs.Own(nativeSheet.Range["AV1:BP1700"]),formulas=refs.Own(all.SpecialCells(-4123));
                    Console.WriteLine("BENCH_NATIVE_RESULT: first="+Convert.ToString(start.Formula)+" last="+Convert.ToString(last.Formula)+" merged="+Convert.ToString(merge.MergeCells)+" numberFormat="+Convert.ToString(start.NumberFormat)+" formulas="+Convert.ToInt64(formulas.Count));
                }
                catch(Exception e){Console.WriteLine("BENCH_NATIVE_UNAVAILABLE: "+e.Message);}
            }
            finally
            {
                try{app.CutCopyMode=false;}catch{}
                for(int i=owned.Count-1;i>=0;i--){((dynamic)owned[i]).Close(false);closed.Add(owned[i]);}
                if(previousClipboard==null)System.Windows.Forms.Clipboard.Clear();else System.Windows.Forms.Clipboard.SetDataObject(previousClipboard,true);
            }
        }
    }
}

public sealed class NotesHeightSheet {public NotesHeightRanges Range=new NotesHeightRanges();}
public sealed class NotesHeightRanges
{
    public readonly List<string> Lookups=new List<string>();
    public NotesHeightRange this[string address]{get{Lookups.Add(address);var parts=address.Split(':');if(int.Parse(parts[0])<=4&&int.Parse(parts[1])>=4)throw new Exception("hidden row must not be touched");return new NotesHeightRange();}}
}
public sealed class NotesHeightRange {public object RowHeight{get{return 17.0;}set{throw new Exception("unchanged height must not be set");}}}
