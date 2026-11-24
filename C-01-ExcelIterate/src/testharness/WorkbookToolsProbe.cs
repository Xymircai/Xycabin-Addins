using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Excel_Iterate;
partial class WorkbookToolsProbe {
    static int checks; static void Check(bool v,string name){if(!v)throw new Exception(name);checks++;}
    static void Reject(Action a,string name){bool refused=false;try{a();}catch{refused=true;}Check(refused,name);}
    static void Package(string file,params string[] links) {
        using(FileStream output=File.Create(file))using(ZipArchive zip=new ZipArchive(output,ZipArchiveMode.Create)) {
            Part(zip,"xl/workbook.xml","<workbook/>");
            Part(zip,"xl/worksheets/sheet1.xml","<worksheet><cell>original formula and cached value</cell></worksheet>");
            Part(zip,"customXml/item1.xml","<content>must remain byte-for-byte</content>");
            for(int i=0;i<links.Length;i++)Part(zip,"xl/externalLinks/_rels/externalLink"+(i+1)+".xml.rels","<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" Target=\""+System.Security.SecurityElement.Escape(links[i])+"\" TargetMode=\"External\"/></Relationships>");
        }
    }
    static void Part(ZipArchive zip,string name,string data){using(StreamWriter w=new StreamWriter(zip.CreateEntry(name).Open(),new UTF8Encoding(false)))w.Write(data);}
    static LinkAssistantSession JournalSession(string report, string[] targets, string oldPath, string newPath, bool opened) {
        var sheet=new LinkFakeSheet(); var reportBook=new LinkFakeBook{FullName=report,Sheet=sheet};
        var books=new LinkFakeBooks{Items=new[]{reportBook}};
        reportBook.Owner=books;
        if(opened) {var target=new LinkFakeBook{FullName=targets[0],Sheet=new LinkFakeSheet(),Owner=books};books.Items=new[]{reportBook,target};}
        var app=new LinkFakeApp{ActiveWorkbook=reportBook,Workbooks=books,EnableEvents=true,ScreenUpdating=true};
        var session=new LinkAssistantSession(app){ReportAnchor="probe",ReportName="外部链接"};
        sheet.Values=new object[targets.Length+1,4];
        for(int c=0;c<4;c++)sheet.Values[0,c]=LinkAssistantSession.Headers[c];
        for(int r=0;r<targets.Length;r++){sheet.Values[r+1,0]=targets[r];sheet.Values[r+1,1]=oldPath;sheet.Values[r+1,2]=newPath;session.Links.Add(new ExternalLinkInfo{OwnerPath=targets[r],SourcePath=oldPath});}
        session.StageValidation=(file,expected)=>Check(new HashSet<string>(expected,StringComparer.OrdinalIgnoreCase).SetEquals(XlsxLinkPackage.Scan(file).Select(l=>l.SourcePath)),"offline stage agrees with expected relationships");
        return session;
    }
    static void JournalChecks(string folder,string a,string b,string c) {
        var journal=new LinkOperationLog();var lines=new StringBuilder();journal.LineAdded+=line=>lines.Append(line);
        journal.Record("开始","中文路径与多次操作");journal.EnablePersistence(Path.Combine(folder,"journals"));journal.Record("结束","第二次操作");
        Check(WpsShared.SourcePaths.ReadAllText(journal.FilePath,Encoding.UTF8)==journal.DiagnosticText,"UTF-8 diagnostic journal contains complete trace");
        byte[] bytes=File.ReadAllBytes(journal.FilePath);Check(!(bytes.Length>2&&bytes[0]==239&&bytes[1]==187&&bytes[2]==191),"journal writes UTF-8 without BOM");
        Check(lines.ToString()==journal.Text,"event sequence equals journal sequence");
        var other=new LinkOperationLog();other.EnablePersistence(Path.Combine(folder,"journals"));Check(other.FilePath!=journal.FilePath,"one unique file per opening");
        int writes=0;var broken=new LinkOperationLog((path,text)=>{if(++writes>1)throw new IOException("fake disk full");});var seen=new StringBuilder();broken.LineAdded+=line=>seen.Append(line);
        broken.EnablePersistence(Path.Combine(folder,"journals"));broken.Record("扫描","仍可继续");broken.Record("结束","结果保留");
        Check(broken.PersistenceError!=null&&broken.Text.Contains("日志保存失败")&&broken.Text.Contains("结果保留"),"disk failure leaves all later records copyable");
        Check(seen.ToString()==broken.Text,"failure warning keeps event order identical");
        Check(writes==2,"disk failure is reported once, no retry storm");
        var blocked=new LinkOperationLog();blocked.Record("选择","窗口记录");blocked.EnablePersistence(a);Check(blocked.PersistenceError!=null&&blocked.Text.Contains("窗口记录"),"directory creation failure also preserves initial journal");
        var observer=new LinkOperationLog();observer.LineAdded+=line=>{throw new Exception("fake UI failure");};observer.Record("结果","业务仍继续");Check(observer.Text.Contains("业务仍继续"),"display exception cannot change business result");
        var scope=new LinkScanSelection(a,journal);Check(scope.Candidates().Single()==a,"current scope stays bound to original receiver");
        Check(scope.SetScope(LinkScanScope.Workbooks)&&!scope.SetScope(LinkScanScope.Workbooks),"duplicate radio notifications do not record duplicate transition");
        scope.Choose(new[]{a,b,a},true);Check(scope.Candidates().Length==2,"selected workbooks deduplicate paths");scope.Choose(new string[0],false);Check(scope.Candidates().Length==2,"cancel preserves original selection");
        scope.SetScope(LinkScanScope.Folder);Reject(()=>scope.Candidates(),"switch clears previous scope selection");
        string scanFolder=Path.Combine(folder,"scope");Directory.CreateDirectory(scanFolder);Package(Path.Combine(scanFolder,"first.xlsx"));Package(Path.Combine(scanFolder,"second.xlsx"));Package(Path.Combine(scanFolder,"~$temp.xlsx"));Package(Path.Combine(scanFolder,".EI_LinkStage_temp.xlsx"));Directory.CreateDirectory(Path.Combine(scanFolder,"nested"));Package(Path.Combine(scanFolder,"nested/ignored.xlsx"));
        scope.Choose(new[]{scanFolder},true);Check(scope.Candidates().Length==2,"folder scans only top-level non-temporary workbooks");
        Check(journal.DiagnosticText.Contains("跳过临时文件 2 个")&&journal.DiagnosticText.Contains(Path.Combine(scanFolder,"first.xlsx"))&&!journal.DiagnosticText.Contains("ignored.xlsx"),"actual paths and temporary skip count recorded");
        string empty=Path.Combine(folder,"empty");Directory.CreateDirectory(empty);scope.Choose(new[]{empty},true);Check(scope.Candidates().Length==0&&journal.DiagnosticText.Contains("扫描未执行"),"zero folder files not reported successful");
        scope.SetScope(LinkScanScope.Current);Check(scope.Candidates().Single()==a&&journal.DiagnosticText.Contains("取消选择"),"scope switching and cancel accumulate history");
        string t1=Path.Combine(folder,"journal-target1.xlsx"),t2=Path.Combine(folder,"journal-target2.xlsx");Package(t1,a);Package(t2,a);
        var session=JournalSession(b,new[]{t1},a,c,false);string result=session.Execute(null);
        Check(result.Contains("成功1条，跳过0条，失败0条")&&XlsxLinkPackage.Scan(t1)[0].SourcePath==c,"closed forward committed");Check(!WpsShared.SourcePaths.GetFiles(folder,"*链接备份*").Any(),"no persistent whole workbook backups");
        Check(WpsShared.SourcePaths.Exists(session.Batch.Journal)&&WpsShared.SourcePaths.Exists(session.Batch.LogPath(folder)),"local link log and mapping");
        foreach(string step in session.Steps(true,()=>false)){}Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&session.Batch.State=="Undone","closed reverse mapping");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);var open=((LinkFakeApp)session.App).Workbooks.Items[1];open.Saved=false;result=session.Execute(null);
        Check(open.SaveCount>=2&&open.CloseCount==0&&XlsxLinkPackage.Scan(t1)[0].SourcePath==c,"open target autosaves never closes");
        foreach(string step in session.Steps(true,()=>false)){}Check(open.CloseCount==0&&XlsxLinkPackage.Scan(t1)[0].SourcePath==a,"open inverse never closes");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);open=((LinkFakeApp)session.App).Workbooks.Items[1];open.SourceStyle="uri";open.RetainHistorical=true;result=session.Execute(null);
        Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==c&&result.Contains("成功1条")&&open.LinkSources(1) is Array,"raw URI name and lower-bound-one stale WPS array accepted");
        foreach(string step in session.Steps(true,()=>false)){}Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&open.CloseCount==0,"inverse ignores cached old names and rereads exact source token");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);open=((LinkFakeApp)session.App).Workbooks.Items[1];open.FailSaveAfterChange=true;result=session.Execute(null);
        Check(result.Contains("失败1条")&&XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&session.Batch.State=="Undone","Save exception reconciles intent then reverses actual change");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);open=((LinkFakeApp)session.App).Workbooks.Items[1];open.NoChange=true;result=session.Execute(null);
        Check(result.Contains("失败1条")&&XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&session.Batch.State=="Undone","no-op host API fails saved-package verification without ambiguous undo");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);open=((LinkFakeApp)session.App).Workbooks.Items[1];open.RefuseChange=true;result=session.Execute(null);Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&open.CloseCount==0&&result.Contains("失败1条"),"unsupported ChangeLink refuses no close fallback");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);Package(t2,a);session=JournalSession(b,new[]{t1,t2},a,c,false);int n=0;session.StageValidation=(f,e)=>{if(++n==2)throw new IOException("injected second validation failure");};result=session.Execute(null);
        Check(result.Contains("失败1条")&&XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&XlsxLinkPackage.Scan(t2)[0].SourcePath==a,"failure reverses prior committed links");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);bool stop=false;foreach(string step in session.Steps(false,()=>stop))if(step.StartsWith("已替换"))stop=true;Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==a&&session.Batch.State=="Undone","running cancellation reverses actual completed mapping");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);var app=(LinkFakeApp)session.App;var sheet=((LinkFakeBook)app.ActiveWorkbook).Sheet;
        Check(session.Scan(new[]{t1})==1&&session.ReportName=="外部链接","fixed report reuse");sheet.Values[1,2]=c;Check(session.Scan(new[]{t1})==1&&app.Workbooks.Items[0].Worksheets.Count==1&&sheet.Values[1,2].ToString()=="","repeat scan overwrites same sheet");sheet.Values[1,2]=c;
        var reopened=new LinkAssistantSession(app);Check(reopened.HasReport&&reopened.Links.Count==1,"reopen loads report and original scope");result=reopened.Execute(null);Check(XlsxLinkPackage.Scan(t1)[0].SourcePath==c,"reopen direct execute rescans and validates");
        foreach(string step in reopened.Steps(true,()=>false)){}reopened.Batch.Confirm(RecycleFiles.Send);
        session=JournalSession(b,new[]{t1},a,c,true);session.Scan(new[]{t1});((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet.Values[1,2]=c;Package(t1,a,c);Reject(()=>session.Execute(null),"new link invalidates persisted report");Check(XlsxLinkPackage.Scan(t1).Count==2,"mismatch leaves original unchanged");session.Batch.Confirm(RecycleFiles.Send);
        var changes=new[]{new LinkChange{Old=a,New=b},new LinkChange{Old=b,New=c}};
        Check(LinkBatch.Order(changes,new[]{a,b},false)[0].Old==b,"acyclic chain runs tail first");Check(LinkBatch.Order(changes,new[]{b,c},true)[0].Old==a,"inverse chain runs head first");
        Reject(()=>LinkBatch.Order(new[]{new LinkChange{Old=a,New=c},new LinkChange{Old=b,New=c}},new[]{a,b},false),"many to one refused");Reject(()=>LinkBatch.Order(new[]{new LinkChange{Old=a,New=b}},new[]{a,b},false),"unchanged target source refused");Reject(()=>LinkBatch.Order(new[]{new LinkChange{Old=a,New=b},new LinkChange{Old=b,New=a}},new[]{a,b},false),"cycle refused");
        Package(t1,a,b);session=JournalSession(b,new[]{t1},a,c,true);session.Links.Add(new ExternalLinkInfo {OwnerPath=t1,SourcePath=b});
        var chainSheet=((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet;
        chainSheet.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{t1,a,b,""},{t1,b,c,""}};
        result=session.Execute(null);Check(new HashSet<string>(XlsxLinkPackage.Scan(t1).Select(l=>l.SourcePath),StringComparer.OrdinalIgnoreCase).SetEquals(new[]{b,c}),"live chain executes without source merging");
        using(var stream=File.Open(t1,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update)){zip.GetEntry("customXml/item1.xml").Delete();Part(zip,"customXml/item1.xml","<content>later user edit</content>");}
        foreach(string step in session.Steps(true,()=>false)){}Check(new HashSet<string>(XlsxLinkPackage.Scan(t1).Select(l=>l.SourcePath),StringComparer.OrdinalIgnoreCase).SetEquals(new[]{a,b}),"live reverse chain preserves original source set");
        using(var stream=File.OpenRead(t1))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))using(var reader=new StreamReader(zip.GetEntry("customXml/item1.xml").Open()))Check(reader.ReadToEnd().Contains("later user edit"),"link inverse preserves unrelated post-execution edits");session.Batch.Confirm(RecycleFiles.Send);
        Package(t1,a);session=JournalSession(b,new[]{t1},a,c,true);session.Execute(null);
        using(var form=new LinkAssistantForm(new LinkAssistantSession(session.App)))Check(form.Stop.Enabled&&form.Confirm.Enabled&&!form.ExecuteButton.Enabled,"reopened link form enables undo and confirmation");
        foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(RecycleFiles.Send);
        Package(b,a);session=JournalSession(b,new[]{b},a,c,false);var receiver=(LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook;receiver.Saved=false;result=session.Execute(null);Check(receiver.CloseCount==0&&XlsxLinkPackage.Scan(b)[0].SourcePath==c&&receiver.Sheet.Values[1,3].ToString().Contains("已替换"),"receiver also target saves and writes report without close");foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(RecycleFiles.Send);Package(b);
    }
    static void ResultChecks(string folder,string a,string b,string c)
    {
        var split=new LinkOperationLog();split.Record("","用户结果");split.Trace("调用ChangeLink","Name=内部参数");
        Check(!split.Text.Contains("内部参数")&&!System.Text.RegularExpressions.Regex.IsMatch(split.Text,@"\d{4}-\d{2}-\d{2}"),"display excludes diagnostics and timestamps");
        Check(split.DiagnosticText.Contains("内部参数")&&System.Text.RegularExpressions.Regex.IsMatch(split.DiagnosticText,@"\d{4}-\d{2}-\d{2}"),"diagnostics retain timestamps and parameters");
        int resets=0;split.DisplayReset+=()=>resets++;split.ClearDisplay();Check(split.Text==""&&resets==1&&split.DiagnosticText.Contains("用户结果"),"scan reset only clears displayed batch");
        string t=Path.Combine(folder,"精简日志 很长的中文 工作簿.xlsx");Package(t,a,b);
        var session=JournalSession(c,new[]{t},a,b,true);var app=(LinkFakeApp)session.App;var report=app.Workbooks.Items[0].Sheet;var target=app.Workbooks.Items[1];
        using(var form=new LinkAssistantForm(session))
        {
            Check(form.LogView.Text.Replace("\r","")==form.Guidance.Replace("\r",""),"fresh window shows only brief guidance");
            session.Scan(new[]{t});string once=session.Operations.Text;
            Check(once.Contains("发现2条链接：")&&once.Contains(a)&&once.Contains(b)&&!form.LogView.Text.Contains(form.Guidance),"scan full paths replace initial guidance");
            session.Scan(new[]{t});Check(session.Operations.Text==once,"repeat scan replaces old user batch");
            string disk=WpsShared.SourcePaths.ReadAllText(session.Batch.LogPath(folder));Check(disk.Split(new[]{"开始扫描："},StringSplitOptions.None).Length==3,"diagnostic increment contains two scans without history duplication");
            report.Values[1,2]=b;report.Values[2,2]="";session.Execute(null);
            Check(session.Operations.Text.Contains("尚未修改链接")&&!session.Operations.Text.Contains("成功："),"mapping conflict refuses with no false success");
            session.Batch.Confirm(RecycleFiles.Send);session.Batch=null;
            session.Scan(new[]{t});report.Values[1,2]=a;report.Values[2,2]="";session.Execute(null);
            Check(session.LastResult.Contains("成功0条，跳过2条")&&!session.Operations.Text.Contains("跳过链接：")&&!session.Operations.Text.Contains("待替换链接："),"blank and same mapping skip only in summary");
            Check(!WpsShared.SourcePaths.ReadAllText(session.Batch.LogPath(folder)).Contains("预检未通过"),"confirmed prior diagnostics are not copied into the new batch");
            string copied=null;form.CopyText=text=>copied=text;
            typeof(System.Windows.Forms.Control).GetMethod("OnClick",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic).Invoke(form.CopyLog,new object[]{EventArgs.Empty});
            Check(copied==session.Operations.Text.Replace("\r","")&&copied==form.LogView.Text,"copy captures current concise batch without touching clipboard");
            session.Batch.Confirm(RecycleFiles.Send);session.Batch=null;
        }
        Package(t,a,b);session=JournalSession(c,new[]{t},a,c,true);report=((LinkFakeApp)session.App).Workbooks.Items[0].Sheet;target=((LinkFakeApp)session.App).Workbooks.Items[1];
        session.Links.Add(new ExternalLinkInfo{OwnerPath=t,SourcePath=b});report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{t,a,c,""},{t,b,Path.Combine(folder,"D.xlsx"),""}};Package(Path.Combine(folder,"D.xlsx"));
        target.FailOnChange=2;string result=session.Execute(null);
        Check(result.Contains("失败1条")&&result.Contains("仍有0条")&&new HashSet<string>(XlsxLinkPackage.Scan(t).Select(l=>l.SourcePath)).SetEquals(new[]{a,b}),"partial success summary reflects automatic full rollback: "+result);
        Check(session.Operations.Text.Contains("成功："+a+"\r\n→ "+c)&&session.Operations.Text.Contains("已还原："+c+"\r\n→ "+a),"partial forward and reverse actual per-link results remain visible");
        Check(!session.Operations.Text.Contains("HRESULT")&&!session.Operations.Text.Contains("ChangeLink")&&session.Operations.DiagnosticText.Contains("调用ChangeLink")&&session.Operations.DiagnosticText.Contains("HRESULT"),"call details and exceptions go only to diagnostic channel");
        var reopened=new LinkAssistantSession(session.App);Check(reopened.Operations.Text==session.Operations.Text,"reopening restores current batch transcript");session.Batch.Confirm(RecycleFiles.Send);
        Package(t,a,b);session=JournalSession(c,new[]{t},a,c,false);report=((LinkFakeApp)session.App).Workbooks.Items[0].Sheet;
        session.Links.Add(new ExternalLinkInfo{OwnerPath=t,SourcePath=b});report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{t,b,Path.Combine(folder,"D.xlsx"),""},{t,a,c,""}};
        session.DiagnosticWriter=(path,text)=>{throw new IOException("injected diagnostic failure");};session.Execute(null);
        Check(session.LastResult.Contains("成功2条")&&session.Operations.Text.Contains("成功："+a)&&session.Operations.Text.Contains("成功："+b),"closed sorted two-map result counts links despite diagnostic failure");
        Check(session.Operations.Text.Split(new[]{"日志保存失败："},StringSplitOptions.None).Length==2,"diagnostic failure warns once per batch");
        foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(RecycleFiles.Send);
        Package(t,a);session=JournalSession(c,new[]{t},a,b,true);report=((LinkFakeApp)session.App).Workbooks.Items[0].Sheet;target=((LinkFakeApp)session.App).Workbooks.Items[1];
        report.FailWrite=true;session.Execute(null);Check(session.LastResult.Contains("成功1条")&&session.Operations.Text.Contains("清单状态未完整保存")&&XlsxLinkPackage.Scan(t)[0].SourcePath==b,"report-write failure does not hide successful replacement");report.FailWrite=false;
        target.RefuseChange=true;foreach(string step in session.Steps(true,()=>false)){}
        Check(session.Batch.State=="UndoFailed"&&session.LastResult.Contains("仍有1条")&&session.Operations.Text.Contains("还原失败："),"undo failure exposes remaining link and prohibits confirmation");Reject(()=>session.Batch.Confirm(_=>{}),"failed undo cannot be confirmed");
        target.RefuseChange=false;foreach(string step in session.Steps(true,()=>false)){}
        Reject(()=>{foreach(string step in session.ConfirmSteps(_=>{throw new IOException("injected recycle failure");})){}},"confirmation failure retains remaining records");
        string before=WpsShared.SourcePaths.ReadAllText(session.Batch.LogPath(folder));session.ReportStopped(new IOException("injected cleanup failure"));Check(WpsShared.SourcePaths.ReadAllText(session.Batch.LogPath(folder))==before&&session.Operations.Text.Contains("重试确认清理"),"cleanup failure does not recreate or append recycled log");
        string journal=session.Batch.Journal,log=session.Batch.LogPath(folder),display=session.Operations.Text;
        foreach(string step in session.ConfirmSteps(RecycleFiles.Send)){}
        Check(session.Batch==null&&!WpsShared.SourcePaths.Exists(journal)&&!WpsShared.SourcePaths.Exists(log)&&session.Operations.Text==display+"已确认，本批日志和恢复记录已清理。\r\n","successful retry appends only confirmation and does not recreate logs");
        Check(XlsxLinkPackage.Scan(t)[0].SourcePath==a,"confirmation preserves already restored file state");
        Package(t);session=JournalSession(c,new[]{t},a,b,true);session.Scan(new[]{t});Check(session.Operations.Text.Contains("未发现可替换链接")&&!session.Operations.Text.Contains("再点击「执行替换」"),"zero-link scan omits replacement suggestion");session.Batch.Confirm(RecycleFiles.Send);
    }
    static void MissingDestinations(string folder) {
        string[] originals = Enumerable.Range(0, 14).Select(i => Path.Combine(folder, "source-" + i + ".xlsx")).ToArray();
        string present = Path.Combine(folder, "available.xlsx");
        string[] missing = Enumerable.Range(0, 12).Select(i => Path.Combine(folder, i % 2 == 0 ? "future" : "", "new-" + i + ".xlsx")).ToArray();
        foreach (bool live in new[] { false, true })
        {
            string owner = Path.Combine(folder, "target-" + live + ".xlsx"), report = Path.Combine(folder, "report-" + live + ".xlsx");
            Package(owner, originals); Package(report); Package(present);
            string before = XlsxLinkPackage.Hash(owner);
            var session = JournalSession(report, new[] { owner }, originals[0], present, live);
            var app = (LinkFakeApp)session.App; app.Workbooks.SupportReopen = true;
            session.Scan(new[] { owner });
            var sheet = ((LinkFakeBook)app.ActiveWorkbook).Sheet;
            for (int i = 0; i < 12; i++) sheet.Values[i + 1, 2] = Path.GetDirectoryName(missing[i]) + @"\\[" + Path.GetFileName(missing[i]) + "]";
            sheet.Values[13, 2] = present;
            var original = live ? app.Workbooks.Items[1] : null;
            foreach (string step in session.Steps(false, () => false))
                if (live && step.StartsWith("备份及链接预检完成")) File.Delete(present);
            int missingCount = live ? 13 : 12;
            Check(session.LastResult.Contains("成功13条，跳过1条，失败0条") && session.LastResult.Contains("目标文件当前缺失" + missingCount + "条"),
                "batch applies entered missing destinations and reports their count: " + session.LastResult);
            var saved = XlsxLinkPackage.ScanReportParts(owner);
            Check(new HashSet<string>(saved.Select(l => l.SourcePath ?? XlsxLinkPackage.AbsoluteFileTarget(l.RawTarget))).SetEquals(missing.Concat(new[] { present, originals[13] })),
                "readback proves all normalized targets and the untouched source");
            Check(saved.Count(l => l.IsMissingPath) == missingCount && !Directory.Exists(Path.GetDirectoryName(missing[0]))
                && missing.All(path => !File.Exists(path)), "replacement creates no target files or directories");
            var reopened = new LinkAssistantSession(app);
            Check(reopened.Links.Count(l => l.Status == LinkAssistantSession.MissingStatus) == missingCount, "reopen reads the committed missing status");
            if (original != null) { Check(original.CloseCount == 0, "original workbook stays open"); original.Close(false); }
            foreach (string step in reopened.Steps(true, () => false)) { }
            Check(XlsxLinkPackage.Hash(owner) == before, "checkpoint cancellation restores exact XLSX bytes");
            reopened.Batch.Confirm(File.Delete);
        }
    }
    static void ClosedCheckpointSharingViolation(string folder)
    {
        string source = Path.Combine(folder, "checkpoint-source.xlsx");
        string destination = Path.Combine(folder, "checkpoint-missing.xlsx");
        string outside = Path.Combine(folder, "checkpoint-outside.xlsx");
        Package(outside, Path.Combine(folder, "concurrent-source.xlsx"));
        string outsideHash = XlsxLinkPackage.Hash(outside);
        foreach (bool concurrentEdit in new[] { false, true })
        {
            string owner = Path.Combine(folder, "checkpoint-owner-" + concurrentEdit + ".xlsx");
            string report = Path.Combine(folder, "checkpoint-report-" + concurrentEdit + ".xlsx");
            Package(owner, source); Package(report);
            string before = XlsxLinkPackage.Hash(owner);
            var session = JournalSession(report, new[] { owner }, source, destination, false);
            var app = (LinkFakeApp)session.App;
            app.Workbooks.SupportReopen = true;
            session.Scan(new[] { owner });
            ((LinkFakeBook)app.ActiveWorkbook).Sheet.Values[1, 2] = destination;

            FileStream held = null;
            int commitRetries = 0, restoreRetries = 0;
            session.Operations.DiagnosticLineAdded += line =>
            {
                if (line.Contains("暂存提交暂时占用"))
                {
                    commitRetries++;
                    held.Dispose(); held = null;
                    if (concurrentEdit) File.Copy(outside, owner, true);
                }
                else if (line.Contains("整簿恢复提交暂时占用"))
                {
                    restoreRetries++;
                    held.Dispose(); held = null;
                }
            };
            try
            {
                foreach (string step in session.Steps(false, () => false))
                    if (step.StartsWith("已处理暂存副本"))
                        held = File.Open(owner, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                Check(commitRetries == 1, "checkpoint commit retries a real sharing violation once");
                if (concurrentEdit)
                {
                    var checkpoint = session.Batch.Checkpoints.Single();
                    Check(session.LastResult.Contains("失败1条") && XlsxLinkPackage.Hash(owner) == outsideHash,
                        "commit and automatic recovery both protect a concurrent edit made before retry");
                    Check(checkpoint.NeedsRestore && XlsxLinkPackage.Hash(checkpoint.Backup) == before,
                        "concurrent edit leaves the original backup and pending verification");
                    continue;
                }
                Check(session.LastResult.Contains("成功1条") && XlsxLinkPackage.Hash(owner) != before,
                    "released share lock allows the validated missing-target replacement");

                held = File.Open(owner, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                foreach (string step in session.Steps(true, () => false)) { }
                Check(restoreRetries == 1 && XlsxLinkPackage.Hash(owner) == before,
                    "checkpoint undo retries a share lock and restores the exact bytes");
                session.Batch.Confirm(File.Delete);
            }
            finally { if (held != null) held.Dispose(); }
        }
    }
    static LinkFakeSheet OutputFixture(bool nullHidden)
    {
        var sheet=new LinkFakeSheet{HiddenHeightNull=nullHidden,Values=new object[4,6]};
        for(int c=1;c<=6;c++){sheet.Widths[c]=70+c;for(int r=1;r<=4;r++){sheet.Values[r-1,c-1]="v"+r+"-"+c;var style=sheet.Cell(r,c);style.Wrap=(r+c)%2==0;style.Bold=r==1&&c%2==0;style.FontColor=10*c;style.FillColor=20*c;style.Horizontal=10+c;style.Vertical=20+c;}}
        for(int r=1;r<=4;r++)sheet.Row(r).Height=20+r;
        sheet.Row(2).Hidden=true;sheet.Row(4).Hidden=true;sheet.Values[3,2]=new string('L',95);return sheet;
    }
    static void OutputFormatChecks()
    {
        foreach(bool nullHidden in new[]{false,true})
        {
            var sheet=OutputFixture(nullHidden);string outside=sheet.Cell(1,1).FontColor+":"+sheet.Width(1)+":"+sheet.Cell(3,6).Wrap;
            WorkbookUi.FitOutput(sheet,3,2,4);
            Check(sheet.Fits.SequenceEqual(new[]{"C1:D4"}),"full non-A output range supplies the column measurement");
            Check(sheet.Width(3)>=95&&sheet.Width(3)<255,"last hidden row participates in width measurement");
            Check(sheet.Cell(1,3).Bold&&sheet.Cell(1,3).Horizontal==-4108&&sheet.Cell(1,3).Vertical==-4108&&!sheet.Cell(1,3).Wrap,"header becomes centered bold and single-line");
            Check(!sheet.Cell(4,3).Wrap&&sheet.Row(1).Height>=26,"ordinary output keeps one-line body and minimum header height");
            Check(sheet.Row(2).Hidden&&sheet.Row(4).Hidden,"hidden zero/null heights preserve both hidden rows");
            Check(sheet.Row(2).Height==22&&sheet.Row(4).Height==24,"measurement preserves latent body heights");
            Check(outside==sheet.Cell(1,1).FontColor+":"+sheet.Width(1)+":"+sheet.Cell(3,6).Wrap,"non-output columns keep their prior formats and widths");
            Check(sheet.RowFits.Contains("C1:D1"),"header height uses a separate native row fit");

            sheet=OutputFixture(nullHidden);sheet.Values[3,2]=new string('L',600);sheet.Cell(1,3).Size=36;WorkbookUi.FitOutput(sheet,3,2,4);
            Check(sheet.Width(3)==255&&sheet.Cell(2,3).Wrap&&sheet.Cell(4,3).Wrap,"host maximum column width wraps only that data column");
            Check(!sheet.Cell(1,3).Wrap&&!sheet.Cell(2,4).Wrap,"maximum-width fallback never wraps headers or unrelated columns");
            Check(sheet.Row(4).Height>24&&sheet.Row(1).Height>=62,"wrapped body and large header receive sufficient row height");
            Check(sheet.Row(2).Hidden&&sheet.Row(4).Hidden,"row AutoFit restores all hidden flags");

            foreach(bool rowFailure in new[]{false,true})
            {
                sheet=OutputFixture(nullHidden);sheet.Values[3,2]=new string('L',600);string original=sheet.State();
                if(rowFailure)sheet.FailRowAutoFitAfter=2;else sheet.FailColumnAutoFitAfter=1;
                Reject(()=>WorkbookUi.FitOutput(sheet,3,2,4),"partial "+(rowFailure?"row":"column")+" fit failure is propagated");
                Check(sheet.State()==original,"partial fit rolls back widths, mixed wrap, header properties, heights and hidden rows");
            }
            sheet=OutputFixture(nullHidden);string beforeHeaderFit=sheet.State();sheet.FailRowAutoFitAfter=1;
            Reject(()=>WorkbookUi.FitOutput(sheet,3,2,4),"header-only row AutoFit failure is propagated");
            Check(sheet.State()==beforeHeaderFit,"header row fit failure also restores full original formats and hidden flags");
            sheet=OutputFixture(nullHidden);sheet.FailColumnAutoFitAfter=1;sheet.FailRestoreWidthColumn=3;sheet.FailRestoreWidthValue=73;
            Reject(()=>WorkbookUi.FitOutput(sheet,3,2,4),"width restoration failure remains visible");
            Check(sheet.Width(4)==74&&sheet.Cell(1,3).FontColor==30&&sheet.Cell(1,3).Horizontal==13,"one width restore failure does not block other widths or the same column header restore");
            Check(sheet.Row(2).Hidden&&sheet.Row(4).Hidden&&sheet.Row(2).Height==22&&sheet.Row(4).Height==24,"width restoration failure still restores every hidden row and height");

            sheet=OutputFixture(nullHidden);sheet.Row(4).Height=0;string badState=sheet.State();
            Reject(()=>WorkbookUi.FitOutput(sheet,3,2,4),"unreadable latent hidden height refuses formatting before AutoFit");
            Check(sheet.State()==badState&&sheet.Fits.Count==0,"snapshot failure restores temporary visibility and changes no formats");
        }
        var empty=new LinkFakeSheet{Values=new object[,]{{"完整表头名称"}}};WorkbookUi.FitOutput(empty,1,1,1);
        Check(empty.Fits.SequenceEqual(new[]{"A1:A1"})&&empty.Width(1)>10&&empty.Row(1).Height>=26,"empty result still fits the full header-only rectangle");
    }
    static void RenameOutputChecks(string folder)
    {
        string bookPath=Path.Combine(folder,"rename-report.xlsx");Package(bookPath);
        var book=new LinkFakeBook{FullName=bookPath,Sheet=new LinkFakeSheet{Name="接收表"}};var books=new LinkFakeBooks{Items=new[]{book}};book.Owner=books;
        var app=new LinkFakeApp{ActiveWorkbook=book,Workbooks=books,EnableEvents=true,ScreenUpdating=true};
        var session=new RenameWorkbookSession(app);var file=new FileSnapshot{Path=@"Z:\导出测试\"+new string('x',100)+".txt",Name=new string('x',100)+".txt"};
        session.Create(folder,new List<FileSnapshot>{file});var sheet=book.SheetItems.Single(s=>s.Name=="文件批量命名");
        Check(sheet.Fits.SequenceEqual(new[]{"A1:D2"})&&sheet.Width(1)>100&&sheet.Width(2)>100,"rename creation fits complete path and filename values without a 30-character cap");
        Check(sheet.Cell(1,1).Bold&&sheet.Cell(1,1).Horizontal==-4108&&sheet.Cell(1,1).Vertical==-4108,"rename headers are centered bold");
        var result=new FileOperationResult();string longStatus="处理中："+new string('s',160);result.Status[file.Path]=longStatus;session.WriteResult(result);
        Check(sheet.Values[1,3].ToString()==longStatus&&sheet.Width(4)>160,"rename status updates refit the actual complete result text");
        Check(book.SavedReports.Last().Contains(longStatus)&&book.SavedReports.Last().Contains(":True:False:"),"rename save persists status and final header/body formats");
        Check(book.SavedHashes.Last()==XlsxLinkPackage.Hash(bookPath)&&book.Saved,"rename save fingerprint represents the fitted report");
        string emptyPath=Path.Combine(folder,"rename-empty.xlsx");Package(emptyPath);book=new LinkFakeBook{FullName=emptyPath,Sheet=new LinkFakeSheet{Name="文件批量命名"}};books=new LinkFakeBooks{Items=new[]{book}};book.Owner=books;
        app=new LinkFakeApp{ActiveWorkbook=book,Workbooks=books,EnableEvents=true,ScreenUpdating=true};session=new RenameWorkbookSession(app);session.Create(folder,new List<FileSnapshot>());
        Check(book.Sheet.Fits.SequenceEqual(new[]{"A1:D1"})&&book.Sheet.Values.GetLength(0)==1,"empty rename list keeps a fitted four-column header");
    }
    static void LinkOutputChecks(string folder,string a,string b,string c)
    {
        string target=Path.Combine(folder,"format-link-owner.xlsx");string longSource=a;Package(target,longSource);
        var session=JournalSession(b,new[]{target},longSource,c,true);var app=(LinkFakeApp)session.App;var receiver=app.Workbooks.Items[0];var sheet=receiver.Sheet;
        session.Scan(new[]{target});Check(sheet.Fits.Last()=="A1:D2"&&sheet.Width(2)>=Math.Min(255,longSource.Length)&&sheet.Width(2)>30,"link scan fits the full source path after all report styling");
        Check(sheet.Cell(1,2).Bold&&sheet.Cell(1,2).Horizontal==-4108&&!sheet.Cell(1,2).Wrap,"link scan headers follow the shared output standard");
        sheet.Values[1,2]=c;sheet.Widths[4]=2;session.Execute(null);
        Check(sheet.Values[1,3].ToString().Contains("已替换")&&sheet.Width(4)>2,"link result rewrites and refits the status column");
        Check(receiver.SavedReports.Last().Contains("已替换")&&receiver.SavedReports.Last()==string.Join("\n",receiver.SheetItems.Select(s=>s.State())),"link save contains the final report value and all fitted formats");
        Check(receiver.SavedHashes.Last()==XlsxLinkPackage.Hash(b),"report digest is captured only after fitted content reaches disk");
        foreach(string step in session.Steps(true,()=>false)){}session.Batch.Confirm(RecycleFiles.Send);
    }
    static void DirectoryOutputChecks(string folder)
    {
        string path=Path.Combine(folder,"directory-format.xlsx");Package(path);
        var first=new LinkFakeSheet{Name="有'引号表",Values=new object[,]{{"原表头"},{"原正文"}}};var second=new LinkFakeSheet{Name="较长名称"+new string('x',20),Visible=2,Values=new object[,]{{"隐藏原表头"},{"隐藏原正文"}}};
        var book=new LinkFakeBook{FullName=path,Sheet=first};second.Parent=book;book.SheetItems.Add(second);var books=new LinkFakeBooks{Items=new[]{book}};book.Owner=books;
        var app=new LinkFakeApp{ActiveWorkbook=book,Workbooks=books,EnableEvents=true,ScreenUpdating=true};
        string source1=first.State(),source2=second.State();int count=DirectoryNavigation.Generate(app);var output=book.SheetItems.Single(s=>s.Name=="目录");
        Check(count==2&&output.Values.GetLength(0)==3,"directory writes one row per source worksheet");
        Check(output.HyperlinkTargets.Count==2&&book.NamedRanges.Count==4,"directory uses real hyperlink and stable named-range calls");
        Check(output.Fits.SequenceEqual(new[]{"A1:C3"})&&output.Width(1)>=second.Name.Length,"directory fits all generated rows including the longest final name");
        for(int r=1;r<=3;r++)for(int c=1;c<=3;c++)Check(output.Cell(r,c).Bold&&output.Cell(r,c).Horizontal==-4108&&output.Cell(r,c).Vertical==-4108,"directory whole output retains centered bold typography at "+r+","+c);
        Check(output.FitSnapshots.Single().Formats[1,1]=="True:-4108:-4108:10:@","directory column AutoFit runs after hyperlink-added styles are replaced by final whole-table formatting");
        Check(output.Values[2,2].ToString()=="非常隐藏"&&second.Visible==2&&first.State()==source1&&second.State()==source2,"directory formats only its new report and preserves source values, formats and visibility");
    }
    static void AdvancedFilterOutputChecks(string folder)
    {
        string path=Path.Combine(folder,"filter-format.xlsx");Package(path);string longText="完整末行内容"+new string('v',110);
        var source=new LinkFakeSheet{Name="原数据",Values=new object[5,4],RegionAddress="B2:D5"};
        source.Values[1,1]="编号";source.Values[1,2]="金额";source.Values[1,3]="说明";
        source.Values[2,1]="保留";source.Values[2,2]=1876.5;source.Values[2,3]="短";
        source.Values[3,1]="跳过";source.Values[3,2]=8.0;source.Values[3,3]="其它";
        source.Values[4,1]="保留";source.Values[4,2]=2678.5;source.Values[4,3]=longText;
        for(int r=3;r<=5;r++){source.Cell(r,2).NumberFormat="@";source.Cell(r,3).NumberFormat="#,##0.00";source.Cell(r,4).NumberFormat="@";}
        // The native engine is replaced only at its call boundary. This fixture
        // validates output formatting, not Excel/WPS criteria or filter accuracy.
        source.NativeFilterResult=new object[,]{{"编号","金额","说明"},{"保留",1876.5,"短"},{"保留",2678.5,longText}};
        var book=new LinkFakeBook{FullName=path,Sheet=source};var books=new LinkFakeBooks{Items=new[]{book}};book.Owner=books;
        var selection=source.Range["B3"];var app=new LinkFakeApp{ActiveWorkbook=book,ActiveSheet=source,Selection=selection,ActiveWindow=new LinkFakeWindow(),Workbooks=books,EnableEvents=true,ScreenUpdating=true};
        string before=source.State();int count=QuickAdvancedFilter.Run(app);var output=book.SheetItems.Single(s=>s.Name=="高筛结果");
        Check(count==2&&output.Values[2,2].ToString()==longText,"high-filter formats the supplied native result without truncating the last row");
        Check(source.NativeFilterCalls==1&&source.LastCriterion[0,0].ToString()=="编号"&&source.LastCriterion[1,0].ToString()=="保留","high-filter invokes the native boundary with the actual selected criterion");
        Check(output.Fits.SequenceEqual(new[]{"A1:C3"})&&output.Width(3)>110,"high-filter column fitting includes the entire returned native rectangle");
        Check(output.Cell(2,2).NumberFormat=="#,##0.00"&&output.Cell(3,3).NumberFormat=="@"&&output.Cell(2,2).Size==10,"high-filter restores source column formats and final font before fitting");
        Check(output.FitSnapshots.Single().Formats[1,1]=="False:19:21:10:#,##0.00","high-filter AutoFit measures the final numeric format and body font, not the native copy defaults");
        for(int c=1;c<=3;c++)Check(output.Cell(1,c).Bold&&output.Cell(1,c).Horizontal==-4108&&output.Cell(1,c).Vertical==-4108&&!output.Cell(1,c).Wrap,"high-filter header is centered bold and single-line at "+c);
        Check(source.State()==before&&book.SheetItems.Count==2&&!book.SheetItems.Any(s=>s.Name.StartsWith("_EI高筛条件")),"high-filter output formatting preserves source cells and removes the criteria-only sheet");
    }
    [STAThread] static int Main(){
        string folder=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"workbook-tools-probe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(folder);
        try{
            string literal=Path.Combine(folder,"中文 %20#来源.xlsx"),owner=Path.Combine(folder,"owner.xlsx");
            var known=new[]{new ExternalLinkInfo{OwnerPath=owner,SourcePath=literal,RawTarget=new Uri(literal).AbsoluteUri}};
            foreach(string token in new[]{literal,new Uri(literal).AbsoluteUri,Path.GetFileName(literal),Path.GetPathRoot(literal)=="\\"?literal:literal.Substring(2)})
            {Array raw=Array.CreateInstance(typeof(object),new[]{1},new[]{1});raw.SetValue(token,1);var parsed=LinkAssistantSession.ParseHostSources(raw,owner,known);Check(parsed.Single().Name==token&&parsed.Single().Path==literal,"host source preserves exact token and percent filename: "+token);}
            Check(LinkAssistantSession.ParseHostSources(null,owner,known).Count==0,"empty host source result");
            Reject(()=>LinkAssistantSession.ParseHostSources(new object[1,1],owner,known),"unexpected multidimensional source result refused");
            var ambiguous=new[]{new ExternalLinkInfo{SourcePath=Path.Combine(folder,"A/same.xlsx")},new ExternalLinkInfo{SourcePath=Path.Combine(folder,"B/same.xlsx")}};
            Reject(()=>LinkAssistantSession.ParseHostSources(new[]{"same.xlsx"},owner,ambiguous),"bare duplicate filename association refused");
            string[] heads={"凭证号","月份","数据"};
            var cells=new List<FilterCriterion>{new FilterCriterion{Row=2,Column=0,Value="001"},new FilterCriterion{Row=2,Column=1,Value=1},new FilterCriterion{Row=12,Column=0,Value="003"},new FilterCriterion{Row=12,Column=1,Value=2}};
            object[,] criteria=QuickFilterCriteria.Build(heads,cells);
            Check(criteria.GetLength(0)==3&&criteria.GetLength(1)==2,"two physical OR rows with inferred headers");
            Check((string)criteria[1,0]=="001"&&(int)criteria[1,1]==1&&(string)criteria[2,0]=="003"&&(int)criteria[2,1]==2,"no Cartesian criteria");
            cells.Add(cells[0]);Check(QuickFilterCriteria.Build(heads,cells).GetLength(0)==3,"overlap deduplicated");
            cells.Add(new FilterCriterion{Row=24,Column=2,Value=""});criteria=QuickFilterCriteria.Build(heads,cells);Check(criteria.GetLength(0)==3,"blank OR row omitted");
            Check(criteria[1,2]==null,"missing column unconstrained");
            cells[0].Value=">100";cells[2].Value="*数据*";criteria=QuickFilterCriteria.Build(heads,cells);Check((string)criteria[1,0]==">100"&&(string)criteria[2,0]=="*数据*","native comparison and wildcard preserved");
            Reject(()=>QuickFilterCriteria.Build(heads,new List<FilterCriterion>{new FilterCriterion{Row=2,Column=0,Value=null}}),"all-blank criteria refused");
            Reject(()=>QuickFilterCriteria.Build(heads,new List<FilterCriterion>{new FilterCriterion{Row=2,Column=7,Value=1}}),"out-of-region column refused");
            Reject(()=>QuickFilterCriteria.Build(heads,new List<FilterCriterion>{new FilterCriterion{Row=2,Column=0,Value=unchecked((int)0x800A07D0)}}),"boxed WPS error criterion refused");
            Check(WorkbookUi.UniqueName("高筛结果",new[]{"高筛结果","高筛结果(2)"})=="高筛结果(3)","existing output sheets preserved");
            Check(WorkbookUi.Quote("隐藏'表")=="'隐藏''表'","apostrophe quoted for hyperlink/names");
            string a=Path.Combine(folder,"A.xlsx"),b=Path.Combine(folder,"B.xlsx"),c=Path.Combine(folder,"C.xlsx");
            Package(a);Package(b);Package(c);
            string original=Path.Combine(folder,"中文 # 100% 测试.xlsx");Package(original,a,b,a);
            string originalHash=XlsxLinkPackage.Hash(original);
            Check(XlsxLinkPackage.Scan(original).Count==3,"all external relationship records read");
            Check(XlsxLinkPackage.Resolve(@"D:\Folder\book.xlsx",@"\Data\源%20文件.xlsx")==@"D:\Data\源 文件.xlsx","drive rooted relationship uses owner drive");
            Check(XlsxLinkPackage.Resolve(original,"https://example.com/book.xlsx")==null,"web link separate");
            Check(XlsxLinkPackage.Resolve(original,"A.xlsx")==a,"relative source path");
            string stage=XlsxLinkPackage.Prepare(original,new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){{a,b},{b,c}});
            var links=XlsxLinkPackage.Scan(stage);Check(links[0].SourcePath==b&&links[1].SourcePath==c&&links[2].SourcePath==b,"A to B / B to C simultaneous and duplicate relationships");
            Check(XlsxLinkPackage.Hash(original)==originalHash,"prepare never changes original");
            XlsxLinkPackage.VerifyRewrite(original,stage,new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){{a,b},{b,c}});checks++;
            string backup=XlsxLinkPackage.Commit(original,stage,originalHash);
            Check(WpsShared.SourcePaths.Exists(backup)&&XlsxLinkPackage.Hash(backup)==originalHash,"backup exact bytes before commit");
            Check(XlsxLinkPackage.Scan(original)[0].SourcePath==b,"commit writes only staged result");
            stage=XlsxLinkPackage.Prepare(original,new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){{b,c}});
            var same=XlsxLinkPackage.Scan(stage);Check(same.All(x=>x.SourcePath==c),"many-to-one mapping supported");
            File.Delete(stage);
            Reject(()=>XlsxLinkPackage.Prepare(original,new Dictionary<string,string>{{b,original}}),"self link refused");
            Reject(()=>XlsxLinkPackage.Prepare(original,new Dictionary<string,string>{{b,Path.Combine(folder,"missing.xlsx")}}),"missing new source refused");
            Reject(()=>XlsxLinkPackage.Prepare(original,new Dictionary<string,string>{{a,c}}),"stale mapping refused");
            Check(!WpsShared.SourcePaths.GetFiles(folder,".EI_LinkStage_*.xlsx").Any(),"failed preparation cleans stage");
            string unicode=Path.Combine(folder,"中文空格 #100%.xlsx");Package(unicode);
            stage=XlsxLinkPackage.Prepare(original,new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){{b,unicode}});
            Check(XlsxLinkPackage.Scan(stage).Count(l=>l.SourcePath==unicode)==2,"Unicode/reserved URI round-trip without double decode");
            Reject(()=>XlsxLinkPackage.Commit(original,stage,"stale hash"),"changed original aborts commit");
            Check(WpsShared.SourcePaths.Exists(stage),"failed commit leaves stage recoverable");File.Delete(stage);
            string bad=Path.Combine(folder,"bad.xlsx");File.WriteAllText(bad,"not a zip");
            Reject(()=>XlsxLinkPackage.Scan(bad),"extension alone insufficient");
            string signed=Path.Combine(folder,"signed.xlsx");Package(signed);using(var stream=File.Open(signed,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update))Part(zip,"_xmlsignatures/sig1.xml","<signed/>");
            Reject(()=>XlsxLinkPackage.Scan(signed),"signed document not silently invalidated");
            var host=new QuietFake{EnableEvents=true,ScreenUpdating=false};
            using(new HostQuiet(host)){Check(!host.EnableEvents&&!host.ScreenUpdating,"host paused");}
            Check(host.EnableEvents&&!host.ScreenUpdating,"both original states restored");
            var quiet=new HostQuiet(host);quiet.Dispose();quiet.Dispose();Check(host.EnableEvents,"dispose idempotent");
            host.FailDisable=true;Reject(()=>new HostQuiet(host),"pause failure refuses work");Check(host.EnableEvents&&!host.ScreenUpdating,"constructor failure restores original host state");
            var fakeSheet=new LinkFakeSheet(); var fakeBook=new LinkFakeBook{FullName=original,Sheet=fakeSheet};
            var target=new LinkFakeBook{FullName=a,Saved=false,Sheet=new LinkFakeSheet()};
            var fakeApp=new LinkFakeApp{ActiveWorkbook=fakeBook,Workbooks=new LinkFakeBooks{Items=new[]{fakeBook,target}}};
            var session=new LinkAssistantSession(fakeApp);session.ReportAnchor="probe";
            session.Links=new List<ExternalLinkInfo>{new ExternalLinkInfo{OwnerPath=a,SourcePath=b}};
            object[,] report={{"文档路径","外部链接","新链接（待填）","处理状态"},{a,b,c,"old status"}};
            fakeSheet.Values=report;
            Check(session.ReadMappings().Count==1,"saved report parses expected mapping identity");
            fakeBook.Saved=false;Check(session.ReadMappings().Count==1&&fakeBook.SaveCount==1,"unsaved report automatically saved");
            fakeBook.ReadOnly=true;Reject(()=>session.ReadMappings(),"read-only report refused");fakeBook.ReadOnly=false;
            report[1,0]=b;Reject(()=>session.ReadMappings(),"edited source owner refused");report[1,0]=a;
            report[1,2]=Path.Combine(folder,"missing.xlsx");Check(session.ReadMappings().Count==1&&session.Links[0].DestinationMissing,"missing destination remains an actionable mapping");report[1,2]=c;
            string directory = Path.Combine(folder, "folder.xlsx"); Directory.CreateDirectory(directory);
            report[1,2]=directory;Reject(()=>session.ReadMappings(),"existing directory is refused");report[1,2]=c;
            report[1,2]="relative.xlsx";Reject(()=>session.ReadMappings(),"relative user new path refused");report[1,2]=c;
            string targetHash=XlsxLinkPackage.Hash(a);target.RefuseChange=true;
            Check(session.Execute(_=>{}).Contains("失败1条"),"unsupported open target reports failure");Check(target.SaveCount>0&&target.CloseCount==0&&XlsxLinkPackage.Hash(a)==targetHash,"dirty target saved without close or overwrite");session.Batch.Confirm(RecycleFiles.Send);session.Batch=null;target.RefuseChange=false;
            report[1,2]="";Check(session.ReadMappings().Count==0,"empty new path skipped");
            report[1,2]=b;Check(session.ReadMappings().Count==0,"identical path skipped");
            report[1,2]=a;Reject(()=>session.ReadMappings(),"report self-reference refused");
            session.Links.Add(new ExternalLinkInfo{OwnerPath=a,SourcePath=c});
            fakeSheet.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{a,c,b,""},{a,b,c,""}};
            Check(session.ReadMappings().Count==2,"sorted report rows map by stable key");
            fakeSheet.Values[2,1]=c;Reject(()=>session.ReadMappings(),"duplicate report identity refused");
            OutputFormatChecks();
            RenameOutputChecks(folder);
            LinkOutputChecks(folder,a,b,c);
            DirectoryOutputChecks(folder);
            AdvancedFilterOutputChecks(folder);
            JournalChecks(folder,a,b,c);
            ResultChecks(folder,a,b,c);
            MissingDestinations(folder);
            ClosedCheckpointSharingViolation(folder);
            BreakLinkChecks(folder,a,b,c);
            NameLinkChecks(folder,a,b,c);
            Console.WriteLine("WorkbookToolsProbe "+checks+" PASS, 0 FAIL");return 0;
        }catch(Exception ex){File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"workbook-tools-probe-failure.log"),ex.ToString(),new UTF8Encoding(false));Console.WriteLine("FAILED: "+ex);return 1;}
        finally {
            string target=Path.GetFullPath(folder),expected=Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory).TrimEnd(Path.DirectorySeparatorChar);
            if(Path.GetDirectoryName(target)!=expected||!Path.GetFileName(target).StartsWith("workbook-tools-probe-"))throw new InvalidOperationException("Unsafe fixture cleanup");
            if(Directory.Exists(target))Directory.Delete(target,true);
        }
    }
}
public class LinkFakeApp {public object ActiveWorkbook,ActiveSheet,Selection,ActiveWindow;public LinkFakeBooks Workbooks;public bool EnableEvents,ScreenUpdating,DisplayAlerts=true;}
public class LinkFakeWindow {public LinkFakeSelectedSheets SelectedSheets{get{return new LinkFakeSelectedSheets();}}}
public class LinkFakeSelectedSheets {public int Count{get{return 1;}}}
public class LinkFakeBooks {public LinkFakeBook[] Items;public bool SupportReopen,ForeignOpen;public LinkFakeSheet ClosedSheet;public int Count{get{return Items.Length;}}public object Item(int index){return Items[index-1];}public object Open(string file,int update,bool readOnly){if(ForeignOpen)return Items[0];if(!SupportReopen)throw new IOException("fake reopen failure");var book=new LinkFakeBook{FullName=file,Sheet=ClosedSheet,Owner=this};Items=Items.Concat(new[]{book}).ToArray();return book;}}
public partial class LinkFakeBook {
    public bool RefuseChange,RetainHistorical,FailSaveAfterChange,NoChange;public string SourceStyle;public int CloseCount,FailOnChange,ChangeCount;private bool failSave;
    private readonly List<string> history=new List<string>();
    public object LinkSources(int type){var paths=XlsxLinkPackage.Scan(FullName).Where(l=>l.SourcePath!=null).Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();if(RetainHistorical){foreach(string p in paths)if(!history.Contains(p))history.Add(p);paths=history.ToList();}var names=paths.Select(p=>SourceStyle=="uri"?new Uri(p).AbsoluteUri:p).Concat(MissingHostLinks().Select(NativeMissingName)).ToList();Array raw=Array.CreateInstance(typeof(object),new[]{names.Count},new[]{1});for(int i=0;i<names.Count;i++)raw.SetValue(names[i],i+1);return raw;}
    public void ChangeLink(string oldPath,string newPath,int type){if(RefuseChange||++ChangeCount==FailOnChange)throw new IOException("宿主不支持本次链接修改");if(SourceStyle=="uri")oldPath=new Uri(oldPath).LocalPath;if(!NoChange&&!ChangeMissingHostLink(oldPath,newPath)){string stage=XlsxLinkPackage.Prepare(FullName,new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase){{oldPath,newPath}});File.Replace(stage,FullName,null);}Saved=false;failSave=FailSaveAfterChange;FailSaveAfterChange=false;}
    public string FullName;public bool Saved=true,ReadOnly,ProtectStructure;public int SaveCount;private LinkFakeSheet sheet;
    public readonly List<LinkFakeSheet> SheetItems=new List<LinkFakeSheet>();public readonly Dictionary<string,LinkFakeName> NamedRanges=new Dictionary<string,LinkFakeName>();
    public readonly List<string> SavedReports=new List<string>();public readonly List<string> SavedHashes=new List<string>();
    public LinkFakeSheet Sheet{get{return sheet;}set{sheet=value??new LinkFakeSheet();sheet.Parent=this;if(!SheetItems.Contains(sheet))SheetItems.Add(sheet);}}
    public LinkFakeBooks Owner;public LinkFakeSheets Worksheets{get{return new LinkFakeSheets{Book=this};}}public LinkFakeSheets Sheets{get{return Worksheets;}}public LinkFakeSheet ActiveSheet{get{return Sheet;}}
    public void Activate(){}public void Close(bool save){CloseCount++;pendingBreak=null;Owner.ClosedSheet=Sheet;Owner.Items=Owner.Items.Where(b=>b!=this).ToArray();}
    public void Save(){SaveCount++;if(failSave){failSave=false;throw new IOException("injected save failure");}if(pendingBreak!=null){File.WriteAllBytes(FullName,pendingBreak);pendingBreak=null;}
        if(SheetItems.Any(s=>s.Dirty))using(var stream=File.Open(FullName,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update))
        {var old=zip.GetEntry("customXml/probe-report.xml");if(old!=null)old.Delete();string state=string.Join("\n",SheetItems.Select(s=>s.State()));using(var writer=new StreamWriter(zip.CreateEntry("customXml/probe-report.xml").Open(),new UTF8Encoding(false)))writer.Write(state);SavedReports.Add(state);foreach(var s in SheetItems)s.Dirty=false;}
        SavedHashes.Add(XlsxLinkPackage.Hash(FullName));Saved=true;}
    public LinkFakeNames Names{get{return new LinkFakeNames{Book=this};}}
}
public class LinkFakeSheets{public LinkFakeBook Book;public int Count{get{return Book.SheetItems.Count;}}public object Item(object name){if(name is int)return Book.SheetItems[(int)name-1];return Book.SheetItems.Single(s=>s.Name==Convert.ToString(name));}public object Add(object before,object after,int count,object type){var sheet=new LinkFakeSheet{Parent=Book,Name="Sheet"+(Count+1)};int at=before is LinkFakeSheet?Book.SheetItems.IndexOf((LinkFakeSheet)before):Book.SheetItems.Count;Book.SheetItems.Insert(at,sheet);return sheet;}}
public class HighlightFakeRange{public HighlightFakeSheet Worksheet;public HighlightFakeFill Interior;}
public class HighlightFakeSheet{public LinkFakeBook Parent;}
public class HighlightFakeFill{public LinkFakeBook Book;private int color=65535;public int Color{get{return color;}set{color=value;Book.Saved=false;}}}
public class LinkFakeNames {public LinkFakeBook Book;public int Count{get{return Book.NamedRanges.Count;}}public object Add(string name,string refers,bool visible){int bang=refers.LastIndexOf('!');string title=refers.Substring(1,bang-1).Trim('\'').Replace("''","'");var range=new LinkFakeName{Book=Book,Sheet=Book.SheetItems.Single(s=>s.Name==title),Name=name,Address=refers.Substring(bang+1)};Book.NamedRanges[name]=range;return range;}public LinkFakeName Item(object name){if(name is int)return Book.NamedRanges.Values.ElementAt((int)name-1);LinkFakeName range;return Book.NamedRanges.TryGetValue(Convert.ToString(name),out range)?range:new LinkFakeName{Book=Book,Sheet=Book.Sheet,Name=Convert.ToString(name)};}}
public class LinkFakeName {public string Name,Address;public LinkFakeBook Book;public void Delete(){Book.NamedRanges.Remove(Name);}public LinkFakeSheet Sheet;public LinkFakeRange RefersToRange{get{return new LinkFakeRange{Worksheet=Sheet,Address=Address};}}}
public class LinkFakeCellState {public string Name="Calibri",NumberFormat="General";public double Size=10;public bool Bold,Wrap;public int FontColor=17,FillColor=23;public int Horizontal=19,Vertical=21;}
public class LinkFakeRowState {public double Height=20;public bool Hidden;}
public class LinkFakeSheet {
    public object[,] Values=new object[1,4],NativeFilterResult,LastCriterion;public string Name="外部链接",RegionAddress;public bool FailWrite,Dirty,HiddenHeightNull,ProtectContents;public LinkFakeBook Parent;
    public int NativeFilterCalls;public readonly List<string> HyperlinkTargets=new List<string>();public readonly List<LinkFakeFitSnapshot> FitSnapshots=new List<LinkFakeFitSnapshot>();
    public int FailColumnAutoFitAfter,FailRowAutoFitAfter,FailRestoreWidthColumn;public double FailRestoreWidthValue;
    public readonly Dictionary<string,LinkFakeCellState> CellStates=new Dictionary<string,LinkFakeCellState>();public readonly Dictionary<int,double> Widths=new Dictionary<int,double>();public readonly Dictionary<int,LinkFakeRowState> RowStates=new Dictionary<int,LinkFakeRowState>();
    public readonly List<string> Fits=new List<string>(),RowFits=new List<string>();
    public LinkFakeCellState Cell(int r,int c){string key=r+","+c;LinkFakeCellState value;if(!CellStates.TryGetValue(key,out value))CellStates[key]=value=new LinkFakeCellState();return value;}
    public LinkFakeRowState Row(int r){LinkFakeRowState value;if(!RowStates.TryGetValue(r,out value))RowStates[r]=value=new LinkFakeRowState();return value;}
    public double Width(int c){double value;return Widths.TryGetValue(c,out value)?value:8.43;}
    public void Touch(){Dirty=true;if(Parent!=null)Parent.Saved=false;}
    public void Put(int r,int c,object value){if(FailWrite)throw new IOException("fake report write failure");if(r>Values.GetLength(0)||c>Values.GetLength(1)){var expanded=new object[Math.Max(r,Values.GetLength(0)),Math.Max(c,Values.GetLength(1))];for(int rr=0;rr<Values.GetLength(0);rr++)for(int cc=0;cc<Values.GetLength(1);cc++)expanded[rr,cc]=Values[rr,cc];Values=expanded;}Values[r-1,c-1]=value;Touch();}
    public object Get(int r,int c){return r<=Values.GetLength(0)&&c<=Values.GetLength(1)?Values[r-1,c-1]:null;}
    public string State(){var text=new StringBuilder(Name);for(int r=1;r<=Values.GetLength(0);r++){text.Append("|R"+r+":"+Row(r).Height+":"+Row(r).Hidden);for(int c=1;c<=Values.GetLength(1);c++){var s=Cell(r,c);text.Append("|"+r+","+c+":"+Get(r,c)+":"+Width(c)+":"+s.Name+":"+s.Size+":"+s.Bold+":"+s.Wrap+":"+s.FontColor+":"+s.FillColor+":"+s.Horizontal+":"+s.Vertical+":"+s.NumberFormat);}}return text.ToString();}
    public LinkFakeRange Cells{get{return Range;}}public void Activate(){}public LinkFakeRange Range{get{return new LinkFakeRange{Worksheet=this};}}public LinkFakeRange UsedRange{get{return Range;}}
    public LinkFakeHyperlinks Hyperlinks{get{return new LinkFakeHyperlinks{Sheet=this};}}
    public void Delete(){Parent.SheetItems.Remove(this);}public int Visible=-1;
}
public class LinkFakeRange {
    public LinkFakeSheet Worksheet;public string Address;public bool MergeCells;
    public LinkFakeRange CurrentRegion{get{return Worksheet.Range[Worksheet.RegionAddress];}}public LinkFakeAreas Areas{get{return new LinkFakeAreas{Range=this};}}public void Select(){}
    public void AdvancedFilter(int action,object criteria,object destination,bool unique){Worksheet.NativeFilterCalls++;Worksheet.LastCriterion=(object[,])((LinkFakeRange)criteria).Value2;var target=(LinkFakeRange)destination;target.Worksheet.Values=(object[,])Worksheet.NativeFilterResult.Clone();target.Worksheet.Range.Each(s=>{s.Size=18;s.Bold=false;s.NumberFormat="General";});}
    public int Row{get{return Bounds[0];}}public int Column{get{return Bounds[1];}}
    public int[] Bounds{get{if(string.IsNullOrEmpty(Address))return new[]{1,1,Worksheet.Values.GetLength(0),Worksheet.Values.GetLength(1)};string[] parts=Address.Replace("$","").Split(':');int[] a=Parse(parts[0]),b=Parse(parts[parts.Length-1]);return new[]{a[0],a[1],b[0],b[1]};}}
    static int[] Parse(string value){int col=0,i=0;for(;i<value.Length&&char.IsLetter(value[i]);i++)col=col*26+char.ToUpperInvariant(value[i])-64;return new[]{i==value.Length?1:int.Parse(value.Substring(i)),col};}
    public LinkFakeRange this[string address]{get{return new LinkFakeRange{Worksheet=Worksheet,Address=address};}}
    public void Each(Action<LinkFakeCellState> action){int[] b=Bounds;for(int r=b[0];r<=b[2];r++)for(int c=b[1];c<=b[3];c++)action(Worksheet.Cell(r,c));Worksheet.Touch();}
    public object Read(Func<LinkFakeCellState,object> read){int[] b=Bounds;object first=read(Worksheet.Cell(b[0],b[1]));for(int r=b[0];r<=b[2];r++)for(int c=b[1];c<=b[3];c++)if(!object.Equals(first,read(Worksheet.Cell(r,c))))return null;return first;}
    public object Font{get{return new LinkFakeFont{Range=this};}}public object Interior{get{return new LinkFakeFill{Range=this};}}
    public object NumberFormat{get{return Read(s=>s.NumberFormat);}set{Each(s=>s.NumberFormat=Convert.ToString(value));}}
    public object WrapText{get{return Read(s=>s.Wrap);}set{Each(s=>s.Wrap=Convert.ToBoolean(value));}}
    public object HorizontalAlignment{get{return Read(s=>s.Horizontal);}set{Each(s=>s.Horizontal=Convert.ToInt32(value));}}public object VerticalAlignment{get{return Read(s=>s.Vertical);}set{Each(s=>s.Vertical=Convert.ToInt32(value));}}
    public object ColumnWidth{get{int[] b=Bounds;double first=Worksheet.Width(b[1]);for(int c=b[1];c<=b[3];c++)if(first!=Worksheet.Width(c))return null;return first;}set{int[] b=Bounds;double width=Convert.ToDouble(value);for(int c=b[1];c<=b[3];c++){if(c==Worksheet.FailRestoreWidthColumn&&width==Worksheet.FailRestoreWidthValue)throw new IOException("fake width restoration failure");Worksheet.Widths[c]=width;}Worksheet.Touch();}}
    public object RowHeight{get{int[] b=Bounds;object first=Height(b[0]);for(int r=b[0];r<=b[2];r++)if(!object.Equals(first,Height(r)))return null;return first;}set{int[] b=Bounds;for(int r=b[0];r<=b[2];r++)Worksheet.Row(r).Height=Convert.ToDouble(value);Worksheet.Touch();}}
    object Height(int r){return Worksheet.Row(r).Hidden?(Worksheet.HiddenHeightNull?null:(object)0):Worksheet.Row(r).Height;}
    public object Hidden{get{int[] b=Bounds;bool first=Worksheet.Row(b[0]).Hidden;for(int r=b[0];r<=b[2];r++)if(first!=Worksheet.Row(r).Hidden)return null;return first;}set{int[] b=Bounds;for(int r=b[0];r<=b[2];r++)Worksheet.Row(r).Hidden=Convert.ToBoolean(value);Worksheet.Touch();}}
    public object EntireRow{get{return this;}}
    public LinkFakeAxis Columns{get{return new LinkFakeAxis{Range=this,IsRows=false};}}public LinkFakeAxis Rows{get{return new LinkFakeAxis{Range=this,IsRows=true};}}
    public void Clear(){Worksheet.Values=new object[1,4];Worksheet.CellStates.Clear();Worksheet.Touch();}
    public object Value2{get{int[] b=Bounds;if(b[0]==b[2]&&b[1]==b[3])return Worksheet.Get(b[0],b[1]);if(b[0]==1&&b[1]==1&&b[2]==Worksheet.Values.GetLength(0)&&b[3]==Worksheet.Values.GetLength(1))return Worksheet.Values;var data=new object[b[2]-b[0]+1,b[3]-b[1]+1];for(int r=b[0];r<=b[2];r++)for(int c=b[1];c<=b[3];c++)data[r-b[0],c-b[1]]=Worksheet.Get(r,c);return data;}set{int[] b=Bounds;var data=value as Array;if(data==null){Worksheet.Put(b[0],b[1],value);return;}for(int r=0;r<data.GetLength(0);r++)for(int c=0;c<data.GetLength(1);c++)Worksheet.Put(b[0]+r,b[1]+c,data.GetValue(r+data.GetLowerBound(0),c+data.GetLowerBound(1)));}}
}
public class LinkFakeFont {public LinkFakeRange Range;public object Bold{get{return Range.Read(s=>s.Bold);}set{Range.Each(s=>s.Bold=Convert.ToBoolean(value));}}public object Color{get{return Range.Read(s=>s.FontColor);}set{Range.Each(s=>s.FontColor=Convert.ToInt32(value));}}public object Name{get{return Range.Read(s=>s.Name);}set{Range.Each(s=>s.Name=Convert.ToString(value));}}public object Size{get{return Range.Read(s=>s.Size);}set{Range.Each(s=>s.Size=Convert.ToDouble(value));}}}
public class LinkFakeFill {public LinkFakeRange Range;public object Color{get{return Range.Read(s=>s.FillColor);}set{Range.Each(s=>s.FillColor=Convert.ToInt32(value));}}}
public class LinkFakeAreas {public LinkFakeRange Range;public int Count{get{return 1;}}public object Item(int i){return Range;}}
public class LinkFakeHyperlinks {public LinkFakeSheet Sheet;public object Add(object anchor,string address,string subAddress,string tip,string text){var cell=(LinkFakeRange)anchor;Sheet.HyperlinkTargets.Add(subAddress);cell.Value2=text;cell.Each(s=>{s.Bold=false;s.Horizontal=-4131;s.Size=17;});return cell;}}
public class LinkFakeFitSnapshot {public string Address;public string[,] Formats;public LinkFakeFitSnapshot(LinkFakeRange range){Address=range.Address;int[] b=range.Bounds;Formats=new string[b[2]-b[0]+1,b[3]-b[1]+1];for(int r=b[0];r<=b[2];r++)for(int c=b[1];c<=b[3];c++){var s=range.Worksheet.Cell(r,c);Formats[r-b[0],c-b[1]]=s.Bold+":"+s.Horizontal+":"+s.Vertical+":"+s.Size+":"+s.NumberFormat;}}}
public class LinkFakeAxis {public LinkFakeRange Range;public bool IsRows;public int Count{get{int[] b=Range.Bounds;return IsRows?b[2]-b[0]+1:b[3]-b[1]+1;}}
    public static double Need(object value,LinkFakeCellState style){string text=value is double&&style.NumberFormat=="#,##0.00"?((double)value).ToString("N2",System.Globalization.CultureInfo.InvariantCulture):Convert.ToString(value);int count=text.Sum(c=>c>127?2:1);return Math.Max(2,count*style.Size/10*(style.Bold?1.08:1)+2);}
    public void AutoFit(){var sheet=Range.Worksheet;int[] b=Range.Bounds;if(IsRows){sheet.RowFits.Add(Range.Address??"used");for(int r=b[0];r<=b[2];r++){double height=0;for(int c=b[1];c<=b[3];c++){var s=sheet.Cell(r,c);double lines=s.Wrap?Math.Max(1,Math.Ceiling(Need(sheet.Get(r,c),s)/sheet.Width(c))):1;height=Math.Max(height,(s.Size*1.5+8)*lines);}sheet.Row(r).Height=height;sheet.Touch();if(sheet.FailRowAutoFitAfter>0&&r-b[0]+1==sheet.FailRowAutoFitAfter){sheet.FailRowAutoFitAfter=0;throw new IOException("fake partial row AutoFit failure");}}}else{sheet.Fits.Add(Range.Address??"used");sheet.FitSnapshots.Add(new LinkFakeFitSnapshot(Range));for(int c=b[1];c<=b[3];c++){double width=2;for(int r=b[0];r<=b[2];r++)if(!sheet.Row(r).Hidden)width=Math.Max(width,Need(sheet.Get(r,c),sheet.Cell(r,c)));sheet.Widths[c]=Math.Min(255,width);sheet.Touch();if(sheet.FailColumnAutoFitAfter>0&&c-b[1]+1==sheet.FailColumnAutoFitAfter){sheet.FailColumnAutoFitAfter=0;throw new IOException("fake partial column AutoFit failure");}}}}
}
public class QuietFake {
    bool events=true; public bool FailDisable;
    public bool EnableEvents {get{return events;}set{if(FailDisable&&!value)throw new InvalidOperationException("fake pause refusal");events=value;}}
    public bool ScreenUpdating {get;set;}
}

