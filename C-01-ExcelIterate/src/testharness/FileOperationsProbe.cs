using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel_Iterate;

class FileOperationsProbe
{
    static int checks; static string sandbox;
    static void Check(bool value,string message) { if(!value)throw new Exception(message); checks++; }
    static void Reject(Action work,string message) { bool refused=false;try{work();}catch{refused=true;}Check(refused,message); }
    static string Root() { string root=Path.Combine(sandbox,Guid.NewGuid().ToString("N"));Directory.CreateDirectory(root);return root; }
    static string Put(string root,string relative,string content) { string path=Path.Combine(root,relative);Directory.CreateDirectory(Path.GetDirectoryName(path));File.WriteAllText(path,content);return path; }
    static FileOperationResult Organize(OrganizePlan plan,Action<string> recycle,Func<bool> cancel,Action<string,string> move) { return FileOrganizer.Execute(plan,recycle,cancel??(()=>false),_=>{},move); }
    static RenamePlan Rename(string root,params string[] names) { var notes=new List<string>();var files=FilePaths.Scan(root,false,notes);return new RenamePlan{Root=root,Items=files.Select((f,i)=>new RenameItem{Source=f,Input=names[i]}).ToList()}; }
    static FileOperationResult RunRename(RenamePlan plan,Func<bool> cancel,Action<string,string> move) { return FileRenamer.Execute(plan,cancel??(()=>false),_=>{},move); }
    static Dictionary<string,string> Contents(string root){return FilePaths.Scan(root,true,new List<string>()).ToDictionary(f=>f.Path,f=>WpsShared.SourcePaths.ReadAllText(f.Path),StringComparer.OrdinalIgnoreCase);}
    static void Same(string root,Dictionary<string,string> before,string label){var now=Contents(root);Check(before.Count==now.Count&&before.All(p=>now.ContainsKey(p.Key)&&now[p.Key]==p.Value),label);}
    static void JournalLockCheck()
    {
        var batch=OperationBatch.Create(Root(),"Rename");
        var held=File.Open(batch.Journal,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);
        var release=new System.Threading.Thread(delegate() {System.Threading.Thread.Sleep(125);held.Dispose();});
        release.IsBackground=true;release.Start();
        try {batch.State="Completed";batch.Persist();}
        finally {release.Join();}
        Check(OperationBatch.Find(batch.Root).State=="Completed","journal commits after real sharing lock is released");
    }
    sealed class JournalFailure : IOException
    {
        internal JournalFailure(int code) : base("injected journal failure") {HResult=unchecked((int)0x80070000)|code;}
    }
    static void JournalPersistenceChecks()
    {
        int attempts=0;var waits=new List<int>();
        var initial=new OperationBatch {Root=Root(),Id=Guid.NewGuid().ToString("N"),Kind="Rename"};
        initial.Persist((temporary,journal)=>{attempts++;throw new JournalFailure(1175);},waits.Add);
        Check(attempts==0&&waits.Count==0&&OperationBatch.Find(initial.Root).State=="Active","first journal is created without replacement");
        initial.State="Completed";
        initial.Persist((temporary,journal)=>{attempts++;},waits.Add);
        Check(attempts==1&&waits.Count==0,"successful replacement returns without retry or delay");
        initial.Persist();
        Check(OperationBatch.Find(initial.Root).State=="Completed"&&!File.Exists(initial.Journal+".tmp"),"real replacement commits XML and consumes temporary snapshot");
        foreach(int code in new[]{32,33,1175})
        {
            var batch=OperationBatch.Create(Root(),"Rename");batch.State="Completed";
            attempts=0;waits.Clear();byte[] snapshot=null;bool same=true;
            batch.Persist((temporary,journal)=>
            {
                byte[] candidate=File.ReadAllBytes(temporary);
                if(snapshot==null)snapshot=candidate;else same&=snapshot.SequenceEqual(candidate);
                if(++attempts<=2)throw new JournalFailure(code);
                File.Replace(temporary,journal,null);
            },waits.Add);
            Check(attempts>=3&&attempts<=5&&waits.SequenceEqual(Enumerable.Range(0,attempts-1).Select(i=>50<<i)),"transient replacement uses bounded backoff "+code);
            Check(same&&OperationBatch.Find(batch.Root).State=="Completed"&&!File.Exists(batch.Journal+".tmp"),"retry commits the same complete snapshot "+code);
        }
        foreach(int code in new[]{32,33,1175,5,112,1176,1177})
        {
            var batch=OperationBatch.Create(Root(),"Rename");byte[] original=File.ReadAllBytes(batch.Journal);batch.State="Completed";
            attempts=0;waits.Clear();IOException failure=null;JournalFailure injected=null;
            try {batch.Persist((temporary,journal)=>{attempts++;throw injected=new JournalFailure(code);},waits.Add);}
            catch(IOException ex) {failure=ex;}
            int expected=code==32||code==33||code==1175?5:1;
            Check(failure!=null&&attempts==expected&&waits.SequenceEqual(Enumerable.Range(0,expected-1).Select(i=>50<<i)),"persistent or nonretryable error is bounded "+code);
            Check(object.ReferenceEquals(failure.InnerException,injected)&&failure.Message.Contains("0x"+injected.HResult.ToString("X8"))&&failure.Message.Contains(batch.Journal)&&failure.Message.Contains(batch.Journal+".tmp")&&failure.Message.Contains("尝试"+expected+"次"),"terminal journal error preserves cause and diagnostic context "+code);
            var candidate=new System.Xml.XmlDocument();candidate.Load(batch.Journal+".tmp");
            Check(original.SequenceEqual(File.ReadAllBytes(batch.Journal))&&OperationBatch.Find(batch.Root).State=="Active"&&candidate.DocumentElement["State"].InnerText=="Completed","failed replacement retains valid original and candidate "+code);
            batch.Persist();Check(OperationBatch.Find(batch.Root).State=="Completed","retained journal can be committed after failure "+code);
        }
        foreach(bool removeTemporary in new[]{false,true})
        foreach(bool duringWait in new[]{false,true})
        {
            var batch=OperationBatch.Create(Root(),"Rename");batch.State="Completed";attempts=0;waits.Clear();
            string removed=removeTemporary?batch.Journal+".tmp":batch.Journal;IOException failure=null;
            try
            {
                batch.Persist((temporary,journal)=>{attempts++;if(!duringWait)File.Delete(removed);throw new JournalFailure(1175);},
                    delay=>{waits.Add(delay);if(duringWait)File.Delete(removed);});
            }
            catch(IOException ex) {failure=ex;}
            Check(failure!=null&&attempts==1&&waits.Count==(duringWait?1:0),"missing replacement input stops retries "+removeTemporary+"/"+duringWait);
            var retained=new System.Xml.XmlDocument();retained.Load(removeTemporary?batch.Journal:batch.Journal+".tmp");
            Check(retained.DocumentElement["State"].InnerText==(removeTemporary?"Active":"Completed"),"remaining recovery evidence is readable "+removeTemporary+"/"+duringWait);
        }
        var denied=OperationBatch.Create(Root(),"Rename");attempts=0;waits.Clear();bool rejected=false;
        try {denied.Persist((temporary,journal)=>{attempts++;throw new UnauthorizedAccessException("injected permission denial");},waits.Add);}
        catch(UnauthorizedAccessException) {rejected=true;}
        Check(rejected&&attempts==1&&waits.Count==0&&File.Exists(denied.Journal)&&File.Exists(denied.Journal+".tmp"),"permission errors propagate immediately with both records retained");

        string root=Root();Put(root,"A/a.xlsx","restore after journal lock");var before=Contents(root);
        var organized=Organize(FileOrganizer.Plan(root,new OrganizeOptions {Classify=true}),null,null,null);
        FileStream held=null;IOException blocked=null;
        try
        {
            try {organized.Batch.Undo(null,path=>{Directory.Delete(path,false);held=File.Open(organized.Batch.Journal,FileMode.Open,FileAccess.Read,FileShare.ReadWrite);});}
            catch(IOException ex) {blocked=ex;}
        }
        finally {if(held!=null)held.Dispose();}
        Check(blocked!=null&&File.Exists(organized.Batch.Journal)&&File.Exists(organized.Batch.Journal+".tmp"),"persistent real lock retains recovery evidence after files are restored");
        var reopened=OperationBatch.Find(root);Check(reopened.CanUndo,"directory progress remains recoverable after persistence failure");
        var undo=reopened.Undo(null);Check(!undo.RecoveryIncomplete,"reopened undo completes after journal lock release: "+undo.Text);
        Same(root,before,"journal retry restores original paths and contents");
        reopened.Confirm(RecycleFiles.Send);Check(OperationBatch.Find(root)==null,"confirmation cleans recovered batch records");
    }
    static void OrganizerChecks()
    {
        string previewRoot=Root();for(int i=0;i<250;i++)Put(previewRoot,"原目录/文件"+i+".txt","content");
        var compact=FileOrganizer.Plan(previewRoot,new OrganizeOptions{Classify=true});
        Check(compact.Confirmation=="移动250项；同名备份0项；未选类型暂存0项。\r\n确认操作将回收以下原目录及剩余内容：\r\n"+Path.Combine(previewRoot,"原目录"),"large batch confirmation contains counts and directory only, without blank lines");
        Check(compact.Preview.Contains("文件249.txt")&&!compact.Confirmation.Contains("文件249.txt"),"complete file details stay in log only");
        compact=FileOrganizer.Plan(Root(),new OrganizeOptions());Check(compact.Confirmation.EndsWith("无原目录需要回收")&&!compact.Confirmation.Contains("\r\n\r\n"),"zero batch compact confirmation");
        for(int mask=0;mask<8;mask++)
        {
            string root=Root();Put(root,"a.xlsx","older");Put(root,"sub/a.xlsx","latest");Put(root,"deep/note.docx","note");Put(root,"plain","no suffix");Put(root,"备份/user.txt","user-owned");
            File.SetLastWriteTimeUtc(Path.Combine(root,"a.xlsx"),DateTime.UtcNow.AddDays(-2));var before=Contents(root);
            var options=new OrganizeOptions {Replace=(mask&1)!=0,Classify=(mask&2)!=0,Filter=(mask&4)!=0};options.Types.Add("XLSX");
            var plan=FileOrganizer.Plan(root,options);Check(plan.Preview.Contains("暂存")&&plan.Preview.Contains("确认操作"),"deferred backup preview "+mask);
            var result=Organize(plan,_=>{throw new Exception("must not recycle before confirmation");},null,null);
            Check(result.Failed==0,"all option combinations "+mask+"\r\n"+result.Text);Check(Directory.Exists(Path.Combine(root,"deep")),"source folders retained");
            string dest=options.Classify?Path.Combine(root,"xlsx"):root;Check(WpsShared.SourcePaths.Exists(Path.Combine(dest,"a.xlsx")),"destination "+mask);
            if(options.Replace)Check(WpsShared.SourcePaths.ReadAllText(Path.Combine(dest,"a.xlsx"))=="latest","latest wins");else Check(WpsShared.SourcePaths.GetFiles(dest,"*.xlsx").Length==2,"numbering retains all");
            Check(result.Batch.Journal.StartsWith(root+"\\")&&WpsShared.SourcePaths.Exists(result.Batch.LogPath),"journal and log local");
            Reject(()=>Organize(FileOrganizer.Plan(root,options),null,null,null),"pending batch blocks repeat");
            var restored=OperationBatch.Find(root);Check(restored!=null&&restored.CanUndo,"reopen restores pending batch");
            var undo=restored.Undo(null);Check(!undo.RecoveryIncomplete,"complete inverse "+mask+"\r\n"+undo.Text);Same(root,before,"all original paths and contents restored "+mask);
            Check(!options.Classify||!Directory.Exists(Path.Combine(root,"xlsx")),"new classification directory removed on undo "+mask);
            restored.Confirm(RecycleFiles.Send);Check(OperationBatch.Find(root)==null,"confirmation removes only owned logs");Check(WpsShared.SourcePaths.ReadAllText(Path.Combine(root,"备份/user.txt"))=="user-owned","user backup retained");
        }
        string r=Root();Put(r,"a.txt","existing");Put(r,"one/a.txt","child");Put(r,"two/a.txt","child2");Put(r,"a（1）.txt","reserved");
        var plan1=FileOrganizer.Plan(r,new OrganizeOptions());var res=Organize(plan1,null,null,null);Check(WpsShared.SourcePaths.Exists(Path.Combine(r,"a（2）.txt"))&&WpsShared.SourcePaths.Exists(Path.Combine(r,"a（3）.txt")),"Chinese collision suffix");res.Batch.Undo(null);res.Batch.Confirm(RecycleFiles.Send);
        r=Root();string old=Put(r,"a.txt","existing"),nw=Put(r,"sub/a.txt","newer");File.SetLastWriteTimeUtc(old,DateTime.UtcNow.AddDays(-1));var plan2=FileOrganizer.Plan(r,new OrganizeOptions {Replace=true});
        res=Organize(plan2,null,null,(x,y)=>{throw new IOException("backup denied");});Check(res.Failed==1&&WpsShared.SourcePaths.ReadAllText(old)=="existing"&&WpsShared.SourcePaths.Exists(nw),"failed backup never overwrites");res.Batch.Confirm(RecycleFiles.Send);
        r=Root();old=Put(r,"sub/a.txt","content");plan2=FileOrganizer.Plan(r,new OrganizeOptions());File.AppendAllText(old,"changed");res=Organize(plan2,null,null,null);Check(res.Failed==1&&WpsShared.SourcePaths.Exists(old),"changed source not moved");res.Batch.Confirm(RecycleFiles.Send);
        r=Root();old=Put(r,"a.txt","old");nw=Put(r,"sub/a.txt","new");File.SetLastWriteTimeUtc(old,DateTime.UtcNow.AddDays(-1));var original=Contents(r);bool stop=false;
        res=Organize(FileOrganizer.Plan(r,new OrganizeOptions{Replace=true}),null,()=>stop,(x,y)=>{File.Move(x,y);stop=true;});Check(res.Cancelled&&!res.RecoveryIncomplete,"cancel after backup rolls back");Same(r,original,"cancel restored original files");res.Batch.Confirm(RecycleFiles.Send);
        r=Root();old=Put(r,"sub/a.txt","content");res=Organize(FileOrganizer.Plan(r,new OrganizeOptions()),null,null,null);File.AppendAllText(Path.Combine(r,"a.txt"),"external edit");var failed=res.Batch.Undo(null);Check(failed.RecoveryIncomplete&&WpsShared.SourcePaths.Exists(Path.Combine(r,"a.txt")),"external edit blocks undo");Reject(()=>res.Batch.Confirm(RecycleFiles.Send),"incomplete undo blocks confirmation");
        Reject(()=>FileOrganizer.Plan(r,new OrganizeOptions {Filter=true}),"filter requires a type");
        var sink=new RecycleSink();Check(sink.PreDeleteItem(0,null)<0&&sink.PreDeleteItem(0x80,null)==0,"strict recycle no permanent fallback");
        r=Root();old=Put(r,"sub/excluded.txt","excluded");var filter=new OrganizeOptions {Filter=true};filter.Types.Add("xlsx");res=Organize(FileOrganizer.Plan(r,filter),null,null,null);Check(!WpsShared.SourcePaths.Exists(old)&&WpsShared.SourcePaths.Exists(res.Batch.Files.Single().Current),"excluded file held in reversible backup");
        string journal=res.Batch.Journal;Reject(()=>res.Batch.Confirm(_=>{throw new IOException("recycle denied");}),"recycle failure reported");Check(WpsShared.SourcePaths.Exists(journal)&&OperationBatch.Find(r).State=="Cleaning","cleanup failure retains retry record");res.Batch.Confirm(RecycleFiles.Send);Check(!WpsShared.SourcePaths.Exists(journal),"real recycle retry succeeds isolated batch only");
    }
    static void DirectoryChecks()
    {
        string emptyRoot=Root();Directory.CreateDirectory(Path.Combine(emptyRoot,"A"));var emptyBatch=Organize(FileOrganizer.Plan(emptyRoot,new OrganizeOptions()),null,null,null).Batch;
        Check(emptyBatch.CanUndo&&!emptyBatch.Undo(null).RecoveryIncomplete,"cancel no-move batch cancels pending original-directory cleanup");emptyBatch.Confirm(RecycleFiles.Send);Check(Directory.Exists(Path.Combine(emptyRoot,"A")),"empty original remains after no-move undo confirmation");
        string root=Root();Put(root,"A/a.xlsx","a");Put(root,"B/b.docx","b");Directory.CreateDirectory(Path.Combine(root,"C"));Put(root,"备份/user.txt","owned");
        var result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        Check(result.Batch.OriginalDirectories.Count==3,"snapshot original first-level directories including empty C");
        Check(WpsShared.SourcePaths.Exists(Path.Combine(root,"备份/user.txt")),"existing backup excluded from scanning");
        Check(result.Batch.ConfirmationText().Contains("后来新增内容")&&result.Batch.ConfirmationText().Contains(Path.Combine(root,"A")),"confirmation names complete directory recycle");
        Put(root,"A/late.txt","late");Directory.CreateDirectory(Path.Combine(root,"later"));
        result.Batch.Confirm(RecycleFiles.Send);
        Check(!Directory.Exists(Path.Combine(root,"A"))&&!Directory.Exists(Path.Combine(root,"B"))&&!Directory.Exists(Path.Combine(root,"C")),"actual recycle original A B C with remaining content");
        Check(WpsShared.SourcePaths.Exists(Path.Combine(root,"xlsx/a.xlsx"))&&WpsShared.SourcePaths.Exists(Path.Combine(root,"docx/b.docx")),"classified results retained");
        Check(Directory.Exists(Path.Combine(root,"later"))&&WpsShared.SourcePaths.Exists(Path.Combine(root,"备份/user.txt")),"later directory and original backup protected");
        root=Root();Put(root,"A/a.xlsx","a");Put(root,"docx/existing.docx","b");
        result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        Check(!result.Batch.OriginalDirectories.Any(d=>d.Path==Path.Combine(root,"docx")),"existing classification result excluded from original cleanup");
        var undo=result.Batch.Undo(null);Check(!undo.RecoveryIncomplete&&!Directory.Exists(Path.Combine(root,"xlsx"))&&Directory.Exists(Path.Combine(root,"docx")),"undo deletes newly-created result only");
        result.Batch.Confirm(RecycleFiles.Send);Check(Directory.Exists(Path.Combine(root,"A")),"confirm undone keeps original directories");
        root=Root();Put(root,"A/a.xlsx","a");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        Put(root,"xlsx/outsider.txt","outside");undo=result.Batch.Undo(null);
        Check(undo.RecoveryIncomplete&&WpsShared.SourcePaths.Exists(Path.Combine(root,"A/a.xlsx"))&&WpsShared.SourcePaths.Exists(Path.Combine(root,"xlsx/outsider.txt")),"restore files but preserve outsider directory content");
        var reopened=OperationBatch.Find(root);Check(reopened.CanUndo&&reopened.State=="UndoFailed","directory-only undo retry survives reopening");
        Reject(()=>reopened.Confirm(RecycleFiles.Send),"incomplete directory undo blocks confirmation");
        File.Move(Path.Combine(root,"xlsx/outsider.txt"),Path.Combine(root,"outsider.txt"));undo=reopened.Undo(null);
        Check(!undo.RecoveryIncomplete&&!Directory.Exists(Path.Combine(root,"xlsx")),"directory-only retry completes");reopened.Confirm(RecycleFiles.Send);
        root=Root();Put(root,"A/a.xlsx","a");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        undo=result.Batch.Undo(null,_=>{throw new IOException("injected directory denial");});Check(undo.RecoveryIncomplete&&result.Batch.CanUndo,"delete failure remains retryable after files restored");
        Check(!OperationBatch.Find(root).Undo(null).RecoveryIncomplete,"directory delete retry from persisted record");OperationBatch.Find(root).Confirm(RecycleFiles.Send);
        root=Root();Put(root,"A/a.xlsx","a");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        Directory.Move(Path.Combine(root,"xlsx"),Path.Combine(root,"moved-result"));Directory.CreateDirectory(Path.Combine(root,"xlsx"));File.Move(Path.Combine(root,"moved-result/a.xlsx"),Path.Combine(root,"xlsx/a.xlsx"));
        undo=result.Batch.Undo(null);Check(undo.RecoveryIncomplete&&Directory.Exists(Path.Combine(root,"xlsx")),"replacement created directory protected even when empty after restoring file");
        root=Root();Put(root,"A/a.xlsx","a");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        Directory.Move(Path.Combine(root,"A"),Path.Combine(root,"replacement-old"));Directory.CreateDirectory(Path.Combine(root,"A"));Put(root,"A/new.txt","new");
        Reject(()=>result.Batch.Confirm(RecycleFiles.Send),"same-path original replacement stops recycling");Check(WpsShared.SourcePaths.Exists(Path.Combine(root,"A/new.txt"))&&WpsShared.SourcePaths.Exists(result.Batch.Journal),"replacement and cleanup record retained");
        root=Root();Put(root,"A/a.xlsx","a");Put(root,"B/b.docx","b");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);
        string a=Path.Combine(root,"A"),b=Path.Combine(root,"B");Reject(()=>result.Batch.Confirm(p=>{if(p==b)throw new IOException("injected recycle denial");RecycleFiles.Send(p);}),"partial cleanup failure");
        var retry=OperationBatch.Find(root);Check(!Directory.Exists(a)&&Directory.Exists(b)&&retry.OriginalDirectories.First(d=>d.Path==a).Done,"per-directory cleanup progress persisted");
        Directory.CreateDirectory(a);Put(root,"A/new.txt","new");retry.Confirm(RecycleFiles.Send);Check(WpsShared.SourcePaths.Exists(Path.Combine(root,"A/new.txt"))&&!Directory.Exists(b),"cleanup retry never touches replacement of already recycled directory");
        root=Root();Put(root,"A/a.xlsx","a");result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,null,null);result.Batch.OriginalDirectories.Clear();result.Batch.Persist();OperationBatch.Find(root).Confirm(RecycleFiles.Send);
        Check(Directory.Exists(Path.Combine(root,"A")),"old record without original snapshot never infers cleanup scope");
        root=Root();Put(root,"A/a.xlsx","a");bool stop=false;result=Organize(FileOrganizer.Plan(root,new OrganizeOptions{Classify=true}),null,()=>stop,(x,y)=>{File.Move(x,y);stop=true;});
        Check(result.Cancelled&&!result.RecoveryIncomplete&&!Directory.Exists(Path.Combine(root,"xlsx"))&&WpsShared.SourcePaths.Exists(Path.Combine(root,"A/a.xlsx")),"in-flight cancellation removes new category");result.Batch.Confirm(RecycleFiles.Send);
    }
    static void RenameChecks()
    {
        Check(FileRenamer.NewName("新文件","old.xlsx")=="新文件.xlsx","stem keeps suffix");Check(FileRenamer.NewName("new.docx","old.xlsx")=="new.docx","explicit suffix");
        foreach(string name in new[]{"../x","CON","nul.txt","a/b","bad:","trailing.","LPT1.pdf"})Reject(()=>FileRenamer.NewName(name,"old.txt"),"invalid name "+name);
        string root=Root();Put(root,"A.txt","A");Put(root,"B.txt","B");Put(root,"C.txt","C");Put(root,"sub/ignored.txt","ignored");var original=Contents(root);
        var plan=Rename(root,"B.txt","C.txt","A.txt");Check(plan.Items.Count==3,"first layer only");var result=RunRename(plan,null,null);Check(result.Success==3&&WpsShared.SourcePaths.ReadAllText(Path.Combine(root,"A.txt"))=="C","cycle commit");
        var reopened=OperationBatch.Find(root);Check(reopened.CanUndo,"reopen rename pending");Check(!reopened.Undo(null).RecoveryIncomplete,"inverse cycle");Same(root,original,"cycle rollback original bytes");reopened.Confirm(RecycleFiles.Send);
        plan=Rename(root,"a.txt","","");result=RunRename(plan,null,null);Check(result.Success==1&&result.Skipped==2&&WpsShared.SourcePaths.GetFiles(root).Any(f=>Path.GetFileName(f)=="a.txt"),"case-only forward");result.Batch.Undo(null);Check(WpsShared.SourcePaths.GetFiles(root).Any(f=>Path.GetFileName(f)=="A.txt"),"case-only inverse exact casing");result.Batch.Confirm(RecycleFiles.Send);
        plan=Rename(root,"dup","dup","");Check(RunRename(plan,null,null).Failed>0,"duplicate target pauses");plan=Rename(root,"B.txt","","");Check(RunRename(plan,null,null).Failed>0,"untouched target occupied");
        plan=Rename(root,"new1","new2","new3");using(File.Open(plan.Items[0].Source.Path,FileMode.Open,FileAccess.ReadWrite,FileShare.None))Check(RunRename(plan,null,null).Failed>0,"locked source pauses");
        foreach(int failAt in new[]{2,4,6})
        {
            root=Root();Put(root,"A.txt","A");Put(root,"B.txt","B");Put(root,"C.txt","C");original=Contents(root);plan=Rename(root,"B.txt","C.txt","A.txt");int n=0;
            result=RunRename(plan,null,(x,y)=>{if(++n==failAt)throw new IOException("injected move failure");File.Move(x,y);});Check(result.Failed>0&&!result.RecoveryIncomplete,"failed submit restores "+failAt);Same(root,original,"whole rollback "+failAt);Check(WpsShared.SourcePaths.Exists(result.RecoveryPath),"local recovery retained");result.Batch.Confirm(RecycleFiles.Send);
        }
        root=Root();Put(root,"A.txt","A");Put(root,"B.txt","B");original=Contents(root);int polls=0;result=RunRename(Rename(root,"B.txt","A.txt"),()=>++polls==3,null);Check(result.Cancelled&&!result.RecoveryIncomplete,"cancel staging inverse");Same(root,original,"cancel restores originals");result.Batch.Confirm(RecycleFiles.Send);
        int calls=0;result=RunRename(Rename(root,"B.txt","A.txt"),null,(x,y)=>{if(++calls>=3)throw new IOException("blocked commit and restore");File.Move(x,y);});Check(result.RecoveryIncomplete&&WpsShared.SourcePaths.Exists(result.RecoveryPath),"failed inverse retains record");Reject(()=>result.Batch.Confirm(RecycleFiles.Send),"incomplete rename refuses cleanup");
        root=Root();Put(root,"A.xlsx","A");Put(root,"B.docx","B");plan=Rename(root,"newA","newB.docx");object[,] values=new object[3,4];for(int c=0;c<4;c++)values[0,c]=RenameWorkbookSession.Headers[c];for(int i=0;i<2;i++){values[i+1,0]=plan.Items[1-i].Source.Path;values[i+1,1]=plan.Items[1-i].Source.Name;values[i+1,2]="new"+i;}
        Check(RenameWorkbookSession.Parse(root,plan.Items.Select(i=>i.Source).ToList(),values).Items[0].Source.Name=="B.docx","sorted identities");values[0,1]="edit";Reject(()=>RenameWorkbookSession.Parse(root,plan.Items.Select(i=>i.Source).ToList(),values),"edited header refused");values[0,1]=RenameWorkbookSession.Headers[1];values[1,0]="changed";Reject(()=>RenameWorkbookSession.Parse(root,plan.Items.Select(i=>i.Source).ToList(),values),"edited original refused");
    }
    static void SavingChecks()
    {
        string root=Root(),file=Put(root,"book.xlsx","placeholder");var book=new SavingFake{FullName=file,Saved=false};WorkbookSaving.Save(book,true);Check(book.SaveCount==1&&book.Saved,"dirty workbook saved through real method");
        book.Saved=false;book.Fail=true;Reject(()=>WorkbookSaving.Save(book,true),"save failure propagated");book.Fail=false;book.StillDirty=true;Reject(()=>WorkbookSaving.Save(book,true),"dirty after save refused");
        book.ReadOnly=true;Reject(()=>WorkbookSaving.Save(book,true),"readonly refused");
        book=new SavingFake{FullName="Book1",Saved=false};WorkbookSaving.ChooseNewPath=()=>null;Reject(()=>WorkbookSaving.Save(book,true),"new workbook save cancellation");
        WorkbookSaving.ChooseNewPath=()=>Path.Combine(root,"first.xlsx");WorkbookSaving.Save(book,true);Check(book.SaveAsCount==1&&WpsShared.SourcePaths.Exists(book.FullName),"new workbook native-choice seam SaveAs format51");WorkbookSaving.ChooseNewPath=null;
        using(var form=new OrganizeFilesForm())Check(!form.Replace.Checked&&!form.Classify.Checked&&!form.Filter.Checked,"all organize options default off");
        root=Root();string original=Put(root,"sub/a.txt","UI pending restore");var batchResult=Organize(FileOrganizer.Plan(root,new OrganizeOptions()),null,null,null);
        using(var form=new OrganizeFilesForm())
        {
            IntPtr handle=form.Handle;form.FolderPath.Text=root;
            Check(form.Stop.Enabled&&form.Confirm.Enabled&&!form.Preview.Enabled,"reopened completed form enables undo and confirm, blocks next batch");
            typeof(Button).GetMethod("OnClick",System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic).Invoke(form.Stop,new object[]{EventArgs.Empty});
            DateTime until=DateTime.UtcNow.AddSeconds(10);
            while(!WpsShared.SourcePaths.Exists(original)&&DateTime.UtcNow<until){Application.DoEvents();System.Threading.Thread.Sleep(10);}
            // The worker's file move can precede its posted UI completion.
            for(int i=0;i<30;i++){Application.DoEvents();System.Threading.Thread.Sleep(10);}
            Check(WpsShared.SourcePaths.Exists(original)&&!form.Stop.Enabled&&form.Confirm.Enabled,"actual cancel click restores completed batch and keeps confirmation available");
        }
        OperationBatch.Find(root).Confirm(RecycleFiles.Send);
    }
    static void DeleteOwnedTree(string path)
    {
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("Do not traverse a test reparse point: " + path);
        foreach (string file in WpsShared.SourcePaths.GetFiles(path))
        {
            if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("Do not delete a test link: " + file);
            File.Delete(file);
        }
        foreach (string directory in Directory.GetDirectories(path)) DeleteOwnedTree(directory);
        Directory.Delete(path);
    }
    [STAThread] static int Main()
    {
        sandbox=Path.Combine(Path.GetTempPath(),"EI-FeatureProbe-"+Guid.NewGuid().ToString("N"));Directory.CreateDirectory(sandbox);
        int exit=0;
        try {JournalLockCheck();JournalPersistenceChecks();OrganizerChecks();DirectoryChecks();RenameChecks();SavingChecks();}
        catch(Exception ex){Console.WriteLine("FAIL "+ex);exit=1;}
        finally
        {
            WorkbookSaving.ChooseNewPath=null;
            if(exit==0)
            {
                try
                {
                    string owned=Path.GetFullPath(sandbox), temp=Path.GetFullPath(Path.GetTempPath()).TrimEnd(Path.DirectorySeparatorChar);
                    if(!string.Equals(Path.GetDirectoryName(owned),temp,StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(owned).StartsWith("EI-FeatureProbe-",StringComparison.Ordinal))
                        throw new IOException("Unsafe test cleanup target: "+owned);
                    DeleteOwnedTree(owned);
                }
                catch(Exception ex){Console.WriteLine("FAIL test cleanup: "+ex);exit=1;}
            }
            if(exit!=0)Console.WriteLine("Failed isolated evidence retained: "+sandbox);
        }
        if(exit==0)Console.WriteLine("PASS FileOperationsProbe: "+checks+" checks; isolated files cleaned; real recycle tested.");
        return exit;
    }
}
public class SavingFake
{
    public string FullName;public bool Saved,ReadOnly,Fail,StillDirty;public int SaveCount,SaveAsCount;
    public void Save(){SaveCount++;if(Fail)throw new IOException("fake save failure");Saved=!StillDirty;}
    public void SaveAs(string path,int format){if(format!=51)throw new Exception("format");SaveAsCount++;FullName=path;File.WriteAllText(path,"new book");Saved=true;}
}
