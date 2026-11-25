using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using Excel_Iterate;

partial class WorkbookToolsProbe
{
    static void BreakLinkChecks(string folder,string a,string b,string c)
    {
        string target=Path.Combine(folder,"断链 #100% .xlsx");
        Package(target,a,b);WriteFormulaSample(target,a,b);
        byte[] before=File.ReadAllBytes(target);
        var session=JournalSession(c,new[]{target},a,"  "+LinkAssistantSession.BreakCommand+"  ",true);
        var book=((LinkFakeApp)session.App).Workbooks.Items[1];
        session.Execute(null);
        Check(session.Batch.Changes.Single().Action==LinkAction.Break&&session.Batch.Changes[0].Applied,"exact trimmed command persists a break action");
        Check(XlsxLinkPackage.Scan(target).Select(l=>l.SourcePath).SequenceEqual(new[]{b}),"break removes selected source and preserves other link");
        var xml=SheetXml(target);Check(xml.SelectSingleNode("//c[@r='A1']/f")==null&&xml.SelectSingleNode("//c[@r='A1']/v").InnerText=="12345","break converts selected formula to cached value");
        Check(xml.SelectSingleNode("//c[@r='B1']/f")!=null&&xml.SelectSingleNode("//c[@r='B1']/v").InnerText=="10","unselected formula and value preserved");
        Check(book.CloseCount == 0 && session.Batch.Checkpoints.Count == 0 &&
            !Directory.EnumerateFiles(folder, ".EI_LinkBackup_" + session.Batch.Id + "_*.xlsx").Any(),
            "open break saves the result and removes the temporary protection copy");

        Package(target, a, b); before = File.ReadAllBytes(target);
        session = JournalSession(c, new[] { target }, a, LinkAssistantSession.BreakCommand, true);
        book = ((LinkFakeApp)session.App).Workbooks.Items[1];
        StopAfterSave(session, () => { });
        var checkpoint = session.Batch.Checkpoints.Single();
        Check(session.Batch.State == "UndoFailed" && book.CloseCount == 0 && File.ReadAllBytes(checkpoint.Backup).SequenceEqual(before) &&
            session.Operations.Text.Contains(checkpoint.Backup),
            "incomplete automatic rollback leaves the open workbook and reports its verified copy");
        using (var form = new LinkAssistantForm(new LinkAssistantSession(session.App)))
            Check(form.ScanButton.Enabled && !form.ExecuteButton.Enabled, "an unresolved batch does not block a new scan");

        Package(target, a, b); before = File.ReadAllBytes(target);
        session = JournalSession(c, new[] { target }, a, LinkAssistantSession.BreakCommand, false);
        ((LinkFakeApp)session.App).Workbooks.SupportReopen = true;
        long changedLength = 0;
        StopAfterSave(session, () =>
        {
            using (var stream = File.Open(target, FileMode.Open, FileAccess.ReadWrite))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Update))
            {
                zip.GetEntry("customXml/item1.xml").Delete();
                Part(zip, "customXml/item1.xml", "<content>new user edit</content>");
            }
            changedLength = new FileInfo(target).Length;
        });
        Check(session.Batch.State == "UndoFailed" && new FileInfo(target).Length == changedLength &&
            File.ReadAllBytes(session.Batch.Checkpoints.Single().Backup).SequenceEqual(before),
            "concurrent edits prevent whole-file overwrite and retain the original protection copy");

        Package(target,a,b);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        session.Links.Add(new ExternalLinkInfo {OwnerPath=target,SourcePath=b});var report=((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet;
        report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{target,a,LinkAssistantSession.BreakCommand,""},{target,b,c,""}};
        StopAfterSave(session, () =>
        {
            Check(XlsxLinkPackage.Scan(target).Select(l => l.SourcePath).SequenceEqual(new[] { c }) &&
                session.Batch.Changes.All(x => x.Applied), "mixed owner replaces first and breaks only the requested source");
            Check(((LinkFakeApp)session.App).Workbooks.Count == 1, "the only closed workbook was the assistant's stage");
        });
        Check(File.ReadAllBytes(target).SequenceEqual(before) && session.Batch.Checkpoints.Count == 0,
            "automatic whole-file rollback restores both changes and clears protection");

        Package(target,a,b);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        session.Links.Add(new ExternalLinkInfo {OwnerPath=target,SourcePath=b});report=((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).Sheet;
        report.Values=new object[,]{{"文档路径","外部链接","新链接（待填）","处理状态"},{target,a,LinkAssistantSession.BreakCommand,""},{target,b,a,""}};
        session.Execute(null);Check(File.ReadAllBytes(target).SequenceEqual(before)&&session.Batch.Checkpoints.Count==0,"replacement cannot merge into a source being broken");

        foreach(bool refuse in new[]{true,false})
        {
            Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.RefuseBreak=refuse;book.NoBreak=!refuse;
            session.Execute(null);
            Check(session.LastResult.Contains("失败1条") && book.CloseCount == 0 && File.ReadAllBytes(target).SequenceEqual(before),
                "unsupported or ineffective native break leaves the original formula bytes and reports failure");
        }
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.FailSaveAfterBreak=true;
        session.Execute(null);
        Check(session.Batch.State == "UndoFailed" && session.Batch.Changes[0].Intent && book.CloseCount == 0 &&
            File.ReadAllBytes(session.Batch.Checkpoints.Single().Backup).SequenceEqual(before),
            "save failure leaves the uncertain open workbook and its original bytes available for diagnosis");

        string missing=Path.Combine(folder,"already missing.xlsx");Package(target,missing);session=JournalSession(c,new[]{target},missing,LinkAssistantSession.BreakCommand,true);session.Execute(null);Check(XlsxLinkPackage.Scan(target).Count==0,"missing source file still supports break from cached values");
        Package(target,a);session=JournalSession(c,new[]{target},a,"请断开该链接！",true);Reject(()=>session.Execute(null),"near command text is not a break action");Check(!session.Links[0].BreakRequested,"near command leaves action unset");
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        bool stop=false;var iterator=session.Steps(()=>stop).GetEnumerator();Check(iterator.MoveNext(),"cancellation fixture reaches checkpoint");stop=true;while(iterator.MoveNext()){}iterator.Dispose();Check(File.ReadAllBytes(target).SequenceEqual(before)&&session.Batch.Checkpoints.Count==0,"cancellation before host call commits nothing or retains a copy");
        Package(target,a);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];book.SourceStyle="uri";book.RetainHistorical=true;session.Execute(null);Check(XlsxLinkPackage.Scan(target).Count==0,"raw URI name and stale in-memory source handled");
        var legacy="<LinkBatch><ReportBook>"+System.Security.SecurityElement.Escape(c)+"</ReportBook><Changes><LinkChange><Owner>old</Owner><Old>A</Old><New>B</New></LinkChange></Changes></LinkBatch>";
        using(var reader=new StringReader(legacy)){var old=(LinkBatch)new XmlSerializer(typeof(LinkBatch)).Deserialize(reader);Check(old.Changes[0].Action==LinkAction.Replace&&old.Checkpoints.Count==0,"legacy journals default to replace and no checkpoint");}
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,true);book=((LinkFakeApp)session.App).Workbooks.Items[1];
        iterator=session.Steps(()=>false).GetEnumerator();Check(iterator.MoveNext(),"new edit fixture reaches backup");book.Saved=false;while(iterator.MoveNext()){}iterator.Dispose();Check(!book.Saved&&book.SaveCount==0&&File.ReadAllBytes(target).SequenceEqual(before),"new unsaved edits after checkpoint are never silently saved or overwritten");
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        stop=false;iterator=session.Steps(()=>stop).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage cancel fixture reaches native stage result");Check(File.ReadAllBytes(target).SequenceEqual(before),"stage native call leaves original untouched until commit");stop=true;while(iterator.MoveNext()){}iterator.Dispose();Check(File.ReadAllBytes(target).SequenceEqual(before)&&((LinkFakeApp)session.App).Workbooks.Count==1,"stage cancellation closes owned copy and restores nothing into original");
        Package(target, a); before = File.ReadAllBytes(target);
        session = JournalSession(c, new[] { target }, a, LinkAssistantSession.BreakCommand, false);
        ((LinkFakeApp)session.App).Workbooks.SupportReopen = true;
        StopAfterSave(session, () =>
        {
            checkpoint = session.Batch.Checkpoints.Single();
            checkpoint.PreparedHash = checkpoint.LastHash;
            checkpoint.LastHash = checkpoint.BeforeHash;
            checkpoint.CommitIntent = true;
            checkpoint.Applied = false;
            session.Batch.Changes[0].Applied = false;
            session.Batch.Changes[0].Intent = true;
            session.Batch.Persist();
        });
        Check(File.ReadAllBytes(target).SequenceEqual(before) && session.Batch.Checkpoints.Count == 0,
            "interrupted commit is reconciled by prepared digest during automatic rollback");

        Package(target, a);
        session = JournalSession(c, new[] { target }, a, LinkAssistantSession.BreakCommand, false);
        ((LinkFakeApp)session.App).Workbooks.SupportReopen = true;
        string appliedHash = null;
        StopAfterSave(session, () =>
        {
            File.AppendAllText(session.Batch.Checkpoints.Single().Backup, "damaged");
            appliedHash = XlsxLinkPackage.Hash(target);
        });
        Check(session.Batch.State == "UndoFailed" && XlsxLinkPackage.Hash(target) == appliedHash &&
            File.Exists(session.Batch.Checkpoints.Single().Backup),
            "damaged protection copy prevents rollback and leaves the current data untouched");

        Package(target, a);
        session = JournalSession(target, new[] { target }, a, LinkAssistantSession.BreakCommand, false);
        session.Execute(null);
        var fittedReceiver=(LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook;
        Check(fittedReceiver.SavedReports.Last()==fittedReceiver.Sheet.State()&&fittedReceiver.SavedReports.Last().Contains("已断开"),"break receiver saves final status together with the complete fitted formats");
        using(var stream=File.OpenRead(target))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))using(var reader=new StreamReader(zip.GetEntry("customXml/probe-report.xml").Open()))Check(reader.ReadToEnd()==fittedReceiver.SavedReports.Last(),"break checkpoint fingerprints actual disk report values and formats, not only a Save call");
        Check(session.Batch.Checkpoints.Count == 0 && fittedReceiver.SavedHashes.Last() == XlsxLinkPackage.Hash(target) &&
            fittedReceiver.Sheet.Cell(1,4).Bold && fittedReceiver.Sheet.Cell(1,4).Horizontal == -4108 && !fittedReceiver.Sheet.Cell(1,4).Wrap,
            "report write persists final formatting and completes the batch");
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.ForeignOpen=true;session.Execute(null);
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&((LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook).CloseCount==0,"foreign object returned for stage is never owned or closed");
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        iterator=session.Steps(()=>false).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage edit fixture reaches saved stage");var editedStage=((LinkFakeApp)session.App).Workbooks.Items[1];editedStage.Saved=false;while(iterator.MoveNext()){}iterator.Dispose();
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&editedStage.CloseCount==0&&WpsShared.SourcePaths.Exists(editedStage.FullName),"unexpected user edits in owned stage are preserved rather than closed or committed");editedStage.Close(false);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(c,new[]{target},a,LinkAssistantSession.BreakCommand,false);((LinkFakeApp)session.App).Workbooks.SupportReopen=true;
        stop=false;iterator=session.Steps(()=>stop).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"stage edit cancellation reaches saved stage");editedStage=((LinkFakeApp)session.App).Workbooks.Items[1];editedStage.Saved=false;stop=true;while(iterator.MoveNext()){}iterator.Dispose();
        Check(File.ReadAllBytes(target).SequenceEqual(before)&&editedStage.CloseCount==0&&WpsShared.SourcePaths.Exists(editedStage.FullName),"cancellation preserves newly edited stage without touching original");editedStage.Close(false);
        Package(target,a);before=File.ReadAllBytes(target);session=JournalSession(target,new[]{target},a,LinkAssistantSession.BreakCommand,false);book=(LinkFakeBook)((LinkFakeApp)session.App).ActiveWorkbook;
        iterator=session.Steps(()=>false).GetEnumerator();Check(iterator.MoveNext()&&iterator.MoveNext(),"report-write failure reaches saved break");book.Sheet.FailWrite=true;while(iterator.MoveNext()){}iterator.Dispose();
        Check(session.Batch.State == "Completed" && session.Batch.Checkpoints.Count == 0 &&
            XlsxLinkPackage.Scan(target).Count == 0 && session.Operations.Text.Contains("清单状态未完整保存"),
            "report write failure is logged and keeps the successfully saved break");
    }
    static XmlDocument SheetXml(string file)
    {var doc=new XmlDocument();using(var stream=File.OpenRead(file))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))using(var reader=zip.GetEntry("xl/worksheets/sheet1.xml").Open())doc.Load(reader);return doc;}
    static void WriteFormulaSample(string file,string first,string second)
    {
        using(var stream=File.Open(file,FileMode.Open,FileAccess.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Update))
        {zip.GetEntry("xl/worksheets/sheet1.xml").Delete();Part(zip,"xl/worksheets/sheet1.xml","<worksheet><c r=\"A1\"><f>'["+Path.GetFileName(first)+"]Sheet1'!A1</f><v>12345</v></c><c r=\"B1\"><f>'["+Path.GetFileName(second)+"]Sheet1'!A1</f><v>10</v></c></worksheet>");}
    }
}
public partial class LinkFakeBook
{
    public bool RefuseBreak,NoBreak,FailSaveAfterBreak;
    private byte[] pendingBreak;
    public void BreakLink(string name,int type)
    {
        if(RefuseBreak)throw new IOException("fake native BreakLink refusal");if(NoBreak){Saved=false;return;}
        string source=SourceStyle=="uri"?new Uri(name).LocalPath:name;
        using(var memory=new MemoryStream())
        {
            byte[] original=File.ReadAllBytes(FullName);memory.Write(original,0,original.Length);
            using(var zip=new ZipArchive(memory,ZipArchiveMode.Update,true))
            {
                foreach(var entry in zip.Entries.Where(e=>e.FullName.StartsWith("xl/externalLinks/_rels/")).ToList())
                {
                    var xml=new XmlDocument();using(var stream=entry.Open())xml.Load(stream);
                    if(xml.DocumentElement.ChildNodes.OfType<XmlElement>().Any(r=>string.Equals(XlsxLinkPackage.Resolve(FullName,r.GetAttribute("Target")),source,StringComparison.OrdinalIgnoreCase)))entry.Delete();
                }
                var sheet=zip.GetEntry("xl/worksheets/sheet1.xml");var doc=new XmlDocument();using(var stream=sheet.Open())doc.Load(stream);
                foreach(XmlNode formula in doc.SelectNodes("//f").Cast<XmlNode>().ToList())if(formula.InnerText.Contains("["+Path.GetFileName(source)+"]"))formula.ParentNode.RemoveChild(formula);
                sheet.Delete();using(var writer=new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open(),new UTF8Encoding(false)))writer.Write(doc.OuterXml);
            }
            pendingBreak=memory.ToArray();
        }
        Saved=false;if(FailSaveAfterBreak){failSave=true;FailSaveAfterBreak=false;}
    }
}
