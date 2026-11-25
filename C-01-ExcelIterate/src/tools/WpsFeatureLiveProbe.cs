using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Forms;
using System.Linq;
using System.IO.Compression;
using System.Xml;
using System.Text.RegularExpressions;
class WpsFeatureLiveProbe
{
    static List<object> owned=new List<object>();
    static dynamic Own(object x){if(x!=null)owned.Add(x);return x;}
    static BindingFlags Hidden=BindingFlags.Instance|BindingFlags.Static|BindingFlags.Public|BindingFlags.NonPublic;
    static Assembly dll;
    static object Call(string type,string method, params object[] args){return dll.GetType("Excel_Iterate."+type,true).GetMethod(method,Hidden).Invoke(null,args);}
    static object New(string type, params object[] args){return Activator.CreateInstance(dll.GetType("Excel_Iterate."+type,true),Hidden,null,args,null);}
    static object Invoke(object obj,string method,params object[]args){return obj.GetType().GetMethod(method,Hidden).Invoke(obj,args);}
    static string Hash(string f){return (string)Call("XlsxLinkPackage","Hash",f);}
    static void Check(bool value,string name){if(!value)throw new Exception("FAIL: "+name);Console.WriteLine("PASS: "+name);Console.Out.Flush();}
    [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int reserved,out IRunningObjectTable table);
    [DllImport("ole32.dll")] static extern int CreateBindCtx(int reserved,out IBindCtx context);
    static object Attach(string original)
    {
        try {return Marshal.GetActiveObject("KET.Application");}
        catch(COMException ex) {if(ex.HResult!=unchecked((int)0x800401E3))throw;}
        // Some WPS sessions publish the workbook moniker, but not the Application
        // moniker. Bind only the already-open supplied workbook; never activate
        // a new application or start a process as a connection workaround.
        IRunningObjectTable rot=null;IBindCtx ctx=null;IEnumMoniker enumerator=null;
        try {
            Marshal.ThrowExceptionForHR(GetRunningObjectTable(0,out rot));Marshal.ThrowExceptionForHR(CreateBindCtx(0,out ctx));rot.EnumRunning(out enumerator);
            IMoniker[] moniker=new IMoniker[1];
            while(enumerator.Next(1,moniker,IntPtr.Zero)==0)try {
                string name;moniker[0].GetDisplayName(ctx,null,out name);
                if(!string.Equals(name,original,StringComparison.OrdinalIgnoreCase))continue;
                object workbook=null;try {rot.GetObject(moniker[0],out workbook);object app=((dynamic)workbook).Application;Console.WriteLine("Connected through existing workbook ROT moniker: "+name);return app;}
                finally {if(workbook!=null&&Marshal.IsComObject(workbook))Marshal.ReleaseComObject(workbook);}
            }finally {Marshal.ReleaseComObject(moniker[0]);}
            throw new InvalidOperationException("没有找到已运行的 WPS 表格实例或所提供工作簿的 ROT 登记，请打开该工作簿后重试。");
        }finally {if(enumerator!=null)Marshal.ReleaseComObject(enumerator);if(ctx!=null)Marshal.ReleaseComObject(ctx);if(rot!=null)Marshal.ReleaseComObject(rot);}
    }
    [STAThread] static int Main(string[] args)
    {
        Console.OutputEncoding=new UTF8Encoding(false);
        dll=Assembly.LoadFrom(args[0]); string folder=Path.GetFullPath(args[1]);Directory.CreateDirectory(folder);
        string original=Path.GetFullPath(args[2]);bool originalMode=args.Length>3&&args[3]=="--mapped-links-original";
        string copy=originalMode?original:Path.Combine(folder,"功能测试-"+Guid.NewGuid().ToString("N")+".xlsx");
        string hash=Hash(original); if(!originalMode)File.Copy(original,copy,false);
        object app=null, test=null, navigation=null; object originalBook=null,originalSheet=null,originalSelection=null;
        bool? saved=null;object events=null,screen=null;int status=1;string originalAddress=null;
        try {
            app=Attach(original);dynamic a=app;
            Console.WriteLine("HOST="+a.Name+"; VERSION="+a.Version+"; DLL="+args[0]+"; SHA256="+Hash(args[0]));
            originalBook=Own(a.ActiveWorkbook); originalSheet=Own(a.ActiveSheet);originalSelection=Own(a.Selection);
            originalAddress=(string)((dynamic)originalSelection).Address;
            saved=(bool)((dynamic)originalBook).Saved;events=a.EnableEvents;screen=a.ScreenUpdating;
            dynamic books=Own(a.Workbooks); test=originalMode?Own(InvokeStaticFind(app,original)):Own(books.Open(copy,0,false));if(test==null)throw new Exception("指定实验工作簿未打开。");dynamic book=test,sheets=Own(book.Worksheets);
            if(args.Length>3&&args[3].StartsWith("--mapped-links",StringComparison.Ordinal)){TestMappedLinks(app,test,copy,original,folder,originalMode);status=0;return status;}
            // Directory checks must not depend on a manually maintained advanced-
            // filter reference sheet or on a particular source worksheet name.
            if(args.Length>3 && args[3]=="--directory-only") {TestDirectory(app,test);status=0;return status;}
            if(args.Length>3 && args[3]=="--links-only") {TestLinks(app,test,copy,folder,originalBook,originalSheet,originalSelection);test=null;status=0;return status;}
            dynamic sheet=Own(sheets.Item("实验区"));
            sheet.Activate();dynamic condition=Own(sheet.Range["A2:B2,A12:B12"]);condition.Select();
            int count=(int)Call("QuickAdvancedFilter","Run",app);
            Check(count==9,"Native AdvancedFilter: two AND rows / OR = 9 records");
            dynamic result=Own(sheets.Item("高筛结果(2)")), resultRows=Own(result.UsedRange.Rows);
            Check((int)resultRows.Count==10,"Existing result preserved; new output has header + 9");
            object[,] actual=Own(result.Range["A1:E10"]).Value2 as object[,];
            object[,] expected=Own(sheets.Item("高筛结果").Range["A1:E10"]).Value2 as object[,];
            Check(actual!=null&&expected!=null,"WPS result array readable");
            for(int r=1;r<=10;r++)for(int c=1;c<=5;c++)
            {
                bool same=Convert.ToString(actual[r,c])==Convert.ToString(expected[r,c]);
                if(!same)Console.WriteLine("Reference mismatch: actual="+Convert.ToString(actual[r,c])+" ("+(actual[r,c]==null?"null":actual[r,c].GetType().Name)+"), expected="+Convert.ToString(expected[r,c])+" ("+(expected[r,c]==null?"null":expected[r,c].GetType().Name)+")");
                Check(same,"Result "+r+","+c);
            }
            sheet.Activate();navigation=New("DirectoryNavigation",app);
            dynamic hidden=Own(sheets.Add(Type.Missing,result,1,Type.Missing));hidden.Name="隐藏'测试";hidden.Visible=2;
            int directoryCount=(int)Call("DirectoryNavigation","Generate",app);
            dynamic directory=Own(sheets.Item("目录")), links=Own(directory.Hyperlinks);
            Check(directoryCount==(int)sheets.Count-1,"Directory lists all original worksheets");
            dynamic directoryCells=Own(directory.Range["A1:C"+((int)sheets.Count)]), directoryFont=Own(directoryCells.Font);
            Check(Convert.ToBoolean(directoryFont.Bold),"Every directory cell, including header and links, is bold");
            Check(Convert.ToInt32(directoryCells.HorizontalAlignment)==-4108,"Every directory cell is horizontally centered");
            Check(Convert.ToInt32(directoryCells.VerticalAlignment)==-4108,"Every directory cell is vertically centered");
            dynamic firstLink=Own(links.Item(1)), firstAnchor=Own(firstLink.Range), linkFont=Own(firstAnchor.Font);
            Check(Convert.ToBoolean(linkFont.Bold)&&Convert.ToInt32(linkFont.Underline)!=-4142,"Hyperlink remains underlined and bold");
            Console.WriteLine("Directory hyperlink color="+Convert.ToString(linkFont.Color));
            firstLink.Follow(false);
            Check((string)Own(a.ActiveSheet).Name=="实验区","Visible worksheet hyperlink still navigates");
            Check(!Convert.ToBoolean(Own(Own(result.Range["A2"]).Font).Bold),"Directory formatting does not change advanced-filter results");
            directory.Activate();
            object[,] names=Own(directory.Range["A1:C"+((int)sheets.Count)]).Value2 as object[,];
            int index=0;for(int r=2;r<=names.GetLength(0);r++)if((string)names[r,1]=="隐藏'测试")index=r-1;
            Check(index>0,"Very-hidden sheet included");
            dynamic link=Own(links.Item(index));link.Follow(false);
            var until=DateTime.UtcNow.AddSeconds(2);while(DateTime.UtcNow<until){Application.DoEvents();System.Threading.Thread.Sleep(30);}
            Check((int)hidden.Visible==-1,"Real SheetFollowHyperlink unhides very-hidden worksheet");
            Check((string)Own(a.ActiveSheet).Name=="隐藏'测试","Hidden apostrophe-name hyperlink navigates");
            ((IDisposable)navigation).Dispose();navigation=null;
            TestLinks(app,test,copy,folder,originalBook,originalSheet,originalSelection);test=null;
            status=0;
        } catch(Exception ex) {Console.WriteLine(ex.ToString());}
        finally {
            if(navigation!=null)((IDisposable)navigation).Dispose();
            if(test!=null&&!originalMode)try{((dynamic)test).Close(false);}catch{}
            if(app!=null&&!originalMode)try{object remaining=InvokeStaticFind(app,copy);if(remaining!=null){try{((dynamic)remaining).Close(false);}finally{Marshal.ReleaseComObject(remaining);}}}catch(Exception ex){Console.WriteLine("Own test copy cleanup: "+ex.Message);}
            if(originalBook!=null)try{((dynamic)originalBook).Activate();((dynamic)originalSheet).Activate();((dynamic)originalSelection).Select();}catch(Exception ex){Console.WriteLine("Restore view: "+ex.Message);}
            if(app!=null){Check(Equals(((dynamic)app).EnableEvents,events),"Original EnableEvents restored");Check(Equals(((dynamic)app).ScreenUpdating,screen),"Original ScreenUpdating restored");}
            if(originalBook!=null)Check((bool)((dynamic)originalBook).Saved==saved,"User workbook Saved unchanged");
            if(app!=null&&originalAddress!=null){object selection=((dynamic)app).Selection;try{Check((string)((dynamic)selection).Address==originalAddress,"User original selection restored");}finally{Marshal.ReleaseComObject(selection);}}
            if(!originalMode)Check(Hash(original)==hash,"User original file SHA256 unchanged");
            for(int i=owned.Count-1;i>=0;i--)try{if(Marshal.IsComObject(owned[i]))Marshal.ReleaseComObject(owned[i]);}catch{}
            if(app!=null)Marshal.ReleaseComObject(app);
            Console.WriteLine("TEST COPY="+copy+"; WPS process left running.");
        }
        return status;
    }
    static Dictionary<string,string> FormulaBindings(string file)
    {
        var result=new Dictionary<string,string>();
        using(var input=File.Open(file,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))using(var zip=new ZipArchive(input,ZipArchiveMode.Read))
        {
            Func<string,XmlDocument> read=name=>{var doc=new XmlDocument{XmlResolver=null};using(var stream=zip.GetEntry(name).Open())doc.Load(stream);return doc;};
            var relations=read("xl/_rels/workbook.xml.rels").DocumentElement.ChildNodes.OfType<XmlElement>().ToDictionary(e=>e.GetAttribute("Id"),e=>e.GetAttribute("Target"));
            var paths=new List<string>();var found=((System.Collections.IEnumerable)Call("XlsxLinkPackage","Scan",file)).Cast<object>().ToList();
            foreach(XmlElement reference in read("xl/workbook.xml").SelectNodes("//*[local-name()='externalReference']"))
            {
                string target=relations[reference.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")];
                string part="xl/externalLinks/_rels/"+Path.GetFileName(target)+".rels";
                object source=found.Single(l=>(string)l.GetType().GetField("Part",Hidden).GetValue(l)==part);paths.Add((string)source.GetType().GetField("SourcePath",Hidden).GetValue(source));
            }
            foreach(var entry in zip.Entries.Where(e=>e.FullName.StartsWith("xl/worksheets/sheet")&&e.FullName.EndsWith(".xml")))
                foreach(XmlElement cell in read(entry.FullName).SelectNodes("//*[local-name()='c'][*[local-name()='f']]"))
                {
                    string formula=cell.SelectSingleNode("*[local-name()='f']").InnerText;if(!Regex.IsMatch(formula,@"\[\d+\]"))continue;
                    result[entry.FullName+"!"+cell.GetAttribute("r")]=Regex.Replace(formula,@"\[(\d+)\]",m=>"["+paths[int.Parse(m.Groups[1].Value)-1]+"]");
                }
        }
        return result;
    }
    static void TestMappedLinks(object app,object target,string file,string original,string evidence,bool originalMode)
    {
        var oldPaths=Enumerable.Range(1,4).Select(i=>Path.Combine(Path.GetDirectoryName(original),"链接","链接"+i+".xlsx")).ToArray();
        var newPaths=Enumerable.Range(5,4).Select(i=>Path.Combine(Path.GetDirectoryName(original),"链接","链接"+i+".xlsx")).ToArray();
        var sourceHashes=oldPaths.Concat(newPaths).ToDictionary(p=>p,Hash);
        Func<HashSet<string>> actual=()=>new HashSet<string>(((System.Collections.IEnumerable)Call("XlsxLinkPackage","Scan",file)).Cast<object>().Select(l=>(string)l.GetType().GetField("SourcePath",Hidden).GetValue(l)),StringComparer.OrdinalIgnoreCase);
        Check(actual().SetEquals(oldPaths),"Four original sources verified before mapped test");
        var before=FormulaBindings(file);Check(before.Count==4,"Four external formula bindings captured");
        dynamic book=target;book.Activate();object savedReport=null;object priorSheet=null;
        try{priorSheet=Own(Own(book.Worksheets).Item("外部链接"));savedReport=Own(((dynamic)priorSheet).Range["A1:D5"]).Value2;}catch{priorSheet=null;}
        object session=null;bool restored=false;
        try
        {
            session=New("LinkAssistantSession",app);Check((int)Invoke(session,"Scan",(object)new[]{file})==4,"Scan and fixed report contain four links");
            dynamic report=Own(Own(book.Worksheets).Item("外部链接"));
            for(int row=2;row<=5;row++){string old=Convert.ToString(Own(report.Range["B"+row]).Value2);int index=Array.FindIndex(oldPaths,p=>string.Equals(p,old,StringComparison.OrdinalIgnoreCase));Check(index>=0,"Report row matches original source "+row);Own(report.Range["C"+row]).Value2=newPaths[index];}
            Check(!(bool)book.Saved,"Four new paths are real unsaved edits");
            session = New("LinkAssistantSession", app);
            bool stop = false; int saved = 0; Exception observationError = null;
            foreach (object step in (System.Collections.IEnumerable)Invoke(session, "Steps", new Func<bool>(() => stop)))
            {
                Console.WriteLine(step);
                if (!step.ToString().StartsWith("已替换：", StringComparison.Ordinal) || ++saved != oldPaths.Length) continue;
                stop = true;
                try
                {
                    Check(actual().SetEquals(newPaths), "All four sources changed to links 5 6 7 8");
                    var after = FormulaBindings(file);
                    Check(before.Count == after.Count && before.All(pair =>
                    {
                        string expected = pair.Value;
                        for (int i = 0; i < 4; i++) expected = expected.Replace(oldPaths[i], newPaths[i]);
                        return after.ContainsKey(pair.Key) && after[pair.Key] == expected;
                    }), "Every formula preserves its intended source mapping");
                    object stillOpen = InvokeStaticFind(app, file);
                    Check(stillOpen != null, "Mapped target stays open throughout replacement");
                    if (stillOpen != null) Marshal.ReleaseComObject(stillOpen);
                }
                catch (Exception error) { observationError = error; }
            }
            restored = actual().SetEquals(oldPaths);
            Check(restored && stop, "Stopping the running batch automatically restores all four original sources");
            var rolledBack = FormulaBindings(file);
            Check(before.Count == rolledBack.Count && before.All(p => rolledBack.ContainsKey(p.Key) && rolledBack[p.Key] == p.Value),
                "Automatic rollback restores each formula binding");
            if (observationError != null) throw observationError;
        }
        finally
        {
            if(restored&&originalMode&&priorSheet!=null){Own(((dynamic)priorSheet).Range["A1:D5"]).Value2=savedReport;book.Save();}
            if(session!=null){object log=session.GetType().GetField("Operations",Hidden).GetValue(session);File.WriteAllText(Path.Combine(evidence,originalMode?"original-mapping-events.log":"copy-mapping-events.log"),(string)log.GetType().GetProperty("Text",Hidden).GetValue(log,null),new UTF8Encoding(false));}
            Check(restored&&actual().SetEquals(oldPaths),"Final saved workbook contains links 1 2 3 4");
            Check(sourceHashes.All(p=>Hash(p.Key)==p.Value),"All eight source files unchanged");
        }
        Check((bool)book.Saved,"Final workbook saved");
    }
    static void TestLinks(object app,object test,string copy,string folder,params object[] preserved)
    {
        dynamic book=test,a=app,books=Own(a.Workbooks);book.Activate();book.Save();
        string replacement=Path.Combine(folder,"本轮新来源-"+Guid.NewGuid().ToString("N")+".xlsx");
        dynamic source=Own(books.Add()),sourceSheet=Own(Own(source.Worksheets).Item(1));Own(sourceSheet.Range["A1"]).Value2=123;source.SaveAs(replacement,51);source.Close(false);
        book.Activate();object session=New("LinkAssistantSession",app);int count=(int)Invoke(session,"Scan",(object)new string[]{copy});Check(count>0,"Workspace copy has external links");
        string reportName=(string)session.GetType().GetField("ReportName",Hidden).GetValue(session);Check(reportName=="外部链接","Fixed report name");
        Check((int)Invoke(session,"Scan",(object)new string[]{copy})==count,"Repeat scan reuses fixed report");
        dynamic report=Own(Own(book.Worksheets).Item(reportName));Own(report.Range["C2"]).Value2=replacement;
        object reopenedSession=New("LinkAssistantSession",app);string result=(string)Invoke(reopenedSession,"Execute",new Action<string>(Console.WriteLine));Console.WriteLine(result);
        Check(result.Contains("成功1条，跳过"+(count-1)+"条，失败0条"),"Reopened assistant rescans and executes via ChangeLink");
        object same=InvokeStaticFind(app,copy);Check(same!=null,"Test target remains open");if(Marshal.IsComObject(same))Marshal.ReleaseComObject(same);
        var links=(System.Collections.IEnumerable)Call("XlsxLinkPackage","Scan",copy);bool found=false;foreach(object l in links)if((string)l.GetType().GetField("SourcePath",Hidden).GetValue(l)==replacement)found=true;Check(found,"Saved replacement points to new source");
        object batch=reopenedSession.GetType().GetField("Batch",Hidden).GetValue(reopenedSession);
        string journal=(string)batch.GetType().GetProperty("Journal",Hidden).GetValue(batch,null);
        Check(WpsShared.SourcePaths.Exists(journal)&&journal.StartsWith(folder+"\\"),"Workspace-local link record stays available");
        Check((int)Invoke(reopenedSession, "Scan", (object)new[] { copy }) == count,
            "saved replacement can be scanned again immediately");
        ((dynamic)test).Close(false);
    }
    static void TestDirectory(object app, object workbook)
    {
        dynamic a=app, book=workbook, sheets=Own(book.Worksheets);
        dynamic last=Own(sheets.Item((int)sheets.Count));
        string token=Guid.NewGuid().ToString("N").Substring(0,8), visibleName="目录测试-"+token, hiddenName="隐藏'测试-"+token;
        dynamic visible=Own(sheets.Add(Type.Missing,last,1,Type.Missing));visible.Name=visibleName;Own(visible.Range["A1"]).Value2="测试副本";
        dynamic hidden=Own(sheets.Add(Type.Missing,visible,1,Type.Missing));hidden.Name=hiddenName;hidden.Visible=2;
        using((IDisposable)New("DirectoryNavigation",app))
        {
            int count=(int)Call("DirectoryNavigation","Generate",app);
            dynamic directory=Own(a.ActiveSheet), cells=Own(directory.Range["A1:C"+(count+1)]), font=Own(cells.Font), links=Own(directory.Hyperlinks);
            Check(count==(int)sheets.Count-1,"Directory lists all original worksheets");
            Check(Convert.ToBoolean(font.Bold),"Every directory cell, including header and links, is bold");
            Check(Convert.ToInt32(cells.HorizontalAlignment)==-4108,"Every directory cell is horizontally centered");
            Check(Convert.ToInt32(cells.VerticalAlignment)==-4108,"Every directory cell is vertically centered");
            object[,] names=cells.Value2 as object[,];int visibleIndex=0,hiddenIndex=0;
            for(int r=2;r<=count+1;r++){if((string)names[r,1]==visibleName)visibleIndex=r-1;if((string)names[r,1]==hiddenName)hiddenIndex=r-1;}
            Check(visibleIndex>0&&hiddenIndex>0,"Visible and very-hidden targets both included");
            dynamic normalLink=Own(links.Item(visibleIndex)), anchor=Own(normalLink.Range), linkFont=Own(anchor.Font);
            Check((string)anchor.Value2=="跳转至本表"&&Convert.ToInt32(linkFont.Underline)!=-4142,"Hyperlink text and underline preserved");
            Console.WriteLine("Directory hyperlink color="+Convert.ToString(linkFont.Color));
            normalLink.Follow(false);
            Check((string)Own(a.ActiveSheet).Name==visibleName,"Visible worksheet hyperlink still navigates");
            directory.Activate();Own(links.Item(hiddenIndex)).Follow(false);
            var until=DateTime.UtcNow.AddSeconds(2);while(DateTime.UtcNow<until){Application.DoEvents();System.Threading.Thread.Sleep(30);}
            Check((int)hidden.Visible==-1,"Real SheetFollowHyperlink unhides very-hidden worksheet");
            Check((string)Own(a.ActiveSheet).Name==hiddenName,"Hidden apostrophe-name hyperlink still navigates");
            // The generic style used by reports and advanced-filter results must
            // keep regular body text; only the generated directory is centered.
            Call("WorkbookUi","StyleTable",(object)visible,2,2);
            dynamic plain=Own(visible.Range["A2:B2"]), plainFont=Own(plain.Font);
            Check(!Convert.ToBoolean(plainFont.Bold),"Shared report body style remains regular");
            Check(Convert.ToInt32(plain.HorizontalAlignment)!=-4108,"Shared report body alignment remains unchanged");
        }
    }
    static object InvokeStaticFind(object app,string copy){return dll.GetType("Excel_Iterate.LinkAssistantSession",true).GetMethod("FindBook",Hidden).Invoke(null,new object[]{app,copy});}
}
