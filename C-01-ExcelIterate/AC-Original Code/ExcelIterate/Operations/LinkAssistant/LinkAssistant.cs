using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class HostLinkSource
    {
        internal string Name, Path;
    }
    // All host calls run on the Ribbon/WinForms STA. No RCW is retained for a
    // workbook that may close; report identity survives reopening via a hidden name.
    internal sealed partial class LinkAssistantSession
    {
        internal static readonly string[] Headers = { "文档路径", "外部链接", "新链接（待填）", "处理状态" };
        internal readonly object App;
        internal readonly LinkOperationLog Operations = new LinkOperationLog();
        internal Action<string, IEnumerable<string>> StageValidation; // Offline probe seam; null always uses the real host.
        internal readonly string ReportBookPath;
        internal string ReportName, ReportAnchor;
        internal List<ExternalLinkInfo> Links = new List<ExternalLinkInfo>();
        internal LinkBatch Batch; internal string ScanFolder;
        internal bool CanExecute { get { return ReportAnchor != null && (Batch == null || Batch.State == "Scanned"); } }
        private readonly Dictionary<string,int> diagnosticOffsets=new Dictionary<string,int>(StringComparer.OrdinalIgnoreCase);
        private int diagnosticStart;
        internal Action<string,string> DiagnosticWriter;
        private Exception reportedError;
        private string displayedOwner;
        private bool reportChanged, reportSaved, deferDiagnostics;
        private void FlushDiagnostics()
        {
            if (Batch == null || deferDiagnostics || Operations.PersistenceError != null) return;
            string text=Operations.DiagnosticText;
            try {foreach(string dir in Batch.LogDirectories){string path=Batch.LogPath(dir);int offset;if(!diagnosticOffsets.TryGetValue(path,out offset))offset=diagnosticStart;if(offset>=text.Length)continue;
                if(DiagnosticWriter!=null)DiagnosticWriter(path,text.Substring(offset));else Batch.LogTo(dir,text.Substring(offset));diagnosticOffsets[path]=text.Length;}}
            catch(Exception ex){LogFailure=ex.Message;Operations.ReportPersistenceFailure(ex);}
        }
        private void Owner(string path)
        {if(!string.Equals(displayedOwner,path,StringComparison.OrdinalIgnoreCase)){Operations.Record("","工作簿：\r\n"+path);displayedOwner=path;}}
        private static string Reason(Exception error)
        {
            while(error.InnerException!=null)error=error.InnerException;
            if(error is System.Runtime.InteropServices.COMException)return "宿主链接接口或保存调用失败，请检查WPS支持情况、工作簿保护及文件权限。";
            return error.Message.Split(new[]{"\r\n","\n"},StringSplitOptions.None)[0];
        }
        private void Pair(LinkChange c,bool reverse,bool success,Exception error)
        {
            Owner(c.Owner);
            if(c.Action==LinkAction.Break){Operations.Record("",(success?(reverse?"断链已还原：":"已断开："):(reverse?"断链还原失败：":"断链失败："))+c.Old+(error==null?"":"\r\n原因："+Reason(error)));return;}
            Operations.Record("",(success?(reverse?"已还原：":"成功："):(reverse?"还原失败：":"失败："))+(reverse?c.New:c.Old)+"\r\n→ "+(reverse?c.Old:c.New)+(error==null?"":"\r\n原因："+Reason(error)));
        }
        internal void ReportStopped(Exception error)
        {
            if(ReferenceEquals(reportedError,error))return;reportedError=error;
            Operations.Trace("操作异常",error.ToString());
            Operations.Record("操作停止", Reason(error) + "\r\n尚未开始本次链接处理。");
        }
        internal LinkAssistantSession(object app)
        {
            App = app;
            using (ComBag bag = new ComBag())
            {
                object book = bag.Own(((dynamic)app).ActiveWorkbook);
                if (book == null) throw new InvalidOperationException("请先打开并保存用于接收链接清单的 xlsx 工作簿。");
                Connect.PrepareWorkbookSave(Convert.ToString(((dynamic)book).FullName));
                ReportBookPath = WorkbookSaving.Save(book, true);
            }
            Batch=LinkBatch.Load(ReportBookPath);
            if(Batch!=null){ReportAnchor=Batch.Anchor;ReportName=Batch.ReportName;Links=Batch.Rows.Select(r=>r.Link()).ToList();ScanFolder=Batch.ScopeFolder;}
            Operations.DiagnosticLineAdded+=delegate {FlushDiagnostics();};
            if (Batch != null)
            {
                LastResult = Batch.LastSummary;
                if (!string.IsNullOrEmpty(Batch.UserLog)) Operations.RestoreDisplay(Batch.UserLog);
                else if (!string.IsNullOrEmpty(LastResult)) Operations.Record("", LastResult);
                if (Batch.State == "Active")
                {
                    Operations.Record("上次处理已中断", "重新扫描前请核对工作簿和链接记录：" + Batch.Journal);
                    RecordUnresolvedCopies();
                }
            }
            Operations.Trace("打开助手", "清单接收工作簿：" + ReportBookPath);
        }
        internal string LogFailure;
        internal static object FindBook(object app, string path)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    object book = books.Item(i); bool keep = false;
                    try { keep = string.Equals(Path.GetFullPath((string)((dynamic)book).FullName), path, StringComparison.OrdinalIgnoreCase); if (keep) return book; }
                    finally { if (!keep) using (ComBag one = new ComBag()) one.Own(book); }
                }
            }
            return null;
        }
        private static void RequireSaved(object book)
        {
            Connect.PrepareWorkbookSave(Convert.ToString(((dynamic)book).FullName));
            WorkbookSaving.Save(book, true);
        }
        internal int Scan(IEnumerable<string> files)
        {
            reportChanged = reportSaved = false;
            deferDiagnostics = Batch == null || Batch.State != "Scanned";
            if (deferDiagnostics) diagnosticStart = Operations.DiagnosticText.Length;
            try { return ScanCore(files); }
            catch (Exception ex)
            {
                reportedError = ex;
                if (reportChanged) ReportAnchor = null;
                Operations.Trace("扫描异常", ex.ToString());
                Operations.Record("扫描中止", Reason(ex) + "\r\n" +
                    (reportSaved ? "清单已保存，链接记录未完整建立；请重新扫描。" :
                        reportChanged ? "清单已更新，但未完整保存；请重新扫描。" : "本次扫描未完成，尚未更新清单。"));
                throw;
            }
            finally { deferDiagnostics = false; }
        }
        private int ScanCore(IEnumerable<string> files)
        {
            var previousBatch = Batch;
            if (previousBatch != null) previousBatch.PreserveUnresolvedRecord();
            List<ExternalLinkInfo> collected = new List<ExternalLinkInfo>();
            string[] targets = files.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            Operations.ClearDisplay();LogFailure=null;displayedOwner=null;reportedError=null;
            Operations.Record("开始扫描", targets.Length + "个工作簿。");
            if (targets.Length == 0) { Operations.Record("", "没有可扫描的工作簿。");LastResult="没有可扫描的工作簿。"; return 0; }
            using (ComBag receiver = new ComBag()) RequireSaved(receiver.Own(FindBook(App, ReportBookPath)));
            foreach (string item in targets)
            {
                Owner(item);
                using (ComBag one = new ComBag()) { object open = one.Own(FindBook(App, item)); if (open != null) RequireSaved(open); }
                var found = XlsxLinkPackage.ScanReport(item);
                LinkNameAnalysis.Describe(item, found);
                int previous = collected.Count;
                foreach (ExternalLinkInfo link in found)
                    if (!collected.Any(l => string.Equals(l.Key, link.Key, StringComparison.OrdinalIgnoreCase))) collected.Add(link);
                int added = collected.Count - previous;
                Operations.Record("",added==0?"未发现可替换链接。":"发现"+added+"条链接：\r\n"+string.Join("\r\n",collected.Skip(previous).Select((l,i)=>(i+1)+". "+(l.SourcePath??l.RawTarget))));
            }
            using (ComBag bag = new ComBag())
            {
                dynamic book = bag.Own(FindBook(App, ReportBookPath));
                if (book == null) throw new InvalidOperationException("接收清单的原工作簿已关闭，请重新打开链接助手。");
                RequireSaved((object)book);
                using (HostQuiet quiet = new HostQuiet(App))
                {
                    object sheet = null, anchor = null; bool success = false, created = false;
                    try
                    {
                        dynamic sheets=bag.Own(book.Worksheets);
                        try {sheet=bag.Own(sheets.Item("外部链接"));}catch {sheet=null;}
                        reportChanged=true;
                        if(sheet==null){sheet=WorkbookUi.NewSheet((object)book,"外部链接",bag,false);created=true;}
                        else {dynamic cells=bag.Own(((dynamic)sheet).Cells);cells.Clear();}
                        // Hidden names prove ownership of historic numbered reports.
                        dynamic oldNames=bag.Own(book.Names);
                        for(int i=Convert.ToInt32(oldNames.Count);i>=1;i--)
                        {
                            dynamic old=bag.Own(oldNames.Item(i));string label=Convert.ToString(old.Name);
                            if(!label.Contains("_EI_LinkReport_"))continue;
                            try {dynamic oldRange=bag.Own(old.RefersToRange),oldSheet=bag.Own(oldRange.Worksheet);string oldTitle=Convert.ToString(oldSheet.Name);
                                if(oldTitle!="外部链接"&&System.Text.RegularExpressions.Regex.IsMatch(oldTitle,@"^外部链接[（(]\d+[）)]$"))WorkbookUi.DeleteSheet(App,(object)oldSheet);
                            }catch { /* Unowned or broken references are not deleted as sheets. */ }
                            old.Delete();
                        }
                        string token = "_EI_LinkReport_" + Guid.NewGuid().ToString("N");
                        dynamic names = bag.Own(book.Names);
                        anchor = bag.Own(names.Add(token, "=" + WorkbookUi.Quote((string)((dynamic)sheet).Name) + "!$A$1", false));
                        object[,] data = new object[collected.Count + 1, 4];
                        for (int c = 0; c < 4; c++) data[0, c] = Headers[c];
                        for (int r = 0; r < collected.Count; r++) { data[r + 1, 0] = collected[r].OwnerPath; data[r + 1, 1] = collected[r].SourcePath ?? collected[r].RawTarget; data[r + 1, 2] = ""; data[r + 1, 3] = collected[r].Status; }
                        dynamic range = bag.Own(((dynamic)sheet).Range["A1:D" + (collected.Count + 1)]); range.NumberFormat = "@"; range.Value2 = data;
                        WorkbookUi.StyleTable(sheet, 4, collected.Count + 1);
                        dynamic edits = bag.Own(((dynamic)sheet).Range["C2:C" + Math.Max(2, collected.Count + 1)]), fill = bag.Own(edits.Interior); fill.Color = 0xDCF4FF;
                        WorkbookUi.FitOutput(sheet, 1, 4, collected.Count + 1);
                        ((dynamic)sheet).Activate();
                        Links = collected; ReportName = (string)((dynamic)sheet).Name; ReportAnchor = token; success = true;
                    }
                    finally { if (!success) { if (anchor != null) ((dynamic)anchor).Delete(); if (sheet != null && created) WorkbookUi.DeleteSheet(App, sheet); } }
                }
                RequireSaved((object)book);
                reportSaved=true;
            }
            var scanned = new LinkBatch { ReportBook = ReportBookPath, Anchor = ReportAnchor, ReportName = ReportName, ScopeFolder = ScanFolder,
                Owners = targets.ToList(), Rows = Links.Select(LinkRowRecord.FromLink).ToList() };
            scanned.LogDirectories = ScanFolder != null ? new List<string> { ScanFolder } : targets.Select(Path.GetDirectoryName).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (previousBatch != null && previousBatch.State == "Scanned")
            {
                scanned.Id = previousBatch.Id;
                scanned.LogDirectories = scanned.LogDirectories.Concat(previousBatch.LogDirectories).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            }
            scanned.Persist();
            Batch = scanned;
            deferDiagnostics = false;
            FlushDiagnostics();
            Operations.Trace("导出清单", ReportBookPath + "；工作表「" + ReportName + "」");
            LastResult=Links.Any(l=>l.UnavailableReason==null)?"扫描完成。在「外部链接」第三列填完整新路径，或对普通链接填「断开链接」，再点击「执行替换」。":"扫描完成，未发现可处理链接。";
            SaveSummary();
            return Links.Count;
        }
        private object ReportSheet(object book, ComBag bag)
        {
            if (ReportAnchor == null) throw new InvalidOperationException("请先扫描链接。");
            dynamic names = bag.Own(((dynamic)book).Names), name = bag.Own(names.Item(ReportAnchor)), range = bag.Own(name.RefersToRange);
            return bag.Own(range.Worksheet);
        }
        internal const string MissingStatus="链接已替换，目标文件当前缺失";
        internal const string BreakCommand="断开链接";
        internal static bool MissingTarget(string path) {
            try { if((File.GetAttributes(path)&FileAttributes.Directory)!=0)throw new IOException("新链接指向文件夹，请填写 xlsx 文件路径。");return false; }
            catch(FileNotFoundException){return true;}catch(DirectoryNotFoundException){return true;}
        }
        private static bool ValidateReplacementTarget(string path)
        {
            if (MissingTarget(path)) return true;
            try { XlsxLinkPackage.Scan(path); }
            catch (FileNotFoundException) { return true; }
            catch (DirectoryNotFoundException) { return true; }
            return false;
        }
        private static void RequireAvailableReplacementTarget(string path)
        {
            if (MissingTarget(path)) throw new IOException("目标文件于替换预检后变为缺失。请重新扫描并执行换链：" + path);
        }
        private static string ReplacementPath(string input)
        {
            string path = input.Replace('/', '\\');
            int start = path.LastIndexOf('\\') + 1;
            if (start > 0 && path.Length > start + 2 && path[start] == '[' && path[path.Length - 1] == ']')
                path = path.Substring(0, start) + path.Substring(start + 1, path.Length - start - 2);
            if (!XlsxLinkPackage.IsFullWindowsPath(path)) throw new InvalidOperationException("新链接请填写完整文件路径：" + input);
            path = Path.GetFullPath(path);
            if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("新链接需要填写 xlsx 完整路径：" + input);
            return path;
        }
        internal Dictionary<string, string> ReadMappings()
        {
            using (ComBag bag = new ComBag())
            {
                object book = bag.Own(FindBook(App, ReportBookPath));
                if (book == null) throw new InvalidOperationException("清单工作簿已关闭，请重新打开后再执行。");
                RequireSaved(book); object sheet = ReportSheet(book, bag);
                dynamic used = bag.Own(((dynamic)sheet).UsedRange), rows = bag.Own(used.Rows);
                int count = Convert.ToInt32(rows.Count);
                if (count != Links.Count + 1) throw new InvalidOperationException("清单行数已变化，请重新扫描；不要增删清单行。");
                dynamic range = bag.Own(((dynamic)sheet).Range["A1:D" + count]); Array values = range.Value2 as Array;
                if (values == null) throw new InvalidOperationException("无法可靠读取链接清单。");
                int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
                for (int c = 0; c < 4; c++) if (Convert.ToString(values.GetValue(r0, c0 + c)) != Headers[c]) throw new InvalidOperationException("清单表头已变化，请重新扫描。");
                Dictionary<string, ExternalLinkInfo> known = Links.ToDictionary(l => l.Key, StringComparer.OrdinalIgnoreCase);
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                for (int r = 1; r < count; r++)
                {
                    string key = Convert.ToString(values.GetValue(r0 + r, c0)) + "\n" + Convert.ToString(values.GetValue(r0 + r, c0 + 1));
                    ExternalLinkInfo link;
                    if (!known.TryGetValue(key, out link) || !seen.Add(key)) throw new InvalidOperationException("文档路径或外部链接已变化，或有重复行；请重新扫描。");
                    string input = Convert.ToString(values.GetValue(r0 + r, c0 + 2)).Trim();
                    var source = link.ValidatedSource ?? link;
                    link.NewPath = null;
                    link.DestinationMissing = false;
                    link.BreakRequested=false;
                    if (input.Length == 0) { link.Status = link.UnavailableReason ?? "未填新链接，跳过"; Operations.Trace("跳过链接", link.Key + "；" + link.Status); continue; }
                    try {
                    if (link.UnavailableReason != null) throw new InvalidOperationException("不能替换非文件外部链接：" + link.RawTarget);
                    if ((link.IsMissingPath || source.IsMissingPath) && input == BreakCommand) throw new InvalidOperationException("来源路径待核对，请填写完整新路径替换：" + link.RawTarget);
                    if(input==BreakCommand){link.BreakRequested=true;link.Status="待断开，将处理相关公式及未使用的名称引用";Operations.Trace("待断开链接",link.Key);continue;}
                    if(input=="请断开该链接")throw new InvalidOperationException("断链指令已改为「断开链接」，请修改清单后重试。");
                    string path = ReplacementPath(input);
                    if (string.Equals(path, link.OwnerPath, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("新链接不能指向目标工作簿自身。");
                    link.DestinationMissing = ValidateReplacementTarget(path);
                    string currentPath = source.SourcePath ?? (link.DestinationMissing ? XlsxLinkPackage.AbsoluteFileTarget(source.RawTarget) : null);
                    if (string.Equals(path, currentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        link.Status = link.DestinationMissing ? "链接路径已一致，目标文件当前缺失" : "链接相同，无需替换";
                        Operations.Trace("路径已核对", link.Key + "；" + link.Status);
                        continue;
                    }
                    if (link.DestinationMissing) link.Status = "待替换，目标文件当前缺失";
                    link.NewPath = path; result.Add(key, path);
                    Operations.Trace("待替换链接", link.Key + "；新链接：" + path);
                    } catch(Exception ex){reportedError=ex;Operations.Trace("映射预检异常",ex.ToString());Pair(new LinkChange{Owner=link.OwnerPath,Old=link.SourceId,New=input},false,false,ex);Operations.Record("","预检未通过，尚未修改链接。");throw;}
                }
                return result;
            }
        }
        internal string Execute(Action<string> progress)
        { foreach(string step in Steps(() => false))if(progress!=null)progress(step);return LastResult; }
        internal string LastResult;
        private static IEnumerable<string> HostLinkNames(object raw)
        {
            if(raw==null||raw==DBNull.Value)return Enumerable.Empty<string>();
            Array array=raw as Array;
            if(array!=null&&array.Rank!=1)throw new IOException("宿主链接来源不是一维数组，无法可靠匹配。");
            if(array==null&&!(raw is string))throw new IOException("无法识别宿主链接来源类型："+raw.GetType().FullName);
            IEnumerable<object> values=array==null?new object[]{raw}:array.Cast<object>();
            return values.Select(value => Convert.ToString(value)).Where(name => !string.IsNullOrWhiteSpace(name));
        }
        internal static List<HostLinkSource> ParseHostSources(object raw,string owner,IEnumerable<ExternalLinkInfo> package)
        {
            var known=package.Where(l=>l.SourcePath!=null).ToList();var result=new List<HostLinkSource>();
            foreach(string name in HostLinkNames(raw))
            {
                string resolved=null;
                // Host names are Windows filenames, not OOXML URI references.
                // Preserve literal %20/%23 in filenames and the exact call token.
                if(System.Text.RegularExpressions.Regex.IsMatch(name,@"^(?:[A-Za-z]:[\\/]|\\\\)"))resolved=System.IO.Path.GetFullPath(name.Replace('/', '\\'));
                else if(name.StartsWith("file:",StringComparison.OrdinalIgnoreCase))resolved=XlsxLinkPackage.Resolve(owner,name);
                else if(!name.Contains("://"))
                {
                    resolved=name.StartsWith("\\",StringComparison.Ordinal)?System.IO.Path.GetFullPath(System.IO.Path.Combine(System.IO.Path.GetPathRoot(owner),name.TrimStart('\\'))):System.IO.Path.GetFullPath(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(owner),name));
                    if(!known.Any(l=>string.Equals(l.SourcePath,resolved,StringComparison.OrdinalIgnoreCase)))
                    {
                        var exact=known.Where(l=>string.Equals(l.RawTarget,name,StringComparison.Ordinal)).Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                        if(exact.Count>1)throw new IOException("宿主原始来源匹配不唯一："+name);
                        if(exact.Count==1)resolved=exact[0];
                        else if(name.IndexOfAny(new[]{'\\','/'})<0)
                        {
                            var leaf=known.Where(l=>string.Equals(System.IO.Path.GetFileName(l.SourcePath),name,StringComparison.OrdinalIgnoreCase)).Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                            if(leaf.Count>1)throw new IOException("宿主仅返回文件名，来源匹配不唯一："+name);
                            if(leaf.Count==1)resolved=leaf[0];
                        }
                    }
                }
                result.Add(new HostLinkSource{Name=name,Path=resolved});
            }
            return result;
        }
        private static List<HostLinkSource> ActiveHostSources(object book,string owner,Action<string> trace)
        {
            var package=XlsxLinkPackage.Scan(owner).Where(l=>l.SourcePath!=null).ToList();
            object raw=((dynamic)book).LinkSources(1);
            var host=ParseHostSources(raw,owner,package);var active=new List<HostLinkSource>();
            if(trace!=null)trace("返回类型："+(raw==null?"null":raw.GetType().FullName)+"；原始来源：\r\n"+string.Join("\r\n",host.Select(s=>s.Name)));
            foreach(string path in package.Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var matches=host.Where(s=>string.Equals(s.Path,path,StringComparison.OrdinalIgnoreCase)).GroupBy(s=>s.Name,StringComparer.Ordinal).Select(g=>g.First()).ToList();
                if(matches.Count!=1)throw new IOException("宿主来源与已保存工作簿无法唯一对应："+path+"；匹配 "+matches.Count+" 项。");
                active.Add(matches[0]);
            }
            // WPS 12.1 retains replaced names in its in-memory LinkSources list,
            // even after Save. Saved package relationships identify active links.
            if(trace!=null&&host.Count>active.Count)trace("宿主保留 "+(host.Count-active.Count)+" 条历史来源；按已保存XLSX关系核对实际链接。");
            return active;
        }
        private static List<string> Sources(object book,string path)
        {return ActiveHostSources(book,path,null).Select(s=>s.Path).ToList();}
        private static string Failure(Exception error)
        {return error.GetType().Name+" (HRESULT:0x"+error.HResult.ToString("X8")+")："+error.Message;}
        private List<ExternalLinkInfo> Fresh(string[] owners)
        {
            var rows=new List<ExternalLinkInfo>();
            foreach(string path in owners)
            {using(var bag=new ComBag()){object book=bag.Own(FindBook(App,path));if(book!=null)RequireSaved(book);}var found=XlsxLinkPackage.ScanReport(path);LinkNameAnalysis.Describe(path,found);foreach(var link in found)if(!rows.Any(r=>string.Equals(r.Key,link.Key,StringComparison.OrdinalIgnoreCase)))rows.Add(link);}
            return rows;
        }
        private static bool SameReportSource(ExternalLinkInfo previous, ExternalLinkInfo current)
        {
            if (!string.Equals(previous.OwnerPath, current.OwnerPath, StringComparison.OrdinalIgnoreCase)) return false;
            if (previous.IsMissingPath == current.IsMissingPath && string.Equals(previous.SourceId, current.SourceId, StringComparison.OrdinalIgnoreCase)) return true;
            if (!previous.IsMissingPath && !current.IsMissingPath) return false;
            string previousPath = previous.SourcePath ?? XlsxLinkPackage.AbsoluteFileTarget(previous.RawTarget);
            string currentPath = current.SourcePath ?? XlsxLinkPackage.AbsoluteFileTarget(current.RawTarget);
            if (previousPath != null && currentPath != null) return string.Equals(previousPath, currentPath, StringComparison.OrdinalIgnoreCase);
            // Saving the report can turn an unqualified missing target into an
            // ordinary relative target. Only accept the same recorded filename.
            string filename = Uri.UnescapeDataString(previous.RawTarget ?? string.Empty);
            return filename.Length > 0 && filename.IndexOfAny(Path.GetInvalidFileNameChars()) < 0
                && string.Equals(filename, Uri.UnescapeDataString(current.RawTarget ?? string.Empty), StringComparison.OrdinalIgnoreCase);
        }
        private void Rescan()
        {
            if (Batch == null) return;
            var fresh = Fresh(Batch.Candidates());
            const string changed = "当前链接与原扫描清单不一致，暂停执行；请重新扫描。第三列输入未被覆盖。";
            if (fresh.Count != Links.Count) throw new InvalidOperationException(changed);
            var used = new HashSet<ExternalLinkInfo>();
            foreach (var link in Links)
            {
                var matches = fresh.Where(current => SameReportSource(link, current)).ToList();
                if (matches.Count != 1 || !used.Add(matches[0])) throw new InvalidOperationException(changed);
                link.ValidatedSource = matches[0];
            }
        }
        private static LinkChange ToChange(ExternalLinkInfo link)
        {
            var source = link.ValidatedSource ?? link;
            bool reclassified = source.IsMissingPath != link.IsMissingPath
                || !string.Equals(source.SourceId, link.SourceId, StringComparison.OrdinalIgnoreCase);
            return new LinkChange
            {
                Owner = link.OwnerPath,
                Old = source.SourceId,
                ReportSource = reclassified ? link.SourceId : null,
                New = link.NewPath,
                DestinationMissing = link.DestinationMissing,
                IsMissingPath = source.IsMissingPath,
                Action = link.BreakRequested ? LinkAction.Break : LinkAction.Replace
            };
        }
        private void ClosedChange(string owner,List<LinkChange> changes,bool reverse)
        {
            string stage=null;string hash=XlsxLinkPackage.Hash(owner);
            var mapping=changes.ToDictionary(c=>reverse?c.New:c.Old,c=>reverse?c.Old:c.New,StringComparer.OrdinalIgnoreCase);
            if((File.GetAttributes(owner)&FileAttributes.ReadOnly)!=0)throw new IOException("工作簿为只读："+owner);
            try
            {
                stage=XlsxLinkPackage.Prepare(owner,mapping);XlsxLinkPackage.VerifyRewrite(owner,stage,mapping);
                if(StageValidation!=null)StageValidation(stage,XlsxLinkPackage.Scan(stage).Select(l=>l.SourcePath).Where(p=>p!=null));
                string prepared=XlsxLinkPackage.Hash(stage);
                if(XlsxLinkPackage.Hash(owner)!=hash)throw new IOException("验证期间目标文件已变化，未提交："+owner);
                foreach(var c in changes){c.Intent=true;c.IntentReverse=reverse;}Batch.Persist();
                File.Replace(stage,owner,null);stage=null;
                if(XlsxLinkPackage.Hash(owner)!=prepared)throw new IOException("已提交，但提交后摘要核对失败。");
                foreach(var c in changes){c.EverApplied=true;c.Applied=!reverse;c.Intent=false;}Batch.Persist();
            }
            finally {if(stage!=null&&File.Exists(stage))File.Delete(stage);}
        }
        private void ReconcileIntent(string owner)
        {
            var pending=Batch.Changes.Where(c=>c.Owner==owner&&c.Intent).ToList();if(pending.Count==0)return;
            List<string> sources;
            using(var bag=new ComBag()){object book=bag.Own(FindBook(App,owner));if(book!=null)RequireSaved(book);sources=book==null?XlsxLinkPackage.Scan(owner).Where(l=>l.SourcePath!=null).Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList():Sources(book,owner);}
            Func<bool,HashSet<string>> expected=completed=>new HashSet<string>(Batch.Rows.Where(r=>r.Owner==owner&&r.Source!=null).Select(r=>{
                var c=Batch.Changes.FirstOrDefault(x=>x.Owner==owner&&string.Equals(x.Old,r.Source,StringComparison.OrdinalIgnoreCase));
                bool applied=c!=null&&(c.Intent?(completed?!c.IntentReverse:c.IntentReverse):c.Applied);return applied?c.New:r.Source;
            }),StringComparer.OrdinalIgnoreCase);
            bool done=expected(true).SetEquals(sources),untouched=expected(false).SetEquals(sources);
            if(done==untouched)throw new IOException("链接中断状态不明确，已保留链接记录："+owner);
            foreach(var c in pending){c.Applied=done?!c.IntentReverse:c.IntentReverse;c.EverApplied|=c.Applied||c.IntentReverse;c.Intent=false;}
            Batch.Persist();
        }
        private sealed class ProcessingResult
        {
            internal readonly HashSet<LinkChange> Failures = new HashSet<LinkChange>();
            internal int Success;
            internal bool Cancelled;
            internal Exception Error;
        }

        internal IEnumerable<string> Steps(Func<bool> cancelled)
        {
            if (!CanExecute) throw new InvalidOperationException("请先扫描链接。");
            displayedOwner = null;
            reportedError = null;
            try { Rescan(); ReadMappings(); } catch (Exception ex) { ReportStopped(ex); throw; }

            if (Batch == null) // Offline seam; production scans always have a persisted identity.
                Batch = new LinkBatch { ReportBook = ReportBookPath, Anchor = ReportAnchor, ReportName = ReportName,
                    Owners = Links.Select(l => l.OwnerPath).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                    Rows = Links.Select(LinkRowRecord.FromLink).ToList(), LogDirectories = new List<string> { Path.GetDirectoryName(ReportBookPath) } };
            foreach (var row in Batch.Rows)
            {
                var link = Links.FirstOrDefault(l => l.Key == row.Owner + "\n" + (row.Source ?? row.Raw));
                if (link != null) row.Status = link.Status;
            }
            Batch.Changes = Links.Where(l => l.NewPath != null || l.BreakRequested).Select(ToChange).ToList();
            Batch.Skipped = Links.Count - Batch.Changes.Count;
            Batch.State = "Active";
            try { Batch.Persist(); }
            catch { Batch.State = "Scanned"; Batch.Changes.Clear(); throw; }
            FlushDiagnostics();
            bool includesBreak = Batch.Changes.Any(c => c.Action == LinkAction.Break);
            Operations.Record(includesBreak ? "开始链接处理" : "开始替换", Batch.Changes.Count + "条链接。");

            var forward = new ProcessingResult();
            foreach (string step in RunChanges(false, cancelled, forward)) yield return step;
            bool rollback = forward.Cancelled || forward.Failures.Count > 0 || forward.Error != null;
            if (rollback)
            {
                if (forward.Error != null)
                {
                    Operations.Trace("操作异常", forward.Error.ToString());
                    Operations.Record("链接处理异常", Reason(forward.Error));
                }
                Operations.Record("", forward.Cancelled ? "已停止后续处理，正在尝试回退本批改动。" : "链接处理失败，正在尝试回退本批改动。");
                var reverse = new ProcessingResult();
                foreach (string step in RunChanges(true, () => false, reverse)) yield return step;
                if (reverse.Error != null)
                {
                    Operations.Trace("回退异常", reverse.Error.ToString());
                    Operations.Record("自动回退中止", Reason(reverse.Error));
                }
                int remaining = PendingChanges();
                Batch.State = remaining > 0 ? "UndoFailed" : "Undone";
                int unprocessed = Batch.Changes.Count - forward.Success - forward.Failures.Count;
                LastResult = "链接处理停止：仍有" + remaining + "条改动或待核对步骤，跳过" + Batch.Skipped +
                    "条，失败" + forward.Failures.Count + "条，未执行" + unprocessed + "条。\r\n" +
                    (remaining > 0 ? "自动回退未完成，请按操作日志核对实际结果。" : "本批链接已恢复原状。");
            }
            else
            {
                Batch.State = "Completed";
                LastResult = (includesBreak ? "链接处理完成（替换/断开）：成功" : "替换完成：成功") + forward.Success +
                    "条，跳过" + Batch.Skipped + "条，失败0条。\r\n" +
                    (forward.Success > 0 ? "结果已自动保存。可扫描下一批。" : "没有修改链接。可扫描下一批。");
                int missing = Batch.Changes.Count(c => c.Applied && c.DestinationMissing);
                if (missing > 0) LastResult += "\r\n已替换的链接中，目标文件当前缺失" + missing + "条。";
            }

            RefreshStatuses(rollback);
            try { UpdateReport(rollback); }
            catch (Exception ex)
            {
                Operations.Trace("清单回写异常", ex.ToString());
                Operations.Record("清单状态未完整保存", Reason(ex) + "\r\n链接处理结果以本日志为准。");
            }
            ClearTemporaryProtection(rollback);
            SaveSummary();
        }

        private IEnumerable<string> RunChanges(bool rollback, Func<bool> cancelled, ProcessingResult outcome)
        {
            using (var steps = ChangeWorkbooks(rollback, cancelled, outcome).GetEnumerator())
            {
                while (true)
                {
                    bool more;
                    try { more = steps.MoveNext(); }
                    catch (Exception error) { outcome.Error = error; more = false; }
                    if (!more) yield break;
                    yield return steps.Current;
                }
            }
        }

        private IEnumerable<string> ChangeWorkbooks(bool undo, Func<bool> cancelled, ProcessingResult outcome)
        {
            var failures = outcome.Failures;
            bool abort = false;
            foreach(string owner in Batch.Changes.Select(c=>c.Owner).Distinct(StringComparer.OrdinalIgnoreCase).ToArray())
            {
                if(!undo&&cancelled()){abort=true;break;}
                var ownerChanges = Batch.Changes.Where(c => c.Owner == owner).ToList();
                if (!undo)
                {
                    try
                    {
                        foreach (var replacement in ownerChanges.Where(c => c.Action == LinkAction.Replace))
                            replacement.DestinationMissing = MissingTarget(replacement.New);
                    }
                    catch (Exception preflightError)
                    {
                        foreach (var change in ownerChanges)
                        {
                            change.FailureReason = Reason(preflightError); failures.Add(change); Pair(change, false, false, preflightError);
                        }
                        continue;
                    }
                }
                if (Checkpoint(owner) != null || ownerChanges.Any(c => c.Action == LinkAction.Break || c.IsMissingPath || c.ReportSource != null || c.DestinationMissing))
                {
                    var checkpointResult = new CheckpointOutcome();
                    foreach (string step in CheckpointSteps(owner, undo, cancelled, checkpointResult)) yield return step;
                    outcome.Success += checkpointResult.Success;
                    abort |= checkpointResult.Cancelled;
                    if (checkpointResult.Error != null)
                        foreach (var change in ownerChanges.Where(c => undo || !c.Applied || c.Intent))
                        {
                            failures.Add(change);
                            if (!undo) change.FailureReason = Reason(checkpointResult.Error);
                            Pair(change, undo, false, checkpointResult.Error);
                        }
                    if (!undo && checkpointResult.Success > 0)
                    {
                        yield return "已保存工作簿：" + Path.GetFileName(owner);
                        abort |= cancelled();
                    }
                    if (abort) break;
                    continue;
                }
                Exception error=null;List<LinkChange> ordered=null;
                try
                {
                    ReconcileIntent(owner);
                    var changes=Batch.Changes.Where(c=>c.Owner==owner&&(undo?c.Applied:!c.Applied)).ToList();
                    if(changes.Count==0)continue;
                    using(var bag=new ComBag())
                    {object book=bag.Own(FindBook(App,owner));if(book!=null)RequireSaved(book);
                        var actual=book==null?XlsxLinkPackage.Scan(owner).Where(l=>l.SourcePath!=null).Select(l=>l.SourcePath).Distinct(StringComparer.OrdinalIgnoreCase).ToList():Sources(book,owner);
                        foreach(var c in changes)if(!actual.Contains(undo?c.New:c.Old,StringComparer.OrdinalIgnoreCase))throw new IOException("链接来源已变化，暂停该工作簿："+owner);
                        ordered=LinkBatch.Order(changes,actual,undo);
                    }
                }catch(Exception ex){error=ex;}
                if(error!=null){Operations.Trace("工作簿预检异常",owner+"；"+error);foreach(var c in Batch.Changes.Where(c=>c.Owner==owner&&(undo?c.Applied||c.Intent:!c.Applied))){failures.Add(c);if(!undo)c.FailureReason=Reason(error);Pair(c,undo,false,error);}Operations.Record("",undo?"该工作簿未还原，恢复记录已保留。":"该工作簿预检未通过，尚未修改链接。");continue;}
                yield return (undo?"还原：":"替换：")+Path.GetFileName(owner);
                bool live;
                using(var bag=new ComBag())live=bag.Own(FindBook(App,owner))!=null;
                if(live)
                {
                    foreach(var c in ordered)
                    {
                        if(!undo&&cancelled()){abort=true;break;}
                        try
                        {
                            using(var bag=new ComBag())
                            {
                                object book=bag.Own(FindBook(App,owner));if(book==null)throw new IOException("目标工作簿已关闭；本步骤停止。");
                                if (!undo) RequireAvailableReplacementTarget(c.New);
                                RequireSaved(book);var sourceNames=ActiveHostSources(book,owner,s=>Operations.Trace("宿主来源",s));var before=sourceNames.Select(s=>s.Path).ToList();string from=undo?c.New:c.Old,to=undo?c.Old:c.New;
                                if(!before.Contains(from,StringComparer.OrdinalIgnoreCase)||before.Contains(to,StringComparer.OrdinalIgnoreCase))throw new IOException("执行期间链接来源发生变化，暂停处理。");
                                c.Intent=true;c.IntentReverse=undo;Batch.Persist();
                                string rawName=sourceNames.Single(s=>string.Equals(s.Path,from,StringComparison.OrdinalIgnoreCase)).Name;
                                Operations.Trace("调用ChangeLink","Name="+rawName+"；NewName="+to+"；Type=1");
                                using(var quiet=new HostQuiet(App))((dynamic)book).ChangeLink(rawName,to,1);
                                // Intent survives Save failure. Do not treat WPS's stale
                                // cached old names as an unsuccessful replacement.
                                WorkbookSaving.Save(book,true);
                                var expected=new HashSet<string>(before.Select(p=>string.Equals(p,from,StringComparison.OrdinalIgnoreCase)?to:p),StringComparer.OrdinalIgnoreCase);
                                if(!expected.SetEquals(ActiveHostSources(book,owner,s=>Operations.Trace("保存后核对",s)).Select(s=>s.Path)))throw new IOException("保存后实际链接核对失败；保留恢复记录。");
                                c.EverApplied=true;c.Applied=!undo;c.Intent=false;Batch.Persist();
                            }
                            outcome.Success++;Pair(c,undo,true,null);
                        }catch(Exception ex){error=ex;failures.Add(c);if(!undo)c.FailureReason=Reason(ex);Operations.Trace("链接处理异常",owner+"；"+Failure(ex)+"\r\n"+ex);Pair(c,undo,false,ex);break;}
                        yield return (undo?"已还原：":"已替换：")+Path.GetFileName(owner);
                    }
                }
                else
                {
                    // Closed package work stays off the host STA; the iterator
                    // yields while it runs so the UI can accept cancellation.
                    Exception workerError=null;var done=new System.Threading.ManualResetEvent(false);
                    var worker=new System.Threading.Thread(delegate(){try{if(ordered.Count>0)ClosedChange(owner,ordered,undo);}catch(Exception ex){workerError=ex;}finally{done.Set();}});
                    worker.SetApartmentState(System.Threading.ApartmentState.STA);worker.IsBackground=true;worker.Start();
                    while(!done.WaitOne(0))yield return (undo?"正在还原文件：":"正在核对文件：")+Path.GetFileName(owner);
                    done.Dispose();error=workerError;
                    if(error==null){foreach(var c in ordered){outcome.Success++;Pair(c,undo,true,null);}}
                    else {Operations.Trace("文件处理异常",owner+"；"+error);foreach(var c in ordered){failures.Add(c);if(!undo)c.FailureReason=Reason(error);Pair(c,undo,false,error);}}
                }
                if(!undo&&cancelled())abort=true;
                if(abort)break;
            }
            outcome.Cancelled = abort;
        }

        private void ClearTemporaryProtection(bool rollback)
        {
            foreach (var checkpoint in Batch.Checkpoints.ToArray())
            {
                if (rollback && (checkpoint.NeedsRestore ||
                    Batch.Changes.Any(c => c.Owner == checkpoint.Owner && (c.Applied || c.Intent))))
                    continue;
                try { Batch.RemoveTemporaryCopy(checkpoint); }
                catch (Exception ex)
                {
                    Operations.Trace("临时副本清理异常", ex.ToString());
                    Operations.Record("临时保护副本未清理", checkpoint.Backup + "\r\n" + Reason(ex));
                }
            }
            if (rollback && Batch.State == "UndoFailed")
            {
                Operations.Record("保留待核对记录", Batch.Journal);
                RecordUnresolvedCopies();
            }
        }

        private void RecordUnresolvedCopies()
        {
            foreach (var checkpoint in Batch.Checkpoints.Where(c => c.NeedsRestore))
                Operations.Record("保留待核对副本", checkpoint.Owner + "\r\n" + checkpoint.Backup);
        }

        private void RefreshStatuses(bool rollback)
        {
            foreach (var link in Links)
            {
                var change = Batch.Changes.FirstOrDefault(c => string.Equals(c.Owner, link.OwnerPath, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(c.ReportSourceId, link.SourceId, StringComparison.OrdinalIgnoreCase));
                if (change == null) continue;

                if (change.Intent) link.Status = "处理结果待核对，请查看操作日志";
                else if (change.Applied)
                {
                    link.Status = change.Action == LinkAction.Break ? "链接已断开，结果已保存" :
                        change.DestinationMissing ? MissingStatus : "链接已替换，结果已保存";
                    if (rollback) link.Status += "；自动回退未完成";
                }
                else if (rollback && change.EverApplied) link.Status = "本批处理已回退";
                else link.Status = change.FailureReason == null ? "未执行" : "未执行；失败：" + change.FailureReason;

                var row = Batch.Rows.FirstOrDefault(r => string.Equals(r.Owner, link.OwnerPath, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(r.Source ?? r.Raw, link.SourceId, StringComparison.OrdinalIgnoreCase));
                if (row != null) row.Status = link.Status;
            }
        }

        private void SaveSummary()
        {
            Operations.Record("", LastResult);
            Batch.LastSummary = LastResult;
            Batch.UserLog = Operations.Text;
            try { Batch.Persist(); }
            catch (Exception ex)
            {
                Operations.Trace("链接记录保存异常", ex.ToString());
                Operations.Record("链接记录未完整保存", Reason(ex) + "\r\n请复制窗口日志。");
                LogFailure = ex.Message;
            }
        }
        private void UpdateReport(bool rollback)
        {
            using (ComBag bag = new ComBag())
            {
                dynamic book = bag.Own(FindBook(App, ReportBookPath)); if (book == null) throw new InvalidOperationException("清单工作簿没有打开。");
                var checkpoint=Checkpoint(ReportBookPath);
                if (rollback && ((checkpoint != null && checkpoint.NeedsRestore) ||
                    Batch.Changes.Any(c => c.Owner == ReportBookPath && c.Intent)))
                    throw new IOException("清单工作簿的回退结果待核对，请查看操作日志。");
                if (checkpoint != null && !checkpoint.Restored) GuardCheckpoint(checkpoint, (object)book);
                else RequireSaved((object)book);
                dynamic sheet = ReportSheet((object)book, bag);
                dynamic used = bag.Own(sheet.UsedRange), rows = bag.Own(used.Rows);
                if (Convert.ToInt32(rows.Count) != Links.Count+1) throw new InvalidOperationException("处理期间清单行数已变化，实际结果见日志。");
                dynamic data = bag.Own(sheet.Range["A1:D" + (Links.Count + 1)]); Array values = data.Value2 as Array;
                if (values == null) throw new InvalidOperationException("无法读取清单身份。");
                var known = Links.ToDictionary(l => l.Key, StringComparer.OrdinalIgnoreCase); int r0 = values.GetLowerBound(0), c0 = values.GetLowerBound(1);
                for (int c=0;c<4;c++) if (Convert.ToString(values.GetValue(r0,c0+c))!=Headers[c]) throw new InvalidOperationException("处理期间清单表头变化，实际结果见日志。");
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i=1;i<=Links.Count;i++) if (!seen.Add(Convert.ToString(values.GetValue(r0+i,c0))+"\n"+Convert.ToString(values.GetValue(r0+i,c0+1))) || !known.ContainsKey(Convert.ToString(values.GetValue(r0+i,c0))+"\n"+Convert.ToString(values.GetValue(r0+i,c0+1)))) throw new InvalidOperationException("处理期间清单行身份变化，实际结果见日志。");
                if(checkpoint!=null&&!checkpoint.Restored){checkpoint.Uncertain=true;Batch.Persist();}
                for (int i = 1; i <= Links.Count; i++)
                {
                    string key = Convert.ToString(values.GetValue(r0 + i, c0)) + "\n" + Convert.ToString(values.GetValue(r0 + i, c0 + 1)); ExternalLinkInfo link;
                    if (!known.TryGetValue(key, out link)) throw new InvalidOperationException("清单行身份变化，请查看操作日志。");
                    dynamic cell = bag.Own(sheet.Range["D" + (i + 1)]); cell.Value2 = link.Status;
                }
                WorkbookUi.FitOutput((object)sheet, 1, 4, Links.Count + 1);
                sheet.Activate(); WorkbookSaving.Save((object)book, true);
                if(checkpoint!=null&&!checkpoint.Restored){checkpoint.LastHash=XlsxLinkPackage.Hash(ReportBookPath);checkpoint.Uncertain=false;Batch.Persist();}
            }
        }
    }

    internal sealed class LinkAssistantForm : Form
    {
        internal readonly RadioButton Current=new RadioButton {Text="当前工作簿",Checked=true,AutoSize=true},Other=new RadioButton {Text="自选其他工作簿",AutoSize=true},Folder=new RadioButton {Text="自选其他文件夹",AutoSize=true};
        internal readonly TextBox LocationBox=new TextBox {ReadOnly=true};
        internal readonly Label Status=Ui.Label("选择范围后扫描链接。",9,true),ScopeLabel=Ui.Label("扫描范围",10,false);
        internal readonly RichTextBox LogView=new RichTextBox {ReadOnly=true,WordWrap=true,ScrollBars=RichTextBoxScrollBars.Vertical,BorderStyle=BorderStyle.None,BackColor=Color.White,Dock=DockStyle.Fill};
        internal readonly Button CopyLog = Ui.Button("复制日志", 96), ScanButton = Ui.Button("扫描链接", 116),
            ExecuteButton = Ui.Button("执行替换", 116), Browse = Ui.Button("选择…", 96);
        internal readonly LinkOperationLog Operations;
        internal Action<string> CopyText=text=>Clipboard.SetDataObject(text,true);
        internal readonly string Guidance = "先扫描链接，在「外部链接」第三列填写完整新路径，或填写「断开链接」，再执行替换。\r\n" +
            "断链会将所选来源的公式转为当前值，并处理未使用的名称。名称仍被使用或无法安全核对时停止。结果自动保存，失败时会尝试回退。";
        private readonly UiCard scope=new UiCard(),logCard=new UiCard();
        private readonly Label logTitle=Ui.Label("操作日志",10,true),help=Ui.Label("当前工作簿接收清单；文件夹只扫描本层.xlsx，跳过临时文件。",9,false);
        private readonly FlowLayoutPanel options=new FlowLayoutPanel {WrapContents=true},actions=new FlowLayoutPanel {WrapContents=true};
        private readonly TableLayoutPanel layout=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=4,Padding=new Padding(16)};
        private readonly LinkScanSelection selection;
        private readonly LinkAssistantSession session;
        private readonly System.Windows.Forms.Timer stepTimer=new System.Windows.Forms.Timer {Interval=20};
        private IEnumerator<string> steps;
        private bool running,cancelRequested,closeAfter,scopeDirty;
        internal bool IsWorking {get{return running;}}
        internal LinkAssistantForm(LinkAssistantSession bound)
        {
            session=bound;Text="链接助手";DialogSizing.Prepare(this,"links.assistant","link",Ui.Blue);
            Operations=session==null?new LinkOperationLog():session.Operations;selection=new LinkScanSelection(session==null?"离线预览.xlsx":session.ReportBookPath,Operations);
            foreach(var opt in new[]{Current,Other,Folder}){opt.Font=Ui.MakeFont(10,FontStyle.Bold,Font);opt.ForeColor=Ui.Blue;options.Controls.Add(opt);}
            foreach(var label in new[]{Status,ScopeLabel,help,logTitle}){label.AutoEllipsis=false;label.AutoSize=true;}Status.ForeColor=Ui.Blue;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.Percent,100));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var scopeRows=new TableLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=1,RowCount=3,Margin=Padding.Empty};
            scopeRows.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));for(int i=0;i<3;i++)scopeRows.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            options.Dock=DockStyle.Top;options.AutoSize=true;options.AutoSizeMode=AutoSizeMode.GrowAndShrink;options.Margin=new Padding(0,0,0,8);
            foreach(var opt in new[]{Current,Other,Folder})opt.Margin=new Padding(0,0,16,4);
            var pathRow=new TableLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=3,RowCount=1,Margin=new Padding(0,0,0,8)};
            pathRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            ScopeLabel.Dock=DockStyle.None;ScopeLabel.Anchor=AnchorStyles.Left;ScopeLabel.Margin=new Padding(0,0,8,0);LocationBox.Anchor=AnchorStyles.Left|AnchorStyles.Right;
            pathRow.Controls.Add(ScopeLabel,0,0);pathRow.Controls.Add(LocationBox,1,0);pathRow.Controls.Add(Browse,2,0);
            scopeRows.Controls.Add(options,0,0);scopeRows.Controls.Add(pathRow,0,1);scopeRows.Controls.Add(help,0,2);
            scope.Dock = DockStyle.Top;
            scope.Controls.Add(scopeRows);
            Action fitScope = delegate
            {
                int height = scopeRows.Height + scope.Padding.Vertical;
                if (scope.Height != height) scope.Height = height;
            };
            scopeRows.SizeChanged += delegate { fitScope(); };
            var logRows=new TableLayoutPanel {Dock=DockStyle.Fill,ColumnCount=1,RowCount=2,Margin=Padding.Empty};logRows.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));logRows.RowStyles.Add(new RowStyle(SizeType.AutoSize));logRows.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            var header=new TableLayoutPanel {Dock=DockStyle.Top,AutoSize=true,AutoSizeMode=AutoSizeMode.GrowAndShrink,ColumnCount=2,RowCount=1,Margin=new Padding(0,0,0,8)};header.RowStyles.Add(new RowStyle(SizeType.AutoSize));header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100));header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));header.Controls.Add(logTitle,0,0);header.Controls.Add(CopyLog,1,0);
            logRows.Controls.Add(header,0,0);logRows.Controls.Add(LogView,0,1);logCard.Dock=DockStyle.Fill;logCard.Controls.Add(logRows);
            actions.Controls.AddRange(new Control[] { ScanButton, ExecuteButton });
            actions.Dock=DockStyle.Top;actions.AutoSize=true;actions.AutoSizeMode=AutoSizeMode.GrowAndShrink;actions.Margin=Padding.Empty;
            foreach(Control button in actions.Controls)button.Margin=new Padding(0,0,8,4);
            Status.Margin=new Padding(0,0,0,8);Status.Dock=DockStyle.Top;
            layout.Controls.Add(scope,0,0);layout.Controls.Add(logCard,0,1);layout.Controls.Add(Status,0,2);layout.Controls.Add(actions,0,3);Controls.Add(layout);
            Action fitWidth = delegate
            {
                int width = Math.Max(80, layout.ClientSize.Width - layout.Padding.Horizontal);
                int innerWidth = Math.Max(80, width - scope.Padding.Horizontal);
                actions.MaximumSize = new Size(width, 0);
                options.MaximumSize = new Size(innerWidth, 0);
                help.MaximumSize = new Size(innerWidth, 0);
                Status.MaximumSize = new Size(width, 0);
                fitScope();
            };
            Resize += delegate { fitWidth(); };
            fitWidth();
            if(Operations.Text.Length==0){LogView.SelectionFont=Ui.MakeFont(9,FontStyle.Bold,Font);LogView.AppendText(Guidance);}LogView.SelectionFont=Ui.MakeFont(9,FontStyle.Regular,Font);LogView.AppendText(Operations.Text);
            Operations.LineAdded+=AppendLog;
            Operations.DisplayReset+=ResetLog;
            CopyLog.Click+=delegate {try{CopyText(LogView.Text);Status.Text="已复制操作日志。";}catch(Exception ex){Status.Text="复制失败，可选择日志文字复制："+ex.Message;}};
            LocationBox.Text=session==null?"离线预览.xlsx":session.ReportBookPath;
            Current.CheckedChanged+=delegate {UpdateScope();};Other.CheckedChanged+=delegate {UpdateScope();};Folder.CheckedChanged+=delegate {UpdateScope();};
            Browse.Click+=delegate {if(Other.Checked){string[] selected=HostDialogService.OpenFiles(this);SelectPaths(selected,selected.Length>0);}else {string selected=HostDialogService.Folder(this,"选择仅扫描本层的文件夹",true);SelectPaths(selected==null?new string[0]:new[]{selected},selected!=null);}};
            ScanButton.Click+=delegate {Run(false);};ExecuteButton.Click+=delegate {Run(true);};
            stepTimer.Tick+=delegate {Advance();};
            FormClosing += delegate(object sender, FormClosingEventArgs e)
            {
                if (!running) return;
                cancelRequested = closeAfter = true;
                e.Cancel = true;
                Status.Text = "等待当前步骤结束并核对结果后关闭。";
            };
            FormClosed+=delegate {stepTimer.Stop();stepTimer.Dispose();if(steps!=null)steps.Dispose();Operations.LineAdded-=AppendLog;Operations.DisplayReset-=ResetLog;};
            if(session!=null&&session.Batch!=null)
            {
                if(session.Batch.ScopeFolder!=null){Folder.Checked=true;SelectPaths(new[]{session.Batch.ScopeFolder},true);}
                else if(session.Batch.Owners.Count!=1||!string.Equals(session.Batch.Owners[0],session.ReportBookPath,StringComparison.OrdinalIgnoreCase)){Other.Checked=true;SelectPaths(session.Batch.Owners.ToArray(),true);}
                scopeDirty = false;
                Status.Text = session.CanExecute ? "已加载扫描清单，可核对后执行替换。" : "可扫描下一批，上一批结果见操作日志。";
            }
            Shown+=delegate {layout.PerformLayout();LogView.SelectionStart=0;LogView.SelectionLength=0;LogView.ScrollToCaret();};RefreshActions();
        }
        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData != Keys.Escape) return base.ProcessDialogKey(keyData);
            Close();
            return true;
        }
        private void ResetLog(){LogView.Clear();LogView.SelectionFont=Ui.MakeFont(9,FontStyle.Regular,Font);}
        private void AppendLog(string line){LogView.AppendText(line);LogView.SelectionStart=LogView.TextLength;LogView.ScrollToCaret();}
        internal void SelectPaths(string[] paths,bool accepted){selection.Choose(paths,accepted);if(accepted){LocationBox.Text=string.Join("；",selection.Picked);scopeDirty=true;RefreshActions();}}
        private void RefreshActions()
        {
            ScanButton.Enabled = !running;
            ExecuteButton.Enabled = !running && !scopeDirty && session != null && session.CanExecute;
            Browse.Enabled = !running && !Current.Checked;
            options.Enabled = !running;
        }
        private void UpdateScope()
        {
            if(running||(!Current.Checked&&!Other.Checked&&!Folder.Checked))return;
            if(!selection.SetScope(Current.Checked?LinkScanScope.Current:Other.Checked?LinkScanScope.Workbooks:LinkScanScope.Folder))return;
            scopeDirty=true;LocationBox.Text=Current.Checked?session==null?"离线预览.xlsx":session.ReportBookPath:"请点击右侧选择…";RefreshActions();
        }
        private IEnumerable<string> ScanSteps()
        {
            string[] files=selection.Candidates();if(files.Length==0)throw new InvalidOperationException("没有可扫描的工作簿。");
            session.ScanFolder=selection.Scope==LinkScanScope.Folder?selection.Picked[0]:null;
            yield return "正在扫描并更新固定清单…";
            if(cancelRequested)yield break;
            session.Scan(files);scopeDirty=false;
        }
        private void Run(bool execute){if(running||session==null)return;Start(execute ? session.Steps(() => cancelRequested) : ScanSteps());}
        private void Start(IEnumerable<string> work){if(running)return;running=true;cancelRequested=false;steps=work.GetEnumerator();RefreshActions();stepTimer.Start();}
        private void Advance()
        {
            stepTimer.Stop();bool more=false;
            try {more=steps.MoveNext();if(more)Status.Text=steps.Current;else Status.Text=session.LastResult??"操作结束。";}
            catch (Exception ex) { Status.Text = "已停止：" + ex.Message; session.ReportStopped(ex); }
            if(more)stepTimer.Start();else {running=false;steps.Dispose();steps=null;RefreshActions();if(session.LogFailure!=null)Status.Text+="；诊断日志未完整写入，请复制窗口日志。";if(closeAfter)Close();}
        }
    }
}
