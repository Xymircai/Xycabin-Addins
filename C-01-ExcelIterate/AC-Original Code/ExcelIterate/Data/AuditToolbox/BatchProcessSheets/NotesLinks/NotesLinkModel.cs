using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class NotesLinkConfig
    {
        internal string Parent="",Folder="",LeadSheet="Lead",ParentCodeCell="K2",CodeColumn="K",NameColumn="M",FileColumn="N";
        internal int FirstRow=3,LastRow;
        internal string SourceSheet="Notes",TargetSheet="Notes",SourceRange="D1:X1700",FirstStart="AV1";
        internal int ColumnSpacing=22;
        internal bool NativePaste=true;
        internal NotesLinkConfig Copy(){return (NotesLinkConfig)MemberwiseClone();}
        internal string LeadKey { get{return string.Join("\u001f",new[]{Parent,Folder,LeadSheet,ParentCodeCell,CodeColumn,NameColumn,FileColumn,FirstRow.ToString(),LastRow.ToString()});} }
        internal string Key { get{return LeadKey+"\u001f"+string.Join("\u001f",new[]{SourceSheet,TargetSheet,SourceRange,FirstStart,ColumnSpacing.ToString(),NativePaste.ToString()});} }
        internal MigrationArea Validate()
        {
            NotesLinkRules.RequireWorkbook(Parent);
            if(!Path.IsPathRooted(Folder)||!Directory.Exists(Folder))throw new InvalidOperationException("请选择存在的绝对子公司文件夹路径。");
            if(FirstRow<1||FirstRow>100000||LastRow<0||LastRow>100000||(LastRow>0&&LastRow<FirstRow))throw new InvalidOperationException("Lead名单行数应在1至100000之间，结束行0表示自动读取。");
            MigrationPlanner.ColumnIndex(CodeColumn);MigrationPlanner.ColumnIndex(NameColumn);MigrationPlanner.ColumnIndex(FileColumn);
            if(MigrationArea.Parse(ParentCodeCell).Rows!=1||MigrationArea.Parse(ParentCodeCell).Columns!=1)throw new InvalidOperationException("父公司编码位置应为一个单元格。");
            var source=MigrationArea.Parse(SourceRange);source.Check(1048576,16384);
            if((long)source.Rows*source.Columns>100000)throw new InvalidOperationException("每家公司最多链接100000格，请缩小来源范围，不能选择整列。");
            NotesLinkRules.Target(source,FirstStart,ColumnSpacing,0);return source;
        }
    }
    internal sealed class NotesLinkTask
    {
        internal int LeadRow,Slot;
        internal bool Selected=true,Confirmed,NameNeedsConfirmation,NativeCompatible;
        internal string Code="",Name="",File="",Sheet="",Status="待预检",Hash="";
        internal List<string> SheetNames=new List<string>();
        internal List<MigrationArea> SourceMerges=new List<MigrationArea>(),Blocks=new List<MigrationArea>(),Anchors=new List<MigrationArea>();
        internal MigrationArea Target;
        internal int Occupied,Skipped;
        internal NotesLinkTask Copy(){var copy=(NotesLinkTask)MemberwiseClone();copy.SheetNames=new List<string>(SheetNames);copy.SourceMerges=new List<MigrationArea>(SourceMerges);copy.Blocks=new List<MigrationArea>(Blocks);copy.Anchors=new List<MigrationArea>(Anchors);return copy;}
    }
    internal sealed class NotesLinkPlan
    {
        internal NotesLinkConfig Config;
        internal string ParentHash,TargetSheet;
        internal MigrationArea Source;
        internal MigrationArea Cleanup;
        internal NotesSheetFacts Original;
        internal readonly List<NotesLinkTask> Tasks=new List<NotesLinkTask>();
        internal int Occupied {get{return Tasks.Sum(t=>t.Occupied);}}
    }
    internal sealed class NotesLinkBatchItem
    {
        internal string Parent,Status="待预检",Backup,CommittedHash;
        internal bool Selected=true;
        internal NotesLinkPlan Plan;
        internal bool Ready {get{return Plan!=null&&Plan.Tasks.Count>0&&Plan.Tasks.All(t=>t.Status.StartsWith("通过",StringComparison.Ordinal));}}
    }
    internal static class NotesLinkBatch
    {
        internal static List<string> Files(IEnumerable<string> paths)
        {
            var result=new List<string>();var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var path in paths){string full=Path.GetFullPath(path);string name=Path.GetFileName(full),ext=Path.GetExtension(full);if(name.StartsWith("~$",StringComparison.Ordinal)||name.StartsWith(".xy-",StringComparison.OrdinalIgnoreCase))continue;if(!ext.Equals(".xlsx",StringComparison.OrdinalIgnoreCase)&&!ext.Equals(".xlsm",StringComparison.OrdinalIgnoreCase))continue;if(seen.Add(full))result.Add(full);}
            return result;
        }
        internal static string PickerParent(NotesLinkConfig shared,IEnumerable<NotesLinkBatchItem> input,string preferred)
        {
            var selected=input.Where(i=>i.Selected).ToList();
            if(selected.Count==0)throw new InvalidOperationException("请先在列表中勾选汇总文件，再选择参数位置。");
            var ordered=selected.OrderBy(i=>string.Equals(i.Parent,preferred,StringComparison.OrdinalIgnoreCase)?0:1).ToList();
            // Prepared plans already identify actual summary workbooks without reopening files.
            var ready=ordered.FirstOrDefault(i=>i.Ready);if(ready!=null)return ready.Parent;
            foreach(var item in ordered)
            {
                var cfg=shared.Copy();cfg.Parent=item.Parent;
                try{if(NotesLinkRules.ReadLead(cfg).Count>0)return item.Parent;}catch(InvalidOperationException){}catch(IOException){}
            }
            // A mispicked header must still be correctable: only the reference-file
            // identification uses the standard owner row, never the execution config.
            foreach(var item in ordered)
            {
                var cfg=shared.Copy();cfg.Parent=item.Parent;cfg.ParentCodeCell=CellAddress.Name(2,MigrationPlanner.ColumnIndex(cfg.CodeColumn));
                try{if(NotesLinkRules.ReadLead(cfg).Count>0)return item.Parent;}catch(InvalidOperationException){}catch(IOException){}
            }
            throw new InvalidOperationException("没有找到可用的汇总文件。请检查是否只添加了子公司明细文件，并核对Lead的编码列和名单起始行。");
        }
        internal static List<NotesLinkBatchItem> Prepare(NotesLinkConfig shared,IEnumerable<NotesLinkBatchItem> input,bool confirmed,Action<string> log,Func<bool> cancelled,IDictionary<string,string> blockers=null)
        {
            var result=new List<NotesLinkBatchItem>();
            foreach(var item in input.Where(i=>i.Selected))
            {
                if(cancelled!=null&&cancelled())throw new OperationCanceledException("已停止预检，未修改文件。");
                var next=new NotesLinkBatchItem{Parent=item.Parent,Backup=item.Backup,CommittedHash=item.CommittedHash};result.Add(next);
                try
                {
                    string error;if(blockers!=null&&blockers.TryGetValue(item.Parent,out error))throw new InvalidOperationException(error);
                    var cfg=shared.Copy();cfg.Parent=item.Parent;
                    if(log!=null)log(Path.GetFileName(item.Parent)+"｜名单设置："+cfg.LeadSheet+"；汇总公司编码在 "+cfg.ParentCodeCell+"；编码列 "+cfg.CodeColumn+"；从第 "+cfg.FirstRow+" 行读至 "+(cfg.LastRow==0?"末行（自动）":cfg.LastRow.ToString()));
                    var rows=NotesLinkRules.ReadLead(cfg);foreach(var row in rows)row.Confirmed=confirmed;
                    next.Plan=NotesLinkPlanner.Build(cfg,rows,message=>{if(log!=null)log(Path.GetFileName(item.Parent)+"｜"+message);},cancelled);
                    int blocked=next.Plan.Tasks.Count(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal));
                    next.Status=blocked==0?"通过｜公司 "+rows.Count+" 家，覆盖 "+next.Plan.Occupied+" 格":"阻止｜公司 "+rows.Count+" 家，其中 "+blocked+" 项未通过";
                }
                catch(OperationCanceledException){throw;}
                catch(Exception error){next.Plan=null;next.Status="阻止："+error.Message;}
                if(log!=null)log(Path.GetFileName(next.Parent)+"｜"+next.Status);
            }
            return result;
        }
        // Process a parent before another selected workbook that it references.
        // No link refresh is needed, and later writes cannot invalidate an earlier source hash.
        internal static List<NotesLinkBatchItem> Order(IEnumerable<NotesLinkBatchItem> input)
        {
            var remaining=input.Where(i=>i.Selected&&i.Ready).ToList();var result=new List<NotesLinkBatchItem>();
            while(remaining.Count>0)
            {
                var next=remaining.FirstOrDefault(i=>!remaining.Any(other=>other!=i&&other.Plan.Tasks.Any(t=>string.Equals(Path.GetFullPath(t.File),Path.GetFullPath(i.Parent),StringComparison.OrdinalIgnoreCase))));
                if(next==null)throw new InvalidOperationException("所选汇总文件之间存在循环引用，请取消其中一个文件后重新预检。");
                result.Add(next);remaining.Remove(next);
            }
            return result;
        }
    }
    internal sealed class NotesLinkBatchRun
    {
        private readonly object application;
        private readonly List<NotesLinkBatchItem> items;
        private readonly Action<string> log;
        internal bool Cancelled;
        internal int Completed,Succeeded,Failed;
        internal NotesLinkBatchRun(object app,IEnumerable<NotesLinkBatchItem> prepared,Action<string> output)
        {application=app;items=NotesLinkBatch.Order(prepared);log=output??delegate{};}
        internal IEnumerable<string> Run()
        {
            foreach(var item in items)
            {
                if(Cancelled)break;
                item.Status="正在粘贴链接";
                using(var execution=new NotesLinkExecution(application,item.Plan,message=>log(Path.GetFileName(item.Parent)+"｜"+message)))
                using(var steps=execution.Run().GetEnumerator())
                {
                    Exception failure=null;
                    while(true)
                    {
                        bool more=false;string message=null;
                        try{execution.Cancelled=Cancelled;more=steps.MoveNext();if(more)message=steps.Current;}
                        catch(Exception error){failure=error;}
                        if(failure!=null||!more)break;
                        yield return Path.GetFileName(item.Parent)+"｜"+message;
                    }
                    if(execution.Committed)
                    {item.Backup=execution.Backup;item.CommittedHash=execution.CommittedHash;item.Status=failure==null?"已完成｜公司 "+item.Plan.Tasks.Count+" 家":"已更新，有后处理提示："+failure.Message;Succeeded++;Completed++;}
                    else if(failure is OperationCanceledException||Cancelled){item.Status="已停止｜未更新";Cancelled=true;}
                    else{item.Status="失败｜未更新："+(failure==null?"未完成提交":failure.Message);Failed++;Completed++;}
                    if(failure!=null)log(Path.GetFileName(item.Parent)+"｜"+item.Status);
                }
                yield return Path.GetFileName(item.Parent)+"｜"+item.Status;
                if(Cancelled)break;
            }
            if(Cancelled)foreach(var item in items.Where(i=>i.Status.StartsWith("通过",StringComparison.Ordinal)))item.Status="未执行｜已停止";
        }
    }
    internal static class NotesLinkRules
    {
        internal static void RequireWorkbook(string file)
        {string ext=Path.GetExtension(file);if(!Path.IsPathRooted(file)||!File.Exists(file)||!(ext.Equals(".xlsx",StringComparison.OrdinalIgnoreCase)||ext.Equals(".xlsm",StringComparison.OrdinalIgnoreCase)))throw new InvalidOperationException("请选择存在的.xlsx或.xlsm工作簿的完整路径："+file);}
        internal static bool Intersects(MigrationArea a,MigrationArea b){return a.Row<b.Row+b.Rows&&b.Row<a.Row+a.Rows&&a.Column<b.Column+b.Columns&&b.Column<a.Column+a.Columns;}
        internal static bool Same(MigrationArea a,MigrationArea b){return a.Row==b.Row&&a.Column==b.Column&&a.Rows==b.Rows&&a.Columns==b.Columns;}
        internal static bool Inside(MigrationArea area,int row,int col){return row>=area.Row&&row<area.Row+area.Rows&&col>=area.Column&&col<area.Column+area.Columns;}
        internal static MigrationArea Target(MigrationArea source,string first,string second,int slot)
        {
            var a=MigrationArea.Parse(first);var b=MigrationArea.Parse(second);
            if(a.Rows!=1||a.Columns!=1||b.Rows!=1||b.Columns!=1||a.Row!=b.Row)throw new InvalidOperationException("两个粘贴起点必须是同一行的单个单元格。");
            int stride=b.Column-a.Column;if(stride<source.Columns)throw new InvalidOperationException("第二个起点应在右侧，列间距不能小于来源宽度"+source.Columns+"列。");
            long column=a.Column+(long)slot*stride;if(column+source.Columns-1>16384)throw new InvalidOperationException("第"+(slot+1)+"家公司目标超出Excel最大列数。");
            var result=new MigrationArea{Row=a.Row,Column=(int)column,Rows=source.Rows,Columns=source.Columns};result.Check(1048576,16384);return result;
        }
        internal static MigrationArea Target(MigrationArea source,string first,int spacing,int slot)
        {
            var start=MigrationArea.Parse(first);
            if(start.Rows!=1||start.Columns!=1)throw new InvalidOperationException("第一家公司的放置位置请选一个单元格。");
            if(spacing<source.Columns||spacing>16384)throw new InvalidOperationException("列间距至少为"+source.Columns+"列（每家公司占用的列数），最多16384列，避免两家公司重叠。");
            if(slot<0)throw new InvalidOperationException("公司顺序无效。");
            long column=start.Column+(long)slot*spacing;
            if(column+source.Columns-1>16384)throw new InvalidOperationException("第"+(slot+1)+"家公司超出Excel最右列，请减少公司数或调整放置位置。");
            var result=new MigrationArea{Row=start.Row,Column=(int)column,Rows=source.Rows,Columns=source.Columns};result.Check(1048576,16384);return result;
        }
        internal static List<string> Candidates(IEnumerable<string> names,string requested)
        {
            string key=(requested??"").Trim();var all=names.ToList();
            var literal=all.Where(n=>n==requested).ToList();if(literal.Count==1)return literal;
            var exact=all.Where(n=>string.Equals(n.Trim(),key,StringComparison.OrdinalIgnoreCase)).ToList();if(exact.Count>0)return exact;
            if(key.Equals("Note",StringComparison.OrdinalIgnoreCase)||key.Equals("Notes",StringComparison.OrdinalIgnoreCase))
            {var aliases=all.Where(n=>n.Trim().Equals("Note",StringComparison.OrdinalIgnoreCase)||n.Trim().Equals("Notes",StringComparison.OrdinalIgnoreCase)).ToList();if(aliases.Count>0)return aliases;}
            return all.Where(n=>key.Length>0&&n.Trim().IndexOf(key,StringComparison.OrdinalIgnoreCase)>=0).ToList();
        }
        internal static bool ExactOrAlias(string name,string requested)
        {string n=(name??"").Trim(),r=(requested??"").Trim();return n.Equals(r,StringComparison.OrdinalIgnoreCase)||((n.Equals("Note",StringComparison.OrdinalIgnoreCase)||n.Equals("Notes",StringComparison.OrdinalIgnoreCase))&&(r.Equals("Note",StringComparison.OrdinalIgnoreCase)||r.Equals("Notes",StringComparison.OrdinalIgnoreCase)));}
        internal static string Resolve(IEnumerable<string> names,string requested)
        {var list=Candidates(names,requested);if(list.Count!=1)throw new InvalidOperationException(list.Count==0?"找不到工作表“"+requested+"”，请手动选择。":"存在多个候选表，请明确选择："+string.Join("、",list));return list[0];}
        internal static bool IsDirect(string parent,string child)
        {string p=CompanyStructureTasks.Key(parent),c=CompanyStructureTasks.Key(child);if(p==c)return false;int dash=c.LastIndexOf('-');return p=="0"?dash<0:dash>=0&&c.Substring(0,dash)==p;}
        internal static string LeadCodeKey(string value,string sheet,string address,bool parent)
        {
            try{return CompanyStructureTasks.Key(value);}
            catch(InvalidOperationException error)
            {
                string display=string.IsNullOrEmpty(value)?"（空白）":string.Concat(value.Take(120).Select(c=>char.IsControl(c)||char.GetUnicodeCategory(c)==System.Globalization.UnicodeCategory.Format?"[不可见字符U+"+((int)c).ToString("X4")+"]":c.ToString()));
                throw new InvalidOperationException(sheet+"!"+address+"："+(parent?"汇总公司编码":"公司编码")+"读到“"+display+"”。"+error.Message+(parent?" 请检查“汇总公司编码在”，通常填K2。":" 请检查“公司编码列”和“名单从哪行读”；名单不包含标题行，通常从第3行开始。"),error);
            }
        }
        internal static List<NotesLinkTask> ReadLead(NotesLinkConfig cfg)
        {
            RequireWorkbook(cfg.Parent);var names=DataMigrationHost.PackageSheets(cfg.Parent);string lead=Resolve(names,cfg.LeadSheet);
            var data=MigrationSheetSnapshot.Load(cfg.Parent,lead);var p=MigrationArea.Parse(cfg.ParentCodeCell);if(p.Rows!=1||p.Columns!=1)throw new InvalidOperationException("父公司编码应选择一个单元格。");
            string parent=data.Value(p.Row,p.Column);LeadCodeKey(parent,lead,CellAddress.Name(p.Row,p.Column),true);
            int code=MigrationPlanner.ColumnIndex(cfg.CodeColumn),name=MigrationPlanner.ColumnIndex(cfg.NameColumn),file=MigrationPlanner.ColumnIndex(cfg.FileColumn);
            int last=cfg.LastRow==0?Math.Min(data.MaxRow,100000):cfg.LastRow;var seen=new HashSet<string>();var result=new List<NotesLinkTask>();
            for(int row=cfg.FirstRow;row<=last;row++)
            {
                if(row==p.Row)continue;string value=data.Value(row,code),filename=data.Value(row,file);
                if(string.IsNullOrWhiteSpace(value)){if(!string.IsNullOrWhiteSpace(filename))throw new InvalidOperationException("Lead第"+row+"行有文件名但缺少编码，请检查列和名单范围。");continue;}
                string key=LeadCodeKey(value,lead,CellAddress.Name(row,code),false);if(!seen.Add(key))throw new InvalidOperationException("Lead编码重复："+value+"（第"+row+"行）。");
                if(!IsDirect(parent,value))continue;
                var task=new NotesLinkTask{LeadRow=row,Slot=result.Count,Code=value,Name=data.Value(row,name)};
                try{string safe=CompanyWorkbookRules.FileName(filename);if(Path.GetExtension(safe).Length==0)safe+=".xlsx";task.File=Path.Combine(cfg.Folder,safe);}
                catch(Exception error){task.File="";task.Status="文件名无效："+error.Message;}
                result.Add(task);
            }
            if(result.Count==0)throw new InvalidOperationException("没有识别到直系子公司；请检查父编码单元格、Lead列及名单行范围。");return result;
        }
        internal static string ZeroFormat(string format)
        {
            if(SummaryAreaHost.PreserveNumberFormat(format))return format;
            if(Regex.IsMatch(format??"",@"\[[<>＝=]"))throw new InvalidOperationException("目标含条件数字格式，请先调整为普通金额格式。");
            var parts=new List<string>();var current=new System.Text.StringBuilder();bool quote=false,escape=false;
            foreach(char c in format??"General")
            {if(escape){current.Append(c);escape=false;continue;}if(c=='\\'){current.Append(c);escape=true;continue;}if(c=='"')quote=!quote;if(c==';'&&!quote){parts.Add(current.ToString());current.Length=0;}else current.Append(c);}
            parts.Add(current.ToString());string positive=parts[0];if(positive=="@"||positive.Length==0)positive="General";
            return positive+";"+(parts.Count>1?parts[1]:"-"+positive)+";\"-\";"+(parts.Count>3?parts[3]:"@");
        }
        internal static string ErrorText(object value)
        {
            var wrapper=value as System.Runtime.InteropServices.ErrorWrapper;
            uint raw=wrapper!=null?unchecked((uint)wrapper.ErrorCode):value is uint?(uint)value:unchecked((uint)Convert.ToInt32(value));
            switch(raw&65535u){case 2000:return "#NULL!";case 2007:return "#DIV/0!";case 2015:return "#VALUE!";case 2023:return "#REF!";case 2029:return "#NAME?";case 2036:return "#NUM!";case 2042:return "#N/A";case 2043:return "#GETTING_DATA";default:return "#ERROR("+(raw&65535u)+")";}
        }
        internal static List<MigrationArea> Rectangles(MigrationArea area,bool[,] excluded)
        {
            var result=new List<MigrationArea>();var active=new Dictionary<string,MigrationArea>();
            for(int r=0;r<area.Rows;r++)
            {
                var next=new Dictionary<string,MigrationArea>();
                for(int c=0;c<area.Columns;)
                {
                    if(excluded[r,c]){c++;continue;}int start=c;while(c<area.Columns&&!excluded[r,c])c++;string key=start+":"+c;MigrationArea block;
                    if(!active.TryGetValue(key,out block))result.Add(block=new MigrationArea{Row=area.Row+r,Column=area.Column+start,Rows=0,Columns=c-start});
                    block.Rows++;next[key]=block;
                }
                active=next;
            }
            return result;
        }
        internal static bool CanPasteWhole(MigrationArea source,NotesLinkTask task,NotesSheetFacts target)
        {
            var sourceMerges=task.SourceMerges.Where(m=>Intersects(m,source)).Select(m=>(m.Row-source.Row)+","+(m.Column-source.Column)+","+m.Rows+","+m.Columns).OrderBy(x=>x).ToArray();
            var targetMerges=target.Snapshot.Merges.Where(m=>Intersects(m,task.Target)).ToList();
            if(!sourceMerges.SequenceEqual(targetMerges.Select(m=>(m.Row-task.Target.Row)+","+(m.Column-task.Target.Column)+","+m.Rows+","+m.Columns).OrderBy(x=>x)))return false;
            // Native Paste Link may populate merged followers. Preserve any preexisting hidden content.
            foreach(var merge in targetMerges)for(int r=merge.Row;r<merge.Row+merge.Rows;r++)for(int c=merge.Column;c<merge.Column+merge.Columns;c++)
                if((r!=merge.Row||c!=merge.Column)&&target.Occupied(r,c))return false;
            return true;
        }
        internal static void MapMerges(NotesLinkTask task,MigrationArea source,NotesSheetFacts parent,NotesSheetFacts sourceFacts=null)
        {
            var excluded=new bool[source.Rows,source.Columns];
            foreach(var merge in task.SourceMerges.Where(m=>Intersects(m,source)))
            {
                if(!source.Contains(merge))throw new InvalidOperationException("来源选区截断合并区域："+merge.Address);
                var shifted=new MigrationArea{Row=task.Target.Row+merge.Row-source.Row,Column=task.Target.Column+merge.Column-source.Column,Rows=merge.Rows,Columns=merge.Columns};
                foreach(var targetMerge in parent.Snapshot.Merges.Where(m=>Intersects(m,shifted)))if(!Same(targetMerge,shifted))throw new InvalidOperationException("来源与目标合并结构不兼容："+targetMerge.Address);
                bool matched=parent.Snapshot.Merges.Any(m=>Same(m,shifted));
                for(int row=merge.Row;row<merge.Row+merge.Rows;row++)for(int col=merge.Column;col<merge.Column+merge.Columns;col++)
                {
                    if(row==merge.Row&&col==merge.Column)continue;
                    int tr=task.Target.Row+row-source.Row,tc=task.Target.Column+col-source.Column;long key=MigrationSheetSnapshot.Key(tr,tc);
                    if(!matched&&parent.Occupied(tr,tc))throw new InvalidOperationException("来源合并内部对应目标已有内容："+CellAddress.Name(tr,tc));
                    excluded[row-source.Row,col-source.Column]=true;task.Skipped++;
                }
            }
            foreach(var merge in parent.Snapshot.Merges.Where(m=>Intersects(m,task.Target)))
            {
                if(!task.Target.Contains(merge))throw new InvalidOperationException("目标合并区域跨越公司边界或选区："+merge.Address);
                var original=new MigrationArea{Row=source.Row+merge.Row-task.Target.Row,Column=source.Column+merge.Column-task.Target.Column,Rows=merge.Rows,Columns=merge.Columns};
                if(!task.SourceMerges.Any(m=>Same(m,original)))
                {
                    bool empty=sourceFacts!=null;
                    for(int r=original.Row;r<original.Row+original.Rows&&empty;r++)for(int c=original.Column;c<original.Column+original.Columns;c++)if(sourceFacts.Occupied(r,c)){empty=false;break;}
                    if(!empty)throw new InvalidOperationException("目标合并结构与来源不一致：目标 "+merge.Address+" 对应来源 "+original.Address+"；来源含内容或公式，不能忽略来源格。");
                    for(int r=merge.Row;r<merge.Row+merge.Rows;r++)for(int c=merge.Column;c<merge.Column+merge.Columns;c++)
                    {if(r==merge.Row&&c==merge.Column)continue;excluded[r-task.Target.Row,c-task.Target.Column]=true;task.Skipped++;}
                }

                excluded[merge.Row-task.Target.Row,merge.Column-task.Target.Column]=true;
                task.Anchors.Add(new MigrationArea{Row=merge.Row,Column=merge.Column,Rows=1,Columns=1});
            }
            task.Blocks=Rectangles(task.Target,excluded);
        }
    }
}
