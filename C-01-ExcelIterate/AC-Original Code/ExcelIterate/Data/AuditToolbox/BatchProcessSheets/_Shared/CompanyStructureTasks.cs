using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class CompanyStructureTask
    {
        internal int Row;
        internal string Sequence, ShortName, FullName, FileName, Level, ParentFullName;
        internal bool Leaf;
    }
    internal sealed class CompanyStructureColumns
    {
        internal int FirstRow = 3, Sequence = 1, ShortName = 3, FullName = 5, FileName = 7, Level = 4, ParentName = 6;
    }
    internal sealed class CompanyParentLinkTask
    {
        internal CompanyStructureTask Parent;
        internal readonly List<CompanyStructureTask> Children = new List<CompanyStructureTask>();
    }
    internal sealed class CompanyStructureSelection
    {
        internal List<CompanyStructureTask> All, Selected;
    }
    internal static class CompanyStructureTasks
    {
        internal static string Key(string sequence)
        {
            // Business markers are ignored only for hierarchy comparisons; original fields remain intact.
            string text = (sequence ?? "").Trim().TrimStart('!', '！').Trim();
            if (!Regex.IsMatch(text, @"^[0-9]+(?:-[0-9]+)*-?$")) throw new InvalidOperationException("公司编码格式不正确，请使用03-、03-01-这样的分段数字，开头可带！或!。");
            return string.Join("-", text.TrimEnd('-').Split('-').Select(s => s.TrimStart('0').Length == 0 ? "0" : s.TrimStart('0')).ToArray());
        }
        internal static void Classify(IList<CompanyStructureTask> rows, bool useLevel)
        {
            var keys = new Dictionary<string, CompanyStructureTask>(StringComparer.Ordinal);
            foreach (var row in rows)
            {
                string key;
                try { key = Key(row.Sequence); }
                catch (Exception error) { throw new InvalidOperationException("架构表第" + row.Row + "行：" + error.Message); }
                if (keys.ContainsKey(key)) throw new InvalidOperationException("架构表序号重复：第" + keys[key].Row + "行与第" + row.Row + "行，" + row.Sequence);
                keys.Add(key, row);
            }
            foreach (var entry in keys)
                entry.Value.Leaf = !keys.Keys.Any(k => k.StartsWith(entry.Key + "-", StringComparison.Ordinal));
            if (useLevel)
            {
                var roots = rows.Where(r => Regex.IsMatch((r.Level ?? "").Trim(), @"^0\s*级?$" )).ToList();
                if (roots.Count > 1) throw new InvalidOperationException("架构表存在多个0级总集团，无法确定归属，请核对级次列。");
                if (roots.Count == 1 && rows.Count > 1) roots[0].Leaf = false;
            }
        }
        internal static List<CompanyWorkbookRecord> Records(IEnumerable<CompanyStructureTask> tasks)
        {
            return tasks.Select(t => new CompanyWorkbookRecord { Row = t.Row, Name = t.FileName, Code = t.ShortName ?? "", FullName = t.FullName }).ToList();
        }
        internal static List<CompanyParentLinkTask> DirectParentLinks(IList<CompanyStructureTask> all,IList<CompanyStructureTask> selected)
        {
            Classify(all,true);
            var keys=all.ToDictionary(t=>Key(t.Sequence),StringComparer.Ordinal);
            var chosen=new HashSet<string>(selected.Select(t=>Key(t.Sequence)),StringComparer.Ordinal);
            if(chosen.Count==0)throw new InvalidOperationException("请勾选要链接到父公司的末级公司。");
            if(chosen.Count!=selected.Count||chosen.Any(k=>!keys.ContainsKey(k)))throw new InvalidOperationException("所选公司不属于当前架构表或存在重复选择。");
            var root=all.SingleOrDefault(t=>Regex.IsMatch((t.Level??"").Trim(),@"^0\s*级?$"));
            var groups=new Dictionary<string,CompanyParentLinkTask>(StringComparer.Ordinal);
            foreach(var child in all.Where(t=>chosen.Contains(Key(t.Sequence))))
            {
                if(!child.Leaf)throw new InvalidOperationException("第"+child.Row+"行不是末级公司；本模式只处理末级公司到直属父公司。");
                var parent=ParentOf(keys,root,child);
                CompanyWorkbookRules.FileName(child.FileName);
                if(!string.Equals(Path.GetExtension(child.FileName),".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("第"+child.Row+"行来源文件须为.xlsx，当前外链规则不自动转换其他格式。");
                string parentKey=Key(parent.Sequence);CompanyParentLinkTask group;
                if(!groups.TryGetValue(parentKey,out group)){group=new CompanyParentLinkTask{Parent=parent};groups.Add(parentKey,group);}
                group.Children.Add(child);
            }
            var ordered=all.Where(t=>groups.ContainsKey(Key(t.Sequence))).Select(t=>groups[Key(t.Sequence)]).ToList();
            ResolveFiles(ordered.Select(g=>g.Parent),Path.GetTempPath());
            foreach(var group in ordered)ResolveFiles(group.Children,Path.GetTempPath());
            return ordered;
        }
        private static CompanyStructureTask ParentOf(IDictionary<string,CompanyStructureTask> keys,CompanyStructureTask root,CompanyStructureTask child)
        {
            string key=Key(child.Sequence);int separator=key.LastIndexOf('-');CompanyStructureTask parent=null;
            if(separator>=0)keys.TryGetValue(key.Substring(0,separator),out parent);else if(root!=null&&!ReferenceEquals(root,child))parent=root;
            if(parent==null)throw new InvalidOperationException("第"+child.Row+"行找不到直属父公司，不会自动跳到更高层。");
            string declared=(child.ParentFullName??"").Trim();
            if((declared=="母公司本部"&&!string.Equals(child.FullName,parent.FullName,StringComparison.Ordinal))||(declared.Length>0&&declared!="母公司本部"&&!string.Equals(declared,(parent.FullName??"").Trim(),StringComparison.Ordinal)))throw new InvalidOperationException("第"+child.Row+"行F列并表母公司与序号层级不一致，请核对后重试。");
            return parent;
        }
        internal static List<CompanyStructureTask> DirectChildren(IList<CompanyStructureTask> all,CompanyStructureTask selected)
        {
            Classify(all,true);var keys=all.ToDictionary(t=>Key(t.Sequence),StringComparer.Ordinal);
            CompanyStructureTask parent;if(!keys.TryGetValue(Key(selected.Sequence),out parent)||parent.Leaf)throw new InvalidOperationException("请选择有下属公司的合并主体。");
            var root=all.SingleOrDefault(t=>Regex.IsMatch((t.Level??"").Trim(),@"^0\s*级?$"));
            // Resolve only this parent's immediate children. A top-level row has no
            // parent and must not invalidate every unrelated consolidation task.
            if(root==null)keys.TryGetValue("0",out root);
            string parentKey=Key(parent.Sequence);
            var children=all.Where(t=>
            {
                if(ReferenceEquals(t,parent)||ReferenceEquals(t,root))return false;
                string key=Key(t.Sequence);int separator=key.LastIndexOf('-');
                return separator>=0?key.Substring(0,separator)==parentKey:ReferenceEquals(parent,root);
            }).ToList();
            foreach(var child in children)ParentOf(keys,root,child);
            var body=children.Where(t=>(t.ParentFullName??"").Trim()=="母公司本部"&&t.FullName==parent.FullName).ToList();
            if(body.Count!=1)throw new InvalidOperationException("该主体须在架构表中明确列出且只列出一条母公司本部。");
            return body.Concat(children.Where(t=>!ReferenceEquals(t,body[0]))).ToList();
        }
        internal static List<string> ResolveFiles(IEnumerable<CompanyStructureTask> tasks,string folder)
        {
            var result=new List<string>();var names=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var task in tasks)
            {
                string file;
                try{file=CompanyWorkbookRules.FileName(task.FileName);}
                catch(Exception error){throw new InvalidOperationException("架构表第"+task.Row+"行："+error.Message,error);}
                if(!names.Add(file))throw new InvalidOperationException("架构表文件名重复："+file);
                string extension=Path.GetExtension(file);if(!string.Equals(extension,".xlsx",StringComparison.OrdinalIgnoreCase)&&!string.Equals(extension,".xls",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("架构表第"+task.Row+"行文件格式暂不支持："+file);
                result.Add(Path.Combine(Path.GetFullPath(folder),file));
            }
            if(result.Count==0)throw new InvalidOperationException("请至少勾选一个架构任务。");return result;
        }
    }
}
