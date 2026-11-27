using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class ConsolidatedLeadEntry
    {
        internal string Code, Name, FilePart, FileName;
        internal static ConsolidatedLeadEntry From(CompanyStructureTask task)
        {
            CompanyStructureTasks.Key(task.Sequence);string code=task.Sequence.Trim().TrimEnd('-')+"-";
            string file=CompanyWorkbookRules.FileName(task.FileName);
            if(!file.EndsWith(".xlsx",StringComparison.OrdinalIgnoreCase)||!file.StartsWith(code,StringComparison.Ordinal))throw new InvalidOperationException("文件名须保留架构表完整序号，并使用.xlsx后缀："+file);
            string part=file.Substring(code.Length),name=part.Substring(0,part.Length-5),full=task.FullName??"";
            if(string.IsNullOrWhiteSpace(full)||(task.Leaf?name!=full:name!=full+"（合并）"&&name!=full+"(合并)"))throw new InvalidOperationException("公司全称、合并标识与文件名不一致："+file);
            return new ConsolidatedLeadEntry{Code=code,Name=name,FilePart=part,FileName=file};
        }
    }
    internal sealed class ConsolidatedWorkbookPlan
    {
        internal CompanyWorkbookRecord Record;
        internal readonly List<ConsolidatedLeadEntry> Entries=new List<ConsolidatedLeadEntry>();
        internal int ClearLast;
        internal int[] Columns;
    }
    internal sealed class ConsolidatedTemplateProfile
    {
        internal int OldLast=2, CapacityLast=100000;
        internal int[] Columns=new[]{11,12,13,14,15};
        private int firstUnsafeRow=int.MaxValue;
        private void UnsafeRange(string address)
        {
            var match=Regex.Match(address??"",@"^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$",RegexOptions.IgnoreCase);
            if(!match.Success)throw new InvalidOperationException("Lead包含无法安全识别的合并或数组区域。");
            int first=Column(match.Groups[1].Value),last=match.Groups[3].Success?Column(match.Groups[3].Value):first;
            int row=int.Parse(match.Groups[2].Value,CultureInfo.InvariantCulture),end=match.Groups[4].Success?int.Parse(match.Groups[4].Value,CultureInfo.InvariantCulture):row;
            if(Columns.Any(c=>c>=first&&c<=last)&&end>=2)firstUnsafeRow=Math.Min(firstUnsafeRow,Math.Max(2,row));
        }
        private static int Column(string name){int value=0;foreach(char c in name.ToUpperInvariant())value=value*26+c-'A'+1;return value;}
        private const string Ns="http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly Regex CompanyRange=new Regex(@"(?<![\w\]])(?:'Lead'|Lead)!\$?[K-O]\$?[34]:\$?[K-O]\$?(\d+)",RegexOptions.IgnoreCase|RegexOptions.CultureInvariant);
        private static XmlDocument Xml(ZipArchiveEntry entry){if(entry==null)throw new InvalidOperationException("模板缺少必要的工作簿结构。");var doc=new XmlDocument{XmlResolver=null};using(var stream=entry.Open())doc.Load(stream);return doc;}
        private static XmlNamespaceManager Names(XmlDocument doc){var ns=new XmlNamespaceManager(doc.NameTable);ns.AddNamespace("s",Ns);return ns;}
        private void Bound(string formula){foreach(Match match in CompanyRange.Matches(formula??""))CapacityLast=Math.Min(CapacityLast,int.Parse(match.Groups[1].Value,CultureInfo.InvariantCulture));}
        internal static ConsolidatedTemplateProfile Read(string path){return Read(path,null);}
        internal static ConsolidatedTemplateProfile Read(string path,int[] columns)
        {
            if(!string.Equals(Path.GetExtension(path),".xlsx",StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("合并模板生成目前使用.xlsx模板。");
            var profile=new ConsolidatedTemplateProfile();if(columns!=null){if(columns.Length!=5||columns.Any(c=>c<1||c>16384)||columns.Distinct().Count()!=5)throw new InvalidOperationException("五个字段必须映射到不同的有效列。");profile.Columns=columns;}
            using(var input=File.OpenRead(path))using(var zip=new ZipArchive(input,ZipArchiveMode.Read))
            {
                var book=Xml(zip.GetEntry("xl/workbook.xml"));var bn=Names(book);
                if(book.SelectSingleNode("/s:workbook/s:workbookProtection[@lockStructure='1']",bn)!=null)throw new InvalidOperationException("模板工作簿结构受保护。");
                var sheet=book.SelectNodes("/s:workbook/s:sheets/s:sheet",bn).Cast<XmlElement>().FirstOrDefault(s=>s.GetAttribute("name")=="Lead");if(sheet==null)throw new InvalidOperationException("合并模板须包含Lead工作表。");
                string id=sheet.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships");var rels=Xml(zip.GetEntry("xl/_rels/workbook.xml.rels"));
                var rel=rels.DocumentElement.ChildNodes.Cast<XmlElement>().First(r=>r.GetAttribute("Id")==id);string target=rel.GetAttribute("Target").Replace('\\','/');
                string leadPath=target.StartsWith("/",StringComparison.Ordinal)?target.TrimStart('/'):(target.StartsWith("xl/",StringComparison.Ordinal)?target:"xl/"+target);
                var strings=new List<string>();var shared=zip.GetEntry("xl/sharedStrings.xml");if(shared!=null){var xml=Xml(shared);var ns=Names(xml);foreach(XmlNode si in xml.SelectNodes("/s:sst/s:si",ns))strings.Add(string.Concat(si.SelectNodes(".//s:t",ns).Cast<XmlNode>().Select(t=>t.InnerText)));}
                var lead=Xml(zip.GetEntry(leadPath));var ln=Names(lead);var values=new Dictionary<string,string>();var occupied=new HashSet<string>();
                foreach(XmlElement cell in lead.SelectNodes("//s:sheetData/s:row/s:c",ln))
                {
                    string address=cell.GetAttribute("r");var value=cell.SelectSingleNode("s:v",ln);string text=value==null?"":value.InnerText;
                    if(cell.GetAttribute("t")=="s")text=strings[int.Parse(text,CultureInfo.InvariantCulture)];else if(cell.GetAttribute("t")=="inlineStr")text=string.Concat(cell.SelectNodes("s:is//s:t",ln).Cast<XmlNode>().Select(t=>t.InnerText));
                    values[address]=text;if(text.Length>0||cell.SelectSingleNode("s:f",ln)!=null)occupied.Add(address);var m=Regex.Match(address,@"^([K-O])(\d+)$");
                    if(m.Success&&int.Parse(m.Groups[2].Value,CultureInfo.InvariantCulture)>=2&&(text.Length>0||cell.SelectSingleNode("s:f",ln)!=null))profile.OldLast=Math.Max(profile.OldLast,int.Parse(m.Groups[2].Value,CultureInfo.InvariantCulture));
                }
                if(columns==null)
                {
                    var aliases=new[]{new[]{"编码","公司编码","序号"},new[]{"简称","公司简称","编码&简称"},new[]{"公司全称","全称","公司名称"},new[]{"完整文件名","完整文件名称","带编码文件名"},new[]{"文件名称","文件名","不含编码文件名"}};
                    for(int i=0;i<5;i++)
                    {
                        var matches=values.Where(v=>Regex.IsMatch(v.Key,@"^[A-Z]+1$")&&aliases[i].Contains(v.Value.Trim(),StringComparer.OrdinalIgnoreCase)).Select(v=>Column(Regex.Match(v.Key,@"^[A-Z]+").Value)).ToArray();
                        if(matches.Length==1)profile.Columns[i]=matches[0];
                        else if(!(i>=3&&values.ContainsKey(i==3?"N1":"O1")&&values[i==3?"N1":"O1"].Trim()=="文件名称"))throw new InvalidOperationException("无法唯一识别 Lead 字段，请手动映射第 "+(i+1)+" 个字段列。");
                    }
                    if(profile.Columns.Distinct().Count()!=5)throw new InvalidOperationException("自动识别列重复，请手动映射。");
                }
                profile.OldLast=2;
                foreach(var pair in values){var m=Regex.Match(pair.Key,@"^([A-Z]+)(\d+)$");if(m.Success&&profile.Columns.Contains(Column(m.Groups[1].Value))&&occupied.Contains(pair.Key))profile.OldLast=Math.Max(profile.OldLast,int.Parse(m.Groups[2].Value));}

                if(lead.SelectSingleNode("/s:worksheet/s:sheetProtection",ln)!=null)throw new InvalidOperationException("模板Lead受保护，请使用可填写的模板。");
                foreach(XmlElement merge in lead.SelectNodes("//s:mergeCell",ln))profile.UnsafeRange(merge.GetAttribute("ref"));
                foreach(XmlElement formula in lead.SelectNodes("//s:f[@t='array']",ln))profile.UnsafeRange(formula.HasAttribute("ref")?formula.GetAttribute("ref"):((XmlElement)formula.ParentNode).GetAttribute("r"));
                foreach(var entry in zip.Entries.Where(e=>e.FullName.StartsWith("xl/worksheets/",StringComparison.Ordinal)&&e.FullName.EndsWith(".xml",StringComparison.Ordinal)))
                {var doc=Xml(entry);var ns=Names(doc);foreach(XmlNode formula in doc.SelectNodes("//s:f",ns))profile.Bound(formula.InnerText);}
                foreach(XmlNode name in book.SelectNodes("//s:definedName",bn))profile.Bound(name.InnerText);
            }
            if(profile.OldLast>100000)throw new InvalidOperationException("模板Lead公司区域过大，请整理模板。");return profile;
        }
        internal ConsolidatedWorkbookPlan Plan(IList<CompanyStructureTask> all,CompanyStructureTask parent,string output)
        {
            var children=CompanyStructureTasks.DirectChildren(all,parent);var plan=new ConsolidatedWorkbookPlan{Columns=(int[])Columns.Clone()};
            plan.Entries.Add(ConsolidatedLeadEntry.From(parent));plan.Entries.AddRange(children.Select(ConsolidatedLeadEntry.From));
            int last=1+plan.Entries.Count;if(last>CapacityLast)throw new InvalidOperationException("该主体含本部共"+children.Count+"家公司，需写至Lead第"+last+"行；模板公式只覆盖至第"+CapacityLast+"行，已阻止以免漏取。");
            plan.ClearLast=Math.Max(OldLast,last);if(firstUnsafeRow<=plan.ClearLast)throw new InvalidOperationException("Lead第"+firstUnsafeRow+"行的公司写入范围含合并或数组公式，请先整理模板。");string file=CompanyWorkbookRules.FileName(parent.FileName);
            plan.Record=new CompanyWorkbookRecord{Row=parent.Row,Name=file,Code=plan.Entries[0].Code,FullName=plan.Entries[0].Name,OutputPath=Path.Combine(Path.GetFullPath(output),file)};
            if(!Directory.Exists(output))throw new InvalidOperationException("请选择输出文件夹。");if(plan.Record.OutputPath.Length>=240)throw new InvalidOperationException("输出路径过长。");
            if(File.Exists(plan.Record.OutputPath)||Directory.Exists(plan.Record.OutputPath))throw new InvalidOperationException("同名输出已存在，不覆盖："+file);return plan;
        }
    }
    internal static class ConsolidatedLeadWriter
    {
        internal static void Write(object book,ConsolidatedWorkbookPlan plan)
        {
            using(var bag=new ComBag())
            {
                dynamic sheets=bag.Own(((dynamic)book).Worksheets),lead=bag.Own(sheets.Item("Lead"));
                string mapped=string.Join(",",plan.Columns.Select(c=>CellAddress.ColumnLetters(c)+"2:"+CellAddress.ColumnLetters(c)+plan.ClearLast));
                dynamic range=bag.Own(lead.Range[mapped]);
                object merged=range.MergeCells,array=range.HasArray;
                if(Convert.ToBoolean(lead.ProtectContents)||merged==null||merged==DBNull.Value||Convert.ToBoolean(merged)||array==null||array==DBNull.Value||Convert.ToBoolean(array))throw new InvalidOperationException("Lead公司写入范围受保护，或包含合并、数组单元格。");
                try{object spill=range.HasSpill;if(spill==null||spill==DBNull.Value||Convert.ToBoolean(spill))throw new InvalidOperationException("Lead写入范围包含溢出公式。");}
                catch(Microsoft.CSharp.RuntimeBinder.RuntimeBinderException){}
                catch(System.Runtime.InteropServices.COMException error){if(error.ErrorCode!=unchecked((int)0x80020003)&&error.ErrorCode!=unchecked((int)0x80020006))throw;}
                range.ClearContents();
                for(int i=0;i<plan.Entries.Count;i++)
                {
                    var item=plan.Entries[i];int row=i+2;
                    CompanyWorkbookHost.WriteText(bag.Own(lead.Range[CellAddress.ColumnLetters(plan.Columns[0])+row]),item.Code);CompanyWorkbookHost.WriteText(bag.Own(lead.Range[CellAddress.ColumnLetters(plan.Columns[2])+row]),item.Name);CompanyWorkbookHost.WriteText(bag.Own(lead.Range[CellAddress.ColumnLetters(plan.Columns[4])+row]),item.FilePart);
                    ((dynamic)bag.Own(lead.Range[CellAddress.ColumnLetters(plan.Columns[1])+row])).Formula="="+CellAddress.ColumnLetters(plan.Columns[0])+row+"&"+CellAddress.ColumnLetters(plan.Columns[2])+row;((dynamic)bag.Own(lead.Range[CellAddress.ColumnLetters(plan.Columns[3])+row])).Formula="="+CellAddress.ColumnLetters(plan.Columns[0])+row+"&"+CellAddress.ColumnLetters(plan.Columns[4])+row;
                }
            }
            Verify(book,plan);
        }
        internal static void Verify(object book,ConsolidatedWorkbookPlan plan)
        {
            using(var bag=new ComBag())
            {
                dynamic sheets=bag.Own(((dynamic)book).Worksheets),lead=bag.Own(sheets.Item("Lead")),range=bag.Own(lead.Range["A1"]);
                for(int i=0;i<plan.ClearLast-1;i++)for(int c=0;c<5;c++)
                {
                    string expected="";if(i<plan.Entries.Count){var item=plan.Entries[i];int row=i+2;expected=c==0?item.Code:c==1?"="+CellAddress.ColumnLetters(plan.Columns[0])+row+"&"+CellAddress.ColumnLetters(plan.Columns[2])+row:c==2?item.Name:c==3?"="+CellAddress.ColumnLetters(plan.Columns[0])+row+"&"+CellAddress.ColumnLetters(plan.Columns[4])+row:item.FilePart;}
                    if(CompanyWorkbookRules.Text(((dynamic)bag.Own(lead.Cells[i+2,plan.Columns[c]])).Formula)!=expected)throw new InvalidOperationException("合并Lead写入或旧名单清理核验失败：第"+(i+2)+"行第"+(c+11)+"列。");
                    if(i<plan.Entries.Count){dynamic cell=bag.Own(lead.Cells[i+2,plan.Columns[c]]);if(Convert.ToBoolean(cell.HasFormula)!=(c==1||c==3))throw new InvalidOperationException("Lead文本与公式类型核验失败。");}
                }
            }
        }
    }
}
