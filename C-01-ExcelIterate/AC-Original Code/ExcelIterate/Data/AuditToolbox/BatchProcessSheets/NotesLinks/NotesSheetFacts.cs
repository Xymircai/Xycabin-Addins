using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace Excel_Iterate
{
    // Read-only package facts supplement the existing bounded sheet snapshot.
    internal sealed class NotesSheetFacts
    {
        internal MigrationSheetSnapshot Snapshot;
        internal double StandardWidth,StandardHeight;
        internal readonly HashSet<int> HiddenRows=new HashSet<int>(),HiddenColumns=new HashSet<int>();
        internal bool Protected,Tables,HiddenSheet;
        internal readonly List<MigrationArea> Arrays=new List<MigrationArea>();
        internal readonly Dictionary<long,string> Contents=new Dictionary<long,string>();
        private readonly Dictionary<long,int> cellStyles=new Dictionary<long,int>();
        private readonly Dictionary<int,int> rowStyles=new Dictionary<int,int>();
        private readonly List<int[]> columnStyles=new List<int[]>();
        private readonly Dictionary<int,string> appearances=new Dictionary<int,string>(),formats=new Dictionary<int,string>();
        private string geometry;
        private static XmlDocument Read(ZipArchiveEntry entry)
        {if(entry==null)throw new InvalidOperationException("工作簿缺少必要结构。");var doc=new XmlDocument{XmlResolver=null};using(var stream=entry.Open())using(var xml=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))doc.Load(xml);return doc;}
        private static XmlNode Node(XmlNode node,string name){return node.SelectSingleNode("*[local-name()='"+name+"']");}
        private static int Int(XmlElement node,string name,int fallback){int n;return int.TryParse(node.GetAttribute(name),out n)?n:fallback;}
        private static string Canonical(XmlNode node)
        {
            if(node==null)return "";
            var element=node as XmlElement;if(element==null)return node.Value??"";
            string attrs=string.Join(";",element.Attributes.Cast<XmlAttribute>().Where(a=>a.Prefix!="xmlns"&&a.Name!="xmlns").OrderBy(a=>a.LocalName).Select(a=>a.LocalName+"="+a.Value));
            return element.LocalName+"["+attrs+"]"+string.Concat(element.ChildNodes.Cast<XmlNode>().Select(Canonical));
        }
        internal static NotesSheetFacts Load(string path,string sheet)
        {
            var result=new NotesSheetFacts{Snapshot=MigrationSheetSnapshot.Load(path,sheet)};
            using(var stream=File.Open(path,FileMode.Open,FileAccess.Read,FileShare.ReadWrite))using(var zip=new ZipArchive(stream,ZipArchiveMode.Read))
            {
                var book=Read(zip.GetEntry("xl/workbook.xml"));var item=book.SelectNodes("//*[local-name()='sheet']").Cast<XmlElement>().Single(n=>string.Equals(n.GetAttribute("name"),sheet,StringComparison.OrdinalIgnoreCase));
                result.HiddenSheet=item.GetAttribute("state")=="hidden"||item.GetAttribute("state")=="veryHidden";
                string id=item.Attributes.Cast<XmlAttribute>().First(a=>a.LocalName=="id").Value;
                var relationships=Read(zip.GetEntry("xl/_rels/workbook.xml.rels"));var link=relationships.DocumentElement.ChildNodes.OfType<XmlElement>().Single(n=>n.GetAttribute("Id")==id);
                string part=Uri.UnescapeDataString(new Uri(new Uri("http://package/xl/workbook.xml"),link.GetAttribute("Target")).AbsolutePath.TrimStart('/'));
                var doc=Read(zip.GetEntry(part));result.Protected=doc.SelectSingleNode("//*[local-name()='sheetProtection']")!=null;
                result.Tables=doc.SelectSingleNode("//*[local-name()='tablePart' or local-name()='pivotTablePart']")!=null;
                var defaults=doc.SelectSingleNode("//*[local-name()='sheetFormatPr']") as XmlElement;
                result.geometry=string.Join(";",new[]{"zeroHeight","thickTop","thickBottom"}.Select(n=>defaults!=null&&(defaults.GetAttribute(n)=="1"||defaults.GetAttribute(n)=="true")?"1":"0"));
                foreach(XmlElement row in doc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']"))
                {
                    int number=Int(row,"r",0);if(row.HasAttribute("s"))result.rowStyles[number]=Int(row,"s",0);
                    if(row.GetAttribute("hidden")=="1")result.HiddenRows.Add(number);
                    double height;if(double.TryParse(row.GetAttribute("ht"),System.Globalization.NumberStyles.Float,System.Globalization.CultureInfo.InvariantCulture,out height))result.Snapshot.Heights[number]=height;
                    foreach(XmlElement cell in row.ChildNodes.OfType<XmlElement>().Where(n=>n.LocalName=="c"))
                    {
                        var area=MigrationArea.Parse(cell.GetAttribute("r"));long key=MigrationSheetSnapshot.Key(area.Row,area.Column);
                        if(cell.HasAttribute("s"))result.cellStyles[key]=Int(cell,"s",0);
                        var formula=Node(cell,"f");string type=cell.GetAttribute("t"),value=result.Snapshot.Value(area.Row,area.Column);
                        result.Contents[key]=formula==null?(value.Length==0?"":(type=="s"||type=="inlineStr"||type=="str"?"text:":"value:")+value):"formula:"+formula.InnerText;
                        if(formula!=null&&((XmlElement)formula).GetAttribute("t")=="array")result.Arrays.Add(MigrationArea.Parse(((XmlElement)formula).GetAttribute("ref").Length>0?((XmlElement)formula).GetAttribute("ref"):cell.GetAttribute("r")));
                        else if(cell.HasAttribute("cm"))result.Arrays.Add(area);
                    }
                }
                foreach(XmlElement col in doc.SelectNodes("//*[local-name()='cols']/*[local-name()='col']"))
                {if(col.HasAttribute("style"))result.columnStyles.Add(new[]{Int(col,"min",1),Int(col,"max",1),Int(col,"style",0)});if(col.GetAttribute("hidden")=="1")for(int c=Int(col,"min",1);c<=Math.Min(16384,Int(col,"max",1));c++)result.HiddenColumns.Add(c);}
                var styles=Read(zip.GetEntry("xl/styles.xml"));var fonts=Node(styles.DocumentElement,"fonts").ChildNodes.OfType<XmlElement>().ToList();var fills=Node(styles.DocumentElement,"fills").ChildNodes.OfType<XmlElement>().ToList();var borders=Node(styles.DocumentElement,"borders").ChildNodes.OfType<XmlElement>().ToList();
                var custom=new Dictionary<int,string>();var num=Node(styles.DocumentElement,"numFmts");if(num!=null)foreach(XmlElement n in num.ChildNodes.OfType<XmlElement>())custom[Int(n,"numFmtId",0)]=n.GetAttribute("formatCode");
                int index=0;
                foreach(XmlElement xf in Node(styles.DocumentElement,"cellXfs").ChildNodes.OfType<XmlElement>())
                {
                    int font=Int(xf,"fontId",0),fill=Int(xf,"fillId",0),border=Int(xf,"borderId",0),format=Int(xf,"numFmtId",0);
                    string attributes=string.Join(";",xf.Attributes.Cast<XmlAttribute>().Where(a=>!new[]{"fontId","fillId","borderId","numFmtId","applyNumberFormat","quotePrefix"}.Contains(a.LocalName)).OrderBy(a=>a.LocalName).Select(a=>a.LocalName+"="+a.Value));
                    result.appearances[index]=attributes+Canonical(fonts[font])+Canonical(fills[fill])+Canonical(borders[border])+string.Concat(xf.ChildNodes.Cast<XmlNode>().Select(Canonical));
                    string code;result.formats[index]=custom.TryGetValue(format,out code)?code:"builtin:"+format;index++;
                }
            }
            return result;
        }
        internal int Style(int row,int col)
        {int n;if(cellStyles.TryGetValue(MigrationSheetSnapshot.Key(row,col),out n))return n;if(rowStyles.TryGetValue(row,out n))return n;var column=columnStyles.LastOrDefault(c=>col>=c[0]&&col<=c[1]);return column==null?0:column[2];}
        internal string Content(int row,int col){string value;return Contents.TryGetValue(MigrationSheetSnapshot.Key(row,col),out value)?value:"";}
        internal bool Occupied(int row,int col){return Content(row,col).Length>0;}
        internal double Height(int row){double value;return Snapshot.Heights.TryGetValue(row,out value)?value:StandardHeight;}
        internal double Width(int col){double value;return Snapshot.Widths.TryGetValue(col,out value)?value:StandardWidth;}
        internal void ValidateNumberFormats(MigrationArea area)
        {
            var seen=new HashSet<int>();for(int r=area.Row;r<area.Row+area.Rows;r++)for(int c=area.Column;c<area.Column+area.Columns;c++)seen.Add(Style(r,c));
            foreach(int style in seen){string format=formats[style];if(!format.StartsWith("builtin:",StringComparison.Ordinal))NotesLinkRules.ZeroFormat(format);}
        }
        private static Dictionary<int,List<int[]>> Intervals(IEnumerable<MigrationArea> areas)
        {
            var rows=new Dictionary<int,List<int[]>>();
            foreach(var area in areas)for(int r=area.Row;r<area.Row+area.Rows;r++)
            {List<int[]> list;if(!rows.TryGetValue(r,out list))rows[r]=list=new List<int[]>();list.Add(new[]{area.Column,area.Column+area.Columns});}
            foreach(var list in rows.Values)list.Sort((a,b)=>a[0].CompareTo(b[0]));return rows;
        }
        private static bool Contains(Dictionary<int,List<int[]>> rows,int row,int column)
        {
            List<int[]> list;if(!rows.TryGetValue(row,out list))return false;int low=0,high=list.Count-1;
            while(low<=high){int mid=(low+high)/2;if(column<list[mid][0])high=mid-1;else if(column>=list[mid][1])low=mid+1;else return true;}return false;
        }
        internal void VerifyAgainst(NotesSheetFacts original,IList<NotesLinkTask> tasks)
        {
            if(geometry!=original.geometry||Math.Abs(StandardHeight-original.StandardHeight)>0.01||Math.Abs(StandardWidth-original.StandardWidth)>0.01||!HiddenRows.SetEquals(original.HiddenRows)||!HiddenColumns.SetEquals(original.HiddenColumns))throw new InvalidOperationException("Notes默认尺寸或隐藏状态发生变化，未交付。");
            foreach(int c in Snapshot.Widths.Keys.Concat(original.Snapshot.Widths.Keys))if(Math.Abs(Width(c)-original.Width(c))>0.01)throw new InvalidOperationException("Notes列宽变化："+CellAddress.ColumnLetters(c)+"（"+original.Width(c)+" → "+Width(c)+"），未交付。");
            foreach(int r in Snapshot.Heights.Keys.Concat(original.Snapshot.Heights.Keys))if(Math.Abs(Height(r)-original.Height(r))>0.01)throw new InvalidOperationException("Notes行高变化："+r+"（"+original.Height(r)+" → "+Height(r)+"），未交付。");
            if(!Snapshot.Merges.Select(x=>x.Address).OrderBy(x=>x).SequenceEqual(original.Snapshot.Merges.Select(x=>x.Address).OrderBy(x=>x)))throw new InvalidOperationException("Notes合并结构发生变化，未交付。");
            var keys=new HashSet<long>(Contents.Keys.Concat(original.Contents.Keys).Concat(cellStyles.Keys).Concat(original.cellStyles.Keys));
            var writes=Intervals(tasks.SelectMany(t=>t.Blocks.Concat(t.Anchors)));var authorization=Intervals(tasks.Select(t=>t.Target));
            foreach(long key in keys)
            {
                int row=(int)(key/20000),col=(int)(key%20000);bool writable=Contains(writes,row,col);bool authorized=Contains(authorization,row,col);
                if(!writable&&Content(row,col)!=original.Content(row,col))throw new InvalidOperationException("未授权单元格内容变化："+CellAddress.Name(row,col));
                int old=original.Style(row,col),now=Style(row,col);
                if(appearances[now]!=original.appearances[old]||(!authorized&&formats[now]!=original.formats[old]))throw new InvalidOperationException("Notes样式变化超出授权："+CellAddress.Name(row,col));
            }
        }
    }
    internal static class NotesLinkPlanner
    {
        internal static NotesLinkPlan Build(NotesLinkConfig input,IList<NotesLinkTask> rows,Action<string> log,Func<bool> cancelled)
        {
            var cfg=input.Copy();var source=cfg.Validate();string hash=DataMigrationHost.Digest(cfg.Parent);
            string target=NotesLinkRules.Resolve(DataMigrationHost.PackageSheets(cfg.Parent),cfg.TargetSheet);
            if(!NotesLinkRules.ExactOrAlias(target,cfg.TargetSheet))throw new InvalidOperationException("请在合并Notes表下拉框明确选择："+target);
            var facts=NotesSheetFacts.Load(cfg.Parent,target);if(facts.Protected||facts.Tables)throw new InvalidOperationException("目标Notes受保护或包含表格／透视表对象，不能安全写入。");
            var plan=new NotesLinkPlan{Config=cfg,ParentHash=hash,TargetSheet=target,Source=source,Original=facts};
            plan.Cleanup=FindUnusedTemplate(facts,source,cfg,rows);
            if(plan.Cleanup!=null&&log!=null)log("将清理右侧多余模板："+plan.Cleanup.Address+"；保留列位置及左侧合计公式。");
            long estimatedCells=facts.Snapshot.Values.Count;
            if(rows.Count(r=>r.Selected)==0)throw new InvalidOperationException("请勾选至少一家子公司。");
            if((long)rows.Count(r=>r.Selected)*source.Rows*source.Columns>5000000)throw new InvalidOperationException("本批超过500万格，请分批勾选公司。");
            var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach(var row in rows.Where(r=>r.Selected))
            {
                if(cancelled!=null&&cancelled())throw new OperationCanceledException("已取消预检。");var task=row.Copy();task.Blocks.Clear();task.Anchors.Clear();task.Skipped=task.Occupied=0;
                try
                {
                    task.Target=NotesLinkRules.Target(source,cfg.FirstStart,cfg.ColumnSpacing,task.Slot);
                    NotesLinkRules.RequireWorkbook(task.File);task.File=Path.GetFullPath(task.File);
                    if(!seen.Add(task.File)||string.Equals(task.File,Path.GetFullPath(cfg.Parent),StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("来源文件重复或指向父公司自身。");
                    task.Hash=DataMigrationHost.Digest(task.File);task.SheetNames=DataMigrationHost.PackageSheets(task.File);
                    string requested=string.IsNullOrWhiteSpace(task.Sheet)?cfg.SourceSheet:task.Sheet;
                    task.Sheet=NotesLinkRules.Resolve(task.SheetNames,requested);bool fuzzy=task.NameNeedsConfirmation||!NotesLinkRules.ExactOrAlias(task.Sheet,requested);task.NameNeedsConfirmation=fuzzy;
                    var sourceFacts=NotesSheetFacts.Load(task.File,task.Sheet);var profile=sourceFacts.Snapshot;task.SourceMerges=profile.Merges;
                    bool feature=profile.Merges.Any(m=>m.Column==source.Column&&m.Columns==source.Columns);
                    if((fuzzy||!feature)&&!task.Confirmed)throw new InvalidOperationException((fuzzy?"表名为模糊匹配；":"")+(!feature?"未找到横跨来源列的大型合并区域；":"")+"请检查后勾选“已核对特殊表格”。");
                    task.Target=NotesLinkRules.Target(source,cfg.FirstStart,cfg.ColumnSpacing,task.Slot);
                    if(facts.Arrays.Any(a=>NotesLinkRules.Intersects(a,task.Target)))throw new InvalidOperationException("目标含数组／溢出公式，不能安全覆盖。");
                    NotesLinkRules.MapMerges(task,source,facts,sourceFacts);
                    task.NativeCompatible=cfg.NativePaste&&!facts.HiddenSheet&&!sourceFacts.HiddenSheet&&NotesLinkRules.CanPasteWhole(source,task,facts);
                    foreach(var block in (task.NativeCompatible?new[]{task.Target}:task.Blocks.Concat(task.Anchors)))for(int rowNumber=block.Row;rowNumber<block.Row+block.Rows;rowNumber++)for(int col=block.Column;col<block.Column+block.Columns;col++)if(facts.Snapshot.Value(rowNumber,col).Length==0)estimatedCells++;
                    if(estimatedCells>MigrationSheetSnapshot.CellLimit)throw new InvalidOperationException("预计结果Notes超过读取器60万非空单元格上限，请缩小行数或减少任务；本批未写入。");
                    for(int r=task.Target.Row;r<task.Target.Row+task.Target.Rows;r++)for(int c=task.Target.Column;c<task.Target.Column+task.Target.Columns;c++)if(facts.Occupied(r,c))task.Occupied++;
                    if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在预检期间变化。");
                    task.Status="通过；覆盖"+task.Occupied+"格；合并内部"+task.Skipped+"格";
                }
                catch(Exception error){task.Status="阻止："+error.Message;}
                plan.Tasks.Add(task);if(log!=null)log(task.Code+task.Name+"｜"+task.Status);
            }
            if(DataMigrationHost.Digest(cfg.Parent)!=hash)throw new InvalidOperationException("父公司文件在预检期间变化，请重新预检。");
            return plan;
        }
        internal static MigrationArea FindUnusedTemplate(NotesSheetFacts facts,MigrationArea source,NotesLinkConfig cfg,IList<NotesLinkTask> rows)
        {
            if(rows.Count==0)return null;
            var first=NotesLinkRules.Target(source,cfg.FirstStart,cfg.ColumnSpacing,0);
            // Only a contiguous series of full-width, single-row company headers
            // identifies template slots. Never use UsedRange as a deletion boundary.
            int slots=0;
            while(facts.Snapshot.Merges.Any(m=>m.Row==first.Row&&m.Rows==1&&m.Column==first.Column+slots*cfg.ColumnSpacing&&m.Columns==source.Columns))slots++;
            int actual=rows.Max(t=>t.Slot)+1;
            if(slots<=actual)return null;
            var area=new MigrationArea{Row=first.Row,Column=first.Column+actual*cfg.ColumnSpacing,Rows=Math.Max(source.Rows,facts.Snapshot.MaxRow-first.Row+1),Columns=(slots-actual-1)*cfg.ColumnSpacing+source.Columns};
            if(facts.Arrays.Any(a=>NotesLinkRules.Intersects(a,area))||facts.Snapshot.Merges.Any(m=>NotesLinkRules.Intersects(m,area)&&!(m.Row>=area.Row&&m.Column>=area.Column&&m.Row+m.Rows<=area.Row+area.Rows&&m.Column+m.Columns<=area.Column+area.Columns)))throw new InvalidOperationException("右侧模板清理区域包含数组公式或跨界合并，无法安全清理。");
            return area;
        }
    }
}
