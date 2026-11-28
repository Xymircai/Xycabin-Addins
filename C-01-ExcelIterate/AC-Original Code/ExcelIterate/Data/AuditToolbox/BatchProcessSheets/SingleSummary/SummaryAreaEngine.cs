using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class SummaryLeadRow
    {
        internal int Row, CodeColumn, NameColumn;
    }
    internal sealed class SummaryAreaSettings
    {
        internal readonly List<SummaryAreaPlan> Template = new List<SummaryAreaPlan>(), Lead = new List<SummaryAreaPlan>();
        internal SummaryAreaPlan Paste;
        internal string Folder;
        internal bool CheckFiles, HeadersEnabled, FormulaHeaders, RepeatNames, VariableHeaders;
        internal SummaryAreaPlan CompanyHeader, YearSource, YearTarget, RepeatSource, RepeatTarget;
        internal static string Encode(IEnumerable<SummaryAreaPlan> parts)
        { return string.Join(";", parts.Select(p => string.Join(",", new[] { p.Row, p.Column, p.Rows, p.Columns }.Select(n => n.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray())).ToArray()); }
        internal static List<SummaryAreaPlan> Decode(string text, string sheet)
        {
            if (string.IsNullOrEmpty(sheet) || string.IsNullOrEmpty(text)) throw new InvalidOperationException("预设缺少工作表或选区，请重新选择后保存。");
            var result = new List<SummaryAreaPlan>();
            foreach (string token in text.Split(';'))
            {
                string[] fields = token.Split(','); var n = new int[4];
                if (fields.Length != 4) throw new InvalidOperationException("预设区域格式无效。");
                for (int i = 0; i < 4; i++) if (!int.TryParse(fields[i], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out n[i]) || n[i] < 1) throw new InvalidOperationException("预设区域坐标无效。");
                if ((long)n[0] + n[2] - 1 > 1048576 || (long)n[1] + n[3] - 1 > 16384 || (long)n[2] * n[3] > SummaryAreaEngine.CellLimit) throw new InvalidOperationException("预设区域越界或过大。");
                var area = new SummaryAreaPlan { Sheet = sheet, Row = n[0], Column = n[1], Rows = n[2], Columns = n[3] }; area.Address = SummaryAreaEngine.Address(area); result.Add(area);
            }
            return result;
        }
        internal static SummaryAreaPlan OptionalHeader(Dictionary<string,string> values, string key)
        {
            if (!values.ContainsKey(key) || string.IsNullOrEmpty(values[key])) return null;
            if (!values.ContainsKey(key + "Sheet")) throw new InvalidOperationException("表头预设缺少工作表：" + key);
            var parts = Decode(values[key], values[key + "Sheet"]);
            if(parts.Count != 1) throw new InvalidOperationException("表头预设只接受一个连续区域。");
            return parts[0];
        }
        internal static SummaryAreaSettings Read(BatchPreset preset)
        {
            BatchPresetStore.Validate(preset, "summary"); var values = preset.Values;
            if (!values.ContainsKey("mode") || values["mode"] != "area-mvp") throw new InvalidOperationException("这是旧整列参数预设，已保留在库中；请重新指定已编制区域和粘贴起点后另存区域预设，不会覆盖旧预设。");
            foreach (string key in new[] { "templateSheet", "templateAreas", "pasteSheet", "pasteArea", "leadSheet", "leadAreas", "projectFolder", "checkFiles" }) if (!values.ContainsKey(key)) throw new InvalidOperationException("区域预设缺少参数：" + key);
            var result = new SummaryAreaSettings { Folder = values["projectFolder"] };
            if (!bool.TryParse(values["checkFiles"], out result.CheckFiles)) throw new InvalidOperationException("来源文件检查选项无效。");
            result.Template.AddRange(Decode(values["templateAreas"], values["templateSheet"])); result.Lead.AddRange(Decode(values["leadAreas"], values["leadSheet"]));
            var paste = Decode(values["pasteArea"], values["pasteSheet"]);
            if (paste.Count != 1 || paste[0].Rows != 1 || paste[0].Columns != 1 || values["pasteSheet"] != values["templateSheet"]) throw new InvalidOperationException("预设粘贴起点须为已编制区域同表的一个单元格。");
            if (values.ContainsKey("headersEnabled") && !bool.TryParse(values["headersEnabled"], out result.HeadersEnabled)) throw new InvalidOperationException("表头开关无效。");
            if(values.ContainsKey("headersFormula")&&!bool.TryParse(values["headersFormula"],out result.FormulaHeaders))throw new InvalidOperationException("表头公式选项无效。");
            if(values.ContainsKey("headersRepeat")&&!bool.TryParse(values["headersRepeat"],out result.RepeatNames))throw new InvalidOperationException("名称重复选项无效。");
            if(values.ContainsKey("variableHeaders")&&!bool.TryParse(values["variableHeaders"],out result.VariableHeaders))throw new InvalidOperationException("变量表头选项无效。");
            result.RepeatSource=OptionalHeader(values,"repeatSource");result.RepeatTarget=OptionalHeader(values,"repeatTarget");
            result.CompanyHeader = OptionalHeader(values, "companyHeader"); result.YearSource = OptionalHeader(values, "yearSource"); result.YearTarget = OptionalHeader(values, "yearTarget");
            if (result.HeadersEnabled && !result.VariableHeaders && (result.CompanyHeader == null || result.YearSource == null || result.YearTarget == null)) throw new InvalidOperationException("启用表头的预设缺少选区。");
            if ((result.CompanyHeader != null && (result.CompanyHeader.Rows != 1 || result.CompanyHeader.Columns != 1 || result.CompanyHeader.Sheet != values["templateSheet"])) || (result.YearTarget != null && (result.YearTarget.Rows != 1 || result.YearTarget.Columns != 1 || result.YearTarget.Sheet != values["templateSheet"])) || (result.YearSource != null && result.YearSource.Rows != 1 && result.YearSource.Columns != 1)) throw new InvalidOperationException("表头预设的目标或年度来源形状无效。");
            result.Paste = paste[0]; SummaryAreaEngine.LeadRows(result.Lead); SummaryAreaEngine.Targets(result.Template, result.Paste.Row, result.Paste.Column, 1);
            int row = result.Template.Min(p => p.Row), column = result.Template.Min(p => p.Column); foreach (var part in result.Template) { part.RowOffset = part.Row - row; part.ColumnOffset = part.Column - column; }
            return result;
        }
    }
    internal sealed class SummaryHeaderArray
    {
        internal SummaryAreaPlan Area;
        internal string Formula;
        internal bool Dynamic;
    }
    internal static class SummaryPeriodOrder
    {
        internal static DateTime Parse(object value,string display,bool date1904)
        {
            if(value is DateTime)return (DateTime)value;
            string text=(display??"").Trim();
            var year=Regex.Match(text,@"^(\d{4})(?:年度|年)?$");
            if(year.Success){int y=int.Parse(year.Groups[1].Value);if(y>=1900&&y<=9999)return new DateTime(y,1,1);}
            var date=Regex.Match(text,@"^(\d{4})[./年-](\d{1,2})[./月-](\d{1,2})日?$");
            if(date.Success){try{return new DateTime(int.Parse(date.Groups[1].Value),int.Parse(date.Groups[2].Value),int.Parse(date.Groups[3].Value));}catch(ArgumentOutOfRangeException){}}
            if(value is double||value is float||value is int||value is decimal)
            {double n=Convert.ToDouble(value);if(n>=1&&n<2958466){try{return DateTime.FromOADate(n+(date1904?1462:0));}catch(ArgumentException){}}}
            throw new InvalidOperationException("无法识别期间："+text+"；请选择年度或明确的年月日。");
        }
        internal static List<int> Descending(IList<DateTime> values)
        {if(values.Count==0||values.Distinct().Count()!=values.Count)throw new InvalidOperationException("期间为空或包含重复时间，不能唯一排序。");return Enumerable.Range(0,values.Count).OrderByDescending(i=>values[i]).ToList();}
    }
    internal static class SummaryHeaderFormula
    {
        internal static string Reference(SummaryAreaPlan area)
        { return "'"+area.Sheet.Replace("'","''")+"'!$"+CellAddress.ColumnLetters(area.Column)+"$"+area.Row+":$"+CellAddress.ColumnLetters(area.Column+area.Columns-1)+"$"+(area.Row+area.Rows-1); }
        internal static string Cell(string sheet,int row,int column)
        {return "='"+sheet.Replace("'","''")+"'!$"+CellAddress.ColumnLetters(column)+"$"+row;}
        internal static string Years(SummaryAreaPlan source)
        {
            if(source.Rows!=1&&source.Columns!=1)throw new InvalidOperationException("年度来源必须为单列或单行连续区域。");
            return source.Rows>1?"=TRANSPOSE("+Reference(source)+")":"="+Reference(source);
        }
    }
    internal sealed class SummaryAreaCell
    {
        internal int Row, Column;
        internal string Text, R1C1;
        internal bool Formula, Literal;
        internal SummaryExistingFormula Existing;
    }
    internal sealed class SummaryReferenceSource
    {
        internal string Key, File, Locations, Reason, Signature, Decision;
    }
    // Reuse the existing span scanner; literals are masked and never rewritten.
    internal sealed class SummaryExistingFormula
    {
        internal string Original;
        internal readonly List<RefSpan> Spans = new List<RefSpan>();
        internal readonly Dictionary<string, bool> Decisions = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        internal static string Prefix(string formula, RefSpan span)
        {
            string raw = formula.Substring(span.PrefixStart, span.PrefixLength);
            return raw.StartsWith("'", StringComparison.Ordinal) ? raw.Substring(1, raw.Length-2).Replace("''", "'") : raw;
        }
        internal static string Source(string formula, RefSpan span)
        {
            string prefix=Prefix(formula,span);
            return span.Kind==RefKind.CrossSheet ? "本工作簿!"+prefix : prefix.Substring(0,prefix.LastIndexOf(']')+1);
        }
        internal static string File(string formula, RefSpan span)
        {
            if(span.Kind!=RefKind.CrossBook)return "";
            string prefix=Prefix(formula,span);int open=prefix.LastIndexOf('['),close=prefix.LastIndexOf(']');
            if(open<0||close<=open+1)throw new InvalidOperationException("无法识别外链文件。");
            return prefix.Substring(open+1,close-open-1);
        }
        internal static SummaryExistingFormula Read(string formula)
        {
            var result=new SummaryExistingFormula{Original=formula};result.Spans.AddRange(FormulaRefRewriter.Scan(formula));
            var mask=new System.Text.StringBuilder(formula);
            foreach(var span in result.Spans.OrderByDescending(x=>x.Start))
            {
                if(span.Kind==RefKind.Local)continue;
                if(span.Kind==RefKind.CrossBook)File(formula,span);
                // Reject unsupported name/3D/structured/dynamic references via residual validation.
                string prefix=Prefix(formula,span);if(span.Kind==RefKind.CrossSheet&&prefix.Contains(":"))throw new InvalidOperationException("不支持三维工作表引用。");
                mask.Remove(span.PrefixStart,span.Start+span.Length-span.PrefixStart);mask.Insert(span.PrefixStart,"0");
            }
            SummaryAreaEngine.NativeFormula(mask.ToString());return result;
        }
        internal string Render(string folder,string company,int rows,int columns)
        {
            var result=new System.Text.StringBuilder(Original);
            foreach(var span in Spans.OrderByDescending(x=>x.Start))
            {
                int start=span.Start,length=span.Length;string replacement;
                if(span.Kind==RefKind.Local)
                {
                    replacement=Cell(span.Column1,span.Row1,span.Column1Absolute,span.Row1Absolute,span.IsWholeColumn,rows,columns);
                    if(span.IsRange)replacement+=":"+Cell(span.Column2,span.Row2,span.Column2Absolute,span.Row2Absolute,span.IsWholeColumn,rows,columns);
                }
                else
                {
                    bool replace;if(!Decisions.TryGetValue(Source(Original,span),out replace))throw new InvalidOperationException("引用来源尚未确认："+Source(Original,span));
                    if(!replace)continue;
                    if(span.Kind!=RefKind.CrossBook)throw new InvalidOperationException("本工作簿跨表来源只能保留原引用。");
                    string prefix=Prefix(Original,span);string sheet=prefix.Substring(prefix.LastIndexOf(']')+1);
                    // External validates paths, filenames and worksheet names. Keep original address verbatim.
                    string sample=SummaryTextReference.External(folder,company,"'"+sheet.Replace("'","''")+"'!A1");
                    replacement=sample.Substring(1,sample.LastIndexOf('!'))+Original.Substring(span.Start,span.Length);
                    start=span.PrefixStart;length=span.Start+span.Length-start;
                }
                result.Remove(start,length);result.Insert(start,replacement);
            }
            return result.ToString();
        }
        private static string Cell(int column,int row,bool absoluteColumn,bool absoluteRow,bool wholeColumn,int rows,int columns)
        {
            long c=absoluteColumn?column:(long)column+columns,r=wholeColumn?1:absoluteRow?row:(long)row+rows;
            if(c<1||c>16384||r<1||r>1048576)throw new InvalidOperationException("本表公式平移后引用越界。");
            return (absoluteColumn?"$":"")+CellAddress.ColumnLetters((int)c)+(wholeColumn?"":(absoluteRow?"$":"")+r.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
    internal static class SummaryAreaEngine
    {
        internal const int CellLimit = 100000;
        // Pair actual row identities rather than zipping flattened Areas: a
        // missing cell in either column must never shift subsequent companies.
        internal static List<SummaryLeadRow> LeadRows(IList<SummaryAreaPlan> areas)
        {
            if (areas.Count == 0) throw new InvalidOperationException("请选择 Lead 公司字段。");
            var columns = new SortedSet<int>(); var rows = new Dictionary<int, HashSet<int>>(); var order = new List<int>(); long count = 0;
            foreach (var area in areas)
            {
                if (area.Sheet != areas[0].Sheet) throw new InvalidOperationException("Lead 选区必须来自同一工作表。");
                if (area.Rows < 1 || area.Columns < 1 || area.Columns > 2 || area.Row < 1 || area.Column < 1 || (long)area.Row + area.Rows - 1 > 1048576 || (long)area.Column + area.Columns - 1 > 16384)
                    throw new InvalidOperationException("Lead 每个 Area 只能包含一列或两列的有效数据区域。");
                count += (long)area.Rows * area.Columns;
                if (count > CellLimit) throw new InvalidOperationException("Lead 选区最多十万单元格，请选择实际数据行。");
                for (int c = area.Column; c < area.Column + area.Columns; c++) columns.Add(c);
                if (columns.Count > 2) throw new InvalidOperationException("Lead 只能选择两列：左列编码简称，右列公司全称。");
                for (int r = area.Row; r < area.Row + area.Rows; r++)
                {
                    HashSet<int> selected;
                    if (!rows.TryGetValue(r, out selected)) { selected = new HashSet<int>(); rows.Add(r, selected); order.Add(r); }
                    for (int c = area.Column; c < area.Column + area.Columns; c++)
                        if (!selected.Add(c)) throw new InvalidOperationException("Lead 第 " + r + " 行有重复或重叠选区。");
                }
            }
            if (columns.Count < 1 || columns.Count > 2) throw new InvalidOperationException("请选择公司名称一列。");
            var result = new List<SummaryLeadRow>();
            foreach (int row in order)
            {
                if (rows[row].Count != columns.Count) throw new InvalidOperationException("Lead 第 " + row + " 行未同时选中编码和全称，请按同一行补齐两列。");
                result.Add(new SummaryLeadRow { Row = row, CodeColumn = columns.Min, NameColumn = columns.Max });
            }
            return result;
        }
        internal static bool Intersects(SummaryAreaPlan a, SummaryAreaPlan b)
        { return a.Row <= (long)b.Row + b.Rows - 1 && b.Row <= (long)a.Row + a.Rows - 1 && a.Column <= (long)b.Column + b.Columns - 1 && b.Column <= (long)a.Column + a.Columns - 1; }
        internal static List<SummaryAreaPlan> Targets(IList<SummaryAreaPlan> areas, int pasteRow, int pasteColumn, int companies)
        {
            if (areas.Count == 0 || companies < 1) throw new InvalidOperationException("请选择已编制区域和至少一家公司。");
            int minRow = areas.Min(a => a.Row), minColumn = areas.Min(a => a.Column);
            int width = areas.Max(a => checked(a.Column + a.Columns - 1)) - minColumn + 1;
            long count = 0;
            for (int i = 0; i < areas.Count; i++)
            {
                var a = areas[i];
                if (a.Rows < 1 || a.Columns < 1 || a.Row < 1 || a.Column < 1 || (long)a.Row + a.Rows - 1 > 1048576 || (long)a.Column + a.Columns - 1 > 16384)
                    throw new InvalidOperationException("已编制区域坐标无效。");
                if (a.Sheet != areas[0].Sheet) throw new InvalidOperationException("已编制区域不能跨工作表。");
                count += (long)a.Rows * a.Columns;
                for (int j = 0; j < i; j++) if (Intersects(a, areas[j])) throw new InvalidOperationException("已编制 Area 相互重叠。");
            }
            if (count * companies > CellLimit) throw new InvalidOperationException("MVP 每批目标最多十万单元格，请缩小区域或公司数量。");
            var result = new List<SummaryAreaPlan>();
            for (int company = 0; company < companies; company++) foreach (var area in areas)
            {
                var target = area.Target(minRow, minColumn, pasteRow, pasteColumn, width, company);
                if (areas.Any(source => Intersects(source, target))) throw new InvalidOperationException("已编制区域与目标重叠，原始已编制区域必须保持不变。");
                result.Add(target);
            }
            return result;
        }
        internal static void NativeFormula(string formula)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') throw new InvalidOperationException("公式类型无法确认。");
            string mask = Regex.Replace(formula, "\"(?:[^\"]|\"\")*\"", "\"\"");
            if (mask.IndexOfAny(new[] { '[', ']', '!', '#', '@', '{', '}', '\'', ';' }) >= 0)
                throw new InvalidOperationException("已有外链、跨表、数组、溢出或特殊公式不在 MVP 原生复制白名单内。");
            mask = Regex.Replace(mask, @"\$?[A-Za-z]{1,3}\$?[1-9][0-9]*", "0");
            var supported = new HashSet<string>(new[] { "ROUND", "SUM", "IF", "IFERROR", "ABS", "MIN", "MAX", "AVERAGE", "COUNT", "COUNTA", "AND", "OR", "NOT", "TRUE", "FALSE" });
            foreach (Match token in Regex.Matches(mask, @"[A-Za-z_][A-Za-z0-9_.]*"))
                if (!supported.Contains(token.Value.ToUpperInvariant())) throw new InvalidOperationException("已有公式含未确认名称或函数：" + token.Value);
        }
        internal static string Address(SummaryAreaPlan area)
        { return CellAddress.Name(area.Row, area.Column) + ":" + CellAddress.Name(area.Row + area.Rows - 1, area.Column + area.Columns - 1); }
        internal static object[,] FormulaMatrix(object copied, SummaryAreaPlan area, IList<SummaryAreaCell> cells, string folder, string company, int rows = 0, int columns = 0)
        {
            var result = new object[area.Rows, area.Columns];
            for (int r = 0; r < area.Rows; r++) for (int c = 0; c < area.Columns; c++) result[r, c] = SummaryArrays.Pick(copied, r, area.Rows, c, area.Columns);
            foreach (var cell in cells) if (cell.Existing != null) result[cell.Row-area.Row,cell.Column-area.Column] = cell.Existing.Render(folder,company,rows,columns); else if (!cell.Formula && !cell.Literal) result[cell.Row - area.Row, cell.Column - area.Column] = SummaryTextReference.External(folder, company, cell.Text);
            return result;
        }
        internal static void VerifyMatrix(object actual, object r1c1, SummaryAreaPlan area, IList<SummaryAreaCell> cells, string folder, string company, int rows = 0, int columns = 0)
        {
            foreach (var cell in cells.Where(c=>!c.Literal))
            {
                int r = cell.Row - area.Row, c = cell.Column - area.Column;
                string expected = cell.Existing != null ? cell.Existing.Render(folder,company,rows,columns) : cell.Formula ? cell.R1C1 : SummaryTextReference.External(folder, company, cell.Text);
                string found = Convert.ToString(SummaryArrays.Pick(cell.Formula && cell.Existing == null ? r1c1 : actual, r, area.Rows, c, area.Columns));
                if (!string.Equals(found, expected, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("公式批量读回核验失败，已编制区域位置：" + CellAddress.Name(cell.Row, cell.Column));
            }
        }
        internal static bool IsError(object value)
        { return value is System.Runtime.InteropServices.ErrorWrapper || (value is int && (((uint)(int)value & 0xFFFF0000u) == 0x800A0000u)) || (value is uint && (((uint)value & 0xFFFF0000u) == 0x800A0000u)); }
        internal static string XmlState(string xml)
        {
            var doc = new System.Xml.XmlDocument { XmlResolver = null };
            var options = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64L * 1024 * 1024 };
            using (var reader = System.Xml.XmlReader.Create(new System.IO.StringReader(xml), options)) doc.Load(reader);
            var ns = new System.Xml.XmlNamespaceManager(doc.NameTable); ns.AddNamespace("s", "urn:schemas-microsoft-com:office:spreadsheet");
            var styles = doc.SelectSingleNode("/s:Workbook/s:Styles", ns); var tables = doc.SelectNodes("/s:Workbook/s:Worksheet/s:Table", ns);
            if (styles == null || tables.Count != 1) throw new InvalidOperationException("宿主不支持完整区域 XML 状态。");
            var table = tables[0];
            // Cached results may change on recalculation; formula text and all
            // table/style/comment markup remain in the read-only fingerprint.
            foreach (System.Xml.XmlNode cache in table.SelectNodes(".//s:Cell[@s:Formula]/s:Data", ns)) cache.ParentNode.RemoveChild(cache);
            return styles.OuterXml + table.OuterXml;
        }
        internal static void ValidateShift(string formula, int rows, int columns)
        {
            NativeFormula(formula);
            foreach (var span in FormulaRefRewriter.Scan(formula))
            {
                CheckReference(span.Row1, span.Column1, span.Row1Absolute, span.Column1Absolute, rows, columns);
                if (span.IsRange) CheckReference(span.Row2, span.Column2, span.Row2Absolute, span.Column2Absolute, rows, columns);
            }
        }
        private static void CheckReference(int row, int column, bool rowAbsolute, bool columnAbsolute, int deltaRow, int deltaColumn)
        {
            long r = rowAbsolute ? row : (long)row + deltaRow, c = columnAbsolute ? column : (long)column + deltaColumn;
            if (r < 1 || r > 1048576 || c < 1 || c > 16384) throw new InvalidOperationException("已有公式复制后会产生越界引用。");
        }
    }
}
