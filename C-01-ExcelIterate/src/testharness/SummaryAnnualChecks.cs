using System;
using System.Collections.Generic;
using System.IO;
using Excel_Iterate;

public sealed class SummaryAnnualFakeCell { public bool MergeCells { get; set; } public bool HasArray { get; set; } public object Value2 { get; set; } }
public sealed class SummaryAnnualFakeCells
{
    private readonly Dictionary<string, SummaryAnnualFakeCell> cells = new Dictionary<string, SummaryAnnualFakeCell>();
    public SummaryAnnualFakeCell this[int row, int column]
    {
        get { string key = row + ":" + column; if (!cells.ContainsKey(key)) cells[key] = new SummaryAnnualFakeCell(); return cells[key]; }
    }
}
public sealed class SummaryAnnualFakeRows { public int Count { get { return 65536; } } }
public sealed class SummaryAnnualFakeSheet
{
    public SummaryAnnualFakeRows Rows { get { return new SummaryAnnualFakeRows(); } }
    public SummaryAnnualFakeCells Cells { get; private set; }
    public SummaryAnnualFakeSheet() { Cells = new SummaryAnnualFakeCells(); }
}
public sealed class SummaryAnnualFakeBook { public string FullName { get; set; } }
public sealed class SelectionCount { public int Count { get; set; } }
public sealed class SelectionSheet
{
    public object Parent { get; set; }
    public SelectionCount Rows { get { return new SelectionCount { Count = 1048576 }; } }
    public SelectionCount Columns { get { return new SelectionCount { Count = 16384 }; } }
}
public sealed class SelectionRange
{
    public object Parent { get; set; } public int Row { get; set; } public int Column { get; set; }
    public SelectionCount Areas { get; private set; } public SelectionCount Rows { get; private set; } public SelectionCount Columns { get; private set; }
    public SelectionRange() { Areas = new SelectionCount { Count = 1 }; Rows = new SelectionCount { Count = 1 }; Columns = new SelectionCount { Count = 1 }; }
}
internal static class SummaryAnnualChecks
{
    private static void SelectionChecks(Action<string, bool, string> check)
    {
        Action<Action, string> reject = delegate(Action action, string name) {
            bool rejected = false; try { action(); } catch (InvalidOperationException) { rejected = true; }
            check(name, rejected, "invalid selection accepted");
        };
        SummarySelectionRules.Validate(3, 4, 1, 16384, false, false);
        SummarySelectionRules.Validate(1, 4, 1048576, 1, true, false);
        check("整行及整列仅取坐标，不受迁移单元格数限制", true, "");
        reject(delegate { SummarySelectionRules.Validate(1, 1, 2, 1, false, false); }, "表头跨行拒绝");
        reject(delegate { SummarySelectionRules.Validate(1, 1, 1, 2, true, false); }, "模板跨列拒绝");
        reject(delegate { SummarySelectionRules.Validate(1, 1, 1, 1, false, true); }, "合并区域拒绝");
        reject(delegate { SummarySelectionRules.Validate(1, 1, 1, 1, false, DBNull.Value); }, "混合合并区域拒绝");
        var book = new object(); var sheet = new SelectionSheet { Parent = book };
        var selection = new SelectionRange { Parent = sheet, Row = 3, Column = 4 };
        var coordinates = DataMigrationHost.ReadCoordinates(selection, book, sheet);
        check("读取选区使用绑定工作簿和工作表", coordinates.Row == 3 && coordinates.Column == 4, "wrong coordinates");
        reject(delegate { DataMigrationHost.ReadCoordinates(selection, book, new SelectionSheet { Parent = book }); }, "同名其他工作表身份拒绝");
        reject(delegate { DataMigrationHost.ReadCoordinates(selection, new object(), sheet); }, "其他工作簿拒绝");
        selection.Areas.Count = 2;
        reject(delegate { DataMigrationHost.ReadCoordinates(selection, book, sheet); }, "不连续选区拒绝");
        reject(delegate { DataMigrationHost.ReadCoordinates(new object(), book, sheet); }, "非单元格选区拒绝");
        reject(delegate { DataMigrationHost.ReadCoordinates(null, book, sheet); }, "空选区拒绝");
        using (var form = new SingleSummaryForm())
        {
            check("年度来源检查首次打开不勾选", !form.CheckYearFiles.Checked && !new SummaryRequest().CheckYearFiles, "default enabled");
            form.SetContext("C:\\test.xlsx", new List<string> { "lead", "summary" }, "lead", "summary", "D");
            form.SetPreview("source.xlsx", true, "ready");
            form.ApplySelection("header", 5, 4, "selected", "title");
            check("读取汇总表头使预检失效并保留手动参数", form.BuildRequest().SheetHeaderRow == 5 && !form.Execute.Enabled, "stale preview");
            form.ApplySelection("column", 7, 6, "selected", "title");
            check("模板选区回填实际列号", form.BuildRequest().TemplateColumn == 6, "wrong column");
            form.SetUndoAvailable(true);
            check("待还原批次禁止读取选区", !form.Controls.Find("summaryPick_column", true)[0].Enabled, "button unlocked");
        }
    }
    private static SummaryRequest Request()
    {
        var request = new SummaryRequest { Annual = true, FirstYear = 2023, LastYear = 2026, YearHeaderRow = 3, SheetHeaderRow = 2,
            LeadSheet = "lead", TargetSheet = "汇总", TemplateColumn = 4, SourceFile = "！模板公司.xlsx", CheckYearFiles = false };
        request.Companies.Add(new SummaryCompany { Row = 2, Code = "甲公司", Name = "！甲公司" });
        request.Companies.Add(new SummaryCompany { Row = 3, Code = "乙公司", Name = "！乙公司" });
        return request;
    }
    internal static void Run(Action<string, bool, string> check)
    {
        SelectionChecks(check);
        check("F列校验公式复制到G列随列平移", SummaryFormulaShift.Shift("=ROUND(F76-F121,2)", 1) == "=ROUND(G76-G121,2)", "formula stayed on F");
        check("复制多列使用真实列偏移", SummaryFormulaShift.Shift("=ROUND(F76-F121,2)", 4) == "=ROUND(J76-J121,2)", "wrong offset");
        check("绝对列不移动而相对列移动", SummaryFormulaShift.Shift("=$F76+F$121+$F$9", 1) == "=$F76+G$121+$F$9", "mixed anchors changed");
        check("同表范围和整列按相对规则移动", SummaryFormulaShift.Shift("=SUM(F76:F121)+SUM(F:F)+SUM($F:$F)", 1) == "=SUM(G76:G121)+SUM(G:G)+SUM($F:$F)", "range shift failed");
        check("字符串及跨表跨簿地址沿用原规则", SummaryFormulaShift.Shift("=F76+BS!L14+'N:\\[甲.xlsx]BS'!F6+\"F76\"", 1) == "=G76+BS!L14+'N:\\[甲.xlsx]BS'!F6+\"F76\"", "qualified references changed");
        check("表格结构引用及函数名不误识别", SummaryFormulaShift.Shift("=LOG10(F76)+Table1[Column1]", 1) == "=LOG10(G76)+Table1[Column1]", "function/table rewritten");
        bool overflow = false;
        try { SummaryFormulaShift.Shift("=XFD1", 1); } catch (InvalidOperationException) { overflow = true; }
        check("公式引用列越界在写入前阻断", overflow, "overflow accepted");
        var ordinary = Request(); ordinary.Annual = false; ordinary.TemplateColumn = 6;
        var destination = new SummaryYearColumn { Column = 7, Company = ordinary.Companies[0] };
        check("普通模式换链使用同表平移后的公式", ordinary.Rewrite("=ROUND(F76-F121,2)", destination).Text == "=ROUND(G76-G121,2)", "ordinary mode wrong");
        ordinary.Annual = true; ordinary.YearPathPattern = "GROUP_DEMO_{yyyy}1231"; destination.Year = 2024;
        check("年度模式混合公式仅移动同表引用", ordinary.Rewrite("=F76+'N:\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!F6", destination).Text == "=G76+'N:\\GROUP_DEMO_20241231\\[！甲公司.xlsx]BS'!F6", "annual mixed formula wrong");
        using (var form = new SingleSummaryForm())
        {
            form.SetContext("C:\\fixture.xlsx", new List<string> { "lead", "summary" }, "lead", "summary", "F");
            check("四个参数均提供在表中选择入口", form.Controls.Find("summaryRange_lead", true).Length == 1 && form.Controls.Find("summaryRange_header", true).Length == 1 && form.Controls.Find("summaryRange_column", true).Length == 1 && form.Controls.Find("summaryRange_year", true).Length == 1, "picker missing");
            form.SetUndoAvailable(true);
            check("待还原时禁止原生选择", !form.Controls.Find("summaryRange_column", true)[0].Enabled, "picker unlocked");
        }
        Action<Action, string> reject = delegate(Action action, string title)
        { bool failed = false; try { action(); } catch (InvalidOperationException) { failed = true; } check(title, failed, "应阻断"); };
        try
        {
            var request = Request(); var columns = request.Columns();
            check("两家公司四个年度生成八列", request.Validate() == null && columns.Count == 8 && columns[0].Column == 4 && columns[7].Column == 11, "列范围错误");
            check("公司外层分组，年度默认倒序", columns[0].Year == 2026 && columns[3].Year == 2023 && columns[4].Company.Code == "乙公司" && columns[4].Year == 2026, "顺序错误");
            check("每个年度列重复公司名称", columns[0].Company.Code == columns[3].Company.Code && columns[4].Company.Code == columns[7].Company.Code && columns[0].YearLabel == "2026年度", "公司名或年度标题错误");
            request.DescendingYears = false; check("可选择升序年度", request.Columns()[0].Year == 2023 && request.Columns()[3].Year == 2026, "升序未生效"); request.DescendingYears = true;
            request.YearHeaderRow = 2; check("年度行必须在公司行下面", request.ValidateYears() != null, "未拦截"); request.YearHeaderRow = 3;
            request.LastYear = 2022; check("倒置年度范围被拒绝", request.ValidateYears() != null, "未拦截"); request.LastYear = 2026;
            request.TemplateColumn = 16380; check("公司乘年度计入列越界检查", request.Validate() != null, "未拦截"); request.TemplateColumn = 4;
            request.Companies[1].Selected = false; check("取消勾选公司不生成该公司任何年度", request.Columns().Count == 4, "公司仍出现"); request.Companies[1].Selected = true;
            request.Annual = false; request.YearHeaderRow = 0; check("旧单列模式不要求年度设置", request.Columns().Count == 2 && request.Columns()[0].Year == 0 && request.Validate() == null, "旧模式受影响"); request = Request();

            string formula = "='N:\\Data\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!$AA$6+2025+\"2025 [！模板公司.xlsx]\"+'N:\\Data\\GROUP_DEMO_20241231\\[其他公司.xlsx]BS'!$AA$6";
            var formulas = new List<string> { formula }; SummaryAnnualLinks.Prepare(request, formulas);
            check("只从模板公司链接自动识别年度目录", request.YearPathPattern == "GROUP_DEMO_{yyyy}1231" && request.TemplateYear == 2025, request.YearPathPattern);
            var rewrite = request.Rewrite(formula, request.Columns()[0]);
            check("改公司文件名和年度，保留感叹号与绝对地址", rewrite.Text.Contains(@"GROUP_DEMO_20261231\[！甲公司.xlsx]BS'!$AA$6") && rewrite.Count == 1, rewrite.Text);
            check("不修改常量和其他公司链接", rewrite.Text.Contains("+2025+\"2025 [！模板公司.xlsx]\"") && rewrite.Text.Contains(@"GROUP_DEMO_20241231\[其他公司.xlsx]BS'!$AA$6"), rewrite.Text);
            var same = new SummaryYearColumn { Company = new SummaryCompany { Name = "！模板公司" }, Year = 2024 };
            check("同一公司也必须按年度换目录", request.Rewrite(formula, same).Text.Contains(@"GROUP_DEMO_20241231\[！模板公司.xlsx]"), "未换年度");
            check("内部引用不发生相对平移", request.Rewrite("=BS!L14-BS!G14", columns[0]).Text == "=BS!L14-BS!G14", "内部公式变化");
            string quoted = SummaryAnnualLinks.Rewrite("='N:\\O''Brien\\GROUP_DEMO_20251231\\[O''Brien.xlsx]BS'!$AA$6", "O'Brien.xlsx", "A'B.xlsx", request.YearPathPattern, 2024).Text;
            check("单引号转义和目录字符保留", quoted == "='N:\\O''Brien\\GROUP_DEMO_20241231\\[A''B.xlsx]BS'!$AA$6", quoted);
            string bare = SummaryAnnualLinks.Rewrite(@"=N:\Data\GROUP_DEMO_20251231\[！模板公司.xlsx]BS!$AA$6", request.SourceFile, "甲.xlsx", request.YearPathPattern, 2023).Text;
            check("未加引号的外链同样按目录段换年", bare == @"=N:\Data\GROUP_DEMO_20231231\[甲.xlsx]BS!$AA$6", bare);
            var ambiguous = Request();
            reject(delegate { SummaryAnnualLinks.Prepare(ambiguous, new List<string> { "='N:\\ARCHIVE_20200101\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!A1" }); }, "多个日期目录要求用户指定");
            ambiguous.YearPathPattern = "GROUP_DEMO_{yyyy}1231";
            SummaryAnnualLinks.Prepare(ambiguous, new List<string> { "='N:\\ARCHIVE_20200101\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!A1" });
            check("明确规则只改指定段", ambiguous.Rewrite("='N:\\ARCHIVE_20200101\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!A1", ambiguous.Columns()[0]).Text.Contains(@"ARCHIVE_20200101\GROUP_DEMO_20261231"), "其他日期目录被改");
            reject(delegate { SummaryAnnualLinks.Prepare(Request(), new List<string> { formula, formula.Replace("20251231", "20241231") }); }, "模板混合不同年度不猜测同比关系");
            reject(delegate { SummaryAnnualLinks.Prepare(Request(), new List<string> { "=[！模板公司.xlsx]BS!A1" }); }, "没有年度目录时明确阻止");
            request.YearPathPattern = "A_{yyyy}_{yyyy}"; check("目录规则不能有两个年度占位", request.ValidateYears() != null, "规则未校验"); request = Request();
            var selection = new SummaryLeadSnapshot { HeaderRow = 1, PickWriteColumn = 3 };
            selection.Companies.Add(new SummaryCompany { Row = 2 }); selection.Companies.Add(new SummaryCompany { Row = 5 });
            SummaryLeadReader.ApplySelection(selection, new List<SummaryCompany> { selection.Companies[0] });
            var lead = new SummaryAnnualFakeSheet(); lead.Cells[3, 3].Value2 = "保留间隔行";
            SummaryLeadReader.WriteSelection(lead, selection);
            check("公司勾选按实际来源行回写且保留间隔行", object.Equals(lead.Cells[2, 3].Value2, true) && object.Equals(lead.Cells[5, 3].Value2, false) && object.Equals(lead.Cells[3, 3].Value2, "保留间隔行"), "勾选行错位或未选公司被选用");

            string folder = Path.Combine(Path.GetTempPath(), "summary-years-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
            try
            {
                foreach (int year in new[] { 2023, 2024, 2025, 2026 })
                {
                    string dir = Path.Combine(folder, "GROUP_DEMO_" + year + "1231"); Directory.CreateDirectory(dir);
                    foreach (var company in request.Companies) File.WriteAllText(Path.Combine(dir, company.FileName), "fixture");
                }
                var probe = new SummaryTemplateProbe { LastRow = 8 };
                probe.Formulas.Add("='" + folder.Replace("'", "''") + "\\GROUP_DEMO_20251231\\[！模板公司.xlsx]BS'!$AA$6");
                request.CheckYearFiles = true; request.YearHeaderRow = 10;
                var sheet = new SummaryAnnualFakeSheet(); var book = new SummaryAnnualFakeBook { FullName = Path.Combine(folder, "summary.xlsx") };
                var plan = SummaryTemplateReader.AnnualPreview(book, sheet, request, probe);
                check("预检列出八列实际年度文件并扩大还原行范围", plan.Count == 9 && probe.LastRow == 10 && plan[1].StartsWith("D｜甲公司｜2026年度"), "计划缺失");
                string missing = Path.Combine(folder, "GROUP_DEMO_20231231", request.Companies[1].FileName); File.Delete(missing);
                reject(delegate { SummaryTemplateReader.AnnualPreview(book, sheet, request, probe); }, "缺少任何公司年度文件阻断生成");
                request.CheckYearFiles = false; plan = SummaryTemplateReader.AnnualPreview(book, sheet, request, probe);
                check("显式关闭文件检查时注明未验证", plan[plan.Count - 1].Contains("不代表来源文件存在"), "没有提示");
                sheet.Cells[10, 11].MergeCells = true;
                reject(delegate { SummaryTemplateReader.AnnualPreview(book, sheet, request, probe); }, "最后一列年度行也禁止合并"); sheet.Cells[10, 11].MergeCells = false;
                request.YearHeaderRow = 65537;
                reject(delegate { SummaryTemplateReader.AnnualPreview(book, sheet, request, probe); }, "检查实际工作表行上限");
                Directory.Delete(folder, true);
            }
            catch { Console.WriteLine("Retained annual fixture: " + folder); throw; }
            using (var form = new SingleSummaryForm())
            {
                check("年度模式可选且默认倒序", !form.AnnualMode.Checked && form.DescendingYears.Checked, "默认设置不符");
                form.AnnualMode.Checked = true; form.FirstYear.Value = 2023; form.LastYear.Value = 2026; form.YearRow.Value = 4;
                check("年度控件写入请求", form.BuildRequest().Annual && form.BuildRequest().YearCount == 4 && form.BuildRequest().YearHeaderRow == 4, "配置未传递");
            }
        }
        catch (Exception error) { check("多年度托管夹具", false, error.ToString()); }
    }
}
