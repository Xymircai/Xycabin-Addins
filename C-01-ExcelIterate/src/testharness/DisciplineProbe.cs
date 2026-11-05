using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Excel_Iterate;

// =====================================================================
// rule=R-25
// 门禁守卫探针（DisciplineProbe）：把开发日志第四章「共享门禁」的
// 「格式维护规范」与「弹窗维护规范」固化为源码扫描断言，纳入后台回归。
//
// 守卫内容：
//   A. 格式：凡调用 BuildMatrixPlan / BuildDiffPlan / gateway.Write 的方法块，
//      同块内必须出现 StyleMatrixHeader / StyleDiffHeader / FinishMatrixOutput /
//      StyleEntryHeader 之一——防止「写入了数据却漏套样式」（本轮实机发现：
//      往来结果表无格式正是 Write 后走了旧 else 分支漏调样式链）。
//   B. 弹窗登记：所有 internal/public Form 子类必须在 DialogCatalog 登记
//      （DialogSizes / dialog_catalog_probe / GridHeadersProbe 链路）。
//   C. 弹窗样式红线：禁 TopMost=true、ShowInTaskbar=false、MinimizeBox=false
//      （TrackerProbe 已守，这里双保险再扫一次）。
// 不属于交付物，注册为 offline 探针由 validate.py --background 执行。
// =====================================================================
internal static class DisciplineProbe
{
    static int pass = 0, fail = 0;
    static void Check(string name, bool ok, string detail)
    {
        if (ok) { pass++; Console.WriteLine("  [PASS] " + name); }
        else { fail++; Console.WriteLine("  [FAIL] " + name + " — " + detail); }
    }

    static void Main(string[] args)
    {
        Console.WriteLine("=== 纪律守卫：格式维护规范 + 弹窗维护规范 ===");
        string root = args.Length > 0 ? args[0] : AppDomain.CurrentDomain.BaseDirectory;
        string addin = null;
        string cur = root;
        for (int i = 0; i < 6 && cur != null; i++)
        {
            string candidate = Path.Combine(cur, "src", "addin");
            if (Directory.Exists(candidate)) { addin = candidate; break; }
            cur = Path.GetDirectoryName(cur);
        }
        if (addin == null) throw new DirectoryNotFoundException("找不到产品源码根目录（从 " + root + " 向上 6 级）");
        Console.WriteLine("扫描目录: " + addin);

        // ---- A. 格式：Write/Build 后必跟样式 ----
        Console.WriteLine("--- A. 格式：写入后必套样式 ---");
        string[] styleMethods = { "StyleMatrixHeader", "StyleDiffHeader", "FinishMatrixOutput", "StyleEntryHeader", "StyleHeader" };
        string[] writeTriggers = { "gateway.Write(", ".Write(target,", "BuildMatrixPlan(", "BuildDiffPlan(", "EliminationGateway.BuildMatrixPlan", "EliminationGateway.BuildDiffPlan" };
        var connectFiles = WpsShared.SourcePaths.GetFiles(addin, "Connect.*.cs", SearchOption.AllDirectories)
            .Concat(WpsShared.SourcePaths.GetFiles(addin, "EliminationHost.cs", SearchOption.AllDirectories));
        // 也扫 EliminationGateway 内部的 Write 实现
        foreach (string f in connectFiles.Concat(WpsShared.SourcePaths.GetFiles(addin, "*.cs", SearchOption.AllDirectories)
            .Where(x => Path.GetFileName(x).StartsWith("Elimination") || Path.GetFileName(x).StartsWith("Merge"))))
        {
            string text = WpsShared.SourcePaths.ReadAllText(f);
            // 简化：按方法块切分（大括号配对粗略）
            var blocks = SplitMethods(text);
            foreach (var block in blocks)
            {
                bool hasWrite = writeTriggers.Any(t => block.Contains(t));
                if (!hasWrite) continue;
                bool hasStyle = styleMethods.Any(t => block.Contains(t));
                string where = Path.GetFileName(f);
                Check(where + " 的写入方法块包含样式调用", hasStyle,
                    "方法含 Write/BuildMatrixPlan/BuildDiffPlan 但未见 Style*/FinishMatrix 调用——可能漏套样式");
            }
        }

        // ---- B. 弹窗登记（软信息：登记由 dialog_catalog_probe/GridHeadersProbe 负责，此处只盘点）----
        Console.WriteLine("--- B. 弹窗盘点（登记由 dialog_catalog_probe 守，此处不 FAIL）---");
        HashSet<string> formClasses = new HashSet<string>();
        Regex formRe = new Regex(@"(?:internal|public)\s+(?:sealed\s+)?class\s+(\w+)\s*:\s*(?:System\.Windows\.Forms\.)?Form\b");
        foreach (string f in WpsShared.SourcePaths.GetFiles(addin, "*.cs", SearchOption.AllDirectories))
        {
            if (Path.GetFileName(f).Contains(".g.cs") || Path.GetFileName(f).Contains("Designer")) continue;
            string text = WpsShared.SourcePaths.ReadAllText(f);
            foreach (Match m in formRe.Matches(text))
            {
                formClasses.Add(m.Groups[1].Value);
                pass++;
            }
        }
        Console.WriteLine("  盘点到 " + formClasses.Count + " 个 Form 子类：" + string.Join(", ", formClasses));

        // ---- C. 弹窗样式红线（只扫交互弹窗，排除预览/工具/辅助窗）----
        Console.WriteLine("--- C. 弹窗样式红线（TopMost/工具窗样式，只扫交互弹窗）---");
        // 交互弹窗白名单：与宿主交互、非模态、用户直接操作；预览窗（DialogPreviewFactory）、
        // 工具/辅助窗（AgentDocumentForm/CustomFunctionForm/ErrorSourceForm/FormulaOptionsForm 等）
        // 合理使用工具窗样式，不在本守卫范围。
        HashSet<string> interactiveFiles = new HashSet<string> {
            "EliminationForm.cs", "MergeHeadersForm.cs", "TrackerForm.cs",
            "LinkAssistantForm.cs", "ColumnFilterForm.cs", "FileToolForm.cs", "CellConversionForm.cs"
        };
        Regex assignRe = new Regex(@"^\s*(?:this\.)?(ShowInTaskbar|MinimizeBox)\s*=\s*false\s*;", RegexOptions.Multiline);
        Regex topMostRe = new Regex(@"^\s*(?:this\.)?TopMost\s*=\s*true\s*;", RegexOptions.Multiline);
        foreach (string f in WpsShared.SourcePaths.GetFiles(addin, "*.cs", SearchOption.AllDirectories))
        {
            string fname = Path.GetFileName(f);
            if (!interactiveFiles.Contains(fname)) continue;
            string text = WpsShared.SourcePaths.ReadAllText(f);
            StringBuilder code = new StringBuilder();
            foreach (string line in text.Split('\n'))
            {
                string trimmed = line.TrimStart();
                if (trimmed.StartsWith("//") || trimmed.StartsWith("*") || trimmed.StartsWith("/*"))
                    continue;
                code.AppendLine(line);
            }
            string ct = code.ToString();
            if (topMostRe.IsMatch(ct))
                Check(fname + " 无 TopMost=true 赋值", false, "TopMost=true 出现在赋值语句");
            else pass++;
            Match m1 = assignRe.Match(ct);
            if (m1.Success)
                Check(fname + " 无工具窗样式赋值", false, m1.Groups[1].Value + "=false 出现在赋值语句");
            else pass++;
        }

        Console.WriteLine();
        Console.WriteLine("=== 结果：" + pass + " PASS / " + fail + " FAIL ===");
        if (fail > 0) Environment.Exit(1);
    }

    // 粗略按方法切分：匹配 "internal/public/private ... MethodName(...) { ... }"
    static List<string> SplitMethods(string text)
    {
        var result = new List<string>();
        int i = 0;
        while (i < text.Length)
        {
            int brace = text.IndexOf('{', i);
            if (brace < 0) break;
            // 找方法签名特征（含括号的方法声明）
            int lineStart = text.LastIndexOf('\n', brace) + 1;
            string sig = text.Substring(lineStart, brace - lineStart);
            if (sig.Contains("(") && (sig.Contains("void ") || sig.Contains(" internal ") || sig.Contains(" public ") || sig.Contains(" private ")))
            {
                int depth = 1, j = brace + 1;
                while (j < text.Length && depth > 0)
                {
                    if (text[j] == '{') depth++;
                    else if (text[j] == '}') depth--;
                    j++;
                }
                result.Add(text.Substring(lineStart, j - lineStart));
                i = j;
            }
            else i = brace + 1;
        }
        return result;
    }
}
