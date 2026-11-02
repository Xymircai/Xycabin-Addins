using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class CompanyWorkbookProgress
    {
        internal int Total { get; private set; }
        internal int Completed { get; private set; }
        internal int Succeeded { get; private set; }
        internal int Failed { get; private set; }
        internal int Remaining { get { return Total - Completed; } }
        internal double Percent { get { return Total == 0 ? 0 : Math.Floor(Completed * 1000.0 / Total) / 10; } }
        internal void Reset(int total)
        {
            if (total < 0) throw new ArgumentOutOfRangeException("total");
            Total = total; Completed = Succeeded = Failed = 0;
        }
        internal void Complete(bool succeeded)
        {
            if (Completed >= Total) throw new InvalidOperationException("没有待完成的任务。");
            Completed++; if (succeeded) Succeeded++; else Failed++;
        }
        internal string Counts(string state)
        {
            return string.Format(System.Globalization.CultureInfo.InvariantCulture,
                "{0} · 已处理 {1} / {2} 项（{3:0.0}%） · 成功 {4} · 失败 {5} · 剩余 {6}",
                state, Completed, Total, Percent, Succeeded, Failed, Remaining);
        }
        internal string Timing(TimeSpan elapsed, bool running)
        {
            string text = "已用 " + Duration(elapsed.TotalSeconds);
            if (Completed == 0 || elapsed.TotalSeconds <= 0)
                return text + (running ? " · 平均耗时与预计剩余：首项完成后计算" : "");
            double average = elapsed.TotalSeconds / Completed;
            text += " · 平均 " + (average < 0.1 ? "<0.1" : average.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)) + " 秒/项";
            if (running && Remaining > 0) text += " · 预计剩余约 " + Duration(average * Remaining);
            return text;
        }
        private static string Duration(double seconds)
        {
            long whole = (long)Math.Ceiling(Math.Max(0, seconds));
            return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:00}:{1:00}:{2:00}", whole / 3600, whole / 60 % 60, whole % 60);
        }
    }
    internal sealed class CompanyWorkbookRequest
    {
        internal string SourceSheet, TargetSheet, NameField, CodeField, FullNameField, TargetCodeField, TargetNameField, Output;
        internal int HeaderRow, FirstRow, LastRow, TargetHeaderRow, TargetRow;
        internal bool UpdateLead;
        internal void Validate()
        {
            if (HeaderRow < 1 || FirstRow <= HeaderRow || LastRow < FirstRow)
                throw new InvalidOperationException("请明确填写来源表头行、起始行和结束行；起始行须在表头之后。");
            if (string.IsNullOrWhiteSpace(SourceSheet) || string.IsNullOrWhiteSpace(NameField))
                throw new InvalidOperationException("请选择来源工作表和文件命名字段。");
            if (!Directory.Exists(Output)) throw new InvalidOperationException("请选择已存在的输出文件夹。");
            if (UpdateLead && (TargetHeaderRow < 1 || TargetRow <= TargetHeaderRow || string.IsNullOrWhiteSpace(TargetSheet)
                || string.IsNullOrWhiteSpace(CodeField) || string.IsNullOrWhiteSpace(FullNameField)
                || string.IsNullOrWhiteSpace(TargetCodeField) || string.IsNullOrWhiteSpace(TargetNameField)))
                throw new InvalidOperationException("请明确填写 Lead 的表头行、唯一写入行及两个来源/目标字段；只修改指定行。");
        }
    }
    internal sealed class CompanyWorkbookRecord
    {
        internal int Row;
        internal string Name, Code, FullName, OutputPath;
    }
    internal static class CompanyWorkbookRules
    {
        internal static string Text(object value)
        {
            if (value is System.Runtime.InteropServices.ErrorWrapper) throw new InvalidOperationException("单元格包含 Excel 错误值。");
            // Excel Value2 returns error variants as signed HRESULTs on some hosts.
            if (value is int && (int)value <= -2146820000 && (int)value >= -2146830000)
                throw new InvalidOperationException("单元格包含 Excel 错误值。");
            return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "";
        }
        internal static string FileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException("文件命名字段为空。");
            if (value != value.Trim() || value.EndsWith(".", StringComparison.Ordinal))
                throw new InvalidOperationException("文件名含前后空格或末尾句点，请修正来源数据。");
            if (value.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || value == "." || value == "..")
                throw new InvalidOperationException("文件名含 Windows 禁止字符。");
            string stem = value.Split('.')[0].TrimEnd(' ');
            if (Regex.IsMatch(stem, @"^(CON|PRN|AUX|NUL|CLOCK\$|CONIN\$|CONOUT\$|COM[1-9¹²³]|LPT[1-9¹²³])$", RegexOptions.IgnoreCase))
                throw new InvalidOperationException("文件名为 Windows 保留名称。");
            return value;
        }
        internal static int Column(IDictionary<string, int> headers, string field)
        {
            int column;
            if (!headers.TryGetValue(field.Trim(), out column)) throw new InvalidOperationException("找不到字段「" + field + "」。");
            return column;
        }
        internal static void ValidateRecords(IList<CompanyWorkbookRecord> rows, string output, string extension)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in rows)
            {
                try
                {
                    string name = FileName(row.Name);
                    if (!names.Add(name)) throw new InvalidOperationException("公司文件名重复（忽略大小写）。");
                    row.OutputPath = Path.Combine(Path.GetFullPath(output), name + extension);
                    if (row.OutputPath.Length >= 240 || (name + extension).Length > 255) throw new InvalidOperationException("输出路径过长。");
                    if (File.Exists(row.OutputPath) || Directory.Exists(row.OutputPath)) throw new InvalidOperationException("同名输出已存在，不会覆盖：" + row.OutputPath);
                }
                catch (Exception error) { throw new InvalidOperationException("来源第 " + row.Row + " 行：" + error.Message, error); }
            }
            if (rows.Count == 0) throw new InvalidOperationException("指定范围内没有有效公司记录。");
        }
    }
}
