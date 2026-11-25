using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Excel_Iterate
{
    // User results and diagnostics are independent. Neither observer nor disk
    // failures may change the outcome of a workbook operation.
    internal sealed class LinkOperationLog
    {
        private readonly StringBuilder history = new StringBuilder(), diagnostic = new StringBuilder();
        private readonly Action<string, string> writer;
        internal event Action<string> LineAdded, DiagnosticLineAdded;
        internal event Action DisplayReset;
        internal string FilePath { get; private set; }
        internal string PersistenceError { get; private set; }
        internal string Text { get { return history.ToString(); } }
        internal string DiagnosticText { get { return diagnostic.ToString(); } }
        internal LinkOperationLog() : this((path, value) => File.AppendAllText(path, value, new UTF8Encoding(false))) { }
        internal LinkOperationLog(Action<string, string> append) { writer = append; }
        internal void ClearDisplay()
        {
            history.Clear(); PersistenceError = null;
            var handlers = DisplayReset;
            if (handlers != null) foreach (Action handler in handlers.GetInvocationList()) try { handler(); } catch { }
        }
        internal void RestoreDisplay(string text)
        {ClearDisplay();text=text.Replace("\r\n","\n").Replace("\r","\n").Replace("\n","\r\n");history.Append(text);Notify(LineAdded,text);}
        internal void EnablePersistence(string directory)
        {
            if (FilePath != null || PersistenceError != null) return;
            try
            {
                Directory.CreateDirectory(directory);
                FilePath = Path.Combine(directory, "session-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N") + ".log");
                writer(FilePath, DiagnosticText);
                Trace("日志文件", FilePath);
            }
            catch (Exception ex) { ReportPersistenceFailure(ex); }
        }
        internal void Record(string kind, string message)
        {
            string line = (string.IsNullOrEmpty(kind) ? "" : kind + "：") + message + "\r\n";
            history.Append(line); Notify(LineAdded, line);
            Trace(string.IsNullOrEmpty(kind) ? "用户结果" : kind, message);
        }
        internal void Trace(string kind, string message)
        {
            string line = DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz") + "  " + kind + "：" + message + "\r\n";
            diagnostic.Append(line); Notify(DiagnosticLineAdded, line);
            if (FilePath != null && PersistenceError == null)
                try { writer(FilePath, line); } catch (Exception ex) { ReportPersistenceFailure(ex); }
        }
        internal void ReportPersistenceFailure(Exception ex)
        {
            if (PersistenceError != null) return;
            PersistenceError = ex.Message;
            diagnostic.Append(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss.fff zzz") + "  日志保存失败：" + ex + "\r\n");
            string line = "日志保存失败：诊断记录未完整写入；操作结果仍有效，请复制窗口日志。\r\n";
            history.Append(line); Notify(LineAdded, line);
        }
        private static void Notify(Action<string> handlers, string line)
        {
            if (handlers == null) return;
            foreach (Action<string> handler in handlers.GetInvocationList())
                try { handler(line); } catch { }
        }
    }

    internal enum LinkScanScope { Current, Workbooks, Folder }
    internal sealed class LinkScanSelection
    {
        private readonly string current;
        private readonly LinkOperationLog log;
        internal LinkScanScope Scope { get; private set; }
        internal string[] Picked { get; private set; }
        internal LinkScanSelection(string path, LinkOperationLog journal)
        {
            current = path; log = journal; Picked = new string[0];
            log.Trace("扫描范围", "当前工作簿；" + path);
        }
        internal bool SetScope(LinkScanScope value)
        {
            if (Scope == value) return false;
            Scope = value; Picked = new string[0];
            log.Trace("切换范围", value == LinkScanScope.Current ? "当前工作簿；" + current : value == LinkScanScope.Workbooks ? "自选其他工作簿" : "自选其他文件夹（仅本层）");
            return true;
        }
        internal void Choose(string[] values, bool accepted)
        {
            string type = Scope == LinkScanScope.Folder ? "文件夹" : "工作簿";
            if (!accepted) { log.Trace("取消选择", type + "；保留原选择"); return; }
            Picked = values.Select(Path.GetFullPath).Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            foreach (string path in Picked) log.Trace("选择" + type, path);
        }
        internal string[] Candidates()
        {
            log.Trace("开始扫描", Scope == LinkScanScope.Current ? "当前工作簿" : Scope == LinkScanScope.Workbooks ? "自选其他工作簿" : "文件夹本层");
            if (Scope == LinkScanScope.Current) return new[] { current };
            if (Picked.Length == 0) throw new InvalidOperationException("请先选择工作簿或文件夹。");
            if (Scope == LinkScanScope.Workbooks) return (string[])Picked.Clone();
            int skipped;
            string[] files = FolderCandidates(Picked[0], out skipped);
            log.Trace("文件夹扫描", Picked[0] + "；纳入 " + files.Length + " 个工作簿；跳过临时文件 " + skipped + " 个；不扫描子文件夹");
            foreach (string path in files) log.Trace("纳入扫描", path);
            if (files.Length == 0) log.Trace("扫描未执行", "文件夹本层没有可扫描的 xlsx 工作簿，未导出清单");
            return files;
        }

        internal static string[] FolderCandidates(string folder, out int skipped)
        {
            string[] all = Directory.GetFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly);
            string[] files = all.Where(p => !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal) &&
                    !Path.GetFileName(p).StartsWith(".EI_LinkStage_", StringComparison.OrdinalIgnoreCase) &&
                    !OperationBatch.ManagedFile(p))
                .OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
            skipped = all.Length - files.Length;
            return files;
        }
    }

    internal sealed class LinkFileOutcome
    {
        internal readonly string Path;
        internal bool Committed, Verified;
        internal string Backup, Error, ReopenError, CleanupError;
        internal LinkFileOutcome(string path) { Path = path; }
        internal bool HasWarning { get { return (Committed && !Verified) || ReopenError != null || CleanupError != null; } }
        internal void Record(LinkOperationLog log)
        {
            string result = !Committed ? "替换失败，原文件未提交" : !Verified ? "已提交，但后续核对失败" : ReopenError != null ? "已替换，但重新打开失败" : "替换成功，提交后核对通过";
            log.Record("文件结果", Path + "；" + result + (Backup == null ? "" : "；备份：" + Backup) + (Error == null ? "" : "；" + Error) + (ReopenError == null ? "" : "；重开错误：" + ReopenError) + (CleanupError == null ? "" : "；临时文件清理失败：" + CleanupError));
        }
    }
}
