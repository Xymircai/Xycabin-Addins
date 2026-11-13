using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Excel_Iterate
{
    // Commit only an already closed and verified staging file. File.Replace
    // captures the displaced target as the recovery file in the same operation.
    internal static class MigrationOutputTransaction
    {
        internal static string Commit(string stage, string target, string expectedHash, string batch)
        {
            stage = Path.GetFullPath(stage); target = Path.GetFullPath(target);
            if (string.Equals(stage, target, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(Path.GetDirectoryName(stage), Path.GetDirectoryName(target), StringComparison.OrdinalIgnoreCase))
                throw new IOException("临时文件必须与目标位于同一目录，且不能是目标本身。");
            if (string.IsNullOrEmpty(batch) || batch != Path.GetFileName(batch) || batch == "." || batch == "..")
                throw new IOException("备份批次标识无效。");
            PlainPath(target); PlainPath(stage);
            if (DataMigrationHost.Digest(target) != expectedHash) throw new IOException("目标文件在预检后已变化，请重新预检。");
            // Fail before commit on ordinary sharing/read-only errors. There is
            // no delete-then-move fallback when atomic replacement is unavailable.
            using (File.Open(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            string verified = DataMigrationHost.Digest(stage);
            string directory = Path.Combine(Path.GetDirectoryName(target), "复制备份", batch);
            PlainPath(directory); Directory.CreateDirectory(directory); PlainPath(directory);
            string backup = Path.Combine(directory, Path.GetFileName(target));
            if (File.Exists(backup) || Directory.Exists(backup)) throw new IOException("备份文件已存在，未覆盖目标：" + backup);
            // Recheck after directory creation and immediately before the switch.
            if (DataMigrationHost.Digest(target) != expectedHash) throw new IOException("目标文件在提交前已变化，请重新预检。");
            File.Replace(stage, target, backup, false);
            try
            {
                if (DataMigrationHost.Digest(backup) != expectedHash) throw new IOException("提交期间目标发生并发变化。");
                if (DataMigrationHost.Digest(target) != verified) throw new IOException("替换后目标摘要与已校验内容不一致。");
            }
            catch (Exception error)
            {
                // Restore only our own installed bytes. A subsequent external
                // writer must never be overwritten by a speculative rollback.
                try
                {
                    if (DataMigrationHost.Digest(target) != verified)
                        throw new IOException("目标已被其他程序改变，未自动回滚；恢复文件：" + backup);
                    string failed = Path.Combine(directory, "未交付_" + Guid.NewGuid().ToString("N") + Path.GetExtension(target));
                    File.Replace(backup, target, failed, false);
                }
                catch (Exception rollback) { throw new IOException("提交核验失败；恢复文件位于 " + directory + "。" + rollback.Message, error); }
                throw new IOException("提交核验失败，已恢复替换前目标。", error);
            }
            return backup;
        }
        // Put a migration backup back in place. The migrated version is moved
        // next to the backup instead of being discarded, and nothing happens
        // unless the target still holds exactly the bytes the migration wrote.
        internal static string Restore(string target, string backup, string expectedCurrentHash)
        {
            target = Path.GetFullPath(target); backup = Path.GetFullPath(backup);
            PlainPath(target); PlainPath(backup);
            if (!File.Exists(backup)) throw new IOException("找不到备份文件：" + backup);
            if (!File.Exists(target)) throw new IOException("目标文件不存在：" + target);
            if (DataMigrationHost.Digest(target) != expectedCurrentHash) throw new IOException("复制后该文件又被修改过，未还原；备份位于：" + backup);
            try { using (File.Open(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { } }
            catch (IOException) { throw new IOException("文件正在被使用，请先在 WPS 中关闭后再还原。"); }
            catch (UnauthorizedAccessException) { throw new IOException("文件为只读或没有写入权限，未还原。"); }
            string original = DataMigrationHost.Digest(backup);
            string stage = Path.Combine(Path.GetDirectoryName(target), ".xy-migration-restore-" + Guid.NewGuid().ToString("N") + Path.GetExtension(target));
            string displaced = Path.Combine(Path.GetDirectoryName(backup), "还原前_" + Guid.NewGuid().ToString("N").Substring(0, 8) + "_" + Path.GetFileName(target));
            File.Copy(backup, stage, false);
            try
            {
                if (DataMigrationHost.Digest(stage) != original) throw new IOException("复制备份时内容不一致，未还原。");
                if (DataMigrationHost.Digest(target) != expectedCurrentHash) throw new IOException("还原前文件发生变化，未还原。");
                File.Replace(stage, target, displaced, false);
            }
            catch
            {
                try { if (File.Exists(stage)) File.Delete(stage); } catch { }
                throw;
            }
            if (DataMigrationHost.Digest(target) != original) throw new IOException("还原后内容与备份不一致；复制后的版本保存在：" + displaced);
            return displaced;
        }
        // Permanently delete backup files this window created, then remove
        // the batch folder and the "复制备份" folder when they end up empty.
        // Only the listed files are touched; anything else is left in place.
        internal static void DeleteBackups(IEnumerable<string> files)
        {
            var folders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string file in files)
            {
                string full = Path.GetFullPath(file);
                PlainPath(full);
                if (File.Exists(full)) { File.SetAttributes(full, FileAttributes.Normal); File.Delete(full); }
                folders.Add(Path.GetDirectoryName(full));
            }
            foreach (string batch in folders)
            {
                if (Directory.Exists(batch) && !Directory.EnumerateFileSystemEntries(batch).Any()) Directory.Delete(batch);
                string root = Path.GetDirectoryName(batch);
                if ((string.Equals(Path.GetFileName(root), "复制备份", StringComparison.Ordinal) || string.Equals(Path.GetFileName(root), "迁移备份", StringComparison.Ordinal)) && Directory.Exists(root) && !Directory.EnumerateFileSystemEntries(root).Any()) Directory.Delete(root);
            }
        }
        private static void PlainPath(string path)
        {
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if ((File.Exists(current) || Directory.Exists(current)) && (File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new IOException("不通过重解析点替换或备份文件：" + current);
        }
    }
}
