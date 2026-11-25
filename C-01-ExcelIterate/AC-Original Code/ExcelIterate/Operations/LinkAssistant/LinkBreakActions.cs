using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Excel_Iterate
{
    internal sealed partial class LinkAssistantSession
    {
        private sealed class CheckpointOutcome { internal int Success; internal bool Cancelled; internal Exception Error; }
        private int PendingChanges()
        {
            var owners = new HashSet<string>(Batch.Checkpoints.Where(c => c.NeedsRestore).Select(c => c.Owner), StringComparer.OrdinalIgnoreCase);
            return Batch.Changes.Count(c => c.Applied || c.Intent || owners.Contains(c.Owner));
        }
        private LinkWorkbookCheckpoint Checkpoint(string owner) { return Batch.Checkpoints.FirstOrDefault(c => string.Equals(c.Owner, owner, StringComparison.OrdinalIgnoreCase)); }
        private void GuardCheckpoint(LinkWorkbookCheckpoint checkpoint, object book)
        {
            if (book != null && !ComFormulaHost.Boolean(((dynamic)book).Saved, "Saved")) throw new IOException("建立临时保护后工作簿有新编辑，处理已停止。");
            if (XlsxLinkPackage.Hash(checkpoint.Owner) != checkpoint.LastHash) throw new IOException("建立临时保护后文件已变化，处理已停止。");
        }
        private LinkWorkbookCheckpoint CreateCheckpoint(string owner)
        {
            var checkpoint = Checkpoint(owner); if (checkpoint != null) { Batch.ValidateBackup(checkpoint); return checkpoint; }
            if ((File.GetAttributes(owner) & (FileAttributes.ReadOnly | FileAttributes.ReparsePoint)) != 0) throw new IOException("目标为只读文件或重解析点，未修改：" + owner);
            string hash = XlsxLinkPackage.Hash(owner);
            checkpoint = new LinkWorkbookCheckpoint { Owner = owner, BeforeHash = hash, LastHash = hash,
                Backup = Path.Combine(Path.GetDirectoryName(owner), ".EI_LinkBackup_" + Batch.Id + "_" + Guid.NewGuid().ToString("N") + ".xlsx") };
            File.Copy(owner, checkpoint.Backup, false);
            Batch.Checkpoints.Add(checkpoint); Batch.Persist();
            Batch.ValidateBackup(checkpoint); GuardCheckpoint(checkpoint, null);
            Operations.Trace("建立临时保护", owner + "\r\n" + checkpoint.Backup);
            return checkpoint;
        }
        private static HashSet<string> SavedSources(string path)
        { return new HashSet<string>(XlsxLinkPackage.Scan(path).Where(l => l.SourcePath != null).Select(l => l.SourcePath), StringComparer.OrdinalIgnoreCase); }
        private static ExternalLinkInfo MissingSource(List<ExternalLinkInfo> links, string raw)
        {
            var matches = links.Where(link => link.IsMissingPath && string.Equals(link.RawTarget, raw, StringComparison.OrdinalIgnoreCase)).ToList();
            if (matches.Count != 1) throw new IOException("已保存的待替换来源无法唯一识别，请重新扫描或在 WPS 编辑链接中核对：" + raw);
            return matches[0];
        }
        private static string SavedSourceIdentity(ExternalLinkInfo link)
        {
            return link.SourcePath ?? (link.IsMissingPath ? XlsxLinkPackage.AbsoluteFileTarget(link.RawTarget) : null) ?? link.RawTarget;
        }
        internal static string MatchMissingHostName(object rawNames, ExternalLinkInfo source, List<ExternalLinkInfo> links)
        {
            string absolute = XlsxLinkPackage.AbsoluteFileTarget(source.RawTarget);
            string recorded = Uri.UnescapeDataString(source.RawTarget).Replace('/', '\\');
            if (absolute == null && (string.IsNullOrWhiteSpace(recorded) || recorded.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0))
                throw new IOException("待替换来源缺少可核对的完整路径或文件名：" + source.RawTarget);
            if (absolute != null && links.Count(link => string.Equals(SavedSourceIdentity(link), absolute, StringComparison.OrdinalIgnoreCase)) != 1)
                throw new IOException("已保存工作簿含多个同路径来源，无法唯一匹配：" + source.RawTarget);
            string leaf = Path.GetFileName(absolute ?? recorded);
            bool uniqueLeaf = links.Count(link => string.Equals(LinkLeaf(link), leaf, StringComparison.OrdinalIgnoreCase)) == 1;
            var matches = HostLinkNames(rawNames).Distinct(StringComparer.Ordinal).Where(name =>
            {
                string nativePath = XlsxLinkPackage.AbsoluteFileTarget(name);
                if (absolute != null && nativePath != null) return string.Equals(absolute, nativePath, StringComparison.OrdinalIgnoreCase);
                return uniqueLeaf && string.Equals(leaf, nativePath == null ? name : Path.GetFileName(nativePath), StringComparison.OrdinalIgnoreCase);
            }).ToList();
            if (matches.Count != 1) throw new IOException("WPS 没有提供唯一的待替换来源，请在「编辑链接」中核对：" + source.RawTarget);
            return matches[0];
        }
        private static string LinkLeaf(ExternalLinkInfo link)
        {
            string decoded = (link.SourcePath ?? XlsxLinkPackage.AbsoluteFileTarget(link.RawTarget)
                ?? Uri.UnescapeDataString(link.RawTarget)).Replace('/', '\\');
            return decoded.Substring(decoded.LastIndexOf('\\') + 1);
        }
        private static void VerifySavedSources(string path, LinkChange change, HashSet<string> expected)
        {
            var saved = XlsxLinkPackage.ScanReportParts(path);
            if (change.Action == LinkAction.Replace && !saved.Any(link =>
                string.Equals(link.SourcePath ?? (change.DestinationMissing && link.IsMissingPath
                    ? XlsxLinkPackage.AbsoluteFileTarget(link.RawTarget) : null), change.New, StringComparison.OrdinalIgnoreCase)))
                throw new IOException("保存后无法核对新路径。");
            if (!expected.SetEquals(saved.Select(SavedSourceIdentity)))
                throw new IOException("保存后未能确认新路径及其余链接。");
        }
        private void ReplaceCheckpointFile(string stage, string stageHash, string owner, string ownerHash, string phase)
        {
            const int SharingViolation = unchecked((int)0x80070020);
            const int LockViolation = unchecked((int)0x80070021);
            const int UnableToRemoveReplaced = unchecked((int)0x80070497);
            const int MaxAttempts = 5;
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    if (XlsxLinkPackage.Hash(owner) != ownerHash || XlsxLinkPackage.Hash(stage) != stageHash)
                        throw new IOException(phase + "前摘要变化，操作已停止。");
                    File.Replace(stage, owner, null);
                    return;
                }
                catch (IOException error)
                {
                    bool occupied = error.HResult == SharingViolation || error.HResult == LockViolation ||
                        error.HResult == UnableToRemoveReplaced;
                    bool retry = occupied && attempt < MaxAttempts && File.Exists(owner) && File.Exists(stage);
                    Operations.Trace(phase + (retry ? "暂时占用" : "失败"), owner + "，第 " + attempt +
                        " 次，HRESULT=0x" + error.HResult.ToString("X8") + "，" + error.Message);
                    if (!retry) throw;
                    Thread.Sleep(50 << (attempt - 1));
                }
            }
        }
        private void RestoreCheckpoint(LinkWorkbookCheckpoint checkpoint)
        {
            if (checkpoint == null) return;
            if (!checkpoint.NeedsRestore && !Batch.Changes.Any(c => c.Owner == checkpoint.Owner && (c.Applied || c.Intent))) { checkpoint.Restored = true; Batch.Persist(); return; }
            Batch.ValidateBackup(checkpoint);
            string current = XlsxLinkPackage.Hash(checkpoint.Owner);
            using (var bag = new ComBag())
            {
                object book = bag.Own(FindBook(App, checkpoint.Owner));
                if (book != null && (current != checkpoint.BeforeHash || !ComFormulaHost.Boolean(((dynamic)book).Saved, "Saved")))
                    throw new IOException("工作簿仍打开，无法安全回退整簿改动：" + checkpoint.Owner);
            }
            if (checkpoint.CommitIntent && current == checkpoint.PreparedHash)
            { checkpoint.LastHash = current; checkpoint.Applied = true; foreach (var c in Batch.Changes.Where(c => c.Owner == checkpoint.Owner)) c.EverApplied = true; }
            if (current != checkpoint.BeforeHash && current != checkpoint.LastHash)
                throw new IOException("文件已变化或提交状态不明确，无法安全回退：" + checkpoint.Owner);
            if (current != checkpoint.BeforeHash)
            {
                string stage = Path.Combine(Path.GetDirectoryName(checkpoint.Owner), ".EI_LinkStage_" + Guid.NewGuid().ToString("N") + ".xlsx");
                try
                {
                    File.Copy(checkpoint.Backup, stage, false);
                    ReplaceCheckpointFile(stage, checkpoint.BeforeHash, checkpoint.Owner, current, "整簿恢复提交");
                    if (XlsxLinkPackage.Hash(checkpoint.Owner) != checkpoint.BeforeHash) throw new IOException("回退后摘要核对失败。");
                }
                finally { if (File.Exists(stage)) File.Delete(stage); }
            }
            foreach (var c in Batch.Changes.Where(c => c.Owner == checkpoint.Owner)) { c.Applied = c.Intent = false; }
            checkpoint.Applied = checkpoint.Uncertain = checkpoint.CommitIntent = false; checkpoint.Restored = true; checkpoint.LastHash = checkpoint.BeforeHash; Batch.Persist();
        }
        private IEnumerable<string> CheckpointSteps(string owner, bool undo, Func<bool> cancelled, CheckpointOutcome outcome)
        {
            if (undo)
            {
                try
                {
                    var checkpoint = Checkpoint(owner);
                    if (checkpoint == null && Batch.Changes.Any(c => c.Owner == owner && (c.Applied || c.Intent))) throw new IOException("缺少整簿临时保护副本，无法安全回退。");
                    var restored = Batch.Changes.Where(c => c.Owner == owner && (c.Applied || c.Intent || c.EverApplied)).ToList();
                    RestoreCheckpoint(checkpoint);
                    foreach (var c in restored) Pair(c, true, true, null);
                    outcome.Success = restored.Count;
                }
                catch (Exception error) { outcome.Error = error; }
                yield return outcome.Error == null ? "已核对整簿回退：" + Path.GetFileName(owner) : "整簿回退待核对：" + Path.GetFileName(owner);
                yield break;
            }
            LinkWorkbookCheckpoint snapshot = null; List<LinkChange> ordered = null;
            try
            {
                using (var bag = new ComBag()) { var open = bag.Own(FindBook(App, owner)); if (open != null) RequireSaved(open); }
                var actual = SavedSources(owner);
                var changes = Batch.Changes.Where(c => c.Owner == owner).ToList();
                foreach (var change in changes.Where(c => c.Action == LinkAction.Replace))
                    change.DestinationMissing = ValidateReplacementTarget(change.New);
                if (changes.Any(c => c.IsMissingPath))
                {
                    var reportLinks = XlsxLinkPackage.ScanReportParts(owner);
                    foreach (var change in changes.Where(c => c.IsMissingPath)) MissingSource(reportLinks, change.Old);
                    actual.UnionWith(reportLinks.Where(link => link.IsMissingPath).Select(link => link.RawTarget));
                }
                if (changes.Any(c => !actual.Contains(c.Old))) throw new IOException("链接来源已变化，请重新扫描：" + owner);
                ordered = LinkBatch.Order(changes.Where(c => c.Action == LinkAction.Replace), actual, false);
                ordered.AddRange(changes.Where(c => c.Action == LinkAction.Break));
                if (changes.Any(c => c.Action == LinkAction.Break))
                {
                    var names = LinkNameAnalysis.Read(owner);
                    foreach (var change in changes.Where(c => c.Action == LinkAction.Break)) names.Plan(change.Old);
                }
                if (ordered.Count > 0) snapshot = CreateCheckpoint(owner);
            }
            catch (Exception error) { outcome.Error = error; }
            if(outcome.Error!=null || ordered.Count == 0)yield break;
            yield return "链接预检完成：" + Path.GetFileName(owner);
            if (cancelled()) { outcome.Cancelled = true; yield break; }
            using (var bag = new ComBag())
            {
                object original = bag.Own(FindBook(App, owner)), work = original, active = null;
                string stage = null,stageHash=null; bool ownCopy = false,preserveCopy=false;
                try
                {
                    try
                    {
                        GuardCheckpoint(snapshot, original);
                        if (original == null)
                        {
                            stage = Path.Combine(Path.GetDirectoryName(owner), ".EI_LinkStage_" + Guid.NewGuid().ToString("N") + ".xlsx");
                            File.Copy(owner, stage, false);
                            active = bag.Own(((dynamic)App).ActiveWorkbook);
                            dynamic books = bag.Own(((dynamic)App).Workbooks);
                            using (var quiet = new HostQuiet(App)) work = bag.Own(books.Open(stage, 0, false));
                            if (work == null || !string.Equals(Path.GetFullPath(Convert.ToString(((dynamic)work).FullName)), stage, StringComparison.OrdinalIgnoreCase)) throw new IOException("宿主没有返回本次暂存副本，已停止；不会关闭其他工作簿。");
                            ownCopy = true;
                            stageHash=XlsxLinkPackage.Hash(stage);
                        }
                    }
                    catch (Exception error) { outcome.Error = error; }
                    if (outcome.Error != null) yield break;
                    foreach (var change in ordered)
                    {
                        if (cancelled()) { outcome.Cancelled = true; break; }
                        try
                        {
                            if (original != null) GuardCheckpoint(snapshot, work); else GuardCheckpoint(snapshot, null);
                            if(original==null&&(!ComFormulaHost.Boolean(((dynamic)work).Saved,"Saved")||XlsxLinkPackage.Hash(stage)!=stageHash))
                            {preserveCopy=true;throw new IOException("暂存副本有新增编辑，已保留副本且未提交原文件："+stage);}
                            string path = original == null ? stage : owner;
                            var saved = XlsxLinkPackage.ScanReportParts(path);
                            var expected = new HashSet<string>(saved.Select(SavedSourceIdentity), StringComparer.OrdinalIgnoreCase);
                            string sourceName;
                            string formulaSource = change.Old;
                            if (change.IsMissingPath)
                            {
                                var old = MissingSource(saved, change.Old);
                                sourceName = MatchMissingHostName((object)((dynamic)work).LinkSources(1), old, saved);
                                expected.Remove(SavedSourceIdentity(old));
                                formulaSource = XlsxLinkPackage.Resolve(path, old.RawTarget);
                            }
                            else
                            {
                                var sources = ActiveHostSources(work, path, s => Operations.Trace("宿主来源", s));
                                var source = sources.SingleOrDefault(s => string.Equals(s.Path, change.Old, StringComparison.OrdinalIgnoreCase));
                                if (source == null) throw new IOException("不能唯一定位待处理来源：" + change.Old);
                                sourceName = source.Name;
                                expected.Remove(change.Old);
                            }
                            if (change.Action == LinkAction.Replace)
                            {
                                change.DestinationMissing = ValidateReplacementTarget(change.New);
                                if (!expected.Add(change.New)) throw new IOException("新链接与另一来源冲突，已停止该工作簿：" + change.New);
                            }
                            var formulaCheck = change.DestinationMissing ? LinkNameAnalysis.Read(path) : null;
                            LinkNameAnalysis nameAnalysis = null; List<LinkDefinedName> namePlan = null;
                            List<object> nativeNames = null;
                            using (var nameReferences = new ComBag())
                            {
                                if (change.Action == LinkAction.Break)
                                {
                                    nameAnalysis = LinkNameAnalysis.Read(path); namePlan = nameAnalysis.Plan(change.Old);
                                    nativeNames = nameAnalysis.NativeNames(work, namePlan, nameReferences);
                                }
                                GuardCheckpoint(snapshot, work);
                                if (original == null && XlsxLinkPackage.Hash(stage) != stageHash)
                                { preserveCopy = true; throw new IOException("名称读取期间暂存副本发生新编辑，已保留副本且未提交原文件。"); }
                                if (original != null) { change.Intent = true; snapshot.Uncertain = true; Batch.Persist(); }
                                using (var quiet = new HostQuiet(App))
                                {
                                    if (change.Action == LinkAction.Break)
                                    {
                                        foreach (object name in nativeNames) ((dynamic)name).Delete();
                                        if (namePlan.Count > 0) Operations.Trace("清理来源名称", change.Old + "；" + namePlan.Count + " 个未使用名称");
                                        // Name deletion can remove the native source immediately;
                                        // match fresh raw names without requiring the old package to agree.
                                        var current = ParseHostSources((object)((dynamic)work).LinkSources(1), path, XlsxLinkPackage.Scan(path))
                                            .Where(s => string.Equals(s.Path, change.Old, StringComparison.OrdinalIgnoreCase)).GroupBy(s => s.Name, StringComparer.Ordinal).Select(g => g.First()).ToList();
                                        if (current.Count > 1) throw new IOException("清理名称后宿主来源不唯一。");
                                        if (current.Count == 1) { Operations.Trace("调用BreakLink", "Name=" + current[0].Name + "；Type=1"); ((dynamic)work).BreakLink(current[0].Name, 1); }
                                    }
                                    else
                                    {
                                        Operations.Trace("调用ChangeLink", "Name=" + sourceName + "；NewName=" + change.New + "；Type=1");
                                        ((dynamic)work).ChangeLink(sourceName, change.New, 1);
                                    }
                                    WorkbookSaving.Save(work, true);
                                }
                                if (original != null) { snapshot.LastHash = XlsxLinkPackage.Hash(owner); snapshot.Applied = true; snapshot.Uncertain = false; change.EverApplied |= snapshot.LastHash != snapshot.BeforeHash; Batch.Persist(); }
                                else stageHash=XlsxLinkPackage.Hash(stage);
                                VerifySavedSources(path, change, expected);
                                if (formulaCheck != null) formulaCheck.VerifyReplacement(LinkNameAnalysis.Read(path), formulaSource, change.New);
                                if (nameAnalysis != null) nameAnalysis.Verify(LinkNameAnalysis.Read(path), change.Old, namePlan);
                                if (original != null) { change.Applied = change.EverApplied = true; change.Intent = false; Batch.Persist(); outcome.Success++; Pair(change, false, true, null); }
                            }
                        }
                        catch (Exception error) { outcome.Error = error; break; }
                        yield return (original == null ? "已处理暂存副本：" : "已保存链接操作：") + Path.GetFileName(owner);
                    }
                    if (original == null && outcome.Error == null && !outcome.Cancelled)
                    {
                        try
                        {
                            if (cancelled()) { outcome.Cancelled = true; }
                            else
                            {
                                if(!ComFormulaHost.Boolean(((dynamic)work).Saved,"Saved")||XlsxLinkPackage.Hash(stage)!=stageHash)
                                {preserveCopy=true;throw new IOException("暂存副本提交前出现新增编辑，已保留副本且未提交原文件："+stage);}
                                ((dynamic)work).Close(false); ownCopy = false;
                                GuardCheckpoint(snapshot, null);
                                snapshot.PreparedHash = stageHash; snapshot.CommitIntent = true;
                                foreach (var c in ordered) c.Intent = true; Batch.Persist();
                                ReplaceCheckpointFile(stage, snapshot.PreparedHash, owner, snapshot.LastHash, "暂存提交");
                                stage = null;
                                if (XlsxLinkPackage.Hash(owner) != snapshot.PreparedHash) throw new IOException("暂存提交后摘要核对失败。");
                                snapshot.LastHash = snapshot.PreparedHash; snapshot.CommitIntent = false; snapshot.Applied = true;
                                foreach (var c in ordered) { c.Applied = c.EverApplied = true; c.Intent = false; }
                                Batch.Persist(); outcome.Success = ordered.Count;
                                foreach (var c in ordered) Pair(c, false, true, null);
                            }
                        }
                        catch (Exception error) { outcome.Error = error; }
                    }
                }
                finally
                {
                    if (ownCopy && !preserveCopy)
                    {
                        try
                        {
                            if (!ComFormulaHost.Boolean(((dynamic)work).Saved, "Saved") || stage == null || XlsxLinkPackage.Hash(stage) != stageHash)
                            { preserveCopy = true; Operations.Record("暂存副本已保留", "副本存在未保存编辑或内容变化，未关闭、未提交原文件：" + stage); }
                        }
                        catch (Exception error) { preserveCopy = true; if (outcome.Error == null) outcome.Error = error; Operations.Trace("暂存副本核对失败", stage + "；" + error); }
                    }
                    if (ownCopy&&!preserveCopy) try { ((dynamic)work).Close(false); } catch (Exception error) { if (outcome.Error == null) outcome.Error = error;preserveCopy=true; }
                    if (active != null) try { ((dynamic)active).Activate(); } catch { }
                    if (stage != null && File.Exists(stage)&&!preserveCopy) try { File.Delete(stage); } catch (Exception error) { Operations.Trace("暂存副本保留", stage + "；" + error.GetType().Name); }
                }
            }
        }
    }
}
