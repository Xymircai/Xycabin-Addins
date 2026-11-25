using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Excel_Iterate
{
    public sealed class LinkRowRecord
    {
        public string Owner, Source, Raw, Status;
        public bool IsMissingPath;
        internal static LinkRowRecord FromLink(ExternalLinkInfo link)
        {
            return new LinkRowRecord { Owner=link.OwnerPath, Source=link.SourcePath, Raw=link.RawTarget, Status=link.Status, IsMissingPath=link.IsMissingPath };
        }
        internal ExternalLinkInfo Link(){return new ExternalLinkInfo {OwnerPath=Owner,SourcePath=Source,RawTarget=Raw,Status=Status??"待填写新链接",IsMissingPath=IsMissingPath};}
    }
    public enum LinkAction { Replace = 0, Break = 1 }
    public sealed class LinkChange
    {
        public string Owner, Old, New;
        public string ReportSource;
        public string FailureReason;
        public bool Applied, Intent, IntentReverse, EverApplied;
        public bool IsMissingPath, DestinationMissing;
        internal string ReportSourceId { get { return ReportSource ?? Old; } }
        public LinkAction Action;
    }
    public sealed class LinkWorkbookCheckpoint
    {
        public string Owner, Backup, BeforeHash, LastHash, PreparedHash;
        public bool Applied, Uncertain, CommitIntent, Restored;
        [XmlIgnore] internal bool NeedsRestore { get { return !Restored && (Applied || Uncertain || CommitIntent); } }
    }
    public sealed class LinkBatch
    {
        public string ReportBook, Anchor, ReportName, ScopeFolder, State="Scanned", Id=Guid.NewGuid().ToString("N");
        public string LastSummary, UserLog;
        public int Skipped;
        public List<string> Owners=new List<string>(),LogDirectories=new List<string>();
        public List<LinkRowRecord> Rows=new List<LinkRowRecord>();
        public List<LinkChange> Changes=new List<LinkChange>();
        public List<LinkWorkbookCheckpoint> Checkpoints=new List<LinkWorkbookCheckpoint>();
        [XmlIgnore] public string Journal { get { return JournalPath(ReportBook, false); } }

        private static string JournalPath(string report, bool legacy)
        {
            string identity = Path.GetFullPath(report).ToUpperInvariant();
            if (!legacy) identity = "v2\n" + identity;
            using (var sha = System.Security.Cryptography.SHA256.Create())
            {
                string key = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(identity))).Replace("-", "").Substring(0, 24);
                return Path.Combine(Path.GetDirectoryName(report), ".EI_LinkBatch_" + key + ".xml");
            }
        }

        internal static LinkBatch Load(string report)
        {
            string current = JournalPath(report, false);
            if (File.Exists(current)) return Read(current, report);

            // A pre-upgrade scan can be reused. Pending recovery records belong to the old workflow.
            string legacy = JournalPath(report, true);
            if (!File.Exists(legacy)) return null;
            try
            {
                var batch = Read(legacy, report);
                if (batch.State != "Scanned" || batch.Changes.Count != 0 || batch.Checkpoints.Count != 0) return null;
                batch.Id = Guid.NewGuid().ToString("N");
                return batch;
            }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
            catch (InvalidOperationException) { return null; }
            catch (XmlException) { return null; }
        }

        private static LinkBatch Read(string path, string report)
        {
            LinkBatch batch;
            using (var reader = XmlReader.Create(path, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                batch = (LinkBatch)new XmlSerializer(typeof(LinkBatch)).Deserialize(reader);
            Guid id;
            if (!string.Equals(batch.ReportBook, report, StringComparison.OrdinalIgnoreCase) ||
                !Guid.TryParseExact(batch.Id, "N", out id))
                throw new IOException("链接记录的工作簿身份或批次编号不符。");
            return batch;
        }

        internal void Persist()
        {
            ReplaceJournal(Journal);
        }

        private void ReplaceJournal(string path)
        {
            string temp = path + ".tmp";
            using (var writer = XmlWriter.Create(temp, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = true }))
                new XmlSerializer(typeof(LinkBatch)).Serialize(writer, this);
            if (File.Exists(path)) File.Replace(temp, path, null);
            else File.Move(temp, path);
        }

        internal void PreserveUnresolvedRecord()
        {
            if (State != "Active" && State != "UndoFailed" && Checkpoints.Count == 0) return;
            string archive = Path.Combine(Path.GetDirectoryName(ReportBook), ".EI_LinkBatch_" + Id + ".xml");
            if (!File.Exists(archive)) ReplaceJournal(archive);
        }

        internal string LogPath(string dir){return Path.Combine(dir,".EI_LinkBatch_"+Id+".log");}
        internal void LogTo(string dir,string text){File.AppendAllText(LogPath(dir),text,new UTF8Encoding(false));}
        internal string[] Candidates()
        {
            int skipped;
            return ScopeFolder == null ? Owners.ToArray() : LinkScanSelection.FolderCandidates(ScopeFolder, out skipped);
        }

        internal void RemoveTemporaryCopy(LinkWorkbookCheckpoint checkpoint)
        {
            if (File.Exists(checkpoint.Backup))
            {
                ValidateBackup(checkpoint);
                File.Delete(checkpoint.Backup);
            }
            Checkpoints.Remove(checkpoint);
        }
        internal void ValidateBackup(LinkWorkbookCheckpoint checkpoint)
        {
            string owner=Path.GetFullPath(checkpoint.Owner),backup=Path.GetFullPath(checkpoint.Backup);
            if(!string.Equals(Path.GetDirectoryName(owner),Path.GetDirectoryName(backup),StringComparison.OrdinalIgnoreCase)
                || !Path.GetFileName(backup).StartsWith(".EI_LinkBackup_"+Id+"_",StringComparison.Ordinal)
                || Path.GetExtension(backup)!=".xlsx" || (File.GetAttributes(backup)&FileAttributes.ReparsePoint)!=0
                || XlsxLinkPackage.Hash(backup)!=checkpoint.BeforeHash)throw new IOException("临时保护副本的身份或摘要不符，已保留原文件。");
        }
        // ChangeLink is sequential. Reverse topological order avoids merging a
        // chain's intermediate source; cycles require aliases and are refused.
        internal static List<LinkChange> Order(IEnumerable<LinkChange> entries,IEnumerable<string> existing,bool reverse)
        {
            var list=entries.ToList();Func<LinkChange,string> from=c=>reverse?c.New:c.Old,to=c=>reverse?c.Old:c.New;
            var keys=new HashSet<string>(list.Select(from),StringComparer.OrdinalIgnoreCase);
            if(keys.Count!=list.Count||list.Select(to).Distinct(StringComparer.OrdinalIgnoreCase).Count()!=list.Count)throw new InvalidOperationException("多个来源合并到同一链接，无法可靠还原；暂停该工作簿。");
            var untouched=new HashSet<string>(existing.Where(p=>!keys.Contains(p)),StringComparer.OrdinalIgnoreCase);
            if(list.Any(c=>untouched.Contains(to(c))))throw new InvalidOperationException("新链接与未修改来源冲突，无法可靠还原；暂停该工作簿。");
            var ordered=new List<LinkChange>();var remaining=new List<LinkChange>(list);
            while(remaining.Count>0)
            {
                var next=remaining.FirstOrDefault(c=>!remaining.Any(other=>string.Equals(from(other),to(c),StringComparison.OrdinalIgnoreCase)));
                if(next==null)throw new InvalidOperationException("环形链接映射暂不支持；未修改该工作簿。");
                ordered.Add(next);remaining.Remove(next);
            }
            return ordered;
        }
    }
}
