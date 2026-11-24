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
        [XmlIgnore] public string Journal {get {return Path.Combine(Path.GetDirectoryName(ReportBook),".EI_LinkBatch_"+Key(ReportBook)+".xml");}}
        [XmlIgnore] internal bool Pending {get {return State!="Scanned";}}
        [XmlIgnore] internal bool CanUndo {get {return State!="Cleaning"&&State!="Undone"&&(Changes.Any(c=>c.Applied||c.Intent)||Checkpoints.Any(c=>c.NeedsRestore));}}
        private static string Key(string path)
        {using(var sha=System.Security.Cryptography.SHA256.Create())return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Path.GetFullPath(path).ToUpperInvariant()))).Replace("-","").Substring(0,24);}
        internal static LinkBatch Load(string report)
        {
            var model=new LinkBatch {ReportBook=report};if(!File.Exists(model.Journal))return null;
            using(var reader=XmlReader.Create(model.Journal,new XmlReaderSettings {DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null}))model=(LinkBatch)new XmlSerializer(typeof(LinkBatch)).Deserialize(reader);
            if(!string.Equals(model.ReportBook,report,StringComparison.OrdinalIgnoreCase))throw new IOException("链接恢复记录工作簿身份不符。");
            return model;
        }
        internal void Persist()
        {
            string temp=Journal+".tmp";using(var writer=XmlWriter.Create(temp,new XmlWriterSettings {Encoding=new UTF8Encoding(false),Indent=true}))new XmlSerializer(typeof(LinkBatch)).Serialize(writer,this);
            // 已进回收站的旧记录仍被 File.Exists 判为存在，Replace 会因"无法删除要被替换的文件"失败，退化为删除后移动。
            if(File.Exists(Journal)){try{File.Replace(temp,Journal,null);}catch(IOException){File.Delete(Journal);File.Move(temp,Journal);}}else File.Move(temp,Journal);
        }
        internal string LogPath(string dir){return Path.Combine(dir,".EI_LinkBatch_"+Id+".log");}
        internal void Log(string text){foreach(string dir in LogDirectories)LogTo(dir,text);}
        internal void LogTo(string dir,string text){File.AppendAllText(LogPath(dir),text,new UTF8Encoding(false));}
        internal string[] Candidates()
        {
            if(ScopeFolder==null)return Owners.ToArray();
            return Directory.GetFiles(ScopeFolder,"*.xlsx").Where(p=>!Path.GetFileName(p).StartsWith("~$")&&!OperationBatch.ManagedFile(p)).OrderBy(p=>p,StringComparer.OrdinalIgnoreCase).ToArray();
        }
        internal void Confirm(Action<string> recycle)
        {
            if(State=="UndoFailed"||Changes.Any(c=>c.Intent)||Checkpoints.Any(c=>c.Uncertain||c.CommitIntent))throw new InvalidOperationException("链接还原或提交核对未完成，禁止清理恢复记录。");
            State="Cleaning";Persist();
            foreach(var checkpoint in Checkpoints)if(File.Exists(checkpoint.Backup)){ValidateBackup(checkpoint);recycle(checkpoint.Backup);}
            foreach(string dir in LogDirectories)if(File.Exists(LogPath(dir)))recycle(LogPath(dir));
            if(File.Exists(Journal+".tmp"))recycle(Journal+".tmp");if(File.Exists(Journal))recycle(Journal);
        }
        internal void ValidateBackup(LinkWorkbookCheckpoint checkpoint)
        {
            string owner=Path.GetFullPath(checkpoint.Owner),backup=Path.GetFullPath(checkpoint.Backup);
            if(!string.Equals(Path.GetDirectoryName(owner),Path.GetDirectoryName(backup),StringComparison.OrdinalIgnoreCase)
                || !Path.GetFileName(backup).StartsWith(".EI_LinkBackup_"+Id+"_",StringComparison.Ordinal)
                || Path.GetExtension(backup)!=".xlsx" || (File.GetAttributes(backup)&FileAttributes.ReparsePoint)!=0
                || XlsxLinkPackage.Hash(backup)!=checkpoint.BeforeHash)throw new IOException("断链备份身份或摘要不符，已保留恢复记录。");
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
