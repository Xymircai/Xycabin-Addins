using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace Excel_Iterate
{
    // Public DTOs are required by XmlSerializer; no Ribbon API is changed.
    public sealed class BatchFile
    {
        public string Source, Current, Next, Hash;
        public long Length, Modified;
        internal static BatchFile From(FileSnapshot file)
        { return new BatchFile { Source=file.Path, Current=file.Path, Hash=file.Hash, Length=file.Length, Modified=file.Modified.Ticks }; }
        internal FileSnapshot Snapshot()
        { return new FileSnapshot { Path=Source, Name=Path.GetFileName(Source), Hash=Hash, Length=Length, Modified=new DateTime(Modified,DateTimeKind.Utc) }; }
        internal void VerifyCurrent()
        { var now=FileSnapshot.Read(Current); if(now.Hash!=Hash || now.Length!=Length || now.Modified.Ticks!=Modified) throw new IOException("文件已被外部修改，暂停还原："+Current); }
    }
    public sealed class BatchDirectory
    {
        public string Path, Identity;
        public bool Done;
        internal static BatchDirectory Read(string path)
        { return new BatchDirectory { Path=System.IO.Path.GetFullPath(path),Identity=DirectoryIdentity.Read(path) }; }
        internal void Verify()
        { if(Identity==null || DirectoryIdentity.Read(Path)!=Identity)throw new IOException("目录身份已变化，保留并暂停清理："+Path); }
    }
    internal static class DirectoryIdentity
    {
        [StructLayout(LayoutKind.Sequential)] private struct Info
        { internal uint Attributes,CreatedLow,CreatedHigh,AccessLow,AccessHigh,WrittenLow,WrittenHigh,Volume,SizeHigh,SizeLow,Links,IndexHigh,IndexLow; }
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] private static extern SafeFileHandle CreateFile(string path,uint access,uint share,IntPtr security,uint creation,uint flags,IntPtr template);
        [DllImport("kernel32.dll",SetLastError=true)] private static extern bool GetFileInformationByHandle(SafeFileHandle handle,out Info info);
        [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] internal static extern bool CreateDirectory(string path,IntPtr security);
        internal static string Read(string path)
        {
            FilePaths.CheckAncestors(path);
            using(var handle=CreateFile(path,0,7,IntPtr.Zero,3,0x02200000,IntPtr.Zero))
            {
                Info info;if(handle.IsInvalid || !GetFileInformationByHandle(handle,out info))throw new IOException("无法核对目录身份："+path,new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error()));
                if((info.Attributes&0x10)==0 || (info.Attributes&0x400)!=0)throw new IOException("不处理目录联接或替代路径："+path);
                return info.Volume.ToString("X8")+":"+info.IndexHigh.ToString("X8")+info.IndexLow.ToString("X8")+":"+info.CreatedHigh.ToString("X8")+info.CreatedLow.ToString("X8");
            }
        }
    }
    public sealed class OperationBatch
    {
        public string Id, Kind, Root, State="Active", DocumentPath, Anchor, BackupRoot;
        public bool BackupContainerCreated;
        public List<BatchFile> Files=new List<BatchFile>();
        public List<BatchFile> ReportFiles=new List<BatchFile>();
        public List<string> CreatedDirectories=new List<string>();
        public List<BatchDirectory> CreatedDirectoryRecords=new List<BatchDirectory>();
        public List<BatchDirectory> OriginalDirectories=new List<BatchDirectory>();
        public List<string> ProtectedDirectories=new List<string>();
        public bool ConfirmOriginalDirectories;
        [XmlIgnore] public string Journal { get { return Path.Combine(Root,".EI_Operation_"+Id+".xml"); } }
        [XmlIgnore] public string LogPath { get { return Path.Combine(Root,".EI_Operation_"+Id+".log"); } }
        [XmlIgnore] internal bool CanUndo { get { return State!="Cleaning" && State!="Undone" && (Files.Any(f=>!string.Equals(f.Current,f.Source,StringComparison.Ordinal)) || CreatedDirectoryRecords.Any(d=>!d.Done) || Kind=="Organize"&&OriginalDirectories.Any(d=>!d.Done)); } }
        internal static OperationBatch Create(string root,string kind)
        {
            root=FilePaths.Folder(root);
            if(Find(root)!=null) throw new InvalidOperationException("该目录有待确认批次，请先取消还原或确认清理。");
            var batch=new OperationBatch { Id=Guid.NewGuid().ToString("N"),Root=root,Kind=kind };
            batch.Persist(); batch.Log("开始"+kind+"；取消操作还原，确认操作清理。"); return batch;
        }
        internal static OperationBatch Find(string root)
        {
            if(!Directory.Exists(root)) return null;
            var paths=Directory.GetFiles(root,".EI_Operation_*.xml");
            if(paths.Length==0)return null;
            if(paths.Length>1)throw new IOException("发现多个未确认批次，请保留恢复文件并核对："+root);
            OperationBatch batch;
            using(var reader=XmlReader.Create(paths[0],new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null }))
                batch=(OperationBatch)new XmlSerializer(typeof(OperationBatch)).Deserialize(reader);
            if(batch.Id==null || batch.Id.Length!=32 || !batch.Id.All(Uri.IsHexDigit) || !string.Equals(batch.Root,Path.GetFullPath(root),StringComparison.OrdinalIgnoreCase) || !string.Equals(paths[0],batch.Journal,StringComparison.OrdinalIgnoreCase)) throw new IOException("恢复记录身份不符，暂停处理。");
            if(!new[]{"Active","Completed","Undone","UndoFailed","Cleaning"}.Contains(batch.State)||!new[]{"Organize","Rename"}.Contains(batch.Kind))throw new IOException("恢复记录状态不符，暂停处理。");
            foreach(string dir in batch.CreatedDirectories)batch.Within(dir);
            foreach(string dir in batch.ProtectedDirectories)batch.Within(dir);
            foreach(var dir in batch.CreatedDirectoryRecords)batch.Within(dir.Path);
            foreach(var dir in batch.OriginalDirectories)
            {batch.Within(dir.Path);if(!string.Equals(Path.GetDirectoryName(dir.Path),batch.Root,StringComparison.OrdinalIgnoreCase)||batch.ProtectedDirectories.Contains(dir.Path,StringComparer.OrdinalIgnoreCase)||Path.GetFileName(dir.Path)=="备份")throw new IOException("原目录清理范围不符，暂停处理。");}
            foreach(var file in batch.Files)
            {
                batch.Within(file.Source);batch.Within(file.Current);
                if(file.Next!=null)
                {
                    batch.Within(file.Next);
                    if(!File.Exists(file.Current)&&File.Exists(file.Next)) { file.Current=file.Next;file.VerifyCurrent();file.Next=null; }
                    else if(File.Exists(file.Current)&&!File.Exists(file.Next)) { file.VerifyCurrent();file.Next=null; }
                    else throw new IOException("移动中断状态不明确，请保留记录："+file.Current+" → "+file.Next);
                }
            }
            return batch;
        }
        internal void Within(string path)
        {
            if(!Path.GetFullPath(path).StartsWith(Root.TrimEnd('\\')+"\\",StringComparison.OrdinalIgnoreCase))throw new IOException("恢复路径越出所选目录："+path);
            FilePaths.CheckAncestors(path);
        }
        internal void Persist()
        { Persist((temporary,journal)=>File.Replace(temporary,journal,null),System.Threading.Thread.Sleep); }
        internal void Persist(Action<string,string> replace,Action<int> pause)
        {
            string temp=Journal+".tmp";
            using(var writer=XmlWriter.Create(temp,new XmlWriterSettings { Encoding=new UTF8Encoding(false),Indent=true })) new XmlSerializer(typeof(OperationBatch)).Serialize(writer,this);
            if(!File.Exists(Journal)) {File.Move(temp,Journal);return;}
            // Retry only failures that leave both names intact; never delete the last recovery record.
            for(int attempt=1;;attempt++)
            {
                try {replace(temp,Journal);return;}
                catch(IOException ex)
                {
                    bool transient=ex.HResult==unchecked((int)0x80070020)
                        ||ex.HResult==unchecked((int)0x80070021)||ex.HResult==unchecked((int)0x80070497);
                    if(transient&&attempt<5&&File.Exists(temp)&&File.Exists(Journal))
                    {
                        pause(50<<(attempt-1));
                        if(File.Exists(temp)&&File.Exists(Journal))continue;
                    }
                    throw new IOException("恢复记录原子替换失败（尝试"+attempt+"次，HRESULT=0x"+ex.HResult.ToString("X8")
                        +"）；恢复记录："+Journal+"；临时记录："+temp+"。"+ex.Message,ex);
                }
            }
        }
        internal void Log(string value) { File.AppendAllText(LogPath,DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")+"  "+value+"\r\n",new UTF8Encoding(false)); }
        internal void DirectoryFor(string path)
        {
            Within(path); var missing=new Stack<string>();string dir=Path.GetDirectoryName(path);
            while(!Directory.Exists(dir)) { Within(dir);missing.Push(dir);dir=Path.GetDirectoryName(dir); }
            while(missing.Count>0)
            {
                dir=missing.Pop();var record=new BatchDirectory {Path=dir};CreatedDirectoryRecords.Add(record);Persist();
                if(!DirectoryIdentity.CreateDirectory(dir,IntPtr.Zero))
                {int error=Marshal.GetLastWin32Error();record.Done=true;Persist();if(error!=183)throw new IOException("创建目录失败："+dir,new System.ComponentModel.Win32Exception(error));FilePaths.CheckAncestors(dir);continue;}
                record.Identity=DirectoryIdentity.Read(dir);CreatedDirectories.Add(dir);Persist();
            }
        }
        internal BatchFile Track(FileSnapshot file)
        { var entry=BatchFile.From(file);Files.Add(entry);Persist();return entry; }
        internal void Move(BatchFile file,string target,Action<string,string> move)
        {
            Within(file.Current);Within(target);file.VerifyCurrent();DirectoryFor(target);
            file.Next=target;Persist();
            try {move(file.Current,target);file.Current=target;file.Next=null;Persist();Log(file.Source+" → "+target);}
            catch {if(File.Exists(file.Current)&&!File.Exists(target)){file.Next=null;Persist();}throw;}
        }
        internal FileOperationResult Undo(Action<string,string> move,Action<string> deleteDirectory=null)
        {
            var result=new FileOperationResult { Batch=this,RecoveryPath=Journal };move=move??FilePaths.Move;
            if(State=="Cleaning")throw new InvalidOperationException("已开始确认清理，不能取消操作。");
            try
            {
                foreach(var dir in OriginalDirectories)if(!dir.Done)dir.Verify();
                var changed=Files.Where(f=>f.Current!=f.Source).ToList();
                var current=new HashSet<string>(changed.Select(f=>f.Current),StringComparer.OrdinalIgnoreCase);
                foreach(var f in changed) {Within(f.Source);Within(f.Current);f.VerifyCurrent();if(Directory.Exists(f.Source)||(File.Exists(f.Source)&&!current.Contains(f.Source)))throw new IOException("还原目标已被占用："+f.Source);using(File.Open(f.Current,FileMode.Open,FileAccess.Read,FileShare.None)){} }
                // Stage the whole inverse batch so swaps and cycles remain reversible.
                foreach(var f in changed)Move(f,Path.Combine(Root,".EI_Restore_"+Guid.NewGuid().ToString("N")+".tmp"),move);
                foreach(var f in changed) {Move(f,f.Source,move);result.Success++;result.Record(f.Source,"已还原并核对");}
                foreach(var dir in CreatedDirectoryRecords.Where(d=>!d.Done).OrderByDescending(d=>d.Path.Length))
                {
                    Within(dir.Path);
                    if(Directory.Exists(dir.Path))
                    {dir.Verify();if(Directory.EnumerateFileSystemEntries(dir.Path).Any())throw new IOException("新建目录存在其他内容，保留并暂停还原："+dir.Path);if(deleteDirectory==null)Directory.Delete(dir.Path,false);else deleteDirectory(dir.Path);if(Directory.Exists(dir.Path))throw new IOException("新建目录未移除："+dir.Path);result.Log.AppendLine("已移除本批新建目录："+dir.Path);}
                    else if(File.Exists(dir.Path))throw new IOException("新建目录路径被文件占用："+dir.Path);
                    dir.Done=true;Persist();
                }
                // Legacy records have no reliable directory identity: never infer ownership.
                if(CreatedDirectories.Except(CreatedDirectoryRecords.Select(d=>d.Path),StringComparer.OrdinalIgnoreCase).Any(Directory.Exists))result.Log.AppendLine("旧批次缺少目录身份，文件已还原；原记录目录保留供人工核对。");
                State="Undone";Persist();Log("本批改动已全部还原；确认操作清理记录。");
            }
            catch(Exception ex) {State="UndoFailed";result.Failed++;result.RecoveryIncomplete=true;result.Log.AppendLine("还原未完成："+ex.Message+"；当前路径见恢复记录。");Persist();Log(result.Text);}
            return result;
        }
        internal void Confirm(Action<string> recycle)
        {
            if(State=="UndoFailed")throw new InvalidOperationException("还原未完成，禁止清理恢复记录。");
            if(State!="Cleaning")
            {
                foreach(var f in Files)if(f.Current!=f.Source){Within(f.Current);f.VerifyCurrent();}
                ConfirmOriginalDirectories=Kind=="Organize"&&State!="Undone";State="Cleaning";Persist();
            }
            if(ConfirmOriginalDirectories)foreach(var dir in OriginalDirectories.Where(d=>!d.Done))
            {
                Within(dir.Path);
                if(Directory.Exists(dir.Path))
                {dir.Verify();CheckTree(dir.Path);recycle(dir.Path);if(Directory.Exists(dir.Path))throw new IOException("原目录未回收："+dir.Path);}
                else if(File.Exists(dir.Path))throw new IOException("原目录路径被文件占用："+dir.Path);
                dir.Done=true;Persist();Log("已回收原目录："+dir.Path);
            }
            if(BackupRoot!=null&&Directory.Exists(BackupRoot))
            {
                Within(BackupRoot);
                string container=Path.Combine(Root,"备份")+"\\";
                if(!BackupRoot.StartsWith(container+".EI_Backup_",StringComparison.OrdinalIgnoreCase)||Path.GetDirectoryName(BackupRoot)!=container.TrimEnd('\\'))throw new IOException("备份目录身份不符，暂停清理。");
                foreach(string dir in Directory.GetDirectories(BackupRoot,"*",SearchOption.AllDirectories))FilePaths.CheckAncestors(dir);
                // Never recycle unexpected content inserted into an owned backup.
                var expected=new HashSet<string>(Files.Where(f=>f.Current.StartsWith(BackupRoot+"\\",StringComparison.OrdinalIgnoreCase)).Select(f=>f.Current),StringComparer.OrdinalIgnoreCase);
                if(Directory.GetFiles(BackupRoot,"*",SearchOption.AllDirectories).Any(p=>!expected.Contains(p)))throw new IOException("备份中出现其他文件，保留并暂停清理："+BackupRoot);
                recycle(BackupRoot);
            }
            if(BackupContainerCreated)
            {string dir=Path.Combine(Root,"备份");if(Directory.Exists(dir)&&!Directory.EnumerateFileSystemEntries(dir).Any())recycle(dir);}
            if(File.Exists(LogPath))recycle(LogPath);
            if(File.Exists(Journal+".tmp"))recycle(Journal+".tmp");
            recycle(Journal); // Last: a failed cleanup remains discoverable and retryable.
        }
        internal string ConfirmationText()
        {
            var text=new StringBuilder("确认后不再提供本批还原；程序生成的备份、日志及恢复记录将送入回收站。");
            if(Kind=="Organize"&&State!="Undone"&&OriginalDirectories.Any(d=>!d.Done))
            {text.AppendLine("\r\n以下原目录及其中剩余未处理文件、后来新增内容将整体送入回收站；分类结果与原有备份内容保留：");foreach(var dir in OriginalDirectories.Where(d=>!d.Done))text.AppendLine(dir.Path);}
            return text.ToString();
        }
        private static void CheckTree(string root)
        {
            var pending=new Stack<string>();pending.Push(root);
            while(pending.Count>0)
            {string dir=pending.Pop();FilePaths.CheckAncestors(dir);foreach(string item in Directory.EnumerateFileSystemEntries(dir)){var attributes=File.GetAttributes(item);if((attributes&FileAttributes.ReparsePoint)!=0)throw new IOException("原目录含联接或符号链接，保留并暂停回收："+item);if((attributes&FileAttributes.Directory)!=0)pending.Push(item);}}
        }
        internal static bool ManagedFile(string path)
        {return System.Text.RegularExpressions.Regex.IsMatch(Path.GetFileName(path),@"^\.EI_(?:Operation_[a-f0-9]{32}\.(?:xml|log)(?:\.tmp)?|(?:Restore|Rename)_[a-f0-9]{32}\.tmp|LinkBatch_[a-f0-9]{24,32}\.(?:xml|log)(?:\.tmp)?|LinkBackup_[a-f0-9]{32}_[a-f0-9]{32}\.xlsx|LinkStage_[a-f0-9]+\.xlsx)$",System.Text.RegularExpressions.RegexOptions.IgnoreCase);}
    }
}
