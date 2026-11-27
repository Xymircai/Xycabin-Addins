using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    // STA iterator: UI gets control between bounded formula blocks. The parent
    // on disk changes only after saving the private copy and atomic commit.
    internal sealed class NotesLinkExecution : IDisposable
    {
        private readonly NotesLinkPlan plan;
        private readonly object sourceApp;
        private readonly Action<string> log;
        private readonly ComBag references=new ComBag();
        private HiddenWorkbookApplication hidden;
        private object book,sheet;
        private string stage,batch,manifest,saveCopy,previousStage,phase="准备",failureDetail="";
        private bool closed=true,disposed;
        internal bool Cancelled,Committed;
        internal string Backup="",PreBackup="",CommittedHash="";
        internal int Completed,FormulaWrites,ValueErrors,LinkWarnings,NativePastes;
        internal Func<int,bool> FailureForProbe;
        internal NotesLinkExecution(object application,NotesLinkPlan frozen,Action<string> output)
        {sourceApp=application;plan=frozen;log=output??delegate{};}
        internal static void CheckOpenWorkbooks(object app,IEnumerable<string> sources,string parent,bool requireClosed)
        {
            var relevant=new HashSet<string>(sources.Select(Path.GetFullPath),StringComparer.OrdinalIgnoreCase);relevant.Add(Path.GetFullPath(parent));
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)app).Workbooks);
                for(int i=1;i<=Convert.ToInt32(books.Count);i++)using(var one=new ComBag())
                {
                    dynamic current=one.Own(books.Item(i));string file=Convert.ToString(current.FullName);if(!Path.IsPathRooted(file)||!relevant.Contains(Path.GetFullPath(file)))continue;
                    if(!Convert.ToBoolean(current.Saved))throw new InvalidOperationException("请先手动保存工作簿的修改，再预检："+file);
                    if(requireClosed&&string.Equals(Path.GetFullPath(file),Path.GetFullPath(parent),StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("备份后更新原文件需要父公司工作簿处于关闭状态；请手动关闭后重试，不会自动保存或关闭它。");
                }
            }
        }
        private void Check()
        {
            if(Cancelled)throw new OperationCanceledException("已取消；父公司原文件未交付修改。");
        }
        private void CheckParent(){Check();if(DataMigrationHost.Digest(plan.Config.Parent)!=plan.ParentHash)throw new InvalidOperationException("父公司文件在预检后变化，请重新预检。");}
        private void Record(string state)
        {
            if(manifest==null)return;
            var values=MigrationJson.Obj("state",state,"parent",plan.Config.Parent,"originalHash",plan.ParentHash,"stage",stage,"saveCopy",saveCopy,"previousStage",previousStage,"phase",phase,"failure",failureDetail,"preBackup",PreBackup,"backup",Backup,"committedHash",CommittedHash);
            File.WriteAllText(manifest,MigrationJson.Write(values),new System.Text.UTF8Encoding(false));
        }
        private void Open(bool readOnly)
        {
            using(var one=new ComBag())
            {dynamic books=one.Own(((dynamic)hidden.Application).Workbooks);book=references.Own(books.Open(stage,0,readOnly));if(book==null||(!readOnly&&Convert.ToBoolean(((dynamic)book).ReadOnly)))throw new InvalidOperationException("无法打开Notes工作副本。");closed=false;dynamic sheets=one.Own(((dynamic)book).Worksheets);sheet=references.Own(sheets.Item(plan.TargetSheet));}
        }
        private void Close()
        {if(book!=null&&!closed){((dynamic)book).Close(false);closed=true;book=null;sheet=null;}}
        private void Step(string name,Action action)
        {
            Check();phase=name;Record(name);
            try{action();}
            catch(Exception error)
            {
                failureDetail=error.ToString();try{Record("Failed");}catch(Exception recordError){log("失败记录写入提示："+recordError.Message);}
                var com=error as COMException;string code=com==null?"":"，HRESULT 0x"+unchecked((uint)com.ErrorCode).ToString("X8");
                throw new InvalidOperationException(name+"失败"+code+"："+error.Message+"；原文件未更新，备份："+PreBackup,error);
            }
        }
        internal static bool SaveWithFallback(object workbook,string fallback,Action<string> output)
        {
            try{((dynamic)workbook).Save();return false;}
            catch(COMException error)
            {
                if(error.ErrorCode!=unchecked((int)0x80004005))throw;
                if(output!=null)output("保存返回E_FAIL，尝试将已粘贴内容另存为工作副本…");
                ((dynamic)workbook).SaveCopyAs(fallback);
                if(!File.Exists(fallback)||new FileInfo(fallback).Length==0)throw new IOException("另存工作副本未生成有效文件，原文件未更新。",error);
                return true;
            }
        }
        private void SaveStage()
        {
            saveCopy=Path.Combine(Path.GetDirectoryName(stage),".xy-notes-save-"+Guid.NewGuid().ToString("N")+Path.GetExtension(stage));
            // WPS requires a workbook view during Save. The owned application
            // stays hidden; no yield or user-workbook operation occurs here.
            MigrationWindowScope.SetOwnedVisible(book,true);
            try
            {
                if(SaveWithFallback(book,saveCopy,log)){previousStage=stage;stage=saveCopy;log("另存工作副本完成。");}
            }
            finally
            {
                try{MigrationWindowScope.SetOwnedVisible(book,false);}catch(Exception error){log("隐藏保存窗口提示："+error.Message);}
                try{((dynamic)hidden.Application).Visible=false;}catch(Exception error){log("隐藏后台进程提示："+error.Message);}
            }
        }
        private IEnumerable<MigrationArea> Chunks(MigrationArea area)
        {for(int r=0;r<area.Rows;r+=512)yield return new MigrationArea{Row=area.Row+r,Column=area.Column,Rows=Math.Min(512,area.Rows-r),Columns=area.Columns};}
        private object[,] Matrix(NotesLinkTask task,MigrationArea area)
        {
            return DataMigrationEngine.LinkMatrix(task.File,task.Sheet,plan.Source.Row+area.Row-task.Target.Row,plan.Source.Column+area.Column-task.Target.Column,area.Rows,area.Columns);
        }
        private bool TryNativePaste(NotesLinkTask task)
        {
            if(!task.NativeCompatible||!HiddenWorkbookApplication.IsWps(hidden.Application))return false;
            System.Windows.Forms.IDataObject clipboard;
            try{clipboard=System.Windows.Forms.Clipboard.GetDataObject();}catch(System.Runtime.InteropServices.ExternalException){log("剪贴板暂不可用，改用兼容写入。");return false;}
            object sourceBook=null;bool copied=false;
            using(var bag=new ComBag())
            {
                try
                {
                    dynamic books=bag.Own(((dynamic)hidden.Application).Workbooks);sourceBook=bag.Own(books.Open(task.File,0,true));
                    dynamic tabs=bag.Own(((dynamic)sourceBook).Worksheets),sourceSheet=bag.Own(tabs.Item(task.Sheet));
                    Check();((dynamic)sourceBook).Activate();sourceSheet.Activate();dynamic from=bag.Own(sourceSheet.Range[plan.Source.Address]);copied=true;from.Copy();
                    Check();((dynamic)book).Activate();((dynamic)sheet).Activate();dynamic start=bag.Own(((dynamic)sheet).Range[task.Target.Address]);start.Select();
                    ((dynamic)sheet).Paste(Type.Missing,true);
                    VerifyLinkSamples(task);
                    NativePastes++;FormulaWrites+=task.Blocks.Concat(task.Anchors).Sum(a=>a.Rows*a.Columns);
                    return true;
                }
                finally
                {
                    try{((dynamic)hidden.Application).CutCopyMode=false;}catch(Exception error){log("结束复制状态提示："+error.Message);}
                    try{if(sourceBook!=null)((dynamic)sourceBook).Close(false);}
                    finally{if(copied)try{if(clipboard==null)System.Windows.Forms.Clipboard.Clear();else System.Windows.Forms.Clipboard.SetDataObject(clipboard,true);}catch(Exception error){log("还原剪贴板提示："+error.Message);}}
                }
            }
        }
        private void VerifyLinkSamples(NotesLinkTask task)
        {
            var areas=task.Blocks.Concat(task.Anchors).ToList();
            var samples=areas.Take(1).Concat(areas.Skip(areas.Count/2).Take(1)).Concat(areas.Skip(Math.Max(0,areas.Count-1))).ToList();
            foreach(var area in samples)foreach(bool last in new[]{false,true})using(var bag=new ComBag())
            {
                int row=area.Row+(last?area.Rows-1:0),col=area.Column+(last?area.Columns-1:0);
                dynamic cell=bag.Own(((dynamic)sheet).Range[CellAddress.Name(row,col)]);
                string formula=Convert.ToString(cell.Formula).Replace("$","").Replace("''","'");
                string token="["+Path.GetFileName(task.File)+"]",address="!"+CellAddress.Name(plan.Source.Row+row-task.Target.Row,plan.Source.Column+col-task.Target.Column);
                if(!Convert.ToBoolean(cell.HasFormula)||formula.IndexOf(token,StringComparison.OrdinalIgnoreCase)<0||!formula.EndsWith(address,StringComparison.OrdinalIgnoreCase))throw new InvalidOperationException("链接未完整替换："+CellAddress.Name(row,col)+"；仍未指向本次来源文件，原文件未更新。");
            }
        }
        private void RestoreGeometry()
        {
            var rows=new SortedDictionary<int,double>();
            foreach(var task in plan.Tasks)
            {for(int r=task.Target.Row;r<task.Target.Row+task.Target.Rows;r++)if(!plan.Original.HiddenRows.Contains(r))rows[r]=plan.Original.Height(r);}
            RestoreRows(rows);
            // Formula writes do not require changing Hidden; leave hidden rows completely untouched.
        }
        private void RestoreRows(SortedDictionary<int,double> values)
        {
            var all=values.ToList();for(int i=0;i<all.Count;)
            {
                int first=all[i].Key,last=first;double size=all[i].Value;i++;
                while(i<all.Count&&all[i].Key==last+1&&Math.Abs(all[i].Value-size)<0.001){last=all[i].Key;i++;}
                using(var bag=new ComBag())
                {dynamic range=bag.Own(((dynamic)sheet).Range[first+":"+last]);object current=range.RowHeight;if(current==null||current==DBNull.Value||Math.Abs(Convert.ToDouble(current)-size)>0.001)range.RowHeight=size;}
            }
        }
        internal IEnumerable<string> Run()
        {
            if(plan.Tasks.Count==0||plan.Tasks.Any(t=>!t.Status.StartsWith("通过",StringComparison.Ordinal)))throw new InvalidOperationException("请先让所有勾选任务通过预检。");
            CheckOpenWorkbooks(sourceApp,plan.Tasks.Select(t=>t.File),plan.Config.Parent,true);CheckParent();
            try{using(File.Open(plan.Config.Parent,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){}}
            catch(IOException error){throw new IOException("父公司文件被其他Excel进程或同步程序占用，请关闭或等待完成后重试；未覆盖原文件。",error);}
            batch="Notes_"+DateTime.Now.ToString("yyyyMMdd_HHmmss",CultureInfo.InvariantCulture)+"_"+Guid.NewGuid().ToString("N").Substring(0,8);
            string directory=Path.Combine(Path.GetDirectoryName(plan.Config.Parent),"复制备份",batch);Directory.CreateDirectory(directory);
            PreBackup=Path.Combine(directory,"执行前_"+Path.GetFileName(plan.Config.Parent));SummaryFileBatch.CopyInputSnapshot(plan.Config.Parent,PreBackup);
            if(DataMigrationHost.Digest(PreBackup)!=plan.ParentHash)throw new InvalidOperationException("执行前备份核验失败，未写入父公司。");
            stage=Path.Combine(Path.GetDirectoryName(plan.Config.Parent),".xy-notes-"+Guid.NewGuid().ToString("N")+Path.GetExtension(plan.Config.Parent));manifest=Path.Combine(directory,"session.json");
            File.Copy(PreBackup,stage,false);File.SetAttributes(stage,FileAttributes.Normal);Record("Prepared");yield return "已建立执行前备份："+PreBackup;
            Check();hidden=new HiddenWorkbookApplication(sourceApp);Open(false);
            dynamic worker=hidden.Application;worker.Calculation=-4135;worker.CalculateBeforeSave=false;Record("Writing");
            plan.Original.StandardHeight=Convert.ToDouble(((dynamic)sheet).StandardHeight);plan.Original.StandardWidth=Convert.ToDouble(((dynamic)sheet).StandardWidth);
            foreach(var task in plan.Tasks)
            {
                Check();if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在预检后变化："+task.File);
                task.Status="执行中";yield return "开始："+task.Code+task.Name+" → "+plan.TargetSheet+"!"+task.Target.Address;
                bool native=false;Step("整块粘贴链接",delegate{native=TryNativePaste(task);});
                if(native)
                {
                    if(FailureForProbe!=null&&FailureForProbe(FormulaWrites))throw new InvalidOperationException("测试注入的写入异常。");
                    Completed++;task.Status="整块链接已粘贴";yield return "整块粘贴完成："+task.Name;continue;
                }
                log("使用兼容写入："+task.Name+"（合并布局不完全一致、宿主或剪贴板不支持整块链接）");
                var refresh=System.Diagnostics.Stopwatch.StartNew();
                foreach(var part in task.Blocks.Concat(task.Anchors))foreach(var chunk in Chunks(part))
                {
                    Check();
                    using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[chunk.Address]);DataMigrationEngine.AssignFormula((object)range,Matrix(task,chunk));}
                    FormulaWrites+=chunk.Rows*chunk.Columns;
                    if(refresh.ElapsedMilliseconds>=200){yield return "粘贴链接："+task.Name+"｜累计 "+FormulaWrites+" 格";refresh.Restart();}
                    if(FailureForProbe!=null&&FailureForProbe(FormulaWrites))throw new InvalidOperationException("测试注入的写入异常。");
                }
                Completed++;task.Status="链接已粘贴";yield return "已粘贴："+task.Name+"｜累计 "+FormulaWrites+" 格";
            }
            if(plan.Cleanup!=null)
            {
                yield return "正在清理右侧多余模板："+plan.Cleanup.Address;
                Step("清理右侧多余模板",delegate{using(var bag=new ComBag()){dynamic range=bag.Own(((dynamic)sheet).Range[plan.Cleanup.Address]);range.UnMerge();range.Clear();}});
            }
            Check();yield return "正在保留原行高及隐藏状态…";
            Step("恢复行高及隐藏状态",RestoreGeometry);
            yield return "正在保存链接，不更新外链或重新计算…";
            Step("保存临时副本",SaveStage);
            Step("关闭已保存副本",Close);
            Step("检查保存文件结构",delegate{if(!DataMigrationHost.PackageSheets(stage).Contains(plan.TargetSheet))throw new IOException("保存文件缺少目标工作表。");});Record("Saved");
            foreach(var task in plan.Tasks)if(DataMigrationHost.Digest(task.File)!=task.Hash)throw new InvalidOperationException("来源文件在执行期间变化，未更新汇总文件："+task.File);
            CheckParent();yield return "链接已保存，即将更新汇总文件…";CheckParent();
            string verifiedHash;
            try{Backup=MigrationOutputTransaction.Commit(stage,plan.Config.Parent,plan.ParentHash,batch,out verifiedHash);}
            catch(Exception error){log("提交中止，请检查父公司及恢复记录："+manifest+"｜"+error.Message);throw;}
            Committed=true;CommittedHash=verifiedHash;
            try{Record("Committed");}catch(Exception error){log("父公司已更新，但提交记录写入失败；恢复备份："+Backup+"｜"+error.Message);}
            foreach(var task in plan.Tasks)task.Status="链接已粘贴";
            yield return "已更新汇总文件｜粘贴链接 "+FormulaWrites+" 格；未更新外链或检查数值；恢复备份："+Backup;
        }
        public void Dispose()
        {
            if(disposed)return;disposed=true;
            try{Close();}catch(Exception error){log("临时工作簿关闭失败："+error.Message);}
            try{references.Dispose();}catch(Exception error){log("COM释放失败："+error.Message);}
            try{if(hidden!=null)hidden.Dispose();}catch(Exception error){log("后台进程退出失败："+error.Message);}hidden=null;
            if(!Committed)
            {
                try{Record(Cancelled?"Cancelled":"Failed");}catch(Exception error){log("恢复记录写入失败："+error.Message);}

                if(PreBackup.Length>0)log("本批未交付；执行前备份保留："+PreBackup);
            }
            foreach(string path in new[]{stage,previousStage,saveCopy}.Where(x=>!string.IsNullOrEmpty(x)).Distinct(StringComparer.OrdinalIgnoreCase))
                try{if(File.Exists(path))File.Delete(path);}catch(Exception error){log("临时副本保留待清理："+path+"｜"+error.Message);}
        }
    }
}
