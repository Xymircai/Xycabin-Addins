using System;
using System.Collections.Generic;
using System.Threading;

namespace Excel_Iterate
{
    internal enum EliminationPhase { Idle, Reading, Calculating, PreviewReady, Writing, UndoAvailable, Finished }
    internal sealed class EliminationSession : IDisposable
    {
        private readonly object application;
        private readonly int thread=Thread.CurrentThread.ManagedThreadId;
        internal EliminationGateway Gateway;
        internal EliminationTargetHandle Target,DiffTarget;
        internal MergeSnapshot Snapshot,DiffSnapshot;
        internal EliminationDiffTable Diff;
        internal string BackupPath,HeadersKey="";
        internal EliminationPhase Phase=EliminationPhase.Idle;
        internal HostBatchRunner<EliminationRowsBlock,EliminationPrepared> PreviewTask;
        internal SourceChangeWatch SourceWatch;
        internal bool OwnsBusy;
        internal bool IsDisposed {get;private set;}
        internal EliminationSession(object app) {application=app;}
        internal EliminationGateway EnsureGateway()
        {AssertHost();if(IsDisposed)throw new ObjectDisposedException("EliminationSession");if(Gateway==null)Gateway=new EliminationGateway(application);return Gateway;}
        private void AssertHost(){if(thread!=Thread.CurrentThread.ManagedThreadId)throw new InvalidOperationException("抵消会话只能在原宿主线程访问。");}
        internal void CancelPreview(){AssertHost();if(PreviewTask!=null)PreviewTask.Cancel();}
        internal void StopWatch(){AssertHost();var watch=SourceWatch;SourceWatch=null;if(watch!=null)watch.Dispose();}
        internal void RestoreTargets(Action<EliminationTargetHandle,MergeSnapshot> restore)
        {
            AssertHost();
            if(Target!=null){if(Target.IsNewSheet||Snapshot!=null)restore(Target,Snapshot);Target=null;Snapshot=null;}
            if(DiffTarget!=null){if(DiffTarget.IsNewSheet||DiffSnapshot!=null)restore(DiffTarget,DiffSnapshot);DiffTarget=null;DiffSnapshot=null;}
            Phase=EliminationPhase.Idle;
        }
        public void Dispose()
        {
            AssertHost();if(IsDisposed)return;IsDisposed=true;Phase=EliminationPhase.Finished;
            try {if(PreviewTask!=null)PreviewTask.Dispose();}
            finally {PreviewTask=null;try{StopWatch();}finally{if(Gateway!=null)Gateway.Dispose();Gateway=null;Target=DiffTarget=null;Snapshot=DiffSnapshot=null;Diff=null;BackupPath=null;}}
        }
    }
    internal sealed class EliminationRowsBlock
    {
        internal readonly List<EliminationRow> Rows=new List<EliminationRow>();
        internal int Completed,Total;
    }
    internal sealed class EliminationPrepared
    {
        internal EliminationRequest Request;
        internal MergeHeaderSet Headers;
        internal MatrixOptions Options;
        internal EliminationMatrix Matrix;
        internal EliminationDiffTable Diff;
        internal List<EliminationRow> FormulaRows;
    }
    internal sealed class EliminationComputation
    {
        private readonly EliminationRequest request;
        private readonly MergeHeaderSet headers;
        private readonly MatrixOptions options;
        private readonly List<EliminationRow> rows=new List<EliminationRow>();
        private CancellationToken token;
        internal EliminationComputation(EliminationRequest input,MergeHeaderSet snapshot,MatrixOptions settings)
        {request=input;headers=snapshot;options=settings;}
        internal void Accept(EliminationRowsBlock block,CancellationToken cancellation)
        {token=cancellation;token.ThrowIfCancellationRequested();rows.AddRange(block.Rows);}
        internal EliminationPrepared Finish()
        {
            return EliminationCancellation.Run(token,delegate {
                var result=new EliminationPrepared {Request=request,Headers=headers,Options=options};
                if(request.Kind==EliminationKind.Receivable)EliminationDiffBuilder.Build(rows,request.Columns,options,out result.Diff,out result.Matrix);
                else result.Matrix=EliminationMatrixBuilder.Build(rows,request.Columns,options);
                token.ThrowIfCancellationRequested();
                if(!request.StaticValues)result.FormulaRows=rows;
                return result;
            });
        }
    }
    internal static class EliminationCancellation
    {
        [ThreadStatic] private static CancellationToken current;
        internal static void Check() {current.ThrowIfCancellationRequested();}
        internal static T Run<T>(CancellationToken token,Func<T> action)
        {var previous=current;current=token;try{token.ThrowIfCancellationRequested();return action();}finally{current=previous;}}
    }
    internal sealed class EliminationSnapshotReader : IHostBatchReader<EliminationRowsBlock>
    {
        private readonly object application,sheet;
        private readonly SourceBookSnapshot book;
        private readonly WorksheetIdentity identity;
        private readonly IList<EliminationColumn> columns;
        private readonly int header,count;
        private readonly SourceChangeWatch watch;
        private int offset;
        internal bool KeepMonitor;
        public bool CanYield {get{return watch.Available;}}
        public EliminationRowsBlock Current {get;private set;}
        internal SourceChangeWatch Monitor {get{return watch;}}
        internal EliminationSnapshotReader(object app,object workbook,object worksheet,IList<EliminationColumn> mapping,int headerRow,int lastRow)
        {
            application=app;sheet=worksheet;columns=mapping;header=headerRow;count=Math.Max(0,lastRow-headerRow);
            book=new SourceBookSnapshot(workbook);identity=WorksheetIdentity.Capture(sheet);watch=new SourceChangeWatch(app);watch.WatchBook(workbook);
        }
        public void Validate(){watch.Validate();book.Validate(application);if(!identity.Matches(sheet))throw new InvalidOperationException("来源工作表已变化，请重试。");}
        public bool MoveNext()
        {
            Validate();if(offset>=count)return false;
            int mapped=0;foreach(var column in columns)if(column.Role!=EliminationRole.Ignore)mapped++;
            if(mapped==0)return false;
            int height=Math.Min(count-offset,Math.Min(4096,Math.Max(1,65536/mapped)));
            var block=new EliminationRowsBlock {Completed=offset+height,Total=count};
            for(int r=0;r<height;r++)block.Rows.Add(new EliminationRow {SourceRow=header+1+offset+r});
            foreach(var column in columns)
            {
                if(column.Role==EliminationRole.Ignore)continue;
                using(var bag=new ComBag())
                {
                    dynamic range=bag.Own(((dynamic)sheet).Range[CellAddress.Name(header+1+offset,column.Number)+":"+CellAddress.Name(header+offset+height,column.Number)]);
                    object values=range.Value2;
                    for(int r=0;r<height;r++)
                    {
                        object value=HostValueArray.At(values,r,0,height,1);
                        block.Rows[r].Values[column.Role]=value==null?"":Convert.ToString(value,System.Globalization.CultureInfo.InvariantCulture);
                    }
                }
                watch.Validate();
            }
            offset+=height;Current=block;return true;
        }
        public void Dispose(){try{if(!KeepMonitor)watch.Dispose();}finally{identity.Dispose();}}
    }
}
