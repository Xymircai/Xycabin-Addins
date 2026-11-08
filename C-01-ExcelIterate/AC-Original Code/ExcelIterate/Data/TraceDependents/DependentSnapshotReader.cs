using System;
using System.Collections.Generic;
using System.Threading;

namespace Excel_Iterate
{
    internal static class HostValueArray
    {
        internal static object At(object value,int row,int column,int rows,int columns)
        {
            var array=value as Array;
            if(array==null)
            {
                if(value==null)return null;
                if(rows==1&&columns==1)return value;
                throw new InvalidOperationException("宿主返回的标量与读取范围形状不符。");
            }
            if(array.Rank==2&&array.GetLength(0)==rows&&array.GetLength(1)==columns)
                return array.GetValue(row+array.GetLowerBound(0),column+array.GetLowerBound(1));
            if(array.Rank==1&&((columns==1&&array.Length==rows)||(rows==1&&array.Length==columns)))
                return array.GetValue(array.GetLowerBound(0)+(columns==1?row:column));
            throw new InvalidOperationException("宿主返回的数组维度或边界与读取范围不符。");
        }
    }
    internal sealed class SourceBookSnapshot
    {
        internal readonly object Book;
        internal readonly string Name,Path;
        internal SourceBookSnapshot(object book)
        {Book=book;Name=Convert.ToString(((dynamic)book).Name);Path=Convert.ToString(((dynamic)book).FullName);}
        internal void Validate(object application)
        {
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(((dynamic)application).Workbooks);
                int count=Convert.ToInt32(books.Count);
                for(int i=1;i<=count;i++)using(var item=new ComBag())
                {
                    object candidate=item.Own(MergeWorkbookGateway.GetItem(books,i));
                    if(!WorksheetIdentity.SameObject(Book,candidate))continue;
                    if(Name==Convert.ToString(((dynamic)candidate).Name)&&Path==Convert.ToString(((dynamic)candidate).FullName))return;
                    break;
                }
            }
            throw new InvalidOperationException("来源工作簿已关闭、另存或替换，请重试。");
        }
    }
    internal sealed class FormulaBlock
    {
        internal string WorkbookName,WorkbookPath,SheetName;
        internal int FirstRow,FirstColumn,Rows,Columns,WorkbookIndex,WorkbookCount,SheetIndex,SheetCount;
        internal string[] Values;
        internal static FormulaBlock Capture(object values,int rows,int columns)
        {
            var block=new FormulaBlock {Rows=rows,Columns=columns,Values=new string[checked(rows*columns)]};
            for(int r=0;r<rows;r++)for(int c=0;c<columns;c++)block.Values[r*columns+c]=HostValueArray.At(values,r,c,rows,columns) as string;
            return block;
        }
    }
    internal sealed class DependentAccumulator
    {
        internal readonly DependentResult Result;
        internal DependentAccumulator(DependentResult result) {Result=result;}
        internal void Accept(FormulaBlock block,CancellationToken token)
        {
            for(int r=0;r<block.Rows;r++)
            {
                token.ThrowIfCancellationRequested();
                for(int c=0;c<block.Columns;c++)
                {
                    checked {Result.ScannedCells++;}
                    string formula=block.Values[r*block.Columns+c];
                    if(string.IsNullOrEmpty(formula)||!formula.StartsWith("=",StringComparison.Ordinal))continue;
                    Result.ScannedFormulas++;
                    DependentTrackerEngine.MatchFormula(formula,block.WorkbookName,block.WorkbookPath,block.SheetName,
                        CellAddress.ColumnLetters(block.FirstColumn+c)+(block.FirstRow+r).ToString(),
                        Result.TargetWorkbookName,Result.TargetSheetName,Result.TargetCellAddress,Result);
                }
            }
        }
        internal DependentResult Finish() {Result.Status=TrackerStatus.Completed;return Result;}
    }
    internal sealed class DependentSnapshotReader : IHostBatchReader<FormulaBlock>
    {
        private sealed class Sheet
        {
            internal object Value;internal SourceBookSnapshot Book;internal string Name;
            internal int BookIndex,BookCount,SheetIndex,SheetCount;
        }
        private readonly object application;
        private readonly int thread=Thread.CurrentThread.ManagedThreadId;
        private readonly ComBag owned=new ComBag();
        private readonly List<Sheet> sheets=new List<Sheet>();
        private SourceChangeWatch watch;
        private SourceBookSnapshot target;
        private int sheetIndex=-1,rowOffset,firstRow,firstColumn,rows,columns;
        private bool disposed;
        internal bool Valid {get;private set;}
        public bool CanYield {get{return watch!=null&&watch.Available;}}
        public FormulaBlock Current {get;private set;}
        internal DependentSnapshotReader(object app,DependentResult result)
        {
            application=app;
            try
            {
                if(app==null){result.Status=TrackerStatus.Error;result.ErrorMessage="宿主为空";return;}
                object book=owned.Own(((dynamic)app).ActiveWorkbook);
                if(book==null){result.Status=TrackerStatus.NoWorkbook;return;}
                object cell=owned.Own(((dynamic)app).ActiveCell);
                if(cell==null){result.Status=TrackerStatus.NoActiveCell;return;}
                target=new SourceBookSnapshot(book);result.TargetWorkbookName=target.Name;
                dynamic active=owned.Own(((dynamic)app).ActiveSheet);result.TargetSheetName=Convert.ToString(active.Name);
                result.TargetCellAddress=DependentTrackerEngine.NormalizeAddr(Convert.ToString(((dynamic)cell).Address(false,false)));
                watch=new SourceChangeWatch(app);
                dynamic books=owned.Own(((dynamic)app).Workbooks);int count=Convert.ToInt32(books.Count);
                for(int b=1;b<=count;b++)
                {
                    object current=owned.Own(MergeWorkbookGateway.GetItem(books,b));var identity=new SourceBookSnapshot(current);watch.WatchBook(current);
                    dynamic collection=owned.Own(((dynamic)current).Worksheets);int length=Convert.ToInt32(collection.Count);
                    for(int s=1;s<=length;s++)
                    {
                        object sheet=owned.Own(MergeWorkbookGateway.GetItem(collection,s));
                        sheets.Add(new Sheet {Value=sheet,Book=identity,Name=Convert.ToString(((dynamic)sheet).Name),BookIndex=b,BookCount=count,SheetIndex=s,SheetCount=length});
                    }
                }
                Valid=true;
            }
            catch {Dispose();throw;}
        }
        public void Validate()
        {
            if(disposed)throw new ObjectDisposedException("DependentSnapshotReader");
            if(Thread.CurrentThread.ManagedThreadId!=thread)throw new InvalidOperationException("扫描读取只能在原宿主线程执行。");
            if(!Valid)return;watch.Validate();target.Validate(application);
            if(sheetIndex>=0&&sheetIndex<sheets.Count)
            {
                var sheet=sheets[sheetIndex];sheet.Book.Validate(application);
                if(Convert.ToString(((dynamic)sheet.Value).Name)!=sheet.Name)throw new InvalidOperationException("来源工作表已重命名或关闭，请重试。");
            }
        }
        public bool MoveNext()
        {
            Validate();if(!Valid)return false;
            if(sheetIndex<0||rowOffset>=rows)
            {
                sheetIndex++;if(sheetIndex>=sheets.Count)return false;
                using(var bag=new ComBag())
                {
                    dynamic used=bag.Own(((dynamic)sheets[sheetIndex].Value).UsedRange);
                    firstRow=Convert.ToInt32(used.Row);firstColumn=Convert.ToInt32(used.Column);
                    dynamic rowRange=bag.Own(used.Rows),columnRange=bag.Own(used.Columns);
                    rows=Convert.ToInt32(rowRange.Count);columns=Convert.ToInt32(columnRange.Count);
                    if(rows<1||columns<1||columns>65536)throw new InvalidOperationException("来源范围大小无效。");
                }
                rowOffset=0;
            }
            var source=sheets[sheetIndex];int height=Math.Min(rows-rowOffset,Math.Min(4096,65536/columns));
            using(var bag=new ComBag())
            {
                dynamic range=bag.Own(((dynamic)source.Value).Range[CellAddress.Name(firstRow+rowOffset,firstColumn)+":"+CellAddress.Name(firstRow+rowOffset+height-1,firstColumn+columns-1)]);
                Current=FormulaBlock.Capture(range.Formula,height,columns);
            }
            Current.FirstRow=firstRow+rowOffset;Current.FirstColumn=firstColumn;Current.WorkbookName=source.Book.Name;Current.WorkbookPath=source.Book.Path;Current.SheetName=source.Name;
            Current.WorkbookIndex=source.BookIndex;Current.WorkbookCount=source.BookCount;Current.SheetIndex=source.SheetIndex;Current.SheetCount=source.SheetCount;
            rowOffset+=height;return true;
        }
        public void Dispose()
        {
            if(disposed)return;disposed=true;
            try {if(watch!=null)watch.Dispose();}finally {owned.Dispose();}
        }
    }
}
