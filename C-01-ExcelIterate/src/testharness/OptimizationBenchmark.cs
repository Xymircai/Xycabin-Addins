using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Excel_Iterate;

// Offline, deterministic host-shaped fixtures. Never attaches to Office/WPS.
public static class OptimizationBenchmark
{
    private static Stopwatch cancelClock;
    private static bool cancellationRequested;
    public sealed class Collection
    {
        public readonly List<object> Items=new List<object>();
        public int Count {get{return Items.Count;}}
        public object Item(int index){return Items[index-1];}
    }
    public sealed class Dimension {public int CountValue;public int Count {get{return CountValue;}}}
    public sealed class Range
    {
        public int Row=1,Column=1,Height,Width=3;
        public Dimension Rows {get{return new Dimension {CountValue=Height};}}
        public Dimension Columns {get{return new Dimension {CountValue=Width};}}
        public object Formula {
            get {
                if(cancelClock!=null&&!cancellationRequested){cancellationRequested=true;cancelClock.Restart();}
                var array=(object[,])Array.CreateInstance(typeof(object),new[]{Height,Width},new[]{1,1});
                for(int r=1;r<=Height;r++)for(int c=1;c<=Width;c++)array[r,c]=Column+c-1==2?"=A1":Column+c-1==3?"=B1":"1";
                return array;
            }
        }
    }
    public sealed class Ranges
    {
        public object this[string address] {get{
            string[] pair=address.Split(':');int r1,c1,r2,c2;Position(pair[0],out r1,out c1);Position(pair.Length==1?pair[0]:pair[1],out r2,out c2);
            return new Range {Row=r1,Column=c1,Height=r2-r1+1,Width=c2-c1+1};
        }}
        private static void Position(string text,out int row,out int column){int n=0;column=0;while(n<text.Length&&char.IsLetter(text[n])){column=column*26+text[n]-'A'+1;n++;}row=int.Parse(text.Substring(n));}
    }
    public sealed class Sheet
    {
        public string Name="Sheet1";public object Parent;public int Height;
        public Range UsedRange {get{return new Range {Height=Height};}}
        public Ranges Range=new Ranges();
    }
    public sealed class Book {public string Name="Book.xlsx",FullName="Book.xlsx";public Collection Worksheets=new Collection();}
    public sealed class Cell {public string Address(bool row,bool column){return "A1";}}
    public sealed class App
    {
        public Book ActiveWorkbook;public Sheet ActiveSheet;public Cell ActiveCell=new Cell();public Collection Workbooks=new Collection();
        public App(int rows){ActiveWorkbook=new Book();ActiveSheet=new Sheet {Parent=ActiveWorkbook,Height=rows};ActiveWorkbook.Worksheets.Items.Add(ActiveSheet);Workbooks.Items.Add(ActiveWorkbook);}
    }
    private const BindingFlags All=BindingFlags.Public|BindingFlags.NonPublic|BindingFlags.Instance|BindingFlags.Static;
    private static string Signature(IEnumerable<string> values)
    {
        using(var hash=SHA256.Create())using(var stream=new CryptoStream(Stream.Null,hash,CryptoStreamMode.Write))
        {foreach(string value in values){byte[] bytes=Encoding.UTF8.GetBytes(value+"\n");stream.Write(bytes,0,bytes.Length);}stream.FlushFinalBlock();return BitConverter.ToString(hash.Hash).Replace("-","");}
    }
    private static List<EliminationRow> Data(int count,bool revenue)
    {
        var rows=new List<EliminationRow>();
        for(int i=0;i<count;i++)
        {
            var row=new EliminationRow {SourceRow=i+3};int pair=(i/2)%100;
            row.Values[EliminationRole.Entity]=(i%2==0?"A":"B")+pair;
            row.Values[EliminationRole.Counterparty]=(i%2==0?"B":"A")+pair;
            row.Values[EliminationRole.Condition]=i%10==9?"否":"是";
            row.Values[EliminationRole.Deadline]="20261231";row.Values[EliminationRole.Year]="2026";
            row.Values[EliminationRole.Account]=i%2==0?"应收账款":"应付账款";
            row.Values[EliminationRole.Amount]=i%2==0?"100":"80";
            row.Values[EliminationRole.Revenue]="100";row.Values[EliminationRole.Cost]="60";row.Values[EliminationRole.UnsoldRatio]="50%";
            rows.Add(row);
        }
        return rows;
    }
    private static List<EliminationColumn> Columns(bool revenue)
    {
        var columns=new List<EliminationColumn>();
        var roles=revenue?new[]{EliminationRole.Entity,EliminationRole.Counterparty,EliminationRole.Year,EliminationRole.Revenue,EliminationRole.Cost,EliminationRole.UnsoldRatio,EliminationRole.Condition}:
            new[]{EliminationRole.Entity,EliminationRole.Counterparty,EliminationRole.Deadline,EliminationRole.Account,EliminationRole.Amount,EliminationRole.Condition};
        for(int i=0;i<roles.Length;i++)columns.Add(new EliminationColumn {Number=i+1,Role=roles[i]});return columns;
    }
    private static void Isolate()
    {
        var assembly=typeof(Connect).Assembly;string scratch=Path.Combine(Path.GetTempPath(),"optimization-benchmark-"+Guid.NewGuid().ToString("N"));
        assembly.GetType("Excel_Iterate.DialogSizeStore").GetField("RootOverride",All).SetValue(null,scratch);
        assembly.GetType("Excel_Iterate.FormatSession").GetField("SourceOverride",All).SetValue(null,Path.Combine(scratch,"missing.xml"));
        assembly.GetType("Excel_Iterate.UiDiagnostics").GetField("RootOverride",All).SetValue(null,scratch);
    }
    private static Func<string> ArchitectureScenario(string mode, int count)
    {
        var assembly = typeof(Connect).Assembly;
        if (mode == "addresses")
        {
            Type type = assembly.GetType("Excel_Iterate.ConversionRangeAddress");
            MethodInfo normalize = type.GetMethod("Normalize", All), first = type.GetMethod("FirstCell", All);
            string address = "A1:A" + count;
            return delegate { return normalize.Invoke(null, new object[] { address }) + "|" + first.Invoke(null, new object[] { address }); };
        }
        if (mode == "names")
        {
            var targets = new List<string>();
            for (int i = 0; i < 3000; i++) targets.Add(@"C:\fixtures\" + i.ToString("D4") + "-公司" + i + ".xlsx");
            Type indexType = assembly.GetType("Excel_Iterate.MigrationNameIndex");
            Type rowType = assembly.GetType("Excel_Iterate.MigrationTriggerRow");
            MethodInfo match = assembly.GetType("Excel_Iterate.DataMigrationForm").GetMethod("MatchRow", All);
            return delegate
            {
                object index = indexType == null ? null : Activator.CreateInstance(indexType, All, null, new object[] { targets }, null);
                var values = new List<string>();
                for (int i = 0; i < count; i++)
                {
                    string name = "公司" + (i % 3000);
                    if (index != null)
                    {
                        var found = (IEnumerable<string>)indexType.GetMethod("Find", All).Invoke(index, new object[] { name });
                        values.Add(string.Join("|", new List<string>(found).ToArray()));
                    }
                    else
                    {
                        object row = Activator.CreateInstance(rowType, true); rowType.GetField("Name", All).SetValue(row, name);
                        match.Invoke(null, new object[] { row, targets });
                        values.Add((string)rowType.GetField("Path", All).GetValue(row));
                    }
                }
                return Signature(values);
            };
        }
        Type planType = assembly.GetType("Excel_Iterate.MigrationPlan"), sourceType = assembly.GetType("Excel_Iterate.MigrationSourceItem"), mappingType = assembly.GetType("Excel_Iterate.MigrationMapping");
        object source = Activator.CreateInstance(sourceType, true);
        var mappings = (System.Collections.IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(mappingType));
        for (int i = 0; i < count; i++)
        {
            object mapping = Activator.CreateInstance(mappingType, true);
            mappingType.GetField("Source", All).SetValue(mapping, source);
            mappingType.GetField("TargetPath", All).SetValue(mapping, "target" + (i % 100)); mappings.Add(mapping);
        }
        Type planner = assembly.GetType("Excel_Iterate.MigrationPlanner") ?? assembly.GetType("Excel_Iterate.DataMigrationForm");
        MethodInfo group = planner.GetMethod("Group", All);
        return delegate
        {
            object plan = Activator.CreateInstance(planType, true); group.Invoke(null, new object[] { plan, mappings, true });
            var values = new List<string>();
            foreach (object pass in (System.Collections.IEnumerable)planType.GetField("Passes", All).GetValue(plan))
            {
                foreach (object mapping in (System.Collections.IEnumerable)pass.GetType().GetField("Mappings", All).GetValue(pass))
                    values.Add((string)mappingType.GetField("TargetPath", All).GetValue(mapping));
                values.Add("--pass--");
            }
            return Signature(values);
        };
    }
    [STAThread] public static int Main(string[] args)
    {
        Console.OutputEncoding=Encoding.UTF8;Application.EnableVisualStyles();Isolate();
        string mode=args[0];int count=int.Parse(args[1]);bool revenue=mode=="matrix";
        var rows=mode=="diff"||mode=="matrix"?Data(count,revenue):null;var columns=Columns(revenue);
        var options=new MatrixOptions {StaticValues=true,HasConditionColumn=true,SourceFirstRow=3,SourceLastRow=count+2};
        var app=new App(count);string signature="";double cancellationMs=0;Action operation;
        if(mode=="names"||mode=="addresses"||mode=="group") { var scenario=ArchitectureScenario(mode,count); operation=delegate {signature=scenario();}; }
        else if(mode=="startup")operation=delegate{new Connect();};
        else if(mode=="cancel")operation=delegate{
            cancellationRequested=false;cancelClock=new Stopwatch();
            var found=DependentTrackerEngine.TraceDependents(app,null,delegate{return cancellationRequested;});
            cancellationMs=cancelClock.Elapsed.TotalMilliseconds;cancelClock=null;
            if(found.Status==TrackerStatus.Error)throw new Exception(found.ErrorMessage);
            signature="cancelled="+found.Cancelled+":scanned="+found.ScannedCells;
        };
        else if(mode=="scan")operation=delegate{var found=DependentTrackerEngine.TraceDependents(app,null,null);if(found.Status!=TrackerStatus.Completed||found.Items.Count!=count)throw new Exception("scan mismatch: "+found.ErrorMessage);signature=found.Items.Count+":"+found.ScannedCells+":"+found.Items[0].CellAddressA1+":"+found.Items[count-1].CellAddressA1;};
        else if(mode=="preview")
        {
            var matrix=new EliminationMatrix();matrix.Years.Add("2026");matrix.Columns.AddRange(EliminationMatrixBuilder.BuildColumns(matrix.Years,1));
            for(int i=0;i<count;i++)matrix.Rows.Add(new MatrixRow {Entity="A"+(i%100),Customer="B"+(i%100),Account="应收账款",Note="全量预览行 "+i,Adjustment="往来抵消"});
            Type formType=typeof(Connect).Assembly.GetType("Excel_Iterate.EliminationForm");MethodInfo preview=formType.GetMethod("SetPreviewMatrix",All);
            MethodInfo fit=typeof(Connect).Assembly.GetType("Excel_Iterate.Ui").GetMethod("AutoFitColumns",All);
            operation=delegate {using(var form=(Form)Activator.CreateInstance(formType,All,null,new object[]{EliminationKind.Receivable},null)){preview.Invoke(form,new object[]{matrix});var grid=(DataGridView)formType.GetField("Entries",All).GetValue(form);fit.Invoke(null,new object[]{grid});signature=grid.Rows.Count+":"+grid.Columns.Count+":"+grid.Rows[count-1].Cells[2].Value;}};
        }
        else operation=delegate {
            EliminationMatrix matrix;EliminationDiffTable diff=null;
            if(revenue)matrix=EliminationMatrixBuilder.Build(rows,columns,options);else EliminationDiffBuilder.Build(rows,columns,options,out diff,out matrix);
            if(matrix.Errors.Count!=0)throw new Exception(string.Join(";",matrix.Errors));
            var values=new List<string>();values.Add(matrix.Rows.Count+":"+matrix.GroupCount+":"+matrix.FilteredRows+":"+matrix.Balanced);
            if(diff!=null)foreach(var item in diff.Rows)values.Add(item.Creditor+"|"+item.Debtor+"|"+item.Asset.Value+"|"+item.Liability.Value+"|"+item.Offset.Value+"|"+item.Difference.Value+"|"+item.CreditorAsset.Value+"|"+item.CreditorLiability.Value+"|"+item.DebtorAsset.Value+"|"+item.DebtorLiability.Value);
            foreach(var item in matrix.Rows){values.Add(item.Entity+"|"+item.Customer+"|"+item.Account);foreach(var amount in item.Debit)values.Add("D"+amount.Key+"="+amount.Value.Value);foreach(var amount in item.Credit)values.Add("C"+amount.Key+"="+amount.Value.Value);}
            signature=Signature(values);
        };
        Console.WriteLine("scenario,rows,iteration,elapsed_ms,peak_working_set,private_bytes,signature,cancel_response_ms");
        if(mode!="startup")operation();
        for(int iteration=0;iteration<(mode=="startup"?1:5);iteration++)
        {
            GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();var clock=Stopwatch.StartNew();operation();clock.Stop();
            using(var process=Process.GetCurrentProcess())Console.WriteLine(mode+","+count+","+iteration+","+clock.Elapsed.TotalMilliseconds.ToString("F3",CultureInfo.InvariantCulture)+","+process.PeakWorkingSet64+","+process.PrivateMemorySize64+","+signature+","+(mode=="cancel"?cancellationMs.ToString("F3",CultureInfo.InvariantCulture):""));
        }
        return 0;
    }
}
