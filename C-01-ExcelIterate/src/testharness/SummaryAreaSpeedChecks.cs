using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Excel_Iterate;
internal static class SummaryAreaSpeedChecks
{
    static readonly BindingFlags F=BindingFlags.Instance|BindingFlags.NonPublic;
    static void Set(object x,string n,object v){x.GetType().GetField(n,F).SetValue(x,v);}
    static object Get(object x,string n){return x.GetType().GetField(n,F).GetValue(x);}
    static void Call(object x,string n,params object[] a){try{x.GetType().GetMethod(n,F).Invoke(x,a);}catch(TargetInvocationException e){throw e.InnerException;}}
    internal static int Run()
    {
        using(var session=new ProbeSession("SummarySpeed"))
        try
        {
            dynamic app=session.Take(Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application",true)));
            if(!session.Override(()=>app.DisplayAlerts,v=>app.DisplayAlerts=v,false,"DisplayAlerts")||!session.Override(()=>app.EnableEvents,v=>app.EnableEvents=v,false,"EnableEvents")||!session.Override(()=>app.ScreenUpdating,v=>app.ScreenUpdating=v,false,"ScreenUpdating"))throw new Exception("Cannot establish host state");
            using(var bag=new ComBag())
            {
                dynamic books=bag.Own(app.Workbooks);
                foreach(string name in new[]{"样本甲","样本乙"})
                {
                    dynamic source=books.Add();bool closed=false;session.Own((object)source,delegate{if(!closed)source.Close(false);});
                    using(var one=new ComBag()){dynamic sheets=one.Own(source.Worksheets),sheet=one.Own(sheets.Item(1));sheet.Name="BS";dynamic cell=one.Own(sheet.Range["A1"]);cell.Value2=25;}
                    source.SaveAs(session.Output(name+".xlsx"),51);source.Close(false);closed=true;
                }
                dynamic book=books.Add();session.Own((object)book,delegate{book.Close(false);});dynamic sheets2=bag.Own(book.Worksheets),target=bag.Own(sheets2.Item(1));target.Name="汇总";
                dynamic lead=bag.Own(sheets2.Add());lead.Name="Lead";dynamic fields=bag.Own(lead.Range["A1:B2"]);fields.Value2=new object[,]{{"01","样本甲"},{"02","样本乙"}};
                dynamic a=bag.Own(target.Range["L3:P22"]),b=bag.Own(target.Range["L25:P44"]),old=bag.Own(target.Range["Q3:Z44"]);a.Value2=b.Value2="BS!$A$1";old.Value2="old";
                using(var form=new SingleSummaryForm(true))
                {
                    var view=(SummaryAreaView)form.Controls[0];view.Folder.Text=session.DirectoryPath;
                    using(var host=new SummaryAreaHost((object)app,form,view,null,delegate{return true;}))
                    {
                        Set(host,"book",(object)book);Set(host,"bookName",Convert.ToString(book.FullName));Set(host,"templateSheet","汇总");Set(host,"pasteSheet","汇总");Set(host,"leadSheet","Lead");
                        var areas=(List<SummaryAreaPlan>)Get(host,"areas");areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=3,Column=12,Rows=20,Columns=5,Address="L3:P22"});areas.Add(new SummaryAreaPlan{Sheet="汇总",Row=25,Column=12,Rows=20,Columns=5,Address="L25:P44"});
                        Set(host,"paste",new SummaryAreaPlan{Row=3,Column=17,Rows=1,Columns=1});Set(host,"lead",new SummaryAreaPlan{Sheet="Lead",Row=1,Column=1,Rows=2,Columns=2,Address="A1:B2"});
                        int calculation=Convert.ToInt32(app.Calculation);var clock=Stopwatch.StartNew();Call(host,"Preview");long preview=clock.ElapsedMilliseconds;clock.Restart();Call(host,"Start");for(int i=0;i<4;i++)Call(host,"Tick",null,EventArgs.Empty);long execute=clock.ElapsedMilliseconds;
                        if((bool)Get(host,"failed")||!(bool)Get(host,"pending"))throw new Exception("Execution did not reach pending state");
                        dynamic first=bag.Own(target.Range["Q3"]),last=bag.Own(target.Range["Z44"]),gap=bag.Own(target.Range["Q23"]);
                        if(!Convert.ToBoolean(first.HasFormula)||!Convert.ToBoolean(last.HasFormula)||Convert.ToString(gap.Value2)!="old"||Convert.ToInt32(app.Calculation)!=calculation)throw new Exception("Formula/gap/calculation guard failed");
                        clock.Restart();Call(host,"Undo");long undo=clock.ElapsedMilliseconds;if(Convert.ToString(first.Value2)!="old"||Convert.ToString(last.Value2)!="old")throw new Exception("Undo failed");
                        var field=host.GetType().GetField("bulkWrites",F);Console.WriteLine("TIMING host="+Convert.ToString(app.Name)+" version="+Convert.ToString(app.Version)+" template_cells=200 companies=2 formula_cells=400 preview_ms="+preview+" execute_ms="+execute+" undo_ms="+undo+" bulk_formula_writes="+(field==null?-1:Convert.ToInt32(field.GetValue(host))));
                    }
                }
            }
            return session.Finish(0);
        }
        catch(Exception e){Console.WriteLine(e);return session.Finish(1);}
    }
}
