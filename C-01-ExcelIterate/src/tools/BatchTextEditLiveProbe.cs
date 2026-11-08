using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Linq;
using System.Runtime.InteropServices;
using Excel_Iterate;

// LiveProbeSafety owns every created workbook and all COM references. User documents
// are never edited, saved or closed; no add-in installation or Application.Quit.
internal static class BatchTextEditLiveProbe
{
    private static int checks;
    private static void Check(bool value, string message) { if (!value) throw new Exception(message); checks++; Console.WriteLine("PASS " + message); }
    private static ConversionRequest Edit(BatchTextEditOperation operation, BatchTextCountOrigin origin, int count, int first = 1, int last = 1, string text = "X")
    { return new ConversionRequest { Kind = CellConversionKind.BatchInsertTrimDelete, TextEdit = new BatchTextEditOptions { Operation = operation, Origin = origin, Count = count, Start = first, End = last, Text = text } }; }
    private static void Formula(object range, string value)
    { ((dynamic)range).NumberFormat = "General"; try { ((dynamic)range).Formula2 = value; } catch (Exception error) { if (error is MissingMemberException || error is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) ((dynamic)range).Formula = value; else { var com = error as COMException; if (com == null || com.ErrorCode != unchecked((int)0x80020003) && com.ErrorCode != unchecked((int)0x80020006)) throw; ((dynamic)range).Formula = value; } } }
    private static CellConversionResult Run(object app, string address, ConversionRequest request, string recovery, Func<BatchSyntaxReport, bool> confirm = null)
    { var host = new ComConversionHost(app); host.SetAddress(address); return CellConversionEngine.Run(host, request, recovery, confirm); }
    private static void Legacy(object app, string address, string assemblyPath, string recovery)
    {
        var assembly = Assembly.LoadFile(Path.GetFullPath(assemblyPath)); var flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
        object options = Activator.CreateInstance(assembly.GetType("Excel_Iterate.BatchTextEditOptions"), true);
        options.GetType().GetField("Operation", flags).SetValue(options, Enum.Parse(assembly.GetType("Excel_Iterate.BatchTextEditOperation"), "Delete"));
        options.GetType().GetField("Origin", flags).SetValue(options, Enum.Parse(assembly.GetType("Excel_Iterate.BatchTextCountOrigin"), "Start"));
        options.GetType().GetField("Count", flags).SetValue(options, 1);
        object request = Activator.CreateInstance(assembly.GetType("Excel_Iterate.ConversionRequest"), true);
        request.GetType().GetField("Kind", flags).SetValue(request, Enum.Parse(assembly.GetType("Excel_Iterate.CellConversionKind"), "BatchInsertTrimDelete"));
        request.GetType().GetField("TextEdit", flags).SetValue(request, options);
        object host = Activator.CreateInstance(assembly.GetType("Excel_Iterate.ComConversionHost"), flags, null, new[] { app }, null);
        host.GetType().GetMethod("SetAddress").Invoke(host, new object[] { address });
        var clock = Stopwatch.StartNew(); object result = assembly.GetType("Excel_Iterate.CellConversionEngine").GetMethod("Run", flags).Invoke(null, new[] { host, request, recovery });
        Check(!(bool)result.GetType().GetField("Failed", flags).GetValue(result), "legacy baseline completes in owned sample " + address);
        Console.WriteLine("LIVE_BENCHMARK legacy count=" + address.Substring(4) + " total=" + clock.ElapsedMilliseconds + "ms; calculation=manual (same for baseline and fixed)");
    }
    private static void CheckHostBinding(ProbeSession session, object application, object book, object sheet, object global)
    {
        object previous = Connect.HostApplication;
        try
        {
            Connect.HostApplication = global;
            Connect.BindHostApplication(delegate { return application; });
            Check(Connect.EnsureApplication() && WorksheetIdentity.SameObject(Connect.HostApplication, application), "shared command connection uses the current host instead of the global instance");
            IntPtr unknown = Marshal.GetIUnknownForObject(application), requested = IntPtr.Zero;
            try
            {
                Guid iid = typeof(Connect).Assembly.GetType("Excel_Iterate.ThisAddIn+HostExcelApplication").GUID;
                Check(Marshal.QueryInterface(unknown, ref iid, out requested) == 0, "WPS supports the application identity requested from the VSTO provider");
            }
            finally { if (requested != IntPtr.Zero) Marshal.Release(requested); Marshal.Release(unknown); }

            using (var host = new ComConversionHost(Connect.HostApplication)) Check(host.ReadPreview("A1").CellAddress == "A1", "conversion command obtains its current worksheet");
            using (var names = new FileNameWorksheet(Connect.HostApplication)) Check(names.Description == Convert.ToString(((dynamic)sheet).Name), "file-name extraction obtains the current worksheet");
            string path = session.Output("current-host.xlsx"); ((dynamic)book).SaveAs(path, 51);
            var links = new LinkAssistantSession(Connect.HostApplication);
            Check(string.Equals(links.ReportBookPath, path, StringComparison.OrdinalIgnoreCase), "link assistant uses the owned saved workbook in the current host");
            using (var refs = session.References())
            {
                object actual = refs.Take(((dynamic)Connect.HostApplication).ActiveWorkbook);
                Check(WorksheetIdentity.SameObject(actual, book), "all shared entry checks preserve the owned workbook identity");
            }
        }
        finally { Connect.BindHostApplication(null); Connect.HostApplication = previous; }
    }
    [STAThread] private static int Main(string[] args)
    {
        using (var session = new ProbeSession("batch-text-edit-live"))
        {
            try
            {
                bool standalone = args.Length > 0 && args[0] == "--standalone";
                if (standalone) args = args.Skip(1).ToArray();
                bool fresh = args.Length > 0 && args[0] == "--new-instance";
                if (fresh) args = args.Skip(1).ToArray();
                bool binding = args.Length > 0 && args[0] == "--host-binding";
                if (binding) args = args.Skip(1).ToArray();
                string[] ids = standalone ? new[] { "Excel.Application" } : new[] { "KET.Application", "et.Application" };
                object global = binding ? session.Take(ProbeSession.Connect(ids, true)) : null;
                object application = session.Take(fresh ? ProbeSession.Connect(ids, false, id => null, id => { Type type = Type.GetTypeFromProgID(id); return type == null ? null : Activator.CreateInstance(type); }, Console.WriteLine) : ProbeSession.Connect(ids, false));
                if (application == null) { Console.WriteLine("SKIP no WPS COM host available"); return session.Finish(2); }
                dynamic app = application; object oldSheet = null, oldSelection = null;
                if (args.Length == 1 && args[0] == "--state")
                {
                    Action<string, Func<object>> inspect = delegate(string name, Func<object> get) { try { Console.WriteLine(name + "=" + Convert.ToString(get())); } catch (Exception error) { Console.WriteLine(name + " ERROR " + error.Message); } };
                    inspect("Name", () => (object)app.Name); inspect("Version", () => (object)app.Version); inspect("Visible", () => (object)app.Visible);
                    inspect("Workbooks", () => (object)app.Workbooks.Count); inspect("Calculation", () => (object)app.Calculation); inspect("DisplayAlerts", () => (object)app.DisplayAlerts);
                    inspect("Ready", () => (object)app.Ready); inspect("Interactive", () => (object)app.Interactive); inspect("Hwnd", () => (object)app.Hwnd);
                    return session.Finish(0);
                }
                try { oldSheet = session.Take((object)app.ActiveSheet); oldSelection = session.Take((object)app.Selection); } catch (COMException) { }
                object savedSheet = oldSheet, savedSelection = oldSelection;
                session.OnClose(delegate { if (savedSheet != null) ((dynamic)savedSheet).Activate(); if (savedSelection != null) ((dynamic)savedSelection).Select(); });
                session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = Convert.ToBoolean(value), false, "EnableEvents");
                session.Override(() => (object)app.DisplayAlerts, value => app.DisplayAlerts = Convert.ToBoolean(value), false, "DisplayAlerts");
                session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = Convert.ToBoolean(value), false, "ScreenUpdating");
                session.Override(() => (object)app.Visible, value => app.Visible = Convert.ToBoolean(value), (object)app.Visible, "Visible (unchanged)");
                session.Override(() => (object)app.Calculation, value => app.Calculation = Convert.ToInt32(value), -4135, "Calculation");
                using (var refs = session.References())
                {
                    object books = refs.Take((object)app.Workbooks); object book = null;
                    book = session.Own((object)((dynamic)books).Add(), delegate { ((dynamic)book).Close(false); });
                    object sheets = refs.Take((object)((dynamic)book).Worksheets);
                    object data = refs.Take((object)((dynamic)sheets).Item(1)); ((dynamic)data).Name = "ProbeData";
                    object bs = refs.Take((object)((dynamic)sheets).Add()); ((dynamic)bs).Name = "BS";
                    object bsValue = refs.Take((object)((dynamic)bs).Range["L16"]); ((dynamic)bsValue).Value2 = 12.0;
                    ((dynamic)data).Activate();
                    object first = refs.Take((object)((dynamic)data).Range["A1"]);
                    if (binding) CheckHostBinding(session, application, book, data, global);
                    foreach (int count in new[] { 300, 3000 })
                    {
                        string address = "A1:A" + count; object range = refs.Take((object)((dynamic)data).Range[address]);
                        ((dynamic)range).NumberFormat = "General"; Formula(range, "=BS!$L$16");
                        if (args.Length == 1) { Legacy(application, address, args[0], session.DirectoryPath); ((dynamic)range).NumberFormat = "General"; Formula(range, "=BS!$L$16"); }
                        var result = Run(application, address, Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 2), session.DirectoryPath);
                        Check(!result.Failed && result.Changed == count && result.PlainResults == count && result.FormulasAsText == 0, "fixed ordinary batch " + count + ": " + result.Details);
                        Check(!(bool)((dynamic)first).HasFormula && (string)((dynamic)first).Value2 == "S!$L$16" && (string)((dynamic)first).NumberFormat == "General", "no quote prefix or forced text format " + count);
                        Console.WriteLine("LIVE_BENCHMARK fixed count=" + count + " " + result.Timings);
                    }
                    foreach (BatchTextEditOperation operation in new[] { BatchTextEditOperation.Delete, BatchTextEditOperation.Insert, BatchTextEditOperation.Trim })
                    {
                        Formula(first, "=BS!$L$16"); ((dynamic)first).NumberFormat = "0.00";
                        var request = operation == BatchTextEditOperation.Delete ? Edit(operation, BatchTextCountOrigin.Middle, 0, 2, 3)
                            : operation == BatchTextEditOperation.Trim ? Edit(operation, BatchTextCountOrigin.Start, 4) : Edit(operation, BatchTextCountOrigin.Start, 1, 1, 1, "!");
                        int prompted = 0; string expected = BatchTextEdit.Transform("=BS!$L$16", request.TextEdit);
                        var result = Run(application, "A1", request, session.DirectoryPath, report => { prompted++; Console.WriteLine("SYNTAX_REPORT " + report.Message); return false; });
                        Check(result.Cancelled && !result.Failed && prompted == 1 && (string)((dynamic)first).Formula == "=BS!$L$16" && (string)((dynamic)first).NumberFormat == "0.00", "syntax cancellation leaves source intact " + operation + ": " + result.Details);
                        result = Run(application, "A1", request, session.DirectoryPath, report => true);
                        Check(!result.Failed && result.FormulasAsText == 1 && !(bool)((dynamic)first).HasFormula && (string)((dynamic)first).Value2 == expected && (string)((dynamic)first).NumberFormat == "@", "confirmed syntax rejection stores exact literal " + operation + ": " + result.Details);
                    }
                    Formula(first, "=1/0");
                    var errorResult = Run(application, "A1", Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, 1, 1, "+0"), session.DirectoryPath, report => { throw new Exception("error-value formula must not warn"); });
                    Check(!errorResult.Failed && (bool)((dynamic)first).HasFormula && (string)((dynamic)first).Formula == "=1/0+0", "calculation error values remain formulas");
                    ((dynamic)first).NumberFormat = "@"; ((dynamic)first).Value2 = "00123"; ((dynamic)first).NumberFormat = "General";
                    var zeros = Run(application, "A1", Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, 1, 1, "X"), session.DirectoryPath);
                    Check(!zeros.Failed && (string)((dynamic)first).Value2 == "00123X" && (string)((dynamic)first).NumberFormat == "General", "raw leading-zero text is retained with original format");
                    ((dynamic)first).NumberFormat = "General"; ((dynamic)first).Formula = "''abc";
                    Check((string)((dynamic)first).Value2 == "'abc", "fixture contains an intentional literal apostrophe");
                    var quote = Run(application, "A1", Edit(BatchTextEditOperation.Insert, BatchTextCountOrigin.End, 0, 1, 1, "X"), session.DirectoryPath);
                    Check(!quote.Failed && (string)((dynamic)first).Value2 == "'abcX" && (string)((dynamic)first).NumberFormat == "General", "existing literal apostrophe is preserved without an added prefix: " + quote.Details + "; value=" + Convert.ToString(((dynamic)first).Value2));
                    app.Calculation = -4105; Formula(first, "=BS!$L$16");
                    var automatic = Run(application, "A1", Edit(BatchTextEditOperation.Delete, BatchTextCountOrigin.Start, 1), session.DirectoryPath);
                    Check(!automatic.Failed && Convert.ToInt32(app.Calculation) == -4105 && (string)((dynamic)first).Value2 == "BS!$L$16", "original automatic calculation mode is restored");
                }
                Console.WriteLine("PASS BatchTextEditLiveProbe " + checks + " assertions; owned samples only"); return session.Finish(0);
            }
            catch (Exception error) { Console.WriteLine(error); return session.Finish(1); }
        }
    }
}
