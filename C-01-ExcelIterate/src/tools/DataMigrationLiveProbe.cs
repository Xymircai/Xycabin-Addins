using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel_Iterate;

// Explicit integration test. All documents are created in ProbeSession's own
// directory. Never install the add-in, call Quit, or write an existing document.
internal static class DataMigrationLiveProbe
{
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    private sealed class ProbeWindow : IWin32Window { internal IntPtr Value; public IntPtr Handle { get { return Value; } } }
    private static void InteractiveRun(object application, DataMigrationHost host, MigrationRequest request, Action<MigrationProgress> inspect)
    {
        Exception failure = null; int observations = 0, missed = 0; bool completed = false;
        using (var form = new DataMigrationForm()) using (var timeout = new Timer { Interval = 300000 })
        {
            form.Text += " · 焦点验证";
            LoadRequest(form, request);
            Action<MigrationProgress> observe = delegate(MigrationProgress step)
            {
                inspect(step); observations++;
                if (observations <= 3)
                {
                    uint process; GetWindowThreadProcessId(GetForegroundWindow(), out process);
                    using (var info = new MigrationWindowScope(application)) Console.WriteLine("FOCUS_SAMPLE: foreground=" + GetForegroundWindow() + " foregroundProcess=" + process + " dialog=" + form.Handle + " minimized=" + IsIconic(form.Handle) + " hostProcesses=" + info.HostIdentity);
                }
                if (GetForegroundWindow() != form.Handle || IsIconic(form.Handle)) missed++;
            };
            Action<Exception> failed = delegate(Exception error) { failure = error; form.Close(); };
            form.PreviewRequested += delegate
            {
                form.Run(Observed(host.Preview(request), observe), delegate
                {
                    form.RefreshResults(); Console.WriteLine("UI_PREVIEW_READY");
                }, failed);
            };
            form.ExecuteRequested += delegate { form.Run(Observed(host.Execute(), observe), delegate { completed = true; form.RefreshResults(); form.Close(); }, failed); };
            form.CancelRequested += delegate { host.Cancelled = true; };
            timeout.Tick += delegate { failure = new TimeoutException("焦点验证未在限时内完成。"); host.Cancelled = true; form.Close(); };
            IntPtr owner = new IntPtr(Convert.ToInt64(((dynamic)application).Hwnd));
            form.Activated += delegate { Console.WriteLine("UI_ACTIVATED: foreground=" + GetForegroundWindow() + " dialog=" + form.Handle); };
            uint ownerProcess; GetWindowThreadProcessId(owner, out ownerProcess);
            form.Show(new ProbeWindow { Value = owner }); timeout.Start(); Console.WriteLine("UI_FOCUS_TEST_READY: foreground=" + GetForegroundWindow() + " dialog=" + form.Handle + " host=" + owner + " hostProcess=" + ownerProcess); Application.Run(form);
        }
        if (failure != null) throw failure;
        Check(completed && observations > 0 && missed == 0, "迁移窗口在真实宿主读写全过程保持前台且未最小化；观测=" + observations + "，失焦=" + missed);
    }
    // Show the request as one row of the table (source -> every target).
    private static void LoadRequest(DataMigrationForm form, MigrationRequest request)
    {
        form.AddBook(false, request.SourcePath, new[] { request.SourceSheet });
        foreach (var target in request.Targets) form.AddBook(true, target.Path, new[] { target.Sheet });
        var pair = form.AddPair(new MigrationPair());
        pair.Source.Books.Clear(); pair.Source.Books.Add(Path.GetFullPath(request.SourcePath)); pair.Source.Sheet = request.SourceSheet; pair.Source.Range = request.SourceRange;
        pair.Target.Books.Clear(); pair.Target.Books.AddRange(request.Targets.Select(t => Path.GetFullPath(t.Path)));
        if (request.Targets.Count > 0) { pair.Target.Sheet = request.Targets[0].Sheet; pair.Target.Range = request.Targets[0].Start; }
        form.RefreshResults();
    }
    private static void Check(bool passed, string message)
    { if (!passed) throw new InvalidOperationException(message); Console.WriteLine("PASS: " + message); }
    private static void Drain(IEnumerable<MigrationProgress> steps)
        { foreach (var step in steps) Console.WriteLine("STEP: " + step.Message); }
    private static void DrainPumped(IEnumerable<MigrationProgress> sequence)
    {
        Exception failure = null; bool done = false, ticking = false;
        using (var context = new ApplicationContext()) using (var timer = new Timer { Interval = 15 }) using (var steps = sequence.GetEnumerator())
        {
            timer.Tick += delegate { if (ticking) return; ticking = true; timer.Stop(); try { if (!steps.MoveNext()) { done = true; context.ExitThread(); } else Console.WriteLine("PUMP: " + steps.Current.Message); } catch (Exception error) { failure = error; done = true; context.ExitThread(); } finally { ticking = false; if (!done) timer.Start(); } };
            timer.Start(); Application.Run(context);
        }
        if (failure != null) throw failure;
    }
    private static void Reproduce(ProbeSession session, object application, string[] args)
    {
        dynamic app = application;
        string source = session.Output("repro-source.xlsx"); File.Copy(args[1], source);
        var targets = new List<MigrationTarget>();
        for (int i = 2; i <= 3; i++)
        {
            string path = session.Output(Path.GetFileName(args[i])); File.Copy(args[i], path);
            targets.Add(new MigrationTarget { Path = path, Sheet = "Sheet1", Start = args[0] == "--repro" ? "A1" : "B2" });
            using (var refs = session.References())
            {
                dynamic books = refs.Take(app.Workbooks); object owned = books.Open(path, 0, false);
                session.Own(owned, delegate { ((dynamic)owned).Close(false); });
            }
        }
        session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, true, "repro events");
        DataMigrationHost.DiagnosticRootOverride = Path.Combine(session.DirectoryPath, "diagnostics");
        using (var host = new DataMigrationHost(application))
        {
            if (args[0] != "--repro")
            {
                bool recalculated = false;
                var request = new MigrationRequest { SourcePath = source, SourceSheet = "迁移源", SourceRange = "B2:I22", Targets = targets };
                RunFormBatch(host, request, delegate(MigrationProgress step) {
                    if (args[0] != "--diagnose" || recalculated || !step.Message.StartsWith("写入 " + Path.GetFileName(targets[1].Path) + " ", StringComparison.Ordinal)) return;
                    recalculated = true;
                    var handle = (MergeWorkbookHandle)typeof(DataMigrationHost).GetField("source", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(host);
                    Console.WriteLine("DIAGNOSTIC: calculate source sheet without editing any values or formulas");
                    ((dynamic)handle.Sheet).Calculate();
                });
                Console.WriteLine("EVENT_TRACE:\n"+host.EventTrace);
            }
            else { DrainPumped(host.Preview(new MigrationRequest { SourcePath = source, SourceSheet = "迁移源", SourceRange = "A1", AllowEmptySource = true, Targets = targets })); DrainPumped(host.Execute()); }
            foreach (var target in targets) { Console.WriteLine("TARGET: " + target.Path + " " + target.Status + " " + target.Error + "\n" + target.FailureDetails); Check(target.Status == "成功", "two open target workbooks processed independently"); if (target.Output.Length > 0) session.TrackFile(target.Output); }
            if(args[0]=="--repair")foreach(var target in targets){string output=Path.Combine(Path.GetFullPath(args[4]),Path.GetFileNameWithoutExtension(target.Path)+"_迁移修正结果.xlsx");session.Export(target.Output,output);Console.WriteLine("REPAIRED_OUTPUT: "+output);}
        }
        foreach (string file in Directory.GetFiles(session.DirectoryPath, "*", SearchOption.AllDirectories)) session.TrackFile(file);
    }
    private static IEnumerable<MigrationProgress> Observed(IEnumerable<MigrationProgress> steps, Action<MigrationProgress> progress)
    { foreach (var step in steps) { if (progress != null) progress(step); yield return step; } }
    private static void RunFormBatch(DataMigrationHost host, MigrationRequest request, Action<MigrationProgress> progress = null)
    {
        Exception failure = null;
        using (var context = new ApplicationContext()) using (var form = new DataMigrationForm())
        {
            LoadRequest(form, request);
            Action<Exception> failed = delegate(Exception error) { failure = error; context.ExitThread(); };
            form.Run(host.Preview(request), delegate {
                form.RefreshResults();
                form.Run(Observed(host.Execute(),progress), delegate { form.RefreshResults(); Console.WriteLine("FORM_SUMMARY: " + DataMigrationForm.Summary(request)); context.ExitThread(); }, failed);
            }, failed);
            Application.Run(context);
        }
        if (failure != null) throw failure;
    }
    private static void Make(ProbeSession session, object app, string path, int format, string sheetName, Action<object, object> populate)
    {
        using (var refs = session.References())
        {
            dynamic books = refs.Take(((dynamic)app).Workbooks); object created = null; bool closed = false;
            try
            {
                created = session.Own((object)books.Add(), delegate { if (!closed) { ((dynamic)created).Close(false); closed = true; } });
                dynamic book = created; dynamic sheets = refs.Take(book.Worksheets); object sheet = refs.Take(sheets.Item(1));
                ((dynamic)sheet).Name = sheetName; populate(created, sheet); book.SaveAs(path, format);
                book.Close(false); closed = true;
            }
            finally { if (!closed && created != null) { ((dynamic)created).Close(false); closed = true; } }
        }
    }
    private static void Cell(ProbeSession session, object sheet, string address, object value, bool formula)
    {
        using (var refs = session.References())
        {
            dynamic cell = refs.Take(((dynamic)sheet).Range[address]);
            if (formula) cell.Formula = value; else { if (value is string) cell.NumberFormat = "@"; cell.Value2 = value; }
        }
    }
    private static void PickerChecks(object application, string source)
    {
        string hash = DataMigrationHost.Digest(source);
        using (var host = new DataMigrationHost(application))
        {
            Check(host.PickSourceRange(source, "源表", "", s => ((dynamic)s).Range["B1:C3"]) == "B1:C3", "WPS source picker validates native range identity and returns coordinates");
        }
        using (var host = new DataMigrationHost(application)) Check(host.PickSourceRange(source, "源表", "B1:C3", s => false) == null, "WPS source picker cancellation returns no replacement");
        Check(DataMigrationHost.Digest(source) == hash, "source picker preserves original file bytes");
        Check(!Directory.GetFiles(Path.GetDirectoryName(source), ".xy-migration-read-*").Any(), "source picker and cancellation remove their disposable read copies");
    }
    [STAThread] private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        using (var session = new ProbeSession("data-migration-live"))
        {
            int code = 1;
            try
            {
                object application = session.Take(ProbeSession.Connect(new[] { "KET.Application", "et.Application" }, false));
                if (application == null) { Console.WriteLine("SKIP: WPS 表格 COM 不可用。"); return session.Finish(2); }
                dynamic app = application;
                Console.WriteLine("HOST: " + app.Name + " " + app.Version + " PATH=" + app.Path);
                object originalBook = null, originalSheet = null;
                try { originalBook = session.Take(app.ActiveWorkbook); originalSheet = session.Take(app.ActiveSheet); } catch { }
                session.OnClose(delegate { if (originalBook != null) ((dynamic)originalBook).Activate(); if (originalSheet != null) ((dynamic)originalSheet).Activate(); });
                object visible = app.Visible;
                session.Override(() => (object)app.Visible, value => app.Visible = (dynamic)value, visible, "Visible");
                session.Override(() => (object)app.DisplayAlerts, value => app.DisplayAlerts = (dynamic)value, false, "DisplayAlerts");
                session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = (dynamic)value, false, "ScreenUpdating");
                session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, false, "EnableEvents");
                if (args.Length == 0 || (args.Length == 1 && (args[0] == "--quiet-replace" || args[0] == "--quiet-replace-ui")))
                {
                    // Represent the visible user workbook which owns the add-in
                    // dialog. No real user document is used as the test host.
                    string parking = session.Output("静默迁移_宿主测试.xlsx");
                    using (var focus = new MigrationWindowScope(application))
                    {
                        Make(session, application, parking, 51, "宿主", delegate(object b, object s) { Cell(session, s, "A1", "WPS 静默迁移测试宿主", false); });
                        using (var references = session.References())
                        {
                            dynamic books = references.Take(app.Workbooks); object opened = books.Open(parking, 0, true);
                            session.Own(opened, delegate { ((dynamic)opened).Close(false); });
                            MigrationWindowScope.SetOwnedVisible(opened, true);
                        }
                    }
                }
                if ((args.Length==4&&(args[0]=="--repro"||args[0]=="--diagnose"))||(args.Length==5&&args[0]=="--repair")) { Reproduce(session, application, args); return session.Finish(0); }
                string third = session.Output("第三方.xlsx"), source = session.Output("源.xlsx");
                Make(session, application, third, 51, "明细", delegate(object book, object sheet) { Cell(session, sheet, "A1", 37.0, false); });
                Make(session, application, source, 51, "源表", delegate(object book, object sheet)
                {
                    Cell(session, sheet, "A1", 10.0, false);
                    // Declare the name before its first formula use; WPS otherwise
                    // serializes the unresolved identifier as _xleta.Rate.
                    using (var refs = session.References())
                    {
                        dynamic names = refs.Take(((dynamic)book).Names); refs.Take(names.Add("Rate", "='源表'!$A$1", true));
                    }
                    string external = "'" + Path.Combine(Path.GetDirectoryName(third), "[" + Path.GetFileName(third) + "]明细").Replace("'", "''") + "'!A1";
                    Cell(session, sheet, "B1", "=A1+$A$1+" + external + "+Rate", true);
                    Cell(session, sheet, "C1", "=普通文本", false); Cell(session, sheet, "B2", 123.5, false);
                    using (var refs = session.References())
                    {
                        dynamic merge = refs.Take(((dynamic)sheet).Range["B3:C3"]); merge.Merge(); merge.Value2 = "合并标题";
                        dynamic font = refs.Take(merge.Font); font.Bold = true; font.Color = 255;
                        dynamic fill = refs.Take(merge.Interior); fill.Color = 15658734;
                    }
                });
                PickerChecks(application, source);
                if (args.Length == 1 && args[0] == "--picker") return session.Finish(0);
                if (args.Length == 1 && args[0] == "--picker-ui")
                {
                    session.Override(() => (object)app.Visible, value => app.Visible = (dynamic)value, true, "picker visibility");
                    session.Override(() => (object)app.ScreenUpdating, value => app.ScreenUpdating = (dynamic)value, true, "picker redraw");
                    string originalHash = DataMigrationHost.Digest(source);
                    using (var host = new DataMigrationHost(application))
                    {
                        string picked = host.PickSourceRange(source, "源表", "", null);
                        Console.WriteLine("MOUSE_SELECTION: " + picked);
                        Check(picked == "B1:C3", "native mouse drag and OK return B1:C3");
                        Check(host.PickSourceRange(source, "源表", picked, null) == null, "native cancel keeps the existing address available");
                    }
                    Check(DataMigrationHost.Digest(source) == originalHash && !Directory.GetFiles(session.DirectoryPath, ".xy-migration-read-*").Any(), "native mouse selection and cancel leave originals intact and remove own copies");
                    return session.Finish(0);
                }
                var targets = new List<MigrationTarget>();
                foreach (var pair in new[] { Tuple.Create(".xlsx", 51), Tuple.Create(".xlsm", 52), Tuple.Create(".xls", 56), Tuple.Create(".et", -4143) })
                {
                    string target = session.Output("目标" + pair.Item1);
                    Make(session, application, target, pair.Item2, "目标表", delegate(object book, object sheet)
                    {
                        Cell(session, sheet, "A1", 20.0, false); Cell(session, sheet, "A100", 30.0, false); Cell(session, sheet, "B100", "原内容", false);
                        using (var refs = session.References()) { dynamic cell = refs.Take(((dynamic)sheet).Range["B100"]); cell.ColumnWidth = 19; cell.RowHeight = 24; }
                    });
                    targets.Add(new MigrationTarget { Path = target, Sheet = "目标表", Start = "B100" });
                }
                var hashes = targets.Select(t => t.Path).Concat(new[] { source, third }).ToDictionary(p => p, DataMigrationHost.Digest);
                DataMigrationHost.DiagnosticRootOverride = Path.Combine(session.DirectoryPath, "diagnostics");
                using (var host = new DataMigrationHost(application))
                {
                    Drain(host.Preview(new MigrationRequest { SourcePath = source, SourceSheet = "源表", SourceRange = "B1:C3", Targets = targets }));
                    foreach (var target in targets) Check(target.Ready, "预检 " + Path.GetExtension(target.Path) + ": " + target.Error);
                    Drain(host.Execute());
                    foreach (var target in targets)
                    {
                        Check(target.Status == "成功" && File.Exists(target.Output), "迁移并重开校验 " + Path.GetExtension(target.Path) + ": " + target.Error);
                        session.TrackFile(target.Output);
                    }
                }
                foreach (var pair in hashes) Check(DataMigrationHost.Digest(pair.Key) == pair.Value, "原文件摘要不变：" + Path.GetFileName(pair.Key));
                if (args.Length == 1 && (args[0] == "--quiet-replace" || args[0] == "--quiet-replace-ui"))
                {
                    int baseline; using (var references = session.References()) { dynamic books = references.Take(app.Workbooks); baseline = Convert.ToInt32(books.Count); }
                    int peak = baseline; bool hidden = true;
                    Action<MigrationProgress> inspect = delegate(MigrationProgress step)
                    {
                        using (var references = session.References())
                        {
                            dynamic books = references.Take(app.Workbooks); int total = Convert.ToInt32(books.Count); peak = Math.Max(peak, total);
                            for (int i = 1; i <= total; i++)
                            {
                                dynamic book = references.Take(books.Item(i)); string path = Convert.ToString(book.FullName);
                                if (!path.StartsWith(session.DirectoryPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(path).StartsWith(".xy-migration-")) continue;
                                dynamic windows = references.Take(book.Windows);
                                for (int j = 1; j <= Convert.ToInt32(windows.Count); j++) { dynamic window = references.Take(windows.Item(j)); hidden &= !(bool)window.Visible; }
                            }
                        }
                    };
                    var replacement = new MigrationRequest { SourcePath = source, SourceSheet = "源表", SourceRange = "B1:C3", Targets = targets, ReplaceTargets = true };
                    using (var host = new DataMigrationHost(application))
                    {
                        if (args[0] == "--quiet-replace-ui") InteractiveRun(application, host, replacement, inspect);
                        else
                        {
                        DrainPumped(Observed(host.Preview(replacement), inspect));
                        foreach (var target in targets) Check(target.Ready, "覆盖模式预检 " + Path.GetExtension(target.Path) + ": " + target.Error);
                        DrainPumped(Observed(host.Execute(), inspect));
                        }
                    }
                    Check(hidden && peak <= baseline + 2, "临时工作簿在每个消息泵让步点隐藏，数量不随目标数增长；峰值=" + peak);
                    foreach (var target in targets)
                    {
                        Check(target.Status == "成功" && target.Output == target.Path && File.Exists(target.Backup), "原文件已更新并保留恢复版本 " + Path.GetExtension(target.Path) + ": " + target.Error);
                        Check(DataMigrationHost.Digest(target.Backup) == hashes[target.Path], "恢复版本字节匹配迁移前目标 " + Path.GetExtension(target.Path));
                        session.TrackFile(target.Backup);
                        using (var references = session.References())
                        {
                            dynamic books = references.Take(app.Workbooks), book = references.Take(books.Open(target.Path, 0, true));
                            try
                            {
                                dynamic sheets = references.Take(book.Worksheets), sheet = references.Take(sheets.Item("目标表")), cell = references.Take(sheet.Range["B100"]);
                                dynamic windows = references.Take(book.Windows), window = references.Take(windows.Item(1));
                                Check((bool)window.Visible && Convert.ToString(cell.Formula).Contains("A100"), "目标原文件重开可见，公式已迁移 " + Path.GetExtension(target.Path));
                            }
                            finally { book.Close(false); }
                        }
                    }
                    Check(DataMigrationHost.Digest(source) == hashes[source] && DataMigrationHost.Digest(third) == hashes[third], "覆盖模式仍不修改源文件与第三方来源");
                    foreach (string file in Directory.GetFiles(session.DirectoryPath, "*", SearchOption.AllDirectories)) session.TrackFile(file);
                    return session.Finish(0);
                }
                if (args.Length == 1)
                {
                    string imported = session.Output("生成器测试源.xlsx"), targetPath = session.Output("生成器目标.xlsx");
                    File.Copy(Path.GetFullPath(args[0]), imported, false);
                    string importedHash = DataMigrationHost.Digest(imported); int bookCount;
                    using (var refs = session.References()) { dynamic books = refs.Take(app.Workbooks); bookCount = Convert.ToInt32(books.Count); }
                    using (var host = new DataMigrationHost(application)) Check(host.Sheets(imported).SequenceEqual(new[] { "迁移源", "辅助明细" }), "直接读取生成器测试源的表名");
                    using (var refs = session.References()) { dynamic books = refs.Take(app.Workbooks); Check(Convert.ToInt32(books.Count) == bookCount, "选择源文件不打开或关闭宿主工作簿"); }
                    var generatedTargets = new List<MigrationTarget>(); var generatedHashes = new Dictionary<string, string>();
                    foreach (string generatedPath in new[] { targetPath, session.Output("生成器目标 (2).xlsx") })
                    {
                        Make(session, application, generatedPath, 51, "目标表", delegate(object book, object sheet)
                        {
                            Cell(session, sheet, "C6", 10.0, false);
                            using (var refs = session.References()) { dynamic sheets = refs.Take(((dynamic)book).Worksheets); object auxiliary = refs.Take(sheets.Add()); ((dynamic)auxiliary).Name = "辅助明细"; Cell(session, auxiliary, "B2", 100.0, false); }
                        });
                        generatedHashes[generatedPath] = DataMigrationHost.Digest(generatedPath);
                        generatedTargets.Add(new MigrationTarget { Path = generatedPath, Sheet = "目标表", Start = "B100" });
                    }
                    session.Override(() => (object)app.EnableEvents, value => app.EnableEvents = (dynamic)value, true, "generated source form events");
                    using (var host = new DataMigrationHost(application))
                    {
                        RunFormBatch(host, new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "B5:I22", Targets = generatedTargets });
                        foreach (var target in generatedTargets)
                        {
                            Check(target.Status == "成功", "正式窗口消息泵完成两个目标的 144 格迁移与重开校验：" + target.Error);
                            if (target.Output.Length > 0) session.TrackFile(target.Output);
                            using (var refs = session.References())
                            {
                                dynamic books = refs.Take(app.Workbooks); dynamic result = refs.Take(books.Open(target.Output, 0, true));
                                try
                                {
                                    dynamic outputSheet = refs.Take(result.ActiveSheet), header = refs.Take(outputSheet.Range["B100"]), formula = refs.Take(outputSheet.Range["D101"]), date = refs.Take(outputSheet.Range["D113"]);
                                    Check(Convert.ToString(outputSheet.Name) == "目标表" && Convert.ToString(header.Value2) == "测试项目" && Convert.ToString(formula.Formula) == "=C101", "重新打开结果显示目标表且包含标题和实际公式");
                                    Check(Convert.ToDouble(date.Value2) == new DateTime(2026, 10, 29).ToOADate() && Convert.ToString(date.Text) == "2026-10-29", "ISO 日期迁移后仍为正确日历日期而非 1905 年");
                                    Check(Path.GetFileName(target.Output).StartsWith(Path.GetFileNameWithoutExtension(target.Path) + "_迁移结果", StringComparison.Ordinal), "中文正式结果文件名保留目标原名");
                                    dynamic windows = refs.Take(result.Windows), window = refs.Take(windows.Item(1));
                                    Check(Convert.ToInt32(window.ScrollRow) == 100, "重新打开结果直接显示迁移位置而非空白 A1");
                                }
                                finally { result.Close(false); }
                            }
                        }
                        Check(generatedTargets.Select(t => t.Output).Distinct().Count() == 2, "两个目标生成两个独立结果副本");
                    }
                    Check(DataMigrationHost.Digest(imported) == importedHash && generatedHashes.All(p => DataMigrationHost.Digest(p.Key) == p.Value), "生成器样本原文件摘要不变");
                    var circular = new MigrationTarget { Path = targetPath, Sheet = "目标表", Start = "A1" };
                    using (var host = new DataMigrationHost(application))
                    {
                        Drain(host.Preview(new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "B5:I22", Targets = new List<MigrationTarget> { circular } }));
                        Check(!circular.Ready && circular.Error.Contains("C4") && circular.Error.Contains("自身") && circular.Output.Length == 0, "A1 起点引入直接自引用时预检明确阻止，未生成结果");
                    }
                    string reversedPath = session.Output("逆序范围目标.xlsx");
                    Make(session, application, reversedPath, 51, "目标表", delegate(object book, object sheet) { Cell(session, sheet, "C6", 100.0, false); Cell(session, sheet, "B5", 40.0, false); });
                    var reversed = new MigrationTarget { Path = reversedPath, Sheet = "目标表", Start = "C17" };
                    using (var host = new DataMigrationHost(application))
                    {
                        RunFormBatch(host, new MigrationRequest { SourcePath = imported, SourceSheet = "迁移源", SourceRange = "D21", Targets = new List<MigrationTarget> { reversed } });
                        Check(reversed.Status == "成功", "WPS 将逆序范围改为 B5:$C$6 后仍通过等价核验：" + reversed.Error);
                        if (reversed.Output.Length > 0) session.TrackFile(reversed.Output);
                    }
                }
                using (var form = new DataMigrationForm()) Check(!form.Visible, "真实宿主存在时构造窗口不自动运行迁移");
                // Every file under this unique fixture directory was generated by
                // this session; TrackFile still enforces the cleanup boundary.
                foreach (string file in Directory.GetFiles(session.DirectoryPath, "*", SearchOption.AllDirectories)) session.TrackFile(file);
                Console.WriteLine("PASS: four WPS file formats; UI focus and mouse interaction still require manual acceptance."); code = 0;
            }
            catch (Exception error) { Console.WriteLine("FAIL: " + error); }
            finally { DataMigrationHost.DiagnosticRootOverride = null; }
            return session.Finish(code);
        }
    }
}
