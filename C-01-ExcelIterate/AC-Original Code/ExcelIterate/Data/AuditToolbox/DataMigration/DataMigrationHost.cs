using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class MigrationTarget
    {
        internal string Path = "", Sheet = "", Start = "", Status = "待预检", Error = "", Output = "", Fingerprint = "";
        internal int Overwritten;
        internal string FailureDetails = "", Backup = "";
        internal bool Ready;
        internal MigrationArea Area;
        internal MigrationFormulaContext Context;
        internal List<MigrationFormula> Formulas;
        internal List<MigrationName> Names;
        internal MergeWorkbookHandle Handle;
        internal MergeWorkbookGateway Gateway;
        internal List<MigrationName> ExistingNames;
        internal HashSet<string> ExistingLinks;
        internal readonly List<double> Heights = new List<double>(), Widths = new List<double>();
        internal readonly List<MigrationArea> ActualMerges = new List<MigrationArea>();
        internal Dictionary<string, string> MacroParts;
        internal List<string> Samples = new List<string>();
    }
    internal sealed class MigrationRequest
    {
        internal string SourcePath = "", SourceSheet = "", SourceRange = "";
        internal bool AllowEmptySource;
        internal bool ReplaceTargets;
        // 只写数值: formulas in the source are read as their calculated values
        // and written as constants; no formula translation or names involved.
        internal bool ValuesOnly;
        internal int NonEmptyCells, FormulaCells;
        internal List<MigrationTarget> Targets = new List<MigrationTarget>();
    }
    internal sealed class MigrationCell
    {
        internal bool IsFormula;
        internal object Value;
        internal string Formula;
        internal Dictionary<string, object> Style;
    }
    internal sealed class MigrationProgress
    {
        internal string Message;
        internal MigrationProgress(string message) { Message = message; }
    }
    // Only this adapter touches COM. The enumerators are pumped on the original
    // host thread, yielding between cells/targets without DoEvents or Task.Run.
    internal sealed class DataMigrationHost : IDisposable
    {
        private readonly object app;
        internal static string DiagnosticRootOverride;
        private readonly int thread = Thread.CurrentThread.ManagedThreadId;
        private readonly MergeWorkbookGateway gateway;
        private readonly Dictionary<string, string> readCopies = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<MergeWorkbookHandle, string> originalPaths = new Dictionary<MergeWorkbookHandle, string>();
        private SourceChangeWatch watch;
        private SourceChangeWatch copyWatch;
        private MergeWorkbookHandle source;
        private MigrationRequest request;
        private MigrationArea area;
        private readonly List<MigrationCell> cells = new List<MigrationCell>();
        private readonly List<MigrationArea> merges = new List<MigrationArea>();
        private List<MigrationName> sourceNames;
        private Dictionary<string, string> sourceLinks;
        private string sourceHash;
        private MigrationPackagePreservation packagePreservation;
        private object activeBook, activeSheet;
        private object ownedCopy;
        private ComBag copyReferences;
        private readonly List<string> diagnosticPaths = new List<string>();
        private readonly System.Collections.Concurrent.ConcurrentQueue<string> eventTrace = new System.Collections.Concurrent.ConcurrentQueue<string>();
        internal string EventTrace { get { return string.Join(Environment.NewLine,eventTrace.ToArray()); } }
        private SourceChangeWatch Watch(object book,string role)
        { var monitor=new SourceChangeWatch(app,true){Role=role,Trace=delegate(string value){if(eventTrace.Count<2048)eventTrace.Enqueue(DateTime.UtcNow.ToString("o")+" "+value);}};monitor.WatchBook(book);eventTrace.Enqueue(role+" subscribed="+monitor.Available);return monitor; }
        internal bool Cancelled;
        internal string DiagnosticDirectory { get; private set; }
        private static readonly string[] RangeStyle = { "NumberFormat", "HorizontalAlignment", "VerticalAlignment", "WrapText", "Orientation", "IndentLevel", "ShrinkToFit", "Locked", "FormulaHidden" };
        private static readonly string[] FontStyle = { "Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color" };
        private static readonly string[] FillStyle = { "Pattern", "Color", "PatternColor" };

        internal DataMigrationHost(object application)
        {
            app = application; gateway = new MergeWorkbookGateway(app);
            using (var bag = new ComBag())
            { activeBook = gateway.Hold(bag.Detach(bag.Own(((dynamic)app).ActiveWorkbook))); activeSheet = gateway.Hold(bag.Detach(bag.Own(((dynamic)app).ActiveSheet))); }
        }
        private void Check()
        {
            if (Thread.CurrentThread.ManagedThreadId != thread) throw new InvalidOperationException("迁移只能在原宿主线程执行。");
            if (Cancelled) throw new OperationCanceledException("已取消；已成功目标及其备份保留，尚未提交的目标不变。");
        }
        internal List<string> Sheets(string path)
        {
            Check(); CheckFile(path);
            // Choosing an OOXML file must not open/activate/close a workbook in
            // WPS just to populate a dropdown. In particular, freshly generated
            // files can be recalculated on open and become Saved=false.
            string extension = Path.GetExtension(path);
            if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsm", StringComparison.OrdinalIgnoreCase))
                return PackageSheets(path);
            var handle = ResolveOriginal(path, ""); Saved(handle);
            var result = new List<string>();
            using (var bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)handle.Book).Worksheets);
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) using (var item = new ComBag()) { dynamic sheet = item.Own(sheets.Item(i)); result.Add(Convert.ToString(sheet.Name)); }
            }
            RestoreActive(); return result;
        }
        internal static List<string> PackageSheets(string path)
        {
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                var entry = package.GetEntry("xl/workbook.xml");
                if (entry == null) throw new InvalidOperationException("工作簿缺少工作表目录：" + path);
                var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024 };
                var result = new List<string>();
                using (var xml = XmlReader.Create(entry.Open(), settings))
                    while (xml.Read())
                        if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "sheet" &&
                            (xml.NamespaceURI == "http://schemas.openxmlformats.org/spreadsheetml/2006/main" || xml.NamespaceURI == "http://purl.oclc.org/ooxml/spreadsheetml/main"))
                        {
                            string name = xml.GetAttribute("name");
                            if (string.IsNullOrEmpty(name) || result.Contains(name)) throw new InvalidOperationException("工作表目录包含空名称或重复名称。");
                            result.Add(name);
                        }
                if (result.Count == 0) throw new InvalidOperationException("工作簿没有工作表。");
                return result;
            }
        }
        internal static string Digest(string path)
        { using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", ""); }
        internal string PickSourceRange(string path, string sheetName, string current, Func<object, object> picker, string role = "源")
        {
            Check(); CheckFile(path);
            if (string.IsNullOrWhiteSpace(sheetName)) throw new InvalidOperationException("请先选择" + role + "工作表。");
            string full = Path.GetFullPath(path); bool open = false;
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++) using (var item = new ComBag())
                { dynamic book = item.Own(books.Item(i)); if (string.Equals(Convert.ToString(book.FullName), full, StringComparison.OrdinalIgnoreCase)) { open = true; break; } }
            }
            // Selecting coordinates may borrow an unsaved open document; it
            // never saves it. The normal preflight still requires saved inputs.
            // Closed files are shown through our readonly disposable copy.
            var handle = open ? gateway.Resolve(full, sheetName, true) : ResolveOriginal(full, sheetName, null, true);
            ((dynamic)handle.Book).Activate(); ((dynamic)handle.Sheet).Activate();
            object initial = Type.Missing;
            if (!string.IsNullOrWhiteSpace(current)) try { initial = MigrationArea.Parse(current).Address; } catch (InvalidOperationException) { }
            using (var refs = new ComBag())
            {
                object chosen = picker != null ? picker(handle.Sheet) : ((dynamic)app).InputBox(
                    "请用鼠标框选“" + sheetName + "”中的连续单元格区域，然后点确定。取消会保留原区域。",
                    "选择" + role + "区域 · " + Path.GetFileName(full), initial, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (chosen == null || chosen is bool) return null;
                chosen = refs.Own(chosen);
                return SourceRangeAddress(chosen, handle.Book, handle.Sheet);
            }
        }
        internal static string SourceRangeAddress(object chosen, object expectedBook, object expectedSheet)
        {
            using (var refs = new ComBag())
            {
                dynamic range = chosen, sheet = refs.Own(range.Parent), book = refs.Own(sheet.Parent), areas = refs.Own(range.Areas);
                if (!WorksheetIdentity.SameObject(sheet, expectedSheet) || !WorksheetIdentity.SameObject(book, expectedBook)) throw new InvalidOperationException("请选择指定源工作簿和源工作表中的区域，不接受其他工作簿或工作表的选区。");
                if (Convert.ToInt32(areas.Count) != 1) throw new InvalidOperationException("只支持一个连续矩形区域，请不要按 Ctrl 选择多个区域。");
                dynamic rows = refs.Own(range.Rows), columns = refs.Own(range.Columns);
                var area = new MigrationArea { Row = Convert.ToInt32(range.Row), Column = Convert.ToInt32(range.Column), Rows = Convert.ToInt32(rows.Count), Columns = Convert.ToInt32(columns.Count) };
                if ((long)area.Rows * area.Columns > 100000) throw new InvalidOperationException("单批最多选择 100000 个单元格，请缩小区域。");
                if (area.Row < 1 || area.Column < 1 || area.Rows < 1 || area.Columns < 1) throw new InvalidOperationException("选区地址无效。");
                return area.Address;
            }
        }
        private static void CheckFile(string path)
        {
            if (!DataMigrationEngine.SupportedFile(path) || !File.Exists(path) || !Path.IsPathRooted(path)) throw new InvalidOperationException("请选择已保存的 .xlsx/.xlsm/.xls/.et 文件：" + path);
        }
        private void Saved(MergeWorkbookHandle handle)
        {
            string original, ownedPath;
            // Re-resolving our still-open copy returns OpenedByUs=false from
            // the shared gateway. Ownership comes from our exact copy map.
            bool ownedReadCopy = originalPaths.TryGetValue(handle, out original) && readCopies.TryGetValue(original, out ownedPath) && string.Equals(ownedPath, handle.Path, StringComparison.OrdinalIgnoreCase);
            // A readonly inspection copy is allowed to be dirty from WPS's own
            // import/recalculation. Never clear Saved or save the user's file.
            // SourceChangeWatch, identity and original-file hashes still guard it.
            if (!ownedReadCopy && !ComFormulaHost.Boolean(((dynamic)handle.Book).Saved, "Saved")) throw new InvalidOperationException("工作簿有未保存修改，请先保存并重新预检：" + handle.Path);
            if (!string.Equals(Path.GetFullPath(Convert.ToString(((dynamic)handle.Book).FullName)), Path.GetFullPath(handle.Path), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("工作簿身份或路径已改变，请重新预检。");
        }
        private MergeWorkbookHandle ResolveOriginal(string path, string sheet, MergeWorkbookGateway targetGateway = null, bool interactive = false)
        {
            string full = Path.GetFullPath(path), readPath = full; bool alreadyOpen = false;
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++) using (var item = new ComBag())
                {
                    dynamic book = item.Own(books.Item(i));
                    if (string.Equals(Convert.ToString(book.FullName), full, StringComparison.OrdinalIgnoreCase))
                    { if (!ComFormulaHost.Boolean(book.Saved, "Saved")) throw new InvalidOperationException("工作簿有未保存修改，请先保存并重新预检：" + full); alreadyOpen = true; break; }
                }
            }
            // Some WPS binary-format readers update compound-file metadata on
            // Close(false), even after a read-only open. Never open an unopened
            // original: inspect a same-directory disposable copy instead.
            bool needsDateCopy = alreadyOpen && MigrationPackageCompatibility.HasIsoDates(full);
            if ((!alreadyOpen || needsDateCopy) && !readCopies.TryGetValue(full, out readPath))
            {
                readPath = Path.Combine(Path.GetDirectoryName(full), ".xy-migration-read-" + Guid.NewGuid().ToString("N") + Path.GetExtension(full));
                try
                {
                    string before = Digest(full);
                    File.Copy(full, readPath, false); readCopies[full] = readPath;
                    if (Digest(full) != before || Digest(readPath) != before) throw new IOException("复制读取副本期间原文件已变化，请重新预检。");
                    MigrationPackageCompatibility.NormalizeOwnedCopy(readPath); File.SetAttributes(readPath, File.GetAttributes(readPath) | FileAttributes.Hidden);
                }
                catch (Exception error) { throw new InvalidOperationException("创建读取副本失败：" + full + "；" + error.Message, error); }
            }
            var resolver = targetGateway ?? gateway;
            var handle = Quiet<MergeWorkbookHandle>(delegate
            {
                var resolved = resolver.Resolve(readPath, sheet, true);
                if (resolved.OpenedByUs && !interactive) MigrationWindowScope.SetOwnedVisible(resolved.Book, false);
                return resolved;
            }); originalPaths[handle] = full; return handle;
        }
        private void ValidateTarget(MigrationTarget target)
        {
            if (target.Handle != null) ValidateOriginal(target.Handle, target.Fingerprint);
            else if (!File.Exists(target.Path) || Digest(target.Path) != target.Fingerprint)
                throw new InvalidOperationException("目标文件已变化，请重新预检：" + target.Path);
            EnsureTargetClosed(target);
        }
        private void EnsureTargetClosed(MigrationTarget target)
        {
            using (var references = new ComBag())
            {
                dynamic books = references.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic book = references.Own(books.Item(i));
                    if (string.Equals(Path.GetFullPath(Convert.ToString(book.FullName)), Path.GetFullPath(target.Path), StringComparison.OrdinalIgnoreCase))
                    {
                        if (request.ReplaceTargets) throw new InvalidOperationException("更新原文件前请先保存并关闭该目标工作簿，再重新预检：" + target.Path);
                        if (!ComFormulaHost.Boolean(book.Saved, "Saved")) throw new InvalidOperationException("目标有未保存修改，请先保存并重新预检：" + target.Path);
                    }
                }
            }
            if (!request.ReplaceTargets) return;
            if ((File.GetAttributes(target.Path) & FileAttributes.ReadOnly) != 0)
                throw new InvalidOperationException("目标文件为只读，不能更新原文件：" + target.Path);
            using (File.Open(target.Path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        }
        private void ReleaseTarget(MigrationTarget target)
        {
            var resolver = target.Gateway; target.Gateway = null;
            try { if (resolver != null) Quiet(delegate { resolver.Dispose(); }); }
            finally
            {
                target.Handle = null;
                string full = Path.GetFullPath(target.Path), path;
                foreach (var handle in originalPaths.Where(p => string.Equals(p.Value, full, StringComparison.OrdinalIgnoreCase)).Select(p => p.Key).ToArray()) originalPaths.Remove(handle);
                if (readCopies.TryGetValue(full, out path))
                {
                    try { File.Delete(path); readCopies.Remove(full); }
                    catch (Exception error) { UiDiagnostics.Record("migration-target-read-cleanup", error); }
                }
            }
        }
        private void ValidateOriginal(MergeWorkbookHandle handle, string hash)
        {
            Saved(handle); bool found = false;
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                { using (var item = new ComBag()) { object book = item.Own(books.Item(i)); if (WorksheetIdentity.SameObject(book, handle.Book)) { found = true; break; } } }
            }
            string path; if (!originalPaths.TryGetValue(handle, out path)) path = handle.Path;
            if (!found || Digest(path) != hash) throw new InvalidOperationException("原工作簿已关闭、替换或文件发生变化，请重新预检。");
            if (!string.Equals(path, handle.Path, StringComparison.OrdinalIgnoreCase))
                using (var bag = new ComBag())
                {
                    dynamic books = bag.Own(((dynamic)app).Workbooks);
                    for (int i = 1; i <= Convert.ToInt32(books.Count); i++) using (var item = new ComBag())
                    {
                        dynamic book = item.Own(books.Item(i));
                        if (string.Equals(Convert.ToString(book.FullName), path, StringComparison.OrdinalIgnoreCase) && !ComFormulaHost.Boolean(book.Saved, "Saved")) throw new InvalidOperationException("原工作簿出现新的未保存编辑，请先处理后重新预检：" + path);
                    }
                }
        }
        internal IEnumerable<MigrationProgress> Preview(MigrationRequest input)
        {
            Check(); request = input; cells.Clear(); merges.Clear();
            MigrationWindowScope.RequireVisibleHost(app);
            CheckFile(input.SourcePath); area = MigrationArea.Parse(input.SourceRange);
            // A bounded snapshot makes accidental whole-sheet selection actionable.
            if ((long)area.Rows * area.Columns > 100000) throw new InvalidOperationException("单批最多迁移 100000 个单元格，请缩小源区域分批处理。");
            sourceHash = Digest(input.SourcePath);
            source = ResolveOriginal(input.SourcePath, input.SourceSheet); Saved(source); ValidateOriginal(source, sourceHash);
            watch = Watch(source.Book,"迁移来源");
            packagePreservation = MigrationPackagePreservation.Capture(input.SourcePath, input.SourceSheet, area);
            sourceNames = ReadNames(source.Book); sourceLinks = ExternalFiles(source.Book);
            using (var bag = new ComBag())
            {
                dynamic rows = bag.Own(((dynamic)source.Sheet).Rows), columns = bag.Own(((dynamic)source.Sheet).Columns);
                area.Check(Convert.ToInt32(rows.Count), Convert.ToInt32(columns.Count));
            }
            RejectTables(source.Sheet, area);
            for (int r = 0; r < area.Rows; r++) for (int c = 0; c < area.Columns; c++)
            {
                Check(); watch.Validate();
                using (var bag = new ComBag())
                {
                    dynamic cell = bag.Own(((dynamic)source.Sheet).Range[MigrationArea.Cell(area.Row + r, area.Column + c)]);
                    if (!input.ValuesOnly) RejectArray(cell);
                    ReadMerge(cell, area, merges);
                    bool follower = IsMergeFollower(cell), formula = !follower && !input.ValuesOnly && ComFormulaHost.Boolean(cell.HasFormula, "HasFormula");
                    cells.Add(new MigrationCell { IsFormula = formula, Formula = formula ? Formula(cell) : null, Value = formula || follower ? null : (object)cell.Value2, Style = ReadStyle(cell) });
                }
                yield return new MigrationProgress("读取源区域 " + cells.Count + "/" + ((long)area.Rows * area.Columns));
            }
            watch.Validate(); ValidateOriginal(source, sourceHash);
            input.NonEmptyCells = cells.Count(c => c.IsFormula || c.Value != null);
            input.FormulaCells = cells.Count(c => c.IsFormula);
            if (input.NonEmptyCells == 0 && !input.AllowEmptySource) throw new InvalidOperationException("源区域 " + input.SourceSheet + "!" + area.Address + " 全部为空，未生成结果。请选择有数据的区域；如需清空目标，请明确勾选“允许全空白源区域”。");
            foreach (var target in input.Targets)
            {
                try
                {
                    Check(); target.Ready = false; target.Error = ""; target.Backup = ""; target.Fingerprint = ""; target.Overwritten = 0; target.Output = ""; target.Samples.Clear(); target.Heights.Clear(); target.Widths.Clear(); target.Status = "预检中";
                    Exception failure = null;
                    try { PrepareTarget(target); } catch (Exception error) { failure = error; }
                    if (failure != null) { Fail(target, failure); yield return new MigrationProgress(target.Error); continue; }
                    bool stop = false;
                    for (int r = 0; r < area.Rows && !stop; r++) for (int c = 0; c < area.Columns; c++)
                    {
                        Check(); watch.Validate();
                        try
                        {
                            using (var bag = new ComBag())
                            {
                                dynamic cell = bag.Own(((dynamic)target.Handle.Sheet).Range[MigrationArea.Cell(target.Area.Row + r, target.Area.Column + c)]);
                                RejectArray(cell); ReadMerge(cell, target.Area, null);
                                if (ComFormulaHost.Boolean(cell.HasFormula, "HasFormula") || cell.Value2 != null) target.Overwritten++;
                                if (c == 0) target.Heights.Add(Convert.ToDouble(cell.RowHeight));
                                if (r == 0) target.Widths.Add(Convert.ToDouble(cell.ColumnWidth));
                            }
                        }
                        catch (Exception error) { Fail(target, error); stop = true; break; }
                        yield return new MigrationProgress("预检 " + Path.GetFileName(target.Path) + " " + (r * area.Columns + c + 1) + "/" + cells.Count);
                    }
                    if (!stop)
                    {
                        try { ValidateOriginal(target.Handle, target.Fingerprint); target.Ready = true; target.Status = "可执行"; }
                        catch (Exception error) { Fail(target, error); }
                    }
                    yield return new MigrationProgress(target.Status + "：" + target.Path);
                }
                finally
                {
                    try { ReleaseTarget(target); }
                    catch (Exception error) { Fail(target, error); }
                    target.Formulas = null; // Preview retains samples, not N full translated snapshots.
                }
            }
            RestoreActive();
        }
        private void PrepareTarget(MigrationTarget target)
        {
            if (string.IsNullOrWhiteSpace(target.Sheet)) throw new InvalidOperationException("请选择该目标中实际存在的工作表。");
            CheckFile(target.Path);
            EnsureTargetClosed(target);
            if(packagePreservation!=null)packagePreservation.RequireCompatibleTarget(target.Path);
            if (string.Equals(Path.GetFullPath(target.Path), Path.GetFullPath(request.SourcePath), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("目标应为其他工作簿，不能选择源文件。");
            var start = MigrationArea.Parse(target.Start); if (start.Rows != 1 || start.Columns != 1) throw new InvalidOperationException("目标起始位置必须是单个单元格。");
            target.Area = new MigrationArea { Row = start.Row, Column = start.Column, Rows = area.Rows, Columns = area.Columns };
            string fingerprint = Digest(target.Path);
            if (target.Fingerprint.Length != 0 && target.Fingerprint != fingerprint) throw new InvalidOperationException("目标在预检后已变化，请重新预检：" + target.Path);
            target.Fingerprint = fingerprint;
            target.Gateway = new MergeWorkbookGateway(app);
            target.Handle = ResolveOriginal(target.Path, target.Sheet, target.Gateway); Saved(target.Handle); ValidateOriginal(target.Handle, target.Fingerprint);
            if (ComFormulaHost.Boolean(((dynamic)target.Handle.Sheet).ProtectContents, "ProtectContents") || ComFormulaHost.Boolean(((dynamic)target.Handle.Book).ProtectStructure, "ProtectStructure")) throw new InvalidOperationException("目标工作表或工作簿结构受保护。");
            RejectTables(target.Handle.Sheet, target.Area);
            var context = Context(target.Handle.Book, target.Sheet, request.SourcePath);
            context.SourceSheet = request.SourceSheet; context.RowDelta = start.Row - area.Row; context.ColumnDelta = start.Column - area.Column; context.Names = sourceNames; context.ExternalFiles = sourceLinks;
            target.Area.Check(context.MaxRow, context.MaxColumn); target.Context = context;
            target.Formulas = new List<MigrationFormula>();
            foreach (var cell in cells)
            {
                MigrationFormula translated = cell.IsFormula ? DataMigrationEngine.Translate(cell.Formula, context, request.SourceSheet) : null;
                int index = target.Formulas.Count;
                if (translated != null && DataMigrationEngine.DirectSelfReference(translated.Text, target.Sheet, target.Area.Row + index / area.Columns, target.Area.Column + index % area.Columns))
                    throw new InvalidOperationException(MigrationArea.Cell(target.Area.Row + index / area.Columns, target.Area.Column + index % area.Columns) + " 的公式 " + translated.Text + " 将直接引用自身；请调整目标起点或源公式的锁定引用。");
                target.Formulas.Add(translated);
                if (translated != null && target.Samples.Count < 8) target.Samples.Add(cell.Formula + " → " + translated.Text);
            }
            target.Names = DataMigrationEngine.RequiredNames(target.Formulas.Where(f => f != null), context);
            target.ExistingNames = ReadNames(target.Handle.Book); target.ExistingLinks = LinkSet(target.Handle.Book);
            target.MacroParts = MacroFingerprints(target.Path);
            var actualContext = Context(target.Handle.Book, target.Sheet, target.Path);
            foreach (var name in target.Names)
            {
                var actual = target.ExistingNames.FirstOrDefault(n => string.Equals(n.Key, name.Key, StringComparison.OrdinalIgnoreCase));
                if (actual != null && !SameDefinition(actual.Formula, name.Formula, actualContext)) throw new InvalidOperationException("名称定义冲突：" + (name.Scope.Length == 0 ? "" : name.Scope + "!") + name.Name);
                if (actual != null) name.Visible = actual.Visible;
                // A local name in the receiving sheet can shadow a workbook name.
                var shadow = target.ExistingNames.FirstOrDefault(n => name.Scope.Length == 0 && string.Equals(n.Scope, target.Sheet, StringComparison.OrdinalIgnoreCase) && string.Equals(n.Name, name.Name, StringComparison.OrdinalIgnoreCase));
                if (shadow != null && !SameDefinition(shadow.Formula, name.Formula, actualContext)) throw new InvalidOperationException("目标局部名称遮蔽工作簿名称：" + name.Name);
            }
        }
        private MigrationFormulaContext Context(object book, string sheetName, string path)
        {
            var context = new MigrationFormulaContext { SourcePath = path, SourceSheet = sheetName, TargetSheet = sheetName, Names = ReadNames(book), ExternalFiles = ExternalFiles(book) };
            using (var bag = new ComBag())
            {
                dynamic sheets = bag.Own(((dynamic)book).Worksheets);
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) using (var item = new ComBag())
                {
                    dynamic sheet = item.Own(sheets.Item(i)); string name = Convert.ToString(sheet.Name); context.TargetSheets.Add(name);
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                    { dynamic rows = item.Own(sheet.Rows), columns = item.Own(sheet.Columns); context.MaxRow = Convert.ToInt32(rows.Count); context.MaxColumn = Convert.ToInt32(columns.Count); }
                }
            }
            return context;
        }
        // Host settings are restored before each yield so cancellation and normal
        // user interaction never inherit a disabled event loop or alert state.
        private T Quiet<T>(Func<T> action)
        {
            using (var window = new MigrationWindowScope(app))
            {
                dynamic a = app; bool alerts = ComFormulaHost.Boolean(a.DisplayAlerts, "DisplayAlerts"), ask = ComFormulaHost.Boolean(a.AskToUpdateLinks, "AskToUpdateLinks");
                object security = null;
                try { security = a.AutomationSecurity; } catch (Exception error) { if (!ComFormulaHost.Missing(error)) throw; }
                using (var quiet = new HostQuiet(app))
                {
                    try
                    {
                        a.DisplayAlerts = false; a.AskToUpdateLinks = false;
                        if (security != null) a.AutomationSecurity = 3; // Never run workbook macros when opening a copy.
                        return action();
                    }
                    finally
                    { try { if (security != null) a.AutomationSecurity = (dynamic)security; } finally { a.AskToUpdateLinks = ask; a.DisplayAlerts = alerts; } }
                }
            }
        }
        private void Quiet(Action action) { Quiet<bool>(delegate { action(); return true; }); }
        private const int WriteSliceMilliseconds = 150;
        internal IEnumerable<MigrationProgress> Execute()
        {
            Check(); watch.Validate(); ValidateOriginal(source, sourceHash);
            DiagnosticDirectory = Path.Combine(DiagnosticRootOverride ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "DataMigration"), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(DiagnosticDirectory);
            foreach (var target in request.Targets)
            {
                Check(); if (!target.Ready) continue;
                string stage = null; object book = null, sheet = null; ComBag held = new ComBag(); copyReferences = held;
                Exception failure = null;
                try
                {
                    target.Status = "复核源文件";
                    watch.Validate(); ValidateOriginal(source, sourceHash);
                    target.Status = "复核目标文件"; ValidateTarget(target);
                    // Rebuild the preflight immediately before copying to reject
                    // newly introduced name conflicts as well as stale disk data.
                    target.Status = "重新预检目标"; PrepareTarget(target);
                    ReleaseTarget(target);
                    target.Status = "创建结果副本";
                    stage = Path.Combine(Path.GetDirectoryName(target.Path), ".xy-migration-" + Guid.NewGuid().ToString("N") + Path.GetExtension(target.Path));
                    File.Copy(target.Path, stage, false); diagnosticPaths.Add(stage);
                    if (Digest(stage) != target.Fingerprint || Digest(target.Path) != target.Fingerprint) throw new IOException("复制写入副本期间目标文件已变化，请重新预检。");
                    MigrationPackageCompatibility.NormalizeOwnedCopy(stage);
                    target.Status = "打开结果副本"; book = ownedCopy = held.Own(OpenCopy(stage));
                    target.Status = "定位目标工作表"; sheet = held.Own(Sheet(book, target.Sheet));
                    copyWatch = Watch(book,"写入副本 "+Path.GetFileName(target.Path)); RestoreActive();
                    target.ActualMerges.Clear();
                    Quiet(delegate
                    {
                        using (var bag = new ComBag()) { dynamic range = bag.Own(((dynamic)sheet).Range[target.Area.Address]); range.UnMerge(); }
                        foreach (var name in target.Names) AddName(book, name);
                    });
                    target.Status = "写入副本";
                }
                catch (Exception error) { failure = error; }
                try { ReleaseTarget(target); } catch (Exception error) { failure = failure ?? error; }
                if (failure == null)
                {
                    // Write in short time slices instead of one host round-trip
                    // per cell: every Quiet toggles ScreenUpdating/EnableEvents and
                    // the window scope, which made WPS and the dialog flash
                    // continuously. Settings are still restored before each yield.
                    for (int i = 0; i < cells.Count; )
                    {
                        try
                        {
                            Check(); watch.Validate(); copyWatch.Validate(); int first = i;
                            i = Quiet<int>(delegate
                            {
                                var clock = System.Diagnostics.Stopwatch.StartNew(); int index = first;
                                do { WriteCell(sheet, target, index); index++; }
                                while (index < cells.Count && !Cancelled && clock.ElapsedMilliseconds < WriteSliceMilliseconds);
                                return index;
                            });
                        }
                        catch (Exception error) { failure = error; break; }
                        yield return new MigrationProgress("写入 " + Path.GetFileName(target.Path) + " " + i + "/" + cells.Count);
                    }
                }
                if (failure == null)
                {
                    try
                    {
                        Check(); copyWatch.Validate(); target.Status = "保存副本";
                        Quiet(delegate
                        {
                            foreach (var merge in merges) using (var bag = new ComBag())
                            {
                                var shifted = new MigrationArea { Row = merge.Row + target.Context.RowDelta, Column = merge.Column + target.Context.ColumnDelta, Rows = merge.Rows, Columns = merge.Columns };
                                dynamic range = bag.Own(((dynamic)sheet).Range[shifted.Address]); range.Merge();
                            }
                            // WPS recreates border edges when merging, so apply
                            // the captured constituent styles after the merge.
                            foreach (var merge in merges) for (int r = 0; r < merge.Rows; r++) for (int c = 0; c < merge.Columns; c++) using (var bag = new ComBag())
                            {
                                int index = (merge.Row + r - area.Row) * area.Columns + merge.Column + c - area.Column;
                                object cell = bag.Own(((dynamic)sheet).Range[MigrationArea.Cell(merge.Row + r + target.Context.RowDelta, merge.Column + c + target.Context.ColumnDelta)]);
                                WriteStyle(cell, cells[index].Style);
                            }
                            RestoreDimensions(sheet, target);
                            // Save with a visible workbook view so the delivered
                            // file does not reopen hidden. Never yield while it
                            // is visible, and hide again even when Save fails.
                            MigrationWindowScope.SetOwnedVisible(book, true);
                            try
                            {
                                SetOutputView(book, sheet, target); ((dynamic)book).Save();
                                if (!ComFormulaHost.Boolean(((dynamic)book).Saved, "Saved")) throw new IOException("副本保存未完成。");
                            }
                            finally { MigrationWindowScope.SetOwnedVisible(book, false); RestoreActive(); }
                        });
                        copyWatch.Dispose(); copyWatch = null;
                        target.Status = "重开核验";
                        Close(book); book = ownedCopy = null; held.Dispose(); held = new ComBag(); copyReferences = held;
                        if (packagePreservation != null) packagePreservation.Apply(stage, target);
                        book = ownedCopy = held.Own(OpenCopy(stage)); sheet = held.Own(Sheet(book, target.Sheet));
                        copyWatch = Watch(book,"重开副本 "+Path.GetFileName(target.Path)); RestoreActive();
                        VerifyNamesAndLinks(book, target, stage);
                    }
                    catch (Exception error) { failure = error; }
                }
                MigrationFormulaContext verification = null;
                if (failure == null) try { verification = Context(book, target.Sheet, stage); } catch (Exception error) { failure = error; }
                if (failure == null)
                {
                    for (int i = 0; i < cells.Count; i++)
                    {
                        try { Check(); copyWatch.Validate(); VerifyCell(sheet, target, i, verification); }
                        catch (Exception error) { failure = error; break; }
                        yield return new MigrationProgress("重开核验 " + Path.GetFileName(target.Path) + " " + (i + 1) + "/" + cells.Count);
                    }
                }
                try
                {
                    if (failure == null)
                    {
                        VerifyMergeLayout(sheet, target);
                        watch.Validate(); ValidateOriginal(source, sourceHash); ValidateTarget(target);
                    }
                }
                catch (Exception error) { failure = error; }
                try { if (copyWatch != null) { copyWatch.Dispose(); copyWatch = null; } if (book != null) { Close(book); ownedCopy = null; } } catch (Exception error) { failure = failure ?? error; }
                try { held.Dispose(); copyReferences = null; } catch (Exception error) { failure = failure ?? error; }
                if (failure == null)
                {
                    try
                    {
                        Check();
                        var macros = MacroFingerprints(stage);
                        if (packagePreservation != null) packagePreservation.Verify(stage, target);
                        if (target.MacroParts.Count != macros.Count || target.MacroParts.Any(p => !macros.ContainsKey(p.Key) || macros[p.Key] != p.Value)) throw new IOException("保存后 VBA 宏载荷发生变化，副本未交付。");
                        File.SetAttributes(stage, File.GetAttributes(stage) & ~FileAttributes.Hidden);
                        ValidateTarget(target);
                        if (request.ReplaceTargets)
                        {
                            target.Status = "更新目标原文件";
                            target.Backup = MigrationOutputTransaction.Commit(stage, target.Path, target.Fingerprint, Path.GetFileName(DiagnosticDirectory));
                            target.Output = target.Path;
                        }
                        else
                        {
                            string output = DataMigrationEngine.OutputPath(target.Path, p => File.Exists(p) || Directory.Exists(p));
                            File.Move(stage, output); target.Output = output;
                        }
                        diagnosticPaths.Remove(stage); target.Ready = false; target.Status = "成功";
                    }
                    catch (Exception error) { failure = error; }
                }
                if (failure != null)
                {
                    Fail(target, failure);
                    if (stage != null && File.Exists(stage))
                        try
                        {
                            string retained = Path.Combine(DiagnosticDirectory, Path.GetFileNameWithoutExtension(target.Path) + "_未通过校验_" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension(target.Path));
                            File.Move(stage, retained); File.SetAttributes(retained, File.GetAttributes(retained) & ~FileAttributes.Hidden);
                            diagnosticPaths.Remove(stage); diagnosticPaths.Add(retained);
                        }
                        catch (Exception archiveError) { UiDiagnostics.Record("migration-failure-archive", archiveError); }
                }
                target.Formulas = null;
                TryWriteLog(); RestoreActive();
                yield return new MigrationProgress(target.Status + "：" + (target.Output.Length == 0 ? target.Error : target.Output));
                if (failure is OperationCanceledException) { Cancelled = true; break; }
            }
        }
        private object OpenCopy(string path)
        {
            return Quiet<object>(delegate
            {
                using (var bag = new ComBag())
                {
                    dynamic books = bag.Own(((dynamic)app).Workbooks); object book = bag.Own(books.Open(path, 0, false));
                    try { MigrationWindowScope.SetOwnedVisible(book, false); }
                    catch { ((dynamic)book).Close(false); throw; }
                    return bag.Detach(book);
                }
            });
        }
        private void Close(object book) { Quiet(delegate { ((dynamic)book).Close(false); }); if(watch!=null)watch.ProcessPending(); }
        private static object Sheet(object book, string name)
        { using (var bag = new ComBag()) { dynamic sheets = bag.Own(((dynamic)book).Worksheets); return sheets.Item(name); } }
        private void WriteCell(object sheet, MigrationTarget target, int index)
        {
            MigrationCell sourceCell = cells[index];
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[MigrationArea.Cell(target.Area.Row + index / area.Columns, target.Area.Column + index % area.Columns)]);
                if (sourceCell.IsFormula)
                {
                    // A text-formatted destination treats an assigned formula as
                    // text on WPS. Apply a formula-capable format before writing.
                    cell.NumberFormat = "General";
                    try { cell.Formula2 = target.Formulas[index].Text; }
                    catch (Exception error) { if (!ComFormulaHost.Missing(error)) throw; cell.Formula = target.Formulas[index].Text; }
                }
                else if (sourceCell.Value == null)
                {
                    // A blank source cell must really clear existing target
                    // data. WPS can ignore Value2 = null on a filled cell, so
                    // the old content survived and the reopen check failed.
                    cell.ClearContents();
                    if (cell.Value2 != null || ComFormulaHost.Boolean(cell.HasFormula, "HasFormula")) cell.Formula = "";
                }
                else
                {
                    // The target's existing format must not reinterpret the
                    // value (a number written into a text cell becomes text).
                    // Write under a neutral format; WriteStyle then restores
                    // the source format.
                    cell.NumberFormat = sourceCell.Value is string ? "@" : "General";
                    cell.Value2 = sourceCell.Value;
                }
                WriteStyle(cell, sourceCell.Style);
            }
        }
        private void VerifyCell(object sheet, MigrationTarget target, int index, MigrationFormulaContext context)
        {
            MigrationCell expected = cells[index]; string address = MigrationArea.Cell(target.Area.Row + index / area.Columns, target.Area.Column + index % area.Columns);
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[address]);
                bool follower = IsMergeFollower(cell), has = !follower && ComFormulaHost.Boolean(cell.HasFormula, "HasFormula");
                if (has != expected.IsFormula || (has ? !SameDefinition(Formula(cell), target.Formulas[index].Text, context) : !SameValue(follower ? null : (object)cell.Value2, expected.Value))) throw new IOException(address + " 保存重开后内容核对失败。");
                var actual = ReadStyle(cell);
                foreach (var pair in expected.Style) if (!SameValue(actual[pair.Key], pair.Value)) throw new IOException(address + " 样式核对失败：" + pair.Key);
                ReadMerge(cell, target.Area, target.ActualMerges);
                if (index % area.Columns == 0 && Convert.ToDouble(cell.RowHeight) != target.Heights[index / area.Columns]) throw new IOException(address + " 目标行高发生变化：" + cell.RowHeight + "，应为 " + target.Heights[index / area.Columns]);
                if (index < area.Columns && Convert.ToDouble(cell.ColumnWidth) != target.Widths[index]) throw new IOException(address + " 目标列宽发生变化：" + cell.ColumnWidth + "，应为 " + target.Widths[index]);
            }
        }
        private void RestoreDimensions(object sheet, MigrationTarget target)
        {
            // Setting fonts/wrapping can make WPS automatically resize rows.
            // Restore the target's dimensions after all formatting and merging.
            for (int r = 0; r < target.Area.Rows; r++) using (var bag = new ComBag())
            { dynamic cell = bag.Own(((dynamic)sheet).Range[MigrationArea.Cell(target.Area.Row + r, target.Area.Column)]); cell.RowHeight = target.Heights[r]; }
            for (int c = 0; c < target.Area.Columns; c++) using (var bag = new ComBag())
            { dynamic cell = bag.Own(((dynamic)sheet).Range[MigrationArea.Cell(target.Area.Row, target.Area.Column + c)]); cell.ColumnWidth = target.Widths[c]; }
        }
        private static void SetOutputView(object book, object sheet, MigrationTarget target)
        {
            // Persist the destination sheet and viewport in our result copy so
            // opening a B100 result does not look like an empty A1 worksheet.
            ((dynamic)book).Activate(); ((dynamic)sheet).Activate();
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[target.Start]); cell.Select();
                dynamic windows = bag.Own(((dynamic)book).Windows), window = bag.Own(windows.Item(1));
                window.ScrollRow = target.Area.Row; window.ScrollColumn = target.Area.Column;
            }
        }
        private static bool SameValue(object a, object b)
        {
            if (a == null || b == null) return a == b;
            if (a is string || b is string) return a is string && b is string && (string)a == (string)b;
            if (a is bool || b is bool) return a is bool && b is bool && (bool)a == (bool)b;
            var ea = a as ErrorWrapper; var eb = b as ErrorWrapper;
            if (ea != null || eb != null) return ea != null && eb != null && ea.ErrorCode == eb.ErrorCode;
            return string.Equals(Convert.ToString(a, CultureInfo.InvariantCulture), Convert.ToString(b, CultureInfo.InvariantCulture), StringComparison.Ordinal);
        }
        private static bool SameDefinition(string a, string b, MigrationFormulaContext context)
        {
            try { return string.Equals(DataMigrationEngine.ComparisonFormula(a, context), DataMigrationEngine.ComparisonFormula(b, context), StringComparison.Ordinal); }
            catch { return string.Equals(a, b, StringComparison.Ordinal); }
        }
        internal static Dictionary<string, string> MacroFingerprints(string path)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            using (var file = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Binary xls/et and encrypted containers are left to the native
                // reader. ZIP-based workbooks additionally retain VBA bytes.
                if (file.ReadByte() != 'P' || file.ReadByte() != 'K') return result;
                file.Position = 0;
                using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
                    foreach (var entry in zip.Entries.Where(e => e.FullName.IndexOf("vbaProject", StringComparison.OrdinalIgnoreCase) >= 0))
                        using (var stream = entry.Open()) using (var sha = SHA256.Create()) result[entry.FullName] = BitConverter.ToString(sha.ComputeHash(stream));
            }
            return result;
        }
        private static string Formula(dynamic cell)
        { try { return Convert.ToString(cell.Formula2); } catch (Exception error) { if (!ComFormulaHost.Missing(error)) throw; return Convert.ToString(cell.Formula); } }
        private static void RejectArray(dynamic cell)
        {
            if (ComFormulaHost.Boolean(cell.HasArray, "HasArray")) throw new InvalidOperationException("区域包含数组公式。");
            try { if (ComFormulaHost.Boolean(cell.HasSpill, "HasSpill")) throw new InvalidOperationException("区域包含动态溢出公式。"); }
            catch (Exception error) { if (!ComFormulaHost.Missing(error)) throw; }
        }
        private static void RejectTables(object sheet, MigrationArea region)
        {
            using (var bag = new ComBag())
            {
                dynamic tables = null;
                try { tables = bag.Own(((dynamic)sheet).ListObjects); } catch (Exception error) { if (!ComFormulaHost.Missing(error)) throw; return; }
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++) using (var item = new ComBag())
                {
                    dynamic table = item.Own(tables.Item(i)), range = item.Own(table.Range); MigrationArea other = RangeArea(range);
                    if (region.Row < other.Row + other.Rows && other.Row < region.Row + region.Rows && region.Column < other.Column + other.Columns && other.Column < region.Column + region.Columns) throw new InvalidOperationException("区域与结构化表重叠，首版暂不迁移此类区域。");
                }
            }
        }
        private static MigrationArea RangeArea(dynamic range)
        { using (var bag = new ComBag()) { dynamic rows = bag.Own(range.Rows), columns = bag.Own(range.Columns); return new MigrationArea { Row = Convert.ToInt32(range.Row), Column = Convert.ToInt32(range.Column), Rows = Convert.ToInt32(rows.Count), Columns = Convert.ToInt32(columns.Count) }; } }
        private static void ReadMerge(dynamic cell, MigrationArea allowed, List<MigrationArea> collect)
        {
            if (!ComFormulaHost.Boolean(cell.MergeCells, "MergeCells")) return;
            using (var bag = new ComBag())
            {
                dynamic merged = bag.Own(cell.MergeArea); MigrationArea merge = RangeArea(merged);
                if (!allowed.Contains(merge)) throw new InvalidOperationException("区域边界切断合并单元格：" + merge.Address);
                if (collect != null && !collect.Any(m => m.Address == merge.Address)) collect.Add(merge);
            }
        }
        private static bool IsMergeFollower(dynamic cell)
        {
            if (!ComFormulaHost.Boolean(cell.MergeCells, "MergeCells")) return false;
            using (var bag = new ComBag())
            {
                dynamic merged = bag.Own(cell.MergeArea);
                return Convert.ToInt32(cell.Row) != Convert.ToInt32(merged.Row) || Convert.ToInt32(cell.Column) != Convert.ToInt32(merged.Column);
            }
        }
        private void VerifyMergeLayout(object sheet, MigrationTarget target)
        {
            var expected = merges.Select(m => new MigrationArea { Row = m.Row + target.Context.RowDelta, Column = m.Column + target.Context.ColumnDelta, Rows = m.Rows, Columns = m.Columns }.Address).OrderBy(x => x).ToArray();
            if (!expected.SequenceEqual(target.ActualMerges.Select(m => m.Address).OrderBy(x => x))) throw new IOException("保存重开后合并结构核对失败。");
        }
        private static object Get(object value, string name)
        { return value.GetType().InvokeMember(name, BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, value, new object[0], CultureInfo.InvariantCulture); }
        private static void Set(object value, string name, object content)
        { value.GetType().InvokeMember(name, BindingFlags.SetProperty | BindingFlags.OptionalParamBinding, null, value, new[] { content }, CultureInfo.InvariantCulture); }
        private static Dictionary<string, object> ReadStyle(object cell)
        {
            var result = new Dictionary<string, object>();
            using (var bag = new ComBag())
            {
                foreach (string property in RangeStyle) result[property] = Get(cell, property);
                object font = bag.Own(Get(cell, "Font")), fill = bag.Own(Get(cell, "Interior"));
                foreach (string property in FontStyle) result["Font." + property] = Get(font, property);
                foreach (string property in FillStyle) result["Interior." + property] = Get(fill, property);
                dynamic borders = bag.Own(Get(cell, "Borders"));
                for (int i = 5; i <= 12; i++) using (var item = new ComBag())
                {
                    object border = item.Own(borders.Item(i)); string prefix = "Border" + i + "."; object line = Get(border, "LineStyle"); result[prefix + "LineStyle"] = line;
                    if (Convert.ToInt32(line) != -4142) { result[prefix + "Color"] = Get(border, "Color"); result[prefix + "Weight"] = Get(border, "Weight"); }
                }
            }
            return result;
        }
        private static void WriteStyle(object cell, Dictionary<string, object> style)
        {
            using (var bag = new ComBag())
            {
                object font = bag.Own(Get(cell, "Font")), fill = bag.Own(Get(cell, "Interior")); dynamic borders = bag.Own(Get(cell, "Borders"));
                foreach (var pair in style)
                {
                    if (pair.Key == "Interior.Pattern") continue;
                    string[] parts = pair.Key.Split('.');
                    if (parts.Length == 1) Set(cell, parts[0], pair.Value);
                    else if (parts[0] == "Font") Set(font, parts[1], pair.Value);
                    else if (parts[0] == "Interior") Set(fill, parts[1], pair.Value);
                    else using (var item = new ComBag()) { object border = item.Own(borders.Item(int.Parse(parts[0].Substring(6)))); Set(border, parts[1], pair.Value); }
                }
                // Assigning a fill color implicitly switches an empty fill to
                // solid on WPS/Excel. Restore the source pattern LAST.
                Set(fill, "Pattern", style["Interior.Pattern"]);
            }
        }
        private static List<double> Dimensions(object sheet, MigrationArea region, bool rows)
        {
            var result = new List<double>();
            for (int i = 0; i < (rows ? region.Rows : region.Columns); i++) using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[MigrationArea.Cell(region.Row + (rows ? i : 0), region.Column + (rows ? 0 : i))]);
                result.Add(Convert.ToDouble(rows ? cell.RowHeight : cell.ColumnWidth, CultureInfo.InvariantCulture));
            }
            return result;
        }
        private static List<MigrationName> ReadNames(object book)
        {
            var result = new List<MigrationName>();
            using (var bag = new ComBag())
            {
                dynamic names = bag.Own(((dynamic)book).Names);
                for (int i = 1; i <= Convert.ToInt32(names.Count); i++) using (var item = new ComBag())
                {
                    dynamic name = item.Own(names.Item(i)); string label = Convert.ToString(name.Name); int bang = label.LastIndexOf('!');
                    result.Add(new MigrationName { Scope = bang < 0 ? "" : label.Substring(0, bang).Trim('\'').Replace("''", "'"), Name = bang < 0 ? label : label.Substring(bang + 1), Formula = Convert.ToString(name.RefersTo), Visible = ComFormulaHost.Boolean(name.Visible, "Visible") });
                }
            }
            return result;
        }
        private void AddName(object book, MigrationName name)
        {
            if (ReadNames(book).Any(n => string.Equals(n.Key, name.Key, StringComparison.OrdinalIgnoreCase))) return;
            using (var bag = new ComBag())
            {
                dynamic names = bag.Own(((dynamic)book).Names);
                string label = name.Scope.Length == 0 ? name.Name : "'" + name.Scope.Replace("'", "''") + "'!" + name.Name;
                bag.Own(names.Add(label, name.Formula, name.Visible));
            }
        }
        private static HashSet<string> LinkSet(object book)
        {
            object raw = ((dynamic)book).LinkSources(1); var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (raw == null) return result;
            var values = raw as Array; if (values == null) throw new InvalidOperationException("无法可靠读取工作簿外链列表。");
            foreach (object value in values)
            {
                string path = Convert.ToString(value); if (string.IsNullOrEmpty(path)) continue;
                if (path.IndexOf("://", StringComparison.Ordinal) >= 0) throw new InvalidOperationException("首版暂不迁移网络 URL 外链。");
                string directory = Path.GetDirectoryName(Convert.ToString(((dynamic)book).FullName)); result.Add(Path.GetFullPath(Path.Combine(directory, path)));
            }
            return result;
        }
        private static Dictionary<string, string> ExternalFiles(object book)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in LinkSet(book))
            {
                string name = Path.GetFileName(path), old;
                if (result.TryGetValue(name, out old) && !string.Equals(old, path, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("外链中存在同名不同路径文件，无法可靠判定来源：" + name);
                result[name] = path;
            }
            return result;
        }
        private void VerifyNamesAndLinks(object book, MigrationTarget target, string stage)
        {
            var actual = ReadNames(book); var context = Context(book, target.Sheet, stage);
            foreach (var expected in target.ExistingNames.Concat(target.Names))
            {
                var name = actual.FirstOrDefault(n => string.Equals(n.Key, expected.Key, StringComparison.OrdinalIgnoreCase));
                if (name == null || name.Visible != expected.Visible || !SameDefinition(name.Formula, expected.Formula, context)) throw new IOException("保存重开后名称核对失败：" + expected.Name);
            }
            var expectedKeys = new HashSet<string>(target.ExistingNames.Concat(target.Names).Select(n => n.Key), StringComparer.OrdinalIgnoreCase);
            if (actual.Any(n => !expectedKeys.Contains(n.Key))) throw new IOException("副本产生了未计划的名称。");
            var allowed = new HashSet<string>(target.ExistingLinks, StringComparer.OrdinalIgnoreCase); allowed.UnionWith(sourceLinks.Values);
            foreach (string path in LinkSet(book)) if (!allowed.Contains(path)) throw new IOException("副本产生了未计划的外部链接：" + path);
        }
        private static void Fail(MigrationTarget target, Exception error)
        { string phase = target.Status; target.Ready = false; target.Status = error is OperationCanceledException ? "已取消" : "失败"; target.Error = phase + "：" + error.Message; target.FailureDetails = error.ToString(); }
        private void WriteLog()
        {
            var lines = new List<string> { "批量迁移数据；" + (request.ReplaceTargets ? "校验后更新目标原文件，替换前版本保留备份" : "另存结果副本，原文件不修改"), "源：" + request.SourcePath, "源区域：" + request.SourceSheet + "!" + request.SourceRange, "源非空数：" + request.NonEmptyCells + "；公式数：" + request.FormulaCells, "源摘要：" + sourceHash };
            foreach (var target in request.Targets) lines.Add(target.Status + "\t" + target.Path + "\t" + target.Sheet + "!" + target.Start + "\t" + target.Output + "\t备份：" + target.Backup + "\t" + target.Error);
            foreach (var target in request.Targets.Where(t => t.FailureDetails.Length > 0)) lines.Add("异常详情：" + target.Path + Environment.NewLine + target.FailureDetails);
            lines.AddRange(diagnosticPaths.Select(p => "未交付的临时副本：" + p));
            lines.Add("监听事件："+Environment.NewLine+EventTrace);
            File.WriteAllLines(Path.Combine(DiagnosticDirectory, "migration.log"), lines, new UTF8Encoding(false));
        }
        private void TryWriteLog()
        {
            try { WriteLog(); }
            catch (Exception error) { UiDiagnostics.Record("migration-log-write", error); }
        }
        private void RestoreActive()
        {
            using (var window = new MigrationWindowScope(app))
            try
            {
                using (var references = new ComBag())
                {
                    object current = references.Own(((dynamic)app).ActiveWorkbook);
                    if (!WorksheetIdentity.SameObject(current, activeBook) && activeBook != null)
                    { ((dynamic)activeBook).Activate(); if (activeSheet != null) ((dynamic)activeSheet).Activate(); }
                }
            }
            catch (Exception error) { UiDiagnostics.Record("migration-restore-active", error); }
        }
        public void Dispose()
        {
            try
            {
                if (request != null) foreach (var target in request.Targets) if (target.Gateway != null) ReleaseTarget(target);
                if (copyWatch != null) { copyWatch.Dispose(); copyWatch = null; }
                if (ownedCopy != null) { Close(ownedCopy); ownedCopy = null; }
                if (copyReferences != null) { copyReferences.Dispose(); copyReferences = null; }
                if (DiagnosticDirectory != null) TryWriteLog();
            }
            finally
            {
                try { if (watch != null) { watch.Dispose(); watch = null; } }
                finally
                {
                    try { Quiet(delegate { RestoreActive(); gateway.Dispose(); }); }
                    finally
                    {
                        foreach (string path in readCopies.Values) try { File.Delete(path); } catch (Exception error) { UiDiagnostics.Record("migration-read-copy-cleanup", error); }
                        readCopies.Clear();
                    }
                }
            }
        }
    }
}
