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
        // Formula cells that refer to sheets this target lacks: their
        // calculated value is pasted instead (the user is asked first).
        internal readonly HashSet<string> MissingSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal int ValueFallbacks;
        // Precheck: how long opening this target took (for 预计耗时); when
        // the last attempt ended (for the log).
        internal long OpenMilliseconds;
        internal DateTime Finished;
    }
    internal sealed class MigrationRequest
    {
        internal string SourcePath = "", SourceSheet = "", SourceRange = "";
        internal bool AllowEmptySource;
        internal bool ReplaceTargets;
        // 只写数值: formulas in the source are read as their calculated values
        // and written as constants; no formula translation or names involved.
        internal bool ValuesOnly;
        // 数值+格式 / 公式+格式: copy the source cell formats too. Without it
        // the target keeps its own formats (merges still follow the source).
        internal bool KeepFormat;
        // 逐行重算: before reading, copy TriggerFrom into TriggerCell in the
        // source's own read copy and recalculate; TriggerNameCell must still
        // read TriggerExpect (the name the target was matched by).
        internal string TriggerCell = "", TriggerFrom = "", TriggerNameCell = "", TriggerExpect = "";
        internal int NonEmptyCells, FormulaCells;
        // Precheck: cells read and how long reading the source took.
        internal int Cells;
        internal long ReadMilliseconds;
        internal List<MigrationTarget> Targets = new List<MigrationTarget>();
    }
    internal sealed class MigrationCell
    {
        internal bool IsFormula;
        internal object Value;
        internal string Formula;
        // The formula's calculated value, used when the target lacks a sheet it refers to.
        internal object Computed;
        internal Dictionary<string, object> Style;
    }
    internal sealed class MigrationProgress
    {
        internal string Message;
        // The next step is a long host call; show this message right away.
        internal bool Pause;
        internal MigrationProgress(string message) { Message = message; }
        internal MigrationProgress(string message, bool pause) { Message = message; Pause = pause; }
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
        // Digest of each original when its read copy was made.
        private readonly Dictionary<string, string> readDigests = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // HoldTargets: a target's read copy stays open for the target's other
        // passes in this run (opened once), until ReleaseHeld or Dispose.
        internal bool HoldTargets;
        private readonly Dictionary<string, MergeWorkbookGateway> heldTargets = new Dictionary<string, MergeWorkbookGateway>(StringComparer.OrdinalIgnoreCase);
        private SourceChangeWatch watch;
        private SourceChangeWatch copyWatch;
        private MergeWorkbookHandle source;
        private MigrationRequest request;
        private MigrationArea area;
        private readonly List<MigrationCell> cells = new List<MigrationCell>();
        private readonly List<MigrationArea> merges = new List<MigrationArea>();
        private List<MigrationName> sourceNames;
        private Dictionary<string, string> sourceLinks;
        private HashSet<string> sourceAmbiguous = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
                // Very hidden sheets (such as WPS's own WPE_CUSTOM) cannot be chosen.
                for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) using (var item = new ComBag())
                {
                    dynamic sheet = item.Own(sheets.Item(i)); int visible = -1;
                    try { visible = Convert.ToInt32(sheet.Visible); } catch (Exception) { }
                    if (visible != 2) result.Add(Convert.ToString(sheet.Name));
                }
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
                // Relationship id -> type, to keep worksheets only (chart and
                // dialog sheets, and WPS's internal WPE_CUSTOM, are not
                // reachable as worksheets).
                var types = new Dictionary<string, string>(StringComparer.Ordinal);
                var rels = package.GetEntry("xl/_rels/workbook.xml.rels");
                if (rels != null)
                    using (var xml = XmlReader.Create(rels.Open(), settings))
                        while (xml.Read())
                            if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "Relationship")
                            { string id = xml.GetAttribute("Id"), type = xml.GetAttribute("Type"); if (!string.IsNullOrEmpty(id) && type != null) types[id] = type; }
                var result = new List<string>(); var all = new List<string>();
                using (var xml = XmlReader.Create(entry.Open(), settings))
                    while (xml.Read())
                        if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "sheet" &&
                            (xml.NamespaceURI == "http://schemas.openxmlformats.org/spreadsheetml/2006/main" || xml.NamespaceURI == "http://purl.oclc.org/ooxml/spreadsheetml/main"))
                        {
                            string name = xml.GetAttribute("name");
                            if (string.IsNullOrEmpty(name) || all.Contains(name)) throw new InvalidOperationException("工作表目录包含空名称或重复名称。");
                            all.Add(name);
                            if (string.Equals(xml.GetAttribute("state"), "veryHidden", StringComparison.Ordinal)) continue;
                            string id = xml.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships") ?? xml.GetAttribute("id", "http://purl.oclc.org/ooxml/officeDocument/relationships");
                            string type;
                            if (id != null && types.TryGetValue(id, out type) && !type.EndsWith("/worksheet", StringComparison.Ordinal)) continue;
                            result.Add(name);
                        }
                if (result.Count == 0) throw new InvalidOperationException("工作簿没有工作表。");
                return result;
            }
        }
        internal static string Digest(string path)
        { using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", ""); }
        internal string PickSourceRange(string path, string sheetName, string current, Func<object, object> picker, string role = "源", Action<Action> ready = null)
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
            // ready gets a way to put the selection back: the caller clicks
            // into the sheet once for the user (so the first real click
            // already selects), which moves the selection to that cell.
            string restore = initial is string ? (string)initial : null;
            try
            {
                using (var bag = new ComBag())
                {
                    if (restore != null) { dynamic range = bag.Own(((dynamic)handle.Sheet).Range[restore]); range.Select(); }
                    else { dynamic selection = bag.Own(((dynamic)app).Selection); restore = Convert.ToString(selection.Address); }
                }
            }
            catch (Exception) { }
            if (ready != null)
                ready(delegate
                {
                    if (string.IsNullOrEmpty(restore)) return;
                    try { using (var bag = new ComBag()) { dynamic range = bag.Own(((dynamic)handle.Sheet).Range[restore]); range.Select(); } }
                    catch (Exception) { }
                });
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
        // copy: read a disposable copy even when the user has the file open
        // (逐行重算 edits its trigger cell, never in the user's workbook).
        private MergeWorkbookHandle ResolveOriginal(string path, string sheet, MergeWorkbookGateway targetGateway = null, bool interactive = false, bool copy = false)
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
            if ((!alreadyOpen || needsDateCopy || copy) && !readCopies.TryGetValue(full, out readPath))
            {
                readPath = Path.Combine(Path.GetDirectoryName(full), ".xy-migration-read-" + Guid.NewGuid().ToString("N") + Path.GetExtension(full));
                try
                {
                    string before = Digest(full);
                    File.Copy(full, readPath, false); readCopies[full] = readPath; readDigests[full] = before;
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
            EnsureTargetClosed(target, true);
        }
        // Precheck (writing = false) only needs the target saved: an open target
        // is read in place. Replacing the file needs it closed.
        private void EnsureTargetClosed(MigrationTarget target, bool writing)
        {
            using (var references = new ComBag())
            {
                dynamic books = references.Own(((dynamic)app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic book = references.Own(books.Item(i));
                    if (string.Equals(Path.GetFullPath(Convert.ToString(book.FullName)), Path.GetFullPath(target.Path), StringComparison.OrdinalIgnoreCase))
                    {
                        if (request.ReplaceTargets && writing) throw new InvalidOperationException("更新原文件前请先保存并关闭该目标工作簿：" + target.Path);
                        if (!ComFormulaHost.Boolean(book.Saved, "Saved")) throw new InvalidOperationException("目标有未保存修改，请先保存并重新预检：" + target.Path);
                    }
                }
            }
            if (!request.ReplaceTargets) return;
            if ((File.GetAttributes(target.Path) & FileAttributes.ReadOnly) != 0)
                throw new InvalidOperationException("目标文件为只读，不能更新原文件：" + target.Path);
            if (writing) using (File.Open(target.Path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
        }
        private void ReleaseTarget(MigrationTarget target)
        {
            var resolver = target.Gateway; target.Gateway = null;
            if (HoldTargets)
            {
                bool owner = target.Handle != null && target.Handle.OpenedByUs; target.Handle = null;
                if (resolver == null) return;
                string key = Path.GetFullPath(target.Path);
                if (owner && !heldTargets.ContainsKey(key)) { heldTargets[key] = resolver; return; }
                Quiet(delegate { resolver.Dispose(); }); return;
            }
            try { if (resolver != null) Quiet(delegate { resolver.Dispose(); }); }
            finally
            {
                target.Handle = null;
                string full = Path.GetFullPath(target.Path), path;
                foreach (var handle in originalPaths.Where(p => string.Equals(p.Value, full, StringComparison.OrdinalIgnoreCase)).Select(p => p.Key).ToArray()) originalPaths.Remove(handle);
                if (readCopies.TryGetValue(full, out path))
                {
                    try { File.Delete(path); readCopies.Remove(full); readDigests.Remove(full); }
                    catch (Exception error) { UiDiagnostics.Record("migration-target-read-cleanup", error); }
                }
            }
        }
        // Closes a held target read copy and deletes it.
        private void ReleaseHeld(string full)
        {
            MergeWorkbookGateway held;
            try { if (heldTargets.TryGetValue(full, out held)) { heldTargets.Remove(full); Quiet(delegate { held.Dispose(); }); } }
            finally
            {
                string path;
                foreach (var handle in originalPaths.Where(p => string.Equals(p.Value, full, StringComparison.OrdinalIgnoreCase)).Select(p => p.Key).ToArray()) originalPaths.Remove(handle);
                if (readCopies.TryGetValue(full, out path))
                {
                    try { File.Delete(path); readCopies.Remove(full); readDigests.Remove(full); }
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
            // A host may run several 逐行重算 passes on one source copy.
            if (watch != null) { watch.Dispose(); watch = null; }
            MigrationWindowScope.RequireVisibleHost(app);
            CheckFile(input.SourcePath); area = MigrationArea.Parse(input.SourceRange);
            // A bounded snapshot makes accidental whole-sheet selection actionable.
            if ((long)area.Rows * area.Columns > 100000) throw new InvalidOperationException("单批最多迁移 100000 个单元格，请缩小源区域分批处理。");
            sourceHash = Digest(input.SourcePath);
            source = ResolveOriginal(input.SourcePath, input.SourceSheet, null, false, input.TriggerCell.Length > 0); Saved(source); ValidateOriginal(source, sourceHash);
            if (input.TriggerCell.Length > 0)
            {
                // Kept passes of this source must not see the trigger edit as a change.
                foreach (var slice in kept)
                    if (slice.Watch != null && string.Equals(Path.GetFullPath(slice.Request.SourcePath), Path.GetFullPath(input.SourcePath), StringComparison.OrdinalIgnoreCase))
                    { slice.Watch.Dispose(); slice.Watch = null; }
                yield return new MigrationProgress("填入触发值并重算：" + input.TriggerFrom + " → " + input.TriggerCell + "…", true);
                var recalc = System.Diagnostics.Stopwatch.StartNew();
                ApplyTrigger(input);
                Connect.Log("迁移耗时：逐行重算 " + input.TriggerFrom + " " + recalc.ElapsedMilliseconds + " ms");
            }
            watch = Watch(source.Book,"迁移来源");
            packagePreservation = MigrationPackagePreservation.Capture(input.SourcePath, input.SourceSheet, area, input.KeepFormat);
            sourceNames = ReadNames(source.Book); sourceLinks = ExternalFiles(source.Book, out sourceAmbiguous);
            using (var bag = new ComBag())
            {
                dynamic rows = bag.Own(((dynamic)source.Sheet).Rows), columns = bag.Own(((dynamic)source.Sheet).Columns);
                area.Check(Convert.ToInt32(rows.Count), Convert.ToInt32(columns.Count));
            }
            RejectTables(source.Sheet, area);
            var reading = System.Diagnostics.Stopwatch.StartNew();
            for (int r = 0; r < area.Rows; r++) for (int c = 0; c < area.Columns; c++)
            {
                Check(); watch.Validate();
                using (var bag = new ComBag())
                {
                    dynamic cell = bag.Own(((dynamic)source.Sheet).Range[MigrationArea.Cell(area.Row + r, area.Column + c)]);
                    if (!input.ValuesOnly) RejectArray(cell);
                    ReadMerge(cell, area, merges);
                    bool follower = IsMergeFollower(cell), formula = !follower && !input.ValuesOnly && ComFormulaHost.Boolean(cell.HasFormula, "HasFormula");
                    cells.Add(new MigrationCell { IsFormula = formula, Formula = formula ? Formula(cell) : null, Value = formula || follower ? null : (object)cell.Value2, Computed = formula ? (object)cell.Value2 : null, Style = input.KeepFormat ? ReadStyle(cell) : new Dictionary<string, object>() });
                }
                yield return new MigrationProgress("读取源区域 " + cells.Count + "/" + ((long)area.Rows * area.Columns));
            }
            watch.Validate(); ValidateOriginal(source, sourceHash);
            input.Cells = cells.Count; input.ReadMilliseconds = reading.ElapsedMilliseconds;
            input.NonEmptyCells = cells.Count(c => c.IsFormula || c.Value != null);
            input.FormulaCells = cells.Count(c => c.IsFormula);
            if (input.NonEmptyCells == 0 && !input.AllowEmptySource) throw new InvalidOperationException("源区域 " + input.SourceSheet + "!" + area.Address + " 全部为空，未生成结果。请选择有数据的区域；如需清空目标，请明确勾选“允许全空白源区域”。");
            foreach (var target in input.Targets)
            {
                try
                {
                    Check(); target.Ready = false; target.Error = ""; target.Backup = ""; target.Fingerprint = ""; target.Overwritten = 0; target.Output = ""; target.Samples.Clear(); target.Heights.Clear(); target.Widths.Clear(); target.Status = "预检中";
                    yield return new MigrationProgress("正在打开目标 " + Path.GetFileName(target.Path) + " 预检（文件较大时可能需要较长时间）…", true);
                    Exception failure = null; var opened = System.Diagnostics.Stopwatch.StartNew();
                    try { PrepareTarget(target, false); } catch (Exception error) { failure = error; }
                    Connect.Log("迁移耗时：预检打开 " + Path.GetFileName(target.Path) + " " + opened.ElapsedMilliseconds + " ms" + (failure == null ? "" : "（失败）"));
                    target.OpenMilliseconds = opened.ElapsedMilliseconds;
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
                        try
                        {
                            ValidateOriginal(target.Handle, target.Fingerprint); target.Ready = true; target.Status = "可执行"; target.Finished = DateTime.Now;
                            if (target.ValueFallbacks > 0) target.Error = "缺少工作表“" + string.Join("”“", target.MissingSheets.ToArray()) + "”：" + target.ValueFallbacks + " 个公式单元格将改为粘贴数值（执行前会确认）";
                        }
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
        private void PrepareTarget(MigrationTarget target, bool writing = true)
        {
            if (string.IsNullOrWhiteSpace(target.Sheet)) throw new InvalidOperationException("请选择该目标中实际存在的工作表。");
            CheckFile(target.Path);
            EnsureTargetClosed(target, writing);
            if(packagePreservation!=null)packagePreservation.RequireCompatibleTarget(target.Path);
            if (string.Equals(Path.GetFullPath(target.Path), Path.GetFullPath(request.SourcePath), StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("目标应为其他工作簿，不能选择源文件。");
            var start = MigrationArea.Parse(target.Start); if (start.Rows != 1 || start.Columns != 1) throw new InvalidOperationException("目标起始位置必须是单个单元格。");
            target.Area = new MigrationArea { Row = start.Row, Column = start.Column, Rows = area.Rows, Columns = area.Columns };
            string fingerprint = Digest(target.Path);
            if (target.Fingerprint.Length != 0 && target.Fingerprint != fingerprint) throw new InvalidOperationException("目标在预检后已变化，请重新预检：" + target.Path);
            target.Fingerprint = fingerprint;
            target.Gateway = new MergeWorkbookGateway(app);
            target.Handle = ResolveOriginal(target.Path, target.Sheet, target.Gateway); Saved(target.Handle); ValidateOriginal(target.Handle, target.Fingerprint);
            string copied; // A held read copy must still match the file on disk.
            if (readDigests.TryGetValue(Path.GetFullPath(target.Path), out copied) && copied != target.Fingerprint) throw new InvalidOperationException("目标在预检期间已变化，请重新预检：" + target.Path);
            if (ComFormulaHost.Boolean(((dynamic)target.Handle.Sheet).ProtectContents, "ProtectContents") || ComFormulaHost.Boolean(((dynamic)target.Handle.Book).ProtectStructure, "ProtectStructure")) throw new InvalidOperationException("目标工作表或工作簿结构受保护。");
            RejectTables(target.Handle.Sheet, target.Area);
            var context = Context(target.Handle.Book, target.Sheet, request.SourcePath);
            context.SourceSheet = request.SourceSheet; context.RowDelta = start.Row - area.Row; context.ColumnDelta = start.Column - area.Column; context.Names = sourceNames; context.ExternalFiles = sourceLinks; context.AmbiguousFiles = sourceAmbiguous;
            target.Area.Check(context.MaxRow, context.MaxColumn); target.Context = context;
            target.Formulas = new List<MigrationFormula>(); target.MissingSheets.Clear(); target.ValueFallbacks = 0;
            foreach (var cell in cells)
            {
                MigrationFormula translated = null;
                if (cell.IsFormula)
                    try { translated = DataMigrationEngine.Translate(cell.Formula, context, request.SourceSheet); }
                    catch (MigrationMissingSheetException missing) { target.MissingSheets.Add(missing.Sheet); target.ValueFallbacks++; }
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
            HashSet<string> ambiguous; var files = ExternalFiles(book, out ambiguous);
            var context = new MigrationFormulaContext { SourcePath = path, SourceSheet = sheetName, TargetSheet = sheetName, Names = ReadNames(book), ExternalFiles = files, AmbiguousFiles = ambiguous };
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
                yield return new MigrationProgress("正在处理 " + Path.GetFileName(target.Path) + "：复核、复制并打开副本（文件较大时可能需要较长时间）…", true);
                var stepTimer = System.Diagnostics.Stopwatch.StartNew(); long lastMark = 0; string stepFile = Path.GetFileName(target.Path);
                Action<string> mark = delegate(string step) { Connect.Log("迁移耗时：" + stepFile + " " + step + " " + (stepTimer.ElapsedMilliseconds - lastMark) + " ms"); lastMark = stepTimer.ElapsedMilliseconds; };
                try
                {
                    target.Status = "复核源文件";
                    watch.Validate(); ValidateOriginal(source, sourceHash);
                    target.Status = "复核目标文件"; ValidateTarget(target);
                    // Rebuild the preflight immediately before copying to reject
                    // newly introduced name conflicts as well as stale disk data.
                    target.Status = "重新预检目标"; PrepareTarget(target); mark("复核并重新预检");
                    ReleaseTarget(target);
                    target.Status = "创建结果副本";
                    stage = Path.Combine(Path.GetDirectoryName(target.Path), ".xy-migration-" + Guid.NewGuid().ToString("N") + Path.GetExtension(target.Path));
                    File.Copy(target.Path, stage, false); diagnosticPaths.Add(stage);
                    if (Digest(stage) != target.Fingerprint || Digest(target.Path) != target.Fingerprint) throw new IOException("复制写入副本期间目标文件已变化，请重新预检。");
                    MigrationPackageCompatibility.NormalizeOwnedCopy(stage);
                    mark("复制副本"); target.Status = "打开结果副本"; book = ownedCopy = held.Own(OpenCopy(stage)); mark("打开副本");
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
                                if (cells[index].Style.Count > 0) WriteStyle(cell, cells[index].Style);
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
                        copyWatch.Dispose(); copyWatch = null; mark("写入并保存");
                        target.Status = "重开核验";
                        Close(book); book = ownedCopy = null; held.Dispose(); held = new ComBag(); copyReferences = held;
                        if (packagePreservation != null) packagePreservation.Apply(stage, target);
                        book = ownedCopy = held.Own(OpenCopy(stage)); sheet = held.Own(Sheet(book, target.Sheet));
                        copyWatch = Watch(book,"重开副本 "+Path.GetFileName(target.Path)); RestoreActive(); mark("关闭并重开");
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
                        diagnosticPaths.Remove(stage); target.Ready = false; target.Status = "成功"; target.Finished = DateTime.Now;
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
        // ---------------- one save per target ----------------
        // A whole plan in one run: every pass is previewed first and kept,
        // then each target workbook is copied, opened, written by all of its
        // passes, saved, reopened, verified and committed exactly once (one
        // backup). Sources are only read: their copies stay open until the
        // host is disposed; a source the user has open is read in place.
        private sealed class SourceSlice
        {
            internal MigrationRequest Request;
            internal MigrationArea Area;
            internal List<MigrationCell> Cells;
            internal List<MigrationArea> Merges;
            internal MergeWorkbookHandle Source;
            internal string Hash;
            internal SourceChangeWatch Watch;
            internal MigrationPackagePreservation Preservation;
            internal List<MigrationName> Names;
            internal Dictionary<string, string> Links;
            internal HashSet<string> Ambiguous;
        }
        private sealed class BatchEntry
        {
            internal SourceSlice Slice;
            internal MigrationTarget Target;
            // A later entry in the same sheet covers part of this area (a
            // 逐行重算 override the user confirmed): those cells are not verified.
            internal readonly List<MigrationArea> CoveredBy = new List<MigrationArea>();
        }
        private sealed class BatchOutcome
        {
            internal Exception Failure;
            internal BatchEntry Culprit;
        }
        private readonly List<SourceSlice> kept = new List<SourceSlice>();
        // Call right after a successful Preview: keeps what it read for ExecuteKept.
        internal void Keep()
        {
            if (request == null || source == null) return;
            var slice = new SourceSlice
            {
                Request = request, Area = area, Cells = new List<MigrationCell>(cells), Merges = new List<MigrationArea>(merges),
                Source = source, Hash = sourceHash, Watch = watch, Preservation = packagePreservation,
                Names = sourceNames, Links = sourceLinks, Ambiguous = sourceAmbiguous
            };
            watch = null;
            // A 逐行重算 copy changes on purpose between rows; it is guarded by
            // the original's hash instead of change events.
            if (request.TriggerCell.Length > 0 && slice.Watch != null) { slice.Watch.Dispose(); slice.Watch = null; }
            kept.Add(slice);
        }
        internal int KeptTargets { get { return kept.Sum(s => s.Request.Targets.Count(t => t.Ready)); } }
        private void Use(SourceSlice slice)
        {
            request = slice.Request; area = slice.Area; source = slice.Source; sourceHash = slice.Hash; packagePreservation = slice.Preservation;
            sourceNames = slice.Names; sourceLinks = slice.Links; sourceAmbiguous = slice.Ambiguous;
            cells.Clear(); cells.AddRange(slice.Cells); merges.Clear(); merges.AddRange(slice.Merges);
        }
        private void ValidateSlice(SourceSlice slice)
        {
            if (slice.Watch != null) slice.Watch.Validate();
            ValidateOriginal(slice.Source, slice.Hash);
        }
        private void CloseBatch(object book)
        {
            Quiet(delegate { ((dynamic)book).Close(false); });
            foreach (var slice in kept) if (slice.Watch != null) slice.Watch.ProcessPending();
        }
        internal IEnumerable<MigrationProgress> ExecuteKept()
        {
            Check();
            if (watch != null) { watch.Dispose(); watch = null; }
            DiagnosticDirectory = Path.Combine(DiagnosticRootOverride ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "DataMigration"), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(DiagnosticDirectory);
            var groups = kept.SelectMany(s => s.Request.Targets.Where(t => t.Ready).Select(t => new BatchEntry { Slice = s, Target = t }))
                .GroupBy(e => Path.GetFullPath(e.Target.Path), StringComparer.OrdinalIgnoreCase).Select(g => g.ToList()).ToList();
            // Remaining time from the pace so far (cells written and verified).
            double totalCells = groups.Sum(x => x.Sum(e => (double)e.Slice.Cells.Count)) + groups.Count * 50.0, doneCells = 0;
            var runClock = System.Diagnostics.Stopwatch.StartNew();
            for (int g = 0; g < groups.Count && !Cancelled; g++)
            {
                var group = groups[g];
                string label = "[目标 " + (g + 1) + "/" + groups.Count + (doneCells > 0 ? "，剩余约 " + DurationText(runClock.ElapsedMilliseconds / doneCells * (totalCells - doneCells)) : "") + "] ";
                foreach (var entry in group) entry.Target.Status = "等待写入";
                while (group.Count > 0)
                {
                    var outcome = new BatchOutcome();
                    foreach (var step in WriteGroup(group, outcome)) yield return new MigrationProgress(label + step.Message, step.Pause);
                    if (outcome.Failure == null) break;
                    // A failure that belongs to one source range: drop that range
                    // and write the others again, so they are still saved.
                    if (outcome.Culprit != null && group.Count > 1 && !(outcome.Failure is OperationCanceledException))
                    {
                        Fail(outcome.Culprit.Target, outcome.Failure); outcome.Culprit.Target.Formulas = null;
                        group.Remove(outcome.Culprit);
                        yield return new MigrationProgress(label + "已跳过失败的一处写入（" + outcome.Culprit.Target.Error + "），其余 " + group.Count + " 处重新写入…");
                        continue;
                    }
                    foreach (var entry in group) { Fail(entry.Target, outcome.Failure); entry.Target.Formulas = null; }
                    if (outcome.Failure is OperationCanceledException) Cancelled = true;
                    break;
                }
                doneCells += groups[g].Sum(e => (double)e.Slice.Cells.Count) + 50.0;
            }
            if (Cancelled)
                foreach (var target in kept.SelectMany(s => s.Request.Targets))
                    if (target.Ready) { target.Ready = false; target.Status = "已取消"; target.Error = "已取消，未执行"; }
            TryWriteLog(); RestoreActive();
        }
        // Writes every entry of one target workbook into one copy. Sets
        // outcome.Failure (and Culprit when a single entry caused it).
        private IEnumerable<MigrationProgress> WriteGroup(List<BatchEntry> group, BatchOutcome outcome)
        {
            var first = group[0].Target; string file = Path.GetFileName(first.Path);
            string stage = null; object book = null; ComBag held = new ComBag(); copyReferences = held;
            Exception failure = null; BatchEntry culprit = null;
            foreach (var entry in group) entry.CoveredBy.Clear();
            var sheets = new Dictionary<BatchEntry, object>();
            var stepTimer = System.Diagnostics.Stopwatch.StartNew(); long lastMark = 0;
            Action<string> mark = delegate(string step) { Connect.Log("迁移耗时：" + file + "（" + group.Count + " 处写入）" + step + " " + (stepTimer.ElapsedMilliseconds - lastMark) + " ms"); lastMark = stepTimer.ElapsedMilliseconds; };
            yield return new MigrationProgress("正在处理 " + file + "（" + group.Count + " 处写入）：复核、复制并打开副本（文件较大时可能需要较长时间）…", true);
            foreach (var entry in group)
            {
                if (failure != null) break;
                var target = entry.Target;
                try
                {
                    Use(entry.Slice);
                    target.Status = "复核源文件"; ValidateSlice(entry.Slice);
                    target.Status = "复核目标文件"; ValidateTarget(target);
                    target.Status = "重新预检目标"; PrepareTarget(target);
                }
                catch (Exception error) { failure = error; culprit = entry; }
                try { ReleaseTarget(target); } catch (Exception error) { if (failure == null) { failure = error; culprit = entry; } }
                if (failure == null) yield return new MigrationProgress("复核 " + file + " " + (group.IndexOf(entry) + 1) + "/" + group.Count);
            }
            try { ReleaseHeld(Path.GetFullPath(first.Path)); } catch (Exception error) { failure = failure ?? error; }
            if (failure == null)
            {
                mark("复核并重新预检");
                // Later entries win where areas overlap; mark what they cover.
                for (int i = 0; i < group.Count; i++)
                    for (int j = i + 1; j < group.Count; j++)
                        if (string.Equals(group[i].Target.Sheet, group[j].Target.Sheet, StringComparison.OrdinalIgnoreCase) && Overlaps(group[i].Target.Area, group[j].Target.Area))
                            group[i].CoveredBy.Add(group[j].Target.Area);
                try
                {
                    first.Status = "创建结果副本";
                    stage = Path.Combine(Path.GetDirectoryName(first.Path), ".xy-migration-" + Guid.NewGuid().ToString("N") + Path.GetExtension(first.Path));
                    File.Copy(first.Path, stage, false); diagnosticPaths.Add(stage);
                    if (Digest(stage) != first.Fingerprint || Digest(first.Path) != first.Fingerprint) throw new IOException("复制写入副本期间目标文件已变化，请重新预检。");
                    MigrationPackageCompatibility.NormalizeOwnedCopy(stage);
                    mark("复制副本"); book = ownedCopy = held.Own(OpenCopy(stage)); mark("打开副本");
                    copyWatch = Watch(book, "写入副本 " + file); RestoreActive();
                }
                catch (Exception error) { failure = error; }
            }
            // Each entry in turn: unmerge, names, cells, merges, dimensions.
            foreach (var entry in group)
            {
                if (failure != null) break;
                var target = entry.Target; object sheet = null;
                try
                {
                    Use(entry.Slice); target.Status = "定位目标工作表";
                    sheet = held.Own(Sheet(book, target.Sheet)); sheets[entry] = sheet;
                    target.ActualMerges.Clear();
                    Quiet(delegate
                    {
                        using (var bag = new ComBag()) { dynamic range = bag.Own(((dynamic)sheet).Range[target.Area.Address]); range.UnMerge(); }
                        foreach (var name in target.Names) AddName(book, name);
                    });
                    target.Status = "写入副本";
                }
                catch (Exception error) { failure = error; culprit = entry; break; }
                int count = cells.Count;
                for (int i = 0; i < count; )
                {
                    try
                    {
                        Check(); if (entry.Slice.Watch != null) entry.Slice.Watch.Validate(); copyWatch.Validate(); int start = i;
                        i = Quiet<int>(delegate
                        {
                            var clock = System.Diagnostics.Stopwatch.StartNew(); int index = start;
                            do { WriteCell(sheet, target, index); index++; }
                            while (index < count && !Cancelled && clock.ElapsedMilliseconds < WriteSliceMilliseconds);
                            return index;
                        });
                    }
                    catch (Exception error) { failure = error; culprit = entry; break; }
                    yield return new MigrationProgress("写入 " + file + " 第 " + (group.IndexOf(entry) + 1) + "/" + group.Count + " 处 " + i + "/" + count);
                }
                if (failure != null) break;
                try
                {
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
                            if (cells[index].Style.Count > 0) WriteStyle(cell, cells[index].Style);
                        }
                        RestoreDimensions(sheet, target);
                    });
                }
                catch (Exception error) { failure = error; culprit = entry; }
            }
            if (failure == null)
            {
                try
                {
                    Check(); copyWatch.Validate(); first.Status = "保存副本";
                    object view = sheets[group[0]];
                    Quiet(delegate
                    {
                        // Save with a visible workbook view so the delivered
                        // file does not reopen hidden. Never yield while it
                        // is visible, and hide again even when Save fails.
                        MigrationWindowScope.SetOwnedVisible(book, true);
                        try
                        {
                            SetOutputView(book, view, first); ((dynamic)book).Save();
                            if (!ComFormulaHost.Boolean(((dynamic)book).Saved, "Saved")) throw new IOException("副本保存未完成。");
                        }
                        finally { MigrationWindowScope.SetOwnedVisible(book, false); RestoreActive(); }
                    });
                    copyWatch.Dispose(); copyWatch = null; mark("写入并保存（一次）");
                    CloseBatch(book); book = ownedCopy = null; sheets.Clear(); held.Dispose(); held = new ComBag(); copyReferences = held;
                }
                catch (Exception error) { failure = error; }
            }
            foreach (var entry in group)
            {
                if (failure != null || entry.Slice.Preservation == null) continue;
                try { entry.Target.Status = "重开核验"; entry.Slice.Preservation.Apply(stage, entry.Target); }
                catch (Exception error) { failure = error; culprit = entry; }
            }
            if (failure == null)
            {
                try
                {
                    book = ownedCopy = held.Own(OpenCopy(stage));
                    foreach (var entry in group) sheets[entry] = held.Own(Sheet(book, entry.Target.Sheet));
                    copyWatch = Watch(book, "重开副本 " + file); RestoreActive(); mark("关闭并重开");
                    VerifyGroupNamesAndLinks(book, group, stage);
                }
                catch (Exception error) { failure = error; }
            }
            foreach (var entry in group)
            {
                if (failure != null) break;
                var target = entry.Target; MigrationFormulaContext verification = null; object sheet = null;
                try { Use(entry.Slice); target.Status = "重开核验"; sheet = sheets[entry]; verification = Context(book, target.Sheet, stage); }
                catch (Exception error) { failure = error; culprit = entry; break; }
                for (int i = 0; i < cells.Count; i++)
                {
                    try
                    {
                        Check(); copyWatch.Validate();
                        int row = target.Area.Row + i / area.Columns, column = target.Area.Column + i % area.Columns;
                        if (!entry.CoveredBy.Any(a => a.Contains(new MigrationArea { Row = row, Column = column, Rows = 1, Columns = 1 }))) VerifyCell(sheet, target, i, verification);
                    }
                    catch (Exception error) { failure = error; culprit = entry; break; }
                    if (i % 64 == 63 || i == cells.Count - 1) yield return new MigrationProgress("重开核验 " + file + " 第 " + (group.IndexOf(entry) + 1) + "/" + group.Count + " 处 " + (i + 1) + "/" + cells.Count);
                }
                if (failure != null) break;
                try
                {
                    if (entry.CoveredBy.Count == 0) VerifyMergeLayout(sheet, target);
                    ValidateSlice(entry.Slice);
                }
                catch (Exception error) { failure = error; culprit = entry; }
            }
            try { if (failure == null) ValidateTarget(first); } catch (Exception error) { failure = error; }
            try { if (copyWatch != null) { copyWatch.Dispose(); copyWatch = null; } if (book != null) { CloseBatch(book); ownedCopy = null; } } catch (Exception error) { failure = failure ?? error; }
            try { sheets.Clear(); held.Dispose(); copyReferences = null; } catch (Exception error) { failure = failure ?? error; }
            foreach (var entry in group)
            {
                if (failure != null || entry.Slice.Preservation == null) continue;
                try { entry.Slice.Preservation.Verify(stage, entry.Target); }
                catch (Exception error) { failure = error; culprit = entry; }
            }
            if (failure == null)
            {
                try
                {
                    Check();
                    var macros = MacroFingerprints(stage);
                    if (first.MacroParts.Count != macros.Count || first.MacroParts.Any(p => !macros.ContainsKey(p.Key) || macros[p.Key] != p.Value)) throw new IOException("保存后 VBA 宏载荷发生变化，副本未交付。");
                    File.SetAttributes(stage, File.GetAttributes(stage) & ~FileAttributes.Hidden);
                    ValidateTarget(first);
                    string backup = "", output;
                    if (group[0].Slice.Request.ReplaceTargets)
                    {
                        first.Status = "更新目标原文件";
                        backup = MigrationOutputTransaction.Commit(stage, first.Path, first.Fingerprint, Path.GetFileName(DiagnosticDirectory));
                        output = first.Path;
                    }
                    else
                    {
                        output = DataMigrationEngine.OutputPath(first.Path, p => File.Exists(p) || Directory.Exists(p));
                        File.Move(stage, output);
                    }
                    diagnosticPaths.Remove(stage);
                    foreach (var entry in group) { var t = entry.Target; t.Backup = backup; t.Output = output; t.Ready = false; t.Status = "成功"; t.Formulas = null; t.Finished = DateTime.Now; }
                    mark("提交");
                }
                catch (Exception error) { failure = error; culprit = null; }
            }
            if (failure != null && stage != null && File.Exists(stage))
                try
                {
                    string retained = Path.Combine(DiagnosticDirectory, Path.GetFileNameWithoutExtension(first.Path) + "_未通过校验_" + Guid.NewGuid().ToString("N").Substring(0, 8) + Path.GetExtension(first.Path));
                    File.Move(stage, retained); File.SetAttributes(retained, File.GetAttributes(retained) & ~FileAttributes.Hidden);
                    diagnosticPaths.Remove(stage); diagnosticPaths.Add(retained);
                }
                catch (Exception archiveError) { UiDiagnostics.Record("migration-failure-archive", archiveError); }
            outcome.Failure = failure; outcome.Culprit = failure == null ? null : culprit;
            TryWriteLog(); RestoreActive();
            yield return new MigrationProgress(failure == null ? "成功：" + first.Output + "（" + group.Count + " 处写入，保存一次）" : "失败：" + file + "：" + failure.Message);
        }
        private static bool Overlaps(MigrationArea a, MigrationArea b)
        { return a.Row < b.Row + b.Rows && b.Row < a.Row + a.Rows && a.Column < b.Column + b.Columns && b.Column < a.Column + a.Columns; }
        private void VerifyGroupNamesAndLinks(object book, List<BatchEntry> group, string stage)
        {
            var actual = ReadNames(book); var first = group[0].Target; var context = Context(book, first.Sheet, stage);
            var expected = first.ExistingNames.Concat(group.SelectMany(e => e.Target.Names)).ToList();
            foreach (var name in expected)
            {
                var found = actual.FirstOrDefault(n => string.Equals(n.Key, name.Key, StringComparison.OrdinalIgnoreCase));
                if (found == null || found.Visible != name.Visible || !SameDefinition(found.Formula, name.Formula, context)) throw new IOException("保存重开后名称核对失败：" + name.Name);
            }
            var keys = new HashSet<string>(expected.Select(n => n.Key), StringComparer.OrdinalIgnoreCase);
            if (actual.Any(n => !keys.Contains(n.Key))) throw new IOException("副本产生了未计划的名称。");
            var allowed = new HashSet<string>(first.ExistingLinks, StringComparer.OrdinalIgnoreCase);
            foreach (var entry in group) allowed.UnionWith(entry.Slice.Links.Values);
            foreach (string path in LinkSet(book)) if (!allowed.Contains(path)) throw new IOException("副本产生了未计划的外部链接：" + path);
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
                if (sourceCell.IsFormula && target.Formulas[index] == null)
                {
                    // Refers to a sheet this target lacks: paste the value.
                    object value = sourceCell.Computed;
                    if (value == null) cell.ClearContents();
                    else WriteValue(cell, value, request.KeepFormat);
                }
                else if (sourceCell.IsFormula)
                {
                    // A text-formatted destination treats an assigned formula as
                    // text on WPS. Apply a formula-capable format before writing.
                    if (request.KeepFormat || Convert.ToString(cell.NumberFormat) == "@") cell.NumberFormat = "General";
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
                    WriteValue(cell, sourceCell.Value, request.KeepFormat);
                }
                if (sourceCell.Style.Count > 0) WriteStyle(cell, sourceCell.Style);
            }
        }
        // The target's existing format must not reinterpret the value (a
        // number written into a text cell becomes text). With formats, write
        // under a neutral format and let WriteStyle restore the source's.
        // Without, keep the target's number format unless it would change
        // the value: text format for numbers, or text WPS would convert.
        private static void WriteValue(dynamic cell, object value, bool keepFormat)
        {
            if (keepFormat) { cell.NumberFormat = value is string ? "@" : "General"; cell.Value2 = value; return; }
            if (!(value is string) && Convert.ToString(cell.NumberFormat) == "@") cell.NumberFormat = "General";
            cell.Value2 = value;
            if (value is string && !SameValue((object)cell.Value2, value)) { cell.NumberFormat = "@"; cell.Value2 = value; }
        }
        private void VerifyCell(object sheet, MigrationTarget target, int index, MigrationFormulaContext context)
        {
            MigrationCell expected = cells[index]; string address = MigrationArea.Cell(target.Area.Row + index / area.Columns, target.Area.Column + index % area.Columns);
            using (var bag = new ComBag())
            {
                dynamic cell = bag.Own(((dynamic)sheet).Range[address]);
                bool follower = IsMergeFollower(cell), has = !follower && ComFormulaHost.Boolean(cell.HasFormula, "HasFormula");
                bool asFormula = expected.IsFormula && target.Formulas[index] != null;
                object value = expected.IsFormula ? expected.Computed : expected.Value;
                if (has != asFormula || (has ? !SameDefinition(Formula(cell), target.Formulas[index].Text, context) : !SameValue(follower ? null : (object)cell.Value2, value))) throw new IOException(address + " 保存重开后内容核对失败。");
                if (expected.Style.Count > 0)
                {
                    var actual = ReadStyle(cell);
                    foreach (var pair in expected.Style) if (!SameValue(actual[pair.Key], pair.Value)) throw new IOException(address + " 样式核对失败：" + pair.Key);
                }
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
        // Only ever on our own read copy: the user's original is never edited.
        // Trigger cells changed in a source copy, with their original content,
        // so another 逐行重算 set of the same copy starts from the original.
        private sealed class TriggerEdit { internal string Copy, Sheet, Cell, Formula; }
        private readonly List<TriggerEdit> triggerEdits = new List<TriggerEdit>();
        private void ApplyTrigger(MigrationRequest input)
        {
            string original, copy;
            bool owned = originalPaths.TryGetValue(source, out original) && readCopies.TryGetValue(original, out copy) && string.Equals(copy, source.Path, StringComparison.OrdinalIgnoreCase);
            if (!owned) throw new InvalidOperationException("逐行重算只在源的副本中改动触发单元格，未能创建副本：" + input.SourcePath);
            Quiet(delegate
            {
                using (var bag = new ComBag())
                {
                    dynamic sheet = source.Sheet;
                    if (input.TriggerNameCell.Length > 0)
                    {
                        dynamic nameCell = bag.Own(sheet.Range[input.TriggerNameCell]);
                        string actual = (Convert.ToString(nameCell.Value2) ?? "").Trim();
                        if (!string.Equals(actual, input.TriggerExpect.Trim(), StringComparison.Ordinal))
                            throw new InvalidOperationException("触发值所在行的文件名已变化（" + input.TriggerNameCell + "：" + actual + "），请重新预检。");
                    }
                    dynamic from = bag.Own(sheet.Range[input.TriggerFrom]);
                    object value = from.Value2;
                    if (value == null || Convert.ToString(value).Trim().Length == 0) throw new InvalidOperationException("触发值为空：" + input.TriggerFrom);
                    foreach (var edit in triggerEdits.Where(e => string.Equals(e.Copy, source.Path, StringComparison.OrdinalIgnoreCase)
                        && !(string.Equals(e.Sheet, input.SourceSheet, StringComparison.OrdinalIgnoreCase) && string.Equals(e.Cell, input.TriggerCell, StringComparison.OrdinalIgnoreCase))).ToList())
                    {
                        dynamic sheets = bag.Own(((dynamic)source.Book).Worksheets); dynamic other = bag.Own(sheets.Item(edit.Sheet));
                        dynamic restore = bag.Own(other.Range[edit.Cell]); restore.Formula = edit.Formula;
                        triggerEdits.Remove(edit);
                    }
                    dynamic to = bag.Own(sheet.Range[input.TriggerCell]);
                    if (!triggerEdits.Any(e => string.Equals(e.Copy, source.Path, StringComparison.OrdinalIgnoreCase) && string.Equals(e.Sheet, input.SourceSheet, StringComparison.OrdinalIgnoreCase) && string.Equals(e.Cell, input.TriggerCell, StringComparison.OrdinalIgnoreCase)))
                        triggerEdits.Add(new TriggerEdit { Copy = source.Path, Sheet = input.SourceSheet, Cell = input.TriggerCell, Formula = Convert.ToString(to.Formula) });
                    to.Value2 = value;
                    ((dynamic)app).Calculate();
                }
            });
            WaitCalculated();
        }
        // Calculate() is normally synchronous; when the host reports a
        // pending or running calculation, wait for it to finish.
        private void WaitCalculated()
        {
            var clock = System.Diagnostics.Stopwatch.StartNew();
            while (true)
            {
                int state;
                try { state = Convert.ToInt32(((dynamic)app).CalculationState); }
                catch (Exception) { return; }
                if (state == 0) return; // xlDone
                if (clock.ElapsedMilliseconds > 120000) throw new TimeoutException("触发值重算超过 2 分钟仍未完成。");
                Thread.Sleep(50); Check();
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
        // File name -> full path of each external link. A name linked under
        // several paths is left out and listed in ambiguous instead; it only
        // blocks a migration whose formulas actually refer to it.
        internal static Dictionary<string, string> MapExternalFiles(IEnumerable<string> links, out HashSet<string> ambiguous)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ambiguous = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in links)
            {
                string name = Path.GetFileName(path), old;
                if (ambiguous.Contains(name)) continue;
                if (result.TryGetValue(name, out old) && !string.Equals(old, path, StringComparison.OrdinalIgnoreCase)) { result.Remove(name); ambiguous.Add(name); continue; }
                result[name] = path;
            }
            return result;
        }
        private static Dictionary<string, string> ExternalFiles(object book, out HashSet<string> ambiguous) { return MapExternalFiles(LinkSet(book), out ambiguous); }
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
        { string phase = target.Status; target.Ready = false; target.Status = error is OperationCanceledException ? "已取消" : "失败"; target.Error = phase + "：" + error.Message; target.FailureDetails = error.ToString(); target.Finished = DateTime.Now; }
        // "约 2 分 30 秒"-style text for a duration.
        // 2026-11-5-10:30
        internal static string Stamp(DateTime time) { return time.ToString("yyyy-M-d-HH:mm", CultureInfo.InvariantCulture); }
        internal static string DurationText(double milliseconds)
        {
            long seconds = Math.Max(1, (long)Math.Round(milliseconds / 1000));
            if (seconds < 60) return seconds + " 秒";
            if (seconds < 3600) return seconds / 60 + " 分" + (seconds % 60 == 0 ? "" : " " + seconds % 60 + " 秒");
            return seconds / 3600 + " 小时" + (seconds % 3600 / 60 == 0 ? "" : " " + seconds % 3600 / 60 + " 分");
        }
        // Rough length of executing what was just prechecked: the precheck
        // again, then per target: copy, open, save and reopen (about three
        // times its open), and per cell: write, format and verify (about
        // three times its read).
        internal static double EstimateMilliseconds(IEnumerable<MigrationRequest> requests, double precheck)
        {
            var list = requests.ToList();
            long read = list.Sum(r => r.ReadMilliseconds); long readCells = list.Sum(r => (long)r.Cells);
            double perCell = readCells == 0 ? 2 : Math.Max(0.2, (double)read / readCells);
            var ready = list.SelectMany(r => r.Targets.Where(t => t.Ready).Select(t => new { Request = r, Target = t })).ToList();
            double opens = ready.GroupBy(x => Path.GetFullPath(x.Target.Path), StringComparer.OrdinalIgnoreCase).Sum(g => g.Max(x => x.Target.OpenMilliseconds)) * 3.0;
            double cells = ready.Sum(x => (double)x.Request.Cells) * perCell * 3.0;
            return precheck + opens + cells;
        }
        private void WriteLog()
        {
            var lines = new List<string> { "记录时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), "批量迁移数据；" + (request.ReplaceTargets ? "校验后更新目标原文件，替换前版本保留备份" : "另存结果副本，原文件不修改") + (kept.Count > 0 ? "；每个目标只打开、保存一次" : "") };
            var runs = kept.Count > 0 ? kept.Select(k => new KeyValuePair<MigrationRequest, string>(k.Request, k.Hash)).ToList() : new List<KeyValuePair<MigrationRequest, string>> { new KeyValuePair<MigrationRequest, string>(request, sourceHash) };
            foreach (var run in runs)
            {
                var item = run.Key;
                lines.Add("源：" + item.SourcePath); lines.Add("源区域：" + item.SourceSheet + "!" + item.SourceRange + (item.TriggerCell.Length > 0 ? "（逐行重算 " + item.TriggerFrom + " → " + item.TriggerCell + "）" : ""));
                lines.Add("源非空数：" + item.NonEmptyCells + "；公式数：" + item.FormulaCells); lines.Add("源摘要：" + run.Value);
                foreach (var target in item.Targets) lines.Add((target.Finished == DateTime.MinValue ? "--:--:--" : target.Finished.ToString("yyyy-M-d-HH:mm:ss", CultureInfo.InvariantCulture)) + "\t" + target.Status + "\t" + target.Path + "\t" + target.Sheet + "!" + target.Start + "\t" + target.Output + "\t备份：" + target.Backup + "\t" + target.Error);
            }
            foreach (var target in runs.SelectMany(r => r.Key.Targets).Where(t => t.FailureDetails.Length > 0)) lines.Add("异常详情：" + target.Path + Environment.NewLine + target.FailureDetails);
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
                foreach (var slice in kept) foreach (var target in slice.Request.Targets) if (target.Gateway != null) ReleaseTarget(target);
                foreach (string key in heldTargets.Keys.ToList()) ReleaseHeld(key);
                if (copyWatch != null) { copyWatch.Dispose(); copyWatch = null; }
                if (ownedCopy != null) { Close(ownedCopy); ownedCopy = null; }
                if (copyReferences != null) { copyReferences.Dispose(); copyReferences = null; }
                if (DiagnosticDirectory != null) TryWriteLog();
            }
            finally
            {
                try
                {
                    if (watch != null) { watch.Dispose(); watch = null; }
                    foreach (var slice in kept) if (slice.Watch != null) { try { slice.Watch.Dispose(); } catch (Exception error) { UiDiagnostics.Record("migration-kept-watch", error); } slice.Watch = null; }
                }
                finally
                {
                    try { Quiet(delegate { RestoreActive(); gateway.Dispose(); }); }
                    finally
                    {
                        foreach (string path in readCopies.Values) try { File.Delete(path); } catch (Exception error) { UiDiagnostics.Record("migration-read-copy-cleanup", error); }
                        readCopies.Clear(); readDigests.Clear();
                    }
                }
            }
        }
    }
}
