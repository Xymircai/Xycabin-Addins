using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static DataMigrationForm _migrationForm;
        private static DataMigrationHost _migrationHost;
        // The previewed plan and a stop flag for its passes.
        private static MigrationPlan _migrationPlan;
        private static bool _planCancelled;
        public void OnDataMigration(object control)
        {
            if (_busy) return;
            if (_migrationForm != null && !_migrationForm.IsDisposed) { _migrationForm.Activate(); return; }
            try
            {
                if (!EnsureApplication()) throw new InvalidOperationException("请先打开 WPS 表格。");
                var form = new DataMigrationForm(); _migrationForm = form;
                form.PreviewRequested += delegate
                {
                    MigrationAction(form, delegate
                    {
                        ResetMigration(); var plan = form.BuildPlan();
                        if (!ConfirmOverrides(form, plan)) return;
                        if (!ReleaseBusyWorkbooks(form, plan, false)) return;
                        _migrationPlan = plan; _planCancelled = false;
                        form.Run(RunPlan(form, plan, false), delegate { form.RefreshResults(); }, delegate(Exception error) { form.Message(error.Message); ResetMigration(); });
                    });
                };
                form.ExecuteRequested += delegate
                {
                    MigrationAction(form, delegate
                    {
                        if (_migrationPlan == null) throw new InvalidOperationException("请先重新预检。");
                        if (!ReleaseBusyWorkbooks(form, _migrationPlan, true)) return;
                        ExecutePlan(form);
                    });
                };
                form.CancelRequested += delegate { _planCancelled = true; if (_migrationHost != null) _migrationHost.Cancelled = true; };
                form.InputsChanged += ResetMigration;
                // Sheet names come from the package for .xlsx/.xlsm (no WPS open);
                // other formats use a hidden read-only copy, as before.
                form.BooksPicked += delegate(bool target, string[] paths)
                {
                    foreach (string path in paths)
                    {
                        try
                        {
                            using (var host = new DataMigrationHost(_app))
                            { var names = host.Sheets(path); form.AddBook(target, path, names); }
                        }
                        catch (Exception error) { Log("复制工作簿读取失败：" + path + " " + error); form.AddBook(target, path, new string[0], error.Message); }
                    }
                };
                // Pick button in workbook mode: open the workbook read-only (or use
                // it if already open), select with the mouse, then come back.
                form.RangePickRequested += delegate(bool target, string path, string sheet, string current, Action<string> done)
                {
                    if (form.Running) return;
                    string address = null; Exception failure = null;
                    form.Hide();
                    // Give WPS the focus once the workbook is shown, and again
                    // after the selection box opens; otherwise the first click
                    // only activates its window and the user has to click again.
                    IntPtr wps = GetHostWindowHandle();
                    MigrationForeground.Bring(wps);
                    // Once the selection box is up, click into the sheet for the
                    // user so the grid has the focus, then restore the selection.
                    System.Windows.Forms.Timer focusTimer = null;
                    // The frame showing the opened workbook (WPS may open it in
                    // another top-level window than the one we started from).
                    Action<Action> ready = delegate(Action reselect) { IntPtr frame = MigrationForeground.ActiveFrame(_app, wps); MigrationForeground.Bring(frame); focusTimer = MigrationForeground.BringSoon(frame, reselect); };
                    try { using (var host = new DataMigrationHost(_app)) address = host.PickSourceRange(path, sheet, current, null, target ? "目标" : "源", ready); }
                    catch (Exception error) { failure = error; Log("复制框选区域：" + error); }
                    finally { if (focusTimer != null) { focusTimer.Stop(); focusTimer.Dispose(); } form.Show(); form.Activate(); }
                    if (failure != null) form.Message("框选失败：" + failure.Message);
                    else if (address == null) form.Message("已取消框选，区域保持不变。");
                    else done(address);
                };
                form.FormClosed += delegate { ResetMigration(); _migrationForm = null; form.Dispose(); };
                DefaultMigrationSource(form);
                IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception error) { Log("批量复制数据：" + error); UiDialog.Show(error.Message, "批量复制数据", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        // The workbook the user is looking at becomes the first source, with
        // its active sheet. Names come from the open workbook itself, so
        // nothing is opened or read from disk; anything unusual (new unsaved
        // book, chart sheet, unsupported format) leaves the window empty.
        private static void DefaultMigrationSource(DataMigrationForm form)
        {
            try
            {
                using (var bag = new ComBag())
                {
                    dynamic book = bag.Own(((dynamic)_app).ActiveWorkbook);
                    if (book == null) return;
                    string path = Convert.ToString(book.FullName);
                    if (!Path.IsPathRooted(path) || !File.Exists(path) || !DataMigrationEngine.SupportedFile(path)) return;
                    var names = new List<string>();
                    dynamic sheets = bag.Own(book.Worksheets);
                    for (int i = 1; i <= Convert.ToInt32(sheets.Count); i++) using (var item = new ComBag())
                    { dynamic sheet = item.Own(sheets.Item(i)); names.Add(Convert.ToString(sheet.Name)); }
                    if (names.Count == 0) return;
                    string active = "";
                    try { dynamic sheet = bag.Own(((dynamic)_app).ActiveSheet); if (sheet != null) active = Convert.ToString(sheet.Name); } catch { }
                    if (!names.Contains(active)) active = names[0];
                    form.AddDefaultSource(Path.GetFullPath(path), names, active);
                    form.Message("已把当前工作簿“" + Path.GetFileName(path) + "”（工作表“" + active + "”）作为第 1 个源区域；点这一行末的框选按钮选择源区域，再到“目标区域”页添加目标。"
                        + (ComFormulaHost.Boolean(book.Saved, "Saved") ? "" : "当前工作簿有未保存修改，预检前请先保存。"));
                }
            }
            catch (Exception error) { Log("批量复制数据默认源：" + error); }
        }
        // Runs every pass of a plan through the existing single-source engine.
        // A pass that fails only marks its own mappings; the other passes go
        // on, which is what gives each target its partial save.
        private static IEnumerable<MigrationProgress> RunPlan(DataMigrationForm form, MigrationPlan plan, bool execute)
        {
            // Passes share one host, so each source copy is opened once (for
            // 逐行重算 only its trigger cell changes between rows).
            DataMigrationHost host = null; string hostSource = null;
            // Execution opens and saves each target workbook once, after every
            // source has been read: all passes are prechecked first, then
            // written per target. Only when a target is also a source of the
            // plan does the order matter, and then passes run one by one.
            var sourcePaths = new HashSet<string>(plan.Passes.Select(p => Path.GetFullPath(p.Request.SourcePath)), StringComparer.OrdinalIgnoreCase);
            bool chain = plan.Passes.SelectMany(p => p.Request.Targets).Any(t => sourcePaths.Contains(Path.GetFullPath(t.Path)));
            bool batch = execute && !chain;
            // Plain reads first: 逐行重算 edits the trigger cell of the
            // source copy, which ordinary ranges of that source share.
            var passes = plan.Passes.OrderBy(p => p.Request.TriggerCell.Length > 0 ? 1 : 0).ToList();
            var clock = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                for (int index = 0; index < passes.Count; index++)
                {
                    var pass = passes[index];
                    string label = "[" + (index + 1) + "/" + passes.Count + "] " + pass.Source.Label + " " + pass.Request.SourceSheet + "!" + pass.Request.SourceRange + "：";
                    if (_planCancelled) { MarkPass(pass, "已取消", "已取消，未执行"); continue; }
                    // One host for the whole run: every source is opened (as its
                    // read copy) once and closed at the end, never saved. A
                    // source the user has open in WPS is read in place and
                    // left open. A failed pass starts a fresh host.
                    bool reuse = host != null && (hostSource != null || batch);
                    if (!reuse)
                    {
                        if (host != null) { _migrationHost = null; try { host.Dispose(); } catch (Exception error) { Log("复制会话释放：" + error); } }
                        host = new DataMigrationHost(_app) { HoldTargets = !chain }; hostSource = pass.Request.SourcePath;
                    }
                    _migrationHost = host;
                    Exception failure;
                    MigrationProgress step;
                    // Execution re-checks each target right before writing.
                    var preview = host.Preview(pass.Request).GetEnumerator();
                    try { while (Advance(preview, out step, out failure)) yield return new MigrationProgress(label + step.Message, step.Pause); }
                    finally { preview.Dispose(); }
                    if (failure != null)
                    {
                        MarkPass(pass, failure is OperationCanceledException ? "已取消" : "失败", "源区域预检：" + failure.Message);
                        // A failed pass may leave the host half way; start clean
                        // (a batched run keeps its host: it holds the passes read so far).
                        if (!batch) hostSource = null;
                        continue;
                    }
                    if (!execute || !pass.Request.Targets.Any(t => t.Ready)) continue;
                    // Batched: keep what was read; every target is written once below.
                    if (batch) { host.Keep(); continue; }
                    var run = host.Execute().GetEnumerator();
                    try { while (Advance(run, out step, out failure)) yield return new MigrationProgress(label + step.Message, step.Pause); }
                    finally { run.Dispose(); }
                    if (failure != null) { MarkPass(pass, failure is OperationCanceledException ? "已取消" : "失败", "执行：" + failure.Message); hostSource = null; }
                    if (!plan.OutputCopies) foreach (var target in pass.Request.Targets) if (target.Status == "成功") form.RecordRestorePoint(target.Path, target.Backup);
                    if (host.Cancelled) _planCancelled = true;
                }
                if (batch && host != null && host.KeptTargets > 0)
                {
                    _migrationHost = host;
                    if (_planCancelled) host.Cancelled = true;
                    Exception failure; MigrationProgress step;
                    var run = host.ExecuteKept().GetEnumerator();
                    try { while (Advance(run, out step, out failure)) yield return step; }
                    finally { run.Dispose(); }
                    if (failure != null) foreach (var pass in plan.Passes) MarkPass(pass, failure is OperationCanceledException ? "已取消" : "失败", "执行：" + failure.Message);
                    if (!plan.OutputCopies) foreach (var target in plan.Passes.SelectMany(p => p.Request.Targets)) if (target.Status == "成功") form.RecordRestorePoint(target.Path, target.Backup);
                    if (host.Cancelled) _planCancelled = true;
                }
                if (!execute) plan.PrecheckMilliseconds = clock.ElapsedMilliseconds;
                plan.Finished = DateTime.Now; plan.Executed = execute; plan.RunMilliseconds = clock.ElapsedMilliseconds;
                Log("复制耗时：" + (execute ? "执行" : "预检") + "合计 " + clock.ElapsedMilliseconds + " ms");
            }
            finally
            {
                _migrationHost = null;
                if (host != null) try { host.Dispose(); } catch (Exception error) { Log("复制会话释放：" + error); }
            }
        }
        private static bool Advance(IEnumerator<MigrationProgress> steps, out MigrationProgress step, out Exception failure)
        {
            step = null; failure = null;
            try { if (!steps.MoveNext()) return false; step = steps.Current; return true; }
            catch (Exception error) { failure = error; return false; }
        }
        private static void MarkPass(MigrationPass pass, string status, string error)
        {
            foreach (var target in pass.Request.Targets)
                if (target.Status != "成功" && target.Status != "失败")
                { target.Status = status; target.Ready = false; target.Error = error; }
        }
        private static void ExecutePlan(DataMigrationForm form)
        {
            var previewed = _migrationPlan;
            var ready = previewed.Mappings.Where(m => m.Result != null && m.Result.Ready && m.Result.Status == "可执行").ToList();
            if (ready.Count == 0) throw new InvalidOperationException("没有通过预检的写入。");
            int overwrite = ready.Sum(m => m.Result.Overwritten), filled = ready.Count(m => m.Result.Overwritten > 0);
            int targets = ready.Select(m => m.TargetPath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            var empty = previewed.Passes.Where(p => p.Request.NonEmptyCells == 0 && p.Request.Targets.Any(t => t.Ready)).Select(p => p.Source.Label + " " + p.Request.SourceSheet + "!" + p.Request.SourceRange).Distinct().ToList();
            string line = Environment.NewLine + Environment.NewLine;
            string saving = previewed.ReplaceTargets
                ? "将更新 " + targets + " 个目标原文件（共 " + ready.Count + " 处写入）；原版本自动备份到同目录的“复制备份”文件夹，可用“还原操作”恢复。源文件不修改。"
                : "将为 " + targets + " 个目标各另存一个“复制结果”副本（共 " + ready.Count + " 处写入），目标原文件不修改。";
            double estimate = DataMigrationHost.EstimateMilliseconds(previewed.Passes.Select(p => p.Request), previewed.PrecheckMilliseconds);
            saving += line + "预计耗时：约 " + DataMigrationHost.DurationText(estimate) + "（按刚才的预检耗时估算，仅供参考）。";
            string text = (empty.Count > 0 ? "以下源区域全部为空，会清空目标对应区域：" + string.Join("、", empty.Take(5).ToArray()) + (empty.Count > 5 ? " 等" : "") + line : "")
                + (overwrite > 0 ? "目标区域已有数据：" + filled + " 处写入中共有 " + overwrite + " 个单元格已有内容，执行后将被源数据覆盖（源中为空的单元格会清空目标）。" + line : "")
                + saving + line + (overwrite > 0 ? "是否覆盖原数据？" : "是否执行？");
            if (UiDialog.Show(text, overwrite > 0 ? "确认覆盖原数据" : "确认执行复制", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            // Formulas that refer to sheets a target lacks are pasted as
            // values there, but only with the user's consent.
            var fallback = ready.Where(m => m.Result.ValueFallbacks > 0).ToList();
            if (fallback.Count > 0)
            {
                string list = string.Join(Environment.NewLine, fallback.Take(10).Select(m => "· " + Path.GetFileName(m.TargetPath) + "：缺少工作表“" + string.Join("”“", m.Result.MissingSheets.ToArray()) + "”，" + m.Result.ValueFallbacks + " 个公式单元格").ToArray())
                    + (fallback.Count > 10 ? Environment.NewLine + "…… 共 " + fallback.Count + " 处" : "");
                string warn = "以下目标缺少源公式引用的工作表：" + Environment.NewLine + Environment.NewLine + list + Environment.NewLine + Environment.NewLine
                    + "这些单元格将改为粘贴源中的当前计算结果（数值），其他单元格照常复制。" + Environment.NewLine + Environment.NewLine + "是否继续？";
                if (UiDialog.Show(warn, "目标缺少工作表", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) { form.Message("已取消执行：可为目标补上缺少的工作表，或把映射改为“只写数值”。"); return; }
            }
            var plan = form.ExecutionPlan(previewed);
            if (plan.OutputCopies) PrepareOutputCopies(plan);
            _migrationPlan = null; _planCancelled = false;
            form.Run(RunPlan(form, plan, true),
                delegate { FinishOutputCopies(plan); form.RefreshResults(); },
                delegate(Exception error) { FinishOutputCopies(plan); form.Message(error.Message); });
        }
        // Result-copy mode: one "_复制结果" copy per target receives every pass.
        private static void PrepareOutputCopies(MigrationPlan plan)
        {
            foreach (string path in plan.Mappings.Select(m => m.TargetPath).Distinct(StringComparer.OrdinalIgnoreCase).ToList())
            {
                string output = DataMigrationEngine.OutputPath(path, p => File.Exists(p) || Directory.Exists(p));
                File.Copy(path, output, false); plan.Outputs[path] = output;
            }
            foreach (var pass in plan.Passes) foreach (var target in pass.Request.Targets) target.Path = plan.Outputs[target.Path];
        }
        private static void FinishOutputCopies(MigrationPlan plan)
        {
            if (!plan.OutputCopies || plan.Outputs.Count == 0) return;
            var targets = plan.Passes.SelectMany(p => p.Request.Targets).ToList();
            // Backups of our own result copies are scaffolding, not user data.
            try { MigrationOutputTransaction.DeleteBackups(targets.Select(t => t.Backup).Where(b => !string.IsNullOrEmpty(b)).ToList()); }
            catch (Exception error) { Log("复制结果副本备份清理：" + error); }
            foreach (var target in targets) target.Backup = "";
            foreach (var pair in plan.Outputs.ToList())
            {
                var mine = plan.Mappings.Where(m => string.Equals(m.TargetPath, pair.Key, StringComparison.OrdinalIgnoreCase)).ToList();
                if (mine.Any(m => m.Result.Status == "成功")) { foreach (var m in mine) if (m.Result.Status == "成功") m.Result.Output = pair.Value; }
                else try { File.Delete(pair.Value); } catch (Exception error) { Log("未使用的复制结果副本：" + error); }
            }
            plan.Outputs.Clear();
        }
        // Workbooks WPS holds that would make the precheck fail: a source with
        // unsaved changes (must be saved), a target that is open while its
        // original is to be replaced (must be saved and closed), or an open
        // target with unsaved changes in copy mode (must be saved). With the
        // user's consent they are saved (and closed) here; otherwise nothing
        // is touched. Returns false when the user declined or a save failed.
        // Two 逐行重算 rows writing the same target area: the later row wins.
        // Ask before going on instead of refusing.
        private static bool ConfirmOverrides(DataMigrationForm form, MigrationPlan plan)
        {
            if (plan.Overrides.Count == 0) return true;
            string line = Environment.NewLine;
            string text = "以下行会写入同一个目标的同一区域，后面的会覆盖前面的：" + line + line
                + string.Join(line, plan.Overrides.Take(12).Select(o => "· " + o).ToArray()) + (plan.Overrides.Count > 12 ? line + "…… 共 " + plan.Overrides.Count + " 处" : "")
                + line + line + "是否继续？";
            if (UiDialog.Show(form, text, "目标区域重复", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes) return true;
            form.Message("已取消：请在明细中调整这些行的目标后再预检。"); return false;
        }
        // Precheck only needs every workbook saved: open ones are read in
        // place (逐行重算 sources through a copy). Execution replacing target
        // files needs those targets closed.
        private static bool ReleaseBusyWorkbooks(DataMigrationForm form, MigrationPlan plan, bool execute)
        {
            var sources = new HashSet<string>(plan.Passes.Select(p => Path.GetFullPath(p.Request.SourcePath)), StringComparer.OrdinalIgnoreCase);
            var targets = new HashSet<string>(plan.Passes.SelectMany(p => p.Request.Targets).Select(t => Path.GetFullPath(t.Path)), StringComparer.OrdinalIgnoreCase);
            var save = new List<string>(); var close = new List<string>();
            using (var bag = new ComBag())
            {
                dynamic books = bag.Own(((dynamic)_app).Workbooks);
                for (int i = 1; i <= Convert.ToInt32(books.Count); i++)
                {
                    dynamic book = bag.Own(books.Item(i)); string full;
                    try { full = Path.GetFullPath(Convert.ToString(book.FullName)); } catch (Exception) { continue; }
                    bool saved = ComFormulaHost.Boolean(book.Saved, "Saved");
                    if (execute && targets.Contains(full) && plan.ReplaceTargets) close.Add(full);
                    else if ((sources.Contains(full) || targets.Contains(full)) && !saved) save.Add(full);
                }
            }
            if (save.Count == 0 && close.Count == 0) return true;
            string line = Environment.NewLine;
            string text = (execute ? "以下工作簿正在 WPS 中打开，执行前需要处理：" : "以下工作簿有未保存修改，预检读取的是已保存的内容：")
                + (save.Count == 0 ? "" : line + line + "有未保存修改（将保存，不关闭）：" + line + string.Join(line, save.Select(p => "· " + Path.GetFileName(p)).ToArray()))
                + (close.Count == 0 ? "" : line + line + "要更新原文件的目标（将保存并关闭，完成后可重新打开）：" + line + string.Join(line, close.Select(p => "· " + Path.GetFileName(p)).ToArray()))
                + line + line + "保存会把你在这些工作簿中尚未保存的修改一并写入文件。" + line + (close.Count == 0 ? "是否保存后继续？" : "是否保存并关闭后继续？");
            // Saving alone is done for the user and reported; only closing a
            // target the user has open still asks first.
            if (close.Count > 0 && UiDialog.Show(form, text, "工作簿被占用", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            { form.Message(execute ? "已取消：请保存并关闭要更新的目标工作簿后再执行。" : "已取消：请先保存这些工作簿再预检。"); return false; }
            var failures = new List<string>();
            foreach (string path in save.Concat(close))
            {
                bool shut = close.Contains(path, StringComparer.OrdinalIgnoreCase);
                try
                {
                    using (var bag = new ComBag())
                    {
                        dynamic books = bag.Own(((dynamic)_app).Workbooks); dynamic match = null;
                        for (int i = 1; i <= Convert.ToInt32(books.Count) && match == null; i++)
                        {
                            dynamic book = bag.Own(books.Item(i));
                            if (string.Equals(Path.GetFullPath(Convert.ToString(book.FullName)), path, StringComparison.OrdinalIgnoreCase)) match = book;
                        }
                        if (match == null) continue;
                        if (ComFormulaHost.Boolean(match.ReadOnly, "ReadOnly")) throw new InvalidOperationException("文件以只读方式打开，无法保存");
                        if (!ComFormulaHost.Boolean(match.Saved, "Saved")) match.Save();
                        if (shut) match.Close(false);
                    }
                }
                catch (Exception error) { Log("复制释放占用工作簿：" + path + " " + error); failures.Add(Path.GetFileName(path) + "：" + error.Message); }
            }
            if (failures.Count > 0)
            {
                form.Message("未能处理：" + string.Join("；", failures.ToArray()) + "。请手动保存" + (execute ? "或关闭" : "") + "后再试。");
                return false;
            }
            form.Message(close.Count == 0 ? "已替你保存 " + save.Count + " 个有未保存修改的工作簿（未关闭）：" + string.Join("、", save.Select(p => Path.GetFileName(p)).ToArray()) + "，继续" + (execute ? "执行" : "预检") + "。" : "已保存 " + save.Count + " 个、保存并关闭 " + close.Count + " 个工作簿，继续执行。");
            return true;
        }
        private static void MigrationAction(DataMigrationForm form, Action action)
        { try { action(); } catch (Exception error) { Log("批量复制数据：" + error); form.Message(error.Message); } }
        private static void ResetMigration()
        {
            if (_migrationPlan != null) foreach (var mapping in _migrationPlan.Mappings)
                if (mapping.Result != null && mapping.Result.Status == "可执行") { mapping.Result.Ready = false; mapping.Result.Status = "需重新预检"; }
            _migrationPlan = null;
            var host = _migrationHost; _migrationHost = null;
            if (host != null) try { host.Dispose(); } catch (Exception error) { Log("复制会话释放：" + error); }
        }
        internal static void CloseDataMigrationForm()
        {
            var form = _migrationForm; _migrationForm = null;
            if (form != null) { form.Dispose(); } ResetMigration();
        }
        public IPictureDisp GetImage_btnDataMigration(object control) { return GetButtonImageCore("btnDataMigration"); }
    }
    internal static class MigrationForeground
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr window, int command);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool AttachThreadInput(uint attach, uint to, bool on);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr window);
        [System.Runtime.InteropServices.DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] private struct POINT { public int X, Y; }
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out RECT rect);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINT point);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT point);
        [System.Runtime.InteropServices.DllImport("user32.dll")] private static extern void mouse_event(uint flags, int x, int y, uint data, UIntPtr extra);
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)] private static extern int GetClassName(IntPtr window, System.Text.StringBuilder name, int size);
        private static string ClassOf(IntPtr window) { var name = new System.Text.StringBuilder(128); GetClassName(window, name, name.Capacity); return name.ToString(); }
        internal static IntPtr ActiveFrame(object app, IntPtr fallback)
        {
            try
            {
                IntPtr handle = new IntPtr(Convert.ToInt64(((dynamic)app).ActiveWindow.Hwnd));
                IntPtr root = handle == IntPtr.Zero ? IntPtr.Zero : GetAncestor(handle, 2);
                if (root != IntPtr.Zero) return root;
            }
            catch (Exception) { }
            return fallback;
        }
        // Is this screen point on the sheet grid of the WPS frame: inside the
        // frame, clear of the ribbon above and the status/tab bar below, and
        // on the frame itself or a large child of it (not the selection box)?
        private static bool OnGrid(IntPtr main, RECT frame, POINT point)
        {
            int width = frame.Right - frame.Left, height = frame.Bottom - frame.Top;
            if (point.X < frame.Left + width * 0.03 || point.X > frame.Right - width * 0.03 || point.Y < frame.Top + height * 0.25 || point.Y > frame.Bottom - height * 0.08) return false;
            IntPtr hit = WindowFromPoint(point);
            if (hit == main) return true;
            if (hit == IntPtr.Zero || GetAncestor(hit, 2) != main) return false;
            RECT area; return GetWindowRect(hit, out area) && area.Right - area.Left >= width * 0.4 && area.Bottom - area.Top >= height * 0.3;
        }
        // One simulated left click on the sheet grid. Where the pointer is
        // already on the grid, click right there without moving it. Otherwise
        // click a grid point and report where the pointer was (moved), so the
        // caller puts it back a moment later.
        private static bool ClickGrid(IntPtr main, out POINT moved, out bool wasMoved)
        {
            moved = new POINT(); wasMoved = false;
            RECT frame; if (main == IntPtr.Zero || !GetWindowRect(main, out frame)) return false;
            int width = frame.Right - frame.Left, height = frame.Bottom - frame.Top; if (width < 200 || height < 200) return false;
            POINT now; bool known = GetCursorPos(out now);
            if (known && OnGrid(main, frame, now))
            {
                mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
                Connect.Log("复制框选：在鼠标当前位置模拟点击 @" + now.X + "," + now.Y);
                return true;
            }
            foreach (var spot in new[] { new[] { 0.3, 0.55 }, new[] { 0.7, 0.55 }, new[] { 0.3, 0.75 }, new[] { 0.7, 0.75 }, new[] { 0.5, 0.8 }, new[] { 0.35, 0.65 } })
            {
                var point = new POINT { X = frame.Left + (int)(width * spot[0]), Y = frame.Top + (int)(height * spot[1]) };
                if (!OnGrid(main, frame, point) || !SetCursorPos(point.X, point.Y)) continue;
                mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero); mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
                if (known) { moved = now; wasMoved = true; }
                Connect.Log("复制框选：鼠标不在表格上，临时在 @" + point.X + "," + point.Y + " 模拟点击后移回");
                return true;
            }
            Connect.Log("复制框选：未模拟点击（窗口 " + ClassOf(main) + " " + width + "x" + height + "）");
            return false;
        }
        // Windows refuses SetForegroundWindow unless the caller shares input
        // with the current foreground thread, so attach to it briefly.
        internal static void Bring(IntPtr window)
        {
            if (window == IntPtr.Zero) return;
            try
            {
                if (IsIconic(window)) ShowWindow(window, 9);
                uint process; uint foreground = GetWindowThreadProcessId(GetForegroundWindow(), out process), me = GetCurrentThreadId();
                bool attached = foreground != 0 && foreground != me && AttachThreadInput(me, foreground, true);
                try { BringWindowToTop(window); SetForegroundWindow(window); }
                finally { if (attached) AttachThreadInput(me, foreground, false); }
            }
            catch (Exception) { }
        }
        // While a modal box (the WPS selection InputBox) runs, check a few
        // times that WPS still has the focus and take it back if not.
        // After the click (second tick) the selection is put back (third).
        internal static System.Windows.Forms.Timer BringSoon(IntPtr window, Action reselect = null)
        {
            if (window == IntPtr.Zero) return null;
            var timer = new System.Windows.Forms.Timer { Interval = 150 }; int ticks = 0; bool clicked = false, moved = false; POINT back = new POINT();
            uint own = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
            timer.Tick += delegate
            {
                ticks++;
                uint process; GetWindowThreadProcessId(GetForegroundWindow(), out process);
                if (process != own) Bring(window);
                try
                {
                    if (ticks == 2) { Connect.Log("复制框选：前台窗口 " + ClassOf(GetForegroundWindow())); clicked = ClickGrid(window, out back, out moved); }
                    else if (ticks == 3)
                    {
                        // The click has been handled by now: put the pointer back.
                        if (moved) SetCursorPos(back.X, back.Y);
                        if (clicked && reselect != null) reselect();
                    }
                }
                catch (Exception) { }
                if (ticks >= 4) { timer.Stop(); timer.Dispose(); }
            };
            timer.Start();
            return timer;
        }
    }
}
