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
                        if (!ReleaseBusyWorkbooks(form, plan)) return;
                        _migrationPlan = plan; _planCancelled = false;
                        form.Run(RunPlan(form, plan, false), delegate { form.RefreshResults(); }, delegate(Exception error) { form.Message(error.Message); ResetMigration(); });
                    });
                };
                form.ExecuteRequested += delegate
                {
                    MigrationAction(form, delegate
                    {
                        if (_migrationPlan == null) throw new InvalidOperationException("请先重新预检。");
                        if (!ReleaseBusyWorkbooks(form, _migrationPlan)) return;
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
                        catch (Exception error) { Log("迁移工作簿读取失败：" + path + " " + error); form.AddBook(target, path, new string[0], error.Message); }
                    }
                };
                // Pick button in workbook mode: open the workbook read-only (or use
                // it if already open), select with the mouse, then come back.
                form.RangePickRequested += delegate(bool target, string path, string sheet, string current, Action<string> done)
                {
                    if (form.Running) return;
                    string address = null; Exception failure = null;
                    form.Hide();
                    try { using (var host = new DataMigrationHost(_app)) address = host.PickSourceRange(path, sheet, current, null, target ? "目标" : "源"); }
                    catch (Exception error) { failure = error; Log("迁移框选区域：" + error); }
                    finally { form.Show(); form.Activate(); }
                    if (failure != null) form.Message("框选失败：" + failure.Message);
                    else if (address == null) form.Message("已取消框选，区域保持不变。");
                    else done(address);
                };
                form.FormClosed += delegate { ResetMigration(); _migrationForm = null; form.Dispose(); };
                DefaultMigrationSource(form);
                IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception error) { Log("批量迁移数据：" + error); UiDialog.Show(error.Message, "批量迁移数据", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
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
            catch (Exception error) { Log("批量迁移数据默认源：" + error); }
        }
        // Runs every pass of a plan through the existing single-source engine.
        // A pass that fails only marks its own mappings; the other passes go
        // on, which is what gives each target its partial save.
        private static IEnumerable<MigrationProgress> RunPlan(DataMigrationForm form, MigrationPlan plan, bool execute)
        {
            for (int index = 0; index < plan.Passes.Count; index++)
            {
                var pass = plan.Passes[index];
                string label = "[" + (index + 1) + "/" + plan.Passes.Count + "] " + pass.Source.Label + " " + pass.Request.SourceSheet + "!" + pass.Request.SourceRange + "：";
                if (_planCancelled) { MarkPass(pass, "已取消", "已取消，未执行"); continue; }
                var host = new DataMigrationHost(_app); _migrationHost = host;
                try
                {
                    Exception failure;
                    MigrationProgress step;
                    // Execution re-checks each target right before writing,
                    // because an earlier pass may just have updated it.
                    var preview = host.Preview(pass.Request).GetEnumerator();
                    try { while (Advance(preview, out step, out failure)) yield return new MigrationProgress(label + step.Message, step.Pause); }
                    finally { preview.Dispose(); }
                    if (failure != null) { MarkPass(pass, failure is OperationCanceledException ? "已取消" : "失败", "源区域预检：" + failure.Message); continue; }
                    if (!execute || !pass.Request.Targets.Any(t => t.Ready)) continue;
                    var run = host.Execute().GetEnumerator();
                    try { while (Advance(run, out step, out failure)) yield return new MigrationProgress(label + step.Message, step.Pause); }
                    finally { run.Dispose(); }
                    if (failure != null) MarkPass(pass, failure is OperationCanceledException ? "已取消" : "失败", "执行：" + failure.Message);
                    if (!plan.OutputCopies) foreach (var target in pass.Request.Targets) if (target.Status == "成功") form.RecordRestorePoint(target.Path, target.Backup);
                    if (host.Cancelled) _planCancelled = true;
                }
                finally
                {
                    _migrationHost = null;
                    try { host.Dispose(); } catch (Exception error) { Log("迁移会话释放：" + error); }
                }
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
                ? "将更新 " + targets + " 个目标原文件（共 " + ready.Count + " 处写入）；原版本自动备份到同目录的“迁移备份”文件夹，可用“还原操作”恢复。源文件不修改。"
                : "将为 " + targets + " 个目标各另存一个“迁移结果”副本（共 " + ready.Count + " 处写入），目标原文件不修改。";
            string text = (empty.Count > 0 ? "以下源区域全部为空，会清空目标对应区域：" + string.Join("、", empty.Take(5).ToArray()) + (empty.Count > 5 ? " 等" : "") + line : "")
                + (overwrite > 0 ? "目标区域已有数据：" + filled + " 处写入中共有 " + overwrite + " 个单元格已有内容，执行后将被源数据覆盖（源中为空的单元格会清空目标）。" + line : "")
                + saving + line + (overwrite > 0 ? "是否覆盖原数据？" : "是否执行？");
            if (UiDialog.Show(text, overwrite > 0 ? "确认覆盖原数据" : "确认执行迁移", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            var plan = form.ExecutionPlan(previewed);
            if (plan.OutputCopies) PrepareOutputCopies(plan);
            _migrationPlan = null; _planCancelled = false;
            form.Run(RunPlan(form, plan, true),
                delegate { FinishOutputCopies(plan); form.RefreshResults(); },
                delegate(Exception error) { FinishOutputCopies(plan); form.Message(error.Message); });
        }
        // Result-copy mode: one "_迁移结果" copy per target receives every pass.
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
            catch (Exception error) { Log("迁移结果副本备份清理：" + error); }
            foreach (var target in targets) target.Backup = "";
            foreach (var pair in plan.Outputs.ToList())
            {
                var mine = plan.Mappings.Where(m => string.Equals(m.TargetPath, pair.Key, StringComparison.OrdinalIgnoreCase)).ToList();
                if (mine.Any(m => m.Result.Status == "成功")) { foreach (var m in mine) if (m.Result.Status == "成功") m.Result.Output = pair.Value; }
                else try { File.Delete(pair.Value); } catch (Exception error) { Log("未使用的迁移结果副本：" + error); }
            }
            plan.Outputs.Clear();
        }
        // Workbooks WPS holds that would make the precheck fail: a source with
        // unsaved changes (must be saved), a target that is open while its
        // original is to be replaced (must be saved and closed), or an open
        // target with unsaved changes in copy mode (must be saved). With the
        // user's consent they are saved (and closed) here; otherwise nothing
        // is touched. Returns false when the user declined or a save failed.
        private static bool ReleaseBusyWorkbooks(DataMigrationForm form, MigrationPlan plan)
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
                    if (targets.Contains(full) && plan.ReplaceTargets) close.Add(full);
                    else if ((sources.Contains(full) || targets.Contains(full)) && !saved) save.Add(full);
                }
            }
            if (save.Count == 0 && close.Count == 0) return true;
            string line = Environment.NewLine;
            string text = "以下工作簿正在 WPS 中被占用，预检无法读取或替换："
                + (save.Count == 0 ? "" : line + line + "有未保存修改（将保存）：" + line + string.Join(line, save.Select(p => "· " + Path.GetFileName(p)).ToArray()))
                + (close.Count == 0 ? "" : line + line + "目标工作簿已打开（将保存并关闭，迁移后可重新打开）：" + line + string.Join(line, close.Select(p => "· " + Path.GetFileName(p)).ToArray()))
                + line + line + "保存会把你在这些工作簿中尚未保存的修改一并写入文件。" + line + "是否保存并关闭后继续？";
            if (UiDialog.Show(form, text, "工作簿被占用", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) != DialogResult.Yes)
            { form.Message("已取消：请自行保存源工作簿、保存并关闭目标工作簿后再预检。"); return false; }
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
                catch (Exception error) { Log("迁移释放占用工作簿：" + path + " " + error); failures.Add(Path.GetFileName(path) + "：" + error.Message); }
            }
            if (failures.Count > 0)
            {
                form.Message("未能处理：" + string.Join("；", failures.ToArray()) + "。请手动保存或关闭后再预检。");
                return false;
            }
            form.Message("已保存 " + save.Count + " 个、保存并关闭 " + close.Count + " 个被占用的工作簿，继续预检。");
            return true;
        }
        private static void MigrationAction(DataMigrationForm form, Action action)
        { try { action(); } catch (Exception error) { Log("批量迁移数据：" + error); form.Message(error.Message); } }
        private static void ResetMigration()
        {
            if (_migrationPlan != null) foreach (var mapping in _migrationPlan.Mappings)
                if (mapping.Result != null && mapping.Result.Status == "可执行") { mapping.Result.Ready = false; mapping.Result.Status = "需重新预检"; }
            _migrationPlan = null;
            var host = _migrationHost; _migrationHost = null;
            if (host != null) try { host.Dispose(); } catch (Exception error) { Log("迁移会话释放：" + error); }
        }
        internal static void CloseDataMigrationForm()
        {
            var form = _migrationForm; _migrationForm = null;
            if (form != null) { form.Dispose(); } ResetMigration();
        }
        public IPictureDisp GetImage_btnDataMigration(object control) { return GetButtonImageCore("btnDataMigration"); }
    }
}
