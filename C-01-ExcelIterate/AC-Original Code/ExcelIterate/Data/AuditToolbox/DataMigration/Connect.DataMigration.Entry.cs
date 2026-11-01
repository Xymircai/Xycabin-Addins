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
        private static MigrationRequest _migrationRequest;
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
                        ResetMigration(); _migrationRequest = form.Request(); _migrationHost = new DataMigrationHost(_app);
                        form.Run(_migrationHost.Preview(_migrationRequest), delegate { form.RefreshResults(); }, delegate(Exception error) { form.Message(error.Message); ResetMigration(); });
                    });
                };
                form.ExecuteRequested += delegate
                {
                    MigrationAction(form, delegate
                    {
                        if (_migrationHost == null || _migrationRequest == null) throw new InvalidOperationException("请先重新预检。");
                        var ready = _migrationRequest.Targets.Where(t => t.Ready).ToList(); int overwrite = ready.Sum(t => t.Overwritten);
                        if (ready.Count == 0) throw new InvalidOperationException("没有通过预检的目标。");
                        if (_migrationRequest.NonEmptyCells == 0 && UiDialog.Show("源区域 " + _migrationRequest.SourceSheet + "!" + _migrationRequest.SourceRange + " 全部为空。继续会清空 " + ready.Count + " 个目标的对应区域。确定继续？", "全空白源区域", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                        // Existing data in the target area is overwritten only after
                        // an explicit yes; once confirmed it is replaced, not reported
                        // as a failure.
                        string saving = _migrationRequest.ReplaceTargets
                            ? "将更新 " + ready.Count + " 个目标原文件；原版本自动备份到同目录的“迁移备份”文件夹，可用“还原操作”恢复。源文件不修改。"
                            : "将另存 " + ready.Count + " 个“迁移结果”副本，目标原文件不修改。";
                        if (overwrite > 0)
                        {
                            int filled = ready.Count(t => t.Overwritten > 0);
                            if (UiDialog.Show("目标区域已有数据：" + filled + " 个目标中共有 " + overwrite + " 个单元格已有内容，执行后将被源数据覆盖（源中为空的单元格会清空目标）。" + Environment.NewLine + Environment.NewLine + saving + Environment.NewLine + Environment.NewLine + "是否覆盖原数据？", "确认覆盖原数据", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                        }
                        else if (_migrationRequest.ReplaceTargets && UiDialog.Show(saving + "每个文件校验通过后才替换。是否执行？", "确认更新目标原文件", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                        form.Run(_migrationHost.Execute(), delegate { form.RefreshResults(); ResetMigration(); }, delegate(Exception error) { form.Message(error.Message); ResetMigration(); });
                    });
                };
                form.CancelRequested += delegate { if (_migrationHost != null) _migrationHost.Cancelled = true; };
                form.InputsChanged += ResetMigration;
                form.SourcePicked += delegate(string path)
                {
                    MigrationAction(form, delegate { using (var host = new DataMigrationHost(_app)) { var names = host.Sheets(path); form.SetSource(path, names[0], "", names); } });
                };
                form.TargetsPicked += delegate(string[] paths)
                {
                    var selectedArea = MigrationArea.Parse(form.SourceRange.Text.Trim().Length == 0 ? "A1" : form.SourceRange.Text);
                    string start = MigrationArea.Cell(selectedArea.Row, selectedArea.Column);
                    foreach (string path in paths)
                    {
                        try
                        {
                            using (var host = new DataMigrationHost(_app))
                            { var names = host.Sheets(path); form.AddTarget(path, names, form.SourceSheet.Text.Length == 0 ? names[0] : form.SourceSheet.Text, start); }
                        }
                        catch (Exception error) { Log("迁移目标读取失败：" + path + " " + error); form.AddTarget(path, new string[0], "", start, error.Message); }
                    }
                };
                form.SelectionRequested += delegate { MigrationAction(form, delegate { MigrationSelection(form, false); }); };
                form.SourceRangeRequested += delegate
                {
                    if (_busy || form.Running) return;
                    MigrationAction(form, delegate
                    {
                        if (form.SourcePath.Text.Trim().Length == 0 || form.SourceSheet.Text.Trim().Length == 0) throw new InvalidOperationException("请先选择源文件和源工作表，再在表中框选区域。");
                        ResetMigration(); _busy = true;
                        try
                        {
                            form.Hide();
                            using (var host = new DataMigrationHost(_app))
                            {
                                string address = host.PickSourceRange(form.SourcePath.Text.Trim(), form.SourceSheet.Text.Trim(), form.SourceRange.Text.Trim(), null);
                                if (address != null) { form.SourceRange.Text = address; form.Message("已选择源区域 " + address + "，请添加目标并重新预检。"); }
                            }
                        }
                        finally { _busy = false; if (!form.IsDisposed) { form.Show(); form.Activate(); } }
                    });
                };
                form.TargetSelectionRequested += delegate { MigrationAction(form, delegate { MigrationSelection(form, true); }); };
                form.FormClosed += delegate { ResetMigration(); _migrationForm = null; form.Dispose(); };
                IntPtr owner = GetHostWindowHandle(); if (owner == IntPtr.Zero) form.Show(); else form.Show(new WindowWrapper(owner));
            }
            catch (Exception error) { Log("批量迁移数据：" + error); UiDialog.Show(error.Message, "批量迁移数据", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        private static void MigrationAction(DataMigrationForm form, Action action)
        { try { action(); } catch (Exception error) { Log("批量迁移数据：" + error); form.Message(error.Message); } }
        private static void MigrationSelection(DataMigrationForm form, bool target)
        {
            using (var bag = new ComBag())
            {
                dynamic book = bag.Own(((dynamic)_app).ActiveWorkbook), sheet = bag.Own(((dynamic)_app).ActiveSheet), selection = bag.Own(((dynamic)_app).Selection);
                if (book == null || sheet == null || selection == null) throw new InvalidOperationException("请先在已保存工作簿中选择单个矩形区域。");
                dynamic areas = bag.Own(selection.Areas); if (Convert.ToInt32(areas.Count) != 1) throw new InvalidOperationException("只支持单个连续矩形选区。");
                string path = Convert.ToString(book.FullName), name = Convert.ToString(sheet.Name);
                if (!Path.IsPathRooted(path) || !File.Exists(path) || !ComFormulaHost.Boolean(book.Saved, "Saved")) throw new InvalidOperationException("请先保存当前工作簿再读取选区。");
                dynamic rows = bag.Own(selection.Rows), columns = bag.Own(selection.Columns);
                var area = new MigrationArea { Row = Convert.ToInt32(selection.Row), Column = Convert.ToInt32(selection.Column), Rows = Convert.ToInt32(rows.Count), Columns = Convert.ToInt32(columns.Count) };
                using (var host = new DataMigrationHost(_app))
                {
                    var names = host.Sheets(path);
                    if (target) form.AddTarget(path, names, name, MigrationArea.Cell(area.Row, area.Column)); else form.SetSource(path, name, area.Address, names);
                }
            }
        }
        private static void ResetMigration()
        {
            var host = _migrationHost;
            if (_migrationRequest != null) foreach (var target in _migrationRequest.Targets)
                if (target.Status == "可执行") { target.Ready = false; target.Status = "需重新预检"; }
            _migrationHost = null; _migrationRequest = null;
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
