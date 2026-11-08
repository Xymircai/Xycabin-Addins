using System;
using System.Collections.Generic;
using System.IO;
using Excel_Iterate;

internal static class BatchPresetChecks
{
    internal static void Run(string folder, Action<bool, string> check)
    {
        string oldRoot = DialogSizeStore.RootOverride; DialogSizeStore.RootOverride = folder;
        Action<Action, string> rejects = delegate(Action action, string name)
        { bool failed = false; try { action(); } catch { failed = true; } check(failed, name); };
        try
        {
            var value = new BatchPreset { Feature = "company", Name = "中文预设", Context = "template.xlsx", Values = new Dictionary<string, string> { { "field", "编码&全称" } } };
            string file = BatchPresetStore.FileFor("company", value.Name);
            BatchPresetStore.Write(file, value, false);
            check(BatchPresetStore.Read(file, "company").Values["field"] == "编码&全称", "preset Unicode roundtrip");
            byte[] before = File.ReadAllBytes(file);
            rejects(delegate { BatchPresetStore.Write(file, value, false); }, "same-name save cannot silently overwrite");
            check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(file)), "refused overwrite preserves bytes");
            using (var locked = new FileStream(file + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                rejects(delegate { BatchPresetStore.Write(file, value, true); }, "concurrent preset write refused");
            check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(file)), "locked write preserves previous preset");
            value.Values["field"] = "新字段"; BatchPresetStore.Write(file, value, true);
            check(BatchPresetStore.Read(file, "company").Values["field"] == "新字段", "explicit overwrite replaces complete preset");
            rejects(delegate { BatchPresetStore.Read(file, "summary"); }, "cross-feature preset rejected");
            value.Version = 2;
            rejects(delegate { BatchPresetStore.Write(file, value, true); }, "newer schema cannot overwrite existing preset");
            value.Version = 1;
            var oldPreset = BatchPresetStore.Decode(BatchPresetStore.Encode(value), "company");
            check(oldPreset.Context == value.Context && oldPreset.Version == 1, "version-one JSON remains compatible with isolated dialog transport");
            string imported = BatchPresetStore.Import(file, "company");
            check(imported != file && File.Exists(file) && BatchPresetStore.Read(imported, "company").Name.Contains("(2)"), "duplicate import preserves old preset and allocates another name");
            rejects(delegate { BatchPresetStore.Rename(imported, "company", value.Name); }, "rename cannot replace another preset");
            check(File.Exists(file) && File.Exists(imported), "rejected rename preserves both presets");
            string renamed = BatchPresetStore.Rename(imported, "company", "重命名预设");
            check(!File.Exists(imported) && BatchPresetStore.Read(renamed, "company").Name == "重命名预设", "rename changes both library file and serialized name");
            string audit = Path.Combine(folder, "audit.xlsx"); File.WriteAllText(audit, "audit-original");
            BatchPresetStore.Delete(renamed);
            check(!File.Exists(renamed) && File.ReadAllText(audit) == "audit-original", "deleting a preset leaves audit file untouched");
            using (var manager = new BatchPresetManagerForm("company"))
            {
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                var list = (System.Windows.Forms.ListBox)typeof(BatchPresetManagerForm).GetField("list", flags).GetValue(manager);
                var search = (System.Windows.Forms.TextBox)typeof(BatchPresetManagerForm).GetField("search", flags).GetValue(manager);
                var detail = (System.Windows.Forms.RichTextBox)typeof(BatchPresetManagerForm).GetField("detail", flags).GetValue(manager);
                check(typeof(BatchPresetManagerForm).GetField("view", flags).FieldType == typeof(PresetManagerView), "single presets use extracted migration view");
                check(list.DrawMode == System.Windows.Forms.DrawMode.OwnerDrawFixed && list.BorderStyle == System.Windows.Forms.BorderStyle.None, "single presets reuse migration two-line owner-drawn list");
                check(search.Width == 220 && manager.Controls.Count == 1, "single presets reuse migration search width and single layout root");
                check(list.Items.Count == 1 && detail.Text.Contains("template.xlsx"), "manager lists version-one presets with human-readable file details");
                search.Text = "not-found"; check(list.Items.Count == 0, "preset search filters names without altering library files");
                search.Clear(); check(list.Items.Count == 1 && manager.Chosen == null, "browsing and cancelling manager does not apply a preset");
            }
            using (var migration = new MigrationPresetManager(null))
            {
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                check(typeof(MigrationPresetManager).GetField("view", flags).FieldType == typeof(PresetManagerView), "migration and single features share the same production view");
                check(migration.Controls.Count == 1, "migration adapter retains one shared layout root");
            }
            using (var preview = DialogPreviewFactory.Create("batch.preset.manager", DialogCatalog.Find("batch.preset.manager").Default))
                check(preview != null, "shared preset manager participates in inert preview catalog");
            foreach (string name in new[] { "../bad", "CON", "", "a.", "a/b" })
                rejects(delegate { BatchPresetStore.FileFor("company", name); }, "invalid preset filename refused");
            string bad = Path.Combine(folder, "broken.xybatch.json"); File.WriteAllText(bad, "{broken");
            rejects(delegate { BatchPresetStore.Read(bad, "company"); }, "malformed imported JSON refused");
            using (var form = new SingleSummaryForm())
            {
                check(!form.CheckYearFiles.Checked, "preset support does not change first-open check-files default");
                form.SetContext("test.xlsx", new List<string> { "Lead", "汇总" }, "Lead", "汇总", "F");
                form.FirstYear.Value = 2023; form.LastYear.Value = 2026; form.YearRow.Value = 2; form.AnnualMode.Checked = true;
                form.CheckYearFiles.Checked = true;
                var preset = form.CapturePreset(); preset.Name = "年度跨度";
                form.FirstYear.Value = 2025; form.CheckYearFiles.Checked = false;
                form.ApplyPreset(preset);
                check(form.FirstYear.Value == 2023 && form.LastYear.Value == 2026 && form.CheckYearFiles.Checked, "summary preset restores saved annual options on explicit apply");
                form.SetUndoAvailable(true);
                rejects(delegate { form.ApplyPreset(preset); }, "pending undo batch blocks preset application");
                form.SetUndoAvailable(false);
                preset.Values["firstYear"] = "10000";
                rejects(delegate { form.ApplyPreset(preset); }, "invalid saved year rejected before mutation");
                check(form.FirstYear.Value == 2023, "invalid preset preserves old summary values");
                preset.Values["firstYear"] = "2023"; preset.Values["target"] = "missing";
                rejects(delegate { form.ApplyPreset(preset); }, "missing summary sheet rejected");
                check(form.FirstYear.Value == 2023, "missing-sheet preset preserves old values");
                form.SetBusy(true); rejects(delegate { form.ApplyPreset(preset); }, "busy summary refuses presets");
                rejects(delegate { form.ApplyPreparedPreset(preset, "replacement.xlsx", new List<string> { "Lead", "汇总" }); }, "prepared preset validates candidate sheets before changing context");
                check(form.BuildRequest().WorkbookName == "test.xlsx", "rejected replacement retains old summary workbook context");
                preset.Values["target"] = "汇总";
                form.ApplyPreparedPreset(preset, "replacement.xlsx", new List<string> { "Lead", "汇总" });
                check(form.IsBusy && form.BuildRequest().WorkbookName == "replacement.xlsx", "prepared preset changes explicit context without releasing busy guard");
                form.SetUndoAvailable(true);
                rejects(delegate { form.ApplyPreparedPreset(preset, "another.xlsx", new List<string> { "Lead", "汇总" }); }, "prepared preset also respects pending rollback protection");
            }
        }
        finally { DialogSizeStore.RootOverride = oldRoot; }
    }
}
