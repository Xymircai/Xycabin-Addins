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
            }
        }
        finally { DialogSizeStore.RootOverride = oldRoot; }
    }
}
