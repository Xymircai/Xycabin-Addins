using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Same explicit, named local-library workflow as migration presets. The
    // schema is separate: migration mappings must never be applied as settings.
    [DataContract]
    internal sealed class BatchPreset
    {
        [DataMember] internal string Format = "xycabin.batch-settings";
        [DataMember] internal int Version = 1;
        [DataMember] internal string Feature, Name, Context, Saved;
        [DataMember] internal Dictionary<string, string> Values;
    }
    internal static class BatchPresetStore
    {
        internal static string Library(string feature)
        { return Path.Combine(DialogSizeStore.RootOverride ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), typeof(Connect).Assembly.GetName().Name, "单体预设", feature); }
        internal static string FileFor(string feature, string name)
        {
            name = (name ?? "").Trim();
            if (name.Length == 0 || name.Length > 80 || name == "." || name == ".." || name.EndsWith(".") || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                throw new InvalidOperationException("预设名称须为 1–80 个有效文件名字符，不能以点结尾。");
            CompanyWorkbookRules.FileName(name + ".json");
            return Path.Combine(Library(feature), name + ".xybatch.json");
        }
        internal static BatchPreset Read(string path, string feature)
        {
            if (new FileInfo(path).Length > 1024 * 1024) throw new InvalidDataException("预设文件超过 1 MB，无法读取。");
            BatchPreset value;
            using (var stream = File.OpenRead(path))
            {
                if (stream.Length > 1024 * 1024) throw new InvalidDataException("预设文件超过 1 MB，无法读取。");
                value = (BatchPreset)new DataContractJsonSerializer(typeof(BatchPreset)).ReadObject(stream);
            }
            Validate(value, feature); return value;
        }
        internal static void Validate(BatchPreset value, string feature)
        {
            if (value == null || value.Format != "xycabin.batch-settings" || value.Feature != feature) throw new InvalidDataException("预设不属于当前功能。");
            if (value.Version != 1) throw new InvalidDataException("不支持此预设版本，请检查文件或升级加载项。");
            FileFor(feature, value.Name);
            if (value.Values == null || value.Values.Count > 80 || value.Values.Any(p => p.Key == null || p.Value == null || p.Value.Length > 32768))
                throw new InvalidDataException("预设参数无效。");
        }
        internal static void Write(string path, BatchPreset value, bool overwrite)
        {
            Validate(value, value.Feature);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            // Reuse the customization store's lock + same-directory replacement
            // discipline. A failed serialization/replacement retains the old file.
            using (var guard = new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                if (File.Exists(path) && !overwrite) throw new IOException("已存在同名预设，请换一个名称。");
                string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write)) new DataContractJsonSerializer(typeof(BatchPreset)).WriteObject(stream, value);
                    if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
                }
                finally { if (File.Exists(temporary)) File.Delete(temporary); }
            }
        }
        internal static Dictionary<string, string> Capture(IDictionary<string, Control> fields)
        {
            var result = new Dictionary<string, string>();
            foreach (var item in fields)
            {
                var number = item.Value as NumericUpDown; var check = item.Value as CheckBox;
                result[item.Key] = number != null ? number.Value.ToString(CultureInfo.InvariantCulture) : check != null ? check.Checked.ToString() : item.Value.Text;
            }
            return result;
        }
        internal static void ValidateFields(IDictionary<string, Control> fields, Dictionary<string, string> values)
        {
            if (values.Count != fields.Count || fields.Keys.Any(k => !values.ContainsKey(k))) throw new InvalidDataException("预设参数不完整，请重新保存预设。");
            foreach (var item in fields)
            {
                decimal n; bool b;
                var number = item.Value as NumericUpDown; var check = item.Value as CheckBox;
                if (number != null && (!decimal.TryParse(values[item.Key], NumberStyles.None, CultureInfo.InvariantCulture, out n) || n != decimal.Truncate(n) || n < number.Minimum || n > number.Maximum)) throw new InvalidDataException("预设行号或年度超出范围：" + item.Key);
                if (check != null && !bool.TryParse(values[item.Key], out b)) throw new InvalidDataException("预设复选选项无效：" + item.Key);
            }
        }
        internal static void Apply(IDictionary<string, Control> fields, Dictionary<string, string> values)
        {
            ValidateFields(fields, values);
            foreach (var item in fields)
            {
                var number = item.Value as NumericUpDown; var check = item.Value as CheckBox;
                if (number != null) { number.Value = decimal.Parse(values[item.Key], CultureInfo.InvariantCulture); number.Text = values[item.Key]; }
                else if (check != null) check.Checked = bool.Parse(values[item.Key]);
                else item.Value.Text = values[item.Key];
            }
        }
    }
    internal sealed class BatchPresetBar : FlowLayoutPanel
    {
        private readonly ComboBox names = new ComboBox { Width = 160, DropDownStyle = ComboBoxStyle.DropDown, Name = "batchPresetName" };
        private readonly string feature;
        private readonly Func<BatchPreset> capture;
        private readonly Action<BatchPreset> apply;
        private readonly Func<bool> available;
        private readonly Action<string> message;
        internal BatchPresetBar(string feature, Func<BatchPreset> capture, Action<BatchPreset> apply, Func<bool> available, Action<string> message)
        {
            this.feature = feature; this.capture = capture; this.apply = apply; this.available = available; this.message = message;
            AutoSize = true; WrapContents = true; Dock = DockStyle.Top;
            Controls.Add(names);
            Add("保存预设", Save); Add("应用预设", Load); Add("导入…", Import); Add("导出…", Export); Add("删除", Delete);
            Controls.Add(new Label { AutoSize = true, Text = "仅保存参数；先选择工作簿，应用后重新预检。", Margin = new Padding(3, 8, 3, 3) });
            // No disk reads in settings previews or isolated constructor probes.
            if (!DialogPreviewFactory.IsConstructing) try { RefreshNames(); } catch (Exception error) { names.AccessibleDescription = "预设库暂不可读：" + error.Message; }
        }
        private void Add(string text, Action action)
        {
            var button = Ui.Button(text, 90); Controls.Add(button);
            button.Click += delegate
            {
                if (!available()) { message("当前正在执行任务或有待确认的还原批次，不能操作预设。"); return; }
                try { action(); } catch (Exception error) { message("预设操作失败：" + error.Message); }
            };
        }
        private bool Ask(string text)
        { return UiDialog.Show(FindForm(), text, "单体功能预设", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes; }
        private void RefreshNames()
        {
            string prior = names.Text; names.Items.Clear(); string library = BatchPresetStore.Library(feature);
            if (Directory.Exists(library)) foreach (var file in Directory.GetFiles(library, "*.xybatch.json").OrderBy(p => p)) names.Items.Add(Path.GetFileName(file).Replace(".xybatch.json", ""));
            names.Text = prior;
        }
        private BatchPreset Current()
        {
            var value = capture(); value.Feature = feature; value.Name = names.Text.Trim(); value.Saved = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
            BatchPresetStore.Validate(value, feature); return value;
        }
        private void Save()
        {
            FindForm().ValidateChildren(); var value = Current(); string file = BatchPresetStore.FileFor(feature, value.Name);
            bool exists = File.Exists(file); if (exists && !Ask("覆盖同名预设“" + value.Name + "”？")) return;
            BatchPresetStore.Write(file, value, exists); RefreshNames(); message("已保存预设“" + value.Name + "”（仅参数，不含工作表数据）。");
        }
        private void Load()
        {
            var value = BatchPresetStore.Read(BatchPresetStore.FileFor(feature, names.Text), feature);
            if (!Ask("应用预设“" + value.Name + "”将替换当前配置参数，需重新预检。\n保存时的文件：" + value.Context + "\n请先保存需要保留的当前设置。是否继续？")) return;
            apply(value); message("已应用预设“" + value.Name + "”；请核对文件、表头和字段，重新预检。");
        }
        private void Import()
        {
            var files = HostDialogService.OpenFiles(FindForm(), "单体功能预设 (*.xybatch.json)|*.xybatch.json");
            if (files.Length == 0) return; if (files.Length != 1) throw new InvalidOperationException("请每次导入一个预设。");
            var value = BatchPresetStore.Read(files[0], feature); string file = BatchPresetStore.FileFor(feature, value.Name);
            bool exists = File.Exists(file); if (exists && !Ask("导入将覆盖同名预设“" + value.Name + "”，是否继续？")) return;
            BatchPresetStore.Write(file, value, exists); names.Text = value.Name; RefreshNames(); message("预设已导入，请点击“应用预设”加载。");
        }
        private void Export()
        {
            var value = BatchPresetStore.Read(BatchPresetStore.FileFor(feature, names.Text), feature);
            string filename = value.Name + ".xybatch.json";
            string file = HostDialogService.SaveSettings(FindForm(), filename);
            if (file == null) return; BatchPresetStore.Write(file, value, true); message("已导出预设：" + file);
        }
        private void Delete()
        {
            string file = BatchPresetStore.FileFor(feature, names.Text);
            if (!File.Exists(file) || !Ask("删除预设“" + names.Text + "”？此操作只删除本机预设文件。")) return;
            File.Delete(file); names.Text = ""; RefreshNames(); message("预设已删除。");
        }
    }
}
