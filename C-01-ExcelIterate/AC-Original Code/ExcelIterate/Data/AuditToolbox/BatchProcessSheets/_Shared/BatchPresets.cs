using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
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
        internal static string Encode(BatchPreset value)
        { using (var stream = new MemoryStream()) { new DataContractJsonSerializer(typeof(BatchPreset)).WriteObject(stream, value); return Encoding.UTF8.GetString(stream.ToArray()); } }
        internal static BatchPreset Decode(string json, string feature)
        {
            if (json == null || Encoding.UTF8.GetByteCount(json) > 1024 * 1024) throw new InvalidDataException("预设超过大小限制。");
            BatchPreset value;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json))) value = (BatchPreset)new DataContractJsonSerializer(typeof(BatchPreset)).ReadObject(stream);
            Validate(value, feature); return value;
        }
        internal static string UniqueName(string feature, string name)
        { string candidate = name; int suffix = 2; while (File.Exists(FileFor(feature, candidate))) candidate = name.Substring(0, Math.Min(name.Length, 70)) + " (" + suffix++ + ")"; return candidate; }
        internal static string Import(string file, string feature)
        {
            var value = Read(file, feature); value.Name = UniqueName(feature, value.Name);
            string target = FileFor(feature, value.Name); Write(target, value, false); return target;
        }
        internal static string Rename(string file, string feature, string name)
        {
            string target = FileFor(feature, name); var value = Read(file, feature); value.Name = name.Trim();
            if (string.Equals(target, file, StringComparison.OrdinalIgnoreCase)) { Write(file, value, true); return file; }
            using (var guard = new FileStream(file + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            {
                Write(target, value, false);
                try { File.Delete(file); } catch { try { File.Delete(target); } catch { } throw; }
            }
            return target;
        }
        internal static void Delete(string file)
        { using (var guard = new FileStream(file + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)) File.Delete(file); }
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
}
