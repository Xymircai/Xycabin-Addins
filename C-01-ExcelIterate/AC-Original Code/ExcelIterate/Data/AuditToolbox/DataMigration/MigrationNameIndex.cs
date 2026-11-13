using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    internal sealed class MigrationNameIndex
    {
        private readonly Dictionary<string, List<string>> names = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        private static readonly string[] Empty = new string[0];
        internal readonly List<string> Paths;
        internal readonly string Signature;
        internal MigrationNameIndex(IEnumerable<string> paths)
        {
            Paths = new List<string>(paths);
            Signature = string.Join("|", Paths.ToArray());
            foreach (string path in Paths)
            {
                string key = NormalizeName(Path.GetFileName(path)); List<string> matches;
                if (!names.TryGetValue(key, out matches)) names.Add(key, matches = new List<string>());
                matches.Add(path);
            }
        }
        internal static string NormalizeName(string name)
        {
            string text = (name ?? "").Trim();
            foreach (string extension in new[] { ".xlsx", ".xlsm", ".xls", ".et" })
                if (text.EndsWith(extension, StringComparison.OrdinalIgnoreCase)) { text = text.Substring(0, text.Length - extension.Length); break; }
            text = text.Normalize(NormalizationForm.FormKC);
            return Regex.Replace(Regex.Replace(text, @"^[\s!]*(?:\d+[\s\-_.、]+)*", ""), @"[\s!]", "").ToLowerInvariant();
        }
        internal IList<string> Find(string name)
        {
            string key = NormalizeName(name); List<string> matches;
            return key.Length > 0 && names.TryGetValue(key, out matches) ? (IList<string>)matches : Empty;
        }
        internal void Match(MigrationTriggerRow row)
        {
            row.Candidates.Clear(); row.Path = "";
            foreach (string path in Find(row.Name)) row.Candidates.Add(new KeyValuePair<string, double>(path, 1));
            row.Kind = row.Candidates.Count == 0 ? "缺失" : row.Candidates.Count == 1 ? "精确" : "多个候选";
            if (row.Candidates.Count == 1) row.Path = row.Candidates[0].Key;
        }
    }
}
