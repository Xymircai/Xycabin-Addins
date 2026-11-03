using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace WpsShared
{
    // Diagnostic-only source access. Runtime fixture/output paths keep normal IO.
    // Explicit read aliases join split files; directory scans enumerate compiled files.
    internal static class SourcePaths
    {
        [DataContract]
        private sealed class Manifest
        {
            [DataMember(Name = "sources")] public string[] Sources;
            [DataMember(Name = "source_aliases")] public Dictionary<string, string[]> Aliases;
        }

        private static string ProductRoot(string path)
        {
            for (string current = path; !string.IsNullOrEmpty(current); current = Path.GetDirectoryName(current))
                if (File.Exists(Path.Combine(current, "project.json"))
                    && (Path.GetFileName(current) == "C-01-ExcelIterate" || Path.GetFileName(current) == "C-02-WordUpdater"))
                    return current;
            return null;
        }

        private static Manifest ReadManifest(string root)
        {
            var serializer = new DataContractJsonSerializer(typeof(Manifest),
                new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
            // project.json 可能带 UTF-8 BOM（上游提交即如此）。DataContractJsonSerializer
            // 会把 BOM 当正文而抛 XmlException，故先剥掉 BOM 再反序列化；
            // 与 Python 侧统一按 utf-8-sig 读取的口径保持一致。
            byte[] bytes = File.ReadAllBytes(Path.Combine(root, "project.json"));
            int offset = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
            using (var stream = new MemoryStream(bytes, offset, bytes.Length - offset, false))
                return (Manifest)serializer.ReadObject(stream);
        }

        private static string Physical(string root, string reference)
        {
            string relative = reference.Replace('\\', '/');
            string basis = root;
            if (relative.StartsWith("@shared/", StringComparison.Ordinal))
            {
                relative = relative.Substring(8);
                basis = Path.Combine(Path.GetDirectoryName(root), "C-00-Shared");
            }
            if (Path.IsPathRooted(relative) || relative.Split('/').Contains("..") || relative.Contains(":"))
                throw new InvalidDataException("Source path escapes allowed roots: " + reference);
            string full = Path.GetFullPath(Path.Combine(basis, relative));
            if (!full.StartsWith(basis + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) || !File.Exists(full))
                throw new FileNotFoundException("Missing compiled source: " + reference, full);
            return full;
        }

        private static string[] ReadTargets(string path)
        {
            string full = Path.GetFullPath(path);
            string root = ProductRoot(full);
            if (root == null) return new[] { full };
            string legacy = Path.Combine(root, "src", "addin") + Path.DirectorySeparatorChar;
            if (!full.StartsWith(legacy, StringComparison.OrdinalIgnoreCase)) return new[] { full };
            var manifest = ReadManifest(root);
            string name = Path.GetFileName(full);
            string[] references;
            if (manifest.Aliases != null && manifest.Aliases.TryGetValue(name, out references))
            {
                if (references.Length == 0 || references.Any(r => !manifest.Sources.Contains(r)))
                    throw new InvalidDataException("Source read alias must refer to compiled sources: " + name);
                return references.Select(r => Physical(root, r)).ToArray();
            }
            var matches = manifest.Sources.Where(r => Path.GetFileName(r) == name).ToArray();
            if (matches.Length > 1) throw new InvalidDataException("Ambiguous source name: " + name);
            if (matches.Length == 1) return new[] { Physical(root, matches[0]) };
            string common = Path.Combine(Path.GetDirectoryName(root), "C-00-Shared", "src", "addin", name);
            return new[] { !File.Exists(full) && File.Exists(common) ? common : full };
        }

        internal static string Resolve(string path)
        {
            string[] targets = ReadTargets(path);
            if (targets.Length != 1) throw new InvalidDataException("Split source requires ReadAllText: " + path);
            return targets[0];
        }
        internal static bool Exists(string path) { return ReadTargets(path).All(File.Exists); }
        internal static string ReadAllText(string path) { return ReadAllText(path, Encoding.UTF8); }
        internal static string ReadAllText(string path, Encoding encoding)
        { return string.Join("\n", ReadTargets(path).Select(p => File.ReadAllText(p, encoding))); }
        internal static string[] ReadAllLines(string path)
        { return ReadAllText(path).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'); }
        internal static string[] GetFiles(string path, string pattern, SearchOption option)
        {
            string full = Path.GetFullPath(path);
            string root = ProductRoot(full);
            if (root != null && full.TrimEnd(Path.DirectorySeparatorChar).Equals(Path.Combine(root, "src", "addin"), StringComparison.OrdinalIgnoreCase))
            {
                var manifest = ReadManifest(root);
                string expression = "^" + Regex.Escape(pattern).Replace(@"\*", ".*").Replace(@"\?", ".") + "$";
                return manifest.Sources.Select(r => Physical(root, r))
                    .Where(p => Regex.IsMatch(Path.GetFileName(p), expression, RegexOptions.IgnoreCase))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
            }
            return Directory.Exists(full) ? Directory.GetFiles(full, pattern, option) : new string[0];
        }
        internal static string[] GetFiles(string path, string pattern) { return GetFiles(path, pattern, SearchOption.TopDirectoryOnly); }
        internal static string[] GetFiles(string path) { return GetFiles(path, "*"); }
    }
}
