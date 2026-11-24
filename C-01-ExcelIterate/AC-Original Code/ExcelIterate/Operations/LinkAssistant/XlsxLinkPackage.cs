using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Xml;

namespace Excel_Iterate
{
    internal sealed class ExternalLinkInfo
    {
        internal string OwnerPath, SourcePath, RawTarget, Part, Relationship;
        internal string NewPath, Status = "待填写新链接";
        internal bool BreakRequested, IsMissingPath;
        internal bool DestinationMissing;
        internal ExternalLinkInfo ValidatedSource;
        internal string SourceId { get { return SourcePath ?? RawTarget; } }
        internal string Key { get { return OwnerPath + "\n" + SourceId; } }
        internal string UnavailableReason
        {
            get
            {
                return SourcePath == null && !IsMissingPath ? "不支持的网页或非文件链接" : null;
            }
        }
    }
    internal static class XlsxLinkPackage
    {
        private const string Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal const string MissingPathSuffix = "/xlExternalLinkPath/xlPathMissing";
        private const string MissingPathStatus = "来源路径待核对（xlPathMissing），可填写新链接";
        internal static bool IsFullWindowsPath(string path)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(path, @"^(?:[A-Za-z]:[\\/]|\\\\[^\\/]+[\\/][^\\/]+)");
        }
        internal static string AbsoluteFileTarget(string target)
        {
            if (string.IsNullOrWhiteSpace(target)) return null;
            Uri uri;
            if (target.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            {
                if (!Uri.TryCreate(target, UriKind.Absolute, out uri) || !uri.IsFile) return null;
                target = uri.LocalPath;
            }
            return IsFullWindowsPath(target) ? Path.GetFullPath(target) : null;
        }
        internal static string WpsFileUri(string file)
        {
            string escaped = new Uri(Path.GetFullPath(file)).AbsoluteUri;
            // Decode only UTF-8 non-ASCII byte sequences. Keep ASCII escapes
            // such as %23, %25 and %20: SafeUnescaped would expose # as a URI
            // fragment and lose literal percent sequences in a real filename.
            return System.Text.RegularExpressions.Regex.Replace(escaped, "(?:%[0-9A-Fa-f]{2})+", delegate(System.Text.RegularExpressions.Match match)
            {
                List<byte> bytes = new List<byte>(); for (int i = 0; i < match.Value.Length; i += 3) bytes.Add(Convert.ToByte(match.Value.Substring(i + 1, 2), 16));
                System.Text.StringBuilder text = new System.Text.StringBuilder();
                for (int i = 0; i < bytes.Count; )
                {
                    if (bytes[i] < 128) { text.Append('%').Append(bytes[i++].ToString("X2")); continue; }
                    int end = i; while (end < bytes.Count && bytes[end] >= 128) end++;
                    text.Append(new System.Text.UTF8Encoding(false, true).GetString(bytes.GetRange(i, end - i).ToArray())); i = end;
                }
                return text.ToString();
            });
        }
        internal static string Resolve(string owner, string target)
        {
            if (string.IsNullOrWhiteSpace(target)) return null;
            Uri uri;
            if (Uri.TryCreate(target, UriKind.Absolute, out uri))
            {
                if (!uri.IsFile) return null;
                return Path.GetFullPath(uri.LocalPath);
            }
            // Package relationship targets are URI references; decode here,
            // not on arbitrary user-entered Windows filenames containing '%'.
            string decoded = Uri.UnescapeDataString(target).Replace('/', Path.DirectorySeparatorChar);
            if (decoded.StartsWith("\\", StringComparison.Ordinal) && !decoded.StartsWith("\\\\", StringComparison.Ordinal))
                return Path.GetFullPath(Path.Combine(Path.GetPathRoot(owner), decoded.TrimStart('\\')));
            return Path.GetFullPath(Path.Combine(Path.GetDirectoryName(owner), decoded));
        }
        internal static XmlDocument Load(Stream stream)
        {
            XmlDocument xml = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null })) xml.Load(reader);
            return xml;
        }
        internal static List<ExternalLinkInfo> Scan(string file)
        {
            return ScanRelationships(file, false);
        }
        internal static List<ExternalLinkInfo> ScanReport(string file)
        {
            return ScanReportParts(file)
                .GroupBy(link => link.Key, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.FirstOrDefault(link => !link.IsMissingPath) ?? group.First())
                .ToList();
        }
        internal static List<ExternalLinkInfo> ScanReportParts(string file)
        {
            return ScanRelationships(file, true);
        }
        private static List<ExternalLinkInfo> ScanRelationships(string file, bool includeMissingPaths)
        {
            string owner = Path.GetFullPath(file);
            if (!string.Equals(Path.GetExtension(owner), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("本轮仅支持 .xlsx 工作簿。");
            List<ExternalLinkInfo> result = new List<ExternalLinkInfo>();
            using (FileStream stream = new FileStream(owner, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                if (zip.GetEntry("xl/workbook.xml") == null) throw new InvalidOperationException("不是有效的 xlsx 工作簿。");
                if (zip.Entries.Any(e => e.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase))) throw new InvalidOperationException("工作簿含文档数字签名，请先移除文档签名后再更换链接。");
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (!entry.FullName.StartsWith("xl/externalLinks/_rels/", StringComparison.Ordinal) || !entry.FullName.EndsWith(".rels", StringComparison.Ordinal)) continue;
                    XmlDocument xml; using (Stream input = entry.Open()) xml = Load(input);
                    foreach (XmlElement rel in xml.DocumentElement.ChildNodes.OfType<XmlElement>())
                    {
                        string type = rel.GetAttribute("Type");
                        bool missingPath = includeMissingPaths && type.EndsWith(MissingPathSuffix, StringComparison.Ordinal);
                        if (!missingPath && !type.EndsWith("/externalLinkPath", StringComparison.Ordinal)) continue;
                        string target = rel.GetAttribute("Target");
                        // xlPathMissing supplies no base path. Preserve the recorded target for display.
                        var link = new ExternalLinkInfo
                        {
                            OwnerPath = owner,
                            SourcePath = missingPath ? null : Resolve(owner, target),
                            RawTarget = target,
                            Part = entry.FullName,
                            Relationship = rel.GetAttribute("Id"),
                            IsMissingPath = missingPath
                        };
                        link.Status = missingPath ? MissingPathStatus : link.UnavailableReason ?? link.Status;
                        result.Add(link);
                    }
                }
            }
            return result;
        }
        internal static string Hash(string file)
        { using (SHA256 hash = SHA256.Create()) using (FileStream input = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) return BitConverter.ToString(hash.ComputeHash(input)).Replace("-", ""); }
        internal static Dictionary<string, string> EntryHashes(string file)
        {
            Dictionary<string, string> hashes = new Dictionary<string, string>(StringComparer.Ordinal);
            using (FileStream input = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (ZipArchive zip = new ZipArchive(input, ZipArchiveMode.Read))
                foreach (ZipArchiveEntry entry in zip.Entries)
                    using (SHA256 hash = SHA256.Create()) using (Stream data = entry.Open()) hashes.Add(entry.FullName, BitConverter.ToString(hash.ComputeHash(data)));
            return hashes;
        }
        internal static void Rewrite(string stage, string owner, IDictionary<string, string> replacements)
        {
            HashSet<string> changed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (FileStream file = new FileStream(stage, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update))
            {
                foreach (ZipArchiveEntry entry in zip.Entries.ToArray())
                {
                    if (!entry.FullName.StartsWith("xl/externalLinks/_rels/", StringComparison.Ordinal) || !entry.FullName.EndsWith(".rels", StringComparison.Ordinal)) continue;
                    XmlDocument xml; using (Stream input = entry.Open()) xml = Load(input); bool edited = false;
                    foreach (XmlElement rel in xml.DocumentElement.ChildNodes.OfType<XmlElement>())
                    {
                        if (!rel.GetAttribute("Type").EndsWith("/externalLinkPath", StringComparison.Ordinal)) continue;
                        string key = Resolve(owner, rel.GetAttribute("Target")), replacement;
                        if (key != null && replacements.TryGetValue(key, out replacement))
                        {
                            // All relationships are mapped against the ORIGINAL
                            // paths: A->B and B->C do not accidentally become A->C.
                            // WPS decodes percent-encoded non-ASCII file URI bytes
                            // as Latin-1. Match its own OOXML convention: preserve
                            // Unicode while retaining escapes for reserved chars.
                            rel.SetAttribute("Target", WpsFileUri(replacement));
                            rel.SetAttribute("TargetMode", "External"); edited = true; changed.Add(key);
                        }
                    }
                    if (edited)
                    {
                        string name = entry.FullName; entry.Delete(); ZipArchiveEntry updated = zip.CreateEntry(name, CompressionLevel.Optimal);
                        using (Stream output = updated.Open()) xml.Save(output);
                    }
                }
            }
            if (changed.Count != replacements.Count) throw new InvalidOperationException("外部链接已变化，请重新扫描后填写映射。");
        }
        internal static void VerifyRewrite(string original, string stage, IDictionary<string, string> replacements)
        {
            Dictionary<string, string> before = EntryHashes(original), after = EntryHashes(stage);
            if (before.Count != after.Count || before.Keys.Any(k => !after.ContainsKey(k))) throw new InvalidOperationException("副本的文件结构核对失败。");
            HashSet<string> changedParts = new HashSet<string>(Scan(original).Where(l => l.SourcePath != null && replacements.ContainsKey(l.SourcePath)).Select(l => l.Part), StringComparer.Ordinal);
            foreach (string part in before.Keys) if (!changedParts.Contains(part) && before[part] != after[part]) throw new InvalidOperationException("非链接内容发生变化：" + part);
            List<ExternalLinkInfo> expected = Scan(original), actual = Scan(stage);
            if (expected.Count != actual.Count) throw new InvalidOperationException("副本链接数量不一致。");
            foreach (ExternalLinkInfo old in expected)
            {
                ExternalLinkInfo now = actual.Single(l => l.Part == old.Part && l.Relationship == old.Relationship);
                string replacement;
                if (old.SourcePath != null && replacements.TryGetValue(old.SourcePath, out replacement))
                { if (!string.Equals(now.SourcePath, replacement, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("新链接核对失败。"); }
                else if (now.RawTarget != old.RawTarget) throw new InvalidOperationException("未勾选链接发生变化。");
            }
        }
        internal static string Prepare(string original, IDictionary<string, string> replacements)
        {
            foreach (var item in replacements)
            {
                if (!File.Exists(item.Value) || !string.Equals(Path.GetExtension(item.Value), ".xlsx", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("新链接必须是存在的 xlsx 文件：" + item.Value);
                if (string.Equals(original, item.Value, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("新链接不能指向目标工作簿自身。");
                Scan(item.Value); // Validate real OOXML, not merely the extension.
            }
            string stage = Path.Combine(Path.GetDirectoryName(original), ".EI_LinkStage_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try { File.Copy(original, stage); File.SetAttributes(stage, FileAttributes.Normal); Rewrite(stage, original, replacements); VerifyRewrite(original, stage, replacements); return stage; }
            catch { if (File.Exists(stage)) File.Delete(stage); throw; }
        }
        internal static string Commit(string original, string stage, string originalHash)
        { return CommitRecorded(original, stage, originalHash, null); }
        internal static string CommitRecorded(string original, string stage, string originalHash, Action<string> backupReady)
        {
            if (Hash(original) != originalHash) throw new InvalidOperationException("原文件在执行期间发生变化，已停止替换。");
            string folder = Path.Combine(Path.GetDirectoryName(original), "ExcelIterate链接备份"); Directory.CreateDirectory(folder);
            string backup = Path.Combine(folder, Path.GetFileNameWithoutExtension(original) + "-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + "-" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(original, backup, false);
            if (Hash(backup) != originalHash || Hash(original) != originalHash) throw new InvalidOperationException("备份或原文件核对失败，未替换原文件。备份：" + backup);
            if (backupReady != null) backupReady(backup);
            File.Replace(stage, original, null);
            return backup;
        }
    }
}
