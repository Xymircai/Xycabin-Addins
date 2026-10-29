using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    internal sealed class PdfSource
    {
        internal string Path, Folder;
        public override string ToString() { return System.IO.Path.GetFileName(Path) + "  |  " + System.IO.Path.GetExtension(Path) + "  |  " + System.IO.Path.GetDirectoryName(Path); }
    }

    // Source identities and ordering survive filtering. This class never changes source files.
    internal sealed class PdfFileSelection
    {
        internal readonly List<PdfSource> Sources = new List<PdfSource>();
        internal readonly List<string> Inputs = new List<string>();
        internal readonly HashSet<string> Removed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<string> OutputFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal readonly HashSet<string> Generated = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> roots = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        internal static readonly string[] ConvertExtensions = ".doc,.docx,.rtf,.odt,.wps,.xls,.xlsx,.xlsm,.xlsb,.ods,.et,.png,.jpg,.jpeg,.bmp,.tif,.tiff,.gif".Split(',');
        internal static bool Supported(string path, bool convert)
        { return (convert ? ConvertExtensions : new[] { ".pdf" }).Contains(System.IO.Path.GetExtension(path), StringComparer.OrdinalIgnoreCase); }
        internal void AddInput(string path)
        {
            path = System.IO.Path.GetFullPath(path);
            if (!Inputs.Contains(path, StringComparer.OrdinalIgnoreCase)) Inputs.Add(path);
            Removed.Remove(path);
        }
        private string RootName(string path)
        {
            string alias;
            if (roots.TryGetValue(path, out alias)) return alias;
            string basis = new DirectoryInfo(path).Name.TrimEnd(':');
            if (basis.Length == 0) basis = "来源";
            alias = basis;
            for (int i = 2; roots.Values.Contains(alias, StringComparer.OrdinalIgnoreCase); i++) alias = basis + " (" + i + ")";
            roots[path] = alias;
            return alias;
        }
        internal List<PdfSource> Scan(bool recursive, Func<bool> cancelled, List<string> warnings)
        {
            var found = new List<PdfSource>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string input in Inputs)
            {
                if (cancelled()) throw new OperationCanceledException();
                if (File.Exists(input)) { AddFound(found, seen, input, ""); continue; }
                string alias = RootName(input);
                var pending = new Stack<string>(); pending.Push(input);
                while (pending.Count > 0)
                {
                    if (cancelled()) throw new OperationCanceledException();
                    string folder = pending.Pop();
                    try
                    {
                        if ((File.GetAttributes(folder) & FileAttributes.ReparsePoint) != 0) { warnings.Add("跳过目录联接：" + folder); continue; }
                        if (OutputFolders.Any(o => !string.Equals(o, input, StringComparison.OrdinalIgnoreCase) && IsWithin(o, input) && IsWithin(folder, o))) continue;
                        foreach (string file in Directory.GetFiles(folder).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
                        {
                            if (cancelled()) throw new OperationCanceledException();
                            string relative = folder.Substring(input.TrimEnd('\\').Length).TrimStart('\\');
                            AddFound(found, seen, file, relative.Length == 0 ? alias : System.IO.Path.Combine(alias, relative));
                        }
                        if (recursive) foreach (string child in Directory.GetDirectories(folder).OrderByDescending(p => p, StringComparer.OrdinalIgnoreCase)) pending.Push(child);
                    }
                    catch (UnauthorizedAccessException e) { warnings.Add(folder + "：" + e.Message); }
                    catch (IOException e) { warnings.Add(folder + "：" + e.Message); }
                }
            }
            return found;
        }
        private void AddFound(List<PdfSource> found, HashSet<string> seen, string path, string folder)
        {
            if (!Removed.Contains(path) && !Generated.Contains(path) && !IsTemporary(path) && seen.Add(path))
                found.Add(new PdfSource { Path = path, Folder = folder });
        }
        private static bool IsTemporary(string path)
        {
            string name = System.IO.Path.GetFileNameWithoutExtension(path); Guid id;
            return name.StartsWith(".pdf-", StringComparison.OrdinalIgnoreCase) && Guid.TryParseExact(name.Substring(5), "N", out id);
        }
        internal static bool IsWithin(string path, string root)
        { return string.Equals(path.TrimEnd('\\'), root.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase) || path.StartsWith(root.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase); }
        internal void Commit(List<PdfSource> scanned)
        {
            var byPath = scanned.ToDictionary(s => s.Path, StringComparer.OrdinalIgnoreCase);
            var ordered = new List<PdfSource>();
            foreach (var source in Sources) { PdfSource current; if (byPath.TryGetValue(source.Path, out current)) { ordered.Add(current); byPath.Remove(source.Path); } }
            ordered.AddRange(scanned.Where(s => byPath.ContainsKey(s.Path)));
            Sources.Clear(); Sources.AddRange(ordered);
        }
        internal void Clear() { Sources.Clear(); Inputs.Clear(); Removed.Clear(); roots.Clear(); }
        internal void Remove(PdfSource source) { Sources.Remove(source); Removed.Add(source.Path); }
        internal PdfSource[] Tasks(HashSet<string> extensions, bool convert)
        { return Sources.Where(s => Supported(s.Path, convert) && extensions.Contains(System.IO.Path.GetExtension(s.Path))).ToArray(); }
    }
}
