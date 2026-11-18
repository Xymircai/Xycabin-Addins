using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

internal static class PayloadProbe
{
    private static string CorePage(string text)
    {
        const string pattern="(<script id=\"installer-assets\" type=\"application/json\">)(.*?)(</script>)";
        if(Regex.Matches(text,pattern,RegexOptions.Singleline).Count!=1)throw new Exception("Unique installer data area required");
        return Regex.Replace(text,pattern,m=>m.Groups[1].Value+"{\"packages\":{}}"+m.Groups[3].Value,RegexOptions.Singleline);
    }

    private static string Hash(byte[] bytes)
    { using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", ""); }
    private static byte[] Resource(Assembly assembly, string name)
    {
        using (Stream input = assembly.GetManifestResourceStream(name))
        using (var buffer = new MemoryStream())
        {
            if (input == null) throw new Exception("Embedded resource missing: " + name);
            input.CopyTo(buffer); return buffer.ToArray();
        }
    }
    private static void VerifyDocumentAssets(string root, Assembly assembly, string expectedManifest)
    {
        root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        byte[] manifest = Resource(assembly, "WpsShared.Agent.Assets.xml");
        if (Hash(manifest) != Hash(File.ReadAllBytes(expectedManifest)))
            throw new Exception("Embedded document assets differ from the current helper-generated manifest");
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        var xml = new XmlDocument { XmlResolver = null };
        using (var input = new MemoryStream(manifest))
        using (XmlReader reader = XmlReader.Create(input, settings)) xml.Load(reader);
        XmlElement top = xml.DocumentElement;
        if (top.Name != "documentAssets" || top.GetAttribute("version") != "1" || top.Attributes.Count != 1)
            throw new Exception("Invalid document asset manifest root");
        var documents = new HashSet<string>(StringComparer.Ordinal);
        var assets = new HashSet<string>(StringComparer.Ordinal);
        foreach (XmlNode child in top.ChildNodes)
        {
            XmlElement doc = child as XmlElement;
            if (doc == null || doc.Name != "document" || doc.Attributes.Count != 4) throw new Exception("Invalid document asset manifest document");
            string name = doc.GetAttribute("name"), resource = doc.GetAttribute("resource");
            string filename = name == "DevExperience" ? "B2-DevExperience.md" : name == "UserGuide" ? "A-UserGuide.md" : null;
            if (filename == null || !documents.Add(name) || resource != "WpsShared.Agent." + name + ".md"
                || doc.GetAttribute("path") != "C-04-Devlog/" + filename
                || !Regex.IsMatch(doc.GetAttribute("sha256"), "^[0-9a-f]{64}$")
                || !String.Equals(Hash(Resource(assembly, resource)), doc.GetAttribute("sha256"), StringComparison.OrdinalIgnoreCase))
                throw new Exception("Invalid document asset manifest identity/hash");
            var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode entry in doc.ChildNodes)
            {
                XmlElement asset = entry as XmlElement;
                if (asset == null || asset.Name != "asset" || asset.Attributes.Count != 4 || asset.HasChildNodes)
                    throw new Exception("Invalid document asset entry");
                string path = asset.GetAttribute("path"), digest = asset.GetAttribute("sha256"), assetResource = asset.GetAttribute("resource");
                string[] parts = path.Split('/'); string extension = Path.GetExtension(path).ToLowerInvariant();
                bool allowed = parts.Length == 3 && parts[1] == "AB-Graphic" && (extension == ".png" || extension == ".svg")
                    && (parts[0] == "C-01-ExcelIterate" || parts[0] == "C-02-WordUpdater" || parts[0] == "C-03-Website·Xycabin");
                if (!allowed || path.IndexOf('\\') >= 0 || path.IndexOf(':') >= 0 || !paths.Add(path)
                    || !Regex.IsMatch(digest, "^[0-9a-f]{64}$") || assetResource != "WpsShared.Agent.Asset." + digest + extension)
                    throw new Exception("Invalid document asset path/resource");
                foreach (string part in parts) if (part == "." || part == ".." || part.Length == 0) throw new Exception("Document asset traversal");
                string repo = Path.GetDirectoryName(Path.GetFullPath(root));
                string source = Path.GetFullPath(Path.Combine(repo, path.Replace('/', Path.DirectorySeparatorChar)));
                string boundary = repo.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (!source.StartsWith(boundary, StringComparison.OrdinalIgnoreCase)) throw new Exception("Document asset escapes the repository");
                for (string cursor = source; !String.Equals(cursor, repo, StringComparison.OrdinalIgnoreCase); cursor = Path.GetDirectoryName(cursor))
                    if ((File.GetAttributes(cursor) & FileAttributes.ReparsePoint) != 0) throw new Exception("Document asset reparse point");
                long size;
                byte[] bytes = Resource(assembly, assetResource), current = File.ReadAllBytes(source);
                if (!Int64.TryParse(asset.GetAttribute("size"), out size) || size <= 0 || bytes.LongLength != size || current.LongLength != size
                    || !String.Equals(Hash(bytes), digest, StringComparison.OrdinalIgnoreCase) || Hash(bytes) != Hash(current))
                    throw new Exception("Stale embedded document asset: " + path);
                assets.Add(assetResource);
            }
        }
        if (documents.Count != 2) throw new Exception("Both document asset manifests required");
        foreach (string name in assembly.GetManifestResourceNames())
            if (name.StartsWith("WpsShared.Agent.Asset.", StringComparison.Ordinal) && !assets.Contains(name))
                throw new Exception("Orphan embedded document asset: " + name);
        Console.WriteLine("PASS: two document hashes and " + assets.Count + " exact PNG/SVG resources against current source and helper manifest");
    }
    // rule=R-24
    public static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            if (args.Length != 6) throw new ArgumentException("Expected: project-root assembly installer site-html version helper-generated-assets-xml");
            string root = Path.GetFullPath(args[0]), name = args[1], installer = Path.GetFullPath(args[2]);
            string sitePath = Path.GetFullPath(args[3]), version = args[4];
            Assembly setup = Assembly.LoadFile(installer);
            FileVersionInfo file = FileVersionInfo.GetVersionInfo(installer);
            if (setup.GetName().Version.ToString() != version || file.FileVersion != version || file.ProductVersion != version)
                throw new Exception("Installer version mismatch");
            string[] names = { name + ".dll", "Microsoft.Office.Tools.Common.v4.0.Utilities.dll", "Interop.Microsoft.Office.Core.dll", name + ".dll.manifest", name + ".vsto" };
            if (setup.GetManifestResourceNames().Length != 5) throw new Exception("Exactly five payloads required");
            foreach (string resourceName in names)
                using (Stream input = setup.GetManifestResourceStream(resourceName))
                using (var buffer = new MemoryStream())
                {
                    if (input == null) throw new Exception("Missing payload: " + resourceName);
                    input.CopyTo(buffer);
                    if (Hash(buffer.ToArray()) != Hash(File.ReadAllBytes(Path.Combine(root, "build", resourceName))))
                        throw new Exception("Stale payload: " + resourceName);
                }
            Assembly dll = Assembly.LoadFile(Path.Combine(root, "build", name + ".dll"));
            if (dll.GetName().Version.ToString() != version) throw new Exception("DLL version mismatch");
            if (name == "Excel_Iterate" && Array.IndexOf(dll.GetManifestResourceNames(), "Excel_Iterate.MigrationComBinder.exe") >= 0)
                throw new Exception("Excel payload still embeds the retired migration helper");
            VerifyDocumentAssets(root, dll, Path.GetFullPath(args[5]));
            // Check embedded documents without opening the Agent reader windows.
            foreach (string document in new string[] { "DevExperience", "UserGuide" })
                using (Stream input = dll.GetManifestResourceStream("WpsShared.Agent." + document + ".md"))
                {
                    if (input == null) throw new Exception("Embedded document missing: " + document);
                    string filename = document == "DevExperience" ? "EmbeddedDevExperience.md" : "EmbeddedUserGuide.md";
                    string expected = File.ReadAllText(Path.Combine(root, "build", filename), Encoding.UTF8);
                    
                    expected = expected.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", Environment.NewLine);
                    using (var reader = new StreamReader(input, Encoding.UTF8))
                        if (reader.ReadToEnd() != expected) throw new Exception("Embedded document differs from current source: " + filename);
                    Console.WriteLine("PASS: normalized embedded " + filename + " (whole source document)");
                }
            using (Stream input = dll.GetManifestResourceStream("WpsShared.Agent.Home.html"))
            using (var buffer = new MemoryStream())
            {
                if (input == null) throw new Exception("Embedded page missing");
                input.CopyTo(buffer);
                if (Hash(buffer.ToArray()) != Hash(new UTF8Encoding(false).GetBytes(CorePage(File.ReadAllText(sitePath,Encoding.UTF8))))) throw new Exception("Embedded page differs from canonical snapshot");
                Console.WriteLine("Website SHA256: " + Hash(buffer.ToArray()));
            }
            Console.WriteLine("PASS: " + name + " versions, five exact payloads, canonical embedded HTML");
            Console.WriteLine("Installer bytes: " + new FileInfo(installer).Length);
            Console.WriteLine("Installer SHA256: " + Hash(File.ReadAllBytes(installer)));
            return 0;
        }
        catch (Exception error) { Console.WriteLine("FAIL: " + error); return 1; }
    }
}
