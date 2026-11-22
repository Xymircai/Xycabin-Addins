// SHA-256 ClickOnce/VSTO signing. Certificate selection is explicit; project PFX import is opt-in.
// SignCmi <manifest> <thumbprint> [expected-subject]
// SignCmi --check-certificate <thumbprint> <expected-subject>
// SignCmi --import-pfx <pfx-path> <expected-thumbprint> <expected-subject>  (password on stdin)
// SignCmi --remove-certificate <thumbprint>   (test fixtures only)
using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

internal static class SignCmi
{
    private const string SD = "System.Deployment, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    private const BindingFlags ALLI = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
    private const BindingFlags ALLS = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try { return Run(args); }
        catch (Exception ex)
        {
            while (ex.InnerException != null) ex = ex.InnerException;
            Console.WriteLine("FAIL " + ex.GetType().Name + ": " + ex.Message);
            return 1;
        }
    }

    private static int Run(string[] args)
    {
        if (args.Length == 4 && args[0] == "--import-pfx")
        {
            using (var input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false, true), false))
            {
                string password = input.ReadLine();
                if (string.IsNullOrEmpty(password)) throw new InvalidOperationException("PFX password is required on stdin.");
                return ImportPfx(args[1], password, args[2], args[3]);
            }
        }
        if (args.Length == 2 && args[0] == "--remove-certificate")
            return RemoveCertificate(args[1]);
        if (args.Length < 2 || args.Length > 3)
            throw new ArgumentException("usage: SignCmi <manifest|--check-certificate> <thumbprint> [expected-subject]");
        bool certificateOnly = args[0] == "--check-certificate";
        if (certificateOnly && args.Length != 3)
            throw new ArgumentException("Certificate checks require the expected subject.");
        string thumb = NormalizeThumbprint(args[1]);
        string path = certificateOnly ? null : Path.GetFullPath(args[0]);
        if (!certificateOnly && !WpsShared.SourcePaths.Exists(path))
            throw new FileNotFoundException("Manifest is missing.", path);
        using (X509Certificate2 cert = FindCertificate(thumb))
        {
            if (args.Length == 3)
                RequireCertificateIdentity(cert, args[2]);
            RequireUsableCertificate(cert);
            if (certificateOnly)
            {
                Console.WriteLine("CERTIFICATE_OK " + cert.Thumbprint.ToUpperInvariant());
                return 0;
            }
            return SignManifest(path, cert);
        }
    }

    private static string NormalizeThumbprint(string value)
    {
        string thumb = (value ?? "").Trim().ToUpperInvariant();
        if (thumb.Length != 40 || !System.Text.RegularExpressions.Regex.IsMatch(thumb, "^[0-9A-F]{40}$"))
            throw new ArgumentException("Invalid certificate thumbprint.");
        return thumb;
    }

    private static void RequireCertificateIdentity(X509Certificate2 cert, string expectedSubject)
    {
        if (!string.Equals(cert.Subject, expectedSubject, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("Certificate subject does not match the project.");
    }

    private static void RequireUsableCertificate(X509Certificate2 cert)
    {
        DateTime now = DateTime.UtcNow;
        if (now < cert.NotBefore.ToUniversalTime() || now > cert.NotAfter.ToUniversalTime())
            throw new InvalidOperationException("Certificate is outside its validity period.");
        using (RSA rsa = cert.GetRSAPrivateKey())
        {
            if (rsa == null) throw new InvalidOperationException("Certificate private key is unavailable.");
            // A read-only signing challenge checks access to the private key, not just its metadata.
            byte[] data = Encoding.UTF8.GetBytes("WPS_CERTIFICATE_CHECK");
            byte[] signature = rsa.SignData(data, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            if (!rsa.VerifyData(data, signature, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
                throw new CryptographicException("Certificate private-key check failed.");
        }
    }

    private static int SignManifest(string path, X509Certificate2 cert)
    {
        using (RSA rsa = cert.GetRSAPrivateKey())
        {
            if (rsa == null) throw new InvalidOperationException("Certificate private key is unavailable.");
            Assembly asm = Assembly.Load(SD);
            Type signerType = asm.GetType("System.Deployment.Internal.CodeSigning.CmiManifestSigner2", true);
            Type manifestType = asm.GetType("System.Deployment.Internal.CodeSigning.SignedCmiManifest2", true);
            object signer = Activator.CreateInstance(signerType, ALLI, null,
                new object[] { (AsymmetricAlgorithm)rsa, cert, true }, null);
            XmlDocument doc = new XmlDocument();
            doc.PreserveWhitespace = true;
            doc.XmlResolver = null;
            using (XmlReader reader = XmlReader.Create(path,
                new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                doc.Load(reader);
            MethodInfo remove = manifestType.GetMethod("RemoveExistingSignature", ALLS);
            if (remove != null) remove.Invoke(null, new object[] { doc });
            object signed = Activator.CreateInstance(manifestType, ALLI, null, new object[] { doc, true }, null);
            MethodInfo sign = manifestType.GetMethod("Sign", ALLI, null, new Type[] { signerType }, null);
            if (sign == null) throw new MissingMethodException("SHA-256 ClickOnce signer is unavailable.");
            sign.Invoke(signed, new object[] { signer });
            doc.Save(path);
            Console.WriteLine("SIGNED_SHA256 " + cert.Thumbprint.ToUpperInvariant() + " " + path);
            return 0;
        }
    }

    private static int ImportPfx(string path, string password, string thumb, string subject)
    {
        thumb = NormalizeThumbprint(thumb);
        string fullPath = Path.GetFullPath(path);
        if (!File.Exists(fullPath))
            throw new FileNotFoundException("PFX is missing.", fullPath);

        using (X509Certificate2 loaded = new X509Certificate2(fullPath, password,
            X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable))
        {
            if (!string.Equals(loaded.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("PFX certificate does not match the pinned thumbprint.");
            RequireCertificateIdentity(loaded, subject);
            RequireUsableCertificate(loaded);

            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadWrite);
                try
                {
                    foreach (X509Certificate2 existing in store.Certificates)
                    {
                        if (!string.Equals(existing.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                            continue;
                        if (!existing.HasPrivateKey)
                            throw new InvalidOperationException("A certificate with this thumbprint exists in CurrentUser\\My without a private key; remove it manually and run again.");
                        Console.WriteLine("CERTIFICATE_ALREADY_AVAILABLE " + thumb);
                        return 0;
                    }
                    store.Add(loaded);
                }
                finally { store.Close(); }
            }
        }
        using (X509Certificate2 imported = FindCertificate(thumb))
        {
            RequireCertificateIdentity(imported, subject);
            RequireUsableCertificate(imported);
        }
        Console.WriteLine("IMPORTED_OK " + thumb);
        return 0;
    }

    private static int RemoveCertificate(string thumb)
    {
        thumb = NormalizeThumbprint(thumb);
        int removed = 0;
        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            store.Open(OpenFlags.ReadWrite);
            try
            {
                foreach (X509Certificate2 cert in store.Certificates)
                {
                    if (string.Equals(cert.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                    {
                        store.Remove(cert);
                        removed++;
                    }
                }
            }
            finally { store.Close(); }
        }
        Console.WriteLine(removed == 0 ? "CERTIFICATE_NOT_PRESENT " + thumb : "REMOVED_CERTIFICATE " + thumb);
        return 0;
    }

    private static X509Certificate2 FindCertificate(string thumb)
    {
        // The shared PFX is imported into CurrentUser\My only. Searching the same
        // scope prevents a public-key-only copy elsewhere from looking usable.
        thumb = NormalizeThumbprint(thumb);
        using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
        {
            try { store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly); }
            catch (CryptographicException) { throw new InvalidOperationException("CurrentUser\\My store is unavailable: " + thumb); }
            foreach (X509Certificate2 candidate in store.Certificates)
                if (string.Equals(candidate.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                {
                    if (!candidate.HasPrivateKey)
                        throw new InvalidOperationException("Certificate in CurrentUser\\My has no private key: " + thumb);
                    return new X509Certificate2(candidate);
                }
        }
        throw new InvalidOperationException("Pinned certificate was not found: " + thumb);
    }
}
