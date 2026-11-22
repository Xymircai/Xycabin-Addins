// MakeCertCapi —— 用 CertEnroll COM 在「当前用户\个人」证书存储中创建
// 自签名代码签名证书，私钥落在 CAPI（Microsoft Enhanced RSA and AES
// Cryptographic Provider, AT_SIGNATURE），供 ClickOnce/VSTO 清单签名使用。
//
// 用法:
//   MakeCertCapi.exe [主题CN]
//   MakeCertCapi.exe --export-pfx <指纹> <输出路径> <口令>
// 创建成功时输出证书 sha1 指纹；导出模式写出含私钥的 PFX。
using System;
using System.IO;
using System.Text;
using Id = WpsShared.ProjectInfo;

internal static class MakeCertCapi
{
    private const int XCN_AT_SIGNATURE = 2;              // KeySpec: AT_SIGNATURE
    private const int XCN_NCRYPT_ALLOW_EXPORT_FLAG = 1;  // ExportPolicy: 可导出
    private const int ContextUser = 0x1;
    private const int XCN_CRYPT_STRING_BASE64 = 0x1;
    private const int AllowNoOutstandingResponse = 0x2;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        try
        {
            if (args.Length == 4 && args[0] == "--export-pfx")
                return ExportPfx(args[1], args[2], args[3]);
            if (args.Length > 1)
                throw new ArgumentException("usage: MakeCertCapi [subject] | --export-pfx <thumbprint> <output> <password>");
            return CreateCertificate(args.Length == 1 ? args[0] : Id.CertificateSubject);
        }
        catch (Exception ex)
        {
            Console.WriteLine("MakeCertCapi failed: " + ex);
            return 1;
        }
    }

    private static int CreateCertificate(string subject)
    {
        // ---- 私钥（CAPI, AT_SIGNATURE, 可导出） ----
        Type tKey = Type.GetTypeFromProgID("X509Enrollment.CX509PrivateKey");
        dynamic key = Activator.CreateInstance(tKey);
        key.ProviderName = "Microsoft Enhanced RSA and AES Cryptographic Provider";
        key.KeySpec = XCN_AT_SIGNATURE;
        key.Length = 2048;
        key.MachineContext = false;
        key.ExportPolicy = XCN_NCRYPT_ALLOW_EXPORT_FLAG;

        // ---- 证书请求（自签名；EKU 对清单签名非必需，从略） ----
        Type tDn = Type.GetTypeFromProgID("X509Enrollment.CX500DistinguishedName");
        dynamic dn = Activator.CreateInstance(tDn);
        dn.Encode(subject, 0);
        Type tReq = Type.GetTypeFromProgID("X509Enrollment.CX509CertificateRequestCertificate");
        dynamic req = Activator.CreateInstance(tReq);
        req.InitializeFromPrivateKey(ContextUser, key, "");
        req.Subject = dn;
        req.Issuer = dn;
        req.NotBefore = DateTime.Now.AddDays(-1);
        req.NotAfter = DateTime.Now.AddYears(5);

        // ---- 注册（创建并安装到 My 存储） ----
        Type tEnroll = Type.GetTypeFromProgID("X509Enrollment.CX509Enrollment");
        dynamic enroll = Activator.CreateInstance(tEnroll);
        enroll.InitializeFromRequest(req);
        enroll.CertificateFriendlyName = subject;
        string csr = enroll.CreateRequest(XCN_CRYPT_STRING_BASE64);
        enroll.InstallResponse(AllowNoOutstandingResponse, csr, XCN_CRYPT_STRING_BASE64, "");

        // Enrollment.Certificate is this response's certificate. Never select an older
        // certificate merely because several certificates share the same subject.
        string enrolledThumb;
        using (var enrolled = new System.Security.Cryptography.X509Certificates.X509Certificate2(
            Convert.FromBase64String((string)enroll.Certificate[XCN_CRYPT_STRING_BASE64])))
            enrolledThumb = enrolled.Thumbprint;

        // ---- 从存储找回证书打印指纹 ----
        var store = new System.Security.Cryptography.X509Certificates.X509Store(
            "My", System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser);
        store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);
        try
        {
            foreach (var c in store.Certificates)
            {
                if (string.Equals(c.Thumbprint, enrolledThumb, StringComparison.OrdinalIgnoreCase))
                {
                    if (!c.HasPrivateKey) throw new InvalidOperationException("New certificate has no private key.");
                    Console.WriteLine("thumbprint: " + c.Thumbprint);
                    Console.WriteLine("provider  : " + GetProviderName(c));
                    return 0;
                }
            }
        }
        finally { store.Close(); }
        Console.WriteLine("cert created but not found in store");
        return 1;
    }

    private static int ExportPfx(string thumbprint, string outputPath, string password)
    {
        string thumb = (thumbprint ?? "").Trim().ToUpperInvariant();
        if (thumb.Length != 40 || !System.Text.RegularExpressions.Regex.IsMatch(thumb, "^[0-9A-F]{40}$"))
            throw new ArgumentException("Invalid certificate thumbprint.");
        if (string.IsNullOrEmpty(outputPath))
            throw new ArgumentException("An output PFX path is required.");
        if (string.IsNullOrEmpty(password))
            throw new ArgumentException("A PFX password is required.");

        var store = new System.Security.Cryptography.X509Certificates.X509Store(
            "My", System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser);
        store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);
        try
        {
            foreach (var cert in store.Certificates)
            {
                if (!string.Equals(cert.Thumbprint, thumb, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!cert.HasPrivateKey)
                    throw new InvalidOperationException("Certificate has no private key: " + thumb);
                byte[] data = cert.Export(
                    System.Security.Cryptography.X509Certificates.X509ContentType.Pfx, password);
                string fullPath = Path.GetFullPath(outputPath);
                string parent = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(parent))
                    Directory.CreateDirectory(parent);
                File.WriteAllBytes(fullPath, data);
                Console.WriteLine("EXPORTED_PFX " + thumb + " " + fullPath);
                return 0;
            }
        }
        finally { store.Close(); }
        throw new InvalidOperationException("Certificate was not found in CurrentUser\\My: " + thumb);
    }

    private static string GetProviderName(System.Security.Cryptography.X509Certificates.X509Certificate2 c)
    {
        try
        {
            var csp = c.PrivateKey as System.Security.Cryptography.RSACryptoServiceProvider;
            if (csp != null)
                return csp.CspKeyContainerInfo.ProviderName + " (CAPI)";
            return c.PrivateKey != null ? c.PrivateKey.GetType().Name + " (CNG?)" : "(no key)";
        }
        catch { return "?"; }
    }
}
