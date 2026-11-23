using System;
#if WPS_WORD
using WordUpdater;
#else
using Excel_Iterate;
#endif
using Id = WpsShared.ProjectInfo;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WpsShared;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Windows.Forms;
using Microsoft.Win32;

// Shared standalone VSTO installer. Identity is generated from project.json.

#if WPS_WORD
namespace WordUpdaterSetup
#else
namespace Excel_IterateSetup
#endif
{
    internal static class Program
    {
        internal const string AddinKeyPath =
            Id.AddinKeyPath;
        internal const string FriendlyName = Id.FriendlyName;
        internal const string Description = Id.Description;

        // 「上次安装目录」的记忆位置。两个雷区都绕开：
        //   1) 不能放在 AddinKeyPath 下面 —— Uninstall() 会 DeleteSubKeyTree
        //      把整棵加载项键删掉，卸载后就记不住了；
        //   2) 不能放在 Software\WordUpdater 下面 —— 那是已废弃的 COM 安装器
        //      （已归档的 ComRegistration 方案）留下的脏数据键，InstallDir 指向错误路径。
        internal const string SetupKeyPath = Id.SetupKeyPath;
        internal const string LastDirValueName = "LastInstallDir";
        internal const string LastTimeValueName = "LastInstalledAt";

        // ==== ★ 装载成败的第一开关：WPS 自己的加载项白名单 AddinsWL ====
        // 2026-09-26 实测根因：WPS 扫 HKCU\...\Word\Addins 时**只看在榜的名字**，
        // 不在榜的一律静默跳过——不报错、不写事件日志、LoadBehavior 也不会降级，
        // 表面上和"没装"一模一样。这台机器上名单已收录几十家第三方加载项
        // （CPAHelperForWordRe、审计小哥等），漏了这条就永远出不来选项卡。
        // 位置在 HKCU，不需要管理员权限；卸载时要删掉，别留垃圾。
        internal const string AddinsWlKeyPath =
            Id.AddinsWlKeyPath;
        internal const string AddinName = Id.AddinName;

        internal const string VstaSolutionsKeyPath =
            @"Software\Microsoft\VSTA\Solutions";

        // VSTO 信任包含列表只用于卸载清理；安装必须走官方运行库。
        // 正常路径必须走官方流程：由 VSTOInstaller 弹出
        // 「Microsoft Office 自定义项安装程序」，用户点「安装」后 VSTO 自己完成
        // 信任授予 + ClickOnce 订阅。（曾经反过来替它预先登记信任、绕开弹窗，
        // 结果加载项反而没被正确安装——别再走那条路。）
        internal const string InclusionKeyPath =
            @"Software\Microsoft\VSTO\Security\Inclusion";

        internal static readonly string[] Payload = Id.Payload;

        internal static IEnumerable<string> InstalledPayload()
        {
            foreach (string name in Payload) yield return name;
#if WPS_EXCEL
            // An earlier Excel installer deployed this companion.
            yield return "MigrationComBinder.exe";
#endif
        }

        /// <summary>部署清单文件名。换名字时上面的 Payload 也要跟着改。</summary>
        internal const string PayloadVstoName = Id.PayloadVstoName;

        // ==== 旧名（WpsLinkUpdater）残留清理 ====
        // 2026-10-01 项目更名 WordUpdater。新安装器按新名卸载，碰不到旧名
        // 留下的挂载点/白名单/订阅/信任/安装目录/记忆键。保留这份旧名常量，
        // 安装/卸载时顺手清一遍（幂等：新机器上这些键与目录都不存在，逐个跳过）。
        internal const string LegacyAddinKeyPath =
            Id.LegacyAddinKeyPath;
        internal const string LegacyAddinName = Id.LegacyAddinName;
        internal const string LegacyVstoName = Id.LegacyVstoName;
        internal const string LegacySetupKeyPath = Id.LegacySetupKeyPath;
        internal static readonly string[] LegacyPayload = Id.LegacyPayload;

        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(MaintenanceEntry.IsUninstall(args)));
        }
    }

    internal class MainForm : Form
    {
        private readonly TextBox _txtPath;
        private readonly TextBox _log;
        private readonly Button _btnInstall;
        private readonly Button _btnUninstall;
        private readonly Button _btnBrowse;
        private readonly Button _btnClose;
        private readonly object _logLock = new object();
        private readonly string _logPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            Id.AssemblyName, "SetupLogs", "setup-" + DateTime.Now.ToString("yyyyMMdd-HHmmss")
            + "-" + Process.GetCurrentProcess().Id + ".log");
        private bool _busy;
        private readonly bool _uninstallMode;

        internal MainForm() : this(false) { }
        internal MainForm(bool uninstallMode)
        {
            _uninstallMode = uninstallMode;
            Text = InstallerText.WindowTitle;
            if (uninstallMode) Text = "卸载 " + InstallerText.Title;
            SuspendLayout();
            Ui.PrepareForm(this, new Size(744, 576), new Size(560, 440));
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = MinimizeBox = false;
            var root = new TableLayoutPanel {
                Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 1, RowCount = 5
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            root.Controls.Add(CreateBanner(), 0, 0);

            string initNote;
            var pathRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 94));
            pathRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var field = new UiCard { Dock = DockStyle.Fill, BackColor = Ui.GridHead,
                Padding = new Padding(11, 5, 10, 5), Margin = new Padding(0, 0, 10, 0) };
            _txtPath = new TextBox { AccessibleName = "安装位置", BorderStyle = BorderStyle.None,
                Anchor = AnchorStyles.Left | AnchorStyles.Right, Font = Ui.MakeFont(11f, FontStyle.Regular, Font),
                BackColor = field.BackColor, ForeColor = Ui.Ink, Text = ResolveInitialDir(out initNote) };
            var center = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 1 };
            center.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            center.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            center.Controls.Add(_txtPath);
            field.Controls.Add(center);
            pathRow.Controls.Add(field, 0, 0);
            var browse = Ui.Button("浏览…", 94);
            browse.Dock = DockStyle.Fill; browse.Margin = Padding.Empty;
            _btnBrowse = browse; browse.Click += delegate { Browse(); };
            pathRow.Controls.Add(browse, 1, 0);
            root.Controls.Add(Section("安装位置", pathRow), 0, 1);

            _log = new TextBox { AccessibleName = "安装日志", Multiline = true, ReadOnly = true,
                ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.None, Dock = DockStyle.Fill,
                BackColor = Color.White, ForeColor = Ui.Ink, Margin = Padding.Empty,
                Font = Ui.MakeFont(10.5f, FontStyle.Regular, Font) };
            root.Controls.Add(Section("安装日志", _log), 0, 2);

            var bar = Ui.ActionBar();
            bar.FlowDirection = FlowDirection.RightToLeft;
            var install = Ui.Button("安装", 118);
            Ui.StylePrimary(install, Id.Host == "Excel");
            install.Font = Ui.MakeFont(11f, FontStyle.Bold, Font);
            install.Visible = !uninstallMode;
            _btnInstall = install; install.Click += delegate { Install(); };
            var uninstall = Ui.Button("卸载", 86);
            if (uninstallMode) Ui.StyleDanger(uninstall);
            _btnUninstall = uninstall; uninstall.Click += delegate { Uninstall(); };
            var close = Ui.Button("关闭", 86);
            _btnClose = close; close.Click += delegate { Close(); };
            var runtime = Ui.Button("获取运行库", 122);
            runtime.Click += delegate {
                Log("微软 VSTO 运行库官方下载地址：" + RuntimeLocator.DownloadUrl);
                try { Process.Start(new ProcessStartInfo(RuntimeLocator.DownloadUrl) { UseShellExecute = true }); }
                catch (Exception ex) { Log("打开微软下载页失败：" + ex.Message + "\r\n请复制上述地址到浏览器，下载并运行 vstor_redist.exe。"); }
            };
            bar.Controls.AddRange(new Control[] { install, uninstall, runtime, close });
            root.Controls.Add(bar, 0, 3);
            var copyright = Ui.Label(InstallerText.Copyright, 9f, false);
            copyright.ForeColor = Ui.GridHeadInk;
            root.Controls.Add(copyright, 0, 4);
            Controls.Add(root);
            ResumeLayout(true);
            FormClosing += delegate(object sender, FormClosingEventArgs e) {
                if (_busy) {
                    e.Cancel = true;
                    Log("官方安装程序尚未结束。请在其窗口完成安装或点「取消」，再关闭本窗口。");
                }
            };
            Log("日志文件：" + _logPath);
            Log(initNote); Log(uninstallMode ? "就绪。点击「卸载」移除此加载项。" : "就绪。请选择安装位置后点击「安装」。");
            Log("提示：安装无需管理员权限；建议先保存文档并退出 WPS。");
            Log("目标宿主是 " + Id.HostDisplayName + "。授权窗口由微软 VSTO 运行库提供，标题中的 Microsoft Office 不表示更换宿主。");
            Log("安装后可从 Windows「已安装的应用」卸载，也可重新运行本安装包点击「卸载」。");
            if (uninstallMode) { _txtPath.ReadOnly = true; _btnBrowse.Enabled = false; Shown += delegate { _btnUninstall.Select(); }; }
        }

        private static UiCard Section(string caption, Control content)
        {
            var card = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(16, 10, 16, 12) };
            var rows = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            rows.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            rows.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            rows.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var label = Ui.Label(caption, 12f, true);
            label.AutoSize = true;
            label.Margin = new Padding(0, 0, 0, 7);
            rows.Controls.Add(label, 0, 0);
            rows.Controls.Add(content, 0, 1);
            card.Controls.Add(rows);
            return card;
        }

        private UiCard CreateBanner()
        {
            bool excel = Id.Host == "Excel";
            var ink = excel ? Color.FromArgb(17, 73, 64) : Color.FromArgb(27, 54, 106);
            var card = new UiCard { Dock = DockStyle.Fill, BackColor = ink, Padding = new Padding(20, 10, 20, 10) };
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2, BackColor = ink };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 76));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 60));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));
            var mark = new PictureBox { SizeMode = PictureBoxSizeMode.Zoom, Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 14, 4), TabStop = false };
            IntPtr icon;
            if (PrivateExtractIcons(Assembly.GetExecutingAssembly().Location, 0, 128, 128, out icon, IntPtr.Zero, 1, 0) == 1 && icon != IntPtr.Zero)
            {
                try { using (var extracted = Icon.FromHandle(icon)) mark.Image = extracted.ToBitmap(); }
                finally { DestroyIcon(icon); }
                mark.Disposed += delegate { mark.Image.Dispose(); };
            }
            layout.Controls.Add(mark, 0, 0);
            layout.SetRowSpan(mark, 2);
            var title = Ui.Label(InstallerText.Title, 22f, true);
            title.ForeColor = Color.White;
            title.TextAlign = ContentAlignment.BottomLeft;
            layout.Controls.Add(title, 1, 0);
            var subtitle = Ui.Label(InstallerText.Subtitle, 11f, false);
            subtitle.ForeColor = excel ? Color.FromArgb(212, 242, 230) : Color.FromArgb(220, 233, 255);
            subtitle.TextAlign = ContentAlignment.TopLeft;
            layout.Controls.Add(subtitle, 1, 1);
            if (!string.IsNullOrWhiteSpace(InstallerText.FeatureText))
            {
                var details = new LinkLabel { Text = "功能说明", AutoSize = true, Anchor = AnchorStyles.Right,
                    Font = Ui.MakeFont(10f, FontStyle.Bold, Font), LinkColor = Color.White, ActiveLinkColor = subtitle.ForeColor,
                    VisitedLinkColor = Color.White, LinkBehavior = LinkBehavior.HoverUnderline };
                details.LinkClicked += delegate {
                    UiDialog.Show(this, InstallerText.FeatureText, InstallerText.Title + " · 功能说明",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                };
                layout.Controls.Add(details, 2, 0);
                layout.SetRowSpan(details, 2);
            }
            card.Controls.Add(layout);
            return card;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "PrivateExtractIconsW", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern uint PrivateExtractIcons(string path, int index, int width, int height, out IntPtr icon, IntPtr id, uint count, uint flags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr icon);

        private void Browse()
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = "选择安装目录";
                dlg.SelectedPath = _txtPath.Text;
                dlg.ShowNewFolderButton = true;
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    _txtPath.Text = dlg.SelectedPath;
            }
        }

        private async void Install()
        {
            if (_busy) return;
            string dir = NormalizeDir(_txtPath.Text);
            if (!IsUsableDir(dir))
            {
                UiDialog.Show(this, "请选择有效的绝对安装目录。", "安装",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string running = RunningWps();
            if (running != null)
            {
                DialogResult go = UiDialog.Show(this,
                    "检测到 WPS 正在运行（" + running + "）。\r\n\r\n"
                    + "建议先保存文档并完全退出 WPS，以免文件被占用。是否仍要继续？",
                    "安装", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (go != DialogResult.Yes) { Log("已取消安装。"); return; }
            }
            SetBusy(true);
            try
            {
                InstallResult result = await Task.Factory.StartNew(delegate {
                    return InstallFlow.Install(new InstallActions {
                        FindRuntime = FindAndLogRuntime,
                        CleanupSubscriptions = PrepareSubscriptions,
                        ReleasePayload = delegate { PrepareMaintenance(dir); ReleasePayload(dir, Log); },
                        InvokeInstall = delegate(string tool) { return RunStandardInstall(tool, dir); },
                        Register = delegate { WriteRegistry(dir); AddToAddinsWl(); RegisterMaintenance(dir); },
                        Complete = delegate {
                            SaveLastInstallDir(dir, "已记住安装目录：" + dir);
                            Log("官方安装已完成。重新打开 " + Id.HostDisplayName + " 后即可使用。");
                        }
                    });
                });
                ShowOperationResult(result, "安装");
            }
            catch (Exception ex) { ShowOperationResult(new InstallResult(InstallStatus.Failed, ex.Message, null), "安装"); }
            finally { SetBusy(false); }
        }

        private static void ReleasePayload(string dir, Action<string> log)
        {
            Directory.CreateDirectory(dir);
            log("安装目录：" + dir);
            Assembly self = Assembly.GetExecutingAssembly();
            foreach (string name in Program.Payload)
            {
                using (Stream src = self.GetManifestResourceStream(name))
                {
                    if (src == null) throw new FileNotFoundException("安装包内缺少文件：" + name);
                    using (FileStream dst = new FileStream(Path.Combine(dir, name), FileMode.Create, FileAccess.Write))
                        src.CopyTo(dst);
                }
                log("  已释放 " + name);
            }
        }

        private async void Uninstall()
        {
            if (_busy) return;
            if (RunningWps() != null)
            {
                UiDialog.Show(this, "请先保存文档并完全退出 WPS（包括后台进程），再点击卸载。安装器不会替你关闭 WPS。", "卸载", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            DialogResult ask = UiDialog.Show(this,
                "确定卸载「" + Id.DisplayName + "」吗？\r\n官方订阅注销成功后，才会移除加载项登记和程序文件。",
                "卸载", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (ask != DialogResult.Yes) return;
            string dir = NormalizeDir(_txtPath.Text);
            string registeredDir = null;
            SetBusy(true);
            try
            {
                InstallResult result = await Task.Factory.StartNew(delegate {
                    return InstallFlow.Uninstall(new UninstallActions {
                        FindRuntime = FindAndLogRuntime,
                        CleanupSubscriptions = delegate(string tool) {
                            registeredDir = ReadRegisteredDir(Program.AddinKeyPath);
                            return PrepareSubscriptions(tool);
                        },
                        RemovePayloadAndRegistration = delegate {
                            if (!string.IsNullOrEmpty(registeredDir)) {
                                SaveLastInstallDir(registeredDir, "已保留上次安装目录：" + registeredDir);
                                DeletePayload(registeredDir, Program.InstalledPayload());
                            }
                            if (IsUsableDir(dir) && !string.Equals(dir, registeredDir, StringComparison.OrdinalIgnoreCase)
                                && File.Exists(Path.Combine(dir, Program.PayloadVstoName)))
                                DeletePayload(dir, Program.InstalledPayload());
                            Registry.CurrentUser.DeleteSubKeyTree(Program.AddinKeyPath, false);
                            RemoveOwnedInclusion(Program.PayloadVstoName);
                            RemoveFromAddinsWl();
                            RemoveMaintenanceRegistration(dir);
                        }
                    });
                });
                ShowOperationResult(result, "卸载");
            }
            catch (Exception ex) { ShowOperationResult(new InstallResult(InstallStatus.Failed, ex.Message, null), "卸载"); }
            finally { SetBusy(false); }
        }

        private string FindAndLogRuntime()
        {
            string tool = RuntimeLocator.FindInstalledTool();
            Log(tool == null ? "运行库检查失败：未找到 VSTOInstaller.exe。" : "运行库检查通过：" + tool);
            return tool;
        }

        private InstallResult PrepareSubscriptions(string tool)
        {
            InstallResult result = RemoveOwnedSubscriptions(tool, Program.PayloadVstoName);
            if (result.Status != InstallStatus.Success) return result;
            return CleanupLegacyInstallation(tool);
        }

        private void ShowOperationResult(InstallResult result, string operation)
        {
            Log(result.Status + ": " + result.Message);
            string message = result.Message;
            if (result.Status == InstallStatus.Success)
                message += "\r\n\r\n请重新打开 " + Id.HostDisplayName + " 使改动生效。";
            else if (result.Status != InstallStatus.MissingRuntime)
                message += "\r\n\r\n本次" + operation + "未完成；请查看日志后重试。";
            message += "\r\n\r\n日志：" + _logPath;
            MessageBoxIcon icon = result.Status == InstallStatus.Success ? MessageBoxIcon.Information
                : result.Status == InstallStatus.Failed ? MessageBoxIcon.Error : MessageBoxIcon.Warning;
            UiDialog.Show(this, message, operation, MessageBoxButtons.OK, icon);
        }

        private static string ReadRegisteredDir(string keyPath)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                string dir;
                return key != null && TryGetDirFromManifest(key.GetValue("Manifest") as string, out dir) ? dir : null;
            }
        }

        /// <summary>规范化目录：去首尾空白和末尾分隔符，但保留根目录（如 D:\）。</summary>
        private static string NormalizeDir(string raw)
        {
            if (raw == null) return "";
            string d = raw.Trim();
            while (d.Length > 3 &&
                   (d[d.Length - 1] == '\\' || d[d.Length - 1] == '/'))
                d = d.Substring(0, d.Length - 1);
            return d;
        }

        /// <summary>是否可作为安装目录使用（只做合法性检查，不检查是否存在）。</summary>
        private static bool IsUsableDir(string dir)
        {
            if (string.IsNullOrEmpty(dir)) return false;
            if (dir.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;
            if (!Path.IsPathRooted(dir)) return false;
            try
            {
                Path.GetFullPath(dir);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 从 Manifest 值（file:///D:/x/WordUpdater.vsto|vstolocal）反解出目录。
        /// 解析失败返回 false，不抛异常。
        /// </summary>
        private static bool TryGetDirFromManifest(string manifestValue, out string dir)
        {
            dir = "";
            try
            {
                if (string.IsNullOrEmpty(manifestValue)) return false;
                string v = manifestValue;
                int bar = v.IndexOf('|');
                if (bar >= 0) v = v.Substring(0, bar);
                Uri uri;
                if (!Uri.TryCreate(v, UriKind.Absolute, out uri) || !uri.IsFile) return false;
                v = uri.LocalPath;
                string d = Path.GetDirectoryName(v);
                if (!IsUsableDir(d)) return false;
                dir = NormalizeDir(d);
                return dir.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// ① 当前已安装目录（来自加载项键的 Manifest）。
        /// 目录已被删除时返回 null —— 注册表还在但目录没了，这状态已失效。
        /// </summary>
        private static string ReadCurrentInstallDir()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(Program.AddinKeyPath))
                {
                    if (k == null) return null;
                    string dir;
                    if (!TryGetDirFromManifest(k.GetValue("Manifest") as string, out dir))
                        return null;
                    return Directory.Exists(dir) ? dir : null;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>② 上次安装目录（来自独立的 SetupKeyPath，卸载不会被删）。</summary>
        private static string ReadLastInstallDir()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(Program.SetupKeyPath))
                {
                    if (k == null) return null;
                    string v = k.GetValue(Program.LastDirValueName) as string;
                    if (!IsUsableDir(v)) return null;
                    string d = NormalizeDir(v);
                    return d.Length > 0 ? d : null;
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 记住安装目录。任何失败都只记一行日志，绝不向上抛 ——
        /// 记不住路径是非致命的，不能让它把「安装成功」变成「安装失败」。
        /// </summary>
        private void SaveLastInstallDir(string dir, string okMessage)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(Program.SetupKeyPath))
                {
                    if (k == null)
                    {
                        Log("（未能记住安装目录：无法创建注册表项 HKCU\\"
                            + Program.SetupKeyPath + "）");
                        return;
                    }
                    k.SetValue(Program.LastDirValueName, dir, RegistryValueKind.String);
                    k.SetValue(Program.LastTimeValueName,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        RegistryValueKind.String);
                    if (!string.IsNullOrEmpty(okMessage)) Log(okMessage);
                }
            }
            catch (Exception ex)
            {
                Log("（未能记住安装目录，不影响本次操作：" + ex.Message + "）");
            }
        }

        /// <summary>
        /// 解析初始目录：① 当前安装目录 → ② 上次安装目录 → ③ 默认目录。
        /// note 输出给用户看的一行说明。
        /// </summary>
        private string ResolveInitialDir(out string note)
        {
            string cur = ReadCurrentInstallDir();
            if (cur != null)
            {
                note = "已检测到当前安装目录，已自动填入：" + cur;
                return cur;
            }

            string last = ReadLastInstallDir();
            if (last != null)
            {
                if (Directory.Exists(last))
                {
                    note = "已自动填入上次安装目录：" + last + "。如需更换，请点「浏览…」。";
                }
                else
                {
                    note = "上次安装目录 " + last
                        + " 当前不可访问（已删除或盘符未连接），已按原样填入；"
                        + "请确认后安装，或点「浏览…」更换。";
                }
                return last;
            }

            string def = NormalizeDir(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", Id.AssemblyName));
            note = "未找到历史安装记录，已使用默认目录：" + def + "。可点「浏览…」更换。";
            return def;
        }

        private static void WriteRegistry(string dir)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(Program.AddinKeyPath))
            {
                if (k == null) throw new InvalidOperationException("无法创建注册表项。");
                string url = VstoUrlOf(dir);
                k.SetValue("Manifest", url + "|vstolocal", RegistryValueKind.String);
                k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                k.SetValue("FriendlyName", Program.FriendlyName, RegistryValueKind.String);
                k.SetValue("Description", Program.Description, RegistryValueKind.String);
            }
        }

        private static string MaintenanceLauncher()
        { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Id.AssemblyName, "Maintenance", Id.AssemblyName + "_Setup.exe"); }

        private static MaintenanceEntry Maintenance(string dir)
        {
            return new MaintenanceEntry(Id.AssemblyName, InstallerText.Title + "（WPS 加载项）", typeof(MainForm).Assembly.GetName().Version.ToString(), dir, MaintenanceLauncher());
        }

        private void PrepareMaintenance(string dir)
        {
            var entry = Maintenance(dir);
            using (var existing = Registry.CurrentUser.OpenSubKey(entry.KeyPath))
                if (existing != null && !entry.Owns(existing.GetValue("XyCabinProduct") as string, existing.GetValue("UninstallString") as string))
                    throw new IOException("卸载登记被其他程序占用，未覆盖：" + entry.KeyPath);
            string source = Assembly.GetExecutingAssembly().Location;
            if (string.Equals(Path.GetFullPath(source), entry.Launcher, StringComparison.OrdinalIgnoreCase)) return;
            Directory.CreateDirectory(Path.GetDirectoryName(entry.Launcher));
            string temporary = entry.Launcher + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.Copy(source, temporary, false);
            using (var first = SHA256.Create()) using (var second = SHA256.Create())
                if (Convert.ToBase64String(first.ComputeHash(File.ReadAllBytes(source))) != Convert.ToBase64String(second.ComputeHash(File.ReadAllBytes(temporary))))
                    throw new IOException("卸载工具复制核对失败，未继续安装。");
            if (File.Exists(entry.Launcher)) File.Replace(temporary, entry.Launcher, null); else File.Move(temporary, entry.Launcher);
            Log("已保留卸载工具，删除下载的安装包也不影响卸载：" + entry.Launcher);
        }

        private void RegisterMaintenance(string dir)
        {
            var entry = Maintenance(dir);
            using (var key = Registry.CurrentUser.CreateSubKey(entry.KeyPath))
            {
                if (key == null) throw new IOException("加载项已安装，但无法创建 Windows 卸载入口；可重新运行安装包卸载。");
                foreach (var pair in entry.Values()) key.SetValue(pair.Key, pair.Value, pair.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
            }
            Log("已添加 Windows「已安装的应用」卸载入口：" + entry.DisplayName);
        }

        private void RemoveMaintenanceRegistration(string dir)
        {
            var entry = Maintenance(dir); bool owned;
            using (var key = Registry.CurrentUser.OpenSubKey(entry.KeyPath))
                owned = key != null && entry.Owns(key.GetValue("XyCabinProduct") as string, key.GetValue("UninstallString") as string);
            if (owned) { Registry.CurrentUser.DeleteSubKeyTree(entry.KeyPath, false); Log("已移除本加载项的 Windows 卸载入口。"); }
            Log("用户配置、安装日志和维护工具缓存保留；不卸载 VSTO 运行库，不改动工作簿或其他加载项。");
        }

        /// <summary>安装目录对应的 VSTO Url（与包含列表里其他插件的写法一致）。</summary>
        private static string VstoUrlOf(string dir)
        {
            return new Uri(Path.Combine(Path.GetFullPath(dir), Program.PayloadVstoName)).AbsoluteUri;
        }

        private void RemoveOwnedInclusion(string manifestName)
        {
            using (RegistryKey root = Registry.CurrentUser.OpenSubKey(Program.InclusionKeyPath, true))
            {
                if (root == null) return;
                foreach (string name in root.GetSubKeyNames())
                {
                    bool owned;
                    using (RegistryKey key = root.OpenSubKey(name))
                        owned = key != null && RuntimeLocator.IsOwnedUrl(key.GetValue("Url") as string, manifestName);
                    if (owned) { root.DeleteSubKeyTree(name, false); Log("  已清除本加载项信任条目 " + name); }
                }
            }
        }

        // ---------------- WPS 加载项白名单（AddinsWL） ----------------

        /// <summary>
        /// 把加载项名字写进 WPS 白名单。这是能否加载的**第一开关**：
        /// 不在榜上时 WPS 会整条跳过，既不弹窗也不报错，最难排查。
        /// </summary>
        private void AddToAddinsWl()
        {
            using (RegistryKey root = Registry.CurrentUser.CreateSubKey(Program.AddinsWlKeyPath))
            {
                if (root == null) throw new InvalidOperationException("无法写入 WPS 白名单。");
                root.SetValue(Program.AddinName, "", RegistryValueKind.String);
                Log("已加入 WPS 白名单：" + Program.AddinName);
            }
        }

        private void RemoveFromAddinsWl()
        {
            using (RegistryKey root = Registry.CurrentUser.OpenSubKey(Program.AddinsWlKeyPath, true))
                if (root != null) root.DeleteValue(Program.AddinName, false);
            Log("已从 WPS 白名单移除本加载项。");
        }

        // ---------------- 官方标准安装/卸载（VSTOInstaller） ----------------

        private int RunStandardInstall(string tool, string dir)
        {
            // The UI stays responsive while this worker waits for the official dialog.
            // A timeout is never interpreted as successful installation.
            Log("即将打开微软 VSTO 授权窗口（标题含 Microsoft Office），目标仍为 " + Id.HostDisplayName + "；请在该窗口点「安装」或「取消」。");
            int code = LaunchVsto(tool, "/Install \"" + VstoUrlOf(dir) + "\"", true);
            if (code == -400)
            {
                Log("返回 -400：重新核对本加载项订阅后重试一次。");
                InstallResult clean = RemoveOwnedSubscriptions(tool, Program.PayloadVstoName);
                if (clean.Status != InstallStatus.Success) return clean.ExitCode ?? int.MinValue;
                code = LaunchVsto(tool, "/Install \"" + VstoUrlOf(dir) + "\"", true);
            }
            return code;
        }

        private InstallResult RemoveOwnedSubscriptions(string tool, string manifestName)
        {
            List<string> urls = new List<string>();
            using (RegistryKey root = Registry.CurrentUser.OpenSubKey(Program.VstaSolutionsKeyPath))
            {
                if (root != null)
                    foreach (string name in root.GetSubKeyNames())
                        using (RegistryKey key = root.OpenSubKey(name))
                        {
                            string url = key == null ? null : key.GetValue("Url") as string;
                            if (RuntimeLocator.IsOwnedUrl(url, manifestName) && !urls.Contains(url)) urls.Add(url);
                        }
            }
            return SubscriptionCleanup.Remove(urls, manifestName, delegate(string url) {
                return LaunchVsto(tool, "/Uninstall \"" + url + "\" /Silent", false);
            }, Log);
        }

        /// <summary>
        /// 返回正在运行的 WPS 进程描述（如 "wps.exe ×3"），没在跑就返回 null。
        /// WPS 是单实例且后台常驻，只看主窗口会漏判，必须数进程。
        /// </summary>
        private static string RunningWps()
        {
            string target = null;
            foreach (string n in new string[] { "wps", "et", "wpp" })
            {
                int cnt = Process.GetProcessesByName(n).Length;
                if (cnt > 0)
                    target = (target == null ? "" : target + "、")
                        + n + ".exe ×" + cnt;
            }
            return target;
        }

        private int LaunchVsto(string tool, string args, bool interactive)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = tool; psi.Arguments = args; psi.UseShellExecute = false;
            // Do not hide the official trust prompt; only subscription removal is silent.
            psi.CreateNoWindow = !interactive;
            Log("  调用：" + tool + " " + args);
            using (Process process = Process.Start(psi))
            {
                if (process == null) throw new InvalidOperationException("无法启动 VSTOInstaller。");
                process.WaitForExit();
                Log("  VSTOInstaller 已结束，返回 " + process.ExitCode);
                return process.ExitCode;
            }
        }

        private void DeletePayload(string dir, IEnumerable<string> payload)
        {
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) return;
            foreach (string name in payload)
            {
                string file = Path.Combine(dir, name);
                if (File.Exists(file)) { File.Delete(file); Log("  已删除 " + file); }
            }
            if (Directory.GetFileSystemEntries(dir).Length == 0) {
                Directory.Delete(dir); Log("  已移除空目录 " + dir);
            }
        }

        private InstallResult CleanupLegacyInstallation(string tool)
        {
            if (string.IsNullOrEmpty(Program.LegacyAddinName))
                return new InstallResult(InstallStatus.Success, "无需旧名迁移。", 0);
            string oldDir = ReadRegisteredDir(Program.LegacyAddinKeyPath);
            InstallResult result = RemoveOwnedSubscriptions(tool, Program.LegacyVstoName);
            if (result.Status != InstallStatus.Success) return result;
            // Read the directory before official uninstallation can remove its Addins key.
            DeletePayload(oldDir, Program.LegacyPayload);
            string defaultDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", Program.LegacyAddinName);
            if (!string.Equals(oldDir, defaultDir, StringComparison.OrdinalIgnoreCase))
                DeletePayload(defaultDir, Program.LegacyPayload);
            Registry.CurrentUser.DeleteSubKeyTree(Program.LegacyAddinKeyPath, false);
            using (RegistryKey white = Registry.CurrentUser.OpenSubKey(Program.AddinsWlKeyPath, true))
                if (white != null) white.DeleteValue(Program.LegacyAddinName, false);
            RemoveOwnedInclusion(Program.LegacyVstoName);
            Registry.CurrentUser.DeleteSubKeyTree(Program.LegacySetupKeyPath, false);
            return new InstallResult(InstallStatus.Success, "旧名迁移检查完成。", 0);
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            _btnInstall.Enabled = !busy && !_uninstallMode; _btnUninstall.Enabled = !busy;
            _btnBrowse.Enabled = !busy && !_uninstallMode; _btnClose.Enabled = !busy; _txtPath.Enabled = !busy;
            Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        }

        private void Log(string msg)
        {
            string line = DateTime.Now.ToString("HH:mm:ss") + " " + msg + "\r\n";
            try {
                lock (_logLock) {
                    Directory.CreateDirectory(Path.GetDirectoryName(_logPath));
                    File.AppendAllText(_logPath, line, new UTF8Encoding(false));
                }
            }
            catch (Exception ex) { line += "日志文件写入失败：" + ex.Message + "\r\n"; }
            if (_log.InvokeRequired) {
                if (!IsDisposed) _log.BeginInvoke(new Action<string>(AppendLog), line);
            }
            else AppendLog(line);
        }

        private void AppendLog(string line)
        {
            if (!_log.IsDisposed) _log.AppendText(line);
        }
    }
}
