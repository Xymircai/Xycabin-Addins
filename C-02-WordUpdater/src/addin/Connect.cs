using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WordUpdater
{
    // =====================================================================
    // 功能区与业务调度（VSTO 版）
    //
    //  - 实现 Microsoft.Office.Core.IRibbonExtensibility
    //    （该接口内嵌于 Microsoft.Office.Tools.Common.v4.0.Utilities.dll，
    //      IID 000C0396-0000-0000-C000-000000000046，与 WPS/Office 一致）
    //  - ThisAddIn.CreateRibbonExtensibilityObject() 返回本类实例，
    //    vstor 经其 GetCustomUI 取功能区 XML，onAction / getImage 经
    //    IDispatch 晚期绑定回调（AutoDispatch 保证 GetIDsOfNames 可解析）
    //  - WPS Application 经 Marshal.GetActiveObject 懒获取并探活，
    //    OnShutdown 时由 ThisAddIn 置空
    // =====================================================================
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class Connect : Microsoft.Office.Core.IRibbonExtensibility
    {
        private static object _app;
        private static bool _busy;
        private static LinkViewForm _viewForm;   // 「查看链接」窗口（非模态，全程只开一个）

        // 标黄“闪一下”的多槽容器（「更新链接」批量 + 「查看链接」单条共用）。
        // ★ 两套机制必须共用同一个 session：否则各自把自己的 Timer 到点时
        //   把【对方刚打上的黄】读成"原值"，最后还原成 wdYellow，黄色就永久
        //   留在用户文档里了。共用后由 Add() 按引用去重、保留最早的原值。
        private static FlashSession _flash;

        public static FlashSession SharedFlash
        {
            get
            {
                if (_flash == null) _flash = new FlashSession();
                return _flash;
            }
        }

        /// <summary>还原还没到期的标黄（新一轮开始前 / 加载项卸载时调用）。</summary>
        public static void ClearPendingFlash()
        {
            try
            {
                if (_flash != null)
                {
                    _flash.ClearNow();
                    _flash = null;
                }
            }
            catch { }
        }

        // ----------------------- WPS Application -----------------------

        /// <summary>VSTO 入口在 OnShutdown 时置空；按钮点击时也可惰性重建。</summary>
        public static object HostApplication
        {
            get { return _app; }
            set { _app = value; }
        }

        /// <summary>确保拿到运行中的 WPS 文字 Application；拿不到返回 false。</summary>
        public static bool EnsureApplication()
        {
            if (IsAlive(_app)) return true;
            _app = null;
            string[] progIds = { "WPS.Application", "KWPS.Application", "Word.Application" };
            for (int i = 0; i < progIds.Length; i++)
            {
                try
                {
                    object o = Marshal.GetActiveObject(progIds[i]);
                    if (o != null)
                    {
                        _app = o;
                        Log("EnsureApplication: 经 " + progIds[i] + " 获取成功");
                        return true;
                    }
                }
                catch { }
            }
            Log("EnsureApplication: 未找到运行中的 WPS 实例");
            return false;
        }

        private static bool IsAlive(object app)
        {
            if (app == null) return false;
            object documents = null;
            try
            {
                // 轻量属性访问探活（COM 对象已分离会抛异常）
                dynamic d = app;
                documents = d.Documents;
                return documents != null;
            }
            catch { return false; }
            finally { ReleaseTemporary(documents); }
        }

        private static void ReleaseTemporary(object value)
        {
            if (value == null || !Marshal.IsComObject(value)) return;
            try { Marshal.ReleaseComObject(value); }
            catch (Exception ex) { Log("临时 COM 引用释放失败：" + ex.Message); }
        }

        // ------------------------ IRibbonExtensibility ------------------------

        // Ribbon XML 由 配置器脚本 依据脚本顶部的「按钮配置区」生成，
        // 输出到 src/addin/Ribbon.g.cs —— 改标题/提示/顺序请改配置后重跑脚本，
        // 不要手改本文件（回归：本类只负责把回调转发给具体功能）。
        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbonId=" + (RibbonID ?? ""));
            return RibbonCustomization.Effective(true).OuterXml;
        }

        // --------------------------- Ribbon 回调 ---------------------------

        // ------------------ 图标回调（每按钮一个专属入口） ------------------
        //
        // 每个按钮在 Ribbon XML 里都有【专属 getImage 回调名】（GetImage_<按钮id>），
        // 函数名即按钮身份 —— 不再读 control.Id（2026-09-26 日志实测：网络盘
        // 文档的回调代理按会话间歇性读不出 Id，靠 id 分发会把所有按钮认成同一个，
        // 图标全白/全一样）。本地与网络盘文档下表现一致。
        //
        // 新增按钮 = 配置器脚本 加条目 + 这里加一行包装（配置器会校验，漏加当场报错）。

        /// <summary>统一的图标分发：优先用配置里的自定义图片，没有则按配置回落内置图标。</summary>
        private static IPictureDisp GetButtonImageCore(string id)
        {
            try
            {
                IPictureDisp custom = RibbonConfig.GetImage(id);
                if (custom != null) return custom;
                string kind = RibbonConfig.BuiltinIconKind(id);
                if (kind != RibbonConfig.BuiltinUpdate && kind != RibbonConfig.BuiltinPaste
                    && kind != RibbonConfig.BuiltinView && kind != RibbonConfig.BuiltinMail && kind != RibbonConfig.BuiltinTable && kind != RibbonConfig.BuiltinFormula && kind != RibbonConfig.BuiltinFilter && kind != RibbonConfig.BuiltinLink && kind != RibbonConfig.BuiltinDirectory && kind != RibbonConfig.BuiltinAgent && kind != RibbonConfig.BuiltinWeb && kind != RibbonConfig.BuiltinDocument && kind != RibbonConfig.BuiltinSettings)
                    kind = RibbonConfig.BuiltinUpdate;
                return MakeRibbonIcon(kind, id == "btnFeedback" ? Ui.Amber : Ui.Blue, RibbonConfig.GetImageSize(id));
            }
            catch (Exception ex)
            {
                Log("GetButtonImageCore(" + (id ?? "") + ") 异常：" + ex.Message);
                return null;
            }
        }

        // ------------------ 网址回调（每个 url 按钮一个专属入口） ------------------
        //
        // 与图标同理：onAction="OnOpenUrl_<按钮id>"，回调名带来按钮身份。
        // 将来新增网址按钮时同样只加一行包装，不需要任何"唯一网址兜底"。

        /// <summary>统一的打开网页/邮箱分发（行为与地址在配置区定义）。</summary>
        private static void OpenUrlCore(string id)
        {
            try
            {
                string url = RibbonConfig.GetUrl(id);
                if (string.IsNullOrEmpty(url))
                {
                    Log("OnOpenUrl：按钮 " + (id ?? "") + " 未配置网址");
                    return;
                }
                Log("OnOpenUrl -> " + id + " " + url);
                System.Diagnostics.Process.Start(url);
            }
            catch (Exception ex)
            {
                Log("OnOpenUrl 异常：" + ex.Message);
                try
                {
                    UiDialog.Show("无法打开配置的网址。\r\n\r\n错误信息：" + ex.Message,
                        "WordUpdater", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch { }
            }
        }

        private static bool _idFailLogged;

        private static string ControlId(object control)
        {
            // 2026-09-26 实测根因：打开【网络盘/云同步映射盘】上的 docx 时，
            // WPS 传给回调的 control 代理对象读 Id 会失败或返回空——于是
            // getImage 拿到空 id 返回 null（图标白）、onAction 拿到空 id
            // 找不到网址（点了没反应）。本地文件完全正常。
            // 这里多试几种读法；都失败就把真实异常记进日志（只记首次防刷屏）。
            Exception last = null;
            try
            {
                dynamic c = control;
                string id = (string)c.Id;
                if (!string.IsNullOrEmpty(id)) return id;
            }
            catch (Exception ex) { last = ex; }
            try
            {
                object v = control.GetType().InvokeMember("Id",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null, control, new object[0]);
                string id = v as string;
                if (!string.IsNullOrEmpty(id)) return id;
            }
            catch (Exception ex) { last = ex; }
            if (!_idFailLogged)
            {
                _idFailLogged = true;
                string tname;
                try { tname = control.GetType().FullName; }
                catch { tname = "(unknown)"; }
                Log("ControlId 读取失败（control 类型 " + tname + "）："
                    + (last == null ? "Id 返回空串" : last.GetType().Name + " - " + last.Message));
            }
            return "";
        }

        // --------------------------- 「查看链接」回调 ---------------------------

        // --------------------------- 「全选表格」回调 ---------------------------

        /// <summary>卸载/关文档时关掉还开着的查看窗口，避免它挂在已死的宿主上。</summary>
        public static void CloseViewForm()
        {
            LinkViewForm f = _viewForm;
            _viewForm = null;
            if (f == null) return;
            try
            {
                if (!f.IsDisposed && f.IsHandleCreated)
                    f.BeginInvoke((MethodInvoker)delegate { try { f.Close(); } catch { } });
            }
            catch
            {
                try { f.Close(); }
                catch { }
            }
        }

        /// <summary>尽力拿 WPS 主窗口句柄，给非模态窗当 Owner 用。</summary>
        private static IntPtr GetHostWindowHandle()
        {
            object window = null;
            try
            {
                window = ((dynamic)_app).ActiveWindow;
                IntPtr handle = HostDialogService.ValidateOwner(new IntPtr(Convert.ToInt64(((dynamic)window).Hwnd)));
                if (handle != IntPtr.Zero) return handle;
            }
            catch { }
            finally { ReleaseTemporary(window); }
            try
            {
                IntPtr handle = HostDialogService.ValidateOwner(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
                if (handle != IntPtr.Zero) return handle;
                return HostDialogService.ValidateOwner(NativeOle.GetForegroundWindow());
            }
            catch { return IntPtr.Zero; }
        }

        /// <summary>把 Win32 句柄包成 WinForms 认的 IWin32Window。</summary>
        private sealed class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private readonly IntPtr _h;
            public WindowWrapper(IntPtr h) { _h = h; }
            public IntPtr Handle { get { return _h; } }
        }

        // --------------------------- 「链接Excel」回调 ---------------------------

        // --------------------------- 链接粘贴结果文案 ---------------------------

        // --------------------------- 结果文案 ---------------------------

        // --------------------------- 按钮图标 ---------------------------

        private static IPictureDisp MakeRibbonIcon(string kind, Color accent, int logicalSize)
        {
            return RibbonImages.Create(logicalSize,
                delegate(int width, int height) { return UiIcons.CreateRibbon(kind, accent, width, height); }, Log);
        }

        // --------------------------- 日志 ---------------------------

        private static readonly object _logLock = new object();

        public static void Log(string msg)
        {
            try
            {
                // 固定日志位置：VSTO 加载时 Assembly.Location 指向 CLR 的
                // dl3 影子缓存而非部署目录，因此写到 %LocalAppData% 下。
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordUpdater");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "WordUpdater.log");
                lock (_logLock)
                {
                    FileInfo fi = new FileInfo(path);
                    if (fi.Exists && fi.Length > 512 * 1024)
                        File.WriteAllText(path, "");
                    File.AppendAllText(path,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  " + msg + "\r\n");
                }
            }
            catch { }
        }
        public IPictureDisp GetImage_menuBatchProcessing(object control) { return GetButtonImageCore("menuBatchProcessing"); }
    }
}
