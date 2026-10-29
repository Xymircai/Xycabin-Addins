using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using Id = WpsShared.ProjectInfo;
using System.Collections;
using System.Reflection;
using System.Text.RegularExpressions;

namespace WpsShared.Tools
{
    // =====================================================================
    // RibbonProbe —— 按钮配置离线自检工具
    //
    // 用法：build\RibbonProbe.exe build\<程序集名>.dll
    //
    // 不改 AA-Program 配置器的配置区也能看看：
    //   1. 生成出来的功能区 XML 长什么样（选项卡/分组/按钮/提示文案）
    //   2. 每个按钮最终会用什么图标（自定义图 -> IPictureDisp，还是内置矢量图标）
    //   3. action="url" 的按钮最终会打开哪个地址
    // =====================================================================
    internal static class RibbonProbe
    {
        private static MethodInfo _getImage;
        private static MethodInfo _getUrl;
        private static MethodInfo _builtinKind;

        [STAThread]
        public static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            try
            {
                if (args.Length > 2 && args[0] == "--image-worker") return ImageWorker(args[1], args[2]);
                string target = args.Length > 0
                    ? args[0]
                    : System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                             Id.AssemblyName + ".dll");
                Assembly asm = Assembly.LoadFrom(target);
                VerifyImages(asm, Path.GetFullPath(target));
                Type cfg = asm.GetType(Id.AddinNamespace + ".RibbonConfig", true);

                string xml = (string)cfg.GetProperty("RibbonXml",
                    BindingFlags.Public | BindingFlags.Static).GetValue(null, null);
                _getImage = cfg.GetMethod("GetImage",
                    BindingFlags.Public | BindingFlags.Static);
                _getUrl = cfg.GetMethod("GetUrl",
                    BindingFlags.Public | BindingFlags.Static);
                _builtinKind = cfg.GetMethod("BuiltinIconKind",
                    BindingFlags.Public | BindingFlags.Static);

                Console.WriteLine("==================== 功能区 XML ====================");
                XmlDocument document = new XmlDocument(); document.LoadXml(xml);
                Console.WriteLine(xml);
                Console.WriteLine();

                Console.WriteLine("==================== 按钮自检 ====================");
                MatchCollection buttons = Regex.Matches(xml,
                    "<button\\s+id=\"([^\"]+)\"\\s+label=\"([^\"]*)\"[^>]*?/>",
                    RegexOptions.Singleline);
                if (buttons.Count == 0)
                {
                    Console.WriteLine("(功能区里没有任何按钮)");
                    return 1;
                }

                MatchCollection groups = Regex.Matches(xml,
                    "<group\\s+id=\"([^\"]+)\"\\s+label=\"([^\"]*)\">");
                Console.WriteLine("分组：" + groups.Count + " 个，按钮：" + buttons.Count + " 个");
                Console.WriteLine();

                bool bad = false;
                int idx = 0;
                foreach (Match m in buttons)
                {
                    idx++;
                    string id = m.Groups[1].Value;
                    string label = XmlText(m.Groups[2].Value);
                    string tip = Attr(xml, id, "screentip");
                    string action = Attr(xml, id, "onAction");

                    Console.WriteLine(idx + ". " + label + "  [" + id + "]");
                    Console.WriteLine("   onAction   : " + action);

                    // 专属回调名契约：getImage 必须是 "GetImage_<按钮id>"
                    // （函数名即按钮身份，不读 control.Id —— 网络盘文档的代理读不出）
                    string img = Attr(xml, id, "getImage");
                    if (!string.Equals(img, "GetImage_" + id, StringComparison.Ordinal))
                    {
                        Console.WriteLine("   getImage   : !! " + (img ?? "(缺失)")
                            + "（应为 GetImage_" + id + "）");
                        bad = true;
                    }

                    // url 按钮的 onAction 约定为 "OnOpenUrl_<按钮id>"
                    if (action != null && action.StartsWith("OnOpenUrl_", StringComparison.Ordinal))
                    {
                        string url = Invoke(_getUrl, id) as string;
                        if (string.IsNullOrEmpty(url))
                        {
                            Console.WriteLine("   url        : !! 未取到地址");
                            bad = true;
                        }
                        else
                        {
                            Console.WriteLine("   url        : " + url);
                        }
                    }

                    object pic = Invoke(_getImage, id);
                    string kind = Invoke(_builtinKind, id) as string;
                    string dims = PictureSize(pic);
                    if (pic != null)
                    {
                        Console.WriteLine("   图标       : 自定义图片" + dims);
                    }
                    else if (!string.IsNullOrEmpty(kind))
                    {
                        Console.WriteLine("   图标       : 内置矢量图标 " + kind);
                    }
                    else
                    {
                        Console.WriteLine("   图标       : !! 没有任何图标（既不配自定图也没配内置图）");
                        bad = true;
                    }

                    string oneLine = (tip ?? "").Replace("\r", " ").Replace("\n", " ");
                    Console.WriteLine("   悬停提示   : " + (oneLine.Length > 0 ? oneLine : "(无)"));
                    Console.WriteLine();
                }

                if (bad)
                {
                    Console.WriteLine("!! 有按钮配置不完整，见上面标 !! 的行");
                    return 1;
                }
                Console.WriteLine("全部按钮自检通过。");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("自检失败：" + ex.Message);
                return 2;
            }
        }

        private static object Invoke(MethodInfo mi, string id)
        {
            if (mi == null)
            {
                Console.Error.WriteLine("!! 反射找不到方法（Prob e内部错误）");
                return null;
            }
            try
            {
                return mi.Invoke(null, new object[] { id });
            }
            catch (Exception ex)
            {
                Exception real = ex is TargetInvocationException
                    ? ex.InnerException : ex;
                Console.Error.WriteLine("!! " + mi.Name + "(\"" + id + "\") 调用失败："
                                        + real.Message);
                if (real.InnerException != null)
                    Console.Error.WriteLine("   内层：" + real.InnerException.ToString());
                return null;
            }
        }

        private static string PictureSize(object pic)
        {
            if (pic == null) return "";
            SizeF size = LogicalSize(pic);
            return "（" + size.Width.ToString("F2") + "x" + size.Height.ToString("F2") + " 逻辑像素）";
        }

        private const BindingFlags StaticAll = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
        [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")] private static extern uint GetGuiResources(IntPtr process, uint flags);
        [DllImport("gdi32.dll", EntryPoint = "GetObjectW")] private static extern int GetObject(IntPtr handle, int size, out NativeBitmap bitmap);
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeBitmap { public int Type, Width, Height, WidthBytes; public ushort Planes, BitsPixel; public IntPtr Bits; }
        [ComImport, Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface PictureMetrics
        {
            [PreserveSig] int GetHandle(out int value);
            [PreserveSig] int GetPalette(out int value);
            [PreserveSig] int GetPictureType(out short value);
            [PreserveSig] int GetWidth(out int value);
            [PreserveSig] int GetHeight(out int value);
        }
        private static SizeF LogicalSize(object picture)
        {
            int width, height; PictureMetrics p = (PictureMetrics)picture;
            Marshal.ThrowExceptionForHR(p.GetWidth(out width)); Marshal.ThrowExceptionForHR(p.GetHeight(out height));
            return new SizeF((float)(width * 96.0 / 2540.0), (float)(height * 96.0 / 2540.0));
        }
        private static IntPtr PictureHandle(object picture)
        { int handle; Marshal.ThrowExceptionForHR(((PictureMetrics)picture).GetHandle(out handle)); return new IntPtr(unchecked((long)(uint)handle)); }
        private static void Assert(bool condition, string message)
        { if (!condition) throw new InvalidOperationException(message); }
        private static void CheckSize(object picture, int expected, string label)
        {
            Assert(picture != null, label + " returned no picture"); SizeF logical = LogicalSize(picture);
            NativeBitmap bitmap;
            Assert(GetObject(PictureHandle(picture), Marshal.SizeOf(typeof(NativeBitmap)), out bitmap) != 0, label + " invalid HBITMAP");
            Console.WriteLine(label + ": physical=" + bitmap.Width + "x" + bitmap.Height + " logical=" + logical.Width.ToString("F2") + "x" + logical.Height.ToString("F2"));
            Assert(Math.Abs(logical.Width - expected) <= 1 && Math.Abs(logical.Height - expected) <= 1, label + " incorrect logical dimensions");
        }
        private static void ReleasePicture(object picture)
        { if (picture != null && Marshal.IsComObject(picture)) Marshal.FinalReleaseComObject(picture); }

        private static void ValidateTree(XmlDocument document, XmlNamespaceManager ns)
        {
            HashSet<string> ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XmlElement control in document.SelectNodes("//*[@id]"))
                Assert(ids.Add(control.GetAttribute("id")), "Duplicate Ribbon id");
            Assert(document.SelectNodes("//r:box", ns).Count == 0, "Direct native group rows");
            XmlNodeList groups = document.SelectNodes("//r:group", ns);
            Assert(groups.Count == 6, "Four primary groups and two Agent-Evo groups");
            Assert(((XmlElement)groups[3]).GetAttribute("label") == "扩展" && ((XmlElement)groups[4]).GetAttribute("label") == "详情" && ((XmlElement)groups[5]).GetAttribute("label") == "操作", "Extension detail and operation groups");
            Assert(document.SelectNodes("//r:tab", ns).Count == 2, "Two tabs");
            XmlElement extra = (XmlElement)document.SelectSingleNode("//r:tab[@id='tabAgentEvo']", ns);
            Assert(extra.GetAttribute("label") == "Agent-Evo" && extra.GetAttribute("getVisible") == "GetAgentTabVisible", "Extra tab visibility");
            string[] groupLabels = Id.AssemblyName == "Excel_Iterate" ? new string[] { "数据", "操作", "反馈" } : new string[] { "链接", "操作", "反馈" };
            for (int i = 0; i < 3; i++) Assert(((XmlElement)groups[i]).GetAttribute("label") == groupLabels[i], "Group label/order");
            XmlElement group = (XmlElement)groups[0];
            string[] rowIds = Id.AssemblyName == "Excel_Iterate" ? new string[] { "btnTracePrecedents", "btnTraceDependents", "menuFilterEnhanced", "menuAuditToolbox" } : new string[] { "btnUpdateLinks", "btnLinkToExcel", "btnViewLinks" };
            Assert(group.ChildNodes.Count == rowIds.Length, "Native small rows");
            string[] rowSizes = Id.AssemblyName == "Excel_Iterate" ? new string[] { "normal", "normal", "normal", "large" } : new string[] { "normal", "normal", "normal" };
            for (int i = 0; i < rowIds.Length; i++) { XmlElement row = (XmlElement)group.ChildNodes[i]; Assert(row.GetAttribute("id") == rowIds[i] && row.GetAttribute("size") == rowSizes[i], "Native row " + rowIds[i]); }
            int[] rootCounts = Id.AssemblyName == "Excel_Iterate" ? new int[] { 4, 1 } : new int[] { 2, 1 };
            for (int i = 1; i < 3; i++) { Assert(groups[i].ChildNodes.Count == rootCounts[i - 1], "Action/feedback root counts"); foreach (XmlElement item in groups[i].ChildNodes) Assert(item.GetAttribute("size") == "large", "Large action/feedback root"); }
            if (Id.AssemblyName != "Excel_Iterate") Assert(((XmlElement)group.ChildNodes[1]).GetAttribute("label") == "链接Excel", "Word new label");
            foreach (XmlElement menu in document.SelectNodes("//r:menu", ns))
            {
                Assert(menu.GetAttribute("itemSize") == "normal" && !menu.HasAttribute("onAction"), "Native menu contract");
                Assert(menu.ChildNodes.Count > 0, "Empty menu");
                foreach (XmlElement child in menu.ChildNodes)
                    Assert((child.LocalName == "menu" || child.LocalName == "button") && !child.HasAttribute("size"), "Unsized nested control");
            }
            // PDF 工具接入后「文件操作」改为两层：Excel 含批量处理 + 文件转换，Word 只有文件转换。
            bool excelHost = Id.AssemblyName == "Excel_Iterate";
            XmlElement files=(XmlElement)document.SelectSingleNode("//r:menu[@id='menuFileOperations']",ns);
            Assert(files.GetAttribute("label")=="文件操作"&&((XmlElement)files.ParentNode).GetAttribute("id")=="grpAction"&&files.ChildNodes.Count==(excelHost?2:1),"File operations large menu");
            if(excelHost)
            {
                XmlElement batch=(XmlElement)files.ChildNodes[0];
                Assert(batch.GetAttribute("id")=="menuBatchProcessing"&&batch.GetAttribute("label")=="批量处理"&&batch.ChildNodes.Count==3,"Batch processing submenu");
                Assert(((XmlElement)batch.ChildNodes[0]).GetAttribute("onAction")=="OnOrganizeFiles"&&((XmlElement)batch.ChildNodes[1]).GetAttribute("onAction")=="OnRenameFiles"&&((XmlElement)batch.ChildNodes[2]).GetAttribute("onAction")=="OnFileNames","File operation callback order");
            }
            XmlElement convert=(XmlElement)files.LastChild;
            Assert(convert.GetAttribute("id")=="menuFileConversion"&&convert.GetAttribute("label")=="PDF 工具"&&convert.ChildNodes.Count==2,"PDF conversion submenu");
            string[] pdfActions={"OnCurrentPdf","OnWatermarkPdf"};
            for(int i=0;i<pdfActions.Length;i++) Assert(((XmlElement)convert.ChildNodes[i]).GetAttribute("onAction")==pdfActions[i],"PDF conversion callback order");
            if (Id.AssemblyName == "Excel_Iterate")
            {
                XmlElement root = (XmlElement)document.SelectSingleNode("//r:menu[@id='menuOneClickFormula']", ns);
                Assert(root != null && root.GetAttribute("label") == "一键公式" && root.GetAttribute("size") == "large", "Formula root menu");
                Assert(((XmlElement)root.ParentNode).GetAttribute("id") == "grpAction", "Formula action group");
                Assert(root.ChildNodes.Count == 3 && root.ChildNodes[0].LocalName == "menu" && root.ChildNodes[1].LocalName == "button", "Formula first level");
                Assert(((XmlElement)root.ChildNodes[0]).GetAttribute("label") == "进行公式封装" && ((XmlElement)root.ChildNodes[1]).GetAttribute("label") == "移除公式封装", "Formula first-level labels");
                XmlElement wrapping = (XmlElement)root.ChildNodes[0];
                Assert(wrapping.ChildNodes.Count == 4, "Three wrapping functions and custom function");
                XmlElement abs = (XmlElement)wrapping.SelectSingleNode("r:button[@id='btnWrapAbs']", ns);
                Assert(abs != null && abs.GetAttribute("label") == "ABS封装" && abs.GetAttribute("onAction") == "OnWrapAbs", "ABS leaf unchanged");
                XmlElement customFunction = (XmlElement)wrapping.SelectSingleNode("r:button[@id='btnCustomFunction']", ns);
                Assert(customFunction != null && customFunction.GetAttribute("label") == "自定义函数" && customFunction.GetAttribute("onAction") == "OnCustomFunction", "Custom function leaf unchanged");
                Assert(wrapping.LastChild == customFunction, "Custom function sits after ABS");
                string[] functionIds = { "menuIFErrorWrap", "menuRoundWrap" };
                string[] labels = { "IFERROR封装", "ROUND封装" };
                string[] defaults = { "OnWrapIFError", "OnWrapRound" }, custom = { "OnCustomIFError", "OnCustomRound" };
                for (int i = 0; i < labels.Length; i++)
                {
                    // Sibling order belongs to 配置器脚本; validate identity and wiring independently.
                    XmlElement function = (XmlElement)wrapping.SelectSingleNode("r:menu[@id='" + functionIds[i] + "']", ns);
                    Assert(function != null && function.GetAttribute("label") == labels[i] && function.ChildNodes.Count == 2, "Function submenu " + labels[i]);
                    XmlElement standard = (XmlElement)function.ChildNodes[0], editor = (XmlElement)function.ChildNodes[1];
                    Assert(standard.GetAttribute("label") == (i == 0 ? "默认返回空" : "默认保留2位") && standard.GetAttribute("onAction") == defaults[i], "Default execution " + labels[i]);
                    Assert(editor.GetAttribute("label") == "自定义封装" && editor.GetAttribute("onAction") == custom[i], "Custom wrapping " + labels[i]);
                }
                Assert(((XmlElement)root.ChildNodes[1]).GetAttribute("onAction") == "OnUnwrapFormula", "Unwrap callback");
                XmlElement arithmetic = (XmlElement)root.ChildNodes[2];
                Assert(arithmetic.GetAttribute("id") == "menuArithmetic" && arithmetic.ChildNodes.Count == 4, "Arithmetic tree");
                string[] operators = { "Add", "Subtract", "Multiply", "Divide" };
                for (int i = 0; i < 4; i++)
                {
                    XmlElement branch = (XmlElement)arithmetic.ChildNodes[i];
                    Assert(branch.GetAttribute("id") == "menuCalculate" + operators[i] && branch.ChildNodes.Count == 2, "Arithmetic branch");
                    Assert(((XmlElement)branch.FirstChild).GetAttribute("onAction") == "OnCalculate" + operators[i] && ((XmlElement)branch.LastChild).GetAttribute("onAction") == "OnCustomCalculate" + operators[i], "Arithmetic leaf wiring");
                    Assert(((XmlElement)branch.LastChild).GetAttribute("label") == "自定义计算", "Custom arithmetic label");
                }
                XmlElement filter = (XmlElement)document.SelectSingleNode("//r:menu[@id='menuFilterEnhanced']", ns);
                Assert(filter.ChildNodes.Count == 2 && ((XmlElement)filter.FirstChild).GetAttribute("onAction") == "OnFilterColumns" && ((XmlElement)filter.LastChild).GetAttribute("onAction") == "OnQuickAdvancedFilter", "Column filter leaf");
            }
            else
            {
                XmlElement action=(XmlElement)groups[1];
                Assert(action.ChildNodes.Count==2&&((XmlElement)action.ChildNodes[0]).GetAttribute("id")=="btnSelectAllTables"&&((XmlElement)action.ChildNodes[1]).GetAttribute("id")=="menuFileOperations","Word action group keeps select-all and file operations");
            }
            Console.WriteLine("Native Ribbon tree and unsized nested attributes passed");
        }

        private static int ImageWorker(string path, string mode)
        {
            // No parent OLE state is inherited: each mode runs in a fresh process.
            Assembly asm = Assembly.LoadFrom(Path.GetFullPath(path));
            Type native = asm.GetType(Id.AddinNamespace + ".NativeOle", true);
            bool awareFirst = mode != "unaware-aware";
            IntPtr previous = SetThreadDpiAwarenessContext(new IntPtr(awareFirst ? -4 : -1));
            Assert(previous != IntPtr.Zero, "Cannot select test DPI context");
            try
            {
                using (Bitmap sample = new Bitmap(32, 32, PixelFormat.Format24bppRgb))
                {
                    object first = native.GetMethod("FromBitmap", StaticAll).Invoke(null, new object[] { sample });
                    Console.WriteLine("First OLE context " + mode + ": " + LogicalSize(first)); ReleasePicture(first);
                }
                if (mode != "aware") Assert(SetThreadDpiAwarenessContext(new IntPtr(mode == "unaware-aware" ? -4 : -1)) != IntPtr.Zero, "DPI switch failed");
                Type cfg = asm.GetType(Id.AddinNamespace + ".RibbonConfig", true);
                Type connectType = asm.GetType(Id.AddinNamespace + ".Connect", true);
                object connect = Activator.CreateInstance(connectType);
                string xml = (string)cfg.GetProperty("RibbonXml", StaticAll).GetValue(null, null);
                XmlDocument document = new XmlDocument(); document.LoadXml(xml);
                XmlNamespaceManager ns = new XmlNamespaceManager(document.NameTable); ns.AddNamespace("r", document.DocumentElement.NamespaceURI);
                XmlNodeList buttons = document.SelectNodes("//r:button | //r:menu | //r:toggleButton", ns);
                Assert(buttons.Count > 0, "No Ribbon controls");
                ValidateTree(document, ns);
                IDictionary sizes = (IDictionary)cfg.GetField("Sizes", StaticAll).GetValue(null);
                IDictionary icons = (IDictionary)cfg.GetField("Icons", StaticAll).GetValue(null);
                IDictionary cache = (IDictionary)cfg.GetField("Cache", StaticAll).GetValue(null);
                Color plate = (Color)cfg.GetField("PlateColor", StaticAll).GetValue(null);
                string custom;
                using (Bitmap image = new Bitmap(40, 40, PixelFormat.Format32bppArgb))
                using (Graphics g = Graphics.FromImage(image))
                using (MemoryStream bytes = new MemoryStream())
                { g.Clear(Color.Transparent); g.FillRectangle(Brushes.Red, 10, 10, 20, 20); image.Save(bytes, ImageFormat.Png); custom = Convert.ToBase64String(bytes.ToArray()); }
                foreach (XmlElement button in buttons)
                {
                    string id = button.GetAttribute("id"); MethodInfo callback = connectType.GetMethod(button.GetAttribute("getImage"));
                    MethodInfo action = connectType.GetMethod(button.GetAttribute("onAction"));
                    Assert(callback != null && callback.GetParameters().Length == 1 && callback.ReturnType.GUID == typeof(PictureMetrics).GUID, "Image callback contract " + id);
                    Assert(button.LocalName == "menu" ? !button.HasAttribute("onAction") :
                        action != null && action.ReturnType == typeof(void) && action.GetParameters().Length == (button.LocalName == "toggleButton" ? 2 : 1), "Action callback contract " + id);
                    if (button.LocalName == "toggleButton")
                    {
                        Assert(action.GetParameters()[0].ParameterType == typeof(object) && action.GetParameters()[1].ParameterType == typeof(bool), "Toggle action signature");
                        MethodInfo getter = connectType.GetMethod(button.GetAttribute("getPressed"));
                        Assert(getter != null && getter.ReturnType == typeof(bool) && getter.GetParameters().Length == 1, "Toggle getter contract");
                    }
                    Assert(!button.HasAttribute("imageWidth") && !button.HasAttribute("imageHeight"), "Unsupported Ribbon image attributes");
                    bool nested = button.ParentNode.LocalName == "menu";
                    Assert(nested ? !button.HasAttribute("size") : button.GetAttribute("size") == "large" || button.GetAttribute("size") == "normal", "Invalid size");
                    int configuredSize = nested || button.GetAttribute("size") == "normal" ? 16 : 32;
                    Assert(Convert.ToInt32(sizes[id]) == configuredSize, "Generated image size " + id);
                    object configuredPicture = callback.Invoke(connect, new object[] { null });
                    try { CheckSize(configuredPicture, configuredSize, mode + " configured " + id); } finally { ReleasePicture(configuredPicture); }
                    if (Id.AssemblyName == "Excel_Iterate" && id.StartsWith("btnTrace", StringComparison.Ordinal))
                        Assert(button.GetAttribute("label") == (id == "btnTracePrecedents" ? "追踪引用" : "追踪从属"), "Excel label " + id);
                    object originalIcon = icons[id]; object originalSize = sizes[id];
                    try
                    {
                        foreach (int expected in new int[] { 32, 16 })
                        {
                            sizes[id] = expected; icons.Remove(id); cache.Clear();
                            object builtin = callback.Invoke(connect, new object[] { null });
                            try { CheckSize(builtin, expected, mode + " built-in " + id); } finally { ReleasePicture(builtin); }
                            icons[id] = custom;
                            object picture = callback.Invoke(connect, new object[] { null });
                            try
                            {
                                CheckSize(picture, expected, mode + " custom " + id);
                                using (Bitmap rendered = Image.FromHbitmap(PictureHandle(picture)))
                                {
                                    Color corner = rendered.GetPixel(0, 0), center = rendered.GetPixel(rendered.Width / 2, rendered.Height / 2);
                                    Assert(corner.ToArgb() == plate.ToArgb(), "Configured background changed " + id);
                                    Assert(center.R > 240 && center.G < 15 && center.B < 15, "Custom image did not take priority " + id);
                                }
                                Assert(ReferenceEquals(picture, callback.Invoke(connect, new object[] { null })), "Custom cache " + id);
                            }
                            finally { cache.Clear(); ReleasePicture(picture); }
                        }
                    }
                    finally { cache.Clear(); if (originalIcon == null) icons.Remove(id); else icons[id] = originalIcon; sizes[id] = originalSize; }
                }
                // The common configuration catalogue must work in both hosts,
                // even for kinds not currently used by that project's Ribbon.
                IDictionary catalogue = (IDictionary)cfg.GetField("BuiltinIcons", StaticAll).GetValue(null);
                object feedbackKind = catalogue["btnFeedback"], feedbackSize = sizes["btnFeedback"], feedbackIcon = icons["btnFeedback"];
                MethodInfo drawIcon = asm.GetType(Id.AddinNamespace + ".UiIcons").GetMethod("CreateRibbon");
                try
                {
                    icons.Remove("btnFeedback"); cache.Clear();
                    foreach (string kind in new string[] { "update", "paste", "mail", "view", "table", "formula", "filter", "link", "directory", "agent", "web", "document", "settings" })
                    foreach (int expectedSize in new int[] { 16, 32 })
                    {
                        catalogue["btnFeedback"] = kind; sizes["btnFeedback"] = expectedSize;
                        object picture = connectType.GetMethod("GetImage_btnFeedback").Invoke(connect, new object[] { null });
                        try
                        {
                            CheckSize(picture, expectedSize, mode + " catalogue " + kind);
                            using (Bitmap actual = Image.FromHbitmap(PictureHandle(picture)))
                            using (Bitmap expected = (Bitmap)drawIcon.Invoke(null, new object[] { kind, Color.FromArgb(245,158,11), actual.Width, actual.Height }))
                            {
                                bool matches = true;
                                for (int y = 0; y < actual.Height && matches; y++) for (int x = 0; x < actual.Width; x++)
                                    if (actual.GetPixel(x,y).ToArgb() != expected.GetPixel(x,y).ToArgb()) { matches = false; break; }
                                Assert(matches, "Catalogue glyph dispatch " + kind);
                            }
                        }
                        finally { ReleasePicture(picture); }
                    }
                }
                finally { catalogue["btnFeedback"] = feedbackKind; sizes["btnFeedback"] = feedbackSize; if(feedbackIcon != null)icons["btnFeedback"] = feedbackIcon; }
                if (Id.AssemblyName == "WordUpdater")
                {
                    IDictionary builtins = (IDictionary)cfg.GetField("BuiltinIcons", StaticAll).GetValue(null);
                    object original = builtins["btnViewLinks"];
                    object updateSize = sizes["btnUpdateLinks"];
                    sizes["btnUpdateLinks"] = sizes["btnViewLinks"];
                    object expected = connectType.GetMethod("GetImage_btnUpdateLinks").Invoke(connect, new object[] { null });
                    object actual = null;
                    try
                    {
                        builtins["btnViewLinks"] = "unknown-kind";
                        actual = connectType.GetMethod("GetImage_btnViewLinks").Invoke(connect, new object[] { null });
                        using (Bitmap a = Image.FromHbitmap(PictureHandle(actual)))
                        using (Bitmap b = Image.FromHbitmap(PictureHandle(expected)))
                        {
                            Assert(a.Size == b.Size, "Unknown kind fallback dimensions");
                            for (int y = 0; y < a.Height; y++) for (int x = 0; x < a.Width; x++)
                                Assert(a.GetPixel(x, y).ToArgb() == b.GetPixel(x, y).ToArgb(), "Unknown kind must keep update fallback");
                        }
                    }
                    finally { builtins["btnViewLinks"] = original; sizes["btnUpdateLinks"] = updateSize; ReleasePicture(actual); ReleasePicture(expected); }
                    Console.WriteLine("Unknown built-in kind keeps original update fallback");
                }
                using (Process current = Process.GetCurrentProcess())
                {
                    uint before = GetGuiResources(current.Handle, 0);
                    MethodInfo callback = connectType.GetMethod(((XmlElement)buttons[0]).GetAttribute("getImage"));
                    for (int i = 0; i < 60; i++)
                    {
                        object picture = callback.Invoke(connect, new object[] { null }); ReleasePicture(picture);
                    }
                    uint after = GetGuiResources(current.Handle, 0);
                    Assert(after <= before + 3, "HBITMAP leak: " + before + " -> " + after);
                    Console.WriteLine("Repeated image ownership: GDI " + before + " -> " + after);
                }
                Console.WriteLine("Image worker PASS: " + mode); return 0;
            }
            finally { SetThreadDpiAwarenessContext(previous); }
        }

        private class FakePicture { internal int Width, Height; internal bool Released; }
        private static void VerifyImages(Assembly asm, string path)
        {
            foreach (string mode in new string[] { "aware", "unaware-aware", "aware-unaware" })
            {
                ProcessStartInfo start = new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,
                    "--image-worker \"" + path + "\" " + mode);
                start.UseShellExecute = false; start.CreateNoWindow = true; start.RedirectStandardOutput = start.RedirectStandardError = true;
                start.StandardOutputEncoding = start.StandardErrorEncoding = Encoding.UTF8;
                using (Process process = Process.Start(start))
                {
                    var stdout = process.StandardOutput.ReadToEndAsync(); var stderr = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(20000)) { process.Kill(); throw new InvalidOperationException("Image worker timeout " + mode); }
                    Console.Write(stdout.Result); Assert(process.ExitCode == 0, mode + " failed: " + stderr.Result);
                }
            }
            MethodInfo core = asm.GetType(Id.AddinNamespace + ".RibbonImages", true).GetMethod("ConvertCore", StaticAll);
            int checks = 0;
            foreach (float dpi in new float[] { 96, 120, 144, 192 })
                foreach (int expected in new int[] { 16, 32 })
                {
                    double[] calibration = new double[2]; int made = 0, released = 0;
                    Func<int, int, Bitmap> render = delegate(int w, int h) { return new Bitmap(w, h, PixelFormat.Format24bppRgb); };
                    Func<Bitmap, object> wrap = delegate(Bitmap b) { made++; return new FakePicture { Width = b.Width, Height = b.Height }; };
                    Func<object, SizeF> metrics = delegate(object p) { FakePicture f = (FakePicture)p; return new SizeF((float)(Math.Round(f.Width * 2540.0 / dpi) * 96.0 / 2540.0), (float)(Math.Round(f.Height * 2540.0 / dpi) * 96.0 / 2540.0)); };
                    Action<object> release = delegate(object p) { FakePicture f = (FakePicture)p; Assert(!f.Released, "Double release"); f.Released = true; released++; };
                    List<string> logs = new List<string>(); Action<string> log = logs.Add;
                    FakePicture result = (FakePicture)core.Invoke(null, new object[] { expected, render, wrap, metrics, release, log, calibration });
                    SizeF size = metrics(result);
                    Assert(Math.Abs(size.Width - expected) <= 1 && Math.Abs(size.Height - expected) <= 1, "Simulated OLE " + dpi);
                    Assert(!result.Released && made == 2 && released == 1 && logs.Count == 0, "Calibration ownership " + dpi);
                    release(result); checks++;
                    // Stale calibration must verify the result and recover exactly once.
                    if (dpi == 144 && expected == 32)
                    {
                        made = released = 0; calibration = new double[] { 1, 1 }; logs.Clear();
                        result = (FakePicture)core.Invoke(null, new object[] { expected, render, wrap, metrics, release, log, calibration });
                        Assert(Math.Abs(metrics(result).Width - expected) <= 1 && made == 3 && released == 2 && logs.Count == 1, "Stale calibration retry"); release(result); checks++;
                    }
                    made = released = 0; calibration = new double[2]; logs.Clear();
                    Func<object, SizeF> fail = delegate { throw new InvalidOperationException("Injected metrics failure"); };
                    result = (FakePicture)core.Invoke(null, new object[] { expected, render, wrap, fail, release, log, calibration });
                    Assert(result.Width == expected && result.Height == expected && made == 3 && released == 2 && logs.Count == 3 && !result.Released, "Bounded calibration failure fallback");
                    release(result); checks++;
                }
            // Independent width/height mappings and a changed mapping during conversion.
            foreach (string scenario in new string[] { "anisotropic", "changed", "invalid-output", "wrap-failure" })
            {
                int made = 0, released = 0, reads = 0;
                Func<int, int, Bitmap> render = delegate(int w, int h) { return new Bitmap(w, h, PixelFormat.Format24bppRgb); };
                Func<Bitmap, object> wrap = delegate(Bitmap b)
                {
                    made++;
                    if (scenario == "wrap-failure" && b.Width == 32) throw new InvalidOperationException("Injected OLE creation failure");
                    return new FakePicture { Width = b.Width, Height = b.Height };
                };
                Func<object, SizeF> metrics = delegate(object p)
                {
                    reads++; FakePicture f = (FakePicture)p;
                    if (scenario == "invalid-output") return reads % 2 == 1 ? new SizeF(32, 32) : new SizeF(float.NaN, 0);
                    float x = scenario == "changed" && reads > 1 ? 2f : 1.5f;
                    float y = scenario == "anisotropic" ? 2f : x;
                    return new SizeF(f.Width / x, f.Height / y);
                };
                Action<object> release = delegate(object p) { FakePicture f = (FakePicture)p; Assert(!f.Released, "Double release " + scenario); f.Released = true; released++; };
                List<string> logs = new List<string>(); Action<string> log = logs.Add;
                int expected = scenario == "wrap-failure" ? 16 : 32;
                FakePicture result = (FakePicture)core.Invoke(null, new object[] { expected, render, wrap, metrics, release, log, new double[2] });
                Assert(!result.Released && made - released == (scenario == "wrap-failure" ? 3 : 1), "Owned result " + scenario);
                if (scenario == "anisotropic") Assert(result.Width == 48 && result.Height == 64 && logs.Count == 0, "Separate width/height calibration");
                if (scenario == "changed") Assert(result.Width == 64 && result.Height == 64 && made == 4 && logs.Count == 1, "Mapping change recovery");
                if (scenario == "invalid-output") Assert(result.Width == 32 && made == 5 && released == 4 && logs.Count == 3, "Invalid result bounded fallback");
                if (scenario == "wrap-failure") Assert(result.Width == 16 && released == 0 && logs.Count == 3, "Creation failure fallback");
                release(result); checks++;
            }
            Console.WriteLine("Simulated OLE dimensions/ownership/retry/fallback PASS: " + checks + " cases (100/125/150/200%)");
        }

        private static string Attr(string xml, string id, string name)
        {
            Match m = Regex.Match(xml,
                "<button\\s+id=\"" + Regex.Escape(id) + "\"[^>]*?\\s" + name + "=\"([^\"]*)\"",
                RegexOptions.Singleline);
            return m.Success ? XmlText(m.Groups[1].Value) : null;
        }

        private static string XmlText(string s)
        {
            return s.Replace("&#13;&#10;", "\r\n")
                    .Replace("&quot;", "\"")
                    .Replace("&apos;", "'")
                    .Replace("&lt;", "<")
                    .Replace("&gt;", ">")
                    .Replace("&amp;", "&");
        }
    }
}
