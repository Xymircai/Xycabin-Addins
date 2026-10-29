// =====================================================================
// 本文件由 AA-Program 配置器自动生成，请勿手改。
//
//   改标题 / 提示文案 / 图标 / 跳转网址 -> 编辑 AA-Program 配置器顶部的
//    「按钮配置区」，然后重跑：python AA-Program 下的配置器
//
//   负责向 Connect.cs 提供四样东西：
//     RibbonXml        功能区 XML（选项卡 / 分组 / 按钮 / 提示文案）
//     GetImage(id)     自定义图标 -> IPictureDisp；没有就用内置矢量图标
//     GetUrl(id)       action="url" 按钮的目标地址
//     GetResultText(cb) 结果弹窗的自定义文案模板（配置区 RESULT_TEMPLATES）
//
//   注意：XML 里每个按钮的 getImage / url 按钮的 onAction 都是【专属回调名】
//   （GetImage_<按钮id> / OnOpenUrl_<按钮id>），函数名即按钮身份 ——
//   网络盘文档的回调代理读不出 control.Id，靠名字才能认对按钮。
//   新增按钮时记得在 Connect.cs 加对应的一行包装方法（本脚本会校验）。
//
// 生成时间：2026-10-29 14:51:49
// =====================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace WordUpdater
{
    internal static class RibbonConfig
    {
        // 内置图标代号（Connect.cs 分发到功能区专用绘图与共享 OLE 尺寸转换）
        public const string BuiltinUpdate = "update";
        public const string BuiltinPaste = "paste";
        public const string BuiltinMail = "mail";
        public const string BuiltinView = "view";
        public const string BuiltinTable = "table";
        public const string BuiltinFormula = "formula";
        public const string BuiltinFilter = "filter";
        public const string BuiltinLink = "link";
        public const string BuiltinDirectory = "directory";
        public const string BuiltinAgent = "agent";
        public const string BuiltinWeb = "web";
        public const string BuiltinDocument = "document";
        public const string BuiltinSettings = "settings";
        public static string ArithmeticDefault(string key)
        {
            if (key == "add") return "100";
            if (key == "subtract") return "100";
            if (key == "multiply") return "100";
            if (key == "divide") return "10000";
            throw new ArgumentException("Unknown arithmetic default", "key");
        }

        public static string RibbonXml { get { return Xml; } }

        private const string Xml = @"<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" onLoad=""OnAgentRibbonLoad"">
  <ribbon>
    <tabs>
      <tab id=""tabWordUpdater"" label=""WordUpdater"">
        <group id=""grpLink"" label=""链接"">
          <button id=""btnUpdateLinks"" label=""更新链接"" size=""normal"" getImage=""GetImage_btnUpdateLinks"" screentip=""一键批量更新所选范围内的全部链接"" onAction=""OnUpdateLinks"" />
          <button id=""btnLinkToExcel"" label=""链接Excel"" size=""normal"" getImage=""GetImage_btnLinkToExcel"" screentip=""把复制的单元格区域逐格链接到你选中的表格区域"" supertip=""用法：先在 WPS 表格里复制一块区域（可按住 Ctrl 多选几个不连续单元格），&#13;&#10;再回到 WPS 文字，用鼠标拖选一个同样格数的矩形区域，然后点本按钮。&#13;&#10;&#13;&#10;只认你选中的这个矩形：格数必须刚好相等 —— 多了判溢出、少了判未填满，都会拒绝。&#13;&#10;源按「子区域顺序 + 区域内先行后列」展开，目标按「先行后列」展开，两边按位一一对应。&#13;&#10;想贴整张表，就先把整张表选上；只点一下光标 = 只链接 1 格。&#13;&#10;（按住 Ctrl 多选复制时，请保持 WPS 表格处于打开状态。）&#13;&#10;&#13;&#10;链接建立后，源数据变动时选中该表格「更新链接」"" onAction=""OnLinkToExcel"" />
          <button id=""btnViewLinks"" label=""查看链接"" size=""normal"" getImage=""GetImage_btnViewLinks"" screentip=""查看本文档引用了哪些Excel文件及其工作表/区域"" onAction=""OnViewLinks"" />
        </group>
        <group id=""grpAction"" label=""操作"">
          <button id=""btnSelectAllTables"" label=""全选表格"" size=""large"" getImage=""GetImage_btnSelectAllTables"" screentip=""选中本文档中的全部表格（正文不受影响）"" supertip=""自动滚动到每个表格并按住 Ctrl 逐个框选，效果等同手工 Ctrl 多选：选区里只有表格、不含正文。&#13;&#10;&#13;&#10;选中后即可用功能区的表格工具 / 字体等命令对所有表格统一调整。"" onAction=""OnSelectAllTables"" />
          <menu id=""menuFileOperations"" label=""文件操作"" size=""large"" getImage=""GetImage_menuFileOperations"" itemSize=""normal"">
            <menu id=""menuFileConversion"" label=""PDF 工具"" getImage=""GetImage_menuFileConversion"" itemSize=""normal"">
              <button id=""btnCurrentPdf"" label=""转为 PDF"" getImage=""GetImage_btnCurrentPdf"" supertip=""在窗口中切换当前文档与批量文件，支持文件夹和后缀筛选。"" onAction=""OnCurrentPdf"" />
              <button id=""btnWatermarkPdf"" label=""编辑 PDF"" getImage=""GetImage_btnWatermarkPdf"" supertip=""在窗口中切换添加水印、拆分与合并，共用来源清单及顺序。"" onAction=""OnWatermarkPdf"" />
            </menu>
          </menu>
        </group>
        <group id=""grpFeedback"" label=""反馈"">
          <button id=""btnFeedback"" label=""问题反馈"" size=""large"" getImage=""GetImage_btnFeedback"" screentip=""有BUG或新增功能请发送邮箱至2783935558@qq.com，我会尽快回复（或许）。"" supertip=""点此打开 QQ 邮箱网页版，登录后即可撰写反馈邮件。"" onAction=""OnOpenUrl_btnFeedback"" />
        </group>
        <group id=""grpExtension"" label=""扩展"">
          <toggleButton id=""btnAgentExtension"" label=""Agent扩展"" size=""large"" getImage=""GetImage_btnAgentExtension"" screentip=""显示或隐藏 Agent-Evo"" supertip=""包含 AI 网站、离线开发文档和弹窗尺寸设置；按用户记住本项目的选择。"" onAction=""OnToggleAgentExtension"" getPressed=""GetAgentExtensionPressed"" />
        </group>
      </tab>
      <tab id=""tabAgentEvo"" label=""Agent-Evo"" getVisible=""GetAgentTabVisible"">
        <group id=""grpAgentResources"" label=""详情"">
          <menu id=""menuAgentApi"" label=""调用API"" size=""large"" getImage=""GetImage_menuAgentApi"" itemSize=""normal"">
            <menu id=""menuAgentCloud"" label=""智能云-API"" getImage=""GetImage_menuAgentCloud"" itemSize=""normal"">
              <button id=""btnAgentBaidu"" label=""百度智能云"" getImage=""GetImage_btnAgentBaidu"" onAction=""OnOpenUrl_btnAgentBaidu"" />
              <button id=""btnAgentAliyun"" label=""阿里智能云"" getImage=""GetImage_btnAgentAliyun"" onAction=""OnOpenUrl_btnAgentAliyun"" />
            </menu>
            <menu id=""menuAgentSites"" label=""Agent-API"" getImage=""GetImage_menuAgentSites"" itemSize=""normal"">
              <menu id=""menuAgentChina"" label=""国内Agent"" getImage=""GetImage_menuAgentChina"" itemSize=""normal"">
                <button id=""btnAgentWorkBuddy"" label=""WorkBuddy"" getImage=""GetImage_btnAgentWorkBuddy"" onAction=""OnOpenUrl_btnAgentWorkBuddy"" />
                <button id=""btnAgentTrae"" label=""TRAE"" getImage=""GetImage_btnAgentTrae"" onAction=""OnOpenUrl_btnAgentTrae"" />
                <button id=""btnAgentQoder"" label=""Qoder"" getImage=""GetImage_btnAgentQoder"" onAction=""OnOpenUrl_btnAgentQoder"" />
                <button id=""btnAgentDeepSeek"" label=""DeepSeek"" getImage=""GetImage_btnAgentDeepSeek"" onAction=""OnOpenUrl_btnAgentDeepSeek"" />
                <button id=""btnAgentDoubao"" label=""豆包"" getImage=""GetImage_btnAgentDoubao"" onAction=""OnOpenUrl_btnAgentDoubao"" />
                <button id=""btnAgentKimi"" label=""Kimi"" getImage=""GetImage_btnAgentKimi"" onAction=""OnOpenUrl_btnAgentKimi"" />
              </menu>
              <menu id=""menuAgentOverseas"" label=""国外Agent"" getImage=""GetImage_menuAgentOverseas"" itemSize=""normal"">
                <button id=""btnAgentChatGPT"" label=""ChatGPT"" getImage=""GetImage_btnAgentChatGPT"" onAction=""OnOpenUrl_btnAgentChatGPT"" />
                <button id=""btnAgentClaude"" label=""Claude"" getImage=""GetImage_btnAgentClaude"" onAction=""OnOpenUrl_btnAgentClaude"" />
                <button id=""btnAgentGemini"" label=""Gemini"" getImage=""GetImage_btnAgentGemini"" onAction=""OnOpenUrl_btnAgentGemini"" />
              </menu>
            </menu>
          </menu>
          <menu id=""menuAgentWorkflow"" label=""开发流程"" size=""large"" getImage=""GetImage_menuAgentWorkflow"" itemSize=""normal"">
            <button id=""btnAgentReadme"" label=""开发预览"" getImage=""GetImage_btnAgentReadme"" onAction=""OnOpenAgentReadme"" />
            <button id=""btnAgentUserGuide"" label=""用户预览"" getImage=""GetImage_btnAgentUserGuide"" onAction=""OnOpenAgentUserGuide"" />
            <button id=""btnAgentHome"" label=""XyCabin"" getImage=""GetImage_btnAgentHome"" onAction=""OnAgentHome"" />
          </menu>
        </group>
        <group id=""grpAgentOperations"" label=""操作"">
          <menu id=""menuAgentCustomization"" label=""自定配置"" size=""large"" getImage=""GetImage_menuAgentCustomization"" itemSize=""normal"">
            <button id=""btnAgentSort"" label=""自定排序"" getImage=""GetImage_btnAgentSort"" onAction=""OnAgentSort"" />
            <button id=""btnAgentNames"" label=""自定名称"" getImage=""GetImage_btnAgentNames"" onAction=""OnAgentNames"" />
            <button id=""btnAgentFormats"" label=""自定格式"" getImage=""GetImage_btnAgentFormats"" onAction=""OnAgentFormats"" />
            <button id=""btnAgentDialogSizes"" label=""交互弹窗"" getImage=""GetImage_btnAgentDialogSizes"" onAction=""OnOpenAgentDialogSizes"" />
          </menu>
          <menu id=""menuAgentExperiments"" label=""实验计划"" size=""large"" getImage=""GetImage_menuAgentExperiments"" itemSize=""normal"">
            <button id=""btnAgentOcrPlaceholder"" label=""OCR识别"" getImage=""GetImage_btnAgentOcrPlaceholder"" enabled=""false"" onAction=""OnAgentOcrPlaceholder"" />
            <button id=""btnAgentRequirements"" label=""需求记录"" getImage=""GetImage_btnAgentRequirements"" onAction=""OnAgentRequirements"" />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        // 透明区域合成底色（HBITMAP 不支持透明通道，只能先垫一层底）
        private static readonly Color PlateColor = ColorTranslator.FromHtml("#F5F6F7");

        private static readonly Dictionary<string, int> Sizes = MakeSizes();
        private static readonly Dictionary<string, string> Icons = MakeIcons();
        private static readonly Dictionary<string, string> BuiltinIcons = MakeBuiltinIcons();
        private static readonly Dictionary<string, string> Urls = MakeUrls();
        private static readonly Dictionary<string, string> ResultTexts = MakeResultTexts();
        private static readonly Dictionary<string, IPictureDisp> Cache =
            new Dictionary<string, IPictureDisp>();

        /// <summary>自定义图标；该按钮没配图则返回 null（调用方会改用内置图标）。</summary>
        public static IPictureDisp GetImage(string id)
        {
            try
            {
                if (string.IsNullOrEmpty(id)) return null;
                string b64;
                if (!Icons.TryGetValue(id, out b64)) return null;
                lock (Cache)
                {
                    IPictureDisp hit;
                    if (Cache.TryGetValue(id, out hit) && hit != null) return hit;
                }
                int px = GetImageSize(id);
                IPictureDisp pic = ToPicture(Convert.FromBase64String(b64), px);
                lock (Cache) Cache[id] = pic;
                return pic;
            }
            catch (Exception ex)
            {
                // 静默返回 null 的时代 WPS 只"白给你看"，什么线索都没有；
                // 现在记一笔日志，并把缓存清掉让下次回调重建
                try { Connect.Log("GetImage(" + (id ?? "") + ") 异常：" + ex.Message); }
                catch { }
                lock (Cache) Cache.Remove(id);
                return null;
            }
        }

        // Shared by custom and built-in image callbacks; public callbacks stay unchanged.
        internal static int GetImageSize(string id)
        {
            int size;
            return id != null && Sizes.TryGetValue(id, out size) ? size : 32;
        }

        /// <summary>该按钮回落用的内置图标代号；没有则返回 null。</summary>
        public static string BuiltinIconKind(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            string kind;
            return BuiltinIcons.TryGetValue(id, out kind) ? kind : null;
        }

        /// <summary>action="url" 按钮的目标地址；没有则返回 null。</summary>
        public static string GetUrl(string id)
        {
            if (string.IsNullOrEmpty(id)) return null;
            string url;
            return Urls.TryGetValue(id, out url) ? url : null;
        }

        /// <summary>
        /// 结果弹窗的自定义文案模板（键=回调名，如 "OnUpdateLinks"）。
        /// 没有配置则返回 null，调用方（Connect.cs 的 FormatResult / FormatLinkResult）
        /// 会改用内置默认文案。
        /// </summary>
        public static string GetResultText(string callback)
        {
            if (string.IsNullOrEmpty(callback)) return null;
            string tpl;
            return ResultTexts.TryGetValue(callback, out tpl) ? tpl : null;
        }

        // ------------------------------ 数据表 ------------------------------

        private static Dictionary<string, string> MakeIcons()
        {
            var d = new Dictionary<string, string>();
            // 未配置任何自定义图标（全部按钮使用内置矢量图标）
            return d;
        }

        private static Dictionary<string, int> MakeSizes()
        {
            var d = new Dictionary<string, int>();
            d["btnUpdateLinks"] = 16;
            d["btnLinkToExcel"] = 16;
            d["btnViewLinks"] = 16;
            d["btnSelectAllTables"] = 32;
            d["menuFileOperations"] = 32;
            d["menuFileConversion"] = 16;
            d["btnCurrentPdf"] = 16;
            d["btnWatermarkPdf"] = 16;
            d["btnFeedback"] = 32;
            d["btnAgentExtension"] = 32;
            d["menuAgentApi"] = 32;
            d["menuAgentCloud"] = 16;
            d["btnAgentBaidu"] = 16;
            d["btnAgentAliyun"] = 16;
            d["menuAgentSites"] = 16;
            d["menuAgentChina"] = 16;
            d["btnAgentWorkBuddy"] = 16;
            d["btnAgentTrae"] = 16;
            d["btnAgentQoder"] = 16;
            d["btnAgentDeepSeek"] = 16;
            d["btnAgentDoubao"] = 16;
            d["btnAgentKimi"] = 16;
            d["menuAgentOverseas"] = 16;
            d["btnAgentChatGPT"] = 16;
            d["btnAgentClaude"] = 16;
            d["btnAgentGemini"] = 16;
            d["menuAgentWorkflow"] = 32;
            d["btnAgentReadme"] = 16;
            d["btnAgentUserGuide"] = 16;
            d["btnAgentHome"] = 16;
            d["menuAgentCustomization"] = 32;
            d["btnAgentSort"] = 16;
            d["btnAgentNames"] = 16;
            d["btnAgentFormats"] = 16;
            d["btnAgentDialogSizes"] = 16;
            d["menuAgentExperiments"] = 32;
            d["btnAgentOcrPlaceholder"] = 16;
            d["btnAgentRequirements"] = 16;
            return d;
        }

        private static Dictionary<string, string> MakeBuiltinIcons()
        {
            var d = new Dictionary<string, string>();
            d["btnUpdateLinks"] = BuiltinUpdate;
            d["btnLinkToExcel"] = BuiltinPaste;
            d["btnViewLinks"] = BuiltinView;
            d["btnSelectAllTables"] = BuiltinTable;
            d["menuFileOperations"] = BuiltinDocument;
            d["menuFileConversion"] = BuiltinDocument;
            d["btnCurrentPdf"] = BuiltinDocument;
            d["btnWatermarkPdf"] = BuiltinDocument;
            d["btnFeedback"] = BuiltinMail;
            d["btnAgentExtension"] = BuiltinAgent;
            d["menuAgentApi"] = BuiltinWeb;
            d["menuAgentCloud"] = BuiltinWeb;
            d["btnAgentBaidu"] = BuiltinWeb;
            d["btnAgentAliyun"] = BuiltinWeb;
            d["menuAgentSites"] = BuiltinAgent;
            d["menuAgentChina"] = BuiltinAgent;
            d["btnAgentWorkBuddy"] = BuiltinWeb;
            d["btnAgentTrae"] = BuiltinWeb;
            d["btnAgentQoder"] = BuiltinWeb;
            d["btnAgentDeepSeek"] = BuiltinWeb;
            d["btnAgentDoubao"] = BuiltinWeb;
            d["btnAgentKimi"] = BuiltinWeb;
            d["menuAgentOverseas"] = BuiltinAgent;
            d["btnAgentChatGPT"] = BuiltinWeb;
            d["btnAgentClaude"] = BuiltinWeb;
            d["btnAgentGemini"] = BuiltinWeb;
            d["menuAgentWorkflow"] = BuiltinDocument;
            d["btnAgentReadme"] = BuiltinDocument;
            d["btnAgentUserGuide"] = BuiltinDocument;
            d["btnAgentHome"] = BuiltinAgent;
            d["menuAgentCustomization"] = BuiltinSettings;
            d["btnAgentSort"] = BuiltinSettings;
            d["btnAgentNames"] = BuiltinSettings;
            d["btnAgentFormats"] = BuiltinSettings;
            d["btnAgentDialogSizes"] = BuiltinSettings;
            d["menuAgentExperiments"] = BuiltinSettings;
            d["btnAgentOcrPlaceholder"] = BuiltinDocument;
            d["btnAgentRequirements"] = BuiltinDocument;
            return d;
        }

        private static Dictionary<string, string> MakeUrls()
        {
            var d = new Dictionary<string, string>();
            d["btnFeedback"] = @"https://wx.mail.qq.com";
            d["btnAgentBaidu"] = @"https://console.bce.baidu.com";
            d["btnAgentAliyun"] = @"https://home.console.aliyun.com/";
            d["btnAgentWorkBuddy"] = @"https://copilot.tencent.com/work/";
            d["btnAgentTrae"] = @"https://www.trae.cn/";
            d["btnAgentQoder"] = @"https://qoder.com/";
            d["btnAgentDeepSeek"] = @"https://chat.deepseek.com/";
            d["btnAgentDoubao"] = @"https://www.doubao.com/";
            d["btnAgentKimi"] = @"https://www.kimi.com/";
            d["btnAgentChatGPT"] = @"https://chatgpt.com/";
            d["btnAgentClaude"] = @"https://claude.ai/";
            d["btnAgentGemini"] = @"https://gemini.google.com/";
            return d;
        }

        private static Dictionary<string, string> MakeResultTexts()
        {
            var d = new Dictionary<string, string>();
            d["OnLinkToExcel"] = "链接粘贴完成。\r\n已按 {rows} 行 × {cols} 列建立 {cells} 个链接，其中 {refreshed} 个已刷新。\r\n源：{source}";
            d["OnUpdateLinks"] = "链接更新完成。\r\n共发现 {found} 个链接：成功 {success} 个，失败 {failed} 个，数值变动 {changed} 个。\r\n";
            return d;
        }

        // ------------------------- base64 -> IPictureDisp -------------------------

        private static IPictureDisp ToPicture(byte[] raw, int px)
        {
            using (MemoryStream ms = new MemoryStream(raw))
            using (Bitmap src = new Bitmap(ms))
            {
                // Preserve configured plate color and custom-image priority. OLE cannot
                // reliably show alpha in WPS, so compose an opaque 24bpp physical bitmap.
                return RibbonImages.Create(px, delegate(int width, int height)
                {
                    Bitmap bitmap = new Bitmap(width, height, PixelFormat.Format24bppRgb);
                    try
                    {
                        using (Graphics g = Graphics.FromImage(bitmap))
                        {
                            g.Clear(PlateColor);
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.DrawImage(src, 0, 0, width, height);
                        }
                        return bitmap;
                    }
                    catch { bitmap.Dispose(); throw; }
                }, Connect.Log);
            }
        }
    }
}
