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
// 生成时间：2026-10-29 14:50:58
// =====================================================================
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace Excel_Iterate
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
      <tab id=""tabEtTrace"" label=""Excel-Iterate"">
        <group id=""grpTrace"" label=""数据"">
          <button id=""btnTracePrecedents"" label=""追踪引用"" size=""normal"" getImage=""GetImage_btnTracePrecedents"" screentip=""解析当前单元格公式，逐项展开被引用单元格的详情"" supertip=""选中含公式的单元格后点击，列出该公式直接引用的全部单元格：&#13;&#10;工作簿、工作表、地址、引用类型（同表/跨表/跨簿）、绝对/相对。&#13;&#10;&#13;&#10;双击列表某条记录可跳转并高亮闪烁；跨簿未打开时会提示后打开再跳。"" onAction=""OnTracePrecedents"" />
          <button id=""btnTraceDependents"" label=""追踪从属"" size=""normal"" getImage=""GetImage_btnTraceDependents"" screentip=""遍历所有已打开工作簿，找出哪些单元格的公式引用了当前单元格"" supertip=""选中单元格后点击，遍历所有已打开工作簿的所有工作表公式，&#13;&#10;反向找出引用了当前单元格的单元格并列出详情。&#13;&#10;&#13;&#10;大表扫描时显示进度条并可取消；双击列表记录跳转高亮。"" onAction=""OnTraceDependents"" />
          <menu id=""menuFilterEnhanced"" label=""筛选增强"" size=""normal"" getImage=""GetImage_menuFilterEnhanced"" screentip=""按活动单元格所在表格的表头选择显示列"" itemSize=""normal"">
            <button id=""btnFilterColumns"" label=""按列筛选"" getImage=""GetImage_btnFilterColumns"" onAction=""OnFilterColumns"" />
            <button id=""btnQuickAdvancedFilter"" label=""快捷高筛"" getImage=""GetImage_btnQuickAdvancedFilter"" onAction=""OnQuickAdvancedFilter"" />
          </menu>
          <menu id=""menuAuditToolbox"" label=""审计工具箱"" size=""large"" getImage=""GetImage_menuAuditToolbox"" screentip=""审计辅助工具集合：表头匹配合并、合并抵消"" supertip=""表头匹配：按表头列名（忽略大小写与首尾空格）匹配列，&#13;&#10;把来源工作簿的列数据合并到目标工作簿同名列下。&#13;&#10;&#13;&#10;合并抵消：往来抵消按单位对冲生成分录；收入成本抵消按（核算主体,客商,年度）三维聚合生成矩阵式底稿。&#13;&#10;两者列映射角色均可改，写入前备份、可一键还原。"" itemSize=""normal"">
            <button id=""btnMergeHeaders"" label=""表头匹配"" getImage=""GetImage_btnMergeHeaders"" screentip=""按表头列名匹配，把多个来源工作簿的列追加到目标工作簿同名列下"" supertip=""选择目标工作簿与若干来源工作簿，按表头列名匹配列，&#13;&#10;把来源列的数据追加到目标工作簿同名列已有数据的下方。&#13;&#10;&#13;&#10;每一列各自追加到该列已有数据的末尾；目标缺失的列名会先列出并请你确认，&#13;&#10;不会自动新建列。公式按目标位置重写同表引用；跨表与跨簿前缀及列不变，相对行按行差偏移。&#13;&#10;写入前自动备份目标工作簿，完成后可点「取消操作」还原。"" onAction=""OnMergeHeaders"" />
            <menu id=""menuElimination"" label=""合并抵消"" getImage=""GetImage_menuElimination"" screentip=""合并报表抵消分录辅助（往来抵消 / 收入成本抵消）"" supertip=""往来抵消按单位对冲出凭证式分录；收入成本抵消按三维聚合出矩阵式底稿（年度列动态、金额写活公式）。&#13;&#10;列映射角色可改，写入前备份、可一键还原。"" itemSize=""normal"">
              <button id=""btnEliminationReceivable"" label=""往来抵消"" getImage=""GetImage_btnEliminationReceivable"" screentip=""内部往来科目抵消"" supertip=""按（核算主体,对方单位,年度）双向合计后以小抵大，生成「往来抵消差异」与「往来抵消结果」两张表；年度取截止日前 4 位；三个科目对：应收应付、其他应收其他应付、预付账款合同负债。"" onAction=""OnEliminationReceivable"" />
              <button id=""btnEliminationRevenueCost"" label=""收入成本抵消"" getImage=""GetImage_btnEliminationRevenueCost"" screentip=""内部交易收入成本抵消"" supertip=""按（核算主体,客商,年度）三维聚合生成矩阵式抵消底稿：借 营业收入、贷 营业成本、贷 成本还原、贷 存货；年度列按数据动态生成，金额写活公式，写入前备份、可一键还原。"" onAction=""OnEliminationRevenueCost"" />
            </menu>
          </menu>
        </group>
        <group id=""grpAction"" label=""操作"">
          <menu id=""menuOneClickFormula"" label=""一键公式"" size=""large"" getImage=""GetImage_menuOneClickFormula"" screentip=""封装或移除当前选区公式的最外层函数"" supertip=""仅点击最终选项才修改公式；支持连续选区和 Ctrl 多选。"" itemSize=""normal"">
            <menu id=""menuWrapFormula"" label=""进行公式封装"" getImage=""GetImage_menuWrapFormula"" itemSize=""normal"">
              <menu id=""menuIFErrorWrap"" label=""IFERROR封装"" getImage=""GetImage_menuIFErrorWrap"" itemSize=""normal"">
                <button id=""btnWrapIFError"" label=""默认返回空"" getImage=""GetImage_btnWrapIFError"" supertip=""增加 IFERROR(原公式,&quot;&quot;) 一层。"" onAction=""OnWrapIFError"" />
                <button id=""btnCustomIFError"" label=""自定义封装"" getImage=""GetImage_btnCustomIFError"" supertip=""设置错误时返回的文字、数值或公式/引用，确认后处理选区；取消不修改。"" onAction=""OnCustomIFError"" />
              </menu>
              <menu id=""menuRoundWrap"" label=""ROUND封装"" getImage=""GetImage_menuRoundWrap"" itemSize=""normal"">
                <button id=""btnWrapRound"" label=""默认保留2位"" getImage=""GetImage_btnWrapRound"" supertip=""增加 ROUND(原公式,2) 一层，固定保留两位小数。"" onAction=""OnWrapRound"" />
                <button id=""btnCustomRound"" label=""自定义封装"" getImage=""GetImage_btnCustomRound"" supertip=""自定义整数位数，例如 10、0 或 -2，确认后处理选区；取消不修改。"" onAction=""OnCustomRound"" />
              </menu>
              <button id=""btnWrapAbs"" label=""ABS封装"" getImage=""GetImage_btnWrapAbs"" supertip=""增加 ABS(原公式) 一层，使用标准单参数语法。"" onAction=""OnWrapAbs"" />
              <button id=""btnCustomFunction"" label=""自定义函数"" getImage=""GetImage_btnCustomFunction"" supertip=""输入任意函数模板，留一个空参数位表示每个单元格内容，如 =TEXTSPLIT(,&quot;-&quot;,1) 或 =ISBLANK()。"" onAction=""OnCustomFunction"" />
            </menu>
            <button id=""btnUnwrapFormula"" label=""移除公式封装"" getImage=""GetImage_btnUnwrapFormula"" supertip=""只拆一层完整包住整条公式的函数，保留第一个参数，丢弃其他参数。"" onAction=""OnUnwrapFormula"" />
            <menu id=""menuArithmetic"" label=""数值计算"" getImage=""GetImage_menuArithmetic"" itemSize=""normal"">
              <menu id=""menuCalculateAdd"" label=""加"" getImage=""GetImage_menuCalculateAdd"" itemSize=""normal"">
                <button id=""btnCalculateAdd"" label=""默认加100"" getImage=""GetImage_btnCalculateAdd"" onAction=""OnCalculateAdd"" />
                <button id=""btnCustomCalculateAdd"" label=""自定义计算"" getImage=""GetImage_btnCustomCalculateAdd"" onAction=""OnCustomCalculateAdd"" />
              </menu>
              <menu id=""menuCalculateSubtract"" label=""减"" getImage=""GetImage_menuCalculateSubtract"" itemSize=""normal"">
                <button id=""btnCalculateSubtract"" label=""默认减100"" getImage=""GetImage_btnCalculateSubtract"" onAction=""OnCalculateSubtract"" />
                <button id=""btnCustomCalculateSubtract"" label=""自定义计算"" getImage=""GetImage_btnCustomCalculateSubtract"" onAction=""OnCustomCalculateSubtract"" />
              </menu>
              <menu id=""menuCalculateMultiply"" label=""乘"" getImage=""GetImage_menuCalculateMultiply"" itemSize=""normal"">
                <button id=""btnCalculateMultiply"" label=""默认乘100"" getImage=""GetImage_btnCalculateMultiply"" onAction=""OnCalculateMultiply"" />
                <button id=""btnCustomCalculateMultiply"" label=""自定义计算"" getImage=""GetImage_btnCustomCalculateMultiply"" onAction=""OnCustomCalculateMultiply"" />
              </menu>
              <menu id=""menuCalculateDivide"" label=""除"" getImage=""GetImage_menuCalculateDivide"" itemSize=""normal"">
                <button id=""btnCalculateDivide"" label=""默认除10000"" getImage=""GetImage_btnCalculateDivide"" onAction=""OnCalculateDivide"" />
                <button id=""btnCustomCalculateDivide"" label=""自定义计算"" getImage=""GetImage_btnCustomCalculateDivide"" onAction=""OnCustomCalculateDivide"" />
              </menu>
            </menu>
          </menu>
          <menu id=""menuFileOperations"" label=""文件操作"" size=""large"" getImage=""GetImage_menuFileOperations"" screentip=""整理文件、批量重命名及 PDF 转换、水印、拆分与合并"" supertip=""批量处理保留原功能；PDF 支持当前文档、混合批量转换、水印、拆分和合并。"" itemSize=""normal"">
            <menu id=""menuBatchProcessing"" label=""批量处理"" getImage=""GetImage_menuBatchProcessing"" itemSize=""normal"">
              <button id=""btnOrganizeFiles"" label=""文件批量整理"" getImage=""GetImage_btnOrganizeFiles"" supertip=""把子目录文件移到第一层；可分类、处理同名或按类型送入回收站。"" onAction=""OnOrganizeFiles"" />
              <button id=""btnRenameFiles"" label=""文件批量命名"" getImage=""GetImage_btnRenameFiles"" supertip=""扫描文件夹第一层文件，填写清单第三列后执行替换。"" onAction=""OnRenameFiles"" />
              <button id=""btnFileNames"" label=""文件提取名称"" getImage=""GetImage_btnFileNames"" supertip=""扫描并预览文件清单；OCR可设置多个前缀，逐项核对后按顺序用短横线组合结果。"" onAction=""OnFileNames"" />
            </menu>
            <menu id=""menuFileConversion"" label=""PDF 工具"" getImage=""GetImage_menuFileConversion"" itemSize=""normal"">
              <button id=""btnCurrentPdf"" label=""转为 PDF"" getImage=""GetImage_btnCurrentPdf"" supertip=""在窗口中切换当前文档与批量文件，支持文件夹和后缀筛选。"" onAction=""OnCurrentPdf"" />
              <button id=""btnWatermarkPdf"" label=""编辑 PDF"" getImage=""GetImage_btnWatermarkPdf"" supertip=""在窗口中切换添加水印、拆分与合并，共用来源清单及顺序。"" onAction=""OnWatermarkPdf"" />
            </menu>
          </menu>
          <button id=""btnLinkAssistant"" label=""链接助手"" size=""large"" getImage=""GetImage_btnLinkAssistant"" screentip=""扫描 xlsx 外部引用，C列填新路径或「断开链接」后执行；支持清理所选来源的未使用名称，断链可从整簿备份恢复"" onAction=""OnLinkAssistant"" />
          <button id=""btnGenerateDirectory"" label=""生成目录"" size=""large"" getImage=""GetImage_btnGenerateDirectory"" screentip=""链接全部工作表，隐藏表点击后显示并跳转"" onAction=""OnGenerateDirectory"" />
        </group>
        <group id=""grpFeedback"" label=""反馈"">
          <button id=""btnFeedback"" label=""问题反馈"" size=""large"" getImage=""GetImage_btnFeedback"" screentip=""有BUG或新增功能请发送邮箱至2783935558@qq.com，我会尽快回复（尽量）。"" supertip=""点此打开 QQ 邮箱网页版，登录后即可撰写反馈邮件。"" onAction=""OnOpenUrl_btnFeedback"" />
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
            d["btnTracePrecedents"] = 16;
            d["btnTraceDependents"] = 16;
            d["menuFilterEnhanced"] = 16;
            d["btnFilterColumns"] = 16;
            d["btnQuickAdvancedFilter"] = 16;
            d["menuAuditToolbox"] = 32;
            d["btnMergeHeaders"] = 16;
            d["menuElimination"] = 16;
            d["btnEliminationReceivable"] = 16;
            d["btnEliminationRevenueCost"] = 16;
            d["menuOneClickFormula"] = 32;
            d["menuWrapFormula"] = 16;
            d["menuIFErrorWrap"] = 16;
            d["btnWrapIFError"] = 16;
            d["btnCustomIFError"] = 16;
            d["menuRoundWrap"] = 16;
            d["btnWrapRound"] = 16;
            d["btnCustomRound"] = 16;
            d["btnWrapAbs"] = 16;
            d["btnCustomFunction"] = 16;
            d["btnUnwrapFormula"] = 16;
            d["menuArithmetic"] = 16;
            d["menuCalculateAdd"] = 16;
            d["btnCalculateAdd"] = 16;
            d["btnCustomCalculateAdd"] = 16;
            d["menuCalculateSubtract"] = 16;
            d["btnCalculateSubtract"] = 16;
            d["btnCustomCalculateSubtract"] = 16;
            d["menuCalculateMultiply"] = 16;
            d["btnCalculateMultiply"] = 16;
            d["btnCustomCalculateMultiply"] = 16;
            d["menuCalculateDivide"] = 16;
            d["btnCalculateDivide"] = 16;
            d["btnCustomCalculateDivide"] = 16;
            d["menuFileOperations"] = 32;
            d["menuBatchProcessing"] = 16;
            d["btnOrganizeFiles"] = 16;
            d["btnRenameFiles"] = 16;
            d["btnFileNames"] = 16;
            d["menuFileConversion"] = 16;
            d["btnCurrentPdf"] = 16;
            d["btnWatermarkPdf"] = 16;
            d["btnLinkAssistant"] = 32;
            d["btnGenerateDirectory"] = 32;
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
            d["btnTracePrecedents"] = BuiltinView;
            d["btnTraceDependents"] = BuiltinView;
            d["menuFilterEnhanced"] = BuiltinFilter;
            d["btnFilterColumns"] = BuiltinFilter;
            d["btnQuickAdvancedFilter"] = BuiltinFilter;
            d["menuAuditToolbox"] = BuiltinTable;
            d["btnMergeHeaders"] = BuiltinTable;
            d["menuElimination"] = BuiltinTable;
            d["btnEliminationReceivable"] = BuiltinTable;
            d["btnEliminationRevenueCost"] = BuiltinTable;
            d["menuOneClickFormula"] = BuiltinFormula;
            d["menuWrapFormula"] = BuiltinFormula;
            d["menuIFErrorWrap"] = BuiltinFormula;
            d["btnWrapIFError"] = BuiltinFormula;
            d["btnCustomIFError"] = BuiltinFormula;
            d["menuRoundWrap"] = BuiltinFormula;
            d["btnWrapRound"] = BuiltinFormula;
            d["btnCustomRound"] = BuiltinFormula;
            d["btnWrapAbs"] = BuiltinFormula;
            d["btnCustomFunction"] = BuiltinFormula;
            d["btnUnwrapFormula"] = BuiltinFormula;
            d["menuArithmetic"] = BuiltinFormula;
            d["menuCalculateAdd"] = BuiltinFormula;
            d["btnCalculateAdd"] = BuiltinFormula;
            d["btnCustomCalculateAdd"] = BuiltinFormula;
            d["menuCalculateSubtract"] = BuiltinFormula;
            d["btnCalculateSubtract"] = BuiltinFormula;
            d["btnCustomCalculateSubtract"] = BuiltinFormula;
            d["menuCalculateMultiply"] = BuiltinFormula;
            d["btnCalculateMultiply"] = BuiltinFormula;
            d["btnCustomCalculateMultiply"] = BuiltinFormula;
            d["menuCalculateDivide"] = BuiltinFormula;
            d["btnCalculateDivide"] = BuiltinFormula;
            d["btnCustomCalculateDivide"] = BuiltinFormula;
            d["menuFileOperations"] = BuiltinDocument;
            d["menuBatchProcessing"] = BuiltinDocument;
            d["btnOrganizeFiles"] = BuiltinDocument;
            d["btnRenameFiles"] = BuiltinDocument;
            d["btnFileNames"] = BuiltinDocument;
            d["menuFileConversion"] = BuiltinDocument;
            d["btnCurrentPdf"] = BuiltinDocument;
            d["btnWatermarkPdf"] = BuiltinDocument;
            d["btnLinkAssistant"] = BuiltinLink;
            d["btnGenerateDirectory"] = BuiltinDirectory;
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
            // 未自定义结果弹窗文案（用 Connect.cs 里的内置默认文案）
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
