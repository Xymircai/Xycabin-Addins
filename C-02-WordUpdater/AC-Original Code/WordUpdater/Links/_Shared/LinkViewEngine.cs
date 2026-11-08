using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace WordUpdater
{
    // =====================================================================
    // 「查看链接」引擎 —— 纯逻辑，不碰 UI（便于 Harvey 式复用与离线自检）
    //
    //   与 UpdateEngine 的最大区别：UpdateEngine 只处理【选区】里的内容，
    //   本引擎扫描【整个活动文档】，因为"看看这篇文档到底引用了哪些文件"
    //   这种问题不存在"选中范围"的前提。
    //
    //   扫描对象（按各自的 story range 逐层展开，覆盖正文 / 页眉页脚 / 脚注等）：
    //     链接字段   LINK / DDE / DDEAUTO / INCLUDEPICTURE / INCLUDETEXT
    //     链接对象   InlineShapes / Shapes 上带 LinkFormat 的（Excel 图表、图片）
    //
    //   解析出来的一条条记录按【源文件】聚合成 LinkSourceInfo 交给窗口显示，
    //   并基于同一份收集结果提供两个动作：
    //     UpdateAll()  更新本文档全部链接
    //     BreakAll()   中断本文档全部链接（转为静态内容）
    // =====================================================================

    public enum LinkViewStatus
    {
        NoDocument,   // 没有打开的文档
        NoLinks,      // 文档里没有任何链接
        Completed,    // 扫描完成
        Error         // 引擎级错误
    }

    /// <summary>文档中单个链接的可读信息（一个单元格/一块区域/一个对象 = 一条）。</summary>
    public class LinkItemInfo
    {
        public int Index;            // 文档内序号（1 起始，扫描顺序）
        public string ItemKind = ""; // 链接字段 / 嵌入对象 / 浮动形状
        public string KindCode = ""; // LINK / DDE / DDEAUTO / INCLUDEPICTURE / OLE
        public string SourcePath = "";
        public string SheetName = "";
        public string RangeText = "";
        public string UpdateMode = ""; // 自动 / 手动
        public string RawCode = "";    // 原始字段代码（明细展示用）
    }

    /// <summary>按源文件聚合后的一个条目（窗口上表的一行）。</summary>
    public class LinkSourceInfo
    {
        public string FullPath = "";
        public string WorkbookName = "";
        public string FolderPath = "";
        public List<string> Sheets = new List<string>();
        public int LinkCount;
        public int FieldCount;
        public int ObjectCount;
        public int AutoCount;
        public List<LinkItemInfo> Items = new List<LinkItemInfo>();

        public bool PathIsLocal;   // 看起来像本地/网络文件路径
        public bool FileExists;

        public string SheetText = "";   // 工作表汇总（多个用 、 连接）
        public string KindText = "";    // 链接字段 / 链接对象 / 字段 + 对象
        public string UpdateText = "";  // 自动 / 手动 / 混合
        public string StatusText = "";  // 正常 / 源文件缺失 / 非本地路径
    }

    public class LinkViewResult
    {
        public LinkViewStatus Status;
        public string ErrorMessage;
        public string DocumentName = "";
        public int TotalLinks;
        public List<LinkSourceInfo> Sources = new List<LinkSourceInfo>();
    }

    public class LinkActionResult
    {
        public bool Executed;      // 动作是否跑完（false = 引擎级异常，看 ErrorMessage）
        public int Total;
        public int Success;
        public int Failed;
        public List<string> Failures = new List<string>();
        public int CleanedHiddenHosts;
        public string ErrorMessage;
    }

    public static class LinkViewEngine
    {
        private const int WdFieldLink = 56;      // wdFieldLink
        private const int WdAlertsNone = 0;      // wdAlertsNone

        private static readonly Regex QuoteRegex =
            new Regex("\"([^\"]+)\"", RegexOptions.Compiled);

        // ==================================================================
        //                              扫描
        // ==================================================================

        public static LinkViewResult Scan(object application)
        {
            LinkViewResult r = new LinkViewResult();
            if (application == null)
            {
                r.Status = LinkViewStatus.Error;
                r.ErrorMessage = "宿主应用对象为空。";
                return r;
            }

            try
            {
                dynamic app = application;
                dynamic doc = null;
                try { doc = app.ActiveDocument; }
                catch { }
                if (doc == null)
                {
                    r.Status = LinkViewStatus.NoDocument;
                    return r;
                }
                try { r.DocumentName = (string)doc.Name; }
                catch { }

                List<RawLink> links = CollectAllLinks(doc);
                r.TotalLinks = links.Count;
                if (links.Count == 0)
                {
                    r.Status = LinkViewStatus.NoLinks;
                    return r;
                }

                List<LinkItemInfo> infos = new List<LinkItemInfo>();
                for (int i = 0; i < links.Count; i++)
                {
                    links[i].Info.Index = i + 1;
                    infos.Add(links[i].Info);
                }
                r.Sources = Aggregate(infos);
                r.Status = LinkViewStatus.Completed;
                return r;
            }
            catch (Exception ex)
            {
                r.Status = LinkViewStatus.Error;
                r.ErrorMessage = ex.Message;
                return r;
            }
        }

        // ==================================================================
        //                      动作：跳转到某条链接的位置
        // ==================================================================

        /// <summary>
        /// 把光标选到某条链接在文档里的位置（字段代码区 / 对象锚点）。
        /// 返回 null = 成功；返回字符串 = 失败原因（交 UI 弹给用户）。
        ///
        /// 实现：不跨调用保留 COM 对象（窗口非模态，文档随时可能被编辑），
        /// 而是按 LinkItemInfo.Index 重新收集一遍取同序号的链接；
        /// 取之前核对类型代号 / 字段代码 / 源路径，对不上说明列表已过期，
        /// 让 UI 提示用户刷新而不是跳错位置。
        /// </summary>
        public static LinkActionResult UpdateAll(object application) { return RunAll(application, false); }
        public static LinkActionResult BreakAll(object application) { return RunAll(application, true); }
        private static LinkActionResult RunAll(object application, bool unlink)
        {
            var result = new LinkActionResult();
            try
            {
                dynamic app = application;
                List<RawLink> links = CollectAllLinks(app.ActiveDocument);
                result.Total = links.Count;
                if (links.Count == 0) { result.Executed = true; return result; }
                var state = WordHostStateScope.Enter(application);
                HashSet<int> before = unlink ? null : EtCleanup.SnapshotEtPids();
                try
                {
                    for (int step = 0; step < links.Count; step++)
                    {
                        int index = unlink ? links.Count - 1 - step : step;
                        var link = links[index];
                        try
                        {
                            dynamic item = link.Com; bool success = true;
                            if (unlink)
                            {
                                if (link.IsField) item.Unlink(); else item.LinkFormat.BreakLink();
                            }
                            else if (link.IsField)
                            {
                                object value = item.Update();
                                if (value != null) try { success = Convert.ToBoolean(value); } catch { }
                            }
                            else item.LinkFormat.Update();
                            if (success) result.Success++;
                            else { result.Failed++; AddFail(result, LabelOf(link, index), "源返回失败（源文件不可用或链接已损坏）"); }
                        }
                        catch (Exception error)
                        {
                            result.Failed++; var com = error as COMException;
                            AddFail(result, LabelOf(link, index), com == null ? error.Message : string.Format("{0} (0x{1:X8})", com.Message, com.ErrorCode));
                        }
                    }
                }
                finally { state.Dispose(); }
                if (!unlink) result.CleanedHiddenHosts = EtCleanup.CleanupSpawnedHiddenEt(before);
                result.Executed = true;
            }
            catch (Exception error) { result.ErrorMessage = error.Message; result.Executed = false; }
            return result;
        }
        public static string JumpToItem(object application, LinkItemInfo item)
        {
            object flashRange; int prevHighlight;
            return JumpToItemEx(application, item, out flashRange, out prevHighlight);
        }

        // WdColorIndex：wdNoHighlight = 0，wdYellow = 7
        private const int WdYellow = 7;

        /// <summary>
        /// 与 JumpToItem 相同的定位，额外把目标区域**标黄**（wdYellow），
        /// 并把该区域 COM Range 和它原来的高亮值通过 out 参数交给 UI——
        /// 多个链接字段堆放在一起时，仅靠选区看不出跳的是哪一个，
        /// UI 层用定时器在几秒后调 ClearFlash 还原，形成"闪一下"的提示效果。
        /// 标黄失败不影响跳转本身（选区已到位），此时 flashRange 为 null。
        /// </summary>
        public static string JumpToItemEx(object application, LinkItemInfo item,
            out object flashRange, out int prevHighlight)
        {
            flashRange = null;
            prevHighlight = 0;
            if (application == null || item == null) return "参数为空。";
            try
            {
                dynamic app = application;
                dynamic doc = app.ActiveDocument;
                List<RawLink> links = CollectAllLinks(doc);
                int idx = item.Index - 1;
                if (idx < 0 || idx >= links.Count)
                    return "文档内容已变化（链接数量与列表对不上）。\r\n请点「刷新列表」后重试。";
                RawLink rl = links[idx];
                if (rl.Info.KindCode != item.KindCode
                    || rl.Info.RawCode != item.RawCode
                    || rl.Info.SourcePath != item.SourcePath)
                    return "文档内容已变化（该位置已不是原来那条链接）。\r\n请点「刷新列表」后重试。";

                dynamic d = rl.Com;
                object target = null;               // 成功选中的那个 Range（用于标黄）
                if (rl.IsField)
                {
                    // 选中字段代码区：字段以结果显示时，光标会落在结果所在处
                    d.Code.Select();

                    // ★ 标黄目标必须是【看得见的那段】：
                    //   字段以结果显示时，Code 是隐藏文字——高亮打在隐藏文字上
                    //   渲染不出来（v1.3.3~v1.3.5 的实测 bug：跳转成功、用户
                    //   却从没见过黄）。正常取 Result（字段显示出来的内容区）；
                    //   字段代码显示模式（Alt+F9）下 Result 反而是隐藏的，
                    //   改用 Code。取不到一律退回 Code。
                    try
                    {
                        bool showCodes = false;
                        try { showCodes = Convert.ToBoolean(d.ShowCodes); } catch { }
                        target = showCodes ? (object)d.Code : d.Result;
                    }
                    catch { }
                    if (target == null) target = d.Code;
                }
                else
                {
                    // 嵌入对象有 Range；浮动形状只有 Anchor
                    bool done = false;
                    try { d.Range.Select(); target = d.Range; done = true; }
                    catch { }
                    if (!done)
                    {
                        try { d.Anchor.Select(); target = d.Anchor; done = true; }
                        catch { }
                    }
                    if (!done) return "该对象无法定位（可能已被删除）。\r\n请点「刷新列表」。";
                }

                // 标黄提示：先记下原高亮值以便还原，失败则静默放弃（不影响跳转）
                if (target != null)
                {
                    try
                    {
                        dynamic r = target;
                        prevHighlight = (int)r.HighlightColorIndex;
                        r.HighlightColorIndex = WdYellow;
                        flashRange = target;
                    }
                    catch
                    {
                        flashRange = null;
                        prevHighlight = 0;
                    }
                }

                // 把焦点还给宿主主窗口，让用户立刻看到光标位置
                try { app.Activate(); }
                catch { }
                return null;
            }
            catch (Exception ex)
            {
                return "跳转失败：" + ex.Message;
            }
        }

        // ==================================================================
        //                          链接收集
        /// <summary>
        /// 还原 JumpToItemEx 标黄的区域。range 是跳转时返回的 COM Range，
        /// prevHighlight 是它原来的高亮值（WdColorIndex）。任何失败都静默——
        /// 晚几秒还原时文档可能已被编辑/关闭，绝不能因此抛异常打断用户。
        /// </summary>
        public static void ClearFlash(object range, int prevHighlight)
        {
            if (range == null) return;
            try
            {
                dynamic r = range;
                r.HighlightColorIndex = prevHighlight;
            }
            catch { }
        }

        // ==================================================================

        /// <summary>一个物理链接 = Word 对象模型里的 Field 或 Shape + 解析结果。</summary>
        private sealed class RawLink
        {
            public object Com;
            public bool IsField;
            public LinkItemInfo Info;
        }

        private static List<RawLink> CollectAllLinks(dynamic doc)
        {
            List<RawLink> list = new List<RawLink>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);

            List<object> stories = EnumerateStoryRanges(doc);
            for (int s = 0; s < stories.Count; s++)
            {
                string key = "S" + s.ToString();
                dynamic range = stories[s];
                CollectFields(range, key, list, seen);
                CollectInlineShapes(range, key, list, seen);
            }
            CollectDocShapes(doc, list, seen);
            return list;
        }

        /// <summary>展开文档的所有 story（正文/页眉/页脚/脚注…）；拿不到就退化成 doc.Content。</summary>
        private static List<object> EnumerateStoryRanges(dynamic doc)
        {
            List<object> list = new List<object>();
            dynamic ranges = null;
            try { ranges = doc.StoryRanges; }
            catch { }

            if (ranges == null)
            {
                try { list.Add((object)doc.Content); }
                catch { }
                return list;
            }

            int count = 0;
            try { count = (int)ranges.Count; }
            catch { }
            for (int i = 1; i <= count; i++)
            {
                object first = null;
                try { first = WordComCollection.Item(ranges, i); }
                catch { continue; }
                object cur = first;
                for (int guard = 0; guard < 64 && cur != null; guard++)
                {
                    list.Add(cur);
                    object next = null;
                    try { next = ((dynamic)cur).NextStoryRange; }
                    catch { }
                    cur = next;
                }
            }
            if (list.Count == 0)
            {
                try { list.Add((object)doc.Content); }
                catch { }
            }
            return list;
        }

        private static void CollectFields(dynamic range, string storyKey, List<RawLink> list, HashSet<string> seen)
        {
            dynamic fields = null;
            int count = 0;
            try
            {
                fields = range.Fields;
                count = (int)fields.Count;
            }
            catch { return; }

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object of = null;
                    try { of = WordComCollection.Item(fields, i); }
                    catch { of = null; }
                    if (of == null) continue;

                    LinkItemInfo info = FieldInfo(of);
                    if (info == null) continue;   // 不是链接字段

                    string key;
                    try
                    {
                        int pos = (int)((dynamic)of).Code.Start;
                        key = storyKey + "|F|" + pos.ToString() + "|" + info.RawCode;
                    }
                    catch { key = storyKey + "|F|" + i.ToString() + "|" + info.RawCode; }
                    if (!seen.Add(key)) continue;

                    RawLink rl = new RawLink();
                    rl.Com = of;
                    rl.IsField = true;
                    rl.Info = info;
                    list.Add(rl);
                }
                catch { }   // 单个字段异常不影响整体扫描
            }
        }

        private static void CollectInlineShapes(dynamic range, string storyKey, List<RawLink> list, HashSet<string> seen)
        {
            dynamic shapes = null;
            int count = 0;
            try
            {
                shapes = range.InlineShapes;
                count = (int)shapes.Count;
            }
            catch { return; }

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object os = null;
                    try { os = WordComCollection.Item(shapes, i); }
                    catch { os = null; }
                    if (os == null) continue;
                    dynamic sh = os;
                    dynamic lf = null;
                    try { lf = sh.LinkFormat; }
                    catch { }
                    if (lf == null) continue;

                    if (!AddShape(list, seen, storyKey, "嵌入对象", os, sh, lf, i)) continue;
                }
                catch { }
            }
        }

        /// <summary>浮动形状不挂在 story range 上，直接看 Document.Shapes。</summary>
        private static void CollectDocShapes(dynamic doc, List<RawLink> list, HashSet<string> seen)
        {
            dynamic shapes = null;
            int count = 0;
            try
            {
                shapes = doc.Shapes;
                count = (int)shapes.Count;
            }
            catch { return; }

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object os = null;
                    try { os = WordComCollection.Item(shapes, i); }
                    catch { os = null; }
                    if (os == null) continue;
                    dynamic sh = os;
                    dynamic lf = null;
                    try { lf = sh.LinkFormat; }
                    catch { }
                    if (lf == null) continue;

                    AddShape(list, seen, "SH", "浮动形状", os, sh, lf, i);
                }
                catch { }
            }
        }

        private static bool AddShape(List<RawLink> list, HashSet<string> seen, string storyKey,
            string itemKind, object com, dynamic sh, dynamic lf, int i)
        {
            LinkItemInfo info = ShapeInfo(lf, itemKind);
            if (info == null) return false;

            string key;
            try
            {
                int pos = (int)sh.Anchor.Start;
                key = storyKey + "|O|" + pos.ToString() + "|" + info.SourcePath;
            }
            catch { key = storyKey + "|O|" + i.ToString() + "|" + info.SourcePath; }
            if (!seen.Add(key)) return false;      // 与 InlineShapes 重复，跳过

            RawLink rl = new RawLink();
            rl.Com = com;
            rl.IsField = false;
            rl.Info = info;
            list.Add(rl);
            return true;
        }

        // ==================================================================
        //                          单个链接的信息提取
        // ==================================================================

        private static LinkItemInfo FieldInfo(object com)
        {
            dynamic f = com;
            string code = "";
            int type = -1;
            try { type = (int)f.Type; }
            catch { }
            try { code = (string)f.Code.Text; }
            catch { }
            if (code == null) code = "";

            string head = code.TrimStart();
            string headUp = head.ToUpperInvariant();
            bool isLink =
                type == WdFieldLink ||
                headUp.StartsWith("LINK ") ||
                headUp.StartsWith("DDE ") ||
                headUp.StartsWith("DDEAUTO ") ||
                headUp.StartsWith("INCLUDEPICTURE") ||
                headUp.StartsWith("INCLUDETEXT");
            if (!isLink) return null;

            LinkItemInfo info = new LinkItemInfo();
            info.ItemKind = "链接字段";
            info.RawCode = code.Trim();

            string path = "", sheet = "", rangeText = "", kindCode = "", switches = "";

            // 首选：复用「链接Excel」里那套成熟的 LINK/DDE 解析器
            bool multiArea = false;
            string parseError = null;
            LinkFieldCode.SourceTemplate t = null;
            try { t = LinkFieldCode.ParseLinkCode(code, out multiArea, out parseError); }
            catch { t = null; }

            if (t != null)
            {
                kindCode = t.Kind ?? "";
                path = t.WorkbookPath ?? t.PathOrTopic ?? "";
                sheet = TrimSuffix(t.SheetPart, "!");
                try
                {
                    // 多区域（Ctrl 多选复制）要列出全部片区，否则会误显示成只有第 1 片
                    rangeText = multiArea
                        ? LinkFieldCode.DescribeAreas(t)
                        : LinkFieldCode.DescribeRange(t);
                }
                catch { rangeText = ""; }
                if (multiArea) info.ItemKind = "链接字段（复合）";
                switches = (t.Switches ?? "").Trim();
            }
            else
            {
                // 兜底：多区域引用或部分非标准写法，只尽力取路径
                Match m = QuoteRegex.Match(code);
                if (m.Success) path = m.Groups[1].Value;
                if (headUp.StartsWith("LINK")) kindCode = "LINK";
                else if (headUp.StartsWith("DDEAUTO")) kindCode = "DDEAUTO";
                else if (headUp.StartsWith("DDE")) kindCode = "DDE";
                else
                {
                    int sp = head.IndexOf(' ');
                    kindCode = sp > 0 ? head.Substring(0, sp) : head;
                }
                if (headUp.StartsWith("INCLUDEPICTURE")) kindCode = "INCLUDEPICTURE";
                else if (headUp.StartsWith("INCLUDETEXT")) kindCode = "INCLUDETEXT";
                if (kindCode == null) kindCode = "";
                rangeText = multiArea ? "多区域引用" : "";
                info.ItemKind = "链接字段（复合）";
            }

            // DDE 主题形如 [C:\dir\Book.xlsx]Sheet1 —— 尽量把路径和工作表都拆出来：
            //   path  取自解析器已提取好的本地路径；
            //   sheet 只写在原始主题的 "]表名" 尾巴上，得单独补一次。
            string topicFull = (t != null && t.PathOrTopic != null) ? t.PathOrTopic : "";
            if (path.Length == 0) path = topicFull;
            if (sheet.Length == 0 && topicFull.StartsWith("["))
            {
                string onlySheet = "";
                SplitTopic(topicFull, ref onlySheet);
                sheet = onlySheet;
            }
            path = SplitTopic(path, ref sheet);

            info.KindCode = string.IsNullOrEmpty(kindCode) ? "LINK" : kindCode.ToUpperInvariant();
            info.SourcePath = NormalizePath(path);
            info.SheetName = CleanSheet(sheet);
            info.RangeText = string.IsNullOrEmpty(rangeText) ? "—" : rangeText;
            info.UpdateMode = IsAuto(info.KindCode, switches) ? "自动" : "手动";
            return info;
        }

        private static LinkItemInfo ShapeInfo(dynamic lf, string itemKind)
        {
            if (lf == null) return null;
            string src = "";
            try { src = (string)lf.SourceFullName; }
            catch { }
            if (string.IsNullOrEmpty(src))
            {
                try { src = (string)lf.SourceName; }
                catch { }
            }
            if (src == null) src = "";

            LinkItemInfo info = new LinkItemInfo();
            info.ItemKind = itemKind;
            info.KindCode = "OLE";
            info.SourcePath = NormalizePath(src);
            info.SheetName = "";
            info.RangeText = "—";

            bool auto = false;
            int au = -1;
            try { au = (int)lf.AutoUpdate; }
            catch { }
            if (au == 0) auto = false;
            else if (au > 0) auto = true;
            info.UpdateMode = auto ? "自动" : "手动";
            return info;
        }

        // ==================================================================
        //                          按源文件聚合
        // ==================================================================

        private static List<LinkSourceInfo> Aggregate(List<LinkItemInfo> items)
        {
            Dictionary<string, LinkSourceInfo> map =
                new Dictionary<string, LinkSourceInfo>(StringComparer.OrdinalIgnoreCase);
            List<LinkSourceInfo> order = new List<LinkSourceInfo>();

            for (int i = 0; i < items.Count; i++)
            {
                LinkItemInfo it = items[i];
                string path = it.SourcePath ?? "";
                string key = path.Length > 0 ? path : "(未知源)";

                LinkSourceInfo si;
                if (!map.TryGetValue(key, out si))
                {
                    si = new LinkSourceInfo();
                    bool known = key != "(未知源)";
                    si.FullPath = known ? path : "";
                    si.WorkbookName = known ? SafeFileName(path) : "(未能识别来源)";
                    si.FolderPath = known ? SafeDirectory(path) : "";
                    si.PathIsLocal = known && LooksLocal(path);
                    si.FileExists = si.PathIsLocal && SafeExists(path);
                    map[key] = si;
                    order.Add(si);
                }

                si.LinkCount++;
                if (it.ItemKind.IndexOf("字段") >= 0) si.FieldCount++;
                else si.ObjectCount++;
                if (it.UpdateMode == "自动") si.AutoCount++;
                si.Items.Add(it);

                string sh = it.SheetName ?? "";
                if (sh.Length > 0 && !si.Sheets.Contains(sh)) si.Sheets.Add(sh);
            }

            for (int i = 0; i < order.Count; i++)
            {
                LinkSourceInfo si = order[i];
                si.SheetText = si.Sheets.Count == 0 ? "—" : JoinNames(si.Sheets, 3);
                if (si.FieldCount > 0 && si.ObjectCount > 0) si.KindText = "字段 + 对象";
                else if (si.FieldCount > 0) si.KindText = "链接字段";
                else si.KindText = "链接对象";

                if (si.AutoCount == 0) si.UpdateText = "手动";
                else if (si.AutoCount == si.LinkCount) si.UpdateText = "自动";
                else si.UpdateText = string.Format("混合（自动 {0} / 手动 {1}）",
                    si.AutoCount, si.LinkCount - si.AutoCount);

                if (!si.PathIsLocal && string.IsNullOrEmpty(si.FullPath)) si.StatusText = "—";
                else if (!si.PathIsLocal) si.StatusText = "非本地路径";
                else si.StatusText = si.FileExists ? "正常" : "源文件缺失";
            }
            return order;
        }

        // ==================================================================
        //                              工具
        // ==================================================================

        /// <summary>DDE 主题可能是 [C:\x\Book.xlsx]Sheet1，拆成路径 + 工作表名。</summary>
        private static string SplitTopic(string path, ref string sheet)
        {
            if (string.IsNullOrEmpty(path)) return "";
            if (path[0] != '[') return path;
            int close = path.LastIndexOf(']');
            if (close <= 0) return path;
            if (string.IsNullOrEmpty(sheet) && close < path.Length - 1)
                sheet = path.Substring(close + 1);
            return path.Substring(1, close - 1);
        }

        private static bool IsAuto(string kindCode, string switches)
        {
            if (string.Equals(kindCode, "DDEAUTO", StringComparison.OrdinalIgnoreCase))
                return true;
            if (kindCode == "OLE") return false;
            string sw = switches ?? "";
            if (sw.Length == 0) return false;
            // 字段开关里的 \a 表示"打开文档时自动更新"
            int i = 0;
            while (i < sw.Length)
            {
                if (sw[i] == '\\' && i + 1 < sw.Length)
                {
                    char c = sw[i + 1];
                    if (c == 'a' || c == 'A') return true;
                    i += 2;
                }
                else i++;
            }
            return false;
        }

        private static string NormalizePath(string p)
        {
            if (string.IsNullOrEmpty(p)) return "";
            string s = p.Trim();
            if (s.Length >= 2 && s[0] == '"') s = s.Trim('"').Trim();
            // UNC / URL / DDE 服务名等无法本地化的原样保留
            return s;
        }

        private static bool LooksLocal(string p)
        {
            if (string.IsNullOrEmpty(p)) return false;
            if (p.StartsWith("\\\\")) return true;                 // UNC
            if (p.IndexOf("://", StringComparison.Ordinal) >= 0) return false;
            if (p.Length >= 2 && p[1] == ':') return true;          // C:\...
            if (p.IndexOf('\\') >= 0) return true;                  // 含路径分隔符
            return false;                                            // 只有书名的裸主题
        }

        private static string SafeFileName(string p)
        {
            try
            {
                string s = Path.GetFileName(p);
                return string.IsNullOrEmpty(s) ? p : s;
            }
            catch { return p; }
        }

        private static string SafeDirectory(string p)
        {
            try
            {
                string s = Path.GetDirectoryName(p);
                return s ?? "";
            }
            catch { return ""; }
        }

        private static bool SafeExists(string p)
        {
            try { return File.Exists(p); }
            catch { return false; }
        }

        private static string CleanSheet(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            string t = TrimSuffix(s.Trim(), "!");
            if (t.Length >= 2 && t[0] == '\'' && t[t.Length - 1] == '\'')
                t = t.Substring(1, t.Length - 2);          // 'My Sheet' 这种带引号的写法
            return t;
        }

        private static string TrimSuffix(string s, string suffix)
        {
            if (string.IsNullOrEmpty(s)) return "";
            if (s.EndsWith(suffix, StringComparison.Ordinal))
                return s.Substring(0, s.Length - suffix.Length);
            return s;
        }

        private static string JoinNames(List<string> names, int max)
        {
            if (names.Count <= max) return string.Join("、", names.ToArray());
            string[] head = new string[max];
            for (int i = 0; i < max; i++) head[i] = names[i];
            return string.Join("、", head) + string.Format(" 等 {0} 个", names.Count);
        }

        private static string LabelOf(RawLink rl, int index)
        {
            LinkItemInfo info = rl.Info;
            if (info == null) return string.Format("第 {0} 个链接", index + 1);
            string src = string.IsNullOrEmpty(info.SourcePath) ? "(未知源)" : SafeFileName(info.SourcePath);
            string tail = info.SheetName.Length > 0
                ? info.SheetName + "!" + info.RangeText
                : info.RangeText;
            return string.Format("第 {0} 个链接（{1} {2}）", index + 1, src, tail);
        }

        private static void AddFail(LinkActionResult r, string label, string reason)
        {
            if (r.Failures.Count < 20)
                r.Failures.Add(label + "：" + reason);
        }

    }
}
