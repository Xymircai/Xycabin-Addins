using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace WordUpdater
{
    // =====================================================================
    // 批量更新引擎 —— 纯逻辑、无 UI，便于自动化测试与宿主复用。
    //
    // 通过 dynamic 对 WPS 对象模型做后期绑定（Word 兼容子集）：
    //   app.Documents / app.Selection / sel.Range
    //   range.Fields        —— LINK / DDE / INCLUDEPICTURE / INCLUDETEXT 字段
    //   range.InlineShapes  —— 嵌入型链接图片（Excel 粘贴链接图片走这里）
    //   sel.ShapeRange      —— 选区内浮动形状中的链接对象
    // =====================================================================

    public enum RunStatus
    {
        NoDocument,   // 没有打开的文档
        NoSelection,  // 没有选中内容
        NoLinks,      // 选区内没有链接
        Completed,    // 执行完毕（含部分失败）
        Error         // 引擎级错误
    }

    public class LinkFailureInfo
    {
        public int Index;
        public string Kind;    // “链接字段” / “链接对象”
        public string Source;  // 源文件路径（尽力提取）
        public string Reason;  // 失败原因
    }

    /// <summary>
    /// 一条【数值发生变动】的链接：更新前后显示文本不同。
    /// FlashRange 是已打上黄的可见 Range（最多存在 2.5 秒，UI 还原后置空）。
    /// </summary>
    public class LinkChangeInfo
    {
        public int Index;           // 收集序号（1 起始，与失败明细同一套编号，便于对照）
        public string Kind;         // “链接字段”（只有字段参与文本比对）
        public string Where;        // “第 2 行第 2 列” / “正文第 3 段” / “正文（表格外）”
        public string OldText;      // 更新前文本（已归一化）
        public string NewText;      // 更新后文本（已归一化）
        public object FlashRange;   // 被打黄的那个 Range；null = 没打上（只报不变黄）
        public int PrevHighlight;   // 打黄前的高亮值（还原时按原值，不是清零）
    }

    public class RunResult
    {
        public RunStatus Status;
        public int Found;
        public int Success;
        public int Failed;
        public List<LinkFailureInfo> Failures = new List<LinkFailureInfo>();
        public string ErrorMessage;
        public int CleanedHiddenHosts;   // 更新后自动清理的隐藏源程序进程数
        // ↓ 数值变动检测（变动条数 = Changes.Count，不另设计数字段以免两处状态不同步）
        public List<LinkChangeInfo> Changes = new List<LinkChangeInfo>();
        public int NoTextCount;           // 更新成功但【没有可比对文本】的条数（图片/对象 + 读取失败）
    }

    public static class UpdateEngine
    {
        // ---- WPS/Word 常量 ----
        private const int WdFieldLink = 56;          // wdFieldLink
        private const int WdNoSelection = 0;         // wdNoSelection
        private const int WdSelectionIP = 1;          // wdSelectionIP（插入点，未选内容）
        private const int WdAlertsNone = 0;           // wdAlertsNone

        private static readonly Regex QuoteRegex = new Regex("\"([^\"]+)\"", RegexOptions.Compiled);

        public static RunResult Execute(object application)
        {
            RunResult r = new RunResult();
            if (application == null)
            {
                r.Status = RunStatus.Error;
                r.ErrorMessage = "宿主应用对象为空。";
                return r;
            }

            dynamic app = application;
            WordHostStateScope state = null;

            try
            {
                // ---------- 1. 文档检查 ----------
                int docCount = 0;
                try { docCount = (int)app.Documents.Count; }
                catch { }
                if (docCount == 0)
                {
                    r.Status = RunStatus.NoDocument;
                    return r;
                }

                // ---------- 2. 选区检查 ----------
                dynamic sel = null;
                try { sel = app.Selection; }
                catch { }
                if (sel == null)
                {
                    r.Status = RunStatus.NoSelection;
                    return r;
                }

                int selType = -1;
                try { selType = (int)sel.Type; }
                catch { }

                dynamic range = null;
                try { range = sel.Range; }
                catch { }
                if (range == null)
                {
                    r.Status = RunStatus.NoSelection;
                    return r;
                }

                int start = -1, endPos = -1;
                try { start = (int)range.Start; endPos = (int)range.End; }
                catch { }

                if (selType == WdNoSelection || selType == WdSelectionIP || (start >= 0 && start == endPos))
                {
                    r.Status = RunStatus.NoSelection;
                    return r;
                }

                // ---------- 3. 收集选区内的链接（先收集后更新，避免集合失效） ----------
                List<object> items = new List<object>();      // Field 或 Shape/LinkFormat 宿主
                List<string> kinds = new List<string>();
                List<string> sources = new List<string>();
                // WPS 的 sel.ShapeRange 会重复包含嵌入型图片（与 range.InlineShapes 同对象），
                // 按锚点位置去重，避免同一图片更新两次（第二次引用已失效会报 null 绑定错误）。
                HashSet<int> seenPositions = new HashSet<int>();

                CollectLinkFields(range, items, kinds, sources, seenPositions);
                CollectInlineShapes(range, items, kinds, sources, seenPositions);
                CollectShapeRange(sel, items, kinds, sources, seenPositions);

                r.Found = items.Count;
                if (r.Found == 0)
                {
                    r.Status = RunStatus.NoLinks;
                    return r;
                }

                // ---------- 4. 抑制弹窗，逐个更新 ----------
                state = WordHostStateScope.Enter(application);

                // 快照更新前的 et.exe（WPS 表格）进程，用于事后识别
                // 本次更新新拉起的隐藏 DDE/OLE 源程序实例。
                HashSet<int> etBefore = EtCleanup.SnapshotEtPids();

                for (int i = 0; i < items.Count; i++)
                {
                    object item = items[i];
                    string kind = kinds[i];
                    string source = sources[i];
                    int displayIndex = i + 1;
                    try
                    {
                        bool ok;
                        dynamic d = item;
                        if (item is FieldEntryHolder)
                        {
                            dynamic f = ((FieldEntryHolder)item).Field;

                            // ★ 更新前先取值，更新后才能比对出"数值是否变动"
                            string before = ReadResultText(f);

                            // LINK 字段：Update() 返回 Boolean，false 表示失败
                            object ret = f.Update();
                            if (ret == null) ok = true;
                            else
                            {
                                try { ok = Convert.ToBoolean(ret); }
                                catch { ok = true; }  // 返回值无法解析但未抛异常，视为成功
                            }

                            // 只有更新成功才比对（失败项不进变动清单、不打黄）
                            if (ok) RecordChange(r, f, displayIndex, kind, before);
                        }
                        else
                        {
                            // 链接图片/形状：LinkFormat.Update() 失败时抛异常
                            d.LinkFormat.Update();
                            ok = true;
                            r.NoTextCount++;   // 图片/对象没有可比对的数字文本，只算已更新
                        }
                        if (ok) r.Success++;
                        else AddFailure(r, displayIndex, kind, source, "更新返回失败（源文件不可用或链接已损坏）");
                    }
                    catch (COMException ex)
                    {
                        AddFailure(r, displayIndex, kind, source,
                            string.Format("{0} (0x{1:X8})", ex.Message, ex.ErrorCode));
                    }
                    catch (Exception ex)
                    {
                        AddFailure(r, displayIndex, kind, source, ex.Message);
                    }
                }

                // 只诊断本次新出现的隐藏实例；不等待，也不关闭非自有进程。
                r.CleanedHiddenHosts = EtCleanup.CleanupSpawnedHiddenEt(etBefore);

                r.Status = RunStatus.Completed;
                return r;
            }
            catch (Exception ex)
            {
                r.Status = RunStatus.Error;
                r.ErrorMessage = ex.Message;
                return r;
            }
            finally
            {
                if (state != null) try { state.Dispose(); }
                catch (Exception error)
                {
                    r.Status = RunStatus.Error;
                    r.ErrorMessage = (r.ErrorMessage == null ? "" : r.ErrorMessage + "\r\n") + error.Message;
                    UiDiagnostics.Record("word-update-state", error);
                }
            }
        }

        // ---------- 收集：范围内的链接字段 ----------
        private static void CollectLinkFields(dynamic range, List<object> items, List<string> kinds, List<string> sources, HashSet<int> seen)
        {
            dynamic fields = null;
            int count = 0;
            try
            {
                fields = range.Fields;
                count = (int)fields.Count;
            }
            catch { }
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic f = WordComCollection.Item(fields, i);
                    if (IsLinkField(f))
                    {
                        int pos = -1;
                        try { pos = (int)f.Result.Start; } catch { }
                        if (pos >= 0 && !seen.Add(pos)) continue;   // 同位置已收集，跳过
                        items.Add(new FieldEntryHolder(f));
                        kinds.Add("链接字段");
                        sources.Add(FieldSource(f));
                    }
                }
                catch { }  // 单个字段收集失败不影响整体
            }
        }

        // ---------- 收集：范围内的嵌入型链接对象 ----------
        private static void CollectInlineShapes(dynamic range, List<object> items, List<string> kinds, List<string> sources, HashSet<int> seen)
        {
            dynamic shapes = null;
            int count = 0;
            try
            {
                shapes = range.InlineShapes;
                count = (int)shapes.Count;
            }
            catch { }
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic sh = WordComCollection.Item(shapes, i);
                    dynamic lf = null;
                    try { lf = sh.LinkFormat; }
                    catch { }
                    if (lf != null)
                    {
                        int pos = AnchorStart(sh);
                        if (pos >= 0 && !seen.Add(pos)) continue;
                        items.Add(sh);
                        kinds.Add("链接对象");
                        sources.Add(ShapeSource(lf));
                    }
                }
                catch { }
            }
        }

        // ---------- 收集：选区内的浮动链接形状 ----------
        private static void CollectShapeRange(dynamic sel, List<object> items, List<string> kinds, List<string> sources, HashSet<int> seen)
        {
            try
            {
                dynamic sr = sel.ShapeRange;
                int count = 0;
                try { count = (int)sr.Count; }
                catch { }
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic sh = WordComCollection.Item(sr, i);
                        dynamic lf = null;
                        try { lf = sh.LinkFormat; }
                        catch { }
                        if (lf != null)
                        {
                            int pos = AnchorStart(sh);
                            if (pos >= 0 && !seen.Add(pos)) continue;   // 与 InlineShapes 去重
                            items.Add(sh);
                            kinds.Add("链接对象");
                            sources.Add(ShapeSource(lf));
                        }
                    }
                    catch { }
                }
            }
            catch { }  // 选区无形状时 ShapeRange 可能抛异常，忽略
        }

        // ---------- 取形状锚点位置：InlineShape 走 Range，浮动 Shape 走 Anchor ----------
        private static int AnchorStart(dynamic sh)
        {
            try { return (int)sh.Range.Start; }
            catch { }
            try { return (int)sh.Anchor.Start; }
            catch { }
            return -1;
        }

        // ---------- 判断字段是否为链接字段 ----------
        private static bool IsLinkField(dynamic f)
        {
            try
            {
                int t = (int)f.Type;
                if (t == WdFieldLink) return true;
            }
            catch { }
            try
            {
                string code = ((string)f.Code.Text) ?? "";
                string s = code.TrimStart().ToUpperInvariant();
                if (s.StartsWith("LINK ") || s.StartsWith("DDE ") || s.StartsWith("DDEAUTO ")
                    || s.StartsWith("INCLUDEPICTURE") || s.StartsWith("INCLUDETEXT"))
                    return true;
            }
            catch { }
            return false;
        }

        // ---------- 提取字段源路径（字段代码中第一个带引号内容） ----------
        private static string FieldSource(dynamic f)
        {
            try
            {
                string code = (string)f.Code.Text;
                if (!string.IsNullOrEmpty(code))
                {
                    Match m = QuoteRegex.Match(code);
                    if (m.Success) return m.Groups[1].Value;
                }
            }
            catch { }
            return "";
        }

        private static string ShapeSource(dynamic linkFormat)
        {
            try { return (string)linkFormat.SourceFullName; }
            catch { return ""; }
        }

        private static void AddFailure(RunResult r, int index, string kind, string source, string reason)
        {
            LinkFailureInfo f = new LinkFailureInfo();
            f.Index = index;
            f.Kind = kind ?? "";
            f.Source = source ?? "";
            f.Reason = reason ?? "";
            r.Failures.Add(f);
            r.Failed++;
        }

        // ==================================================================
        //          数值变动检测：更新前取值 → 更新 → 重新取值比对 → 标黄
        // ==================================================================

        /// <summary>读字段当前显示出来的文本；读不到返回 null（= 不参与比对）。</summary>
        private static string ReadResultText(dynamic f)
        {
            try { return NormalizeText(f.Result.Text); }
            catch { return null; }
        }

        /// <summary>
        /// 归一化：字段结果显示区常带 \r / \a（单元格结束符之类），
        /// 不去掉会让"值其实没变"的链接被误报成变动。
        /// </summary>
        private static string NormalizeText(object text)
        {
            string s = text as string;
            if (s == null) return null;
            char[] noise = { '\r', '\n', '\a', '\0', ' ', '\t' };
            return s.Trim(noise);
        }

        /// <summary>
        /// 字段更新【成功】后调用：重新取值 → 与更新前比对 → 定位 → 打黄 → 记入清单。
        /// 全程 try/catch：任何一步出问题最多是"这条没标黄但文案提到了变动"，
        /// 绝不能冒泡出去把整个更新变成引擎级错误。
        /// </summary>
        private static void RecordChange(RunResult r, dynamic f, int index, string kind, string before)
        {
            object target = null;
            string after = null;
            try
            {
                bool showCodes = false;
                try { showCodes = Convert.ToBoolean(f.ShowCodes); } catch { }

                // ★ 必须【重新】从 Field 取一次 Result，不能复用更新前拿到的那个 Range
                //   ——WPS 里二次引用已失效的对象会报 null 绑定错误
                //   （本文件收集段关于 ShapeRange 的注释是同一类问题）。
                dynamic res = f.Result;
                after = NormalizeText(res.Text);

                // ★ 黄要打在【看得见的那段】：字段以结果显示时 Code 是隐藏文字，
                //   高亮打上去渲染不出来；ShowCodes（Alt+F9）下则反过来。
                target = showCodes ? (object)f.Code : (object)res;
            }
            catch { }
            if (target == null)
            {
                try { target = (object)f.Code; } catch { }   // 兜底退回 Code
            }

            if (before == null || after == null)
            {
                r.NoTextCount++;   // 前后有一头读不到文本，无从比对
                return;
            }
            if (string.Equals(before, after, StringComparison.Ordinal)) return;   // 没变，安静返回

            LinkChangeInfo c = new LinkChangeInfo();
            c.Index = index;
            c.Kind = kind ?? "";
            c.OldText = before;
            c.NewText = after;
            c.Where = DescribePosition(target);

            int prev = 0;
            if (FlashHighlighter.Apply(target, out prev))
            {
                c.FlashRange = target;
                c.PrevHighlight = prev;
            }
            r.Changes.Add(c);
        }

        /// <summary>
        /// 把 Range 翻译成人看得懂的位置，用于弹窗说"第 2 行第 2 列"。
        /// 任何一级读不到都静默降级，绝不中断更新。
        /// </summary>
        private static string DescribePosition(object targetRange)
        {
            if (targetRange == null) return "（位置未知）";
            dynamic rg = targetRange;

            // ① 是否在表格内：Range.Information(12) = wdWithInTable
            bool inTable = false;
            try { inTable = Convert.ToBoolean(rg.Information(12)); }
            catch { }
            if (!inTable) return DescribeOutsideTable(rg);

            // ② 表格内：取首格的 RowIndex / ColumnIndex
            //    （LinkPasteGeometry 里判定目标矩形用的是同一套读法）
            try
            {
                dynamic cells = rg.Cells;
                int n = (int)cells.Count;
                if (n > 0)
                {
                    dynamic cell = WordComCollection.Item(cells, 1);
                    int row = (int)cell.RowIndex;
                    int col = (int)cell.ColumnIndex;
                    if (row > 0 && col > 0)
                        return string.Format("第 {0} 行第 {1} 列", row, col);
                }
            }
            catch { }
            return "表格内（行列未能识别）";
        }

        /// <summary>表格外的链接：段落号 → 页码 → 兜底文案，三级退让。</summary>
        private static string DescribeOutsideTable(dynamic rg)
        {
            // ① 段落序号：数一数从文档开头到本 Range 起点之间有多少个段落
            try
            {
                dynamic doc = rg.Document;
                int st = (int)rg.Start;
                dynamic head = doc.Range(0, st);
                int cnt = (int)head.Paragraphs.Count;
                if (cnt >= 0) return string.Format("正文第 {0} 段", cnt + 1);
            }
            catch { }
            // ② 页码兜底：Information(3) = wdActiveEndPageNumber
            try
            {
                int page = (int)rg.Information(3);
                if (page > 0) return string.Format("正文（第 {0} 页）", page);
            }
            catch { }
            return "正文（表格外）";
        }

        // 字段包装类：区分“字段”与“形状”两条更新路径
        private class FieldEntryHolder
        {
            public readonly dynamic Field;
            public FieldEntryHolder(dynamic field) { Field = field; }
        }

        // 隐藏源程序（et.exe）清理逻辑已抽取至 EtCleanup.cs（与 LinkPasteEngine 共用）。
    }
}
