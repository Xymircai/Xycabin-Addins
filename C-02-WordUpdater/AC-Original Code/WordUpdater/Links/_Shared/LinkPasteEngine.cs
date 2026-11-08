using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace WordUpdater
{
    // =====================================================================
    // 「链接Excel」引擎 —— 纯逻辑、无 UI，便于自动化测试与宿主复用。
    //
    // 功能：用户在 WPS 表格中复制一块连续单元格区域（如 A1:D4），在
    // WPS 文字中选中同样格数的规则矩形，本引擎按行优先顺序把
    // 选中单元格与源区域对应单元格以“粘贴为链接”的方式绑定。
    //
    // 实现思路（关键：不自行猜测链接字段格式，而是让 WPS 自己生成）：
    //   1. 新建一个临时文档，对剪贴板做 Selection.PasteSpecial(Link:=true)
    //      （依次尝试 无格式文本 / RTF / OLE 对象），让 WPS 用自己的
    //      机制生成一条针对整块区域的链接字段；
    //   2. 读取该字段的代码（LINK ... "工作簿" "Sheet1!R1C1:R4C4" ... 或
    //      DDE/DDEAUTO ...），解析出 类名/服务名、工作簿路径或 DDE 主题、
    //      工作表、源区域行列范围、尾部开关；
    //   3. 全部校验通过后（选区规则、无合并、格数一致、源文件存在），
    //      以同一份 类名/路径/开关 逐格重建链接字段，仅把引用部分换成
    //      对应的单一源单元格 —— 与手工“选择性粘贴→粘贴链接”产生的
    //      字段完全同构，因此天然沿用现有「更新链接」按钮的批量更新机制；
    //   4. 写入阶段任何失败都会回滚已写入字段，不产生残缺链接。
    // =====================================================================

    public enum LinkPasteStatus
    {
        NoDocument,       // 没有打开的文档
        NotInTable,       // 光标未位于表格内
        TableNotUniform,  // 保留原状态值；仅表示所选目标结构不能逐格对应
        TableTooLarge,    // 目标表格单元格数超上限
        ClipboardInvalid, // 剪贴板内容不是可链接的单元格区域
        MultiArea,        // 复制了多个不连续区域（已支持，仅作解析标记）
        SizeMismatch,     // 源格数与目标选区格数不一致（溢出或未填满）
        SourceMissing,    // 源工作簿文件不存在
        Completed,        // 执行完毕（含个别刷新失败）
        Error,            // 引擎级错误
        SelectionEmpty,       // 没有选中任何表格单元格
        SelectionOutOfTable,  // 选区跨出表格边界，或横跨多张表格
        SelectionNotRect      // 选区不是规则矩形
    }

    public class LinkPasteResult
    {
        public LinkPasteStatus Status;
        public string SourceDesc = "";   // 源描述（工作簿 + 区域）
        public int SourceRows;           // 源主区域行数
        public int SourceCols;           // 源主区域列数
        public int SourceCells;          // 源序列长度（格数，多区域已展开）
        public int TableRows;            // 目标表格行数（整表，仅作参考）
        public int TableCols;            // 目标表格列数（整表，仅作参考）
        public int TargetRow1;           // 目标矩形左上角行
        public int TargetCol1;           // 目标矩形左上角列
        public int TargetRows;           // 目标矩形行数
        public int TargetCols;           // 目标矩形列数
        public int TargetCells;          // 目标矩形格数
        public int CellsLinked;          // 写入的链接字段数
        public int Refreshed;            // 立即刷新成功的字段数
        public List<string> Failures = new List<string>(); // 刷新失败原因
        public string ErrorMessage;
        public bool RolledBack;          // 写入中途失败已回滚
        public int CleanedHiddenHosts;   // 刷新后自动清理的隐藏源程序进程数
    }

    public static class LinkPasteEngine
    {
        // ---- WPS/Word 常量 ----
        private const int WdWithInTable = 12;   // wdWithInTable（Information 参数）
        private const int WdFieldEmpty = -1;    // wdFieldEmpty（按代码文本插入字段）
        private const int WdFieldLink = 56;     // wdFieldLink
        private const int WdPasteOle = 0;       // wdPasteOLEObject
        private const int WdPasteRtf = 1;       // wdPasteRTF
        private const int WdPasteText = 2;      // wdPasteText
        internal const int MaxCells = 500;      // 单次支持的目标单元格上限

        public static LinkPasteResult Execute(object application)
        {
            LinkPasteResult r = new LinkPasteResult();
            if (application == null)
            {
                r.Status = LinkPasteStatus.Error;
                r.ErrorMessage = "宿主应用对象为空。";
                return r;
            }

            dynamic app = application;
            WordHostStateScope state = null;
            List<LinkTargetCell> targets = null;

            try
            {
                // ---------- 1. 文档检查 ----------
                int docCount = 0;
                try { docCount = (int)app.Documents.Count; }
                catch { }
                if (docCount == 0)
                {
                    r.Status = LinkPasteStatus.NoDocument;
                    return r;
                }

                // ---------- 2. 选区必须位于表格内 ----------
                dynamic sel = null;
                try { sel = app.Selection; }
                catch { }
                if (sel == null)
                {
                    r.Status = LinkPasteStatus.Error;
                    r.ErrorMessage = "无法获取当前选区。";
                    return r;
                }

                bool inTable = false;
                try { inTable = Convert.ToBoolean(sel.Information(WdWithInTable)); }
                catch { }
                if (!inTable)
                {
                    r.Status = LinkPasteStatus.NotInTable;
                    return r;
                }

                dynamic table = null;
                try { table = WordComCollection.Item(sel.Tables, 1); }
                catch { }
                if (table == null)
                {
                    r.Status = LinkPasteStatus.NotInTable;
                    return r;
                }

                // ---------- 3. 整表尺寸仅作参考，目标合法性只由选中单元格决定 ----------
                try { r.TableRows = (int)table.Rows.Count; } catch { }
                try { r.TableCols = (int)table.Columns.Count; } catch { }

                // ---------- 3b. 目标 = 用户选中的实际单元格矩形 ----------
                // 只认选区、不做任何扩展与兜底：想贴整张表，就得先把整张表选上。
                // 保存实际 Cell 和网格位置；写入时重新读取 Range，不使用漂移的字符偏移。
                LinkTargetRect rect = null;
                string rectDetail = null;
                TargetRectStatus rstat =
                    LinkPasteGeometry.ResolveTargetCells(sel, table, out rect, out targets, out rectDetail);
                if (rstat != TargetRectStatus.Ok)
                {
                    r.ErrorMessage = rectDetail;
                    r.Status = TargetStatus(rstat);
                    return r;
                }
                r.TargetRow1 = rect.Row1;
                r.TargetCol1 = rect.Col1;
                r.TargetRows = rect.Rows;
                r.TargetCols = rect.Cols;
                r.TargetCells = rect.CellCount;
                if (rect.CellCount > MaxCells)
                {
                    r.Status = LinkPasteStatus.TableTooLarge;
                    return r;
                }

                // 目标文档（后续探测与写入都在它上面做，避免新建文档导致界面闪出）
                dynamic doc = null;
                try { doc = table.Range.Document; }
                catch { }
                if (doc == null)
                {
                    r.Status = LinkPasteStatus.Error;
                    r.ErrorMessage = "无法定位目标文档。";
                    return r;
                }

                // 记录原始选区，探测结束后恢复（粘贴会临时移动光标）
                int selStart = -1, selEnd = -1;
                try
                {
                    dynamic sr = sel.Range;
                    selStart = (int)sr.Start;
                    selEnd = (int)sr.End;
                }
                catch { }

                // ---------- 4. 抑制弹窗与刷屏 ----------
                state = WordHostStateScope.Enter(application);

                // ---------- 5. 剪贴板探测：让 WPS 自己生成链接字段 ----------
                string code = ProbeClipboardLinkCode(app, doc);
                // 探测过程在文档末尾做过“粘贴→读取→立即删除”，这里恢复原选区
                if (selStart >= 0 && selEnd >= selStart)
                {
                    try { doc.Range(selStart, selEnd).Select(); }
                    catch { }
                }
                if (string.IsNullOrEmpty(code))
                {
                    r.Status = LinkPasteStatus.ClipboardInvalid;
                    r.ErrorMessage = "对剪贴板内容执行“粘贴链接”未产生链接字段。";
                    return r;
                }
                Connect.Log("链接Excel：剪贴板链接字段代码 = " + code);

                // ---------- 6. 解析链接字段代码 ----------
                bool multiArea;
                string parseErr;
                LinkFieldCode.SourceTemplate t = LinkFieldCode.ParseLinkCode(code, out multiArea, out parseErr);
                // 注意：多区域（Ctrl 多选）**不再拒绝**——它正是要支持的能力，
                // 逐片展开后同样按位对应。LinkPasteStatus.MultiArea 仅作解析标记保留。
                if (t == null)
                {
                    r.Status = LinkPasteStatus.ClipboardInvalid;
                    r.ErrorMessage = parseErr;
                    return r;
                }
                r.SourceDesc = t.Describe();
                r.SourceRows = t.R2 - t.R1 + 1;
                r.SourceCols = t.C2 - t.C1 + 1;
                long rawCount = t.RawCount;
                r.SourceCells = rawCount > int.MaxValue ? int.MaxValue : (int)rawCount;
                Connect.Log(string.Format(
                    "链接Excel：解析成功 kind={0} head={1} 路径/主题={2} 引用={3} 主区域={4}x{5} 总格数={6} A1={7}",
                    t.Kind, t.Head, t.PathOrTopic, t.SheetPart + LinkFieldCode.DescribeRange(t),
                    r.SourceRows, r.SourceCols, rawCount, t.A1Style));

                // ---------- 6b. 路径 B：剪贴板只给「包围盒」时的补救 ----------
                // 实测：WPS 表格里 Ctrl 多选 H1、H3 复制后，WPS 文字生成的字段代码是
                // "Sheet1!R1C8:R3C8"，与连选 H1:H3 **逐字符相同** —— 剪贴板丢掉了不连续性。
                // 好在「复制」不会改变 WPS 表格里的选区，此刻回读 ET 的 Selection.Areas
                // 就能拿到逐片明细（Areas.Count=2：$H$1、$H$3）。
                //
                // 触发条件刻意设得极窄：引用只有 1 片，且该片格数 ≠ 目标格数。
                // 连选复制（格数天然一致）根本走不到这里，因此零回归。
                // 任何一项校验不过都静默降级为原行为，绝不报错、绝不打扰。
                if (t.Areas.Count == 1 && rect.CellCount > 0
                    && t.Areas[0].Count != (long)rect.CellCount)
                {
                    string why;
                    if (TryApplyEtSelection(t, rect.CellCount, out why))
                    {
                        rawCount = t.RawCount;
                        r.SourceDesc = t.Describe();
                        r.SourceRows = t.R2 - t.R1 + 1;
                        r.SourceCols = t.C2 - t.C1 + 1;
                        r.SourceCells = rawCount > int.MaxValue ? int.MaxValue : (int)rawCount;
                        Connect.Log(string.Format(
                            "链接Excel：剪贴板只给包围盒，已改用 WPS 表格实时选区（{0} 片 / {1} 格）：{2}",
                            t.Areas.Count, rawCount, LinkFieldCode.DescribeAreas(t)));
                    }
                    else
                    {
                        Connect.Log("链接Excel：未启用表格实时选区（" + why + "）");
                    }
                }

                // ---------- 7. 数量匹配：源序列长度必须等于目标选区格数 ----------
                // 严格匹配、不做任何扩展或兜底；不等即拒绝，一个字都不写。
                if (rawCount > MaxCells)
                {
                    r.Status = LinkPasteStatus.SizeMismatch;
                    return r;
                }
                List<LinkSourceCell> srcCells = LinkFieldCode.ExpandSource(t);
                if (srcCells.Count != targets.Count)
                {
                    r.Status = LinkPasteStatus.SizeMismatch;
                    Connect.Log(string.Format("链接Excel：数量不匹配 源={0} 目标={1}",
                        srcCells.Count, targets.Count));
                    return r;
                }
                Connect.Log("链接Excel：目标=" + rect.ToString() + "（" + targets.Count +
                    " 格），源序列=" + srcCells.Count + " 格");

                // ---------- 8. 源文件存在性 ----------
                if (!string.IsNullOrEmpty(t.WorkbookPath) && LinkFieldCode.LooksLocal(t.WorkbookPath)
                    && !File.Exists(t.WorkbookPath))
                {
                    r.Status = LinkPasteStatus.SourceMissing;
                    return r;
                }

                // ---------- 9. 写入前复核原选区，不依赖整表 Cell/Rows 索引重新猜目标 ----------
                rstat = LinkPasteGeometry.RevalidateTargets(table, targets, out rectDetail);
                if (rstat != TargetRectStatus.Ok)
                {
                    r.Status = TargetStatus(rstat);
                    r.ErrorMessage = rectDetail;
                    return r;
                }
                object undoRec = null;
                try
                {
                    dynamic ur = app.UndoRecord;
                    ur.StartCustomRecord("链接Excel");
                    undoRec = ur;
                }
                catch { }   // WPS 不支持 UndoRecord 时忽略，不影响功能

                List<object> inserted = new List<object>();
                try
                {
                    // 写入：源序列 ↔ 已验证的实际选中单元格，按行优先一一对应。
                    for (int k = 0; k < targets.Count; k++)
                    {
                        dynamic cell = targets[k].Cell;
                        int tr = targets[k].Row;
                        int tc = targets[k].GridColumn;

                        // 目标格内容范围 = 单元格范围去掉结尾的单元格结束符
                        dynamic cr = cell.Range;
                        int cs = (int)cr.Start;
                        int ce = (int)cr.End;
                        dynamic rng = doc.Range(cs, ce - 1);
                        try { rng.Delete(); }
                        catch { try { rng.Text = ""; } catch { } }

                        string fld = LinkFieldCode.BuildCode(t, srcCells[k]);
                        dynamic field = null;
                        bool ok = false;
                        try { field = rng.Fields.Add(rng, WdFieldEmpty, fld, false); ok = true; }
                        catch { }
                        if (!ok)
                        {
                            // 兜底：显式按 wdFieldLink 类型插入（参数不含 LINK 关键字）
                            try
                            {
                                field = rng.Fields.Add(rng, WdFieldLink, StripKindPrefix(fld), false);
                                ok = true;
                            }
                            catch { }
                        }
                        if (!ok)
                            throw new PasteAbortException(LinkPasteStatus.Error,
                                string.Format("在第 {0} 行第 {1} 列写入链接字段失败。", tr, tc));
                        inserted.Add(field);
                    }
                }
                catch (PasteAbortException pae)
                {
                    Rollback(inserted);
                    EndUndo(undoRec);
                    r.Status = pae.Status;
                    r.RolledBack = true;
                    r.ErrorMessage = pae.Message;
                    r.CellsLinked = 0;
                    return r;
                }
                catch (Exception ex)
                {
                    Rollback(inserted);
                    EndUndo(undoRec);
                    r.Status = LinkPasteStatus.Error;
                    r.RolledBack = true;
                    r.ErrorMessage = "写入链接时发生异常：" + ex.Message;
                    r.CellsLinked = 0;
                    return r;
                }
                EndUndo(undoRec);
                r.CellsLinked = inserted.Count;

                // ---------- 10. 立即刷新，让目标格显示源数据当前值 ----------
                HashSet<int> etBefore = EtCleanup.SnapshotEtPids();
                for (int i = 0; i < inserted.Count; i++)
                {
                    try
                    {
                        dynamic f = inserted[i];
                        object ret = f.Update();
                        if (ret == null) r.Refreshed++;
                        else
                        {
                            try
                            {
                                if (Convert.ToBoolean(ret)) r.Refreshed++;
                                else AddFail(r, i + 1, "更新返回失败（源暂时不可用）");
                            }
                            catch { r.Refreshed++; }
                        }
                    }
                    catch (COMException ex)
                    {
                        AddFail(r, i + 1, string.Format("{0} (0x{1:X8})", ex.Message, ex.ErrorCode));
                    }
                    catch (Exception ex)
                    {
                        AddFail(r, i + 1, ex.Message);
                    }
                }
                // 刷新可能在源文件未打开时拉起隐藏 et.exe，复用既有清理策略
                try { System.Threading.Thread.Sleep(500); }
                catch { }
                r.CleanedHiddenHosts = EtCleanup.CleanupSpawnedHiddenEt(etBefore);

                r.Status = LinkPasteStatus.Completed;
                return r;
            }
            catch (Exception ex)
            {
                r.Status = LinkPasteStatus.Error;
                r.ErrorMessage = ex.Message;
                return r;
            }
            finally
            {
                LinkPasteGeometry.ReleaseTargets(targets);
                try { if (state != null) state.Dispose(); }
                catch (Exception error)
                {
                    r.Status = LinkPasteStatus.Error;
                    r.ErrorMessage += (r.ErrorMessage.Length == 0 ? "" : Environment.NewLine) + error.Message;
                }
            }
        }

        // ================= 剪贴板探测 =================

        /// <summary>
        /// 探测剪贴板链接字段代码。优先在**当前目标文档末尾**做临时粘贴
        /// （粘贴→读字段代码→立刻删除），这样不会新建文档，界面不会闪出
        /// 新的 WPS 文字标签页；只有当落点无法安全确定时，才回退到临时文档。
        /// 依次尝试 无格式文本 / RTF / OLE 对象 三种数据类型；
        /// 每种类型先按命名参数调用（最稳），失败再按位置参数调用。
        /// </summary>
        private static string ProbeClipboardLinkCode(dynamic app, dynamic doc)
        {
            int[] dataTypes = { WdPasteText, WdPasteRtf, WdPasteOle };
            for (int k = 0; k < dataTypes.Length; k++)
            {
                // 主路径：目标文档内临时粘贴（无新文件出现）
                string code = TryPasteInDocument(app, doc, dataTypes[k], true);
                if (!string.IsNullOrEmpty(code)) return code;
                code = TryPasteInDocument(app, doc, dataTypes[k], false);
                if (!string.IsNullOrEmpty(code)) return code;
                // 兜底：临时文档（文档结构特殊、无法安全临时插入时）
                code = TryPasteTempDocument(app, dataTypes[k], true);
                if (!string.IsNullOrEmpty(code)) return code;
                code = TryPasteTempDocument(app, dataTypes[k], false);
                if (!string.IsNullOrEmpty(code)) return code;
            }
            return null;
        }

        /// <summary>
        /// 一次尝试：把光标移到目标文档正文末尾最后一个段落标记之前 →
        /// 粘贴链接 → 读取刚粘贴内容的链接字段代码 → 立即删除粘贴内容。
        /// 由于 ScreenUpdating 已关闭且操作在同一文档内完成，用户不会看到任何变化。
        /// </summary>
        private static string TryPasteInDocument(dynamic app, dynamic doc, int dataType, bool namedArgs)
        {
            int pos = -1;
            try
            {
                int end;
                try { end = (int)doc.Content.End; }
                catch { return null; }
                pos = end - 1;                       // 正文末尾最后一个段落标记之前
                if (pos < 0) return null;

                dynamic anchor = doc.Range(pos, pos);
                // 落点若落在表格单元格内（可能破坏表格结构），放弃本方式
                bool inT = false;
                try { inT = Convert.ToBoolean(anchor.Information(WdWithInTable)); }
                catch { }
                if (inT) return null;

                anchor.Select();
                dynamic s = app.Selection;
                if (namedArgs)
                    s.PasteSpecial(IconIndex: 0, Link: true, Placement: 0,
                        DisplayAsIcon: false, DataType: dataType);
                else
                    s.PasteSpecial(0, true, 0, false, dataType);

                int newEnd;
                try { newEnd = (int)doc.Content.End; }
                catch { newEnd = pos; }
                dynamic pasted = doc.Range(pos, Math.Max(pos, newEnd - 1));
                string code = ReadLinkCode(pasted);
                CleanTail(doc, pos);                 // 无论成败都清掉临时粘贴内容
                return code;
            }
            catch
            {
                if (pos >= 0) CleanTail(doc, pos);
                return null;
            }
        }

        /// <summary>删除 [pos, 文末) 区间内本次粘贴产生的全部内容（保留末尾段落标记）。</summary>
        private static void CleanTail(dynamic doc, int pos)
        {
            try
            {
                int end;
                try { end = (int)doc.Content.End; }
                catch { return; }
                if (end - 1 <= pos) return;          // 本次没有新增内容
                dynamic rng = doc.Range(pos, end - 1);
                rng.Delete();
            }
            catch (Exception ex)
            {
                Connect.Log("链接Excel：清理临时粘贴内容失败：" + ex.Message);
            }
        }

        /// <summary>一次尝试：新建临时文档 → 粘贴链接 → 读字段代码 → 关闭（不保存）。仅在文档内插入不可行时使用。</summary>
        private static string TryPasteTempDocument(dynamic app, int dataType, bool namedArgs)
        {
            dynamic tempDoc = null;
            try
            {
                try { tempDoc = app.Documents.Add(); }
                catch { return null; }
                try { tempDoc.Activate(); }
                catch { }

                try
                {
                    dynamic s = app.Selection;
                    if (namedArgs)
                        s.PasteSpecial(IconIndex: 0, Link: true, Placement: 0,
                            DisplayAsIcon: false, DataType: dataType);
                    else
                        s.PasteSpecial(0, true, 0, false, dataType);
                }
                catch
                {
                    return null;   // 该数据类型粘贴链接失败（或剪贴板无 OLE 源）
                }

                // 读取链接字段代码：普通字段优先，其次内联链接对象（OLE 形态）
                return ReadLinkCode(tempDoc);
            }
            finally
            {
                if (tempDoc != null)
                {
                    try { tempDoc.Close(0); }
                    catch
                    {
                        try { tempDoc.Close(0); }
                        catch { }
                    }
                }
            }
        }

        /// <summary>
        /// 从容器（Range 或 Document）读取第一条链接字段代码：
        /// 普通 LINK/DDE 字段优先，其次内联链接对象（OLE 形态）。
        /// </summary>
        private static string ReadLinkCode(dynamic host)
        {
            try
            {
                int fc = (int)host.Fields.Count;
                if (fc >= 1)
                {
                    dynamic f = WordComCollection.Item(host.Fields, 1);
                    string c = (string)f.Code.Text;
                    if (!string.IsNullOrEmpty(c) && c.Trim().Length > 0) return c;
                }
            }
            catch { }
            try
            {
                int sc = (int)host.InlineShapes.Count;
                if (sc >= 1)
                {
                    dynamic sh = WordComCollection.Item(host.InlineShapes, 1);
                    string c = (string)sh.Field.Code.Text;
                    if (!string.IsNullOrEmpty(c) && c.Trim().Length > 0) return c;
                }
            }
            catch { }
            return null;     // 粘贴成功但没产生链接字段（Link 参数被忽略等）
        }

        /// <summary>
        /// 路径 B：剪贴板只给「包围盒」时，回读 WPS 表格的实时选区拿逐片明细。
        /// 三重校验（工作簿 / 工作表 / 落在包围盒内）全部通过才替换 t.Areas，
        /// 否则一律保持原样静默降级 —— 失败是常态，不是错误。
        /// </summary>
        private static bool TryApplyEtSelection(LinkFieldCode.SourceTemplate t, int targetCount, out string why)
        {
            why = null;

            // 来源工作簿可能只登记文件ROT，而默认Application指向空的隐藏实例。
            List<EtSelectionInfo> infos;
            try { infos = EtSelectionReader.TryReadAll(t.WorkbookPath, out why); }
            catch (Exception ex) { why = "读取表格选区异常：" + ex.Message; return false; }
            if (infos == null || infos.Count == 0) return false;

            return TryApplyEtSelections(t, infos, targetCount, out why);
        }

        private static bool TryApplyEtSelections(LinkFieldCode.SourceTemplate t, List<EtSelectionInfo> infos, int targetCount, out string why)
        {
            string last = "未找到与目标格数一致的多区域来源。";
            List<LinkFieldCode.SourceArea> original = t.Areas, chosen = null;
            for (int i = 0; i < infos.Count; i++)
            {
                if (infos[i] == null) continue;
                long count = 0;
                foreach (LinkFieldCode.SourceArea area in infos[i].Areas) count += area.Count;
                if (count != targetCount) continue;
                string w;
                if (ValidateEtSelection(t, infos[i], out w))
                {
                    if (chosen != null && !SameSourceAreas(chosen, t.Areas))
                    {
                        t.Areas = original;
                        why = "同一来源存在多个不同的窗口选区，无法确认刚复制的区域。";
                        return false;
                    }
                    chosen = t.Areas;
                }
                if (w != null) last = w;
            }
            t.Areas = chosen ?? original;
            why = chosen != null ? null : last;
            return chosen != null;
        }

        private static bool SameSourceAreas(List<LinkFieldCode.SourceArea> first, List<LinkFieldCode.SourceArea> second)
        {
            if (first.Count != second.Count) return false;
            for (int i = 0; i < first.Count; i++)
                if (first[i].R1 != second[i].R1 || first[i].C1 != second[i].C1
                    || first[i].R2 != second[i].R2 || first[i].C2 != second[i].C2) return false;
            return true;
        }

        /// <summary>校验某个 ET 选区能否代表本次剪贴板来源；通过则直接替换 t.Areas。</summary>
        private static bool ValidateEtSelection(LinkFieldCode.SourceTemplate t, EtSelectionInfo info, out string why)
        {
            why = null;
            if (info == null) { why = "无表格选区。"; return false; }

            // 校验 1：工作簿必须是同一个（未保存的工作簿没有路径 → 放弃路径 B）
            if (string.IsNullOrEmpty(info.WorkbookPath) || string.IsNullOrEmpty(t.WorkbookPath))
            {
                why = "工作簿路径缺失（源工作簿可能尚未保存）。";
                return false;
            }
            // 注意：WPS 的 LINK 字段代码里路径分隔符写成**双反斜杠**
            // （"C:\\Users\\..."，实测），而 ET 的 Workbook.FullName 是单斜杠。
            // 比较前必须把连续分隔符折叠掉。
            if (!string.Equals(NormPath(info.WorkbookPath), NormPath(t.WorkbookPath),
                    StringComparison.OrdinalIgnoreCase))
            {
                why = "工作簿不一致：表格=[" + info.WorkbookPath + "] 剪贴板=[" + t.WorkbookPath + "]";
                return false;
            }

            // 校验 2：工作表必须是同一张。
            // LINK 形态工作表在引用前缀里；DDE 形态在工作簿路径所在主题里（[工作簿]工作表）。
            string want = (t.SheetPart ?? "").Trim().TrimEnd('!').Trim('\'');
            if (want.Length == 0 && t.Kind != "LINK")
            {
                string topic = t.PathOrTopic ?? "";
                int close = topic.LastIndexOf(']');
                if (close >= 0 && close < topic.Length - 1)
                    want = topic.Substring(close + 1).Trim().Trim('\'');
            }
            if (want.Length > 0 &&
                !string.Equals(info.SheetName ?? "", want, StringComparison.OrdinalIgnoreCase))
            {
                why = "表格当前工作表与剪贴板来源不是同一张表。";
                return false;
            }

            // 窗口候选可能来自其他视图：逐片在包围盒内、互不重叠，且外边界完全一致。
            if (info.Areas.Count < 2) { why = "来源没有多个子区域。"; return false; }
            int minRow = int.MaxValue, maxRow = 0, minCol = int.MaxValue, maxCol = 0;
            for (int i = 0; i < info.Areas.Count; i++)
            {
                LinkFieldCode.SourceArea a = info.Areas[i];
                if (a.R1 < t.R1 || a.R2 > t.R2 || a.C1 < t.C1 || a.C2 > t.C2)
                {
                    why = "表格选区存在超出剪贴板包围盒的子区域。";
                    return false;
                }
                minRow = Math.Min(minRow, a.R1); maxRow = Math.Max(maxRow, a.R2);
                minCol = Math.Min(minCol, a.C1); maxCol = Math.Max(maxCol, a.C2);
                for (int j = 0; j < i; j++)
                {
                    LinkFieldCode.SourceArea previous = info.Areas[j];
                    if (a.R1 <= previous.R2 && previous.R1 <= a.R2 && a.C1 <= previous.C2 && previous.C1 <= a.C2)
                    { why = "来源多区域包含重叠单元格。"; return false; }
                }
            }
            if (minRow != t.R1 || maxRow != t.R2 || minCol != t.C1 || maxCol != t.C2)
            { why = "来源多区域的外边界与剪贴板包围范围不一致。"; return false; }

            // 工作表前缀一律沿用剪贴板里的写法（带空格的表名 WPS 会加单引号），
            // 不采用 ET 的裸名，避免生成不合格式的引用。
            // 记法同理：ET 给的是 A1 记法，但 WPS 的 LINK 字段实测只认 R1C1
            // （写成 "Sheet1!H1" 能插入、能刷新，取到的却是空值），
            // 所以一律改回剪贴板解析出来的那套记法。
            string sp = t.SheetPart ?? "";
            for (int i = 0; i < info.Areas.Count; i++)
            {
                info.Areas[i].SheetPart = sp;
                info.Areas[i].A1Style = t.A1Style;
            }

            t.Areas = info.Areas;
            return true;
        }


        /// <summary>
        /// 路径比较用的规范化：统一分隔符并折叠连续分隔符（保留 UNC 的开头 \\）。
        /// 只用于比对，绝不写回 —— 生成的字段代码必须与手工粘贴链接逐字一致。
        /// </summary>
        internal static string NormPath(string p)
        {
            if (string.IsNullOrEmpty(p)) return "";
            string s = p.Trim().Replace('/', '\\');
            bool unc = s.StartsWith("\\\\");
            string body = unc ? s.Substring(2) : s;
            while (body.IndexOf("\\\\", StringComparison.Ordinal) >= 0)
                body = body.Replace("\\\\", "\\");
            return unc ? "\\\\" + body : body;
        }

        /// <summary>离线自检门面：给一段字段代码，返回行优先展开出的「工作表!单元格」序列。</summary>
        public static bool TryListSourceRefs(string fieldCode, out List<string> refs, out string error)
        {
            refs = new List<string>();
            bool multi = false;
            LinkFieldCode.SourceTemplate t = LinkFieldCode.ParseLinkCode(fieldCode, out multi, out error);
            if (t == null) return false;
            List<LinkSourceCell> cells = LinkFieldCode.ExpandSource(t);
            for (int i = 0; i < cells.Count; i++) refs.Add(cells[i].SheetPart + cells[i].Ref());
            return true;
        }

        /// <summary>去掉代码开头的 LINK/DDE/DDEAUTO 关键字（wdFieldLink 显式插入用）。</summary>
        private static string StripKindPrefix(string code)
        {
            string s = code.TrimStart();
            string[] kinds = { "DDEAUTO ", "DDE ", "LINK " };
            for (int i = 0; i < kinds.Length; i++)
            {
                if (s.Length >= kinds[i].Length &&
                    s.Substring(0, kinds[i].Length).Equals(kinds[i], StringComparison.OrdinalIgnoreCase))
                    return s.Substring(kinds[i].Length);
            }
            return s;
        }

        // ================= 辅助 =================

        private static LinkPasteStatus TargetStatus(TargetRectStatus status)
        {
            if (status == TargetRectStatus.NoCells) return LinkPasteStatus.SelectionEmpty;
            if (status == TargetRectStatus.NotRect) return LinkPasteStatus.SelectionNotRect;
            if (status == TargetRectStatus.OutOfTable || status == TargetRectStatus.MultiTable)
                return LinkPasteStatus.SelectionOutOfTable;
            return LinkPasteStatus.Error;
        }

        private static void Rollback(List<object> fields)
        {
            for (int i = fields.Count - 1; i >= 0; i--)
            {
                try
                {
                    dynamic f = fields[i];
                    f.Delete();
                }
                catch (Exception ex)
                {
                    Connect.Log("链接Excel：回滚字段失败：" + ex.Message);
                }
            }
        }

        private static void EndUndo(object undoRec)
        {
            if (undoRec == null) return;
            try { ((dynamic)undoRec).EndCustomRecord(); }
            catch { }
        }

        private static void AddFail(LinkPasteResult r, int index, string reason)
        {
            r.Failures.Add(string.Format("第 {0} 格：{1}", index, reason));
        }

        /// <summary>写入阶段的安全中止：携带面向用户的状态与消息。</summary>
        private sealed class PasteAbortException : Exception
        {
            public readonly LinkPasteStatus Status;
            public PasteAbortException(LinkPasteStatus status, string message)
                : base(message)
            {
                Status = status;
            }
        }
    }
}
