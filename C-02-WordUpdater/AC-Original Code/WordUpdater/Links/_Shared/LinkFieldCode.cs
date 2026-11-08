using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordUpdater
{
    internal static class LinkFieldCode
    {
        // ---- 链接字段代码解析 ----
        // LINK  <类名> "<工作簿路径>" (<"Sheet!R1C1:R4C4">|<裸引用>) <尾部开关>
        private static readonly Regex RxLinkQuoted =
            new Regex(@"^\s*LINK\s+(?<cls>\S+)\s+""(?<path>[^""]+)""\s+(?:""(?<itemq>[^""]*)""|(?<item>\S+))(?<tail>.*)$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex RxLinkBare =
            new Regex(@"^\s*LINK\s+(?<cls>\S+)\s+(?<path>\S+)\s+(?:""(?<itemq>[^""]*)""|(?<item>\S+))(?<tail>.*)$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        // DDE / DDEAUTO  <服务名> (<"[工作簿]工作表">|<裸主题>) (<"R1C1:R4C4">|<裸引用>) <尾部>
        private static readonly Regex RxDde =
            new Regex(@"^\s*(?<kind>DDE|DDEAUTO)\s+(?<app>\S+)\s+(?:""(?<topicq>[^""]*)""|(?<topic>\S+))\s+(?:""(?<itemq>[^""]*)""|(?<item>\S+))(?<tail>.*)$",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        // 引用项：工作表前缀 + 区域（R1C1 或 A1，含单格与区域两种）
        private static readonly Regex RxRangeR1C1 =
            new Regex(@"^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$", RegexOptions.Compiled);
        private static readonly Regex RxRangeA1 =
            new Regex(@"^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$", RegexOptions.Compiled);
        // 单个引用片段（多区域时按 , / ; 拆开后的每一片）：
        // 工作表前缀可选，缺省则继承上一片的前缀 —— 因此 "Sheet1!R1C1,R3C1" 也能拆对
        private static readonly Regex RxRefPiece =
            new Regex(@"^(?:(?<sheet>[^!]+)!)?(?<range>R\d+C\d+(?::R\d+C\d+)?|\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$",
                RegexOptions.Compiled);

        internal class SourceTemplate
        {
            public string Kind;         // LINK / DDE / DDEAUTO
            public string Head;         // LINK：类名（Excel.Sheet.12…）；DDE：服务名（Excel…）
            public string PathOrTopic;  // LINK：工作簿路径；DDE：主题（[工作簿]工作表）
            public bool PathQuoted;     // 路径/主题在代码中是否带引号
            public string Switches;     // 尾部开关原样保留（含前导空格，可为空串）
            public bool ItemQuoted;     // 引用项在代码中是否整体带引号
            public string SheetPart;    // LINK：工作表前缀（如 "Sheet1!" / "'My Sheet'!"）；DDE：空
            public int R1, C1, R2, C2;  // 主区域（多区域时 = 第 1 区）；1 起始、含端点
            public bool A1Style;        // 引用记法：true=A1，false=R1C1
            public string WorkbookPath; // 尽力提取的本地路径（可能为 null）

            /// <summary>引用项拆出的子区域，按出现顺序（Ctrl 多选即多片）。</summary>
            public List<SourceArea> Areas = new List<SourceArea>();

            public bool MultiArea { get { return Areas.Count > 1; } }

            /// <summary>全部子区域的单元格总数（已按行优先展开后的长度）。</summary>
            public long RawCount
            {
                get
                {
                    long n = 0;
                    for (int i = 0; i < Areas.Count; i++) n += Areas[i].Count;
                    return n;
                }
            }

            public string Describe()
            {
                // 多区域时逐片列出（最多 3 片），单区域与改造前逐字一致
                string rng = MultiArea ? DescribeAreas(this) : DescribeRange(this);
                if (Kind == "LINK")
                    return PathOrTopic + " → " + SheetPart + rng;
                return Kind + " " + Head + " " + PathOrTopic + " → " + rng;
            }
        }
        internal sealed class SourceArea
        {
            public string SheetPart = "";   // "Sheet1!"；DDE 形态为空串
            public int R1, C1, R2, C2;      // 已归一化：R2>=R1、C2>=C1
            public bool A1Style;

            public long Count
            {
                get { return (long)(R2 - R1 + 1) * (long)(C2 - C1 + 1); }
            }
        }

        // ================= 链接字段代码解析 =================

        /// <summary>已解析的链接源模板：按格重建字段代码所需的全部要素。</summary>
        /// 注意：同为 internal —— LinkViewEngine（查看链接）也要靠它解析已存在的链接。

        /// <summary>引用项中的一片连续区域（Ctrl 多选时会有多片）。</summary>

        internal static string DescribeRange(SourceTemplate t)
        {
            string a = CellRef(t, t.R1, t.C1);
            string b = CellRef(t, t.R2, t.C2);
            if (a == b) return a;
            return a + ":" + b;
        }

        internal static SourceTemplate ParseLinkCode(string code, out bool multiArea, out string error)
        {
            multiArea = false;
            error = null;
            string s = (code ?? "").Trim();

            // 罕见前缀 LOCK
            if (s.Length >= 5 && s.Substring(0, 5).Equals("LOCK ", StringComparison.OrdinalIgnoreCase))
                s = s.Substring(5).TrimStart();
            if (s.Length == 0)
            {
                error = "链接字段代码为空。";
                return null;
            }

            string kind = null, head = null, pathOrTopic = null, itemText = null, tail = null;
            bool pathQuoted = false, itemQuoted = false;

            Match m = RxLinkQuoted.Match(s);
            bool viaQuotedRx = m.Success;
            if (!m.Success) m = RxLinkBare.Match(s);
            if (m.Success)
            {
                kind = "LINK";
                head = m.Groups["cls"].Value;
                pathOrTopic = m.Groups["path"].Value;
                pathQuoted = viaQuotedRx;
                itemText = m.Groups["itemq"].Success ? m.Groups["itemq"].Value : m.Groups["item"].Value;
                itemQuoted = m.Groups["itemq"].Success;
                tail = m.Groups["tail"].Value;
            }
            else
            {
                m = RxDde.Match(s);
                if (m.Success)
                {
                    kind = m.Groups["kind"].Value.ToUpperInvariant();
                    head = m.Groups["app"].Value;
                    pathOrTopic = m.Groups["topicq"].Success ? m.Groups["topicq"].Value : m.Groups["topic"].Value;
                    pathQuoted = m.Groups["topicq"].Success;
                    itemText = m.Groups["itemq"].Success ? m.Groups["itemq"].Value : m.Groups["item"].Value;
                    itemQuoted = m.Groups["itemq"].Success;
                    tail = m.Groups["tail"].Value;
                }
            }

            if (kind == null)
            {
                error = "无法识别的链接字段格式：" + Trunc(s, 120);
                return null;
            }
            if (string.IsNullOrEmpty(itemText))
            {
                error = "链接字段缺少区域引用。";
                return null;
            }

            SourceTemplate t = new SourceTemplate();
            t.Kind = kind;
            t.Head = head;
            t.PathOrTopic = pathOrTopic;
            t.PathQuoted = pathQuoted;
            t.Switches = tail ?? "";
            t.ItemQuoted = itemQuoted;

            if (kind == "LINK")
            {
                t.WorkbookPath = pathOrTopic;
                t.SheetPart = "";
            }
            else
            {
                t.SheetPart = "";
                // DDE 主题形如 [工作簿]工作表；工作簿可能是完整路径或仅书名
                string topic = pathOrTopic ?? "";
                if (topic.StartsWith("["))
                {
                    int close = topic.LastIndexOf(']');
                    if (close > 0)
                    {
                        string p = topic.Substring(1, close - 1);
                        t.WorkbookPath = LooksLocal(p) ? p : null;
                    }
                }
            }

            // ---------- 引用项 → 子区域列表 ----------
            // Ctrl 多选时引用形如 "Sheet1!R1C1,R3C1"，按 , / ; 拆成多片；
            // 只有一片时拆出的结果与改造前完全一致。
            if (!TrySplitItem(itemText, t.Areas, out error)) return null;
            if (t.Areas.Count == 0)
            {
                error = "链接字段缺少有效的区域引用。";
                return null;
            }
            multiArea = t.Areas.Count > 1;

            // 兼容既有语义：R1..C2 = 主区域（第 1 片），供 DescribeRange / 查看链接 展示
            SourceArea a0 = t.Areas[0];
            t.R1 = a0.R1; t.C1 = a0.C1; t.R2 = a0.R2; t.C2 = a0.C2;
            t.A1Style = a0.A1Style;
            t.SheetPart = a0.SheetPart ?? "";
            return t;
        }

        /// <summary>
        /// 把引用项按 , / ; 拆成子区域。每片的工作表前缀可省略，缺省继承上一片
        /// （"Sheet1!R1C1,R3C1" 的第二片仍属于 Sheet1）；DDE 形态整体无前缀则为 ""。
        /// </summary>
        internal static bool TrySplitItem(string itemText, List<SourceArea> areas, out string error)
        {
            error = null;
            areas.Clear();
            string[] pieces = (itemText ?? "").Split(new char[] { ',', ';' });
            string inherit = null;

            for (int i = 0; i < pieces.Length; i++)
            {
                string p = pieces[i].Trim();
                if (p.Length == 0) continue;

                Match im = RxRefPiece.Match(p);
                if (!im.Success)
                {
                    error = "无法识别的区域片段：" + Trunc(p, 40);
                    return false;
                }

                string sheet;
                if (im.Groups["sheet"].Success)
                {
                    sheet = im.Groups["sheet"].Value + "!";
                    inherit = sheet;
                }
                else
                {
                    sheet = inherit != null ? inherit : "";
                }
                string range = im.Groups["range"].Value;

                SourceArea a = new SourceArea();
                a.SheetPart = sheet;
                Match rm = RxRangeR1C1.Match(range);
                if (rm.Success)
                {
                    a.A1Style = false;
                    a.R1 = int.Parse(rm.Groups[1].Value);
                    a.C1 = int.Parse(rm.Groups[2].Value);
                    a.R2 = rm.Groups[3].Success ? int.Parse(rm.Groups[3].Value) : a.R1;
                    a.C2 = rm.Groups[4].Success ? int.Parse(rm.Groups[4].Value) : a.C1;
                }
                else
                {
                    rm = RxRangeA1.Match(range);
                    if (!rm.Success)
                    {
                        error = "无法解析源区域引用：" + Trunc(range, 60);
                        return false;
                    }
                    a.A1Style = true;
                    a.C1 = CellAddress.ColumnNumber(rm.Groups[1].Value, int.MaxValue);
                    a.R1 = int.Parse(rm.Groups[2].Value);
                    a.C2 = rm.Groups[3].Success ? CellAddress.ColumnNumber(rm.Groups[3].Value, int.MaxValue) : a.C1;
                    a.R2 = rm.Groups[4].Success ? int.Parse(rm.Groups[4].Value) : a.R1;
                }

                // 归一化：右下角不小于左上角
                if (a.R2 < a.R1) { int tmp = a.R1; a.R1 = a.R2; a.R2 = tmp; }
                if (a.C2 < a.C1) { int tmp = a.C1; a.C1 = a.C2; a.C2 = tmp; }
                if (a.R1 < 1 || a.C1 < 1 || a.R1 > 1048576 || a.C1 > 16384)
                {
                    error = "源区域引用越界：" + Trunc(range, 60);
                    return false;
                }
                areas.Add(a);
            }

            if (areas.Count == 0)
            {
                error = "链接字段缺少区域引用。";
                return false;
            }
            return true;
        }

        /// <summary>
        /// 把源模板展成一维单元格序列：子区域按出现顺序，区域内先行后列（行优先）。
        /// 这是「源 ↔ 目标」按位对应的基础。
        /// </summary>
        internal static List<LinkSourceCell> ExpandSource(SourceTemplate t)
        {
            List<LinkSourceCell> list = new List<LinkSourceCell>();
            if (t == null) return list;
            for (int i = 0; i < t.Areas.Count; i++)
            {
                SourceArea a = t.Areas[i];
                for (int r = a.R1; r <= a.R2; r++)
                {
                    for (int c = a.C1; c <= a.C2; c++)
                    {
                        LinkSourceCell cell = new LinkSourceCell();
                        cell.SheetPart = a.SheetPart;
                        cell.Row = r;
                        cell.Col = c;
                        cell.A1Style = a.A1Style;
                        list.Add(cell);
                    }
                }
            }
            return list;
        }

        /// <summary>多区域的整体描述（最多列 3 区，超出显示「…等 N 区」）。</summary>
        internal static string DescribeAreas(SourceTemplate t)
        {
            if (t == null || t.Areas.Count == 0) return "";
            string s = "";
            int max = Math.Min(t.Areas.Count, 3);
            for (int i = 0; i < max; i++)
            {
                if (i > 0) s += ",";
                s += AreaText(t.Areas[i]);
            }
            if (t.Areas.Count > max) s += "…等 " + t.Areas.Count + " 区";
            return s;
        }

        private static string AreaText(SourceArea a)
        {
            string p, q;
            if (a.A1Style)
            {
                p = CellAddress.ColumnLetters(a.C1) + a.R1.ToString();
                q = CellAddress.ColumnLetters(a.C2) + a.R2.ToString();
            }
            else
            {
                p = "R" + a.R1.ToString() + "C" + a.C1.ToString();
                q = "R" + a.R2.ToString() + "C" + a.C2.ToString();
            }
            return p == q ? p : p + ":" + q;
        }

        // ================= 按格重建字段代码 =================

        private static string CellRef(SourceTemplate t, int row, int col)
        {
            if (t.A1Style) return CellAddress.ColumnLetters(col) + row.ToString();
            return "R" + row.ToString() + "C" + col.ToString();
        }

        /// <summary>以源模板为底，为序列中某个已展开的源单元格生成链接字段代码。
        /// 与上面 (row,col) 重载逐字同构，只是引用可携带该区域自己的工作表前缀。</summary>
        internal static string BuildCode(SourceTemplate t, LinkSourceCell cell)
        {
            string item = (cell.SheetPart ?? "") + cell.Ref();
            if (t.ItemQuoted) item = "\"" + item + "\"";
            string head = t.PathQuoted ? "\"" + t.PathOrTopic + "\"" : t.PathOrTopic;
            string code = t.Kind + " " + t.Head + " " + head + " " + item + t.Switches;
            return code.TrimEnd();
        }

        internal static bool LooksLocal(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            if (path.Length >= 3 && path[1] == ':' && (path[2] == '\\' || path[2] == '/')) return true;
            if (path.StartsWith("\\\\")) return true;
            return false;
        }

        private static string Trunc(string s, int max)
        {
            if (s == null) return "";
            return s.Length <= max ? s : s.Substring(0, max) + "…";
        }
    }
}
