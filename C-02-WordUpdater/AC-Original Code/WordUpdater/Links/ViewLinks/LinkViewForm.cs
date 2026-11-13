using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace WordUpdater
{
    // =====================================================================
    // 「查看链接」窗口 —— 仿 Word「编辑链接」的用法，但只聚焦本工具关心的信息
    //
    //   上表：当前文档引用了【哪些文件】，一行一个源文件
    //         工作簿名称 / 引用的工作表 / 更新方式 / 链接数 / 状态 / 完整路径
    //   下表：选中某个源文件后，列出它名下的每一条链接
    //         序号 / 工作表 / 区域 / 更新方式 / 字段代码或对象来源
    //
    //   底部动作：
    //     跳转到该字段 —— 选中下表某行后，把文档光标定位到该链接的位置
    //     更新所有链接 —— 整篇文档一次性同步（不要求先选中内容）
    //     中断所有链接 —— 整篇文档一次性断开，链接变成静态内容
    //
    //   注意：本窗口必须是【非模态】的（Connect 用 Show() 打开并绑 Owner），
    //   不能改成 ShowDialog() —— 模态期间 WPS 处于"对话框打开"状态，
    //   「打开源文件」经系统外壳让 WPS 开文件会被它自己拒绝
    //   （"对话框打开时命令无法执行"）。非模态顺带允许边看清单边浏览文档。
    // =====================================================================

    internal class LinkViewForm : Form
    {
        private readonly object _app;
        private LinkViewResult _last;

        private Label lblSummary;
        private Label lblEmpty;
        private Label lblDetail;
        private DataGridView gridSrc;
        private DataGridView gridItem;
        private RoundedButton btnOpen;
        private RoundedButton btnCopy;
        private RoundedButton btnJump;
        private RoundedButton btnRefresh;
        private RoundedButton btnBreak;
        private RoundedButton btnUpdate;

        // 「跳转到该字段」的标黄提示：引擎把目标区域标黄后，由定时器在 2.5 秒
        // 后还原成原来的高亮值，形成"闪一下"的效果。定时器与「更新链接」共用
        // Connect 的 FlashSession（见 JumpToSelectedItem），本窗体不再自己持有。

        public LinkViewForm(object application)
        {
            _app = application;
            bool preview=DialogPreviewFactory.IsConstructing;
            BuildUi();
            Load += delegate { if(!preview)RefreshData(); };
            // 关窗时若还有未还原的标黄，立刻还原，别把黄色留在文档里
            FormClosed += delegate { if(!preview)Connect.ClearPendingFlash(); };
        }

        // ------------------------------ 界面 ------------------------------

        private void BuildUi()
        {
            Text = "查看链接";
            DialogSizing.Prepare(this, "word.links", "view", Ui.Blue);
            // 窗口样式与 Excel 链接助手/追踪/合并抵消对齐：不设 ShowInTaskbar、
            // MinimizeBox——tool window 样式（WS_EX_TOOLWINDOW）在 WPS 宿主线程会
            // 抢焦点/失激活，表现为宿主单元格无法输入、功能区子菜单点不开（R-25）。
            MaximizeBox = true;
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill; root.Padding = new Padding(16); root.ColumnCount = 1; root.RowCount = 3;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 112));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            UiCard header = new UiCard(); header.Dock = DockStyle.Fill;
            TableLayoutPanel heading = new TableLayoutPanel(); heading.Dock = DockStyle.Fill;
            heading.ColumnCount = 1; heading.RowCount = 3;
            heading.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            heading.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));
            heading.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Label title = Ui.Label("查看链接", 16f, true); title.ForeColor = Ui.Blue; heading.Controls.Add(title, 0, 0);
            lblSummary = Ui.Label("正在扫描文档…", 10f, false); heading.Controls.Add(lblSummary, 0, 1);
            ToolTip summaryTip = new ToolTip(); lblSummary.TextChanged += delegate { summaryTip.SetToolTip(lblSummary, lblSummary.Text); };
            Disposed += delegate { summaryTip.Dispose(); };
            Label hint = Ui.Label("选择源文件查看明细 · 双击明细可跳转并临时标黄", 9f, false);
            hint.ForeColor = Ui.InkDim; heading.Controls.Add(hint, 0, 2); header.Controls.Add(heading); root.Controls.Add(header, 0, 0);
            TableLayoutPanel lists = new TableLayoutPanel(); lists.Dock = DockStyle.Fill; lists.Margin = Padding.Empty;
            lists.ColumnCount = 1; lists.RowCount = 2; lists.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            lists.RowStyles.Add(new RowStyle(SizeType.Percent, 48)); lists.RowStyles.Add(new RowStyle(SizeType.Percent, 52));
            gridSrc = MakeGrid();
            gridSrc.Columns.Add(MakeCol(gridSrc, "工作簿名称", 150, 260));
            gridSrc.Columns.Add(MakeCol(gridSrc, "工作表", 140, 220));
            gridSrc.Columns.Add(MakeCol(gridSrc, "更新方式", 110));
            gridSrc.Columns.Add(MakeCol(gridSrc, "链接数", 70));
            gridSrc.Columns.Add(MakeCol(gridSrc, "状态", 90));
            gridSrc.Columns.Add(MakeCol(gridSrc, "文件所在路径", 300));
            gridSrc.Columns[gridSrc.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            gridSrc.SelectionChanged += delegate { FillItems(); SyncButtons(); };
            gridSrc.CellDoubleClick += delegate { btnOpen.PerformClick(); };
            gridSrc.CellToolTipTextNeeded += FullTextTip;
            Panel sourceBody = new Panel(); sourceBody.BackColor = Color.White;
            lblEmpty = Ui.Label("", 11f, false); lblEmpty.TextAlign = ContentAlignment.MiddleCenter;
            lblEmpty.ForeColor = Ui.InkDim; lblEmpty.Visible = false;
            sourceBody.Controls.Add(gridSrc); sourceBody.Controls.Add(lblEmpty);
            lists.Controls.Add(Ui.Card(sourceBody, "源文件"), 0, 0);
            gridItem = MakeGrid();
            gridItem.Columns.Add(MakeCol(gridItem, "序号", 60));
            gridItem.Columns.Add(MakeCol(gridItem, "工作表", 130, 220));
            gridItem.Columns.Add(MakeCol(gridItem, "区域", 130, 180));
            gridItem.Columns.Add(MakeCol(gridItem, "更新方式", 90));
            gridItem.Columns.Add(MakeCol(gridItem, "字段代码 / 对象来源", 380));
            gridItem.Columns[gridItem.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            gridItem.Columns[gridItem.Columns.Count - 1].DefaultCellStyle.Font = Ui.MonoFont(10f, Font);
            gridItem.SelectionChanged += delegate { SyncButtons(); };
            gridItem.CellDoubleClick += delegate { JumpToSelectedItem(); };
            gridItem.CellToolTipTextNeeded += FullTextTip;
            UiCard detail = new UiCard(); detail.Dock = DockStyle.Fill;
            TableLayoutPanel detailLayout = new TableLayoutPanel(); detailLayout.Dock = DockStyle.Fill;
            detailLayout.ColumnCount = 1; detailLayout.RowCount = 2;
            detailLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            detailLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34)); detailLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            lblDetail = Ui.Label("链接明细：选择上方源文件", 10f, true);
            detailLayout.Controls.Add(lblDetail, 0, 0); detailLayout.Controls.Add(gridItem, 0, 1);
            detail.Controls.Add(detailLayout); lists.Controls.Add(detail, 0, 1); root.Controls.Add(lists, 0, 1);
            FlowLayoutPanel bar = Ui.ActionBar();
            btnOpen = MakeButton("打开源文件", 128); btnOpen.Click += delegate { OpenSelectedSource(); }; bar.Controls.Add(btnOpen);
            btnCopy = MakeButton("复制路径", 104); btnCopy.Click += delegate { CopySelectedPath(); }; bar.Controls.Add(btnCopy);
            btnJump = MakeButton("跳转到该字段", 140); btnJump.Click += delegate { JumpToSelectedItem(); }; bar.Controls.Add(btnJump);
            btnRefresh = MakeButton("刷新列表", 104); btnRefresh.Click += delegate { RefreshData(); }; bar.Controls.Add(btnRefresh);
            Label separator = new Label(); separator.Size = new Size(12, 1); separator.Margin = Padding.Empty; bar.Controls.Add(separator);
            btnBreak = MakeButton("中断所有链接", 140); Ui.StyleDanger(btnBreak); btnBreak.Click += delegate { DoBreakAll(); }; bar.Controls.Add(btnBreak);
            btnUpdate = MakeButton("更新所有链接", 140); Ui.StylePrimary(btnUpdate); btnUpdate.Click += delegate { DoUpdateAll(); }; bar.Controls.Add(btnUpdate);
            root.Controls.Add(bar, 0, 2); Controls.Add(root);
            AcceptButton = null; CancelButton = null; KeyPreview = true;
            KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) { e.Handled = true; Close(); } };
        }

        private void FullTextTip(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            DataGridView grid = (DataGridView)sender;
            DataGridViewCell cell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string value = Convert.ToString(cell.Value);
            if (!string.IsNullOrEmpty(value) && TextRenderer.MeasureText(value, cell.InheritedStyle.Font).Width > cell.Size.Width - cell.InheritedStyle.Padding.Horizontal)
                e.ToolTipText = value;
        }

        private static DataGridView MakeGrid()
        {
            DataGridView g = new DataGridView();
            g.Dock = DockStyle.Fill;
            g.ReadOnly = true;
            g.AllowUserToAddRows = false;
            g.AllowUserToDeleteRows = false;
            g.AllowUserToOrderColumns = true;
            g.AllowUserToResizeRows = false;
            g.MultiSelect = false;
            g.RowHeadersVisible = false;
            g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;   // 列宽交给 AutoFitColumns
            g.BackgroundColor = Color.White;
            g.BorderStyle = BorderStyle.None;            // 外框交给外层 1px 细边 Panel
            g.EnableHeadersVisualStyles = false;         // ★ 否则自定义表头色不生效
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;  // 先锁再设高度
            g.ColumnHeadersHeight = 30;
            g.ColumnHeadersDefaultCellStyle.BackColor = Ui.GridHead;
            g.ColumnHeadersDefaultCellStyle.ForeColor = Ui.GridHeadInk;
            g.ColumnHeadersDefaultCellStyle.Font = Ui.MakeFont(9f, FontStyle.Bold, null);
            g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            g.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;   // 表头不换行（配合 AutoFitColumns 的表头量宽）
            g.RowTemplate.Height = 24;                   // ★ 必须在 Rows.Add 之前设
            g.DefaultCellStyle.BackColor = Color.White;
            g.DefaultCellStyle.ForeColor = Ui.Ink;
            g.DefaultCellStyle.Font = Ui.MakeFont(9f, FontStyle.Regular, null);
            g.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
            g.DefaultCellStyle.SelectionBackColor = Ui.SelectBack;
            g.DefaultCellStyle.SelectionForeColor = Ui.SelectInk;
            g.AlternatingRowsDefaultCellStyle.BackColor = Ui.ZebraBack;
            g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;  // 只留横线
            g.GridColor = Ui.GridLine;
            Ui.StyleGrid(g);
            return g;
        }

        /// <summary>
        /// 建列：表头按实际渲染字体量宽（TextRenderer 是 DPI 感知的），
        /// 只增不减；长文本列把封顶上限暂存在列 Tag（DataGridViewColumn 没有
        /// MaximumWidth），AutoFitColumns 贴合后再按它手动封顶。
        /// </summary>
        private static DataGridViewTextBoxColumn MakeCol(DataGridView grid, string header, int width, int maxWidth = 0)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = header;
            c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            c.MinimumWidth = 50;
            c.SortMode = DataGridViewColumnSortMode.NotSortable;

            Font hf = grid.ColumnHeadersDefaultCellStyle.Font ?? grid.Font;
            int need = TextRenderer.MeasureText(header, hf).Width + Ui.Px(grid, 32);   // 24 = 左右内边距 + 分隔线 + 余量
            c.Width = Math.Max(width, need);

            if (maxWidth > 0) c.Tag = Math.Max(maxWidth, need);
            return c;
        }

        /// <summary>
        /// 数据填充后对每列「贴合最宽内容」（等价双击列宽线 = AutoResize(AllCells)）。
        /// 跳过末列（Fill 列自行补满窗口）；长文本列贴合后再按列 Tag 里的上限手动封顶。
        /// </summary>
        private static void AutoFitColumns(DataGridView grid)
        { Ui.AutoFitColumns(grid); }

        private RoundedButton MakeButton(string text, int width)
        { return Ui.Button(text, width); }

        // ------------------------------ 数据 ------------------------------

        private void RefreshData()
        {
            Cursor cur = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                LinkViewResult r = LinkViewEngine.Scan(_app);
                _last = r;

                gridSrc.Rows.Clear();
                gridItem.Rows.Clear();

                string docName = string.IsNullOrEmpty(r.DocumentName) ? "(未命名文档)" : r.DocumentName;

                if (r.Status == LinkViewStatus.NoDocument)
                {
                    ShowEmpty("当前没有打开的文档。\r\n\r\n请在 WPS 文字中打开要检查的文档后，再点击「查看链接」。");
                    lblSummary.Text = "未检测到打开的文档";
                    return;
                }
                if (r.Status == LinkViewStatus.Error)
                {
                    ShowEmpty("扫描链接时发生错误：\r\n\r\n" + (r.ErrorMessage ?? "未知错误"));
                    lblSummary.Text = "扫描失败";
                    return;
                }
                if (r.Status == LinkViewStatus.NoLinks)
                {
                    ShowEmpty("当前文档中没有发现任何链接。\r\n\r\n"
                        + "提示：先用「链接Excel」或在 WPS 里「选择性粘贴 → 粘贴链接」，"
                        + "才能在这里看到被引用的文件。");
                    lblSummary.Text = "文档：" + docName + " ｜ 未发现链接";
                    return;
                }

                lblEmpty.Visible = false;
                gridSrc.Visible = true;
                lblSummary.Text = string.Format(
                    "文档：{0} ｜ 引用外部文件 {1} 个 ｜ 链接总数 {2} 条",
                    docName, r.Sources.Count, r.TotalLinks);

                for (int i = 0; i < r.Sources.Count; i++)
                {
                    LinkSourceInfo si = r.Sources[i];
                    int row = gridSrc.Rows.Add(
                        si.WorkbookName, si.SheetText, si.UpdateText,
                        si.LinkCount.ToString(), si.StatusText, si.FullPath);
                    gridSrc.Rows[row].Tag = si;
                    if (!si.FileExists && si.PathIsLocal)
                    {
                        gridSrc.Rows[row].DefaultCellStyle.ForeColor = Ui.Danger;
                    }
                }
                if (gridSrc.Rows.Count > 0)
                {
                    // 时序坑：第一行被 DataGridView 自动选中发生在 Tag 赋值之前，
                    // 那次 SelectionChanged 里 FillItems 读到的是 null；
                    // 之后 Selected=true（状态没变）和"点击已选中的行"都不会再
                    // 触发 SelectionChanged —— 只有一个源文件时明细会永远空白。
                    // 所以填充完必须显式同步一次，不能只依赖事件。
                    gridSrc.Rows[0].Selected = true;
                }
                FillItems();
                AutoFitColumns(gridSrc);
            }
            finally
            {
                Cursor.Current = cur;
            }
            SyncButtons();
        }

        private void ShowEmpty(string message)
        {
            lblEmpty.Text = message;
            lblEmpty.Visible = true;
            lblEmpty.BringToFront();
            gridSrc.Visible = false;
            FillItems();
        }

        private void FillItems()
        {
            gridItem.Rows.Clear();
            LinkSourceInfo si = SelectedSource();
            if (si == null)
            {
                lblDetail.Text = "链接明细：选中上方任一源文件，查看它名下每一条链接";
                return;
            }
            lblDetail.Text = string.Format("链接明细：{0} 名下共 {1} 条链接",
                si.WorkbookName, si.LinkCount);

            for (int i = 0; i < si.Items.Count; i++)
            {
                LinkItemInfo it = si.Items[i];
                string sheet = string.IsNullOrEmpty(it.SheetName) ? "—" : it.SheetName;
                string detail = it.ItemKind.IndexOf("字段") >= 0
                    ? it.KindCode + "  " + it.RawCode
                    : it.ItemKind + "（OLE 链接）";
                int row = gridItem.Rows.Add(
                    it.Index.ToString(), sheet, it.RangeText, it.UpdateMode, detail);
                gridItem.Rows[row].Tag = it;   // 「跳转到该字段」用
            }
            AutoFitColumns(gridItem);
        }

        private LinkSourceInfo SelectedSource()
        {
            try
            {
                if (gridSrc.SelectedRows.Count == 0) return null;
                return gridSrc.SelectedRows[0].Tag as LinkSourceInfo;
            }
            catch { return null; }
        }

        private LinkItemInfo SelectedItem()
        {
            try
            {
                if (gridItem.SelectedRows.Count == 0) return null;
                return gridItem.SelectedRows[0].Tag as LinkItemInfo;
            }
            catch { return null; }
        }

        private void SyncButtons()
        {
            LinkSourceInfo si = SelectedSource();
            bool has = si != null;
            // 只有"看起来像本地/网络路径"的来源才能打开、复制；
            // DDE 裸主题（只有书名）之类没有可打开的路径。
            bool openable = has && si.PathIsLocal && !string.IsNullOrEmpty(si.FullPath);
            btnOpen.Enabled = openable;
            btnCopy.Enabled = openable;
            btnJump.Enabled = SelectedItem() != null;
            if (_last != null && (_last.Status != LinkViewStatus.Completed || _last.TotalLinks == 0))
            {
                btnUpdate.Enabled = false;
                btnBreak.Enabled = false;
            }
            else
            {
                btnUpdate.Enabled = true;
                btnBreak.Enabled = true;
            }
        }

        // ------------------------------ 动作 ------------------------------

        private void OpenSelectedSource()
        {
            LinkSourceInfo si = SelectedSource();
            if (si == null || string.IsNullOrEmpty(si.FullPath)) return;
            try
            {
                if (!si.PathIsLocal)
                {
                    UiDialog.Show(this,
                        "这条链接的来源不是本地文件路径（例如 DDE 裸主题或网络地址）：\r\n\r\n"
                        + si.FullPath + "\r\n\r\n无法直接打开，可先用「复制路径」手动处理。",
                        "查看链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!si.FileExists)
                {
                    DialogResult dr = UiDialog.Show(this,
                        "源文件已不在原位置：\r\n\r\n" + si.FullPath
                        + "\r\n\r\n要打开它所在的文件夹看看吗？",
                        "查看链接", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                        MessageBoxDefaultButton.Button1);
                    if (dr == DialogResult.Yes) LocateInExplorer(si.FullPath);
                    return;
                }
                Process.Start(si.FullPath);
            }
            catch (Exception ex)
            {
                // 兜底：系统打不开（关联异常/被安全软件拦截等），退而求其次定位文件
                DialogResult dr = UiDialog.Show(this,
                    "无法直接打开源文件：\r\n\r\n" + ex.Message
                    + "\r\n\r\n要改为在资源管理器中定位该文件吗？",
                    "查看链接", MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button1);
                if (dr == DialogResult.Yes) LocateInExplorer(si.FullPath);
            }
        }

        /// <summary>在资源管理器里定位文件（不依赖任何文件关联，永远不会被"对话框打开"挡住）。</summary>
        private static void LocateInExplorer(string path)
        {
            try
            {
                Process.Start("explorer.exe", "/select,\"" + path + "\"");
            }
            catch (Exception ex)
            {
                UiDialog.Show("定位失败：\r\n\r\n" + ex.Message,
                    "查看链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CopySelectedPath()
        {
            LinkSourceInfo si = SelectedSource();
            if (si == null || string.IsNullOrEmpty(si.FullPath)) return;
            try { Clipboard.SetText(si.FullPath); }
            catch (Exception ex)
            {
                UiDialog.Show(this, "复制失败：\r\n\r\n" + ex.Message,
                    "查看链接", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>把文档光标定位到明细表选中那条链接的位置（窗口保持开着）。</summary>
        private void JumpToSelectedItem()
        {
            if(DialogPreviewFactory.IsPreview(this))return;
            LinkItemInfo it = SelectedItem();
            if (it == null) return;
            // Restore the previous flash before the engine reads original color
            // and paints the next target. The same field can return another Range
            // wrapper; clearing afterward would erase new yellow and retain 7.
            Connect.ClearPendingFlash();
            object flash; int prev;
            string err = LinkViewEngine.JumpToItemEx(_app, it, out flash, out prev);
            if (err != null)
            {
                UiDialog.Show(this, err, "跳转到该字段",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // Target is now yellow with its real original color recorded. Reuse
            // the update command's shared session and restore after the timer.
            if (flash != null)
            {
                Connect.SharedFlash.Add(flash, prev);
                Connect.SharedFlash.Restart(2500);
            }
        }

        private void DoUpdateAll()
        {
            if (_last == null || _last.TotalLinks == 0) return;
            DialogResult dr = UiDialog.Show(this,
                string.Format("即将更新本文档中的全部 {0} 条链接。\r\n\r\n"
                    + "这会用源文件的当前值刷新文档内容，已脱离源文件的链接会报错。"
                    + "建议先保存文档，以便需要时可撤销。\r\n\r\n确定更新吗？",
                    _last.TotalLinks),
                "更新所有链接", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dr != DialogResult.Yes) return;

            RunAction(true);
        }

        private void DoBreakAll()
        {
            if (_last == null || _last.TotalLinks == 0) return;
            DialogResult dr = UiDialog.Show(this,
                string.Format("即将中断本文档中的全部 {0} 条链接。\r\n\r\n"
                    + "中断后链接会变成文档里的静态内容与图片，不再随源文件变化，"
                    + "且无法用本工具恢复（只能重新「链接Excel」）。\r\n\r\n"
                    + "确定要中断吗？",
                    _last.TotalLinks),
                "中断所有链接", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (dr != DialogResult.Yes) return;

            DialogResult dr2 = UiDialog.Show(this,
                "再确认一次：中断操作不可逆。\r\n\r\n仍要中断全部链接吗？",
                "中断所有链接", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (dr2 != DialogResult.Yes) return;

            RunAction(false);
        }

        private void RunAction(bool update)
        {
            Cursor cur = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            LinkActionResult r;
            try
            {
                r = update ? LinkViewEngine.UpdateAll(_app) : LinkViewEngine.BreakAll(_app);
            }
            catch (Exception ex)
            {
                Cursor.Current = cur;
                UiDialog.Show(this, "操作未能完成：\r\n\r\n" + ex.Message,
                    "查看链接", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Cursor.Current = cur;

            Connect.Log(string.Format("查看链接 -> {0} total={1} ok={2} fail={3}",
                update ? "更新所有链接" : "中断所有链接", r.Total, r.Success, r.Failed));

            UiDialog.Show(this, BuildActionMessage(r, update),
                update ? "更新所有链接" : "中断所有链接",
                MessageBoxButtons.OK,
                (r.Executed && r.Failed == 0) ? MessageBoxIcon.Information : MessageBoxIcon.Warning);

            RefreshData();
        }

        private static string BuildActionMessage(LinkActionResult r, bool update)
        {
            if (!r.Executed)
            {
                return (update ? "更新" : "中断") + "未完成，发生引擎级错误：\r\n\r\n"
                    + (r.ErrorMessage ?? "未知错误");
            }
            if (r.Total == 0) return "文档里已经没有任何链接了。";

            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0}完成：共 {1} 条链接，成功 {2} 条",
                update ? "更新" : "中断", r.Total, r.Success);
            if (r.Failed > 0) sb.AppendFormat("，失败 {0} 条。", r.Failed);
            else sb.Append("。");

            if (r.Failed > 0)
            {
                sb.Append("\r\n\r\n失败明细：\r\n");
                int max = Math.Min(r.Failures.Count, 10);
                for (int i = 0; i < max; i++) sb.Append("  " + r.Failures[i] + "\r\n");
                if (r.Failures.Count > max)
                    sb.AppendFormat("  …… 其余 {0} 条略。", r.Failures.Count - max);
                sb.Append("\r\n提示：源文件被移动或删除时会出现这种情况，"
                    + "可在列表里查看标记为「源文件缺失」的条目。");
            }

            if (update && r.CleanedHiddenHosts > 0)
            {
                sb.AppendFormat("\r\n\r\n（已自动释放 {0} 个更新时拉起的后台 WPS 表格进程。）",
                    r.CleanedHiddenHosts);
            }
            if (!update && r.Success > 0)
            {
                sb.Append("\r\n\r\n文档内容已保留为静态结果；如需恢复自动同步，"
                    + "请重新使用「链接Excel」。");
            }
            return sb.ToString();
        }
    }
}
