using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    // Only the controller runs on the host STA. Isolated dialogs receive
    // strings/JSON and return strings/JSON, never a host Form or COM object.
    internal sealed class BatchPresetActions
    {
        internal readonly RoundedButton Gear = Ui.Button("⚙", 44);
        private readonly Form owner;
        private readonly string feature;
        private readonly Func<BatchPreset> capture;
        private readonly Action<BatchPreset> apply;
        private readonly Func<bool> available;
        private readonly Action<string> message;
        private readonly ContextMenuStrip menu = new ContextMenuStrip();
        private string savedState;
        internal BatchPresetActions(Form owner, string feature, Func<BatchPreset> capture, Action<BatchPreset> apply, Func<bool> available, Action<string> message)
        {
            this.owner = owner; this.feature = feature; this.capture = capture; this.apply = apply; this.available = available; this.message = message;
            Gear.Name = "batchPresetSettings"; Gear.AccessibleName = "预设设置"; Gear.AccessibleDescription = "保存为预设或打开预设管理";
            var gearFont = Ui.MakeFont(15, FontStyle.Regular, Gear.Font); var menuFont = Ui.MakeFont(10, FontStyle.Regular, owner.Font);
            Gear.Font = gearFont; Gear.Margin = new Padding(8, 0, 0, 0);
            var tooltip = new ToolTip(); tooltip.SetToolTip(Gear, "保存为预设 / 预设管理");
            menu.ShowImageMargin = false; menu.BackColor = Color.White; menu.Font = menuFont;
            menu.Items.Add("保存为预设…", null, delegate { owner.BeginInvoke((MethodInvoker)delegate { Run(delegate { Save(); }); }); });
            menu.Items.Add("预设管理…", null, delegate { owner.BeginInvoke((MethodInvoker)delegate { Run(Manage); }); });
            Gear.Click += delegate { if (available()) menu.Show(Gear, new Point(0, Gear.Height)); };
            owner.Disposed += delegate { menu.Dispose(); tooltip.Dispose(); gearFont.Dispose(); menuFont.Dispose(); };
        }
        internal Control Header(Control content)
        {
            var row = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); row.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var flow = content as FlowLayoutPanel; if (flow != null) flow.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            content.Dock = DockStyle.Top; row.Controls.Add(content, 0, 0); row.Controls.Add(Gear, 1, 0); Gear.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            bool fitting = false;
            row.Layout += delegate
            {
                if (fitting) return; fitting = true;
                try
                {
                    int width = Math.Max(100, row.ClientSize.Width - Gear.Width - Gear.Margin.Horizontal - content.Margin.Horizontal);
                    content.MaximumSize = new Size(width, 0);
                    int height = Math.Max(content.GetPreferredSize(new Size(width, 0)).Height + content.Margin.Vertical, Gear.Height + Gear.Margin.Vertical);
                    row.RowStyles[0].SizeType = SizeType.Absolute; row.RowStyles[0].Height = height;
                }
                finally { fitting = false; }
            };
            return row;
        }
        private void Run(Action action)
        {
            if (owner.IsDisposed) return;
            if (!available()) { message("任务执行期间或有待还原结果时不能操作预设。"); return; }
            try { action(); } catch (Exception error) { if (!owner.IsDisposed) message("预设操作失败：" + error.Message); }
        }
        private static string Fingerprint(BatchPreset value)
        { return (value.Context ?? "") + "\n" + string.Join("\n", value.Values.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Key + "=" + p.Value).ToArray()); }
        internal bool Save()
        {
            owner.ValidateChildren(); var value = capture();
            if (string.IsNullOrEmpty(value.Context)) { message("请先选择模板或工作簿并填写配置。"); return false; }
            string suggested = Path.GetFileNameWithoutExtension(value.Context);
            string target; if (value.Values.TryGetValue("target", out target) && target.Length > 0) suggested += " · " + target;
            suggested = suggested.Substring(0, Math.Min(70, suggested.Length));
            string kind = feature;
            string name = HostDialogService.ShowIsolatedResult<string>(owner,
                delegate { return new BatchPresetNameForm(kind, "保存为预设", suggested); },
                delegate(Form dialog) { return ((BatchPresetNameForm)dialog).ChosenName; }, null);
            if (name == null) return false;
            string file = BatchPresetStore.FileFor(feature, name); bool exists = File.Exists(file);
            if (exists && UiDialog.Show(owner, "已有同名预设“" + name + "”，是否覆盖？", "保存为预设", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return false;
            value.Name = name; value.Saved = DateTime.Now.ToString("s", CultureInfo.InvariantCulture);
            BatchPresetStore.Write(file, value, exists); savedState = Fingerprint(value); message("已保存预设“" + name + "”。"); return true;
        }
        internal void Manage()
        {
            string kind = feature;
            string json = HostDialogService.ShowIsolatedResult<string>(owner,
                delegate { return new BatchPresetManagerForm(kind); },
                delegate(Form dialog) { var chosen = ((BatchPresetManagerForm)dialog).Chosen; return chosen == null ? null : BatchPresetStore.Encode(chosen); }, null);
            if (json == null) return;
            var value = BatchPresetStore.Decode(json, feature); var current = capture();
            if (!string.IsNullOrEmpty(current.Context) && Fingerprint(current) != savedState)
            {
                var answer = UiDialog.Show(owner, "应用预设“" + value.Name + "”会替换当前配置。是否先保存当前设置？", "应用预设", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (answer == DialogResult.Cancel || (answer == DialogResult.Yes && !Save())) return;
            }
            if (!ResolveFiles(value)) return;
            if (!available()) { message("当前状态已变化，预设未应用。"); return; }
            apply(value); if (owner.IsDisposed) return; savedState = Fingerprint(value);
            message("已应用预设“" + value.Name + "”。仅恢复设置，请检查并重新预检；未执行生成或写入审计数据。");
        }
        private bool ResolveFiles(BatchPreset value)
        {
            if (string.IsNullOrEmpty(value.Context) || !File.Exists(value.Context))
            {
                if (UiDialog.Show(owner, "找不到预设中的文件：\n" + value.Context + "\n是否选择替代文件？取消将保留当前设置。", "重新定位文件", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return false;
                string filter = feature == "company" ? "模板 (*.xlsx;*.xls)|*.xlsx;*.xls" : "表格工作簿 (*.xlsx;*.xls;*.xlsm;*.et)|*.xlsx;*.xls;*.xlsm;*.et";
                string[] paths = HostDialogService.OpenFiles(owner, filter);
                if (paths.Length == 0) return false;
                if (paths.Length != 1) throw new InvalidOperationException("请只选择一个替代文件。");
                value.Context = paths[0];
            }
            string output;
            if (feature == "company" && value.Values.TryGetValue("output", out output) && output.Length > 0 && !Directory.Exists(output))
            {
                string chosen = HostDialogService.Folder(owner, "预设输出文件夹不存在，请重新选择", true);
                if (string.IsNullOrEmpty(chosen)) return false; value.Values["output"] = chosen;
            }
            return true;
        }
    }

    internal sealed class BatchPresetNameForm : Form
    {
        private readonly TextBox name = new TextBox { Dock = DockStyle.Fill, Name = "batchPresetName" };
        internal string ChosenName { get { return name.Text.Trim(); } }
        internal BatchPresetNameForm(string feature, string title, string initial)
        {
            Text = title; DialogSizing.Prepare(this, "batch.preset.name", "table", Ui.Blue);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 3 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var hint = new Label { AutoSize = true, Text = "预设名称（仅保存在本机，支持导出分享）：", Margin = new Padding(0, 0, 0, 8) };
            name.Text = initial ?? ""; name.MaxLength = 80; root.Controls.Add(hint, 0, 0); root.Controls.Add(name, 0, 1);
            var actions = Ui.ActionBar(); var save = Ui.Button("确定", 90); var cancel = Ui.Button("取消", 90);
            actions.Dock = DockStyle.None; actions.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            actions.Controls.Add(save); actions.Controls.Add(cancel); root.Controls.Add(actions, 0, 2); Controls.Add(root);
            save.Click += delegate { try { BatchPresetStore.FileFor(feature, ChosenName); DialogResult = DialogResult.OK; } catch (Exception error) { UiDialog.Show(this, error.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); } };
            cancel.DialogResult = DialogResult.Cancel; CancelButton = cancel; AcceptButton = save;
            root.SizeChanged += delegate { hint.MaximumSize = new Size(Math.Max(100, root.ClientSize.Width - 32), 0); };
        }
    }

    internal sealed class BatchPresetManagerForm : Form
    {
        private sealed class Entry { internal string File; internal BatchPreset Value; public override string ToString() { return Value.Name; } }
        private readonly string feature;
        private readonly ListBox list = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, HorizontalScrollbar = true, Name = "batchPresetList" };
        private readonly RichTextBox detail = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true, BackColor = Color.White, BorderStyle = BorderStyle.None, Name = "batchPresetDetail" };
        private readonly TextBox search = new TextBox { Width = 180, Name = "batchPresetSearch" };
        private readonly Label status = new Label { AutoSize = true, Dock = DockStyle.Top, ForeColor = Ui.InkDim, Name = "batchPresetStatus" };
        private readonly RoundedButton apply = Ui.Button("应用到窗口", 130), rename = Ui.Button("重命名", 90), remove = Ui.Button("删除", 80), export = Ui.Button("导出…", 90);
        private readonly List<Entry> all = new List<Entry>();
        private readonly Font heading, small;
        internal BatchPreset Chosen;
        internal BatchPresetManagerForm(string feature)
        {
            this.feature = feature; Text = feature == "company" ? "单体模板预设管理" : "单体汇总预设管理";
            DialogSizing.Prepare(this, "batch.preset.manager", "table", Ui.Blue);
            heading = Ui.MakeFont(11, FontStyle.Bold, Font); small = Ui.MakeFont(9, FontStyle.Regular, Font);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var top = new TableLayoutPanel { AutoSize = true, Dock = DockStyle.Top, ColumnCount = 3, Margin = new Padding(0, 0, 0, 10) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var import = Ui.Button("导入预设…", 120); top.Controls.Add(import, 0, 0); top.Controls.Add(new Label { AutoSize = true, Text = "搜索名称", Anchor = AnchorStyles.Right }, 1, 0); top.Controls.Add(search, 2, 0);
            var middle = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 36)); middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 64)); middle.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var left = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(1), Margin = new Padding(0, 0, 6, 0) }; left.Controls.Add(list);
            var right = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(14), Margin = new Padding(6, 0, 0, 0) }; right.Controls.Add(detail); middle.Controls.Add(left, 0, 0); middle.Controls.Add(right, 1, 0);
            var bottom = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, Margin = new Padding(0, 12, 0, 0) };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var leftActions = Ui.ActionBar(); var rightActions = Ui.ActionBar(); rightActions.FlowDirection = FlowDirection.RightToLeft;
            leftActions.WrapContents = rightActions.WrapContents = false;
            var close = Ui.Button("关闭", 90); leftActions.Controls.AddRange(new Control[] { rename, remove, export }); rightActions.Controls.AddRange(new Control[] { close, apply });
            bottom.Controls.Add(leftActions, 0, 0); bottom.Controls.Add(rightActions, 1, 0);
            status.Text = "预设只保存设置，不含表格数据；应用后请重新预检。";
            root.Controls.Add(top, 0, 0); root.Controls.Add(middle, 0, 1); root.Controls.Add(status, 0, 2); root.Controls.Add(bottom, 0, 3); Controls.Add(root);
            search.TextChanged += delegate { Fill(null); }; list.SelectedIndexChanged += delegate { ShowDetail(); };
            import.Click += delegate { Safe(Import); }; export.Click += delegate { Safe(Export); }; rename.Click += delegate { Safe(Rename); }; remove.Click += delegate { Safe(Remove); };
            apply.Click += delegate { if (Current != null) { Chosen = Current.Value; DialogResult = DialogResult.OK; } };
            close.DialogResult = DialogResult.Cancel; CancelButton = close;
            root.SizeChanged += delegate { status.MaximumSize = new Size(Math.Max(200, root.ClientSize.Width - 32), 0); };
            Disposed += delegate { heading.Dispose(); small.Dispose(); };
            if (DialogPreviewFactory.IsConstructing)
            {
                all.Add(new Entry { File = "", Value = new BatchPreset { Feature = feature, Name = "示例年度预设", Context = "示例模板.xlsx", Saved = "2026-01-01T10:00:00", Values = new Dictionary<string, string> { { "source", "公司主题汇总" }, { "target", "Lead" }, { "header", "1" } } } }); Fill(null);
            }
            else Reload(null);
        }
        private Entry Current { get { return list.SelectedItem as Entry; } }
        private void Safe(Action action) { try { action(); } catch (Exception error) { UiDialog.Show(this, error.Message, "预设管理", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); } }
        private void Reload(string selected)
        {
            all.Clear(); int broken = 0;
            try
            {
                string folder = BatchPresetStore.Library(feature);
                if (Directory.Exists(folder)) foreach (string file in Directory.GetFiles(folder, "*.xybatch.json"))
                    try { all.Add(new Entry { File = file, Value = BatchPresetStore.Read(file, feature) }); } catch { broken++; }
                all.Sort((a, b) => string.Compare(b.Value.Saved, a.Value.Saved, StringComparison.Ordinal));
                status.Text = "共 " + all.Count + " 个预设；仅保存参数，应用后重新预检。" + (broken > 0 ? "有 " + broken + " 个文件无法读取，请检查格式和版本。" : "");
            }
            catch (Exception error) { status.Text = "预设库暂不可读：" + error.Message; }
            Fill(selected);
        }
        private void Fill(string selected)
        {
            string prior = selected ?? (Current == null ? null : Current.File); list.BeginUpdate(); list.Items.Clear();
            foreach (var item in all.Where(e => e.Value.Name.IndexOf(search.Text.Trim(), StringComparison.CurrentCultureIgnoreCase) >= 0)) list.Items.Add(item);
            int index = -1;
            for (int i = 0; i < list.Items.Count; i++) if (string.Equals(((Entry)list.Items[i]).File, prior, StringComparison.OrdinalIgnoreCase)) index = i;
            if (index < 0 && list.Items.Count > 0) index = 0; list.SelectedIndex = index; list.EndUpdate(); ShowDetail();
        }
        private void Line(string text, bool title)
        {
            detail.SelectionFont = title ? heading : small; detail.SelectionColor = title ? Ui.Blue : Ui.Ink; detail.AppendText(text + Environment.NewLine);
        }
        private void ShowDetail()
        {
            var entry = Current; apply.Enabled = rename.Enabled = remove.Enabled = export.Enabled = entry != null; detail.Clear();
            if (entry == null) { Line(all.Count == 0 ? "还没有预设。请在主窗口的齿轮菜单中选择“保存为预设…”，或导入已有预设文件。" : "没有名称匹配的预设。", false); return; }
            var value = entry.Value; Line(value.Name, true); Line("保存时间：" + (value.Saved ?? "未知"), false); Line("文件位置", true); Line(value.Context ?? "（未记录）", false);
            string[][] groups = feature == "company" ? new[] {
                new[] { "命名", "naming|命名名称", "removeTail|末尾去除字符数", "stripSuffix|去除 Excel 后缀" },
                new[] { "来源", "source|工作表", "header|表头行", "first|数据起始行", "last|数据结束行", "code|替换字段的编码来源", "full|替换字段的全称来源" },
                new[] { "目标", "update|更新 Lead", "target|工作表", "targetHeader|表头行", "targetRow|写入行", "targetCode|编码字段", "targetName|全称字段", "output|输出文件夹" }
            } : new[] {
                new[] { "工作表与模板", "lead|Lead 工作表", "header|Lead 表头行", "target|汇总工作表", "targetHeader|汇总表头行", "column|模板列", "placeholder|占位名", "sourceFile|模板公司文件名" },
                new[] { "年度", "annual|按年度展开", "firstYear|开始年度", "lastYear|结束年度", "yearRow|年度表头行", "descending|年度倒序", "checkFiles|检查年度来源文件", "yearPattern|年度目录段" }
            };
            foreach (var group in groups) { Line("\n" + group[0], true); foreach (string field in group.Skip(1)) { var pair = field.Split('|'); string text; if (!value.Values.TryGetValue(pair[0], out text) || text.Length == 0) text = "（未设置）"; if (text == "True") text = "是"; if (text == "False") text = "否"; Line(pair[1] + "：" + text, false); } }
            detail.SelectionStart = 0; detail.ScrollToCaret();
        }
        private void Import()
        {
            string[] files = HostDialogService.OpenFiles(this, "单体预设 (*.xybatch.json)|*.xybatch.json|JSON (*.json)|*.json");
            var failures = new List<string>(); string selected = null;
            foreach (string file in files) try { selected = BatchPresetStore.Import(file, feature); } catch (Exception error) { failures.Add(Path.GetFileName(file) + "：" + error.Message); }
            if (files.Length == 0) return; search.Clear(); Reload(selected);
            if (failures.Count > 0) UiDialog.Show(this, string.Join(Environment.NewLine, failures.ToArray()), "部分预设未导入", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
        }
        private void Export()
        {
            var entry = Current; if (entry == null) return;
            string file = HostDialogService.SaveSettings(this, entry.Value.Name + ".xybatch.json"); if (file == null) return;
            BatchPresetStore.Write(file, entry.Value, true); status.Text = "已导出：" + file;
        }
        private void Rename()
        {
            var entry = Current; if (entry == null) return; string kind = feature, initial = entry.Value.Name;
            string name = HostDialogService.ShowIsolatedResult<string>(this, delegate { return new BatchPresetNameForm(kind, "重命名预设", initial); }, delegate(Form form) { return ((BatchPresetNameForm)form).ChosenName; }, null);
            if (name == null || name == initial) return; string file = BatchPresetStore.Rename(entry.File, feature, name); search.Clear(); Reload(file);
        }
        private void Remove()
        {
            var entry = Current; if (entry == null) return;
            if (UiDialog.Show(this, "删除预设“" + entry.Value.Name + "”？只删除本机预设，不影响工作簿。", "删除预设", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            BatchPresetStore.Delete(entry.File); Reload(null);
        }
    }
}
