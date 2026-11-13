using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class BatchInsertTrimDeletePanel : UserControl
    {
        private sealed class Page
        {
            internal RadioButton Start, End, Middle;
            internal NumericUpDown Count, First, Last;
            internal TextBox Text;
            internal FlowLayoutPanel CountRow, MiddleRow;
            internal FlowLayoutPanel Content;
            internal TabPage Tab;
        }
        private readonly Page[] pages = new Page[3];
        internal readonly TabControl Tabs;
        internal event Action OptionsChanged;
        private bool fitting;
        internal BatchInsertTrimDeletePanel()
        {
            Name = "batchEditPanel"; Dock = DockStyle.Top; Height = 244; Margin = new Padding(0, 8, 0, 8); BackColor = Ui.Canvas;
            Tabs = new TabControl { Name = "batchEditTabs", Dock = DockStyle.Fill, Padding = new Point(18, 6) }; Controls.Add(Tabs);
            string[] titles = { "插入", "截取", "删除" };
            for (int i = 0; i < pages.Length; i++)
            {
                int index = i; var page = new Page(); pages[i] = page;
                var tab = new TabPage(titles[i]) { BackColor = Ui.Canvas, Padding = new Padding(10) }; Tabs.TabPages.Add(tab);
                var content = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
                page.Content = content; page.Tab = tab;
                tab.Controls.Add(content);
                var modes = Row(); page.Start = Radio("从开端计数", "batch" + i + "Start", true); page.End = Radio("从末尾计数", "batch" + i + "End", false);
                modes.Controls.AddRange(new Control[] { page.Start, page.End });
                if (i != 0) { page.Middle = Radio("中间计数", "batch" + i + "Middle", false); modes.Controls.Add(page.Middle); }
                content.Controls.Add(modes);
                page.CountRow = Row(); page.CountRow.Controls.Add(Caption(i == 0 ? "插入位置" : "字符数量"));
                page.Count = Number("batch" + i + "Count", i == 0 ? 0 : 1); page.CountRow.Controls.Add(page.Count); page.CountRow.Controls.Add(Caption(i == 0 ? "（0 为所选一端）" : "个字符")); content.Controls.Add(page.CountRow);
                page.MiddleRow = Row(); page.MiddleRow.Controls.Add(Caption("开头位置")); page.First = Number("batch" + i + "First", 1); page.MiddleRow.Controls.Add(page.First);
                page.MiddleRow.Controls.Add(Caption("末尾位置")); page.Last = Number("batch" + i + "Last", 1); page.MiddleRow.Controls.Add(page.Last); page.MiddleRow.Visible = false; content.Controls.Add(page.MiddleRow);
                if (i == 0)
                {
                    content.Controls.Add(Caption("插入文本"));
                    page.Text = new TextBox { Name = "batchInsertText", Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, Width = 540, Height = 60, MaxLength = BatchTextEdit.MaximumTextLength, Margin = new Padding(0, 3, 0, 0) };
                    content.Controls.Add(page.Text); content.SizeChanged += delegate { page.Text.Width = Math.Max(120, content.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 4); };
                    page.Text.TextChanged += delegate { Changed(); };
                }
                else content.Controls.Add(Caption("中间区间从第 1 个字符起算，包含首尾。"));
                page.Start.CheckedChanged += delegate { SwitchMode(index); }; page.End.CheckedChanged += delegate { SwitchMode(index); };
                if (page.Middle != null) page.Middle.CheckedChanged += delegate { SwitchMode(index); };
                page.Count.ValueChanged += delegate { Changed(); }; page.First.ValueChanged += delegate { Changed(); }; page.Last.ValueChanged += delegate { Changed(); };
                page.Count.TextChanged += delegate { Changed(); }; page.First.TextChanged += delegate { Changed(); }; page.Last.TextChanged += delegate { Changed(); };
            }
            Tabs.SelectedIndex = 0;
            Tabs.SelectedIndexChanged += delegate { FitContent(); Changed(); };
            SizeChanged += delegate { FitContent(); };
            FitContent();
        }
        private static FlowLayoutPanel Row() { return new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 3, 0, 3) }; }
        private static Label Caption(string text) { var label = Ui.Label(text, 10, false); label.AutoSize = true; label.Margin = new Padding(0, 6, 10, 3); return label; }
        private static RadioButton Radio(string text, string name, bool selected) { return new RadioButton { Text = text, Name = name, Checked = selected, AutoSize = true, Margin = new Padding(0, 3, 20, 3) }; }
        private static NumericUpDown Number(string name, int value) { return new NumericUpDown { Name = name, Minimum = 0, Maximum = int.MaxValue, Value = value, Width = 112, DecimalPlaces = 0, Margin = new Padding(0, 3, 12, 3) }; }
        private static int ReadNumber(NumericUpDown input)
        { int value; if (!int.TryParse(input.Text, NumberStyles.None, CultureInfo.InvariantCulture, out value)) throw new ArgumentException("位置和字符数必须为非负整数。"); return value; }
        private void SwitchMode(int index)
        { var page = pages[index]; bool middle = page.Middle != null && page.Middle.Checked; page.CountRow.Visible = !middle; page.MiddleRow.Visible = middle; FitContent(); Changed(); }
        private void FitContent()
        {
            if (fitting || Tabs == null) return; fitting = true;
            try
            {
                var page = pages[Math.Max(0, Tabs.SelectedIndex)]; if (page == null || page.Content == null) return;
                page.Content.PerformLayout();
                int chrome = Math.Max(Tabs.Font.Height + 12, Tabs.Height - Tabs.DisplayRectangle.Height);
                int content = page.Content.GetPreferredSize(new Size(Math.Max(1, ClientSize.Width - page.Tab.Padding.Horizontal), 0)).Height;
                Height = chrome + page.Tab.Padding.Vertical + content + 4;
            }
            finally { fitting = false; }
        }
        private void Changed() { var handler = OptionsChanged; if (handler != null) handler(); }
        internal BatchTextEditOptions Options()
        {
            // Native tab handles can briefly report -1 during construction/recreation.
            int index = Math.Max(0, Tabs.SelectedIndex); var page = pages[index];
            var options = new BatchTextEditOptions { Operation = (BatchTextEditOperation)index, Origin = page.Start.Checked ? BatchTextCountOrigin.Start : page.End.Checked ? BatchTextCountOrigin.End : BatchTextCountOrigin.Middle, Text = page.Text == null ? "" : page.Text.Text };
            if (options.Origin == BatchTextCountOrigin.Middle) { options.Start = ReadNumber(page.First); options.End = ReadNumber(page.Last); }
            else options.Count = ReadNumber(page.Count);
            options.Validate(); return options;
        }
    }
}
