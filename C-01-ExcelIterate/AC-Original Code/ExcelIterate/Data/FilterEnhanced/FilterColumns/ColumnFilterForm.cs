using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class ColumnFilterForm : Form
    {
        internal readonly DataGridView Grid;
        internal readonly RoundedButton Apply, Cancel, SelectAll, Reset;
        internal readonly WorksheetUndoButton Undo = new WorksheetUndoButton();
        internal readonly Label Hint;
        internal readonly TextBox Search;
        internal readonly Label SearchCount;
        private IList<ColumnChoice> original;
        private bool clearNext;
        internal event Action<IList<ColumnChoice>> ApplyRequested;
        internal ColumnFilterForm(IList<ColumnChoice> choices) : this(choices, null) { }
        internal ColumnFilterForm(IList<ColumnChoice> choices, string range)
        {
            original = choices;
            Text = "按列筛选";
            DialogSizing.Prepare(this, "columns.filter", "table", Ui.Blue);
            KeyPreview = true;
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Hint = Ui.Label((range == null ? "" : range + "\r\n") + "勾选要显示的列，未勾选列将隐藏。点击“应用”才修改工作表；至少保留一列。", 9, false);
            Hint.AutoSize = true; Hint.MaximumSize = new Size(700, 0); Hint.Margin = new Padding(0, 0, 0, 12); Hint.AutoEllipsis = false;
            Grid = new DataGridView { Dock = DockStyle.Fill, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AllowUserToResizeRows = false,
                RowHeadersVisible = false, MultiSelect = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoGenerateColumns = false, ScrollBars = ScrollBars.Both };
            Ui.StyleGrid(Grid);
            Grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Visible", HeaderText = "显示", Width = 70, MinimumWidth = 60 });
            Grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Letter", HeaderText = "列号", ReadOnly = true, Width = 80, MinimumWidth = 65 });
            Grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Header", HeaderText = "表头内容", ReadOnly = true, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 260 });
            Grid.ShowCellToolTips = true;
            UiCard card = new UiCard { Dock = DockStyle.Fill, Padding = new Padding(12), Margin = Padding.Empty };
            card.Controls.Add(Grid);
            foreach (ColumnChoice choice in choices)
            {
                int index = Grid.Rows.Add(choice.Visible, choice.Letter, choice.Header);
                Grid.Rows[index].Tag = choice;
                Grid.Rows[index].Cells[2].ToolTipText = choice.Header;
            }
            Grid.CurrentCellDirtyStateChanged += delegate { if (Grid.IsCurrentCellDirty) Grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            TableLayoutPanel searchRow = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 5, Margin = new Padding(0, 0, 0, 12) };
            searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Label searchLabel = Ui.Label("搜索列名", 10, false); searchLabel.AutoSize = true; searchLabel.Anchor = AnchorStyles.Left;
            Search = new TextBox { Font = Ui.MakeFont(10, FontStyle.Regular, null), Margin = new Padding(8, 4, 8, 4), Width = 220 };
            RoundedButton clear = Ui.Button("清除", 72); SearchCount = Ui.Label("", 9, false); SearchCount.AutoSize = true; SearchCount.Anchor = AnchorStyles.Left; SearchCount.Margin = new Padding(8);
            clear.Click += delegate { Search.Clear(); Search.Focus(); };
            searchRow.Controls.Add(searchLabel, 0, 0); searchRow.Controls.Add(Ui.Grip(Search, (Control)null), 1, 0); searchRow.Controls.Add(clear, 2, 0); searchRow.Controls.Add(SearchCount, 3, 0);
            FlowLayoutPanel actions = Ui.ActionBar(); actions.Dock = DockStyle.Fill; actions.Margin = new Padding(0, 12, 0, 0); actions.Padding = Padding.Empty;
            SelectAll = Ui.Button("全选", 88); Reset = Ui.Button("恢复原勾选", 128); Apply = Ui.Button("应用", 96); Ui.StylePrimary(Apply); Cancel = Ui.Button("关闭", 96);
            actions.Controls.AddRange(new Control[] { SelectAll, Reset, Apply, Undo, Cancel });
            SelectAll.Click += delegate { Grid.EndEdit(); foreach (DataGridViewRow row in Grid.Rows) if (row.Visible) row.Cells[0].Value = !clearNext; clearNext = !clearNext; UpdateSelectLabel(); };
            Reset.Click += delegate { Grid.EndEdit(); foreach (DataGridViewRow row in Grid.Rows) row.Cells[0].Value = !((ColumnChoice)row.Tag).OriginalHidden; clearNext = false; UpdateSelectLabel(); };
            Cancel.Click += delegate { Close(); };
            Apply.Click += delegate
            {
                IList<ColumnChoice> pending = Pending(); bool any = false; foreach (ColumnChoice choice in pending) any |= choice.Visible;
                if (!any) { UiDialog.Show(this, "请至少保留一列勾选。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (ApplyRequested != null) ApplyRequested(pending);
            };
            Search.TextChanged += delegate { FilterRows(); };
            layout.Controls.Add(Hint, 0, 0); layout.Controls.Add(searchRow, 0, 1); layout.Controls.Add(card, 0, 2); layout.Controls.Add(actions, 0, 3); Controls.Add(layout);
            FilterRows();
            Apply.Enabled = original.Count > 0;
            Cancel.DialogResult = DialogResult.Cancel; CancelButton = Cancel; AcceptButton = null;
            Resize += delegate { Hint.MaximumSize = new Size(Math.Max(240, ClientSize.Width - Ui.Px(this, 32)), 0); };
        }
        internal void RefreshVisibility(IList<ColumnChoice> choices)
        {
            original = choices;
            var byNumber = new Dictionary<int, ColumnChoice>();
            foreach (var choice in choices) byNumber.Add(choice.Number, choice);
            foreach (DataGridViewRow row in Grid.Rows)
            {
                var choice = byNumber[((ColumnChoice)row.Tag).Number];
                row.Tag = choice;
                row.Cells[0].Value = choice.Visible;
            }
            FilterRows();
        }
        private void FilterRows()
        {
            Grid.EndEdit(); Grid.CurrentCell = null;
            string term = Search.Text.Trim(); int count = 0;
            foreach (DataGridViewRow row in Grid.Rows)
            {
                row.Visible = ((ColumnChoice)row.Tag).Header.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0;
                if (row.Visible) count++;
            }
            SearchCount.Text = count + " / " + original.Count + " 列";
            clearNext = false; UpdateSelectLabel();
        }
        private void UpdateSelectLabel() { SelectAll.Text = clearNext ? "取消全选" : "全选"; }
        internal IList<ColumnChoice> Pending()
        {
            Grid.EndEdit(); List<ColumnChoice> result = new List<ColumnChoice>();
            Dictionary<int, bool> pending = new Dictionary<int, bool>();
            foreach (DataGridViewRow row in Grid.Rows) pending.Add(((ColumnChoice)row.Tag).Number, Convert.ToBoolean(row.Cells[0].Value));
            foreach (ColumnChoice choice in original) result.Add(new ColumnChoice { Number = choice.Number, Letter = choice.Letter, Header = choice.Header, OriginalHidden = choice.OriginalHidden, Visible = pending[choice.Number] });
            return result;
        }
        protected override void OnKeyDown(KeyEventArgs e)
        { if (e.KeyCode == Keys.Escape) { e.Handled = true; e.SuppressKeyPress = true; Close(); } base.OnKeyDown(e); }
    }
}
