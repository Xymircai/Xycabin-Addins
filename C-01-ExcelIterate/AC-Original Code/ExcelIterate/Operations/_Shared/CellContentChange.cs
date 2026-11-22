using System;

namespace Excel_Iterate
{
    internal sealed class CellContentChange : IWorksheetChange
    {
        private readonly ConversionCellState before, after;
        private readonly bool changesFormat;
        public string Address { get; private set; }

        internal CellContentChange(string address, ConversionCellState before, ConversionCellState after, bool changesFormat)
        { Address = address; this.before = before; this.after = after; this.changesFormat = changesFormat; }

        public void Check(object sheet, bool applied)
        {
            using (var bag = new ComBag())
            {
                object range = bag.Own(((dynamic)sheet).Range[Address]);
                var cell = new ComConversionHost.ComCell(range, sheet, Address, capture: false);
                cell.CheckWritable();
                if (ComConversionHost.Boolean((object)((dynamic)range).MergeCells, "MergeCells"))
                {
                    dynamic area = bag.Own(((dynamic)range).MergeArea);
                    if ((int)((dynamic)range).Row != (int)area.Row || (int)((dynamic)range).Column != (int)area.Column)
                        throw new InvalidOperationException("目标已位于合并单元格内部。");
                }
                var expected = applied ? after : before;
                var actual = cell.Read(false);
                if (!expected.Matches(actual, false, true, changesFormat))
                    throw new InvalidOperationException("目标内容、公式或数字格式已变化。");
            }
        }

        public void Write(object sheet, bool applied)
        {
            using (var bag = new ComBag())
            {
                object range = bag.Own(((dynamic)sheet).Range[Address]);
                var cell = new ComConversionHost.ComCell(range, sheet, Address, capture: false);
                var value = applied ? after : before;
                if (!changesFormat)
                {
                    value = value.Copy();
                    value.NumberFormat = (string)((dynamic)range).NumberFormat;
                }
                cell.Write(value);
            }
        }

        public string DescribeBefore() { return before.Describe(); }
        internal string DescribePlan() { return after.Describe(); }
    }
}
