using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Excel_Iterate
{
    public static partial class EliminationMatrixBuilder
    {
        public static readonly string[] FixedHeaders = new string[]
        { "核算主体", "明细【客商/二级明细】", "备注", "调整事项", "会计科目" };

        /// <summary>固定信息列数。</summary>
        public const int FixedColumns = 5;

        /// <summary>固定五列 + 每年度三列（[分隔][借][贷]）。</summary>
        public static List<MatrixColumn> BuildColumns(IList<string> years, int startColumn)
        {
            List<MatrixColumn> list = new List<MatrixColumn>();
            for (int i = 0; i < FixedColumns; i++)
            {
                EliminationCancellation.Check();
                MatrixColumn column = new MatrixColumn();
                column.Index = i;
                column.Number = startColumn + i;
                column.Letter = Letter(column.Number);
                column.Head2 = FixedHeaders[i];
                list.Add(column);
            }
            for (int i = 0; i < years.Count; i++)
            {
                EliminationCancellation.Check();
                int spacerIndex = FixedColumns + i * 3;
                MatrixColumn spacer = new MatrixColumn();
                spacer.Index = spacerIndex;
                spacer.Number = startColumn + spacerIndex;
                spacer.Letter = Letter(spacer.Number);
                spacer.IsSpacer = true;
                list.Add(spacer);

                int debitIndex = spacerIndex + 1;
                MatrixColumn debit = new MatrixColumn();
                debit.Index = debitIndex;
                debit.Number = startColumn + debitIndex;
                debit.Letter = Letter(debit.Number);
                debit.Year = years[i];
                debit.IsDebit = true;
                debit.Head2 = years[i];
                debit.Head3 = "借方";
                list.Add(debit);

                int creditIndex = spacerIndex + 2;
                MatrixColumn credit = new MatrixColumn();
                credit.Index = creditIndex;
                credit.Number = startColumn + creditIndex;
                credit.Letter = Letter(credit.Number);
                credit.Year = years[i];
                credit.IsDebit = false;
                credit.Head2 = years[i];
                credit.Head3 = "贷方";
                list.Add(credit);
            }
            return list;
        }

        /// <summary>
        /// 从年份文本提取四位年份。只认 (19|20)\d{2}，不回落数值解析——
        /// 否则 Excel 日期序列号（如 45658）会被误当成年份。
        /// </summary>
        public static string ParseYear(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            Match match = Regex.Match(text.Trim(), "(19|20)\\d{2}");
            return match.Success ? match.Value : null;
        }
        /// <summary>列号 → 列字母；缺省起点为 A。</summary>
        public static string Letter(int column)
        {
            return CellAddress.ColumnLetters(Math.Max(1, column));
        }
    }
}
