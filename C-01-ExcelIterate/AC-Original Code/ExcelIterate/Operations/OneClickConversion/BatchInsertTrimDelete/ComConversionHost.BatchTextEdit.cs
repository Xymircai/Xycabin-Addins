using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    internal sealed partial class ComConversionHost : IBatchTextAccess, IConversionCalculationControl, IBatchFormulaSyntax
    {
        private bool batchFormula2 = true;
        public int Calculation { get { return Convert.ToInt32(((dynamic)application).Calculation); } set { ((dynamic)application).Calculation = value; } }
        private sealed class BatchComCell : IConversionCell, IBatchTextPosition
        {
            private readonly ComConversionHost owner;
            public string Address { get; private set; }
            public string LocalAddress { get; private set; }
            public ConversionCellState Original { get; private set; }
            internal BatchComCell(ComConversionHost host, string address, ConversionCellState state)
            { owner = host; LocalAddress = address; Address = host.Description + "\r\n单元格：" + address; Original = state; }
            public void CheckWritable() { owner.CheckWritableRange(LocalAddress); }
            public ConversionCellState Read() { return owner.ReadTextRange(LocalAddress, false, false)[LocalAddress]; }
            public void Write(ConversionCellState state)
            {
                var point = ConversionRangeAddress.Coordinates(LocalAddress);
                var change = new BatchEditChange { Cell = this, LocalAddress = LocalAddress, Row = point[0], Column = point[1], Before = Original, After = state };
                owner.WriteTextBlock(BatchEditBlock.Create(new[] { change }, false)[0]);
            }
        }
        private static object Element(object value, int row, int column)
        { var array = value as Array; return array == null ? value : array.GetValue(row + array.GetLowerBound(0), column + array.GetLowerBound(1)); }
        private object FormulaValues(object range)
        {
            try { return batchFormula2 ? ((dynamic)range).Formula2 : ((dynamic)range).Formula; }
            catch (Exception error) { if (!batchFormula2 || !Missing(error)) throw; batchFormula2 = false; return ((dynamic)range).Formula; }
        }
        private IDictionary<string, ConversionCellState> ReadTextRange(string address, bool display, bool skipMerged)
        {
            object range = null; var states = new Dictionary<string, ConversionCellState>(StringComparer.OrdinalIgnoreCase);
            try
            {
                range = ((dynamic)worksheet).Range[address]; dynamic r = range;
                object values = r.Value2, formulas = FormulaValues(range), hasFormula = r.HasFormula, format = r.NumberFormat;
                object merged = skipMerged ? r.MergeCells : (object)false;
                string[] ends = address.Split(':');
                int[] first = ConversionRangeAddress.Coordinates(ends[0]), last = ConversionRangeAddress.Coordinates(ends[ends.Length - 1]);
                for (int row = first[0]; row <= last[0]; row++) for (int column = first[1]; column <= last[1]; column++)
                {
                    string name = CellAddress.Name(row, column); object cell = null;
                    try
                    {
                        bool formula;
                        if (hasFormula is bool) formula = (bool)hasFormula;
                        else { cell = ((dynamic)worksheet).Range[name]; formula = Boolean((object)((dynamic)cell).HasFormula, "HasFormula"); }
                        if (!(merged is bool) || (bool)merged)
                        {
                            if (cell == null) cell = ((dynamic)worksheet).Range[name];
                            if (Boolean((object)((dynamic)cell).MergeCells, "MergeCells"))
                            {
                                object area = ((dynamic)cell).MergeArea;
                                try { if (Convert.ToInt32(((dynamic)cell).Row) != Convert.ToInt32(((dynamic)area).Row) || Convert.ToInt32(((dynamic)cell).Column) != Convert.ToInt32(((dynamic)area).Column)) continue; }
                                finally { Release(area); }
                            }
                        }
                        string numberFormat = format as string;
                        if (numberFormat == null) { if (cell == null) cell = ((dynamic)worksheet).Range[name]; numberFormat = ((dynamic)cell).NumberFormat as string; }
                        if (numberFormat == null) throw new InvalidOperationException(name + " 无法读取数字格式。");
                        var state = new ConversionCellState { IsFormula = formula, Value = Element(values, row - first[0], column - first[1]), NumberFormat = numberFormat, Display = "" };
                        if (formula) { state.Formula = Element(formulas, row - first[0], column - first[1]) as string; if (state.Formula == null) throw new InvalidOperationException(name + " 无法读取公式原文。"); }
                        else if (display) { if (cell == null) cell = ((dynamic)worksheet).Range[name]; state.Display = ((dynamic)cell).Text as string; if (state.Display == null) throw new InvalidOperationException(name + " 无法读取显示文字。"); }
                        states.Add(name, state);
                    }
                    finally { Release(cell); }
                }
            }
            finally { Release(range); }
            return states;
        }
        public IList<IConversionCell> CollectTextEdit()
        {
            RecheckScope(); var result = new List<IConversionCell>();
            foreach (var block in BatchEditBlock.Addresses(ConversionRangeAddress.Cells(Address)))
                foreach (var pair in ReadTextRange(block.Address, true, true)) result.Add(new BatchComCell(this, pair.Key, pair.Value));
            return result;
        }
        private void CheckWritableRange(string address)
        {
            object range = null;
            try
            {
                range = ((dynamic)worksheet).Range[address]; dynamic r = range;
                object arrays = r.HasArray, spills = r.HasSpill;
                bool protectedSheet = Boolean((object)((dynamic)worksheet).ProtectContents, "ProtectContents");
                object locked = protectedSheet ? r.Locked : (object)false;
                if (arrays is bool && !(bool)arrays && spills is bool && !(bool)spills && locked is bool && !(bool)locked) return;
                foreach (string name in ConversionRangeAddress.Cells(address))
                {
                    object cell = ((dynamic)worksheet).Range[name];
                    try
                    {
                        if (Boolean((object)((dynamic)cell).HasArray, "HasArray")) throw new InvalidOperationException(name + " 属于数组公式。");
                        if (Boolean((object)((dynamic)cell).HasSpill, "HasSpill")) throw new InvalidOperationException(name + " 属于动态溢出区域。");
                        if (protectedSheet && Boolean((object)((dynamic)cell).Locked, "Locked")) throw new InvalidOperationException(name + " 受保护且已锁定。");
                    }
                    finally { Release(cell); }
                }
            }
            finally { Release(range); }
        }
        public void CheckTextBlocks(IList<BatchEditBlock> blocks, bool compareResults)
        {
            foreach (var block in blocks)
            {
                CheckWritableRange(block.Address); var current = ReadTextRange(block.Address, false, false);
                foreach (var change in block.Changes)
                {
                    var expected = block.State(change); var actual = current[change.LocalAddress];
                    if (!expected.Matches(actual, compareResults, !block.Restore)) throw new InvalidOperationException(change.LocalAddress + " 内容、公式或格式核验失败。");
                    if (!block.Restore && expected.IsFormula) expected.Formula = actual.Formula;
                }
            }
        }
        public void WriteTextBlock(BatchEditBlock block)
        {
            object range = null;
            try
            {
                range = ((dynamic)worksheet).Range[block.Address]; dynamic r = range;
                var data = new object[block.Bottom - block.Top + 1, block.Right - block.Left + 1];
                foreach (var change in block.Changes)
                {
                    var state = block.State(change);
                    data[change.Row - block.Top, change.Column - block.Left] = block.IsFormula ? (object)state.Formula : block.QuoteText ? "'" + (string)state.Value : state.Value;
                }
                if (block.IsFormula)
                {
                    r.NumberFormat = "General";
                    try
                    {
                        try { if (batchFormula2) r.Formula2 = data; else r.Formula = data; }
                        catch (Exception error) { if (!batchFormula2 || !Missing(error)) throw; batchFormula2 = false; r.Formula = data; }
                    }
                    catch (COMException error) { if (BatchTextEditEngine.IsSyntaxRejection(error)) throw new FormulaInputRejectedException(error); throw; }
                }
                else if (block.QuoteText) { r.NumberFormat = "General"; r.Formula = data; }
                else
                {
                    r.NumberFormat = "@"; r.Value2 = data;
                    // Some WPS builds consume one leading ASCII quote even through Value2.
                    // Escape only a demonstrated one-quote loss; Excel builds that preserve
                    // it already keep the original payload unchanged. Ordinary values never
                    // enter this branch, and no Formula setter is used.
                    if (block.Changes.Any(c => block.State(c).Value is string && ((string)block.State(c).Value).StartsWith("'", StringComparison.Ordinal)))
                    {
                        object actual = r.Value2; bool escaped = false;
                        foreach (var change in block.Changes)
                        {
                            string expected = block.State(change).Value as string;
                            if (expected == null || !expected.StartsWith("'", StringComparison.Ordinal)) continue;
                            int row = change.Row - block.Top, column = change.Column - block.Left;
                            if (Element(actual, row, column) as string == expected.Substring(1)) { data[row, column] = "'" + expected; escaped = true; }
                        }
                        if (escaped) r.Value2 = data;
                    }
                }
                // Temporary text formatting protects raw strings, including leading zeroes,
                // without adding a quote prefix or keeping a forced final Text format.
                r.NumberFormat = block.NumberFormat;
            }
            finally { Release(range); }
        }
        public IList<BatchEditChange> ValidateFormulaSyntax(IList<BatchEditChange> candidates)
        {
            var invalid = new List<BatchEditChange>(); if (candidates.Count == 0) return invalid;
            object sourceBook = null, sheets = null, copy = null, copySheets = null, selected = null, validationSheet = null;
            object oldSelection = null, oldActive = null; bool alertsSet = false, securitySet = false, linksSet = false;
            object alerts = null, security = null, links = null;
            object books = null; var existingBooks = new List<object>();
            try
            {
                RecheckScope(); dynamic a = application; oldSelection = a.Selection; oldActive = a.ActiveSheet;
                alerts = a.DisplayAlerts; alertsSet = true; a.DisplayAlerts = false;
                try { security = a.AutomationSecurity; securitySet = true; a.AutomationSecurity = 3; } catch (Exception error) { if (!Missing(error)) throw; }
                try { links = a.AskToUpdateLinks; linksSet = true; a.AskToUpdateLinks = false; } catch (Exception error) { if (!Missing(error)) throw; }
                sourceBook = ((dynamic)worksheet).Parent; sheets = ((dynamic)sourceBook).Worksheets;
                books = a.Workbooks;
                for (int i = 1; i <= Convert.ToInt32(((dynamic)books).Count); i++) existingBooks.Add(((dynamic)books).Item(i));
                try { ((dynamic)sheets).Copy(); }
                finally
                {
                    // A failing Copy can still create a workbook. Close only an object
                    // absent from the pre-call inventory, never a pre-existing user book.
                    selected = a.ActiveWorkbook;
                    if (selected != null && !existingBooks.Any(book => WorksheetIdentity.SameObject(book, selected))) { copy = selected; selected = null; }
                }
                if (copy == null) throw new InvalidOperationException("无法创建隔离公式预检工作簿，原区域未写入。");
                copySheets = ((dynamic)copy).Worksheets;
                VerifyCopiedContext(sheets, copySheets, sourceBook, copy, candidates);
                object windows = ((dynamic)copy).Windows;
                try
                {
                    for (int i = 1; i <= Convert.ToInt32(((dynamic)windows).Count); i++)
                    {
                        object window = ((dynamic)windows).Item(i);
                        try { ((dynamic)window).Visible = false; if (Boolean((object)((dynamic)window).Visible, "Visible")) throw new InvalidOperationException("无法隐藏自建公式预检窗口。"); }
                        finally { Release(window); }
                    }
                }
                finally { Release(windows); }
                validationSheet = ((dynamic)copySheets).Item(identity.Name);
                bool formula2 = true;
                foreach (var change in candidates)
                {
                    object cell = ((dynamic)validationSheet).Range[change.LocalAddress];
                    try
                    {
                        ((dynamic)cell).NumberFormat = "General";
                        try
                        {
                            try { if (formula2) ((dynamic)cell).Formula2 = change.After.Formula; else ((dynamic)cell).Formula = change.After.Formula; }
                            catch (Exception error) { if (!formula2 || !Missing(error)) throw; formula2 = false; ((dynamic)cell).Formula = change.After.Formula; }
                            if (!Boolean((object)((dynamic)cell).HasFormula, "HasFormula")) invalid.Add(change);
                        }
                        catch (COMException error) { if (!BatchTextEditEngine.IsSyntaxRejection(error)) throw; invalid.Add(change); }
                    }
                    finally { Release(cell); }
                }
            }
            finally
            {
                var errors = new List<string>();
                Release(validationSheet); Release(copySheets); Release(sheets); Release(sourceBook); Release(selected);
                try { if (copy != null) ((dynamic)copy).Close(false); } catch (Exception error) { errors.Add("关闭自建预检工作簿失败：" + error.Message); }
                finally { Release(copy); }
                foreach (var existing in existingBooks) Release(existing); Release(books);
                try { if (oldActive != null) ((dynamic)oldActive).Activate(); if (oldSelection != null) ((dynamic)oldSelection).Select(); } catch (Exception error) { errors.Add("恢复原选区失败：" + error.Message); }
                finally { Release(oldSelection); Release(oldActive); }
                try { if (securitySet) ((dynamic)application).AutomationSecurity = Convert.ToInt32(security); } catch (Exception error) { errors.Add(error.Message); }
                try { if (linksSet) ((dynamic)application).AskToUpdateLinks = Boolean(links, "AskToUpdateLinks"); } catch (Exception error) { errors.Add(error.Message); }
                try { if (alertsSet) ((dynamic)application).DisplayAlerts = Boolean(alerts, "DisplayAlerts"); } catch (Exception error) { errors.Add(error.Message); }
                if (errors.Count > 0) throw new InvalidOperationException(string.Join("；", errors.ToArray()));
            }
            RecheckScope(); return invalid;
        }
        private static void VerifyCopiedContext(object sourceSheets, object copySheets, object sourceBook, object copy, IList<BatchEditChange> candidates)
        {
            int count = Convert.ToInt32(((dynamic)sourceSheets).Count);
            if (Convert.ToInt32(((dynamic)copySheets).Count) != count) throw new InvalidOperationException("隔离预检的工作表上下文不完整。");
            for (int i = 1; i <= count; i++)
            {
                object original = ((dynamic)sourceSheets).Item(i), cloned = ((dynamic)copySheets).Item(i);
                try { if ((string)((dynamic)original).Name != (string)((dynamic)cloned).Name) throw new InvalidOperationException("隔离预检的工作表名称不一致。"); }
                finally { Release(original); Release(cloned); }
            }
            object names = null, copiedNames = null;
            try
            {
                names = ((dynamic)sourceBook).Names; copiedNames = ((dynamic)copy).Names;
                for (int i = 1; i <= Convert.ToInt32(((dynamic)names).Count); i++)
                {
                    object name = ((dynamic)names).Item(i), cloned = null;
                    try { string key = ((dynamic)name).Name; cloned = ((dynamic)copiedNames).Item(key); if (cloned == null) throw new InvalidOperationException("隔离预检缺少名称：" + key); }
                    finally { Release(name); Release(cloned); }
                }
            }
            finally { Release(names); Release(copiedNames); }
            if (!candidates.Any(c => c.After.Formula.IndexOf('[') >= 0)) return;
            for (int i = 1; i <= count; i++)
            {
                object original = ((dynamic)sourceSheets).Item(i), cloned = ((dynamic)copySheets).Item(i), tables = null, copiedTables = null;
                try
                {
                    tables = ((dynamic)original).ListObjects; copiedTables = ((dynamic)cloned).ListObjects;
                    if (Convert.ToInt32(((dynamic)tables).Count) != Convert.ToInt32(((dynamic)copiedTables).Count)) throw new InvalidOperationException("隔离预检的表格上下文不完整。");
                    for (int j = 1; j <= Convert.ToInt32(((dynamic)tables).Count); j++)
                    { object table = ((dynamic)tables).Item(j), other = null; try { other = ((dynamic)copiedTables).Item((string)((dynamic)table).Name); if (other == null) throw new InvalidOperationException("隔离预检缺少结构化表格。"); } finally { Release(table); Release(other); } }
                }
                finally { Release(original); Release(cloned); Release(tables); Release(copiedTables); }
            }
        }
    }
}
