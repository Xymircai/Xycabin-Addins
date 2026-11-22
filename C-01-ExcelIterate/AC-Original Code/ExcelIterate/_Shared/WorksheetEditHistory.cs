using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal interface IWorksheetChange
    {
        string Address { get; }
        void Check(object sheet, bool after);
        void Write(object sheet, bool after);
        string DescribeBefore();
    }

    // Retain IUnknown so a newly opened sheet cannot reuse the original identity.
    // No range or selection references are held in the history.
    internal sealed class WorksheetChangeSet : IDisposable
    {
        private IntPtr identity;
        private readonly object managedSheet;
        private readonly string bookName, sheetName, fullName;
        private readonly bool readOnly;
        internal readonly List<IWorksheetChange> Changes = new List<IWorksheetChange>();
        internal string Label { get { return fullName + " / " + sheetName; } }

        internal WorksheetChangeSet(object sheet)
        {
            using (var bag = new ComBag())
            {
                dynamic book = bag.Own(((dynamic)sheet).Parent);
                bookName = (string)book.Name;
                fullName = (string)book.FullName;
                readOnly = (bool)book.ReadOnly;
                sheetName = (string)((dynamic)sheet).Name;
            }
            if (string.IsNullOrEmpty(bookName) || string.IsNullOrEmpty(fullName) || string.IsNullOrEmpty(sheetName))
                throw new InvalidOperationException("无法记录目标工作簿和工作表的身份。");
            if (Marshal.IsComObject(sheet)) identity = Marshal.GetIUnknownForObject(sheet);
            else managedSheet = sheet;
        }

        internal object Resolve(object app, ComBag bag)
        {
            try
            {
                dynamic books = bag.Own(((dynamic)app).Workbooks);
                dynamic book = bag.Own(books.Item(bookName));
                if (!string.Equals((string)book.FullName, fullName, StringComparison.OrdinalIgnoreCase) || (bool)book.ReadOnly != readOnly)
                    throw new InvalidOperationException("工作簿路径或写入状态已变化。");
                dynamic sheets = bag.Own(book.Worksheets);
                object sheet = bag.Own(sheets.Item(sheetName));
                if (!string.Equals((string)((dynamic)sheet).Name, sheetName, StringComparison.Ordinal) || !SameSheet(sheet))
                    throw new InvalidOperationException("工作表已被替换或改名。");
                CheckGroups((object)book);
                return sheet;
            }
            catch (Exception error)
            {
                throw new InvalidOperationException(Label + " 无法作为操作目标。请确认原工作簿和工作表仍打开且可编辑。\r\n" + error.Message, error);
            }
        }

        internal static void CheckGroups(object book)
        {
            using (var bag = new ComBag())
            {
                dynamic windows = bag.Own(((dynamic)book).Windows);
                for (int i = 1; i <= (int)windows.Count; i++)
                {
                    using (var refs = new ComBag())
                    {
                        dynamic window = refs.Own(windows.Item(i));
                        dynamic selected = refs.Own(window.SelectedSheets);
                        if ((int)selected.Count > 1) throw new InvalidOperationException("请先取消工作表组合选择。");
                    }
                }
            }
        }

        private bool SameSheet(object sheet)
        {
            if (identity == IntPtr.Zero) return ReferenceEquals(managedSheet, sheet);
            if (!Marshal.IsComObject(sheet)) return false;
            IntPtr current = Marshal.GetIUnknownForObject(sheet);
            try { return current == identity; }
            finally { Marshal.Release(current); }
        }

        public void Dispose()
        {
            if (identity != IntPtr.Zero) { Marshal.Release(identity); identity = IntPtr.Zero; }
        }
    }

    internal sealed class WorksheetEditBatch : IDisposable
    {
        private sealed class BoundChange
        {
            internal readonly WorksheetChangeSet Target;
            internal readonly object Sheet;
            internal readonly IWorksheetChange Change;

            internal BoundChange(WorksheetChangeSet target, object sheet, IWorksheetChange change)
            { Target = target; Sheet = sheet; Change = change; }

            internal string Label { get { return Target.Label + "!" + Change.Address; } }

            internal void Check(bool after)
            {
                try { Change.Check(Sheet, after); }
                catch (Exception error) { throw new InvalidOperationException(Label + "：" + error.Message, error); }
            }
            internal void Write(bool after)
            {
                try { Change.Write(Sheet, after); }
                catch (Exception error) { throw new InvalidOperationException(Label + "：" + error.Message, error); }
            }
        }

        private readonly List<WorksheetChangeSet> targets = new List<WorksheetChangeSet>();
        internal readonly string Name;
        internal bool LastExecutionCommitted { get; private set; }
        internal WorksheetEditBatch(string name) { Name = name; }

        internal string DescribeTargets()
        {
            if (targets.Count == 0) return Name;
            var first = targets[0];
            string description = first.Label;
            if (first.Changes.Count > 0)
            {
                description += " · " + first.Changes[0].Address;
                if (first.Changes.Count > 1) description += " 等 " + first.Changes.Count + " 项";
            }
            if (targets.Count > 1) description = targets.Count + " 个工作表 · " + description;
            return description;
        }

        internal WorksheetChangeSet AddSheet(object sheet)
        {
            var target = new WorksheetChangeSet(sheet);
            targets.Add(target);
            return target;
        }

        // Apply and undo both preflight every target. Count a setter even if it throws.
        internal void Execute(object app, bool forward, string recoveryDirectory)
        {
            LastExecutionCommitted = false;
            using (var bag = new ComBag())
            {
                var bound = new List<BoundChange>();
                foreach (var target in targets)
                {
                    if (target.Changes.Count == 0) continue;
                    object sheet = target.Resolve(app, bag);
                    foreach (var change in target.Changes) bound.Add(new BoundChange(target, sheet, change));
                }
                if (bound.Count == 0) return;
                if (!forward) bound.Reverse();

                using (new HostQuiet(app))
                {
                    foreach (var item in bound) item.Check(!forward);
                    int attempted = 0;
                    try
                    {
                        foreach (var item in bound)
                        {
                            attempted++;
                            item.Write(forward);
                            item.Check(forward);
                        }
                    }
                    catch (Exception error)
                    {
                        var failures = new List<string>();
                        for (int i = attempted - 1; i >= 0; i--)
                        {
                            BoundChange item = bound[i];
                            string reason = null;
                            try { item.Change.Write(item.Sheet, !forward); } catch (Exception rollbackError) { reason = rollbackError.Message; }
                            try { item.Check(!forward); }
                            catch (Exception rollbackError)
                            {
                                failures.Add(item.Label + "：" + reason + " " + rollbackError.Message + "\r\n" + item.Change.DescribeBefore());
                            }
                        }
                        string detail = "批次未完成：" + error.Message;
                        if (failures.Count == 0) detail += "\r\n本次已尝试的修改均已回滚并核对。";
                        else detail += WriteRecovery(recoveryDirectory, detail, failures);
                        throw new InvalidOperationException(detail, error);
                    }
                    LastExecutionCommitted = true;
                }
            }
        }

        private static string WriteRecovery(string folder, string detail, List<string> failures)
        {
            string report = detail + "\r\n以下目标未能核对，请检查其内容。\r\n" + string.Join("\r\n", failures);
            try
            {
                Directory.CreateDirectory(folder);
                string path = Path.Combine(folder, "WorksheetEdit-" + Guid.NewGuid().ToString("N") + ".txt");
                File.WriteAllText(path, report, Encoding.UTF8);
                return "\r\n部分目标未能确认还原。恢复详情：" + path;
            }
            catch (Exception error) { return "\r\n恢复日志保存失败：" + error.Message + "\r\n" + string.Join("\r\n", failures); }
        }

        public void Dispose()
        {
            foreach (var target in targets) target.Dispose();
            targets.Clear();
        }
    }

    internal sealed class WorksheetUndoState
    {
        internal readonly long Id;
        internal readonly string Description;
        internal WorksheetUndoState(long id, string description) { Id = id; Description = description; }
    }

    internal sealed class WorksheetUndoButton : RoundedButton
    {
        private readonly ToolTip tooltip = new ToolTip { ShowAlways = true };
        internal WorksheetUndoState State { get; private set; }
        internal WorksheetUndoButton()
        {
            Name = "worksheetUndo"; Text = "撤销操作"; Size = new Size(112, 38); Radius = 10;
            Ui.StyleSecondary(this);
            SetState(null);
        }
        internal void SetState(WorksheetUndoState state)
        {
            State = state;
            Enabled = state != null;
            AccessibleDescription = state == null ? "当前功能没有可撤销的最近批次。" : "原目标：" + state.Description;
            tooltip.SetToolTip(this, AccessibleDescription);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing) tooltip.Dispose();
            base.Dispose(disposing);
        }
    }

    public partial class Connect
    {
        private static WorksheetEditBatch worksheetUndo;
        private static WorksheetUndoState worksheetUndoState;
        private static string worksheetUndoSource;
        private static long worksheetUndoSequence;
        internal static event Action WorksheetUndoChanged;

        internal static WorksheetUndoState UndoState(string source)
        { return worksheetUndoSource == source ? worksheetUndoState : null; }

        internal static void RememberWorksheetEdit(string source, WorksheetEditBatch batch)
        {
            if (batch == null) return;
            var old = worksheetUndo;
            worksheetUndo = batch;
            worksheetUndoSource = source;
            worksheetUndoState = new WorksheetUndoState(++worksheetUndoSequence, batch.DescribeTargets());
            try { if (old != null) old.Dispose(); }
            finally { AnnounceWorksheetUndo(); }
        }

        internal static string UndoWorksheetEdit(string source, long id, object application, string recoveryDirectory)
        {
            var state = UndoState(source);
            if (state == null || state.Id != id) throw new InvalidOperationException("最近的批次已变化，当前操作没有可用的撤销记录。");
            if (_busy) throw new InvalidOperationException("另一项加载项操作正在执行，请稍后重试。");
            _busy = true;
            var batch = worksheetUndo;
            try
            {
                batch.Execute(application, false, recoveryDirectory);
                return "已撤销「" + batch.Name + "」。请按需保存工作簿。";
            }
            catch (Exception error)
            {
                UiDiagnostics.Record("worksheet-undo", error);
                if (batch.LastExecutionCommitted)
                    throw new InvalidOperationException("已还原数据，收尾时出现错误。\r\n" + error.Message, error);
                throw;
            }
            finally
            {
                try { if (batch.LastExecutionCommitted) ClearWorksheetUndo(); }
                finally { _busy = false; }
            }
        }

        private static void AnnounceWorksheetUndo()
        {
            var handlers = WorksheetUndoChanged;
            if (handlers == null) return;
            foreach (Action handler in handlers.GetInvocationList())
                try { handler(); } catch (Exception error) { UiDiagnostics.Record("worksheet-undo-dialog", error); }
        }

        private static void ClearWorksheetUndo()
        {
            WorksheetEditBatch old = worksheetUndo;
            worksheetUndo = null;
            worksheetUndoSource = null;
            worksheetUndoState = null;
            if (old != null)
                try { old.Dispose(); } finally { AnnounceWorksheetUndo(); }
        }
    }
}
