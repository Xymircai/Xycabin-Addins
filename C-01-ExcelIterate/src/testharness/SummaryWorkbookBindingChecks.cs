using System;
using System.Collections.Generic;
using System.IO;
using Excel_Iterate;

public sealed class SummaryChoiceFakeBook
{
    public string Name { get; set; }
    public string FullName { get; set; }
    public bool ReadOnly { get; set; }
    public bool Closed;
    public Action OnClose;
    public void Close(bool save) { Closed = true; if (OnClose != null) OnClose(); }
}
public sealed class SummaryChoiceFakeBooks
{
    public readonly List<SummaryChoiceFakeBook> Items = new List<SummaryChoiceFakeBook>();
    public int Count { get { return Items.Count; } }
    public int Opens;
    public bool NextReadOnly, FailOpen;
    public SummaryChoiceFakeBook Item(int index) { return Items[index - 1]; }
    public SummaryChoiceFakeBook Open(string path, int links, bool readOnly)
    {
        Opens++; if (FailOpen) throw new IOException("injected open failure");
        var book = new SummaryChoiceFakeBook { Name = Path.GetFileName(path), FullName = path, ReadOnly = NextReadOnly };
        book.OnClose = delegate { Items.Remove(book); }; Items.Add(book); return book;
    }
}
public sealed class SummaryChoiceFakeApp
{
    public SummaryChoiceFakeBooks Workbooks { get; private set; }
    public object ActiveWorkbook { get; set; }
    public bool DisplayAlerts { get; set; }
    public bool AskToUpdateLinks { get; set; }
    public int AutomationSecurity { get; set; }
    public bool EnableEvents { get; set; }
    public bool ScreenUpdating { get; set; }
    public SummaryChoiceFakeApp()
    { Workbooks = new SummaryChoiceFakeBooks(); DisplayAlerts = AskToUpdateLinks = EnableEvents = ScreenUpdating = true; AutomationSecurity = 2; }
}
internal static class SummaryWorkbookBindingChecks
{
    internal static void Run(Action<string, bool, string> check)
    {
        var app = new SummaryChoiceFakeApp();
        var a = new SummaryChoiceFakeBook { Name = "A.xlsx", FullName = "C:\\one\\A.xlsx" };
        var b = new SummaryChoiceFakeBook { Name = "B.xlsx", FullName = "C:\\two\\B.xlsx" };
        app.Workbooks.Items.Add(a); app.Workbooks.Items.Add(b); app.ActiveWorkbook = a;
        Action<Action, string> reject = delegate(Action action, string title)
        {
            bool rejected = false;
            try { action(); } catch (InvalidOperationException) { rejected = true; }
            check(title, rejected, "未拒绝不明确的处理对象");
        };
        try
        {
            using (var binding = new SummaryWorkbookBinding(app))
            {
                reject(delegate { binding.RequireBook(); }, "活动工作簿不会成为默认目标");
                var choices = binding.Choices();
                check("列表含文件名和完整路径", choices.Count == 2 && choices[1].Name == b.Name && choices[1].Path == b.FullName, "清单错误");
                reject(delegate { binding.RequireBook(); }, "刷新列表不隐式绑定");
                binding.Choose(a.FullName); app.ActiveWorkbook = b;
                check("切换活动工作簿仍绑定原对象", ReferenceEquals(binding.RequireBook(), a), "被活动工作簿替换");
                binding.UseCurrent();
                check("主动使用当前工作簿才改变绑定", ReferenceEquals(binding.RequireBook(), b), "快捷操作未生效");
                long version = binding.Version; binding.Choose(a.FullName);
                check("重新选择使预检批次身份失效", binding.Version != version, "未失效");
                app.Workbooks.Items.Remove(a);
                reject(delegate { binding.RequireBook(); }, "所选工作簿关闭后禁止回退");
                var reopened = new SummaryChoiceFakeBook { Name = a.Name, FullName = a.FullName };
                app.Workbooks.Items.Add(reopened);
                reject(delegate { binding.RequireBook(); }, "同路径重新打开也要求明确重选");
                binding.Choose(reopened.FullName);
                check("明确重选后可使用新对象", ReferenceEquals(binding.RequireBook(), reopened), "无法重选");
                reopened.FullName = "C:\\changed.xlsx";
                reject(delegate { binding.RequireBook(); }, "另存为其他路径使绑定失效");
                b.ReadOnly = true;
                reject(delegate { binding.Choose(b.FullName); }, "已有只读工作簿拒绝且不关闭");
                check("拒绝只读不会关闭用户文件", !b.Closed, "关闭了用户文件"); b.ReadOnly = false;
                binding.Clear();
                check("浏览已打开文件复用且不重复打开", !binding.Browse(b.FullName) && app.Workbooks.Opens == 0 && ReferenceEquals(binding.RequireBook(), b), "重复打开");
                check("借用用户工作簿不会取得关闭所有权", !binding.OpenedByBrowse, "错误取得所有权");
                string path = Path.Combine(Path.GetTempPath(), "summary-choice-fixture.xlsx");
                check("浏览未打开文件明确打开并绑定", binding.Browse(path) && app.Workbooks.Opens == 1, "无法打开");
                var opened = (SummaryChoiceFakeBook)binding.RequireBook();
                check("仅本次新打开文件被标记为自有", binding.OpenedByBrowse, "缺少所有权标记");
                binding.Dispose();
                check("释放绑定清除打开所有权", !binding.OpenedByBrowse, "所有权未清除");
                check("关闭功能窗口保留已打开文件供保存", !opened.Closed && !b.Closed, "自动关闭了工作簿");
                app.Workbooks.NextReadOnly = true;
                reject(delegate { binding.Browse(Path.Combine(Path.GetTempPath(), "summary-readonly-fixture.xlsx")); }, "浏览到只读副本明确拒绝");
                check("只读拒绝及正常打开均恢复宿主状态", app.DisplayAlerts && app.AskToUpdateLinks && app.EnableEvents && app.ScreenUpdating && app.AutomationSecurity == 2, "状态未恢复");
                app.Workbooks.FailOpen = true;
                bool failed = false;
                try { binding.Browse(Path.Combine(Path.GetTempPath(), "summary-failed-fixture.xlsx")); } catch (IOException) { failed = true; }
                check("打开失败恢复宏安全和事件状态", failed && app.DisplayAlerts && app.AskToUpdateLinks && app.EnableEvents && app.ScreenUpdating && app.AutomationSecurity == 2, "失败未恢复");
            }
        }
        catch (Exception error) { check("工作簿绑定托管夹具", false, error.ToString()); }
    }
}
