using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Diagnostic launcher only. Uses the production PDF windows without installing VSTO.
internal static class PdfTrialLauncher
{
    private const BindingFlags All = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
#if WPS_WORD
    private const string Host = "WordUpdater", Product = "Word PDF 集成验证版";
#else
    private const string Host = "Excel_Iterate", Product = "Excel PDF 集成验证版";
#endif
    private static Assembly assembly;
    private static readonly Dictionary<string, Form> forms = new Dictionary<string, Form>();
    private static object FindApplication()
    {
#if WPS_WORD
        string[] ids = { "KWPS.Application", "WPS.Application", "Word.Application" };
#else
        string[] ids = { "KET.Application", "et.Application", "Excel.Application" };
#endif
        foreach (string id in ids) try { return Marshal.GetActiveObject(id); } catch (COMException) { }
        throw new InvalidOperationException("请先在 WPS／Office 中打开需要导出的文档，再点击当前文档入口。");
    }
    private static string BindCurrent(Form form)
    {
        object app = null, session = null;
        try
        {
            app = FindApplication();
            session = Activator.CreateInstance(assembly.GetType(Host + ".PdfCurrentDocument", true), All, null, new[] { app }, null);
            string name = (string)session.GetType().GetProperty("DisplayName", All).GetValue(session, null);
            Type type = form.GetType(); FieldInfo export = type.GetField("ExportCurrent", All);
            object ownedApp = app, ownedSession = session;
            type.GetField("CloseCurrent", All).SetValue(form, new Action(delegate { try { ((IDisposable)ownedSession).Dispose(); } finally { if (Marshal.IsComObject(ownedApp)) Marshal.ReleaseComObject(ownedApp); } }));
            export.SetValue(form, Delegate.CreateDelegate(export.FieldType, session, session.GetType().GetMethod("Export", All)));
            return name;
        }
        catch { if (session != null) ((IDisposable)session).Dispose(); if (app != null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); throw; }
    }
    private static Form Create(string mode)
    {
        Form form = (Form)Activator.CreateInstance(assembly.GetType(Host + ".PdfToolsForm", true), All, null, new object[] { mode }, null);
        Type type = form.GetType();
        type.GetField("BindCurrent", All).SetValue(form, new Func<string>(() => BindCurrent(form)));
        if (mode == "current")
        {
            try { type.GetMethod("SetCurrentSource", All).Invoke(form, new object[] { BindCurrent(form) }); }
            catch (Exception error) { type.GetMethod("ShowCurrentError", All).Invoke(form, new object[] { (error.InnerException ?? error).Message }); }
        }
        return form;
    }
    private static void Open(string mode)
    {
        string key = mode == "current" || mode == "convert" ? "convert" : "edit";
        Form current; if (forms.TryGetValue(key, out current) && !current.IsDisposed) { current.GetType().GetMethod("SelectMode", All).Invoke(current, new object[] { mode }); current.Activate(); return; }
        try { var form = Create(mode); forms[key] = form; form.FormClosed += delegate { forms.Remove(key); form.Dispose(); }; form.Show(); }
        catch (Exception error) { MessageBox.Show((error.InnerException ?? error).Message, Product); }
    }
    [STAThread] private static int Main(string[] args)
    {
        try
        {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            assembly = Assembly.LoadFrom(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Host + ".dll"));
            if (args.Length > 0 && args[0] == "--check")
            {
                foreach (string mode in new[] { "convert", "watermark", "split", "merge" }) using (Form form = Create(mode)) if (form.Controls.Count == 0) throw new Exception("Empty production form");
                return 0;
            }
            using (var home = new Form { Text = Product, ClientSize = new Size(620, 260), StartPosition = FormStartPosition.CenterScreen, MinimumSize = new Size(580, 300), Font = new Font("Microsoft YaHei UI", 10), BackColor = Color.FromArgb(245, 247, 251) })
            {
                var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(20), ColumnCount = 2, RowCount = 3 };
                root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
                root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 90));
                var title = new Label { Text = "PDF 工具", Dock = DockStyle.Fill, Font = new Font(home.Font, FontStyle.Bold) }; root.Controls.Add(title, 0, 0); root.SetColumnSpan(title, 2);
                string[] modes = { "current", "watermark" }, labels = { "转为 PDF", "编辑 PDF" };
                for (int i = 0; i < modes.Length; i++) { string mode = modes[i]; var button = new Button { Text = labels[i], Dock = DockStyle.Fill, Margin = new Padding(5), BackColor = Color.White, FlatStyle = FlatStyle.Flat }; button.Click += delegate { Open(mode); }; root.Controls.Add(button, i % 2, i / 2 + 1); }
                var note = new Label { Text = "独立验证程序，不是 WPS 安装包。\r\n转为 PDF：当前文档、批量文件；编辑 PDF：水印、拆分、合并。\r\n当前文档导出前保存修改；新文件输出，同名自动编号。", Dock = DockStyle.Fill, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft }; root.Controls.Add(note, 0, 2); root.SetColumnSpan(note, 2); home.Controls.Add(root);
                home.FormClosed += delegate { foreach (var form in new List<Form>(forms.Values)) form.Dispose(); forms.Clear(); };
                Application.Run(home);
            }
            return 0;
        }
        catch (Exception error) { if (args.Length == 0) MessageBox.Show((error.InnerException ?? error).Message, Product); return 1; }
    }
}
