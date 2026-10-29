using System;
using System.IO;
using System.Runtime.InteropServices;

#if WPS_WORD
namespace WordUpdater
#else
namespace Excel_Iterate
#endif
{
    // Created, saved, exported and disposed only on the host STA. No COM crosses to the worker.
    internal sealed class PdfCurrentDocument : IDisposable
    {
        private readonly object app;
        private readonly int ownerThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
        private object document, sheet, selection;
        private string originalName;
        internal PdfCurrentDocument(object application)
        {
            if (application == null) throw new InvalidOperationException("请从已打开的文档中使用此功能。");
            app = application;
            try
            {
#if WPS_WORD
                document = ((dynamic)app).ActiveDocument;
#else
                document = ((dynamic)app).ActiveWorkbook;
#endif
                if (document == null) throw new InvalidOperationException("请先打开文档。");
                originalName = Convert.ToString(((dynamic)document).FullName);
#if !WPS_WORD
                sheet = ((dynamic)app).ActiveSheet;
                selection = ((dynamic)app).Selection;
#endif
            }
            catch { Dispose(); throw; }
        }
        internal string DisplayName { get { return originalName; } }
        private static bool Same(object a, object b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null || !Marshal.IsComObject(a) || !Marshal.IsComObject(b)) return false;
            IntPtr left = IntPtr.Zero, right = IntPtr.Zero;
            try { left = Marshal.GetIUnknownForObject(a); right = Marshal.GetIUnknownForObject(b); return left == right; }
            finally { if (left != IntPtr.Zero) Marshal.Release(left); if (right != IntPtr.Zero) Marshal.Release(right); }
        }
        private void Verify()
        {
            if (System.Threading.Thread.CurrentThread.ManagedThreadId != ownerThread) throw new InvalidOperationException("当前文档必须在原宿主线程导出。");
            if (document == null) throw new InvalidOperationException("原文档会话已结束，请重新打开导出窗口。");
            object active = null;
            try
            {
#if WPS_WORD
                active = ((dynamic)app).ActiveDocument;
#else
                active = ((dynamic)app).ActiveWorkbook;
#endif
                if (!Same(active, document)) throw new InvalidOperationException("活动文档已改变，请返回打开窗口时的文档，或重新打开导出窗口。");
            }
            finally { Release(active); }
        }
        internal string Export(PdfToolRequest request)
        {
            Verify();
            dynamic doc = document;
            if (Convert.ToBoolean(doc.ReadOnly)) throw new InvalidOperationException("当前文档为只读，请先另存为可编辑文件。");
            if (!string.Equals(Convert.ToString(doc.FullName), originalName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("文档路径已改变，请重新打开导出窗口。");
#if !WPS_WORD
            if (request.Scope != "book" && request.Scope != "sheet" && request.Scope != "selection") throw new ArgumentException("未知导出范围。");
            if (request.Scope == "sheet" || request.Scope == "selection")
            {
                object parent = null, sourceSheet = null, areas = null;
                try
                {
                    if (request.Scope == "selection")
                    {
                        sourceSheet = ((dynamic)selection).Worksheet;
                        areas = ((dynamic)selection).Areas;
                        if (Convert.ToInt32(((dynamic)areas).Count) != 1) throw new InvalidOperationException("选区导出仅支持一个连续的单元格区域，请重新选择后打开窗口。");
                        parent = ((dynamic)sourceSheet).Parent;
                    }
                    else parent = ((dynamic)sheet).Parent;
                    if (!Same(parent, document)) throw new InvalidOperationException("打开窗口时的工作表或选区已失效，请重新打开窗口。");
                }
                finally { Release(areas); Release(parent); Release(sourceSheet); }
            }
#endif
            // The host owns its native Save/Save As UI. Cancellation must leave Saved false
            // or FullName without a filesystem path; both stop the export below.
            doc.Save(); Verify();
            string path = Convert.ToString(doc.FullName);
            if (!Convert.ToBoolean(doc.Saved) || !Path.IsPathRooted(path) || !File.Exists(path))
                throw new OperationCanceledException("保存未完成，已停止导出 PDF。");
            originalName = path;
            string folder = Path.Combine(Path.GetTempPath(), "XyCurrentPdf-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(folder); string temp = Path.Combine(folder, "document.pdf");
            try
            {
#if WPS_WORD
                doc.ExportAsFixedFormat(temp, 17, false);
#else
                dynamic target = request.Scope == "book" ? document : request.Scope == "sheet" ? sheet : selection;
                target.ExportAsFixedFormat(0, temp, 0, true, request.Scope == "selection", Type.Missing, Type.Missing, false);
#endif
                if (!File.Exists(temp) || new FileInfo(temp).Length == 0) throw new IOException("宿主没有生成 PDF。");
                request.Mode = "publish"; request.Files = new[] { temp }; request.Name = Path.GetFileNameWithoutExtension(path);
                return folder;
            }
            catch
            {
                try { Directory.Delete(folder, true); }
                catch (Exception cleanup) { System.Diagnostics.Trace.WriteLine("当前文档导出临时目录清理失败：" + folder + "：" + cleanup); }
                throw;
            }
        }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
        public void Dispose() { Release(selection); selection = null; Release(sheet); sheet = null; Release(document); document = null; }
    }
}
