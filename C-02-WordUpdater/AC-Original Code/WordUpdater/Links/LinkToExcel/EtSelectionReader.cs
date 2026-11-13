using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace WordUpdater
{
    internal sealed class EtSelectionInfo
    {
        public List<LinkFieldCode.SourceArea> Areas = new List<LinkFieldCode.SourceArea>();
        public string WorkbookPath;
        public string SheetName;
    }

    // One entry for each acquired COM reference, even if the RCW is reused.
    // Selection reading borrows its Application; it never owns or closes WPS.
    internal sealed class EtReferenceScope : IDisposable
    {
        private readonly List<object> acquired = new List<object>();
        private readonly Func<object, bool> isCom;
        private readonly Action<object> release;
        private readonly Action<string> report;
        private bool disposed;

        internal EtReferenceScope() : this(Marshal.IsComObject,
            value => { Marshal.ReleaseComObject(value); },
            message => Connect.Log("ET选区引用释放失败：" + message)) { }

        internal EtReferenceScope(Func<object, bool> checkCom, Action<object> releaseReference, Action<string> diagnostic)
        { isCom = checkCom; release = releaseReference; report = diagnostic; }

        internal object Own(object value)
        { if (value != null) acquired.Add(value); return value; }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            for (int i = acquired.Count - 1; i >= 0; i--)
                try { if (isCom(acquired[i])) release(acquired[i]); }
                catch (Exception error) { try { report(error.Message); } catch { } }
            acquired.Clear();
        }
    }

    // Clipboard LINK codes may retain only a bounding box. Read all running ET
    // candidates and let the caller match workbook, sheet and selected areas.
    // ROT/GetActiveObject can return an existing RCW: release only acquisitions,
    // never FinalRelease, CreateInstance, Quit or process termination.
    internal static class EtSelectionReader
    {
        private static readonly string[] ProgIds = { "KET.Application", "et.Application", "Excel.Application" };
        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx context);

        internal static List<EtSelectionInfo> TryReadAll(string workbookPath, out string note)
        {
            note = null;
            List<EtSelectionInfo> result = new List<EtSelectionInfo>();
            string last = null;
            using (EtReferenceScope references = new EtReferenceScope())
            {
                List<object> apps = CollectEtApps(workbookPath, references);
                foreach (object app in apps)
                {
                    string why;
                    List<EtSelectionInfo> infos = ReadCandidatesFrom(app, out why, new EtReferenceScope());
                    if (infos.Count > 0) result.AddRange(infos);
                    else if (why != null) last = why;
                }
            }
            note = result.Count > 0 ? null : (last ?? "未找到运行中的 WPS 表格。");
            return result;
        }

        private static List<object> CollectEtApps(string workbookPath, EtReferenceScope references)
        {
            List<object> apps = new List<object>();
            List<string> guids = new List<string>();
            foreach (string id in ProgIds)
                try
                {
                    Type type = Type.GetTypeFromProgID(id);
                    if (type == null) continue;
                    string guid = type.GUID.ToString();
                    if (!guids.Contains(guid)) guids.Add(guid);
                }
                catch { }

            // A registered Application can be an empty hidden ET server. The
            // visible source may publish only its workbook path. Always check
            // that exact moniker, even when an Application was already found.
            try { EnumerateRot(apps, guids, workbookPath, references); } catch { }
            foreach (string id in ProgIds)
                try { AddApp(apps, Marshal.GetActiveObject(id), references); }
                catch { }
            return apps;
        }

        private static void AddApp(List<object> apps, object value, EtReferenceScope references)
        {
            if (value == null) return;
            references.Own(value); // Duplicate discovery still acquired a reference.
            AddUnique(apps, value);
        }

        private static void AddUnique(List<object> apps, object value)
        {
            if (value == null) return;
            foreach (object existing in apps) if (ReferenceEquals(existing, value)) return;
            apps.Add(value);
        }

        private static bool SameWorkbookPath(string first, string second)
        {
            string normalized = LinkPasteEngine.NormPath(first);
            return normalized.Length > 0
                && string.Equals(normalized, LinkPasteEngine.NormPath(second), StringComparison.OrdinalIgnoreCase);
        }

        // The workbook is borrowed from the ROT item scope. Only the returned
        // Application acquisition is retained by the outer candidate scope.
        private static bool AddWorkbookApplication(List<object> apps, object workbook, string workbookPath, EtReferenceScope references)
        {
            if (workbook == null) return false;
            try
            {
                dynamic book = workbook;
                if (!SameWorkbookPath((string)book.FullName, workbookPath)) return false;
                object application = (object)book.Application;
                AddApp(apps, application, references);
                return application != null;
            }
            catch { return false; }
        }

        private static void EnumerateRot(List<object> apps, List<string> guids, string workbookPath, EtReferenceScope applications)
        {
            using (EtReferenceScope rotReferences = new EtReferenceScope())
            {
                IRunningObjectTable rot = null;
                int status;
                try { status = GetRunningObjectTable(0, out rot); }
                finally { rotReferences.Own(rot); }
                if (status != 0 || rot == null) return;
                IEnumMoniker enumerator = null;
                try { rot.EnumRunning(out enumerator); }
                finally { rotReferences.Own(enumerator); }
                if (enumerator == null) return;

                IMoniker[] item = new IMoniker[1];
                while (true)
                    using (EtReferenceScope itemReferences = new EtReferenceScope())
                    {
                        item[0] = null;
                        int next;
                        try { next = enumerator.Next(1, item, IntPtr.Zero); }
                        finally { itemReferences.Own(item[0]); }
                        if (next != 0 || item[0] == null) break;
                        IMoniker moniker = item[0];
                        string name = null;
                        try
                        {
                            IBindCtx context = null;
                            int created;
                            try { created = CreateBindCtx(0, out context); }
                            finally { itemReferences.Own(context); }
                            if (created == 0 && context != null) moniker.GetDisplayName(context, null, out name);
                        }
                        catch { }
                        bool sourceWorkbook = SameWorkbookPath(name, workbookPath), applicationMoniker = false;
                        if (name != null)
                            foreach (string guid in guids)
                                if (string.Equals(name, "!{" + guid + "}", StringComparison.OrdinalIgnoreCase))
                                { applicationMoniker = true; break; }
                        // Never bind unrelated documents or arbitrary ROT objects.
                        if (!sourceWorkbook && !applicationMoniker) continue;
                        object application = null;
                        try
                        {
                            try { rot.GetObject(moniker, out application); }
                            finally { if (sourceWorkbook) itemReferences.Own(application); else applications.Own(application); }
                            if (sourceWorkbook)
                            {
                                if (AddWorkbookApplication(apps, application, workbookPath, applications))
                                    Connect.Log("链接Excel：已通过来源工作簿的 ROT 登记定位表格实例。");
                            }
                            else AddUnique(apps, application);
                        }
                        catch { }
                    }
            }
        }

        private static EtSelectionInfo ReadFrom(object application, out string note)
        { return ReadFrom(application, out note, new EtReferenceScope()); }

        // The injected scope exercises the actual read/early-return lifecycle in
        // offline tests. The passed Application is borrowed and is not registered.
        private static EtSelectionInfo ReadFrom(object application, out string note, EtReferenceScope references)
        {
            List<EtSelectionInfo> infos = ReadCandidatesFrom(application, out note, references);
            return infos.Count > 0 ? infos[0] : null;
        }

        private static List<EtSelectionInfo> ReadCandidatesFrom(object application, out string note, EtReferenceScope references)
        {
            note = null;
            using (references)
            {
                if (application == null) { note = "表格实例为空。"; return new List<EtSelectionInfo>(); }
                dynamic app = application;
                object selection = null;
                try { selection = references.Own((object)app.Selection); } catch { }
                EtSelectionInfo direct = ReadSelection(selection, out note, references);
                if (direct != null) return new List<EtSelectionInfo> { direct };

                // WPS 文字功能区回调内，ET 的 Application.Selection 可能为空。
                // 明确读取已打开工作簿各窗口保存的选区，不 Activate/Select，不猜包围盒。
                string primaryNote = note, windowNote;
                List<EtSelectionInfo> windows = ReadWindowSelections(application, references, out windowNote);
                note = windows.Count > 0 ? null : primaryNote + "；" + windowNote;
                return windows;
            }
        }

        private static List<EtSelectionInfo> ReadWindowSelections(object application, EtReferenceScope references, out string note)
        {
            var result = new List<EtSelectionInfo>();
            note = "未能从已打开工作簿窗口读取多区域选区。";
            try
            {
                dynamic books = references.Own((object)((dynamic)application).Workbooks);
                int bookCount = (int)books.Count;
                for (int b = 1; b <= bookCount; b++)
                {
                    try
                    {
                        dynamic book = references.Own((object)WordComCollection.Item(books, b));
                        dynamic windows = references.Own((object)book.Windows);
                        int windowCount = (int)windows.Count;
                        for (int w = 1; w <= windowCount; w++)
                        {
                            try
                            {
                                dynamic window = references.Own((object)WordComCollection.Item(windows, w));
                                object selected = null;
                                try { selected = references.Own((object)window.RangeSelection); } catch { }
                                string why;
                                EtSelectionInfo info = ReadSelection(selected, out why, references);
                                if (info == null)
                                {
                                    selected = null;
                                    try { selected = references.Own((object)window.Selection); } catch { }
                                    info = ReadSelection(selected, out why, references);
                                }
                                if (info != null) result.Add(info);
                                else if (why != null) note = why;
                            }
                            catch (Exception error) { note = "读取工作簿窗口选区失败：" + error.Message; }
                        }
                    }
                    catch (Exception error) { note = "读取工作簿窗口集合失败：" + error.Message; }
                }
            }
            catch (Exception error) { note = "读取已打开工作簿集合失败：" + error.Message; }
            if (result.Count > 0) note = null;
            return result;
        }

        private static EtSelectionInfo ReadSelection(object value, out string note, EtReferenceScope references)
        {
                note = null;
                if (value == null) { note = "WPS 表格无可用选区。"; return null; }
                dynamic selection = value;
                string address = null;
                try { address = (string)selection.Address; } catch { }
                if (address == null) { note = "当前选区不是单元格区域。"; return null; }

                dynamic areas = null;
                int count = 1;
                try
                {
                    areas = references.Own((object)selection.Areas);
                    count = (int)areas.Count;
                }
                catch { count = 1; }
                if (count <= 1) { note = "选区只有一片，无需逐片明细。"; return null; }

                EtSelectionInfo result = new EtSelectionInfo();
                for (int i = 1; i <= count; i++)
                {
                    dynamic area = null;
                    try { area = references.Own((object)WordComCollection.Item(areas, i)); } catch { }
                    if (area == null) { note = "读取子区域 " + i + " 失败。"; return null; }
                    string areaAddress = null, sheetName = null, workbookPath = null;
                    dynamic sheet = null, book = null;
                    try { areaAddress = (string)area.Address; } catch { }
                    try { sheet = references.Own((object)area.Worksheet); } catch { }
                    if (sheet != null)
                    {
                        try { sheetName = (string)sheet.Name; } catch { }
                        try { book = references.Own((object)sheet.Parent); } catch { }
                    }
                    if (book != null) try { workbookPath = (string)book.FullName; } catch { }
                    if (areaAddress == null) { note = "子区域 " + i + " 无地址。"; return null; }
                    if (result.Areas.Count > 0 && (!string.Equals(result.WorkbookPath, workbookPath, StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(result.SheetName, sheetName, StringComparison.OrdinalIgnoreCase)))
                    { note = "多个子区域不属于同一工作簿和工作表。"; return null; }
                    result.WorkbookPath = workbookPath;
                    result.SheetName = sheetName;
                    LinkFieldCode.SourceArea parsed = ParseAddress(areaAddress, sheetName);
                    if (parsed == null) { note = "无法解析地址 " + areaAddress + "。"; return null; }
                    result.Areas.Add(parsed);
                }
                if (result.Areas.Count < 2) { note = "未取到多个子区域。"; return null; }
                return result;
        }

        private static LinkFieldCode.SourceArea ParseAddress(string address, string sheetName)
        {
            List<LinkFieldCode.SourceArea> areas = new List<LinkFieldCode.SourceArea>();
            string error;
            if (!LinkFieldCode.TrySplitItem(address, areas, out error) || areas.Count != 1) return null;
            LinkFieldCode.SourceArea area = areas[0];
            area.SheetPart = string.IsNullOrEmpty(sheetName) ? "" : sheetName + "!";
            return area;
        }
    }
}
