using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Excel_Iterate
{
    internal sealed class MigrationTargetScan
    {
        internal readonly List<string> Files = new List<string>(), Errors = new List<string>();
        internal int Skipped;
        internal static MigrationTargetScan Expand(IEnumerable<string> selections, string source)
        {
            var result = new MigrationTargetScan(); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string selected in selections)
            {
                try
                {
                    string path = Path.GetFullPath(selected);
                    if (Directory.Exists(path))
                    {
                        // Deliberately inspect only the selected folder, never its subfolders.
                        foreach (string file in Directory.EnumerateFiles(path).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
                        {
                            string name = Path.GetFileName(file), stem = Path.GetFileNameWithoutExtension(file);
                            if (!DataMigrationEngine.SupportedFile(file)) continue;
                            if (name.StartsWith("~$", StringComparison.Ordinal) || name.StartsWith(".xy-", StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(file, source, StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(stem, @"_(?:复制|迁移)(?:结果|修正结果)(?:_\d+)?$")) { result.Skipped++; continue; }
                            if (seen.Add(file)) result.Files.Add(file);
                        }
                    }
                    else if (File.Exists(path) && DataMigrationEngine.SupportedFile(path)) { if (seen.Add(path)) result.Files.Add(path); }
                    else result.Errors.Add("无法读取文件或文件夹：" + selected);
                }
                catch (Exception error) { result.Errors.Add(selected + "：" + error.Message); }
            }
            return result;
        }
    }

    // Windows' normal Open action navigates into a selected folder. A command
    // inside the SAME native picker accepts the selection instead, including
    // folders. Only strings cross HostDialogService's isolated UI STA boundary.
    public static class MigrationTargetPicker
    {
        internal static string[] Show(IWin32Window owner)
        {
            return HostDialogService.Run(owner, uiOwner => ShowNative(uiOwner), new string[0]);
        }
        private static string[] ShowNative(IWin32Window owner)
        {
            Application.OleRequired();
            object native = Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("DC1C5A9C-E88A-4DDE-A5A1-60F82A20AEF7")));
            var dialog = (OpenDialog)native; uint cookie = 0; var events = new PickerEvents(dialog);
            try
            {
                dialog.SetOptions(0x200 | 0x40 | 0x800 | 0x1000); // multi-select, filesystem, path/file must exist
                dialog.SetTitle("添加目标：选择工作簿或文件夹"); dialog.SetOkButtonLabel("添加文件");
                dialog.SetFileTypes(1, new[] { new Filter { Name = "工作簿与文件夹", Pattern = "*.xlsx;*.xlsm;*.xls;*.et" } });
                ((Customize)native).AddPushButton(1001, "添加选中项（文件或文件夹）");
                dialog.Advise(events, out cookie);
                int code = dialog.Show(owner == null ? IntPtr.Zero : owner.Handle);
                if (events.Error != null) throw new InvalidOperationException("选择目标失败：" + events.Error.Message, events.Error);
                if (code == unchecked((int)0x800704C7)) return new string[0];
                if (code < 0) Marshal.ThrowExceptionForHR(code);
                if (events.Paths != null) return events.Paths;
                ShellItems items; dialog.GetResults(out items); return Paths(items);
            }
            finally { try { if (cookie != 0) dialog.Unadvise(cookie); } finally { Marshal.ReleaseComObject(native); GC.KeepAlive(events); } }
        }
        private static string[] Paths(ShellItems items)
        {
            try
            {
                uint count; items.GetCount(out count); var paths = new List<string>();
                for (uint i = 0; i < count; i++)
                {
                    ShellItem item; items.GetItemAt(i, out item);
                    try { string path; item.GetDisplayName(0x80058000, out path); paths.Add(path); }
                    finally { Marshal.ReleaseComObject(item); }
                }
                return paths.ToArray();
            }
            finally { Marshal.ReleaseComObject(items); }
        }
        [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
        public sealed class PickerEvents : DialogEvents, ControlEvents
        {
            private readonly OpenDialog dialog; internal string[] Paths; internal Exception Error;
            internal PickerEvents(object value) { dialog = (OpenDialog)value; }
            public int OnButtonClicked(IntPtr customize, uint id)
            {
                if (id != 1001) return 0;
                try
                {
                    ShellItems items; dialog.GetSelectedItems(out items); Paths = MigrationTargetPicker.Paths(items);
                    if (Paths.Length == 0) return 0;
                    dialog.Close(0);
                }
                catch (Exception error) { Error = error; dialog.Close(unchecked((int)0x80004005)); }
                return 0;
            }
            public int OnFileOk(IntPtr dialog) { return 0; }
            public int OnFolderChanging(IntPtr dialog, IntPtr folder) { return 0; }
            public int OnFolderChange(IntPtr dialog) { return 0; }
            public int OnSelectionChange(IntPtr dialog) { return 0; }
            public int OnShareViolation(IntPtr dialog, IntPtr item, out uint response) { response = 0; return 0; }
            public int OnTypeChange(IntPtr dialog) { return 0; }
            public int OnOverwrite(IntPtr dialog, IntPtr item, out uint response) { response = 0; return 0; }
            public int OnItemSelected(IntPtr dialog, uint control, uint item) { return 0; }
            public int OnCheckButtonToggled(IntPtr dialog, uint control, bool value) { return 0; }
            public int OnControlActivating(IntPtr dialog, uint control) { return 0; }
        }
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct Filter { [MarshalAs(UnmanagedType.LPWStr)] public string Name, Pattern; }
        [ComImport, Guid("D57C7288-D4AD-4768-BE02-9D969532D960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface OpenDialog
        {
            [PreserveSig] int Show(IntPtr owner);
            void SetFileTypes(uint count, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] Filter[] filters);
            void SetFileTypeIndex(uint index); void GetFileTypeIndex(out uint index);
            void Advise([MarshalAs(UnmanagedType.Interface)] DialogEvents events, out uint cookie); void Unadvise(uint cookie);
            void SetOptions(uint flags); void GetOptions(out uint flags);
            void SetDefaultFolder(ShellItem folder); void SetFolder(ShellItem folder); void GetFolder(out ShellItem folder); void GetCurrentSelection(out ShellItem item);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string name); void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string name);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string title); void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string text);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string text); void GetResult(out ShellItem item);
            void AddPlace(ShellItem item, uint placement); void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string extension);
            void Close(int result); void SetClientGuid(ref Guid guid); void ClearClientData(); void SetFilter(IntPtr filter);
            void GetResults(out ShellItems items); void GetSelectedItems(out ShellItems items);
        }
        [ComImport, Guid("E6FDD21A-163F-4975-9C8C-A69F1BA37034"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface Customize
        {
            void EnableOpenDropDown(uint id); void AddMenu(uint id, [MarshalAs(UnmanagedType.LPWStr)] string label);
            void AddPushButton(uint id, [MarshalAs(UnmanagedType.LPWStr)] string label);
        }
        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ShellItem
        {
            void BindToHandler(IntPtr context, ref Guid handler, ref Guid iid, out IntPtr result); void GetParent(out ShellItem parent);
            void GetDisplayName(uint kind, [MarshalAs(UnmanagedType.LPWStr)] out string name); void GetAttributes(uint mask, out uint attributes);
            void Compare(ShellItem other, uint hint, out int order);
        }
        [ComImport, Guid("B63EA76D-1F85-456F-A19C-48159EFA858B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ShellItems
        {
            void BindToHandler(IntPtr context, ref Guid handler, ref Guid iid, out IntPtr result);
            void GetPropertyStore(uint flags, ref Guid iid, out IntPtr result); void GetPropertyDescriptionList(ref Guid key, ref Guid iid, out IntPtr result);
            void GetAttributes(uint flags, uint mask, out uint attributes); void GetCount(out uint count); void GetItemAt(uint index, out ShellItem item); void EnumItems(out IntPtr result);
        }
        [ComVisible(true), Guid("973510DB-7D7F-452B-8975-74A85828D354"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface DialogEvents
        {
            [PreserveSig] int OnFileOk(IntPtr dialog); [PreserveSig] int OnFolderChanging(IntPtr dialog, IntPtr folder);
            [PreserveSig] int OnFolderChange(IntPtr dialog); [PreserveSig] int OnSelectionChange(IntPtr dialog);
            [PreserveSig] int OnShareViolation(IntPtr dialog, IntPtr item, out uint response); [PreserveSig] int OnTypeChange(IntPtr dialog);
            [PreserveSig] int OnOverwrite(IntPtr dialog, IntPtr item, out uint response);
        }
        [ComVisible(true), Guid("36116642-D713-4B97-9B83-7484A9D00433"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ControlEvents
        {
            [PreserveSig] int OnItemSelected(IntPtr dialog, uint control, uint item); [PreserveSig] int OnButtonClicked(IntPtr dialog, uint control);
            [PreserveSig] int OnCheckButtonToggled(IntPtr dialog, uint control, [MarshalAs(UnmanagedType.Bool)] bool value);
            [PreserveSig] int OnControlActivating(IntPtr dialog, uint control);
        }
    }
}
