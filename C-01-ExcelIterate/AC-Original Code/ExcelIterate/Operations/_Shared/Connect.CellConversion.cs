using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        internal static string ConversionRecoveryDirectory { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Excel_Iterate", "ConversionRecovery"); } }
        internal static bool TryBeginConversion() { if (_busy) return false; _busy = true; return true; }
        internal static void EndConversion() { _busy = false; }
        private static IntPtr ConversionOwner()
        {
            object window = null;
            try { window = ((dynamic)_app).ActiveWindow; return HostDialogService.ValidateOwner(new IntPtr(System.Convert.ToInt64(((dynamic)window).Hwnd))); }
            catch { return IntPtr.Zero; }
            finally { if (window != null && Marshal.IsComObject(window)) Marshal.ReleaseComObject(window); }
        }
        private static void RunCellConversion(CellConversionKind kind)
        {
            if (!TryBeginConversion()) return;
            try
            {
                if (!EnsureApplication()) { UiDialog.Show("请先打开工作簿并选择单元格。", "一键转换", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                if (ConversionInteraction.ActivatePending(kind)) return;
                ComConversionHost host;
                try { host = new ComConversionHost(_app); }
                catch (Exception error)
                {
                    if (UndoState(ConversionInteraction.UndoSource(kind)) == null) throw;
                    ConversionInteraction.Open(null, kind, ConversionOwner(), _app, error.Message);
                    return;
                }
                ConversionInteraction.Open(host, kind, ConversionOwner(), _app);
            }
            catch (Exception error) { UiDiagnostics.Record("conversion-command", error); UiDialog.Show("转换未执行：" + error.Message, "一键转换", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { EndConversion(); }
        }
    }
}
