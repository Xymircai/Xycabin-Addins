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
                if (kind == CellConversionKind.TextFilter || kind == CellConversionKind.ReferenceStyle || kind == CellConversionKind.BatchInsertTrimDelete)
                {
                    if (ConversionInteraction.ActivatePending()) return;
                    ConversionInteraction.Open(new ComConversionHost(_app), kind, ConversionOwner());
                }
                else PresentConversionResult(CellConversionEngine.Run(new ComConversionHost(_app), new ConversionRequest { Kind = kind }, ConversionRecoveryDirectory));
            }
            catch (Exception error) { UiDiagnostics.Record("conversion-command", error); UiDialog.Show("转换未执行：" + error.Message, "一键转换", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
            finally { EndConversion(); }
        }
        internal static void PresentConversionResult(CellConversionResult result)
        {
            if (result.Cancelled) return;
            if (result.PlainResults + result.FormulasKept + result.FormulasAsText > 0 && !result.Failed) { ShowStatus(result.Details); return; }
            if (!result.Failed && result.Changed > 0) ShowStatus("转换：已处理 " + result.Changed + " 个单元格" + (result.FormulasKept + result.FormulasAsText > 0 ? "；保留公式 " + result.FormulasKept + " 个，公式转文本 " + result.FormulasAsText + " 个" : ""));
            else UiDialog.Show(result.Details, "转换结果", MessageBoxButtons.OK, result.Failed ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }
    }
}
