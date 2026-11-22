using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static FormulaCapture CaptureFor(FormulaOperation operation)
        { return operation == FormulaOperation.Custom ? FormulaCapture.Values : operation >= FormulaOperation.Add ? FormulaCapture.Numbers : FormulaCapture.Formulas; }

        private static void RunFormula(FormulaOperation operation)
        {
            if (!TryBeginFormulaExecution()) return;
            IntPtr owner = IntPtr.Zero;
            try
            {
                if (!EnsureApplication())
                {
                    FormulaWindowThread.Notice("请先打开 WPS 表格工作簿并选中单元格。", owner, true);
                    return;
                }
                owner = FormulaWindowState.ResolveOwner(_app);
                if (FormulaInteraction.ActivatePending(operation)) return;
                FormulaInteraction.Open(new ComFormulaHost(_app, CaptureFor(operation)), operation, owner, _app);
            }
            catch (Exception ex)
            {
                UiDiagnostics.Record("formula-command", ex);
                FormulaWindowThread.Notice("一键公式发生错误：" + ex.Message, owner, true);
            }
            finally { EndFormulaExecution(); }
        }

        internal static bool TryBeginFormulaExecution() { if (_busy) return false; _busy = true; return true; }

        internal static void EndFormulaExecution() { _busy = false; }

    }
}
