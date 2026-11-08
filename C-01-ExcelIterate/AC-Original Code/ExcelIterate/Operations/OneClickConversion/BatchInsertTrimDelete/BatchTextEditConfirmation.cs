using System.Windows.Forms;

namespace Excel_Iterate
{
    internal static class BatchTextEditConfirmation
    {
        internal static UiMessageForm Create(BatchSyntaxReport report)
        {
            var form = new UiMessageForm(report.Message, "公式语法无效", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            form.Positive.Text = "继续并转为文本";
            form.Positive.Width = System.Math.Max((int)(160 * DialogSizing.Scale(form)), TextRenderer.MeasureText(form.Positive.Text, form.Positive.Font).Width + (int)(24 * DialogSizing.Scale(form)));
            return form;
        }
        internal static bool Show(BatchSyntaxReport report)
        {
            return HostDialogService.ShowIsolatedResult(null, delegate { return Create(report); }) == DialogResult.OK;
        }
    }
}
