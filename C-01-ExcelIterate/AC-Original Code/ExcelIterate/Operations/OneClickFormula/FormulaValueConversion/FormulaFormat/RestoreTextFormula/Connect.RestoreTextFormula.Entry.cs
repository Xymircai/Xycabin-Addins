using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnRestoreTextFormula(object control) { RunCellConversion(CellConversionKind.RestoreTextFormula); }
        public IPictureDisp GetImage_btnRestoreTextFormula(object control) { return GetButtonImageCore("btnRestoreTextFormula"); }
    }
}
