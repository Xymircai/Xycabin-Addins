using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnFormulaAsText(object control) { RunCellConversion(CellConversionKind.FormulaAsText); }
        public IPictureDisp GetImage_btnFormulaAsText(object control) { return GetButtonImageCore("btnFormulaAsText"); }
    }
}
