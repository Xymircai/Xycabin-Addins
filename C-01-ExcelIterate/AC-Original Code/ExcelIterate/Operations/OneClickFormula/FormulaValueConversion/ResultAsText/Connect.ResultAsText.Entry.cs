using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnResultAsText(object control) { RunCellConversion(CellConversionKind.ResultAsText); }
        public IPictureDisp GetImage_btnResultAsText(object control) { return GetButtonImageCore("btnResultAsText"); }
    }
}
