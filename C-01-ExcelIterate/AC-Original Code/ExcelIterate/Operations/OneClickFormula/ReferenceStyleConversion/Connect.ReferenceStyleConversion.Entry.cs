using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnReferenceStyleConversion(object control) { RunCellConversion(CellConversionKind.ReferenceStyle); }
        public IPictureDisp GetImage_btnReferenceStyleConversion(object control) { return GetButtonImageCore("btnReferenceStyleConversion"); }
    }
}
