using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnTextExtractFilter(object control) { RunCellConversion(CellConversionKind.TextFilter); }
        public IPictureDisp GetImage_btnTextExtractFilter(object control) { return GetButtonImageCore("btnTextExtractFilter"); }
    }
}
