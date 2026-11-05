using System;
using System.Runtime.InteropServices;

namespace Excel_Iterate
{
    public partial class Connect
    {
        public void OnValuesOnly(object control) { RunCellConversion(CellConversionKind.ValuesOnly); }
        public IPictureDisp GetImage_btnValuesOnly(object control) { return GetButtonImageCore("btnValuesOnly"); }
    }
}
