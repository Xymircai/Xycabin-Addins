using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// 功能：Excel-Iterate → 操作 → 一键公式 → 数值计算
// 控件：menuArithmetic
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuArithmetic(object control) { return GetButtonImageCore("menuArithmetic"); }
        public void OnArithmetic(object control) { RunFormula(FormulaOperation.Add); }
    }
}
