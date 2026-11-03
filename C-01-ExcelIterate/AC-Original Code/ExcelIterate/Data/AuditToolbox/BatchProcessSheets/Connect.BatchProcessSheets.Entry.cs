using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

// 功能：Excel-Iterate → 数据 → 审计工具箱 → 批量生成过程表
// 控件：menuBatchProcessSheets（父菜单目录只承载图标接线）
namespace Excel_Iterate
{
    public partial class Connect
    {
        public IPictureDisp GetImage_menuBatchProcessSheets(object control)
        { return GetButtonImageCore("menuBatchProcessSheets"); }
    }
}
