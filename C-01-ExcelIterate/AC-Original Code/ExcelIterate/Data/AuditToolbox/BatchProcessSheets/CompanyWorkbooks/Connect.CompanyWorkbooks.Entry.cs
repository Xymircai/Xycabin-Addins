using System;
using System.Windows.Forms;

namespace Excel_Iterate
{
    public partial class Connect
    {
        private static CompanyWorkbookForm _companyWorkbookForm;
        internal static void CloseCompanyWorkbookForm()
        {
            if (_companyWorkbookForm != null && !_companyWorkbookForm.IsDisposed) _companyWorkbookForm.Shutdown();
        }
        public bool LegacyCompanyTemplateVisible(object control){return false;}
        public void OnCompanyWorkbooks(object control){OnConsolidatedWorkbooks(control);}
        public IPictureDisp GetImage_btnCompanyWorkbooks(object control) { return GetButtonImageCore("btnCompanyWorkbooks"); }
    }
}
