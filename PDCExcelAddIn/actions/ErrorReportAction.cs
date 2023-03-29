using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class ErrorReportAction:PDCAction
    {
        public ErrorReportAction(Office.CommandBarPopup aPopup, bool beginGroup)
            : base(aPopup, beginGroup)
        {
        }
        protected override void initialize(Microsoft.Office.Core.CommandBarButton aButton)
        {
            //aButton.Caption = "Error &Reports";
            aButton.Caption = Properties.Resources.Action_ErrorReport_Caption;
            aButton.Enabled = false;
        }
        internal override void PerformAction(SheetInfo aSheetInfo, Microsoft.Office.Core.CommandBarButton aButton, ref bool CancelDefault)
        {
            MessageBox.Show("Not implemented yet!");
        }
    }
}
