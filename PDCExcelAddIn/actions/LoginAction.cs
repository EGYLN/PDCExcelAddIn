using System;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Opens a Login form to enter username and password and logs the user into PDC using the
    /// PDCService.
    /// </summary>
    class LoginAction:PDCAction
    {
        public const string ACTION_TAG = "PDC_LoginAction";

        public LoginAction(bool beginGroup)
            : base(Properties.Resources.Action_Login_Caption, ACTION_TAG, beginGroup)
        {
        }

        protected override bool SafeForCellEditingMode()
        {
          return true;
        }

        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {           

            PDCLoginForm tmpLoginForm = new PDCLoginForm();
            bool tmpCancelled = tmpLoginForm.ShowLogin();
            if (!tmpCancelled)
            {
              try
              {
                // start auto updater
                Globals.PDCExcelAddIn.StartAutoUpdater();
                // check server version

                if (Globals.PDCExcelAddIn.checkVersion(VersionInfoAction.GetVersionNo()))
                {
                   Lib.UserInfo tmpUser  = Globals.PDCExcelAddIn.PdcService.UserInfo;
                   Globals.PDCExcelAddIn.LoggedInUser = tmpUser;
                  // update privileges
                  Globals.PDCExcelAddIn.updatePrivileges();

                 
                }
                else
                {
                  MessageBox.Show(Properties.Resources.MSG_PDCVERSION_FAILED, Properties.Resources.MSG_PDCVERSION_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
              }
              catch (Exception e)
              {
                // catch SQLError from pdcv1 server (stored procedure not present)
                if (e.Message.Contains("SQL"))
                {
                  MessageBox.Show(Properties.Resources.MSG_PDCVERSION_FAILED + " (SQL Error)", Properties.Resources.MSG_PDCVERSION_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                  throw e;
                }
              }
            }
            return new ActionStatus();
        }

        private void StartAutoUpdater()
        {
          
        }
    }
}
