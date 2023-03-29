using System;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Uses the current windows user login to log the user into PDC using the PDCService.
    /// </summary>
    class WinLoginAction : PDCAction
    {
        public WinLoginAction(bool beginGroup)
            : base(Properties.Resources.Action_WinLogin_Caption, ACTION_TAG, beginGroup)
        {
        }
        public const string ACTION_TAG = "PDC_WinLoginAction";

        protected override bool SafeForCellEditingMode()
        {
            return true;
        }
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            try
            {
              // start auto updater
              Globals.PDCExcelAddIn.StartAutoUpdater();
              // check server version
             
//              if (Globals.PDCExcelAddIn.checkVersion(typeof(PDCExcelAddIn).Assembly.GetName().Version.ToString()))
              if (Globals.PDCExcelAddIn.checkVersion(VersionInfoAction.GetVersionNo()))
              {
                Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
                tmpService.Login(null, null, true);
                if (tmpService.UserInfo.HasRoles())
                  Globals.PDCExcelAddIn.LoggedInUser =   Globals.PDCExcelAddIn.PdcService.UserInfo;
                else
                {
                  LoginException tmpLE = new LoginException(PDCFaultMessage.LOGIN_FAILED, 99, null, null);
               
                  throw tmpLE;
                }
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
              Globals.PDCExcelAddIn.ResetLoggedIn();
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
            finally
            {
                Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
                Globals.PDCExcelAddIn.Application.EnableEvents = true;
            }
            if (Globals.PDCExcelAddIn.IsLoggedIn)
            {
                MessageBox.Show(Properties.Resources.MSG_LOGIN_TEXT, Properties.Resources.MSG_LOGIN_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return new ActionStatus();
        }
    }
}
