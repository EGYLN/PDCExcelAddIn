namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class LogoutAction:PDCAction
    {
        public LogoutAction(bool startGroup)
            : base(Properties.Resources.Action_Logout_Caption, ACTION_TAG, startGroup)
        {
        }
        public const string ACTION_TAG = "PDC_LogoutAction";

        protected override bool SafeForCellEditingMode()
        {
            return true;
        }

        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            Globals.PDCExcelAddIn.ResetLoggedIn();
            Globals.PDCExcelAddIn.PdcService.Logout();

            return new ActionStatus();
        }
    }
}
