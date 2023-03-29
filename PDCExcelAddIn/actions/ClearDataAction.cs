using System.Windows.Forms;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// This action removes all input data from the active PDC work sheet
    /// </summary>
    class ClearDataAction: PDCAction
    {
        public ClearDataAction(bool beginGroup)
            : base(Properties.Resources.Action_ClearData_Caption, ACTION_TAG, beginGroup)
        {
          myCommandBarText = Properties.Resources.Action_ClearData_BarTitle;
        }
        /// <summary>
        /// Action Tag for main menu button. Used by OpenLib
        /// </summary>
        public const string ACTION_TAG = "PDC_ClearDataAction";

        # region CanPerformAction
        /// <summary>
        /// checks if the current action can be executed.
        /// </summary>
        /// <param name="actionStatus">this action is being filled if the action cannot be performed</param>
        /// <param name="interactive">if interactive mode -> a messagebox is shown</param>
        /// <returns></returns>
        protected override bool CanPerformAction(out ActionStatus actionStatus, bool interactive)
        {
          actionStatus = null;
          if (!Globals.PDCExcelAddIn.AreAllSheetsForTheSelectedSheetAvailable())
          {
            if (interactive) MessageBox.Show(Properties.Resources.MSG_SHEET_IS_MISSING_TEXT, Properties.Resources.MSG_SHEET_IS_MISSING_TITLE);

            actionStatus = new ActionStatus(new Lib.PDCMessage[] {new Lib.PDCMessage(Properties.Resources.MSG_SHEET_IS_MISSING_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_SHEET_IS_MISSING_TEXT, Lib.PDCMessage.TYPE_ERROR)});
            return false;

          }
          return true;
        }
        # endregion

        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            CheckPDCSheet(sheetInfo);

            Lib.Testdefinition tmpTD = sheetInfo.TestDefinition;
            PDCListObject tmpPDCList = sheetInfo.MainTable;
            ExcelUtils.TheUtils.DeleteShapes(sheetInfo);
            tmpPDCList.ClearContents();
            return new ActionStatus();
        }
    }
}
