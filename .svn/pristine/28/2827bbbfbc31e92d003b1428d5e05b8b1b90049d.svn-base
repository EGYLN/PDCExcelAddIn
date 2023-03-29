using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class ValidationAction:PDCAction
  {
    public ValidationAction(bool beginGroup) : base(Properties.Resources.Action_Validate_Caption, ACTION_TAG, beginGroup)
    {
    }

    public const string ACTION_TAG = "PDC_ValidationAction";
    private SheetInfo sheetInfo;
    private Lib.Testdata testdata;

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
      //Get Active Sheet
      CheckPDCSheet(sheetInfo);
      Globals.PDCExcelAddIn.Application.Cursor = Excel.XlMousePointer.xlWait;
      try
      {
        this.sheetInfo = sheetInfo;
        object tmpResult = ProgressDialog.Show(ActionCompleted, PerformValidation, Properties.Resources.LABEL_VALIDATE_WORKBOOK, true, interactive);
        return ResultToActionStatus(tmpResult, Properties.Resources.MSG_VALIDATE_TITLE, interactive);
      }
      finally
      {                
        Globals.PDCExcelAddIn.Application.Cursor = Excel.XlMousePointer.xlDefault;
        Globals.PDCExcelAddIn.SetStatusText(null);
      }
    }

    private void PerformValidation(ProgressDialog aWindowOwner)
    {
      object tmpResult = null;
      try
      {
        Globals.PDCExcelAddIn.Application.ScreenUpdating = false;
        Excel.Worksheet tmpActiveSheet = sheetInfo.ExcelSheet;
        //Get Testdefinition for Active Sheet
        Lib.Testdefinition tmpTD = sheetInfo.TestDefinition;
        //Get ListObject for Testdefinition
        PDCListObject tmpList = sheetInfo.MainTable;
        tmpList.ValidationHandler.ClearAllValidationMessages();
        bool[] tmpHidden = tmpList.HiddenRows();
        testdata = tmpList.TestDataAdapter.GetTestData(true, tmpHidden,true,true, false);
        if (testdata.IsEmpty())
        {
          tmpResult = Properties.Resources.MSG_NO_DATA_TO_VALIDATE_TEXT;
          return;
        }
        bool tmpValidationStatus = tmpList.ValidationHandler.Validate((ExperimentAndMeasurementValues)testdata.Tag, null);
        if (tmpValidationStatus)
        {
          Globals.PDCExcelAddIn.SetStatusText("Server-side validation");
          tmpResult = PDCService.ValidateTestdata(testdata);
        }
        else
        {
          tmpResult = Properties.Resources.MSG_VALIDATION_FAILED;
        }
      }
      catch (Exception e)
      {
        tmpResult = e;
      }
      finally
      {
        aWindowOwner.StatusCallback(tmpResult);
        Globals.PDCExcelAddIn.Application.ScreenUpdating = true;
      }

    }

    private void ActionCompleted(object aResult, ProgressDialog aWindowOwner, bool interactive)
    {
      if (!(aResult is List<Lib.PDCMessage>))
      {
        return;
      }
      aWindowOwner.CanCancel = false;
      List<Lib.PDCMessage> tmpMessages = (List<Lib.PDCMessage>)aResult;

      if (tmpMessages != null && tmpMessages.Count != 0)
      {
        sheetInfo.MainTable.ValidationHandler.DisplayValidationMessages(tmpMessages, testdata, interactive, true);
      }
      Globals.PDCExcelAddIn.SetStatusText("Updating the worksheet");
      sheetInfo.MainTable.TestDataAdapter.UpdateFromUploadWithoutProcessingBinaryLinks(testdata);

    }
  }
}
