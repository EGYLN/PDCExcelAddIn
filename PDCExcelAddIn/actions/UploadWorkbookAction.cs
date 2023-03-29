using System;
using System.Collections.Generic;
using BBS.ST.BHC.BSP.PDC.Lib;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Action which implements the upload of new test data
    /// </summary>
    class UploadWorkbookAction:PDCAction
  {
    private SheetInfo mySheetInfo;
    private Lib.Testdata myTestData;
    private bool mySuppressDialog = false;
    public const string ACTION_TAG = "PDC_UploadAction";

    #region constructor
    public UploadWorkbookAction(bool beginGroup)
      : base(Properties.Resources.Action_Upload_Workbook, ACTION_TAG, beginGroup)
    {
      myCommandBarText = Properties.Resources.Action_UploadWorkbook_BarTitle;
    }
    #endregion

    #region methods

    #region ActionCompleted
    private void ActionCompleted(object aResult, ProgressDialog aWindowOwner, bool interactive)
    {
      try
      {
        if (!(aResult is List<Lib.PDCMessage>))
        {
          return;
        }
        List<Lib.PDCMessage> tmpMessages = (List<Lib.PDCMessage>)aResult;
        if (interactive)
        {
          if (!mySheetInfo.MainTable.AlreadyUploaded)
          {
            MessageBox.Show(Properties.Resources.MSG_UPLOAD_FAILED_TEXT, Properties.Resources.MSG_UPLOAD_OK_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
          else
          {
            int numberOfExperiments = myTestData.Experiments.Count;
            foreach (object experiment in myTestData.Experiments)
            {
              if (experiment is Lib.PlaceHolderExperiment) numberOfExperiments--;
            }
            MessageBox.Show(string.Format(Properties.Resources.MSG_UPLOAD_OK_TEXT, numberOfExperiments, myTestData.NumberOfAllMeasurementValues), Properties.Resources.MSG_UPLOAD_OK_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
          }
        }
        if (tmpMessages.Count > 0)
        {
          mySheetInfo.MainTable.ValidationHandler.DisplayValidationMessages(tmpMessages, myTestData, interactive,true);
        }
      }
      finally
      {
        Globals.PDCExcelAddIn.Application.EnableEvents = true;
        Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
        Globals.PDCExcelAddIn.SetStatusText(null);
      }
    }
    #endregion

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
      /*
       * Feature 4459: No dialog box is shown, validation messages are displayed on upload instead.
       * This dialog could be used if a general confirmation dialog should be presented.
       * 
      if (interactive && MessageBox.Show(Properties.Resources.MSG_QUESTION_VERIFIED, Properties.Resources.MSG_CONFIRM_TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
      {
        return false;
      }
       */
      return true;
    }
    # endregion

    #region PerformAction
    internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
    {
      CheckPDCSheet(sheetInfo);
     
      bool filtered = ExcelUtils.TheUtils.IsAnyFilterOn(sheetInfo.MainTable.Container);
      if (filtered)
      {
        if (interactive) {
          MessageBox.Show(Properties.Resources.MSG_AUTO_FILTER_SET_UPLOAD, Properties.Resources.MSG_ERROR_TITLE,
                         MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return new ActionStatus(new Lib.PDCMessage[] {
              new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_AUTO_FILTER_SET_UPLOAD, Lib.PDCMessage.TYPE_ERROR)});
      }

      Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;
      bool tmpEnabled = Globals.PDCExcelAddIn.Application.EnableEvents;

      Globals.PDCExcelAddIn.Application.EnableEvents = false;
      mySheetInfo = sheetInfo;
      mySuppressDialog = !interactive;
      object tmpResult = ProgressDialog.Show(ActionCompleted, PerformUpload, Properties.Resources.LABEL_UPLOAD_WORKBOOK,false, interactive);
      return ResultToActionStatus(tmpResult, Properties.Resources.MSG_UPLOAD_TITLE, interactive);
    }
    #endregion

    #region PerformUpload
    private void PerformUpload(ProgressDialog aWindowOwner)
    {
      object tmpResult = null;
     
      Excel.Worksheet tmpSheet = mySheetInfo.ExcelSheet;
      Lib.Testdefinition tmpTD = mySheetInfo.TestDefinition;
      PDCListObject tmpPDCList = mySheetInfo.MainTable;
      ExcelFilterStatus tmpExcelFilterStatus = ExcelUtils.TheUtils.CollectExcelFilters(tmpPDCList);

      try
      {
        tmpPDCList.Container.AutoFilterMode = false;
        Globals.PDCExcelAddIn.Application.ScreenUpdating = false;


        tmpPDCList.ValidationHandler.ClearAllValidationMessages();
        bool[] tmpHidden = tmpPDCList.HiddenRows();
        myTestData = tmpPDCList.TestDataAdapter.GetTestData(true, tmpHidden, false, true, false);
        if (myTestData.IsEmpty())
        {
          tmpResult = Properties.Resources.MSG_NO_DATA_TO_UPLOAD_TEXT;
          return;
        }

        if (!tmpPDCList.ValidationHandler.Validate((ExperimentAndMeasurementValues)myTestData.Tag, tmpHidden))
        {
          tmpResult = Properties.Resources.MSG_VALIDATION_FAILED;
          return;
        }
          bool autoupdate = Autoupdate(tmpPDCList, tmpHidden, ref tmpResult);
    
        if (!autoupdate)
        {
          Globals.PDCExcelAddIn.SetStatusText("Uploading test data");
          List<Lib.PDCMessage> tmpMessages = Globals.PDCExcelAddIn.PdcService.UploadTestdata(myTestData);
          Globals.PDCExcelAddIn.SetStatusText("Updating worksheet");
          tmpPDCList.TestDataAdapter.UpdateFromUpload(myTestData);
          
          tmpResult = tmpMessages;
        }
      }
      catch (Exception e)
      {
        tmpResult = e;
      }
      finally
      {
         ExcelUtils.TheUtils.SetExcelFilters(tmpPDCList, tmpExcelFilterStatus);
         aWindowOwner.StatusCallback(tmpResult);
         Globals.PDCExcelAddIn.Application.ScreenUpdating = true;
      }
    }

      private bool Autoupdate(PDCListObject tmpPDCList, bool[] tmpHidden, ref object tmpResult)
      {
          bool autoupdate = false;
          if (tmpPDCList.Testdefinition.HasExperimentLevelVariables)
          {
              // before upload in case of SMT with exp level data, check for existing data in DB
              // create a copy of the test data to fill experimentnos in
              // TODO check if SMT with exp level parameters

              HashSet<decimal> experimentNos = new HashSet<decimal>();
              foreach (ExperimentData experiment in myTestData.Experiments)
              {
                  if (!(experiment is PlaceHolderExperiment) && experiment.ExperimentNo != null)
                  {
                      experimentNos.Add(experiment.ExperimentNo.Value);
                  }
              }

              Lib.Testdata newTestData = tmpPDCList.TestDataAdapter.GetTestData(true, tmpHidden, false, true, false);
              List<Lib.PDCMessage> tmpMessages = Globals.PDCExcelAddIn.PdcService.Autoupdate(myTestData, newTestData);
              // check for message and display decision dialog: Overwrite or upload, when overwrite use newTestData, when upload normal upload

              foreach (ExperimentData experiment in newTestData.Experiments)
              {
                  if (!(experiment is PlaceHolderExperiment) && experiment.ExperimentNo != null)
                  {
                      experimentNos.Remove(experiment.ExperimentNo.Value);
                  }
              }

              if (tmpMessages != null && tmpMessages.Count > 0)
              {
                  DialogResult result = DialogResult.OK;
                  if (!mySuppressDialog)
                  {
                      result =
                          MessageBox.Show(string.Format(Properties.Resources.MSG_AUTOUPDATE_VALUES, (object) tmpMessages[0].Message),
                              Properties.Resources.MSG_AUTOUPDATE_TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                  }
                  switch (result)
                  {
                      case DialogResult.Cancel:
                          autoupdate = true;
                          break;
                      case DialogResult.OK:
                          Globals.PDCExcelAddIn.SetStatusText("Updating test data");
                          tmpMessages = Globals.PDCExcelAddIn.PdcService.UploadChanges(newTestData, experimentNos);
                          Globals.PDCExcelAddIn.SetStatusText("Updating worksheet");
                          tmpPDCList.TestDataAdapter.UpdateFromUpload(newTestData);
                          tmpResult = tmpMessages;
                          autoupdate = true;
                          break;
                  }
              }
          }
          return autoupdate;
      }

      #endregion

    #endregion
  }
}
