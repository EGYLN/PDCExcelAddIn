using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// This action is responsible for the deletion of the selected rows on the server
    /// </summary>
    class DeleteAction:PDCAction
  {
    private Lib.Testdefinition testDefinition;
    private List<decimal> experimentNos;
    
    private Excel.Range mySelectedRange;
    private Lib.Testdata myTestdata; // testdata of the deleted items!
    private PDCListObject pdcList;
//    private UniqueExperimentKeyHandler myUniqueExperimentKeyHandler;

    public DeleteAction(bool beginGroup) : base(Properties.Resources.Action_Delete_Caption, ACTION_TAG, beginGroup)
    {
    }

    public const string ACTION_TAG = "PDC_DeleteAction";


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

    internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactiveMode)
    {
      CheckPDCSheet(sheetInfo);
      pdcList = sheetInfo.MainTable;
      Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlWait;

      try
      {
        object tmpSelection = Globals.PDCExcelAddIn.Application.Selection;
        if (!(tmpSelection is Excel.Range))
        {
          if (interactiveMode)
          {
              MessageBox.Show(Properties.Resources.MSG_NO_CELLS_SELECTED, Properties.Resources.MSG_ERROR_TITLE,
                          MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
          return new ActionStatus(new Lib.PDCMessage[] {
            new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
            new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_ERROR)
          });
        }

        mySelectedRange = (Excel.Range)tmpSelection;
        if (mySelectedRange.Count == 0)
        {
          if (interactiveMode)
          {
            MessageBox.Show(Properties.Resources.MSG_NO_CELLS_SELECTED, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
          return new ActionStatus(new Lib.PDCMessage[] {
            new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
            new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_ERROR)
          });
        }

        // If there are hidden rows selected, it's not possible to delete.
        foreach (Excel.Range row in mySelectedRange.Rows)
        {
          if ((bool) row.Hidden)
          {
            // if hidden rows obviously caused through 
            bool filtered = ExcelUtils.TheUtils.IsAnyFilterOn(pdcList.Container);
            if (interactiveMode)
            {
              if (filtered)
              {
                MessageBox.Show(Properties.Resources.MSG_AUTO_FILTER_SET_DELETE, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK);
              }
              else
              {
                MessageBox.Show(Properties.Resources.MSG_HIDDEN_ROWS_SELECTED, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK);
              }
            }
            if (filtered)
            {
                return new ActionStatus(new Lib.PDCMessage[] {
              new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_AUTO_FILTER_SET_DELETE, Lib.PDCMessage.TYPE_ERROR)});
            }
            return new ActionStatus(new Lib.PDCMessage[] {
              new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_HIDDEN_ROWS_SELECTED, Lib.PDCMessage.TYPE_ERROR)});
            }
        }

        // Get Selected Rows
        int? tmpColumn = pdcList.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO);
        if (tmpColumn == null)
        {
          if (interactiveMode)
          {
            MessageBox.Show(Properties.Resources.MSG_NO_EXPERIMENTNO_ROW, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
          return new ActionStatus(new Lib.PDCMessage[] {
            new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
            new Lib.PDCMessage(Properties.Resources.MSG_NO_EXPERIMENTNO_ROW, Lib.PDCMessage.TYPE_ERROR)
          });
        }
      
        
        
        // this next two line are for checking the unique keys between singlemeasurement and main sheet
        if (pdcList.HasMeasurementParamHandler && pdcList.MeasurementColumn.HasSingleMeasurementTableHandler)
        {
          bool[] tmpHidden = pdcList.HiddenRows();
          myTestdata = pdcList.TestDataAdapter.GetTestData(true, tmpHidden, false, true, false);
          myTestdata = null;
        }
        Dictionary<int, int> rows = ExcelUtils.TheUtils.SelectedRows(mySelectedRange); 
        bool[] tmpTakeOrLeaveFlags = CreateLeaves(pdcList.DataRange.Row, pdcList.DataRange.Rows.Count, rows);
        
        myTestdata = pdcList.TestDataAdapter.GetTestData(true, tmpTakeOrLeaveFlags,true,false, true);


        List<decimal> tmpExperimentNos = GetExperimentNos(rows);


        if (interactiveMode)
        {
         DialogResult tmpResult = MessageBox.Show(string.Format(Properties.Resources.MSG_CONFIRM_DELETE_DATA, tmpExperimentNos.Count), 
            Properties.Resources.MSG_CONFIRM_TITLE, MessageBoxButtons.YesNo);
          if (DialogResult.No == tmpResult)
          {
            return new ActionStatus();
          }
        }
        testDefinition = pdcList.Testdefinition;
        experimentNos = tmpExperimentNos;

        object tmpActionResult = ProgressDialog.Show(ActionCompleted, PerformDelete,
                      string.Format(Properties.Resources.LABEL_DELETE_DATA,   experimentNos.Count), false, interactiveMode);
        if (tmpActionResult is Exception)
        {
          return new ActionStatus((Exception)tmpActionResult);
        }
      }
      catch (Exception e)
      {
        throw e;
      }
      finally
      {
        Globals.PDCExcelAddIn.Application.Cursor = Microsoft.Office.Interop.Excel.XlMousePointer.xlDefault;
      }
      return new ActionStatus();
    }

    private void PerformDelete(ProgressDialog aWindowOwner)
    {
      Exception tmpException = null;
      try
      {

        PDCService.DeleteTestdata(testDefinition, experimentNos);
        
        if (mySelectedRange != null && pdcList != null && pdcList.Container != null)
        {
          ExcelFilterStatus tmpExcelFilterStatus = ExcelUtils.TheUtils.CollectExcelFilters(pdcList);

          pdcList.Container.AutoFilterMode = false;
          ExcelUtils.TheUtils.DeleteRows(mySelectedRange);
          ExcelUtils.TheUtils.SetExcelFilters(pdcList, tmpExcelFilterStatus);
          if (pdcList.HasMeasurementParamHandler && pdcList.MeasurementColumn.HasSingleMeasurementTableHandler)
          {
            pdcList.TestDataAdapter.DeleteMeasurementData();
            //bool[] tmpTakeOrLeaveFlags = CreateLeaves(pdcList.DataRange.Row, pdcList.DataRange.Rows.Count, rows);
            bool[] tmpHidden = pdcList.HiddenRows();

            pdcList.TestDataAdapter.GetTestData(true, tmpHidden, true, false, false);
          }
          if (pdcList.HasMeasurementParamHandler && pdcList.MeasurementColumn.HasMultiMeasurementTableHandler)
          {
            pdcList.MeasurementColumn.MultiMeasurementTableHandler.RemoveUnreferencedTables(pdcList, pdcList.GetColumnIndex(PDCExcelConstants.MEASUREMENTS).Value);
          }
          SheetInfo aSheetInfo = pdcList.SheetInfo;
          Dictionary<int, int> rows = ExcelUtils.TheUtils.SelectedRows(pdcList.DataRange);
          List<decimal> tmpExperimentNos = GetExperimentNos(rows);
          if (tmpExperimentNos.Count == 0)
          {
            
            CheckPDCSheet(aSheetInfo);
            ExcelUtils.TheUtils.DeleteShapes(aSheetInfo);
            pdcList.ClearContents();
          }
        }
      }
      catch (Exception e)
      {
        tmpException = e;
      }
      finally
      {
        aWindowOwner.StatusCallback(tmpException);
      }
    }

    /// <summary>
    /// Delete rows in sheet or display error message
    /// </summary>
    /// <param name="aResult"></param>
    private void ActionCompleted(object aResult, ProgressDialog aWindowOwner, bool interactive)
    {
      if (aResult is Exception && interactive)
      {
        ExceptionHandler.TheExceptionHandler.handleException((Exception)aResult, aWindowOwner);
        return;
      }
      if (pdcList != null && pdcList.Container != null)
      {
        try
        {
          ((Excel.Range) pdcList.Container.Cells[pdcList.DataRange.Row, pdcList.DataRange.Column]).Select();
        }
#pragma warning disable 0168
        catch (Exception e2)
        {
        }
#pragma warning restore 0168
      }
      mySelectedRange = null;
      pdcList = null;
      testDefinition = null;

      experimentNos = null;
      Globals.PDCExcelAddIn.EventsEnabled = true;
    }
    /// <summary>
    /// Returns an array which specifies the rows which are not selected.
    /// </summary>
    /// <param name="aStartRow">The start row of the table</param>
    /// <param name="aNrOfRows">The end row of the table</param>
    /// <param name="theSelectedRows">The row numbers of the selected cells</param>
    /// <returns></returns>
    private bool[] CreateLeaves(int aStartRow, int aNrOfRows, Dictionary<int, int> theSelectedRows)
    {
      bool[] tmpFlags = new bool[aNrOfRows + 1];
      for (int i = 0; i < tmpFlags.Length; i++)
      {
        tmpFlags[i] = !theSelectedRows.ContainsKey(aStartRow + i);
      }
      return tmpFlags;
    }
    /// <summary>
    /// Collects all Experimentnos in the 
    /// </summary>
    /// <returns></returns>
    private List<decimal>  GetExperimentNos(  Dictionary<int, int>  rows )
    {
        List<decimal> tmpExperimentNos = new List<decimal>();

        int? tmpColumn = pdcList.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO);
        object[,] tmpExperimentNosValues = pdcList.GetColumnValues(tmpColumn.Value);

        pdcList.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO);
        int tmpLower = tmpExperimentNosValues.GetLowerBound(0);

        int tmpDataRowStart = pdcList.DataRange.Row;

        for (int tmpRow = tmpLower; tmpRow <= tmpExperimentNosValues.GetUpperBound(0); tmpRow++)
        {
          int tmpRealRow = (tmpRow - tmpLower) + tmpDataRowStart;
          if (rows.ContainsKey(tmpRealRow))
          {
            decimal? tmpValue = Lib.PDCConverter.Converter.ToDecimal(tmpExperimentNosValues[tmpRow, tmpExperimentNosValues.GetLowerBound(1)], ExcelUtils.TheUtils.GetExcelNumberSeparators());
            if (tmpValue != null)
            {
              tmpExperimentNos.Add(tmpValue.Value);
            }
          }
        }
        return tmpExperimentNos;
    }
  }
}
