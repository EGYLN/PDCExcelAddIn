using System;
using System.Collections.Generic;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Performs the upload of changed rows to the server. The changed rows must
    /// be selected by the user.
    /// </summary>
    class UpdateAction : PDCAction
    {
        public const string ACTION_TAG = "PDC_UpdateAction";
        private PDCListObject myPdcListObject;
        private Lib.Testdata myTestdata;
        private Dictionary<int, int> mySelectedRows;
        private bool myInteractive;
        private HashSet<decimal> myExperimentNosToDelete = new HashSet<decimal>();

        #region constructor
        public UpdateAction(bool beginGroup)
            : base(Properties.Resources.Action_Update_Caption, ACTION_TAG, beginGroup)
        {
        }
        #endregion

        #region methods

        #region ActionCompleted
        private void ActionCompleted(object aResult, ProgressDialog aWindowOwner, bool interactive)
        {
            if (!(aResult is List<Lib.PDCMessage>))
            {
                return;
            }
            List<Lib.PDCMessage> tmpMessages = (List<Lib.PDCMessage>)aResult;
            if (tmpMessages.Count == 0)
            {
                if (interactive)
                {
                    MessageBox.Show(Properties.Resources.MSG_UPDATE_OK_TEXT, Properties.Resources.MSG_UPDATE_OK_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                myPdcListObject.ValidationHandler.DisplayValidationMessages(tmpMessages, myTestdata, interactive, true);
            }

        }
        #endregion

        #region CanPerformAction
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

        #region CreateLeaves
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
        #endregion

        #region PerformAction
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            CheckPDCSheet(sheetInfo);
            myPdcListObject = sheetInfo.MainTable;
            Globals.PDCExcelAddIn.Application.Cursor = XlMousePointer.xlWait;
            try
            {
                //Check Selection 
                //todo similar code in CompoundInfoAction->Refactor to ExcelUtils
                object tmpSelection = Globals.PDCExcelAddIn.Application.Selection;
                if (!(tmpSelection is Excel.Range))
                {
                    if (interactive)
                    {
                        MessageBox.Show(
                            Properties.Resources.MSG_NO_CELLS_SELECTED,
                            Properties.Resources.MSG_ERROR_TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return new ActionStatus(new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_FATAL));
                }

                Range tmpSelectedRange = (Range)tmpSelection;
                if (tmpSelectedRange.Count == 0)
                {
                    if (interactive)
                    {
                        MessageBox.Show(Properties.Resources.MSG_NO_CELLS_SELECTED, Properties.Resources.MSG_ERROR_TITLE,
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return new ActionStatus(new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_FATAL));
                }
                // Get Selected Rows
                mySelectedRows = ExcelUtils.TheUtils.SelectedRows(tmpSelectedRange);
                // If there are hidden rows selected, it's not possible to delete.
                foreach (Range row in tmpSelectedRange.Rows)
                {

                    if ((bool)row.Hidden)
                    {
                        // if hidden rows obviously caused through 
                        bool filtered = ExcelUtils.TheUtils.IsAnyFilterOn(myPdcListObject.Container);
                        if (interactive)
                        {
                            if (filtered)
                            {
                                MessageBox.Show(Properties.Resources.MSG_AUTO_FILTER_SET_UPDATE, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK);
                            }
                            else
                            {
                                MessageBox.Show(Properties.Resources.MSG_HIDDEN_ROWS_SELECTED, Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK);
                            }
                        }
                        if (filtered)
                        {
                            return new ActionStatus(new[] {
              new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_AUTO_FILTER_SET_UPDATE, Lib.PDCMessage.TYPE_ERROR)});
                        }
                        return new ActionStatus(new Lib.PDCMessage[] {
              new Lib.PDCMessage(Properties.Resources.MSG_ERROR_TITLE, Lib.PDCMessage.TYPE_ERROR),
              new Lib.PDCMessage(Properties.Resources.MSG_HIDDEN_ROWS_SELECTED, Lib.PDCMessage.TYPE_ERROR)});
                    }
                }

                if (interactive)
                {
                    DialogResult tmpConfirm = MessageBox.Show(string.Format(Properties.Resources.MSG_CONFIRM_UPDATE_DATA, mySelectedRows.Count), Properties.Resources.MSG_CONFIRM_TITLE, MessageBoxButtons.YesNo);
                    if (DialogResult.No == tmpConfirm)
                    {
                        return new ActionStatus();
                    }
                }
                // This is a ugly workaround to have this variable "interactive" also in Performupdate
                myInteractive = interactive;
                object tmpResult = ProgressDialog.Show(ActionCompleted, PerformUpdate, Properties.Resources.LABEL_UPDATE, false, myInteractive);

                return ResultToActionStatus(tmpResult, Properties.Resources.MSG_UPDATE_TITLE, interactive);
            }
            finally
            {
                Globals.PDCExcelAddIn.Application.Cursor = XlMousePointer.xlDefault;
            }
        }
        #endregion

        #region PerformUpdate
        private void PerformUpdate(ProgressDialog aWindowOwner)
        {
            object tmpResult = null;
            try
            {
                myPdcListObject.ValidationHandler.ClearAllValidationMessages();

                // this next two line are for checking the unique keys between singlemeasurement and main sheet
                if (myPdcListObject.HasMeasurementParamHandler && myPdcListObject.MeasurementColumn.HasSingleMeasurementTableHandler)
                {
                    bool[] tmpHidden = myPdcListObject.HiddenRows();
                    myTestdata = myPdcListObject.TestDataAdapter.GetTestData(true, tmpHidden, false, true, false);
                }
                myTestdata = null;
                bool[] tmpTakeOrLeaveFlags = CreateLeaves(myPdcListObject.DataRange.Row, myPdcListObject.DataRange.Rows.Count, mySelectedRows);
                myTestdata = myPdcListObject.TestDataAdapter.GetTestData(true, tmpTakeOrLeaveFlags, false, false, true);



                if (myTestdata.IsEmpty())
                {
                    tmpResult = Properties.Resources.MSG_NO_DATA_TO_UPLOAD_TEXT;
                    return;
                }

                if (!myPdcListObject.ValidationHandler.Validate((ExperimentAndMeasurementValues)myTestdata.Tag, tmpTakeOrLeaveFlags))
                {
                    tmpResult = Properties.Resources.MSG_VALIDATION_FAILED;
                    return;
                }
                if (!CheckForMeasurementsNotBeenSaved())
                {
                    tmpResult = Properties.Resources.MSG_USER_ABORTED_ACTION;
                    return;
                }
                if (!Globals.PDCExcelAddIn.PdcService.CheckExperimentNos(myTestdata))
                {
                    tmpResult = Properties.Resources.MSG_UPDATE_RECORDS_DELETED;
                    return;
                }
                if (!Autoupdate(myPdcListObject, tmpTakeOrLeaveFlags))
                {
                    return;
                }
                tmpResult = Globals.PDCExcelAddIn.PdcService.UploadChanges(myTestdata, myExperimentNosToDelete);

                ExcelFilterStatus tmpExcelFilterStatus = ExcelUtils.TheUtils.CollectExcelFilters(myPdcListObject);

                myPdcListObject.Container.AutoFilterMode = false;
                myPdcListObject.TestDataAdapter.UpdateFromUpload(myTestdata);

                ExcelUtils.TheUtils.SetExcelFilters(myPdcListObject, tmpExcelFilterStatus);

            }
            catch (Exception e)
            {
                tmpResult = e;
            }
            finally
            {
                aWindowOwner.StatusCallback(tmpResult);
            }
        }

        private bool Autoupdate(PDCListObject tmpPDCList, bool[] tmpHidden)
        {
            if (tmpPDCList.Testdefinition.HasExperimentLevelVariables)
            {
                // before upload in case of SMT with exp level data, check for existing data in DB
                // create a copy of the test data to fill experimentnos in
                // TODO check if SMT with exp level parameters
                foreach (ExperimentData experiment in myTestdata.Experiments)
                {
                    if (!(experiment is PlaceHolderExperiment) && experiment.ExperimentNo != null)
                    {
                        myExperimentNosToDelete.Add(experiment.ExperimentNo.Value);
                    }
                }
                Lib.Testdata newTestData = tmpPDCList.TestDataAdapter.GetTestData(true, tmpHidden, false, true, true);
                List<Lib.PDCMessage> tmpMessages = Globals.PDCExcelAddIn.PdcService.CheckDuplicateExperiments(myTestdata);
                foreach (ExperimentData experiment in myTestdata.Experiments)
                {
                    if (!(experiment is PlaceHolderExperiment) && experiment.ExperimentNo != null)
                    {
                        myExperimentNosToDelete.Remove(experiment.ExperimentNo.Value);
                    }
                }

                // check for message and display decision dialog: Overwrite or upload, when overwrite use newTestData, when upload normal upload

                if (tmpMessages != null && tmpMessages.Count > 0)
                {
                    DialogResult result = DialogResult.OK;
                    if (myInteractive)
                    {
                        result =
                            MessageBox.Show(string.Format(Properties.Resources.MSG_AUTOUPDATE_VALUES, (object)tmpMessages[0].Message),
                                Properties.Resources.MSG_AUTOUPDATE_TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    }
                    switch (result)
                    {
                        case DialogResult.No:
                            return false;
                        case DialogResult.Yes:
                            return true;
                    }
                }
            }
            return true;
        }
        /// <summary>
        /// This method let the useer check if measurements are not been saved.
        /// </summary>
        /// <returns>false, if user don't want to save data without meassurements </returns>
        private bool CheckForMeasurementsNotBeenSaved()
        {
            if (myInteractive)
            {
                bool blnAsk = true;
                for (int i = 0; i < myTestdata.Experiments.Count; i++)
                {
                    if (!(myTestdata[i] is Lib.PlaceHolderExperiment))
                    {
                        if (!myTestdata.Experiments[i].MeasurementsLoaded && myTestdata.Experiments[i].MaxNumberOfMeasurementValues > 0)
                        {
                            if (blnAsk)
                            {
                                if (MessageBox.Show(Properties.Resources.MSG_CONFIRM_UPDATE_WITHOUT_MEASUREMENTS,
                                                    Properties.Resources.MSG_CONFIRM_TITLE,
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Question) == DialogResult.No)
                                {
                                    myTestdata = null;
                                    return false;
                                }
                            }
                            blnAsk = false;
                        }
                    }
                }
            }
            return true;
        }
        #endregion

        #endregion
    }
}
