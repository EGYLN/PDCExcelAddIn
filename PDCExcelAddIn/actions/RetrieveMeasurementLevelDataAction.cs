using System;
using System.Collections.Generic;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class RetrieveMeasurementLevelDataAction : PDCAction
    {
        public const string ACTION_TAG = "PDC_RetrieveMeasurementLevelDataAction";
        private PDCListObject myPdcListObject;
        private Lib.Testdata myTestdata;
        private Lib.Testdefinition myTestDefinition;
        List<Lib.TestdataSearchCriteria> mySearchCriteria;
        private Dictionary<int, int> mySelectedRows;

        #region constructor
        public RetrieveMeasurementLevelDataAction(bool beginGroup)
            : base(Properties.Resources.Action_RetrieveMeasurementLevelData_Caption, ACTION_TAG, beginGroup)
        {
        }
        #endregion

        #region methods

        #region CriteriaRow
        /// <summary>
        /// Returns the list row which should be used to initialize the search criteria.
        /// Takes the row of the active cell or the first list row if the row does not belong to 
        /// the list.
        /// </summary>
        private int CriteriaRow(PDCListObject pdcListObject)
        {
            Excel.Range activeCell = Application.ActiveCell;
            int rowNumber = pdcListObject.ToListRow(activeCell.Row);
            return rowNumber == -1 ? 0 : rowNumber;
        }
        #endregion

        #region GetRowValues
        private object[,] GetRowValues(Excel.Worksheet sheet, PDCListObject pdcListObject, int listRow)
        {
            return pdcListObject[listRow];
        }
        #endregion

        #region CreateLeaves
        /// <summary>
        /// Returns an array which specifies the rows which are not selected.
        /// </summary>
        /// <param name="startRow">The start row of the table</param>
        /// <param name="nrOfRows">The end row of the table</param>
        /// <param name="selectedRows">The row numbers of the selected cells</param>
        /// <returns></returns>
        private bool[] CreateLeaves(int startRow, int nrOfRows, Dictionary<int, int> selectedRows)
        {
            bool[] flags = new bool[nrOfRows + 1];
            for (int i = 0; i < flags.Length; i++)
            {
                flags[i] = !selectedRows.ContainsKey(startRow + i);
            }
            return flags;
        }
        #endregion

        #region PerformAction
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            List<Lib.TestdataSearchCriteria> searchCriterias = new List<Lib.TestdataSearchCriteria>();
            try
            {
                CheckPDCSheet(sheetInfo);
                myPdcListObject = sheetInfo.MainTable;
                Excel.Worksheet sheet = sheetInfo.ExcelSheet;
                Lib.Testdefinition testDefinition = sheetInfo.TestDefinition;
                //Find appriopriate row
                PDCListObject pdcListObject = sheetInfo.MainTable;
                object selection = Globals.PDCExcelAddIn.Application.Selection;
                object activeCell = Globals.PDCExcelAddIn.Application.ActiveCell;

                // Get Selected Rows


                if (!(selection is Excel.Range))
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

                Excel.Range tmpSelectedRange = (Excel.Range)selection;
                if (tmpSelectedRange.Count == 0)
                {
                    if (interactive)
                    {
                        MessageBox.Show(Properties.Resources.MSG_NO_CELLS_SELECTED, Properties.Resources.MSG_ERROR_TITLE,
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return new ActionStatus(new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_FATAL));
                }

                if (ExcelUtils.TheUtils.SameCells(selection, activeCell))
                { //take first table row as criteria
                    int selectedRowNo = CriteriaRow(pdcListObject);
                    //Create Search Criteria
                    //Execute search
                    object[,] values = GetRowValues(sheet, pdcListObject, selectedRowNo);

                    Lib.TestdataSearchCriteria searchCriteria = BuildSearchCriteria(values, pdcListObject, testDefinition);
                    if (searchCriteria != null)
                    {
                        searchCriterias.Add(searchCriteria);
                    }
                }
                else
                {
                    object[,] values = pdcListObject.GetSelectedValues();
                    searchCriterias = BuildSearchCriterias(values, pdcListObject);
                }
                if (searchCriterias.Count > 0)
                {
                    mySelectedRows = ExcelUtils.TheUtils.SelectedRows(tmpSelectedRange);
                    myTestDefinition = testDefinition;
                    mySearchCriteria = searchCriterias;
                    myPdcListObject = pdcListObject;
                    object result = ProgressDialog.Show(SearchCompleted, PerformSearch, Properties.Resources.LABEL_SEARCH_TESTDATA, false, interactive);
                    return ResultToActionStatus(result, Properties.Resources.MSG_LOAD_MEASUREMENTS_TITLE, interactive);
                }
                else
                {
                    MessageBox.Show(Properties.Resources.MSG_NO_SEARCH_CONDITIONS_TEXT, Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return new ActionStatus(new Lib.PDCMessage(Properties.Resources.MSG_NO_CELLS_SELECTED, Lib.PDCMessage.TYPE_FATAL));
                }
            }
            catch (Exception e)
            {
                if (interactive)
                {
                    ExceptionHandler.TheExceptionHandler.handleException(e, null);
                }
                else
                {
                    ActionStatus actioStatus = new ActionStatus(e);
                    return actioStatus;
                }

            }
            finally
            {
                Globals.PDCExcelAddIn.Application.Cursor = Excel.XlMousePointer.xlDefault;
            }
            return new ActionStatus();

        }
        #endregion

        #region PerformSearch
        private void PerformSearch(ProgressDialog windowOwner)
        {
            object result = null;
            try
            {
                // Get Selected Rows
                bool[] tmpHidden = myPdcListObject.HiddenRows();
                // this next two line are for checking the unique keys between singlemeasurement and main sheet
                //myTestdata = myPdcListObject.TestDataAdapter.GetTestData(false, tmpHidden, false, true);
                bool[] tmpTakeOrLeaveFlags = CreateLeaves(myPdcListObject.DataRange.Row, myPdcListObject.DataRange.Rows.Count, mySelectedRows);
                myTestdata = myPdcListObject.TestDataAdapter.GetTestData(false, tmpTakeOrLeaveFlags, false, false, true);


                Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
                Lib.Testdata testData = tmpService.FindTestdata(myTestDefinition, mySearchCriteria);
                if (!windowOwner.Cancelled)
                {
                    FillMeasurementTestData(testData);
                }
            }
            catch (Exception e)
            {
                result = e;
            }
            finally
            {
                windowOwner.StatusCallback(result);
            }
        }
        #endregion

        #region SearchCompleted
        private void SearchCompleted(object result, ProgressDialog windowOwner, bool interactive)
        {
            if (result is string && ((string)result) == ProgressDialog.CANCELLED)
            {
                return;
            }
            if (result is Exception)
            {
                throw (Exception)result;
            }
        }

        #endregion
        #region FillTestData
        private void FillMeasurementTestData(Lib.Testdata testData)
        {
            if (testData == null)
            {
                return;
            }
            Globals.PDCExcelAddIn.EventsEnabled = false;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            try
            {
                PDCLogger.TheLogger.LogStarttime("FillMeasurementTestData", "Filling sheet");

                foreach (Lib.ExperimentData oneOfRetrievedExperiments in testData.Experiments)
                {
                    int i = 0;
                    foreach (Lib.ExperimentData oneOfAllExperiments in myTestdata.Experiments)
                    {

                        if (oneOfAllExperiments.ExperimentNo == oneOfRetrievedExperiments.ExperimentNo)
                        {
                            myTestdata.Experiments[i] = oneOfRetrievedExperiments;
                            break;
                        }
                        i++;
                    }
                }

                PDCListObject pdcListObject = (PDCListObject)testData.TestVersion.Tag;
                pdcListObject.TestDataAdapter.SetMeasurementTestData(myTestdata);
                PDCLogger.TheLogger.LogStoptime("FillMeasurementTestData", "Filling sheet");
            }
            finally
            {
                Globals.PDCExcelAddIn.EnableExcel();
            }
        }
        #endregion

        #region BuildSearchCriteria
        private Lib.TestdataSearchCriteria BuildSearchCriteria(object[,] dataRow, PDCListObject pdcListObject, Lib.Testdefinition testDefinition)
        {
            Lib.TestdataSearchCriteria testDataSearchCriteria = new Lib.TestdataSearchCriteria();
            testDataSearchCriteria.TestDefinition = testDefinition;
            //CompoundNo, PreparationNo, Assay, Reference + Testparameter
            ListColumn experimentNoListColum = pdcListObject.ListColumnByColumnId(Lib.PDCConstants.C_ID_EXPERIMENTNO);

            int experimentNoPos = pdcListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value;

            object value = dataRow[dataRow.GetLowerBound(0), dataRow.GetLowerBound(1) + experimentNoPos];
            if (value == null || (value is string && "".Equals(("" + value).Trim())))
            { //Null -> Ignore column
                return null;
            }
            decimal? decimalValue = Lib.PDCConverter.Converter.ToDecimal(value, ExcelUtils.TheUtils.GetExcelNumberSeparators());
            if (decimalValue.HasValue)
            {
                Lib.TestVariableValue testVariableValue = new Lib.TestVariableValue(experimentNoListColum.Id.Value, decimalValue.Value);
                testDataSearchCriteria[experimentNoListColum.Id.Value, false] = testVariableValue;
            }

            return testDataSearchCriteria;
        }
        #endregion

        #region BuildSearchCriterias
        private List<Lib.TestdataSearchCriteria> BuildSearchCriterias(object[,] values, PDCListObject pdcListObject)
        {
            List<Lib.TestdataSearchCriteria> testDataSearchCriteriaList = new List<Lib.TestdataSearchCriteria>();

            int offSetY = values.GetLowerBound(0);
            int offSetX = values.GetLowerBound(1);
            pdcListObject.CurrentListColumnPlacements();
            Lib.Testdefinition testDefinition = pdcListObject.Testdefinition;


            ListColumn experimentNoListColum = pdcListObject.ListColumnByColumnId(Lib.PDCConstants.C_ID_EXPERIMENTNO);

            int experimentNoPos = pdcListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value;


            for (int y = offSetY; y <= values.GetUpperBound(0); y++)
            {//For all experiments

                Lib.TestdataSearchCriteria testDataSearchCriteria = new Lib.TestdataSearchCriteria();

                object value = values[y, experimentNoPos + offSetX];
                if (value == null || (value is string && "".Equals(("" + value).Trim())))
                {
                    continue;
                }

                decimal? decimalValue = Lib.PDCConverter.Converter.ToDecimal(value, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                if (decimalValue.HasValue)
                {
                    Lib.TestVariableValue testVariableValue = new Lib.TestVariableValue(experimentNoListColum.Id.Value, decimalValue.Value);
                    testDataSearchCriteria[experimentNoListColum.Id.Value, false] = testVariableValue;
                }
                testDataSearchCriteria.TestDefinition = testDefinition;
                testDataSearchCriteriaList.Add(testDataSearchCriteria);
            }
            return testDataSearchCriteriaList;
        }
        #endregion

        #endregion
    }
}
