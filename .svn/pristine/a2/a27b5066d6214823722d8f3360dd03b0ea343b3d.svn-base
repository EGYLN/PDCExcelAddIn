using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.BHC.BSP.PDC.Lib;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Implements the search for test data 
    /// </summary>
    [ComVisible(false)]
    public class SearchTestdataAction : PDCAction
    {
        public const string ACTION_TAG = "PDC_SearchTestdataAction";

        private delegate void Callback();
        private delegate SearchTestdataResultDialog.ResultStatus DisplayResultStatus(Lib.Testdata testData, ProgressDialog windowOwner);

        Lib.Testdefinition myCurrentTD;
        private bool myLoadMeasurements;
        private bool myOmitMeasurements;
        List<Lib.TestdataSearchCriteria> mySearchCriteria;
        private bool myInteractiveFlag; //Temporarily used for worker thread
        #region constructor
        public SearchTestdataAction(bool beginGroup)
            : base(Properties.Resources.Action_SearchTestData_Caption, ACTION_TAG, beginGroup)
        {
            myCommandBarText = Properties.Resources.Action_SearchTestdata_BarTitle;
        }
        #endregion

        #region methods

        #region BuildSearchCriteria
        private Lib.TestdataSearchCriteria BuildSearchCriteria(object[,] dataRow, PDCListObject pdcListObject, Lib.Testdefinition testDefinition)
        {
            bool conditionFound = false;
            Lib.TestdataSearchCriteria testDataSearchCriteria = new Lib.TestdataSearchCriteria();
            testDataSearchCriteria.TestDefinition = testDefinition;
            //CompoundNo, PreparationNo, Assay, Reference + Testparameter
            List<ListColumn> columns = pdcListObject.Columns;
            Dictionary<ListColumn, int> columnMapping = pdcListObject.CurrentListColumnPlacements();
            foreach (ListColumn column in columnMapping.Keys)
            {
                if (IgnoreColumn(column))
                {
                    continue;
                }
                object value = dataRow[dataRow.GetLowerBound(0), dataRow.GetLowerBound(1) + columnMapping[column]];
                if (value == null || (value is string && "".Equals(("" + value).Trim())))
                { //Null -> Ignore column
                    continue;
                }
                conditionFound = true;
                SetValueInCriteria(testDataSearchCriteria, column, value);
            }
            return conditionFound ? testDataSearchCriteria : null;
        }
        #endregion

        #region BuildSearchCriterias
        private List<Lib.TestdataSearchCriteria> BuildSearchCriterias(object[,] values, PDCListObject pdcListObject)
        {
            List<Lib.TestdataSearchCriteria> testDataSearchCriteriaList = new List<BBS.ST.BHC.BSP.PDC.Lib.TestdataSearchCriteria>();

            int offSetY = values.GetLowerBound(0);
            int offSetX = values.GetLowerBound(1);
            Excel.Range dataRange = pdcListObject.DataRange;
            Dictionary<ListColumn, int> columnMapping = pdcListObject.CurrentListColumnPlacements();
            Lib.Testdefinition testDefinition = pdcListObject.Testdefinition;

            for (int y = offSetY; y <= values.GetUpperBound(0); y++)
            {//For all experiments
                Lib.TestdataSearchCriteria testDataSearchCriteria = new BBS.ST.BHC.BSP.PDC.Lib.TestdataSearchCriteria();
                foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
                {
                    if (IgnoreColumn(pair.Key))
                    {
                        continue;
                    }
                    object value = values[y, pair.Value + offSetX];
                    if (value == null || (value is string && "".Equals(("" + value).Trim())))
                    {
                        continue;
                    }
                    SetValueInCriteria(testDataSearchCriteria, pair.Key, value);
                    testDataSearchCriteria.TestDefinition = testDefinition;
                }
                if (testDataSearchCriteria.TestDefinition != null)
                {
                    testDataSearchCriteriaList.Add(testDataSearchCriteria);
                }
            }
            return testDataSearchCriteriaList;
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
            return true;
        }
        # endregion

        #region CriteriaRow
        /// <summary>
        /// Returns the list row which should be used to initialize the search criteria.
        /// Takes the row of the active cell or the first list row if the row does not belong to 
        /// the list.
        /// </summary>
        /// <param name="aPDCList"></param>
        /// <returns></returns>
        private int CriteriaRow(PDCListObject pdcListObject)
        {
            Excel.Range activeCell = Application.ActiveCell;
            int rowNumber = pdcListObject.ToListRow(activeCell.Row);
            return rowNumber == -1 ? 0 : rowNumber;
        }
        #endregion

        #region DisplayStatus
        private SearchTestdataResultDialog.ResultStatus DisplayStatus(Lib.Testdata testdata, ProgressDialog windowOwner)
        {
            if (windowOwner == null || windowOwner.Cancelled || windowOwner.IsDisposed)
            {
                return SearchTestdataResultDialog.ResultStatus.CANCEL;
            }
            if (testdata == null || testdata.Experiments == null || testdata.Experiments.Count == 0)
            {
                MessageBox.Show("No Entries found.", Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return SearchTestdataResultDialog.ResultStatus.CANCEL;
            }
            SearchTestdataResultDialog.ResultStatus status = SearchTestdataResultDialog.Show(windowOwner, testdata.Experiments.Count, myLoadMeasurements, ShouldAskForMeasurments(testdata));
            if (status == SearchTestdataResultDialog.ResultStatus.CANCEL)
            {
                return SearchTestdataResultDialog.ResultStatus.CANCEL; ;
            }
            windowOwner.CanCancel = false;
            windowOwner.Message = Properties.Resources.Action_SearchTestData_Filling_Caption;

            return status;
        }

        private bool ShouldAskForMeasurments(Lib.Testdata testdata)
        {
            if (!myOmitMeasurements)
            {
                return false;
            }
            foreach (ExperimentData experiment in testdata.Experiments)
            {
                if (experiment.MaxNumberOfMeasurementValues != 0)
                {
                    return false;
                }
            }
            return true;
        }
        #endregion

        #region FillTestData
        private void FillTestData(Lib.Testdata testData, bool writeMeasurements)
        {
            if (testData == null)
            {
                return;
            }
            Globals.PDCExcelAddIn.EventsEnabled = false;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            try
            {
                PDCLogger.TheLogger.LogStarttime("FillTestData", "Filling sheet");
                Excel.Worksheet sheet = (Excel.Worksheet)Application.ActiveSheet;

                PDCListObject pdcListObject = (PDCListObject)testData.TestVersion.Tag;
                pdcListObject.TestDataAdapter.SetTestdata(testData, writeMeasurements);
                pdcListObject.SheetInfo.AreMeasurementsLoaded = writeMeasurements;
                PDCLogger.TheLogger.LogStoptime("FillTestData", "Filling sheet");
            }
            finally
            {
                Globals.PDCExcelAddIn.EnableExcel();
            }
        }
        #endregion

        #region GetDataCell
        private Excel.Range GetDataCell(Excel.Worksheet sheet, int rowNumber, string rangeName)
        {
            Excel.Range column = sheet.get_Range(rangeName, missing);
            return (Excel.Range)sheet.Cells[rowNumber, column.Column];
        }
        #endregion

        #region GetRowValues
        private object[,] GetRowValues(Excel.Worksheet sheet, PDCListObject pdcListObject, int listRow)
        {
            return pdcListObject[listRow];
        }
        #endregion

        #region IgnoreColumn
        private bool IgnoreColumn(ListColumn listColumn)
        {
            // todo Implement SingleMeasurementTableHandler
            return listColumn.IsHyperLink ||
              listColumn.ParamHandler is Predefined.MultipleMeasurementTableHandler ||
              listColumn.ParamHandler is Predefined.RowNumberHandler;
        }
        #endregion

        #region PerformAction
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            List<Lib.TestdataSearchCriteria> searchCriterias = new List<BBS.ST.BHC.BSP.PDC.Lib.TestdataSearchCriteria>();
            try
            {
                myLoadMeasurements = sheetInfo.TestDefinition.ShowSingleMeasurement;
                if (!myLoadMeasurements)
                {
                    myOmitMeasurements = sheetInfo.MainTable.HasMeasurementParamHandler;
                }
                else
                {
                    myOmitMeasurements = false;
                }
                CheckPDCSheet(sheetInfo);
                Excel.Worksheet sheet = sheetInfo.ExcelSheet;
                Lib.Testdefinition testDefinition = sheetInfo.TestDefinition;
                //Find appriopriate row
                PDCListObject pdcListObject = sheetInfo.MainTable;
                object selection = Globals.PDCExcelAddIn.Application.Selection;
                object activeCell = Globals.PDCExcelAddIn.Application.ActiveCell;

                if (ExcelUtils.TheUtils.SameCells(selection, activeCell) ||
                    selection == null || !(selection is Excel.Range))
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
                    myCurrentTD = testDefinition;
                    mySearchCriteria = searchCriterias;
                    myInteractiveFlag = interactive;
                    ProgressDialog.Show(SearchCompleted, PerformSearch, Properties.Resources.LABEL_SEARCH_TESTDATA);
                }
                else if (interactive)
                {
                    MessageBox.Show(Properties.Resources.MSG_NO_SEARCH_CONDITIONS_TEXT, Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                    return new ActionStatus(e);
                }
            }
            return new ActionStatus();
        }
        #endregion

        #region PerformSearch
        private void PerformSearch(ProgressDialog windowOwner)
        {
            object result = null;
            Lib.Testdata testData = null;
            try
            {
                Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
                testData = tmpService.FindTestdata(myCurrentTD, mySearchCriteria);
                if (!windowOwner.Cancelled)
                {
                    SearchTestdataResultDialog.ResultStatus status = SearchTestdataResultDialog.ResultStatus.REPLACEWITHOUTMEASUREMENTS;
                    if (myInteractiveFlag)
                    {
                        status = (SearchTestdataResultDialog.ResultStatus)windowOwner.Invoke(new DisplayResultStatus(DisplayStatus),
                            testData, windowOwner);
                    }
                    switch (status)
                    {
                        case SearchTestdataResultDialog.ResultStatus.REPLACE:
                            this.FillTestData(testData, true);
                            break;
                        case SearchTestdataResultDialog.ResultStatus.REPLACEWITHOUTMEASUREMENTS:
                            this.FillTestData(testData, false);
                            break;
                    }
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

        private void SearchCompleted(Lib.Testdata testData, ProgressDialog windowOwner)
        {
            if (windowOwner == null || windowOwner.Cancelled || windowOwner.IsDisposed)
            {
                return;
            }
            if (testData == null || testData.Experiments == null || testData.Experiments.Count == 0)
            {
                MessageBox.Show(Properties.Resources.MSG_NO_ENTRIES_FOUND_TEXT, Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            SearchTestdataResultDialog.ResultStatus status = SearchTestdataResultDialog.Show(windowOwner, testData.Experiments.Count, myLoadMeasurements, ShouldAskForMeasurments(testData));
            if (status == SearchTestdataResultDialog.ResultStatus.CANCEL)
            {
                return;
            }
            windowOwner.CanCancel = false;
            switch (status)
            {
                case SearchTestdataResultDialog.ResultStatus.REPLACE:
                    this.FillTestData(testData, true);
                    break;
                case SearchTestdataResultDialog.ResultStatus.REPLACEWITHOUTMEASUREMENTS:
                    this.FillTestData(testData, false);
                    break;
            }
        }
        #endregion

        #region SetUploadDateCriteria
        private Lib.TestVariableValue SetDateCriteria(Lib.TestdataSearchCriteria testDataSearchCriteria, ListColumn listColumn, object value)
        {
            Lib.TestVariableValue testVariableValue;
            string dateString = null;
            string upperLimit = null;
            string lowerLimit = null;
            DateTime? parsedDate = null;
            if (value is string)
            {
                dateString = ((string)value).Trim();
                if (dateString.StartsWith(">")) // parse lower limit
                {
                    parsedDate = Lib.PDCConverter.Converter.ParseDate(dateString.Substring(1), Lib.PDCConverter.INPUT_DATE_FORMAT);
                    if (parsedDate == null)
                    {
                        throw new Exceptions.InvalidDateException(dateString, listColumn.Label);
                    }
                    lowerLimit = Lib.PDCConverter.Converter.FromDate(parsedDate);
                    dateString = lowerLimit + "-";
                }
                else if (dateString.StartsWith("<")) //parse upper limit
                {
                    parsedDate = Lib.PDCConverter.Converter.ParseDate(dateString.Substring(1), Lib.PDCConverter.INPUT_DATE_FORMAT);
                    if (parsedDate == null)
                    {
                        throw new Exceptions.InvalidDateException(dateString, listColumn.Label);
                    }
                    upperLimit = Lib.PDCConverter.Converter.FromDate(parsedDate);
                    dateString = "-" + upperLimit;
                }
                else if (dateString.Contains("-")) //parse range
                {
                    string[] parts = dateString.Split('-');
                    if (parts.Length == 2)
                    {
                        parsedDate = PDC.Lib.PDCConverter.Converter.ParseDate(parts[0], Lib.PDCConverter.INPUT_DATE_FORMAT);
                        if (parsedDate == null)
                        {
                            throw new Exceptions.InvalidDateException(dateString, listColumn.Label);
                        }
                        lowerLimit = PDC.Lib.PDCConverter.Converter.FromDate(parsedDate);
                        parsedDate = PDC.Lib.PDCConverter.Converter.ParseDate(parts[1], Lib.PDCConverter.INPUT_DATE_FORMAT);
                        if (parsedDate == null)
                        {
                            throw new Exceptions.InvalidDateException(dateString, listColumn.Label);
                        }
                        upperLimit = PDC.Lib.PDCConverter.Converter.FromDate(parsedDate);
                        dateString = lowerLimit + "-" + upperLimit;
                    }
                }
                else //exact date
                {
                    if (!string.Empty.Equals(dateString))
                    {
                        parsedDate = PDC.Lib.PDCConverter.Converter.ParseDate(dateString, PDC.Lib.PDCConverter.INPUT_DATE_FORMAT);
                        if (parsedDate == null)
                        {
                            throw new Exceptions.InvalidDateException(dateString, listColumn.Label);
                        }
                        dateString = PDC.Lib.PDCConverter.Converter.FromDate(parsedDate);
                    }
                    // else do nothing
                }
            }
            else if (value is DateTime)
            { //Not a string 
                dateString = PDC.Lib.PDCConverter.Converter.FromDate(value);
            }
            else
            { //Dont know
                throw new Exceptions.InvalidDateException("" + value, listColumn.Label);
            }
            testVariableValue = new Lib.TestVariableValue(listColumn.Id.Value, dateString);
            testVariableValue.ValueCharUpperLimit = upperLimit;
            testVariableValue.ValueCharLowerLimit = lowerLimit;
            testDataSearchCriteria[listColumn.Id.Value, false] = testVariableValue;
            return testVariableValue;
        }
        #endregion

        #region SetValueInCriteria
        private void SetValueInCriteria(Lib.TestdataSearchCriteria testDataSearchCriteria, ListColumn listColumn, object value)
        {
            Lib.TestVariableValue testVariableValue;
            if (listColumn.TestVariable != null)
            { //
                if (listColumn.TestVariable.IsNumeric())
                {
                    testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
                    string prefix = null;
                    if (value is string)
                    {
                        value = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)value, out prefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                    }
                    if (!Lib.PDCConverter.Converter.DoubleToString(value, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue))
                    {
                        throw new Exceptions.InvalidNumberFormatException(listColumn.TestVariable.VariableName, value, value.GetType().ToString(), ExcelUtils.TheUtils.GetExcelNumberSeparators().NumberDecimalSeparator, ExcelUtils.TheUtils.GetExcelNumberSeparators().NumberGroupSeparator);
                    }
                    PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Testvariable '" + listColumn.TestVariable.VariableName + "' value has been transformed from '" + value + " (" + value.GetType() + ") to '" + testVariableValue.ValueChar + "'");
                    testVariableValue.IsNummeric = true;
                    testVariableValue.Prefix = prefix;

                    testDataSearchCriteria[listColumn.Id.Value, true] = testVariableValue;

                }
                else
                {
                    testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
                    testVariableValue.ValueChar = "" + value;
                    testDataSearchCriteria[listColumn.Id.Value, true] = testVariableValue;
                }
            }
            else if (listColumn.Id == Lib.PDCConstants.C_ID_UPLOADDATE || listColumn.Id == Lib.PDCConstants.C_ID_DATE_RESULT)
            {
                testVariableValue = SetDateCriteria(testDataSearchCriteria, listColumn, value);
            }
            else if (listColumn.Id == Lib.PDCConstants.C_ID_EXPERIMENTNO)
            {
                decimal? decimalValue = Lib.PDCConverter.Converter.ToDecimal(value, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                if (decimalValue.HasValue)
                {
                    testVariableValue = new Lib.TestVariableValue(listColumn.Id.Value, decimalValue.Value);
                    testDataSearchCriteria[listColumn.Id.Value, false] = testVariableValue;
                }
            }
            else if (listColumn.Id == Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER ||
              listColumn.Id == Lib.PDCConstants.C_ID_PREPARATIONNO ||
              listColumn.Id == Lib.PDCConstants.C_ID_ASSAY_REFERENCE ||
              listColumn.Id == Lib.PDCConstants.C_ID_REFERENCE ||
              listColumn.Id == Lib.PDCConstants.C_ID_MCNO)
            {
                testVariableValue = new Lib.TestVariableValue(listColumn.Id.Value, "" + value);
                testDataSearchCriteria[listColumn.Id.Value, false] = testVariableValue;
            }
        }
        #endregion

        #endregion
    }
}
