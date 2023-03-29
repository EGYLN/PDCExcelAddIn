using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined;
//using BBS.ST.BHC.BSP.PDC.ExcelClient.actions;
using System.Diagnostics;
using System.Windows.Forms;
using System.Linq;
using BBS.ST.BHC.BSP.PDC.ExcelClient.actions;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Adapter between test data and a PDCListObject holding test data.
    /// </summary>
    [Serializable]
    [ComVisible(false)]
    public class TestDataTableAdapter
    {
        private PDCListObject myMainPDCListObject;
        UniqueExperimentKeyHandler myUniqueExperimentKeyHandler;

        #region constructor

        /// <summary>
        ///   The test data table adapter constructor.
        /// </summary>
        /// <param name="pdcListObject">
        ///   The main PDC list object.
        /// </param>
        public TestDataTableAdapter(PDCListObject pdcListObject)
        {
            myMainPDCListObject = pdcListObject;
        }

        #endregion

        #region methods

        #region AddMeasurementsFromMMT
        /// <summary>
        /// Adds the values of the measurement table to the experiment
        /// </summary>
        /// <param name="experiment"></param>
        /// <param name="column"></param>
        /// <param name="value"></param>
        /// <param name="setDefaults"></param>
        /// <param name="matrixValues"></param>
        /// <returns></returns>
        private bool AddMeasurementsFromMMT(Lib.ExperimentData experiment, ListColumn column, object value, bool setDefaults,
          ExperimentAndMeasurementValues matrixValues)
        {
            bool found = false;

            List<Lib.TestVariableValue> measurements = new List<BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();

            string tableName = "";
            if (value is string[])
            {
                tableName = ((string[])value)[0];
            }
            else
            {
                tableName = "" + value;
            }
            SheetInfo sheetInfo = myMainPDCListObject.SheetInfo;
            PDCListObject measurementTable = sheetInfo.FindMeasurementTable(tableName);
            if (measurementTable == null)
            {
                return found;
            }
            object[,] values = null;
            if (measurementTable is MeasurementPDCListObject && matrixValues.measurementValues != null)
            {
                values = ((MeasurementPDCListObject)measurementTable).GetValues(matrixValues.measurementValues);
            }
            else
            {
                values = measurementTable.Values;
            }
            values = ValidationHandler.Transpone(values);

            Dictionary<ListColumn, int> columnMapping = measurementTable.CurrentListColumnPlacements();
            int position = 0;
            //Over all Measurements
            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
            {
                position++;
                foreach (ListColumn listColumn in columnMapping.Keys)
                {
                    if (listColumn.TestVariable == null)
                    {
                        continue;
                    }
                    int columnRow = columnMapping[listColumn] + values.GetLowerBound(1);
                    if (columnRow < values.GetLowerBound(1) || columnRow > values.GetUpperBound(1))
                    {
                        continue;
                    }
                    object variableValue = values[i, columnRow];
                    if (variableValue == null)
                    {
                        continue;
                    }
                    found = true;
                    Lib.TestVariableValue testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
                    testVariableValue.Position = position;
                    if (listColumn.TestVariable.IsNumeric())
                    {
                        string prefix = null;
                        object valueObject = variableValue;
                        if (variableValue is string)
                        {
                            valueObject = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)variableValue, out prefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                        }

                        Lib.PDCConverter.Converter.DoubleToString(valueObject, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue);
                        testVariableValue.Prefix = prefix;
                        testVariableValue.IsNummeric = true;

                        //testVariableValue.ValueNum = Lib.PDCConverter.Converter.ToDecimal(variableValue, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                        //if (testVariableValue.ValueNum == null)
                        //{
                        //  testVariableValue.ValueChar = "" + variableValue;
                        //}
                    }
                    else
                    {
                        testVariableValue.ValueChar = "" + variableValue;
                    }
                    measurements.Add(testVariableValue);
                }
            }
            experiment.SetMeasurementValues(measurements);
            return found;
        }
        #endregion

        #region AddTestVariableValue
        /// <summary>
        /// Adds a test variable value to the experiment
        /// </summary>
        /// <param name="experimentData"></param>
        /// <param name="listColumn"></param>
        /// <param name="value"></param>
        private void AddTestVariableValue(Lib.ExperimentData experimentData, ListColumn listColumn, object value, Dictionary<int, Lib.TestVariableValue> valueMap)
        {
            Lib.TestVariableValue testVariableValue =
              valueMap.ContainsKey(listColumn.TestVariable.VariableId) ?
              valueMap[listColumn.TestVariable.VariableId] : null;
            if (testVariableValue == null)
            {
                testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
            }
            //Binary parameter
            if (listColumn.TestVariable.IsBinaryParameter())
            {
                if (value is string[])
                {
                    testVariableValue.Filename = ((string[])value)[1];
                    testVariableValue.Url = ((string[])value)[0];
                }
                else
                {
                    testVariableValue.Filename = "" + value;
                    testVariableValue.Url = "" + value;
                }
            }
            else
            {
                if (listColumn.TestVariable.IsNumeric())
                {
                    object valueObject = value;
                    string prefix = null;
                    if (value is string)
                    {
                        valueObject = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)value, out prefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                    }

                    Lib.PDCConverter.Converter.DoubleToString(valueObject, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue);
                    testVariableValue.Prefix = prefix;
                    testVariableValue.IsNummeric = true;
                }
                else if (value != null)
                {
                    testVariableValue.ValueChar = "" + value;
                }
            }
            valueMap[listColumn.TestVariable.VariableId] = testVariableValue;
        }
        #endregion

        #region AddUploadParameter
        /// <summary>
        /// Adds a general Upload parameter value to the experiment
        /// </summary>
        /// <param name="experimentData"></param>
        /// <param name="listColumn"></param>
        /// <param name="value"></param>
        private bool AddUploadParameter(Lib.ExperimentData experimentData, ListColumn listColumn, object value)
        {
            switch (listColumn.Name)
            {
                case PDCExcelConstants.COMPOUNDNO:
                    experimentData.CompoundNo = "" + value;
                    return true;
                case PDCExcelConstants.PREPARATIONNO:
                    experimentData.PreparationNo = "" + value;
                    return true;
                case PDCExcelConstants.UPLOAD_ID:
                    experimentData.UploadId = Lib.PDCConverter.Converter.ToLong(value);
                    break;
                case PDCExcelConstants.MCNO:
                    experimentData.MCNo = "" + value;
                    return true;
                case PDCExcelConstants.EXPERIMENT_NO:
                    experimentData.ExperimentNo = Lib.PDCConverter.Converter.ToLong(value);
                    break;
                case PDCExcelConstants.DATERESULT:
                    experimentData.DateResult = Lib.PDCConverter.Converter.ToDate(value);
                    break;
                case PDCExcelConstants.REPORT_TO_PIX:
                    experimentData.ReportToPix = "" + value;
                    break;

                case PDCExcelConstants.UPLOADDATE:
                    break;
                case PDCExcelConstants.RUNNO:
                    break;
                case PDCExcelConstants.RESULTSTATUS:
                    break;
                default: //Ignore userdefined parameter
                    break;
            }
            return false;
        }
        #endregion


        #region CreateExperiment
        /// <summary>
        /// Creates an experiment from the specified experiment table row 
        /// and an optionally associated measurement table.
        /// </summary>
        /// <param name="testDefinition"></param>
        /// <param name="values">The values from the sheet.</param>
        /// <param name="row"></param>
        /// <param name="columnMapping"></param>
        /// <param name="matrixValues"></param>
        /// <param name="setDefaults"></param>
        /// <returns></returns>
        private Lib.ExperimentData CreateExperiment(Lib.Testdefinition testDefinition, object[,] values, int row, Dictionary<ListColumn, int> columnMapping,
          ExperimentAndMeasurementValues matrixValues, bool setDefaults)
        {
            Lib.ExperimentData experimentData = new BBS.ST.BHC.BSP.PDC.Lib.ExperimentData(testDefinition);
            experimentData.PersonId = Globals.PDCExcelAddIn.PdcService.UserInfo.Cwid;

            string ReportToPixValue = string.Empty;
            experimentData.PersonIdType = 1;
            bool inputFound = false;

            //Examine each column with a mapping
            ListColumn measurementColumn = null;
            object measurementColumnValue = null;

            Dictionary<int, Lib.TestVariableValue> experimentValues = new Dictionary<int, BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();
            Dictionary<int, Lib.TestVariableValue> experimentLevelValues = new Dictionary<int, BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();
            if (testDefinition.HasMeasurementVariables)
            {
                experimentData.MeasurementsLoaded = false;
            }
            foreach (ListColumn column in columnMapping.Keys)
            {

                int index = columnMapping[column];
                object value = values[row, values.GetLowerBound(1) + index];

                // added this code for pdc dummy prep changes PDC-926
                if (values[row, 2] != null)
                {
                    string compoundNumber = values[row, 2].ToString();

                    if ((column.Name == "PreparationNo" && value == null && !string.IsNullOrEmpty(compoundNumber))) //PDC-926
                    {

                        string[] PreparationNumbers = PremarationNumber.PrepartionNo[row - 1].Split(';').Where(val => val != "UNKNOWN").ToArray() ;
                        PreparationNumbers = PreparationNumbers.Where(val => val != "NOT SPECIFIED").ToArray() ;
                        int prCount = PreparationNumbers.Length;

                        if (prCount == 1)
                        {
                            value = PreparationNumbers[0];
                        }
                        if (prCount == 2 && (PremarationNumber.PrepartionNo[row - 1].Contains("DUMY")))
                        {
                            foreach(string preno in PreparationNumbers)
                            if(!preno.Contains("DUMY"))
                               value = preno;
                        }
                        if (prCount > 2)
                        {
                            value = "NOT SPECIFIED";
                            ReportToPixValue = "No";
                        }
                    }
                    if (column.Name == "Var_1324")
                    {
                        ReportToPixValue = null;
                    }
                   
                    if (value != null && column.Name == "PreparationNo" && !value.Equals("NOT SPECIFIED"))
                    {
                        if (!PremarationNumber.PrepartionNo[row - 1].Contains(value.ToString()))
                        {
                            ReportToPixValue = "No";
                        }
                        if (PremarationNumber.PrepartionNo[row - 1].Contains(value.ToString()))
                        {
                            ReportToPixValue = "Yes";
                        }
                    }
                    if (column.Name == "Report_To_PIx" && !string.IsNullOrEmpty(ReportToPixValue))
                    {
                        value = ReportToPixValue;

                    }

                }

                // added this code for pdc dummy prep changes pdc-926

                if (column.HasMultiMeasurementTableHandler || column.HasSingleMeasurementTableHandler)
                {
                    measurementColumn = column;
                    measurementColumnValue = value;
                    experimentData.MeasurementsLoaded = measurementColumnValue == null || !UniqueExperimentKeyHandler.NOT_LOADED.Equals(measurementColumnValue);
                    continue;
                }
                else if (value == null)
                {
                    continue;
                }
                if (column.TestVariable != null)
                {
                    if (testDefinition.ExperimentLevelVariables.ContainsKey(column.TestVariable.VariableId))
                    {
                        AddTestVariableValue(experimentData, column, value, experimentLevelValues);
                    }
                    AddTestVariableValue(experimentData, column, value, experimentValues);
                    inputFound = true;
                }
                else
                {
                    inputFound = AddUploadParameter(experimentData, column, value) || inputFound;
                }
            }

            foreach (Lib.TestVariableValue tmpValue in experimentValues.Values)
            {
                experimentData.GetExperimentValues().Add(tmpValue);
            }
            foreach (Lib.TestVariableValue tmpValue in experimentLevelValues.Values)
            {
                experimentData.GetExperimentLevelVariableValues().Add(tmpValue);
            }

            // Look for measurement data only if experiment values are present
            if (inputFound && measurementColumn != null)
            {
                // in case, the debug-switch is on, a second handler for SinglemeasurementTableHandler has been invoked
                if (measurementColumn.HasSingleMeasurementTableHandler)
                {
                    myUniqueExperimentKeyHandler.SetExperiment(experimentData, row);
                }
                else
                {
                    AddMeasurementsFromMMT(experimentData, measurementColumn, measurementColumnValue, setDefaults, matrixValues);
                }
            }


            return inputFound ? experimentData : null;
        }
        #endregion

        #region CreateImageServletLink
        /// <summary>
        /// Creates url to the pdc image servlet for the specified experiment and test variable
        /// </summary>
        /// <param name="tmpExperiment">Specifies the experimentno used to find the binary</param>
        /// <param name="testVariable">Specifies the test variable used to find the binary</param>
        /// <returns>URL to the PDC image servlet for the given experiment and variable</returns>
        private string CreateImageServletLink(Lib.ExperimentData tmpExperiment, Lib.TestVariable testVariable)
        {
            string tmpLink = Lib.PDCService.ThePDCService.ServerURL + Properties.Settings.Default.ImageServletPath + "?";
            if (tmpExperiment.ExperimentNo == null)
            {
                return null;
            }
            string tmpSourceSystem = tmpExperiment.TestVersion.Sourcesystem;
            if (tmpSourceSystem == null || tmpSourceSystem == "")
            {
                tmpSourceSystem = "PDC";
            }
            tmpLink += "sourceSystem=" + tmpSourceSystem;
            tmpLink += "&experimentno=" + tmpExperiment.ExperimentNo;
            tmpLink += "&variableid=" + testVariable.VariableId;
            return tmpLink;
        }
        #endregion

        #region CreateSingleColumnValue
        /// <summary>
        /// creates a twodimensinal array with one column and fills it with the value
        /// </summary>
        /// <returns></returns>
        private object[,] CreateSingleColumnValue(List<Lib.ExperimentData> experiments, string value)
        {
            object[,] values = new object[experiments.Count, 1];
            for (int i = 0; i < experiments.Count; i++)
            {
                if (experiments[i] is Lib.PlaceHolderExperiment && experiments[i].ExperimentNo == -666)
                {
                    continue;
                }
                values[i, 0] = value;
            }
            return values;
        }
        #endregion



        #region DeleteMeasurementData
        public void DeleteMeasurementData()
        {
            if (myUniqueExperimentKeyHandler != null)
            {
                myUniqueExperimentKeyHandler.RemoveMeasurementsFromSheet();
            }
        }
        #endregion


        #region FillSingleColumnValue
        /// <summary>
        /// Fill a column with a single value
        /// </summary>
        /// <param name="values">array to be filled</param>
        /// <param name="column">at position</param>
        /// <param name="value">with value</param>
        /// <returns></returns>
        private object[,] FillSingleColumnValue(object[,] values, int column, string value)
        {

            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
            {
                values[i, column] = value;
            }
            return values;
        }
        #endregion

        #region FillTestData
        /// <summary>
        /// Fills the test data (just read from the DB via webservice) into the excel sheet
        /// </summary>
        /// <param name="testData"></param>
        /// <param name="writeMeasurements"></param>
        private void FillTestData(Lib.Testdata testData, bool writeMeasurements)
        {
            Globals.PDCExcelAddIn.EventsEnabled = false;
            bool screenUpdating = Globals.PDCExcelAddIn.Application.ScreenUpdating;
            Globals.PDCExcelAddIn.Application.ScreenUpdating = false;
            MultipleMeasurementTableHandler measurementHandler = null;
            try
            {
                myMainPDCListObject.ClearContents();
                if (myMainPDCListObject.MeasurementColumn != null && myMainPDCListObject.MeasurementColumn.HasMultiMeasurementTableHandler)
                {
                    measurementHandler = myMainPDCListObject.MeasurementColumn.MultiMeasurementTableHandler;
                    if (!writeMeasurements)
                    {
                        measurementHandler.Deactivated = true;
                    }
                    else
                    {
                        measurementHandler.Deactivated = false;
                        measurementHandler.AddMissingTables(myMainPDCListObject);
                    }
                }
                myMainPDCListObject.ensureCapacity(testData.Experiments.Count + 1);

                object[,] values = new object[testData.Experiments.Count, myMainPDCListObject.DataRange.Columns.Count];
                int rowNumber = 0;
                int startRow = myMainPDCListObject.DataRange.Row;
                PDCLogger.TheLogger.LogStarttime("fillTestData", "filling testdata");
                Dictionary<ListColumn, int> columnMapping = myMainPDCListObject.CurrentListColumnPlacements();

                foreach (Lib.ExperimentData experimentData in testData.Experiments)
                {
                    if (experimentData is Lib.PlaceHolderExperiment)
                    {
                        continue;
                    }
                    Dictionary<int, Lib.TestVariableValue> testValues = ToDictionary(experimentData.GetExperimentValues());
                    foreach (ListColumn column in columnMapping.Keys)
                    {
                        Lib.TestVariable testVariable = column.TestVariable;
                        Lib.TestVariableValue value = testVariable == null ? null :
                          testValues.ContainsKey(testVariable.VariableId) ?
                          testValues[testVariable.VariableId] : null;
                        int position = columnMapping[column] + values.GetLowerBound(1);

                        if (column.ParamHandler != null)
                        {
                            if (column.ParamHandler is Predefined.RowNumberHandler)
                            {
                                column.ParamHandler.SetValue(myMainPDCListObject, values, rowNumber, position, column, experimentData, value);
                            }
                            continue;
                        }
                        if (testVariable == null)
                        {
                            FillUploadColumnValue(values, rowNumber, experimentData, column, position);
                            continue;
                        }

                        if (value == null)
                        {
                            values[rowNumber, position] = null;
                            continue;
                        }
                        if (column.IsHyperLink && testVariable.IsBinaryParameter())
                        {
                            values[rowNumber, position] = value.Filename;
                            continue;
                        }
                        if (testVariable.IsNumeric())
                        {
                            values[rowNumber, position] =
                              Lib.PDCConverter.Converter.NumericString2Double(value.ValueChar, value.Prefix, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                        }
                        else
                        {
                            if (value.Prefix != null && value.Prefix.Trim() != "")
                            {
                                values[rowNumber, position] = Lib.PDCConverter.Converter.MergeWithPrefix(value.Prefix, value.ValueChar);
                            }
                            else
                            {
                                values[rowNumber, position] = value.ValueChar;
                            }
                        }
                    }
                    if (rowNumber % 10 == 0) Globals.PDCExcelAddIn.SetStatusText("Processing experiment " + rowNumber + "/" + testData.Experiments.Count);
                    rowNumber++;
                }

                foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
                {
                    ListColumn column = pair.Key;
                    int position = pair.Value + values.GetLowerBound(1);
                    if (column.ParamHandler is Predefined.MultipleMeasurementTableHandler)
                    {
                        column.ParamHandler.SetValues(myMainPDCListObject, values, position, column, testData);
                    }
                    if (column.ParamHandler is Predefined.RowNumberHandler)
                    {
                        column.ParamHandler.SetValues(myMainPDCListObject, values, position, column, testData);
                    }
                    if (column.ParamHandler is Predefined.SingleMeasurementTableHandler)
                    {
                        if (writeMeasurements)
                        {
                            column.ParamHandler.SetValues(myMainPDCListObject, values, position, column, testData);
                        }
                        else
                        {
                            FillSingleColumnValue(values, position, UniqueExperimentKeyHandler.NOT_LOADED);
                        }
                    }
                }
                Globals.PDCExcelAddIn.SetStatusText("Writing Experiments");
                myMainPDCListObject.Values = values;
                ProcessBinaryDataLinks(testData, columnMapping, myMainPDCListObject.HiddenRows(), false);
            }
            finally
            {
                if (measurementHandler != null)
                {
                    measurementHandler.Deactivated = false;
                }
                Globals.PDCExcelAddIn.Application.ScreenUpdating = screenUpdating;
                Globals.PDCExcelAddIn.EventsEnabled = true;
                PDCLogger.TheLogger.LogStoptime("fillTestData", "filling testdata");
            }
        }
        #endregion

        #region FillUploadColumnValue
        /// <summary>
        /// Fills the upload info from the experiment data into the specified list column/row of the matrix.
        /// </summary>
        /// <param name="values"></param>
        /// <param name="rowNumber"></param>
        /// <param name="experimentData"></param>
        /// <param name="listColumn"></param>
        /// <param name="position"></param>
        private void FillUploadColumnValue(object[,] values, int rowNumber, Lib.ExperimentData experimentData, ListColumn listColumn, int position)
        {
            switch (listColumn.Name)
            {
                case PDCExcelConstants.COMPOUNDNO:
                    values[rowNumber, position] = experimentData.CompoundNo; break;
                case PDCExcelConstants.PREPARATIONNO:
                    values[rowNumber, position] = experimentData.PreparationNo; break;
                case PDCExcelConstants.MCNO:
                    values[rowNumber, position] = experimentData.MCNo; break;
                case PDCExcelConstants.UPLOAD_ID:
                    values[rowNumber, position] = experimentData.UploadId; break;
                case PDCExcelConstants.EXPERIMENT_NO:
                    values[rowNumber, position] = experimentData.ExperimentNo; break;
                case PDCExcelConstants.RESULTSTATUS:
                    values[rowNumber, position] = GetResultStatusString(experimentData.ResultStatus); break;
                case PDCExcelConstants.UPLOADDATE:
                    values[rowNumber, position] = experimentData.UploadDate.Value; break;
                case PDCExcelConstants.DATERESULT:
                    values[rowNumber, position] = experimentData.DateResult.Value; break;
                case PDCExcelConstants.PERSONID:
                    values[rowNumber, position] = experimentData.PersonId; break;
                case PDCExcelConstants.REPORT_TO_PIX:
                    values[rowNumber, position] = experimentData.ReportToPix; break;

            }
        }
        #endregion


        #region GetFileName
        /// <summary>
        /// Returns the file name for the given test variable in the experiment or null,
        /// if the variable is not specified within the data or is an empty string.
        /// </summary>
        /// <param name="experimentData"></param>
        /// <param name="testVariable"></param>
        /// <returns></returns>
        private string GetFileName(Lib.ExperimentData experimentData, Lib.TestVariable testVariable)
        {
            List<Lib.TestVariableValue> values = experimentData.GetExperimentValues();
            foreach (Lib.TestVariableValue value in values)
            {
                if (value.VariableId == testVariable.VariableId)
                {
                    if (value.Filename == null) return null;
                    if (value.Filename.Trim() == "") return null;

                    return Lib.Util.StreamUtil.GetShortFileName(value.Filename);
                }
            }
            return null;
        }
        #endregion

        #region GetHyperlinkInfo
        /// <summary>
        /// Returns the updated hyperlink data for all rows for the specified column
        /// </summary>
        /// <param name="testData">The updated test data</param>
        /// <param name="hiddenFlags">Which rows are to be ignored</param>
        /// <param name="pair">Columnmapping of a single list column</param>
        /// <returns>An array consisting of a pair of image url and image name</returns>
        private KeyValuePair<string, string>?[] GetHyperlinkInfo(Lib.Testdata testData, bool[] hiddenFlags, KeyValuePair<ListColumn, int> pair)
        {
            KeyValuePair<string, string>?[] links = new KeyValuePair<string, string>?[hiddenFlags.Length];
            int i = 0;
            foreach (Lib.ExperimentData experimentData in testData.Experiments)
            {
                if (experimentData is Lib.PlaceHolderExperiment)
                {
                    continue;
                }
                while (hiddenFlags[i + hiddenFlags.GetLowerBound(0)])
                {
                    i++;
                }
                string link = GetFileName(experimentData, pair.Key.TestVariable);
                if (link != null)
                {
                    string imageLink = CreateImageServletLink(experimentData, pair.Key.TestVariable);
                    if (imageLink != null)
                    {
                        links[i] = new KeyValuePair<string, string>(imageLink, link);
                    }

                }
                i++;
            }
            return links;
        }
        #endregion

        #region GetResultStatusString
        /// <summary>
        /// Returns a string representation of the specified ResultStatus value
        /// </summary>
        /// <param name="resultStatus"></param>
        /// <returns></returns>
        private object GetResultStatusString(decimal? resultStatus)
        {
            if (resultStatus == null)
            {
                return null;
            }
            int status = (int)resultStatus;
            switch (status)
            {
                case Lib.PDCConstants.RESULT_STATUS_EFFECT:
                    return "Effect";
                case Lib.PDCConstants.RESULT_STATUS_NO_EFFECT:
                    return "No Effect";
                default:
                    return resultStatus.Value;
            }
        }
        #endregion

        #region GetTestData
        /// <summary>
        /// Creates a Testdata object from the receiver and any associated measurement tables.
        /// </summary>
        /// <param name="setDefaults">Replaces empty parameter values with default values if set to true</param>
        /// <param name="leaveFlags"></param>
        /// <returns></returns>
        public Lib.Testdata GetTestData(bool setDefaults, bool[] leaveFlags, bool useExperimentNo, bool notFoundThenThrowException, bool isDataToProcessSelection)
        {
            Globals.PDCExcelAddIn.SetStatusText("Reading Testdata from sheet");
            PDCLogger.TheLogger.LogStarttime("GetTestData", "Reading Testdata");
            Lib.Testdata testData = new Lib.Testdata(myMainPDCListObject.Testdefinition);
            object[,] values = myMainPDCListObject.Values;

            Dictionary<ListColumn, int> columnMapping = myMainPDCListObject.CurrentListColumnPlacements();
            ExperimentAndMeasurementValues experimentAndMeasurementValues = new ExperimentAndMeasurementValues(values);
            myUniqueExperimentKeyHandler = null;
            if (myMainPDCListObject.HasMeasurementParamHandler)
            {
                if (myMainPDCListObject.MeasurementColumn.HasSingleMeasurementTableHandler)
                {
                    myUniqueExperimentKeyHandler = new UniqueExperimentKeyHandler(myMainPDCListObject, myMainPDCListObject.MeasurementColumn.SingleMeasurementTableHandler.MeasurementTable);
                    myUniqueExperimentKeyHandler.UseExperimentNo = useExperimentNo;
                }
            }

            //Check for measurement data
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {

                if (pair.Key.HasSingleMeasurementTableHandler)
                {
                    experimentAndMeasurementValues.measurementValues = new MeasurementSheetData();
                    experimentAndMeasurementValues.measurementValues.Range = pair.Key.SingleMeasurementTableHandler.SheetDataRange;
                    break;
                }
                if (pair.Key.HasMultiMeasurementTableHandler)
                {
                    experimentAndMeasurementValues.measurementValues = new MeasurementSheetData();
                    experimentAndMeasurementValues.measurementValues.Range = pair.Key.MultiMeasurementTableHandler.SheetDataRange;
                    break;
                }
            }
            testData.Tag = experimentAndMeasurementValues;

            if (setDefaults)
            {
                myMainPDCListObject.InitializeDefaultValues(values, leaveFlags);
            }



            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
            {
                if (leaveFlags != null && leaveFlags.Length > (i - values.GetLowerBound(0)) &&
                  leaveFlags[i - values.GetLowerBound(0)])
                {
                    Lib.ExperimentData placeHolder = CreateExperiment(myMainPDCListObject.Testdefinition, values, i, columnMapping, experimentAndMeasurementValues, setDefaults);
                    bool isEmpty = (placeHolder == null);
                    placeHolder = new Lib.PlaceHolderExperiment(myMainPDCListObject.Testdefinition);
                    if (isEmpty)
                    {
                        placeHolder.ExperimentNo = -666;
                    }
                    testData.Add(placeHolder);
                    continue;
                }
                Lib.ExperimentData experimentData = CreateExperiment(myMainPDCListObject.Testdefinition, values, i, columnMapping, experimentAndMeasurementValues, setDefaults);
                if (experimentData != null)
                {
                    testData.Add(experimentData);
                }
                else
                {
                    // markiere  Zeilen im Experiment mit -666, damit kein NOT_LOADED ins Sheet geschrieben wird!
                    Lib.ExperimentData placeHolder = new Lib.PlaceHolderExperiment(myMainPDCListObject.Testdefinition);
                    placeHolder.ExperimentNo = -666;
                    testData.Add(placeHolder);

                }
                //Globals.PDCExcelAddIn.SetStatusText("Collecting Experiment " + i + "/" + values.GetUpperBound(0));
            }
            if (myUniqueExperimentKeyHandler != null)
            {
                myUniqueExperimentKeyHandler.FindAllMeasurementRows(testData, notFoundThenThrowException, isDataToProcessSelection);
                myUniqueExperimentKeyHandler.WriteMeasurementLinks(null);
            }
            Globals.PDCExcelAddIn.SetStatusText(null);
            PDCLogger.TheLogger.LogStoptime("GetTestData", "Reading Testdata");


            return testData;
        }
        #endregion


        #region InitializeUploadColumns
        /// <summary>
        /// Initializes the upload column with the updated data from the test data
        /// </summary>
        /// <param name="testData">The test data containing updated values</param>
        /// <param name="hiddenFlags">Rows to be ignored</param>
        /// <param name="uploadColumnNumber">The column which should be processed</param>
        /// <param name="uploadColumn">The array with the current values from the sheet</param>
        /// <returns>true if the upload or experiment nos are changed</returns>
        private bool InitializeUploadColumns(Lib.Testdata testData, bool[] hiddenFlags, int uploadColumnNumber, object[,] uploadColumn)
        {
            int i = 0;
            bool isUploadColumnsInitialized = false;
            foreach (Lib.ExperimentData experimentData in testData.Experiments)
            {
                while (hiddenFlags[i + hiddenFlags.GetLowerBound(0)])
                {
                    i++;
                }
                if (i >= uploadColumn.GetUpperBound(0))
                {
                    break;
                }
                if (experimentData is Lib.PlaceHolderExperiment)
                {
                    i++;
                    continue;
                }
                if (i >= uploadColumn.GetUpperBound(0))
                {
                    break;
                }
                object value = null;
                switch (uploadColumnNumber)
                {
                    case Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER:
                        value = experimentData.CompoundNo; break;
                    case Lib.PDCConstants.C_ID_PREPARATIONNO:
                        value = experimentData.PreparationNo; break;
                    case Lib.PDCConstants.C_ID_MCNO:
                        value = experimentData.MCNo; break;
                    case Lib.PDCConstants.C_ID_UPLOAD_ID:
                        value = experimentData.UploadId;
                        isUploadColumnsInitialized = isUploadColumnsInitialized || value != null;
                        break;
                    case Lib.PDCConstants.C_ID_EXPERIMENTNO:
                        value = experimentData.ExperimentNo;
                        isUploadColumnsInitialized = isUploadColumnsInitialized || value != null;
                        break;
                    case Lib.PDCConstants.C_ID_RESULT_STATUS:
                        value = GetResultStatusString(experimentData.ResultStatus); break;
                    case Lib.PDCConstants.C_ID_PERSONID:
                        value = experimentData.PersonId;
                        break;
                    case Lib.PDCConstants.C_ID_PDC_ONLY_DATA:
                        value = experimentData.ReportToPix;
                        break;
                    case Lib.PDCConstants.C_ID_UPLOADDATE:
                        if (experimentData.UploadDate != null)
                            value = experimentData.UploadDate.Value;
                        //tmpValue = tmpExperiment.UploadDate == null ? null : tmpExperiment.UploadDate.Value.ToShortDateString();
                        break;
                    case Lib.PDCConstants.C_ID_DATE_RESULT:
                        if (experimentData.DateResult != null)
                            value = experimentData.DateResult.Value;
                        //tmpValue = tmpExperiment.DateResult == null ? null : tmpExperiment.DateResult.Value.ToShortDateString();

                        break;
                }
                uploadColumn[i + uploadColumn.GetLowerBound(0), uploadColumn.GetLowerBound(1)] = value;
                i++;
            }
            return isUploadColumnsInitialized;
        }
        #endregion




        #region ProcessBinaryDataLinks
        /// <summary>
        /// replaces local hyperlinks in test data with links to the pdc portal.
        /// </summary>
        /// <param name="testData">The uploaded testdata</param>
        /// <param name="columnMapping">Column mapping</param>
        /// <param name="hiddenFlags">Rows which should not be changed</param>
        /// <param name="mergeCurrentValues"></param>
        private void ProcessBinaryDataLinks(Lib.Testdata testData, Dictionary<ListColumn, int> columnMapping, bool[] hiddenFlags, bool mergeCurrentValues)
        {
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {
                if (pair.Key.IsHyperLink && pair.Key.TestVariable != null)
                {
                    KeyValuePair<string, string>?[] hyperLinks = GetHyperlinkInfo(testData, hiddenFlags, pair);
                    myMainPDCListObject.SetHyperLinkColumnValues(pair.Value, hyperLinks, false);
                }
            }
        }
        #endregion

        #region ProcessUploadChanges
        /// <summary>
        /// Updates the sheet data which may have been changed by the webservice.
        /// </summary>
        /// <param name="testData">The testdata containing updated values</param>
        /// <param name="processBinaryDataLink">Do not process binary data links (actually only when coming from pure validation)</param>
        /// <returns>true if the upload was successfull, false otherwise</returns>
        private bool ProcessUploadChanges(Lib.Testdata testData, bool processBinaryDataLink)
        {
            //Possibly changed: compound no, preparation no, mc no, experiment no
            Dictionary<ListColumn, int> columnMapping = myMainPDCListObject.CurrentListColumnPlacements();
            bool[] hiddenFlags = myMainPDCListObject.HiddenRows();
            int tmpCount = hiddenFlags.GetLength(0);
            bool uploadSucceeded = false;
            //Update SMT first, otherwise we cannot match the rows anymore
            if (myMainPDCListObject.HasMeasurementParamHandler && myMainPDCListObject.MeasurementColumn.HasSingleMeasurementTableHandler)
            {
                myUniqueExperimentKeyHandler.UpdateAfterUpload(testData);
            }
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {
                uploadSucceeded |= UpdateUploadColumns(testData, hiddenFlags, pair);
            }
            if (uploadSucceeded && processBinaryDataLink) //Upload succeeded, Replace local hyperlinks with hyperlinks to ImageServlet
            {
                myMainPDCListObject.AlreadyUploaded = true;
                ProcessBinaryDataLinks(testData, columnMapping, hiddenFlags, true);
            }
            return uploadSucceeded;
        }
        #endregion


        #region SetTestdata
        /// <summary>
        /// Initializes the table with the specified test data. Also initializes the
        /// measurement tables if applicable. Can only be executed on the top PDCListObject for a test.
        /// The test definition of the test data must be compatible to the test definition of the receiver.
        /// </summary>
        /// <param name="testData"></param>
        /// <param name="writeMeasurements"></param>
        public void SetTestdata(Lib.Testdata testData, bool writeMeasurements)
        {
            myMainPDCListObject.ResetMeasurementHyperLinks();
            myMainPDCListObject.WriteExcelDateFormat();
            FillTestData(testData, writeMeasurements);
            myMainPDCListObject.AlreadyUploaded = true;
        }
        #endregion

        #region SetMeasurementTestData
        /// <summary>
        /// Initializes the table with the specified test data. Also initializes the
        /// measurement tables if applicable. Can only be executed on the top PDCListObject for a test.
        /// The test definition of the test data must be compatible to the test definition of the receiver.
        /// </summary>
        public void SetMeasurementTestData(Lib.Testdata testData)
        {
            Globals.PDCExcelAddIn.EventsEnabled = false;
            try
            {
                PDCLogger.TheLogger.LogStarttime("SetMeasurementTestData", "SetMeasurementTestData");
                ListColumn measurementListColumn = myMainPDCListObject.MeasurementColumn;
                int measurementColumnPosition = myMainPDCListObject.GetColumnIndex(PDCExcelConstants.MEASUREMENTS).Value;
                object[,] values = CreateSingleColumnValue(testData.Experiments, UniqueExperimentKeyHandler.NOT_LOADED);

                // can fill Values:
                measurementListColumn.ParamHandler.SetValues(myMainPDCListObject, values, 0, measurementListColumn, testData);
                //
                if (measurementListColumn.ParamHandler2 != null)
                {
                    measurementListColumn.ParamHandler2.SetValues(myMainPDCListObject, values, measurementColumnPosition, measurementListColumn, testData);
                }
                ExcelFilterStatus tmpExcelFilterStatus = ExcelUtils.TheUtils.CollectExcelFilters(myMainPDCListObject);

                myMainPDCListObject.Container.AutoFilterMode = false;
                myMainPDCListObject.SetColumnValues(measurementColumnPosition, values);
                ExcelUtils.TheUtils.SetExcelFilters(myMainPDCListObject, tmpExcelFilterStatus);
            }
            finally
            {
                Globals.PDCExcelAddIn.EventsEnabled = true;
                PDCLogger.TheLogger.LogStoptime("SetMeasurementTestData", "SetMeasurementTestData");
            }

        }
        #endregion

        #region ToDictionary
        /// <summary>
        /// Returns a map of variableid to test variable value
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        private Dictionary<int, Lib.TestVariableValue> ToDictionary(List<Lib.TestVariableValue> values)
        {
            Dictionary<int, Lib.TestVariableValue> dict = new Dictionary<int, BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();
            foreach (Lib.TestVariableValue value in values)
            {
                dict.Add(value.VariableId, value);
            }
            return dict;
        }
        #endregion


        #region UpdateFromUpload
        /// <summary>
        /// The Test data was uploaded and part of data may be changed or got initialized.
        /// These changes are incorporated into the sheet(s)
        /// </summary>
        /// <param name="testData"></param>
        public void UpdateFromUpload(Lib.Testdata testData)
        {
            if (testData == null || !testData.UploadChangeFlag)
            {
                return;
            }
            ProcessUploadChanges(testData, true);
            testData.UploadChangeFlag = false;
        }
        /// <summary>
        /// The Test data was uploaded and part of data may be changed or got initialized.
        /// These changes are incorporated into the sheet(s)
        /// </summary>
        /// <param name="testData"></param>
        public void UpdateFromUploadWithoutProcessingBinaryLinks(Lib.Testdata testData)
        {
            if (testData == null || !testData.UploadChangeFlag)
            {
                return;
            }
            ProcessUploadChanges(testData, false);
            testData.UploadChangeFlag = false;
        }
        #endregion

        #region UpdateUploadColumns
        /// <summary>
        /// Updates the upload columns of the sheet from the theTestdata 
        /// </summary>
        /// <param name="testData"></param>
        /// <param name="hiddenFlags"></param>
        /// <param name="columnPair"></param>
        /// <returns></returns>
        private bool UpdateUploadColumns(Lib.Testdata testData, bool[] hiddenFlags, KeyValuePair<ListColumn, int> columnPair)
        {
            bool uploadSucceeded = false;
            object[,] values;
            switch (columnPair.Key.Name)
            {
                case PDCExcelConstants.UPLOADDATE:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_UPLOADDATE, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.DATERESULT:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_DATE_RESULT, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.COMPOUNDNO:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.PREPARATIONNO:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_PREPARATIONNO, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.MCNO:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_MCNO, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.UPLOAD_ID:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    uploadSucceeded = InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_UPLOAD_ID, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.REPORT_TO_PIX:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    uploadSucceeded = InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_PDC_ONLY_DATA, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.PERSONID:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    uploadSucceeded = InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_PERSONID, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
                case PDCExcelConstants.EXPERIMENT_NO:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    uploadSucceeded = InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_EXPERIMENTNO, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    if (myUniqueExperimentKeyHandler != null)
                    {
                        myUniqueExperimentKeyHandler.WriteExperimentLinks();
                        myUniqueExperimentKeyHandler.WriteMeasurementLinks(null);
                        myUniqueExperimentKeyHandler.WriteExperimentNo(values);
                    }
                    break;
                case PDCExcelConstants.RESULTSTATUS:
                    values = myMainPDCListObject.GetColumnValues(columnPair.Value);
                    InitializeUploadColumns(testData, hiddenFlags, Lib.PDCConstants.C_ID_RESULT_STATUS, values);
                    myMainPDCListObject.SetColumnValues(columnPair.Value, values);
                    break;
            }
            return uploadSucceeded;
        }
        #endregion

        #endregion
    }
}
