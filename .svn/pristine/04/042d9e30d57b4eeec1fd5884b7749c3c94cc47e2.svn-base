using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using BBS.ST.BHC.BSP.PDC.Lib;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    [ComVisible(false)]
    public class UniqueExperimentKeyHandler
    {
        internal object missing = Type.Missing;

        PDCListObject myMainPDCListObject;
        PDCListObject mySmtPDCListObject;

        internal static string NOT_LOADED = "Not loaded";
        internal static string MEASUREMENTS = "Measurements";
        internal static string NONE_FOUND = "None found";

        Dictionary<string, int> myUniqueExperimentKey4MainSheetDict;
        Dictionary<string, List<int>> myUniqueExperimentKey4SMTSheetDict;
        Dictionary<string, bool> myIgnoreMeasurementCauseNotLoadedFromExperiment;
        bool myUseExperimentNo = false;

        public bool UseExperimentNo
        {
            get { return myUseExperimentNo; }
            set { myUseExperimentNo = value; }
        }

        #region constructor
        public UniqueExperimentKeyHandler(PDCListObject mainPDCListObject, PDCListObject smtPDCListObject)
        {
            myMainPDCListObject = mainPDCListObject;
            mySmtPDCListObject = smtPDCListObject;
            myUniqueExperimentKey4MainSheetDict = new Dictionary<string, int>();
            myUniqueExperimentKey4SMTSheetDict = new Dictionary<string, List<int>>();
            myIgnoreMeasurementCauseNotLoadedFromExperiment = new Dictionary<string, bool>();

        }
        #endregion

        #region methods

        #region AddMeasurementRow
        /// <summary>
        /// Adds a row from SMT Sheet identified by the given experimentData
        /// </summary>
        /// <param name="experimentData"></param>
        /// <param name="rowNumber"></param>
        public void AddMeasurementRow(Lib.ExperimentData experimentData, int rowNumber)
        {
            UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);

            List<int> rowsOfMeasurements;
            if (myUniqueExperimentKey4SMTSheetDict.ContainsKey(uniqueExperimentKey.Key))
            {
                rowsOfMeasurements = myUniqueExperimentKey4SMTSheetDict[uniqueExperimentKey.Key];

            }
            else
            {
                rowsOfMeasurements = new List<int>();
                myUniqueExperimentKey4SMTSheetDict.Add(uniqueExperimentKey.Key, rowsOfMeasurements);
                myIgnoreMeasurementCauseNotLoadedFromExperiment.Add(uniqueExperimentKey.Key, !experimentData.MeasurementsLoaded);
            }

            if (!rowsOfMeasurements.Contains(rowNumber))
            {
                rowsOfMeasurements.Add(rowNumber);
            }

        }
        #endregion

        #region CreateMeasurementLinkCell
        /// <summary>
        /// create at a  ExcelReferneceStyle Address-Range for a row like "A1:H1"
        /// </summary>
        /// <param name="startRow"></param>
        /// <param name="endRow"></param>
        /// <returns></returns>
        internal string CreateMeasurementLinkCell(int startRow, int endRow)
        {
            string dest = "";

            dest += ExcelUtils.TheUtils.GetAddressLocal(mySmtPDCListObject.Container, startRow, 1) + ":";
            dest += ExcelUtils.TheUtils.GetAddressLocal(mySmtPDCListObject.Container, endRow, mySmtPDCListObject.ColumnCount);
            return dest;
        }
        #endregion

        #region CreateUniqueExperimentKey

        /// <summary>
        /// Creates a UniqueExperimentKey Object from an Experiment
        /// </summary>
        /// <param name="sourceTable">Either the main pdc sheet or the single measurement table</param>
        /// <param name="rowValues">all values from the source table</param>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        private UniqueExperimentKey CreateUniqueExperimentKey(PDCListObject sourceTable, object[,] rowValues, int rowNumber)
        {
            UniqueExperimentKey uniqueExperimentKey = new UniqueExperimentKey(myUseExperimentNo);

            uniqueExperimentKey.IsCompoundNoExpLevel = myMainPDCListObject.Testdefinition.IsCompoundNoExpLevel;
            uniqueExperimentKey.IsPrepNoExpLevel = myMainPDCListObject.Testdefinition.IsPrepNoExpLevel;
            uniqueExperimentKey.IsMcNoExpLevel = myMainPDCListObject.Testdefinition.IsMcNoExpLevel;
            uniqueExperimentKey.CompoundNo = GetStringValueFromMatrix(rowValues, rowNumber, sourceTable.GetColumnIndex(PDCExcelConstants.COMPOUNDNO).Value + rowValues.GetLowerBound(1)).ToUpper();
            uniqueExperimentKey.ExperimentNo = GetLongValueFromMatrix(rowValues, rowNumber, sourceTable.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value + rowValues.GetLowerBound(1));
            uniqueExperimentKey.PreparationNo = GetStringValueFromMatrix(rowValues, rowNumber, sourceTable.GetColumnIndex(PDCExcelConstants.PREPARATIONNO).Value + rowValues.GetLowerBound(1)).ToUpper();
            int? mcNoIndex = sourceTable.GetColumnIndex(PDCExcelConstants.MCNO);
            if (uniqueExperimentKey.IsMcNoExpLevel && mcNoIndex != null)
            {
                uniqueExperimentKey.McNo =
                    GetStringValueFromMatrix(rowValues, rowNumber, mcNoIndex.Value + rowValues.GetLowerBound(1))
                        .ToUpper();

            }
            else
            {
                uniqueExperimentKey.McNo = UniqueExperimentKey.NullValue;
            }
            foreach (Lib.TestVariable testVariable in myMainPDCListObject.Testdefinition.ExperimentLevelVariables.Values)
            {
                int columns = sourceTable.GetColumnIndex(testVariable.VariableId).Value + rowValues.GetLowerBound(1);

                string value = (rowValues[rowNumber, columns] == null) ? UniqueExperimentKey.NullValue : rowValues[rowNumber, columns].ToString();
                uniqueExperimentKey.ExperimentLevelVariables.Add(testVariable.VariableName, value);
            }
            return uniqueExperimentKey;
        }
        /// <summary>
        /// Creates a UniqueExperimentKey Object from an Experiment
        /// </summary>
        /// <param name="rowValues">all values from the SimgleMeasurementSheet</param>
        /// <param name="rowNumber"></param>
        /// <returns></returns>
        private UniqueExperimentKey CreateUniqueExperimentKey(object[,] rowValues, int rowNumber)
        {
            return CreateUniqueExperimentKey(mySmtPDCListObject, rowValues, rowNumber);
        }

        /// <summary>
        /// Creates a UniqueExperimentKey Object from an Experiment
        /// </summary>
        /// <param name="experimentData"></param>
        /// <returns></returns>
        private UniqueExperimentKey CreateUniqueExperimentKey(Lib.ExperimentData experimentData)
        {
            UniqueExperimentKey uniqueExperimentKey = new UniqueExperimentKey(myUseExperimentNo);

            uniqueExperimentKey.IsCompoundNoExpLevel = myMainPDCListObject.Testdefinition.IsCompoundNoExpLevel;
            uniqueExperimentKey.IsPrepNoExpLevel = myMainPDCListObject.Testdefinition.IsPrepNoExpLevel;
            uniqueExperimentKey.IsMcNoExpLevel = myMainPDCListObject.Testdefinition.IsMcNoExpLevel;
            uniqueExperimentKey.CompoundNo = experimentData.CompoundNo == null ? UniqueExperimentKey.NullValue : experimentData.CompoundNo.ToUpper();
            uniqueExperimentKey.PreparationNo = (experimentData.PreparationNo == null) ? UniqueExperimentKey.NullValue : experimentData.PreparationNo.ToUpper();
            uniqueExperimentKey.ExperimentNo = experimentData.ExperimentNo;
            uniqueExperimentKey.McNo = experimentData.MCNo == null ? UniqueExperimentKey.NullValue : experimentData.MCNo.ToUpper();

            // Iterate through the ExperimentLevelVariableValues to fill the uniqueExperimentkey for each possible ExperimentLevelVariable (values need not to be set!!) 
            // using the TESTDEFINITION as not all defined ExperimentLevelVariables do have values
            foreach (Lib.TestVariable experimentLevelVariable in myMainPDCListObject.Testdefinition.ExperimentLevelVariables.Values)
            {
                uniqueExperimentKey.ExperimentLevelVariables.Add(experimentLevelVariable.VariableName, UniqueExperimentKey.NullValue);
            }

            foreach (Lib.TestVariableValue experimentLevelVariableValue in experimentData.GetExperimentLevelVariableValues())
            {
                Lib.TestVariable experimentLevelVariable = myMainPDCListObject.Testdefinition.ExperimentLevelVariables[experimentLevelVariableValue.VariableId];
                string value = experimentLevelVariableValue.ValueChar;
                if (!string.IsNullOrEmpty(value))
                {
                    uniqueExperimentKey.ExperimentLevelVariables[experimentLevelVariable.VariableName] = value;
                }
            }
            return uniqueExperimentKey;

        }
        #endregion

        #region CreateUniqueExperimentStringKey
        /// <summary>
        /// Creates a UniqueExperimentKey Object from an Experiment
        /// </summary>
        /// <param name="mainSheetRowNumber"></param>
        /// <returns></returns>
        private string CreateUniqueExperimentStringKey(int mainSheetRowNumber)
        {
            foreach (KeyValuePair<string, int> uniqueKeyString in myUniqueExperimentKey4MainSheetDict)
            {
                if (uniqueKeyString.Value == mainSheetRowNumber) return uniqueKeyString.Key;
            }
            return null;
        }
        #endregion

        #region FindAllMeasurementRowsForASingleExperiment

        /// <summary>
        /// Returns a list of all measurement rows which belong to the given experiment.
        /// </summary>
        internal List<int> FindAllMeasurementRowsForASingleExperiment(UniqueExperimentKey anExperimentKey, object[,] measurmentValues)
        {
            List<int> tmpRownumberList = new List<int>();
            for (int i = measurmentValues.GetLowerBound(0); i <= measurmentValues.GetUpperBound(0); i++)
            {
                UniqueExperimentKey tmpMeasurementKey = CreateUniqueExperimentKey(measurmentValues, i);
                if (anExperimentKey.Equals(tmpMeasurementKey))
                {
                    tmpRownumberList.Add(i);
                }
            }
            return tmpRownumberList;
        }

        public void FindAllMeasurementRowsForASingleExperiment(Lib.ExperimentData experimentData)
        {
            UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);
            List<Lib.TestVariableValue> measurements = new List<BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();
            object[,] measurementValues = mySmtPDCListObject.Values;
            int? measurementposition = 0;
            for (int i = measurementValues.GetLowerBound(0); i < measurementValues.GetUpperBound(0); i++)
            {
                UniqueExperimentKey uniqueExperimentKeyOnSMT = CreateUniqueExperimentKey(measurementValues, i);
                if (uniqueExperimentKeyOnSMT.Equals(uniqueExperimentKey))
                {
                    measurementposition++;
                    AddMeasurementRow(experimentData, i - measurementValues.GetLowerBound(0));

                    Dictionary<ListColumn, int> columnMapping = mySmtPDCListObject.CurrentListColumnPlacements();
                    foreach (ListColumn listColumn in columnMapping.Keys)
                    {
                        Lib.TestVariable testVariable = listColumn.TestVariable;
                        if (testVariable == null) continue;
                        if (testVariable.IsExperimentLevelReferenceForSMT) continue;
                        Lib.TestVariableValue testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
                        int columns = mySmtPDCListObject.GetColumnIndex(testVariable.VariableId).Value + measurementValues.GetLowerBound(1);
                        object singleMeasurementValue = measurementValues[i, columns];
                        if (testVariable.IsNumeric())
                        {
                            string prefix = null;
                            object valueObject = singleMeasurementValue;
                            if (singleMeasurementValue is string)
                            {
                                valueObject = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)singleMeasurementValue, out prefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                            }
                            Lib.PDCConverter.Converter.DoubleToString(valueObject, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue);
                            testVariableValue.Prefix = prefix;
                            testVariableValue.IsNummeric = true;
                        }
                        else
                        {
                            testVariableValue.ValueChar = "" + singleMeasurementValue;
                        }
                        testVariableValue.Position = measurementposition;
                        measurements.Add(testVariableValue);
                    }
                    experimentData.SetMeasurementValues(measurements);
                }
            }
        }
        #endregion

        private Lib.ExperimentData FindExperimentWithKey(Lib.Testdata testData, UniqueExperimentKey uniqueExperimentKeyOnSMT)
        {
            foreach (var experimentData in testData.Experiments)
            {
                if (experimentData is Lib.PlaceHolderExperiment) continue;
                UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);
                if (uniqueExperimentKey.Equals(uniqueExperimentKeyOnSMT))
                {
                    return experimentData;
                }
            }
            return null;
        }

        #region FindAllMeasurementRows

        public void FindAllMeasurementRows(Testdata testData, bool notFoundThenThrowException, bool isDataToProcessSelection)
        {

            //      UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);
            //      List<Lib.TestVariableValue> measurements = new List<BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue>();
            object[,] measurementValues = null;
            if (testData.Tag is ExperimentAndMeasurementValues)
            {
                MeasurementSheetData msd = ((ExperimentAndMeasurementValues)testData.Tag).measurementValues;
                if (msd != null)
                {
                    measurementValues = msd.Values;
                }
            }
            if (measurementValues == null)
            {
                measurementValues = mySmtPDCListObject.Values;
            }
            mySmtPDCListObject.InitializeDefaultValues(measurementValues, null);

            if (isDataToProcessSelection)
            {
                var selectedExperiments = testData.Experiments.Where(experiment => !(experiment is PlaceHolderExperiment)).ToList();
                if (!selectedExperiments.Any())
                {
                    throw new Exceptions.NoExperimentFoundForMeasurementException(1);//ToDo:
                }

                var measurementValueKeys = new Dictionary<int, UniqueExperimentKey>();
                for (int i = measurementValues.GetLowerBound(0); i < measurementValues.GetUpperBound(0); i++)
                {
                    var uniqueExperimentKeyOnSmt = CreateUniqueExperimentKey(measurementValues, i);
                    if (uniqueExperimentKeyOnSmt.IsNull()) continue;
                    measurementValueKeys.Add(i, uniqueExperimentKeyOnSmt);
                }

                foreach (var selectedExperiment in selectedExperiments)
                {
                    var selectedExperimentKey = CreateUniqueExperimentKey(selectedExperiment);
                    var selectedMeasurementValueKeys = measurementValueKeys.Where(keyValuePair => keyValuePair.Value.Equals(selectedExperimentKey)).ToList();
                    if (!selectedMeasurementValueKeys.Any() && notFoundThenThrowException)
                    {
                        PDCLogger.TheLogger.LogWarning(
                            "FindAllMeasurementRows", 
                            string.Format(
                                "No measurements for experiment ({0}) are found. CompNo:{1} PrepNo:{2}", 
                                selectedExperiment.ExperimentNo,
                                selectedExperiment.CompoundNo,
                                selectedExperiment.PreparationNo));
                    }
                    var measurementPositions = new Dictionary<string, int>();
                    foreach (var measurementValueKey in selectedMeasurementValueKeys)
                    {
                        if (measurementPositions.ContainsKey(selectedExperimentKey.Key))
                        {
                            measurementPositions[selectedExperimentKey.Key]++;
                        }
                        else
                        {
                            measurementPositions.Add(selectedExperimentKey.Key, 1);
                        }
                        AddMeasurementRow(selectedExperiment, measurementValueKey.Key - measurementValues.GetLowerBound(0));


                        Dictionary<ListColumn, int> columnMapping = mySmtPDCListObject.CurrentListColumnPlacements();
                        foreach (ListColumn listColumn in columnMapping.Keys)
                        {
                            TestVariable testVariable = listColumn.TestVariable;
                            if (testVariable == null) continue;
                            if (testVariable.IsExperimentLevelReferenceForSMT) continue;
                            var testVariableValue = new TestVariableValue(listColumn.TestVariable.VariableId);
                            int columns = mySmtPDCListObject.GetColumnIndex(testVariable.VariableId).Value + measurementValues.GetLowerBound(1);
                            object singleMeasurementValue = measurementValues[measurementValueKey.Key, columns];
                            if (testVariable.IsNumeric())
                            {
                                string prefix = null;
                                var singleMeasurementStringValue = singleMeasurementValue as string;
                                if (singleMeasurementStringValue != null)
                                {
                                    singleMeasurementValue = PDCConverter.Converter.RemoveWellKnownPrefix(
                                        singleMeasurementStringValue, 
                                        out prefix, 
                                        Globals.PDCExcelAddIn.PdcService.Prefixes());
                                }
                                PDCConverter.Converter.DoubleToString(singleMeasurementValue, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue);
                                testVariableValue.Prefix = prefix;
                                testVariableValue.IsNummeric = true;
                            }
                            else
                            {
                                testVariableValue.ValueChar = "" + singleMeasurementValue;
                            }
                            testVariableValue.Position = measurementPositions[selectedExperimentKey.Key];
                            List<TestVariableValue> testValues = selectedExperiment.GetMeasurementValues();
                            testValues.Add(testVariableValue);
                        }
                    }
                }              
            }
            else
            {
                Dictionary<string, int> measurementPositions = new Dictionary<string, int>();
                for (int i = measurementValues.GetLowerBound(0); i < measurementValues.GetUpperBound(0); i++)
                {
                    UniqueExperimentKey uniqueExperimentKeyOnSMT = CreateUniqueExperimentKey(measurementValues, i);
                    if (uniqueExperimentKeyOnSMT.IsNull()) continue;

                    var experimentData = FindExperimentWithKey(testData, uniqueExperimentKeyOnSMT);

                    // either the experiment is not to be saved or during check all a exception is to be thrown
                    if (experimentData == null)
                    {
                        if (notFoundThenThrowException)
                        {
                            PDCLogger.TheLogger.LogWarning(
                                "UniqueExperimentKeyHandler.FindAllMeasurementRows",
                                "No Experiment found. Key: " + uniqueExperimentKeyOnSMT.FormattedKey(" "));

                            throw new Exceptions.NoExperimentFoundForMeasurementException(i + 1);
                        }
                        continue;
                    }


                    UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);

                    if (uniqueExperimentKeyOnSMT.Equals(uniqueExperimentKey))
                    {
                        if (measurementPositions.ContainsKey(uniqueExperimentKey.Key))
                        {
                            measurementPositions[uniqueExperimentKey.Key]++;
                        }
                        else
                        {
                            measurementPositions.Add(uniqueExperimentKey.Key, 1);
                        }
                        AddMeasurementRow(experimentData, i - measurementValues.GetLowerBound(0));


                        Dictionary<ListColumn, int> columnMapping = mySmtPDCListObject.CurrentListColumnPlacements();
                        foreach (ListColumn listColumn in columnMapping.Keys)
                        {
                            Lib.TestVariable testVariable = listColumn.TestVariable;
                            if (testVariable == null) continue;
                            if (testVariable.IsExperimentLevelReferenceForSMT) continue;
                            Lib.TestVariableValue testVariableValue = new Lib.TestVariableValue(listColumn.TestVariable.VariableId);
                            int columns = mySmtPDCListObject.GetColumnIndex(testVariable.VariableId).Value + measurementValues.GetLowerBound(1);
                            object singleMeasurementValue = measurementValues[i, columns];
                            if (testVariable.IsNumeric())
                            {
                                string prefix = null;
                                object valueObject = singleMeasurementValue;
                                if (singleMeasurementValue is string)
                                {
                                    valueObject = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)singleMeasurementValue, out prefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                                }
                                Lib.PDCConverter.Converter.DoubleToString(valueObject, ExcelUtils.TheUtils.GetExcelNumberSeparators(), testVariableValue);
                                testVariableValue.Prefix = prefix;
                                testVariableValue.IsNummeric = true;
                            }
                            else
                            {
                                testVariableValue.ValueChar = "" + singleMeasurementValue;
                            }
                            testVariableValue.Position = measurementPositions[uniqueExperimentKey.Key];
                            List<Lib.TestVariableValue> testValues = experimentData.GetMeasurementValues();
                            testValues.Add(testVariableValue);
                        }
                    }
                }               
            }


        }
        #endregion

        #region GetLongValueFromMatrix
        private long? GetLongValueFromMatrix(object[,] matrix, int x, int y)
        {
            if (matrix[x, y] == null)
            {
                return null;
            }
            return long.Parse(matrix[x, y].ToString());
        }
        #endregion

        #region GetRowFromMatrix
        private object[] GetRowFromMatrix(object[,] values, int rowNumber)
        {
            object[] row = new object[values.GetUpperBound(1) + 1];
            for (int i = 0; i < values.GetUpperBound(1); i++)
            {
                row[i] = values[rowNumber, i];
            }
            return row;
        }
        #endregion

        #region GetStringValueFromMatrix
        private string GetStringValueFromMatrix(object[,] matrix, int x, int y)
        {
            if (matrix[x, y] == null)
            {
                return UniqueExperimentKey.NullValue;
            }
            return matrix[x, y].ToString();
        }
        #endregion

        #region RemoveMeasurementsFromSheet
        public void RemoveMeasurementsFromSheet()
        {
            int measurementValueRow = mySmtPDCListObject.Rectangle.Top;

            int oldRowNumber = -99;
            int startRow = -99;

            foreach (KeyValuePair<string, List<int>> uniqueExperimentKey in myUniqueExperimentKey4SMTSheetDict)
            {
                foreach (int rowNumber in uniqueExperimentKey.Value)
                {

                    if (startRow == -99)
                    {
                        startRow = rowNumber;
                        oldRowNumber = rowNumber - 1;
                    }
                    if (oldRowNumber != rowNumber - 1)
                    {
                        Excel.Range sourceCell = ExcelUtils.TheUtils.GetRange(mySmtPDCListObject.Container,
                                    (Excel.Range)mySmtPDCListObject.Container.Cells[startRow + measurementValueRow, 1],
                                    (Excel.Range)mySmtPDCListObject.Container.Cells[oldRowNumber + measurementValueRow, 1]).EntireRow;
                        sourceCell.Hyperlinks.Delete();
                        sourceCell.Delete(Type.Missing);
                        measurementValueRow = measurementValueRow - (oldRowNumber - startRow + 1);
                        oldRowNumber = rowNumber;
                        startRow = rowNumber;

                    }
                    else
                    {
                        oldRowNumber = rowNumber;
                    }
                }
            }
            if (startRow != -99)
            {
                Excel.Range sourceCell = ExcelUtils.TheUtils.GetRange(mySmtPDCListObject.Container,
                    (Excel.Range)mySmtPDCListObject.Container.Cells[startRow + measurementValueRow, 1],
                    (Excel.Range)mySmtPDCListObject.Container.Cells[oldRowNumber + measurementValueRow, 1]).EntireRow;
                sourceCell.Hyperlinks.Delete();
                sourceCell.Delete(Type.Missing);
            }
            myUniqueExperimentKey4SMTSheetDict.Clear();
            myIgnoreMeasurementCauseNotLoadedFromExperiment.Clear();
        }
        #endregion

        #region SetExperiment
        /// <summary>
        /// sets the position of the ExperimentKey with mainPDCListObject
        /// </summary>
        /// <param name="experimentData"></param>
        /// <param name="rowNumber"></param>
        public bool SetExperiment(Lib.ExperimentData experimentData, int rowNumber)
        {
            UniqueExperimentKey uniqueExperimentKey = CreateUniqueExperimentKey(experimentData);
            if (myUniqueExperimentKey4MainSheetDict.ContainsKey(uniqueExperimentKey.Key))
            {
                int offset = myMainPDCListObject.Rectangle.Top - 1;
                throw new Exceptions.AmbitiousExperimentsException(myUniqueExperimentKey4MainSheetDict[uniqueExperimentKey.Key] + offset, rowNumber + offset);
            }
            myUniqueExperimentKey4MainSheetDict.Add(uniqueExperimentKey.Key, rowNumber);
            return true;
        }
        #endregion

        #region WriteExperimentLinks
        /// <summary>
        /// write the experiment links on the SMT Sheets. 
        /// </summary>
        internal void WriteExperimentLinks()
        {
        }
        #endregion

        #region UpdateAfterUpload

        /// <summary>
        /// updates the measurement tables after an upload with the correct values for experimentno, compoundno, preparationno
        /// </summary>
        /// <param name="theExperimentData"></param>
        internal void UpdateAfterUpload(Lib.Testdata theUploadedData)
        {
            ExperimentAndMeasurementValues allValues = (ExperimentAndMeasurementValues)theUploadedData.Tag;
            //Get current matrix values
            object[,] experimentValues = allValues.experimentValues;
            object[,] measurementValues = allValues.measurementValues.Values;
            if (measurementValues == null)
            {
                measurementValues = mySmtPDCListObject.Values;
            }

            //For all experiments in theExperimentData find correspondend measurement rows and update column values
            int? tmpExperimentNoColumn = mySmtPDCListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO);
            int? tmpCompoundNoColumn = mySmtPDCListObject.GetColumnIndex(PDCExcelConstants.COMPOUNDNO);
            int? tmpPrepNoColumn = mySmtPDCListObject.GetColumnIndex(PDCExcelConstants.PREPARATIONNO);
            int tmpOffX = measurementValues.GetLowerBound(1);

            int row = experimentValues.GetLowerBound(0) - 1;
            foreach (Lib.ExperimentData tmpExperiment in theUploadedData.Experiments)
            {
                ++row;
                if (tmpExperiment is Lib.PlaceHolderExperiment)
                {
                    continue;
                }
                //create experiment key
                if (tmpExperiment.MeasurementsLoaded)
                {
                    UniqueExperimentKey tmpExperimentKey = CreateUniqueExperimentKey(myMainPDCListObject, experimentValues, row);
                    //get measurement rows for experiment key
                    List<int> tmpMeasurementrows = FindAllMeasurementRowsForASingleExperiment(tmpExperimentKey, measurementValues);
                    //update experimentno column, prepno column, compoundno column
                    foreach (int tmpRow in tmpMeasurementrows)
                    {
                        if (tmpExperimentNoColumn != null)
                        {
                            measurementValues[tmpRow, tmpExperimentNoColumn.Value + tmpOffX] = tmpExperiment.ExperimentNo;
                        }
                        if (tmpCompoundNoColumn != null)
                        {
                            measurementValues[tmpRow, tmpCompoundNoColumn.Value + tmpOffX] = tmpExperiment.CompoundNo;
                        }
                        if (tmpPrepNoColumn != null)
                        {
                            measurementValues[tmpRow, tmpPrepNoColumn.Value + tmpOffX] = tmpExperiment.PreparationNo;
                        }
                    }
                }
            }
            //Write back matrix into sheet
            //(A) May be it would be better to update selected columns only?
            mySmtPDCListObject.Values = measurementValues;
        }

        /// <summary>
        /// write the experiment No on the SMT Sheets. Every row is one experiment. so many values COULD be written for one Experiment!
        /// </summary>
        /// <param name="experimentNos"></param>
        internal void WriteExperimentNo(object[,] experimentNos)
        {
            int experimentNoPos = mySmtPDCListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value + 1;
            int smtsOffset = mySmtPDCListObject.Rectangle.Top;
            int smtsRangeRows = mySmtPDCListObject.Rectangle.Size.Height;
            object[,] experimentNosForMeasurementSheet = new object[smtsRangeRows, 1];
            object[,] experimentNosForMeasurementSheetDefault = mySmtPDCListObject.GetColumnValues(experimentNoPos - 1);
            for (int i = experimentNosForMeasurementSheet.GetLowerBound(0); i < experimentNosForMeasurementSheet.GetUpperBound(0); i++)
            {
                experimentNosForMeasurementSheet[i, 0] = experimentNosForMeasurementSheetDefault[i + 1, 1];
            }
            mySmtPDCListObject.GetColumnValues(experimentNoPos);
            for (int i = experimentNos.GetLowerBound(0); i < experimentNos.GetUpperBound(0); i++)
            {
                if (experimentNos[i, 1] == null) continue;
                string uniqueExperimentStringKey = CreateUniqueExperimentStringKey(i);
                if (uniqueExperimentStringKey == null) continue;
                if (myUniqueExperimentKey4SMTSheetDict.ContainsKey(uniqueExperimentStringKey))
                {
                    List<int> destRows = myUniqueExperimentKey4SMTSheetDict[uniqueExperimentStringKey];
                    foreach (int destRow in destRows)
                    {
                        experimentNosForMeasurementSheet[destRow, 0] = experimentNos[i, 1];
                    }
                }

            }
            mySmtPDCListObject.SetColumnValues(mySmtPDCListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value, experimentNosForMeasurementSheet);
        }
        #endregion
        #region WriteMeasurementLinks
        // this function simulates the optional parameter for POS!
        internal void WriteMeasurementLinks(object[,] values)
        {
            WriteMeasurementLinks(values, 0);
        }
        internal void WriteMeasurementLinks(object[,] values, int pos)
        {
            int experimentValueRow = myMainPDCListObject.Rectangle.Top;
            int measuremntCol = myMainPDCListObject.GetColumnIndex(PDCExcelConstants.MEASUREMENTS).Value;
            int measurementValueRow = mySmtPDCListObject.Rectangle.Top;
            Excel.Range sourceCell;
            foreach (KeyValuePair<string, int> uniqueKeyString in myUniqueExperimentKey4MainSheetDict)
            {

                if (myUniqueExperimentKey4SMTSheetDict.ContainsKey(uniqueKeyString.Key))
                {
                    if (myIgnoreMeasurementCauseNotLoadedFromExperiment[uniqueKeyString.Key]) continue;
                    List<int> rowNumbers = myUniqueExperimentKey4SMTSheetDict[uniqueKeyString.Key];
                    string containerName = "'" + mySmtPDCListObject.Container.Name + "'!";
                    string seperator = "";
                    string dest = "";

                    // todo clever calculation of Range in an own Function
                    int oldRowNumber = -99;
                    int startRow = -99;
                    foreach (int rowNumber in rowNumbers)
                    {
                        if (startRow == -99)
                        {
                            startRow = rowNumber;
                            oldRowNumber = rowNumber - 1;
                        }

                        if (oldRowNumber != rowNumber - 1)
                        {
                            dest += seperator + containerName + CreateMeasurementLinkCell(startRow + measurementValueRow, oldRowNumber + measurementValueRow);
                            seperator = ";";
                            oldRowNumber = rowNumber;
                            startRow = rowNumber;
                        }
                        else
                        {
                            oldRowNumber = rowNumber;
                        }
                    }
                    if (startRow != -99)
                    {
                        dest += seperator + containerName + CreateMeasurementLinkCell(startRow + measurementValueRow, oldRowNumber + measurementValueRow);
                    }

                    sourceCell = (Excel.Range)(myMainPDCListObject.Container.Cells[uniqueKeyString.Value + experimentValueRow - 1, measuremntCol + 1]);
                    sourceCell.Hyperlinks.Delete();

                    myMainPDCListObject.Container.Hyperlinks.Add(sourceCell, "", dest, missing, MEASUREMENTS);
                    if (values != null)
                    {
                        values[uniqueKeyString.Value - 1, pos] = MEASUREMENTS;
                    }
                }
                else
                {
                    if (values != null)
                    {
                        values[uniqueKeyString.Value - 1, pos] = NONE_FOUND;
                    }
                }
            }
        }
        #endregion

        #endregion
    }

}
