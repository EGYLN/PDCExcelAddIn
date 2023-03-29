using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /// <summary>
    /// Implements the necessary logic to hold the measurement data in a single table.
    /// </summary>
    [Serializable]
    [ComVisible(false)]
    public class SingleMeasurementTableHandler : PredefinedParameterHandler
    {
        protected const string LIST_PREFIX = "SMea_";
        protected const string MEASUREMENT_DATA_RANGE = "Measurement_data";
        protected string myBaseSheetName;

        PDCListObject myPdcListObject;

        // id of the first sheet to save the measurements (usefull in the case there are more than 65536 rows)
        object myInitialSheetId;
        protected Lib.Testdefinition myTestDefinition;

        [NonSerialized]
        protected UniqueExperimentKeyHandler myUniqueExperimentKeyHandler;
        // keeps the sheetinfo on which the the measurements are shown
        protected SheetInfo mySheetInfo;

        #region constructor
        /// <summary>
        ///   The constructor for the horizontal measurement handler.
        /// </summary>
        /// <param name="testDef">
        ///   The test definition of the selected test.
        /// </param>
        public SingleMeasurementTableHandler(Lib.Testdefinition testDef)
        {
            myTestDefinition = testDef;
            mySheetInfo = null;
            myBaseSheetName = LIST_PREFIX + testDef.TestName;
            if (myBaseSheetName.Length > 15)
            {
                myBaseSheetName = myBaseSheetName.Substring(0, 15);
            }
            myBaseSheetName = ExcelUtils.GetSheetNameWithoutSpecialCharacter(myBaseSheetName);
            myBaseSheetName += "(" + testDef.TestNo + "_" + testDef.Version + ")";
        }
        #endregion

        #region events


        #endregion

        #region methods

        #region ClearContents
        /// <summary>
        ///   Clears the clearedValues array the the measurement table.
        /// </summary>
        /// <param name="pdcTable"></param>
        /// <param name="listColumn"></param>
        /// <param name="clearedValues"></param>
        public override void ClearContents(PDCListObject pdcTable, KeyValuePair<ListColumn, int> listColumn, object[,] clearedValues)
        {
            ClearMeasurementTables();
        }
        #endregion

        #region ClearMeasurementTables
        /// <summary>
        ///   Clears the single measurement table.
        /// </summary>
        protected virtual void ClearMeasurementTables()
        {
            Excel.Range sheetRange = SheetDataRange;
            if (sheetRange == null)
            {
                myPdcListObject.ClearContents();

                return;
            }
            sheetRange.set_Value(missing, null);

            sheetRange.Hyperlinks.Delete();
            ExcelUtils.TheUtils.ResetRowHeights(sheetRange);
        }
        #endregion

        #region InitializeSheet
        /// <summary>
        ///   Creates measurement tables for each new row.
        /// </summary>
        /// <param name="mainPdcListObject">
        ///   The PDCListObject of the main sheet.
        /// </param>
        public override void InitializeSheet(PDCListObject mainPdcListObject)
        {
            bool tmpEventsEnabled = Globals.PDCExcelAddIn.EventsEnabled;
            Globals.PDCExcelAddIn.EventsEnabled = false;

            try
            {
                SheetInfo mainSheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(mainPdcListObject.Container);
                PDCLogger.TheLogger.LogStarttime("Measurement.InitializeNewCells", "Initializing new cells");
                Excel.Worksheet currentSheet;

                if (mySheetInfo == null)
                {
                    Excel.Workbook workbook = (Excel.Workbook)mainPdcListObject.Container.Parent;
                    currentSheet = ExcelUtils.TheUtils.CreateNewSheet(workbook, myBaseSheetName, mainPdcListObject.Testdefinition);
                }
                else
                {
                    currentSheet = mySheetInfo.ExcelSheet;
                }

                PDCListObject currentPdcListObject = new PDCListObject(MEASUREMENT_DATA_RANGE, currentSheet, 3, 1, myTestDefinition, 2, false, false);
                mySheetInfo = Globals.PDCExcelAddIn.RegisterMeasurementTable(currentPdcListObject, currentSheet, mainSheetInfo);
                InitTableColumns(currentPdcListObject);
                if (currentPdcListObject.Container == null) //Loaded a saved PDC workbook
                {
                    if (Globals.PDCExcelAddIn.GetSheetInfo(myInitialSheetId) == null)
                    {
                        return;
                    }
                    currentPdcListObject.Container = Globals.PDCExcelAddIn.GetSheetInfo(myInitialSheetId).ExcelSheet;
                }
                // save sheetid  from first SheetInfo
                if (myInitialSheetId == null)
                {
                    myInitialSheetId = mySheetInfo.Identifier;
                }
                myPdcListObject = currentPdcListObject;
            }
            catch (Exception e)
            {
                ExceptionHandler.TheExceptionHandler.handleException(e, null);
            }
            finally
            {

                Globals.PDCExcelAddIn.EventsEnabled = tmpEventsEnabled;
                PDCLogger.TheLogger.LogStoptime("Measurement.InitializeNewCells", "Initiailized new cells");
            }
        }

        #endregion

        #region InitTableColumns
        /// <summary>
        ///   Initializes the Measurement table columns.
        /// </summary>
        /// <param name="pdcListObjectForMeasurements">
        ///   The PDC list object for the measurements.
        /// </param>
        protected void InitTableColumns(PDCListObject pdcListObjectForMeasurements)
        {
            List<ListColumn> columns = new List<ListColumn>();
            Lib.ClientConfiguration scheme = Globals.PDCExcelAddIn.ClientConfiguration;
            columns.Add(new ListColumn(PDCExcelConstants.COMPOUNDNO, "Compound No", Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER,
             scheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
            columns.Add(new ListColumn(PDCExcelConstants.PREPARATIONNO, "Preparation No", Lib.PDCConstants.C_ID_PREPARATIONNO,
              scheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
                        if (myTestDefinition.IsMcNoExpLevel)
            {
                columns.Add(new ListColumn(PDCExcelConstants.MCNO, "MC No", Lib.PDCConstants.C_ID_MCNO,
                    scheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
            }

            ListColumn experimentNoColumn = new ListColumn(PDCExcelConstants.EXPERIMENT_NO, "Experiment No", Lib.PDCConstants.C_ID_EXPERIMENTNO,
              scheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, true);
            experimentNoColumn.Hidden = true;
            columns.Add(experimentNoColumn);

            foreach (KeyValuePair<int, Lib.TestVariable> experimentLevelVariables in myTestDefinition.ExperimentLevelVariables)
            {
                ListColumn listColumn = PDCListObject.CreateColumn(experimentLevelVariables.Value);
                columns.Add(listColumn);
            }
            foreach (KeyValuePair<int, Lib.TestVariable> measurementVariables in myTestDefinition.MeasurementVariables)
            {
                ListColumn listColumn = PDCListObject.CreateColumn(measurementVariables.Value);
                columns.Add(listColumn);
            }
            ListColumn experimentColumn = new ListColumn(PDCExcelConstants.EXPERIMENTS, "Experiments", 9999, Globals.PDCExcelAddIn.ClientConfiguration[Lib.ClientConfiguration.HEADER_VARIABLE].SystemColor);
            experimentColumn.IsHyperLink = true;
            experimentColumn.ReadOnly = true;
            experimentColumn.Hidden = true;

            columns.Add(experimentColumn);
            pdcListObjectForMeasurements.AddColumns(columns);
        }
        #endregion


        #region SetValues
        /// <summary>
        /// Initializes all measurement tables at once with the new values
        /// </summary>
        /// <param name="mainPdcListObject"></param>
        /// <param name="values"></param>
        /// <param name="pos"></param>
        /// <param name="column"></param>
        /// <param name="testData"></param>
        public override void SetValues(PDCListObject mainPdcListObject, object[,] values, int pos, ListColumn column, Lib.Testdata testData)
        {

            PDCLogger.TheLogger.LogStarttime("SingleMeasurement.SetValues", "SingleMeasurement.SetValues");
            try
            {

                myUniqueExperimentKeyHandler = new UniqueExperimentKeyHandler(mainPdcListObject, myPdcListObject);

                int numOfMeasurements = testData.NumberOfAllMeasurementValues;
                if (numOfMeasurements > 65530)
                {
                    Exceptions.TooManyMeasurementtables ex = new Exceptions.TooManyMeasurementtables();
                    ex.AddArgument(numOfMeasurements);
                    throw ex;
                }

                // todo Check, if Delete has to be called when there is JUST a SingleMeasurementTableHandler

                mainPdcListObject.DeleteSMTMeasurementHyperLinks();

                myPdcListObject.ensureCapacity(numOfMeasurements + 1);

                ValueMatrixForSMT valueMatrixForSMT = new ValueMatrixForSMT(myPdcListObject.DataRange.Rows.Count, myPdcListObject.ColumnCount);

                // todo isntead of selfhandled experimentRow use RowNumberColumnHandler

                int experimentRow = 0;
                myUniqueExperimentKeyHandler.UseExperimentNo = true;

                foreach (Lib.ExperimentData experiment in testData.Experiments)
                {
                    experimentRow++;
                    valueMatrixForSMT.WriteValues(experiment, myPdcListObject, myUniqueExperimentKeyHandler, experimentRow);
                    if (experimentRow % 10 == 0) Globals.PDCExcelAddIn.SetStatusText("Processing measurements for experiment " + experimentRow + "/" + testData.Experiments.Count);
                }
                Globals.PDCExcelAddIn.SetStatusText("Calculating Measurement Links");
                myUniqueExperimentKeyHandler.WriteMeasurementLinks(values, pos);
                Globals.PDCExcelAddIn.SetStatusText("Writing Measurements");
                SheetDataRange.set_Value(Type.Missing, valueMatrixForSMT.Values);
                Globals.PDCExcelAddIn.SetStatusText("Calculating Experiment Links");
                myUniqueExperimentKeyHandler.WriteExperimentLinks();
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("SingleMeasurement.SetValues", "SingleMeasurement.SetValues");
            }
        }
        #endregion



        #endregion



        #region properties

        #region MeasurementTable
        internal PDCListObject MeasurementTable
        {
            get
            {
                return myPdcListObject;
            }
        }
        #endregion

        #region SheetDataRange
        /// <summary>
        /// Returns a range which contains the data range of the single measurement table.
        /// The range does not encompass the header range.
        /// Returns null if the operation is not supported (only sheets from the initial qa phase)
        /// </summary>
        internal virtual Excel.Range SheetDataRange
        {
            get
            {
                if (myPdcListObject == null)
                {
                    return null;
                }


                System.Drawing.Rectangle rectangle = myPdcListObject.Rectangle;

                return ExcelUtils.TheUtils.GetRange(myPdcListObject.Container, rectangle.Top, rectangle.X, rectangle.Bottom - 1, rectangle.Width);
             }
        }
        #endregion

        #region ToDict
        protected Dictionary<int, int> ToDict(Dictionary<ListColumn, int> dict)
        {
            Dictionary<int, int> dictionary = new Dictionary<int, int>();
            foreach (KeyValuePair<ListColumn, int> pair in dict)
            {
                if (pair.Key.TestVariable == null)
                {
                    continue;
                }
                dictionary.Add(pair.Key.TestVariable.VariableId, pair.Value);
            }
            return dictionary;
        }
        #endregion

        #endregion
    }
}
