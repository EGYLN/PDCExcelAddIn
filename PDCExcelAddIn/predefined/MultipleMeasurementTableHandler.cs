using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    [Serializable]
    [ComVisible(false)]
    public class MultipleMeasurementTableHandler : PredefinedParameterHandler
    {
        protected const string ALLOW_EDIT_RANGE_NAME = "Mdata";
        protected const string LIST_PREFIX = "Meas_";
        protected const string PROTECT_PASSWORD = "pdcPRotectME";

        protected string myBaseSheetName;
        private bool myDeactivated = false;
        private int? myDeactivatedStartRow = null;

        MeasurementPDCListObject myFirstTable;
        object myInitialSheetId;
        protected int myNrOfTables;
        protected int myNrOfSheets;
        protected Lib.Testdefinition myTestDefinition;

        /// <summary>
        /// Mapping from measurement table no to PDCListObject
        /// </summary>
        protected IDictionary<int, PDCListObject> myMeasurementTableMap = new Dictionary<int, PDCListObject>();
        protected IDictionary<int, SheetInfo> mySheets = new Dictionary<int, SheetInfo>();

        /// <summary>
        /// Measurement list range name to sheet
        /// </summary>
        protected IDictionary<string, SheetInfo> tableLinks = new Dictionary<string, SheetInfo>();

        #region constructor

        /// <summary>
        /// The constructor.
        /// </summary>
        /// <param name="testDefinition"></param>
        public MultipleMeasurementTableHandler(Lib.Testdefinition testDefinition)
        {
            myTestDefinition = testDefinition;
            myNrOfTables = 0;
            myNrOfSheets = 0;
            myBaseSheetName = LIST_PREFIX + testDefinition.TestName;
            if (myBaseSheetName.Length > 15)
            {
                myBaseSheetName = myBaseSheetName.Substring(0, 15);
            }
            myBaseSheetName = ExcelUtils.GetSheetNameWithoutSpecialCharacter(myBaseSheetName);
            myBaseSheetName += "(" + testDefinition.TestNo + "_" + testDefinition.Version + ")";
        }

        #endregion

        #region methods

        #region CheckAndFixFirstTable
        /// <summary>
        /// Fixes a problem from an older version, where the columns of the first table got wrong names
        /// </summary>
        private void CheckAndFixFirstTable()
        {
            if (myFirstTable == null)
            {
                return;
            }
            try
            {
                int min = int.MaxValue;
                IDictionary<ListColumn, int> columnMapping = myFirstTable.CurrentListColumnPlacements();
                foreach (int column in columnMapping.Values)
                {
                    min = Math.Min(min, column);
                }
                foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
                {
                    if (pair.Key.Name.EndsWith("_1"))
                    {
                        pair.Key.Name = pair.Key.Name.Substring(0, pair.Key.Name.Length - 2) + "_0";
                        Excel.Range tmpRange = (Excel.Range)myFirstTable.Container.Cells[myFirstTable.Rectangle.Y + (pair.Value - min), myFirstTable.Rectangle.X - 1];
                        myFirstTable.Container.Names.Add(pair.Key.Name, tmpRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
                        myFirstTable.ClearCurrentMapping();
                    }
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Exception while fixing column names of measurment template table", e);
            }
        }
        #endregion

        #region CleanupReferences
        /// <summary>
        /// Cleans up references by removing links to removed tables/sheets
        /// </summary>
        internal virtual void CleanupReferences()
        {
            PDCLogger.TheLogger.LogStarttime("MeasurementCleanUp", "MeasurementCleanUp");
            if (myFirstTable != null && !myFirstTable.ExcelTableExists())
            {
                SheetInfo initialSheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(myInitialSheetId);
                if (initialSheetInfo == null)
                {
                    return;
                }
                myFirstTable.Container = initialSheetInfo.ExcelSheet;
                if (!myFirstTable.ExcelTableExists())
                {
                    myFirstTable = null;
                }
            }
            PDCLogger.TheLogger.LogStoptime("MeasurementCleanUp", "MeasurementCleanUp");
        }
        #endregion

        #region ClearContents
        /// <summary>
        ///   Clears the clearedValues array the the measurement tables.
        /// </summary>
        /// <param name="pdcTable"></param>
        /// <param name="listColumn"></param>
        /// <param name="clearedValues"></param>
        public override void ClearContents(PDCListObject pdcTable, KeyValuePair<ListColumn, int> listColumn, object[,] clearedValues)
        {
            for (int i = clearedValues.GetLowerBound(0); i <= clearedValues.GetUpperBound(0); i++)
            {
                clearedValues[i, listColumn.Value] = "Measurement";
            }
            ClearMeasurementTables();
        }
        #endregion

        #region ClearMeasurementTables
        /// <summary>
        ///   Clears the list of measurement tables
        /// </summary>
        protected virtual void ClearMeasurementTables()
        {
            Excel.Range sheetRange = GetSheetRange(true);
            ExcelUtils.TheUtils.UnprotectSheet(myFirstTable.Container, PROTECT_PASSWORD);
            try
            {
                if (sheetRange == null)
                {
                    foreach (PDCListObject list in myMeasurementTableMap.Values)
                    {
                        list.ClearContents();
                    }
                    return;
                }
                sheetRange.set_Value(missing, null);
                sheetRange.Hyperlinks.Delete();
                //        ExcelUtils.TheUtils.ResetRowHeights(sheetRange);
            }
            finally
            {
                ExcelUtils.TheUtils.ProtectSheet(myFirstTable.Container, PROTECT_PASSWORD, sheetRange, ALLOW_EDIT_RANGE_NAME);
            }
        }
        #endregion

        #region ColumnDeleted
        /// <summary>
        /// Warns the user that this column should not be deleted.
        /// </summary>
        /// <param name="pdcListObject"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public override bool ColumnDeleted(PDCListObject pdcListObject, ListColumn column)
        {
            MessageBox.Show(string.Format(Properties.Resources.LIST_MEASUREMENT_COLUMN_DELETED, column.Name), Properties.Resources.MSG_INFO_TITLE);
            return true;
        }
        #endregion

        #region ColumnsNeeded
        /// <summary>
        ///   Returns the number of measurement variables.
        /// </summary>
        /// <returns>
        ///   The number of measurement variables.
        /// </returns>
        protected int ColumnsNeeded()
        {
            return myTestDefinition.MeasurementVariables.Count;
        }
        #endregion

        #region CompareColumns
        public static int CompareColumns(ListColumn column1, ListColumn column2)
        {
            int order = column1.Label.CompareTo(column2.Label);

            if (column1.TestVariable == null && column2.TestVariable != null)
            {
                return -1;
            }
            if (column1.TestVariable != null && column2.TestVariable == null)
            {
                return 1;
            }
            if (column1.TestVariable == null)
            {
                return order;
            }
            if (column1.TestVariable.VariableClass == column1.TestVariable.VariableClass)
            {
                return order;
            }
            if (column1.TestVariable.VariableClass == Lib.TestVariable.VAR_CLASS_VARIABLE)
            {
                return -1;
            }
            if (column1.TestVariable.VariableClass == Lib.TestVariable.VAR_CLASS_RESULT &&
              column2.TestVariable.VariableClass != Lib.TestVariable.VAR_CLASS_VARIABLE)
            {
                return -1;
            }
            return 1;
        }
        #endregion

        #region CreateTableName
        /// <summary>
        /// Creates the table name for the specified nr
        /// </summary>
        /// <param name="nr"></param>
        /// <returns></returns>
        protected string CreateTableName(int nr)
        {
            return LIST_PREFIX + myTestDefinition.TestNo + "_" + myTestDefinition.Version + "_" + nr;
        }
        #endregion

        #region Delete
        public override void Delete(ListColumn column, PDCListObject pdcList, bool completeList)
        {
            if (!completeList)
            {
                return;
            }
            Globals.PDCExcelAddIn.Application.DisplayAlerts = false;
            //Delete all PDCListObjects
            try
            {
                SheetInfo sheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(pdcList.Container);
                sheetInfo.ClearMeasurementTables();
                //Delete all created sheets
                Dictionary<Excel.Worksheet, Excel.Worksheet> deleted = new Dictionary<Microsoft.Office.Interop.Excel.Worksheet, Microsoft.Office.Interop.Excel.Worksheet>();
                foreach (Excel.Worksheet sheet in mySheets.Values)
                {
                    if (!deleted.ContainsKey(sheet))
                    {
                        deleted.Add(sheet, sheet);
                        SheetInfo measureSheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(sheet);
                        if (measureSheetInfo != null)
                        {
                            Globals.PDCExcelAddIn.RemoveSheetInfo(measureSheetInfo);
                        }
                        sheet.Delete();
                    }
                }
                myMeasurementTableMap.Clear();
                mySheets.Clear();
                tableLinks.Clear();
                myNrOfSheets = 0;
                myNrOfTables = 0;
            }
            finally
            {
                Globals.PDCExcelAddIn.Application.DisplayAlerts = true;
            }
        }
        #endregion

        #region DeleteMeasurementtable
        /// <summary>
        /// Deletes the specified measurement table
        /// Since MeasurementPDCListObject don't implement any event handling, we
        /// have to explicitly change the location of all sucessive tables.
        /// </summary>
        /// <param name="tableToSheetPair">Maps the measurementtable name to the sheetinfo to which the table belongs</param>
        /// <param name="measurementTable">The Measurementtable which is to be deleted</param>
        /// <returns>The nr of deleted rows</returns>
        protected virtual int DeleteMeasurementtable(KeyValuePair<string, SheetInfo> tableToSheetPair, PDCListObject measurementTable)
        {
            int lines = measurementTable.Delete();
            tableToSheetPair.Value.MeasurementTables.Remove(tableToSheetPair.Key);
            int row = measurementTable.Rectangle.Bottom;
            foreach (PDCListObject list in MeasurementTables)
            {
                System.Drawing.Rectangle rect = list.Rectangle;
                if (list is MeasurementPDCListObject && rect.Top > row)
                {
                    rect.Y -= lines;
                    list.Rectangle = rect;
                }
            }
            return lines;
        }

        #endregion

        #region GetMeasurementTable
        protected PDCListObject GetMeasurementTable(Excel.Worksheet sheet, int row, int pos)
        {
            PDCLogger.TheLogger.LogStarttime("MeasurementHandler.GetMeasurementTable", "Getting Measurement table for row " + row);
            try
            {
                SheetInfo sheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(sheet);
                Excel.Range cell = (Excel.Range)sheet.Cells[row, pos];
                System.Collections.IEnumerator enumerator = cell.Hyperlinks.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Excel.Hyperlink link = (Excel.Hyperlink)enumerator.Current;
                    string linkAddres = link.SubAddress;
                    if (linkAddres != null)
                    {
                        return sheetInfo.FindMeasurementTable(linkAddres);
                    }
                }
                return null;
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("MeasurementHandler.GetMeasurementTable", "Getting Measurement table for row " + row);
            }
        }
        #endregion

        #region GetNextTablePosition
        /// <summary>
        /// Calculates the start position of the next measurement table
        /// </summary>
        /// <returns></returns>
        protected int GetNextTablePosition()
        {
            int nextPos = 1;
            if (myFirstTable != null)
            {
                Excel.Range used = myFirstTable.Container.UsedRange;

                return used.Row + used.Rows.Count + 1;
            }
            foreach (PDCListObject list in myMeasurementTableMap.Values)
            {
                nextPos = Math.Max(nextPos, list.Rectangle.Bottom + 1);
            }
            return nextPos;
        }
        #endregion

        #region GetUsedSheetInfos
        public ICollection<SheetInfo> GetUsedSheetInfos()
        {
            return mySheets == null ? null : mySheets.Values;
        }
        #endregion

        #region AddMissingTables
        public void AddMissingTables(PDCListObject pdcList)
        {
            int? columnIdx = pdcList.GetColumnIndex(PDCExcelConstants.MEASUREMENTS);
            if (columnIdx == null)
            {
                return;
            }
            Range columnRange = pdcList.ColumnRange(columnIdx.Value, true);
            int startRow = columnRange.Row;
            int column = columnRange.Column;
            int stopRow = columnRange.Rows.Count + startRow;


            for (int i = startRow; i < stopRow; i++)
            {
                if (GetMeasurementTable(pdcList.Container, i, column) == null)
                {
                    Range range = (Range)pdcList.Container.Cells[i, column];
                    InitializeNewCells(range, pdcList);
                }
            }
        }
        #endregion
        #region InitializeNewCells

        /// <summary>
        /// Creates measurement tables for each new row
        /// </summary>
        /// <param name="range"></param>
        /// <param name="pdcList"></param>
        public override void InitializeNewCells(Excel.Range range, PDCListObject pdcList)
        {
            if (myDeactivated)
            {
                return;
            }
            bool tmpEventsEnabled = Globals.PDCExcelAddIn.EventsEnabled;
            Globals.PDCExcelAddIn.EventsEnabled = false;

            try
            {
                int rangeColumn = range.Column;
                Dictionary<ListColumn, int> columnMapping = pdcList.CurrentListColumnPlacements();
                Dictionary<ListColumn, int> measMapping = null;
                SheetInfo mainSheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(pdcList.Container);
                PDCLogger.TheLogger.LogStarttime("Measurement.InitializeNewCells", "Initializing new cells");
                int startRow = range.Row;
                int lastRow = range.Rows.Count - 1 + startRow;
                Excel.Worksheet currentSheet = null;
                SheetInfo currentSheetInfo = null;

                bool registerSheetInfo = false;
                if (myNrOfTables > 0)
                {
                    currentSheetInfo = mySheets[myNrOfTables];
                    currentSheet = currentSheetInfo.ExcelSheet;
                    ExcelUtils.TheUtils.UnprotectSheet(currentSheet, PROTECT_PASSWORD);
                }
                else
                {
                    Excel.Workbook wb = (Excel.Workbook)pdcList.Container.Parent;
                    currentSheet = ExcelUtils.TheUtils.CreateNewSheet(wb, myBaseSheetName + myNrOfSheets, pdcList.Testdefinition);
                    registerSheetInfo = true;
                    myNrOfTables = 1;
                    myNrOfSheets = 1;
                }
                int tableStartRow = GetNextTablePosition();

                //Create new Table(s)
                for (int i = startRow; i <= lastRow; i++)
                {
                    Globals.PDCExcelAddIn.SetStatusText("Creating Measurement table " + CreateTableName(myNrOfTables));
                    PDCLogger.TheLogger.LogStarttime("Measurement.CreateTable", "Initializing new measurement table");
                    if (tableStartRow + ColumnsNeeded() > 65530)
                    {
                        throw new Exceptions.TooManyMeasurementtables();
                    }

                    MeasurementPDCListObject measurementPDCListObject = null;
                    if (myFirstTable == null)
                    {
                        measurementPDCListObject = new MeasurementPDCListObject(CreateTableName(0), currentSheet, tableStartRow, 4, myTestDefinition, true);
                        Globals.PDCExcelAddIn.RegisterMeasurementTable(measurementPDCListObject, currentSheet, mainSheetInfo);
                        myFirstTable = measurementPDCListObject;
                        InitTableColumns(measurementPDCListObject, 0);
                        myFirstTable.ListRangeByName.EntireRow.Hidden = true;
                        tableStartRow += myFirstTable.ColumnCount + 1;
                    }
                    else if (myFirstTable.Container == null) //Loaded a saved PDC workbook
                    {
                        SheetInfo initialSheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(myInitialSheetId);
                        if (initialSheetInfo == null)
                        {
                            return;
                        }
                        myFirstTable.Container = initialSheetInfo.ExcelSheet;
                    }

                    if (measMapping == null)
                    {
                        measMapping = myFirstTable.CurrentListColumnPlacements();
                    }

                    measurementPDCListObject = myFirstTable.CopyMeasurementTable(CreateTableName(myNrOfTables), 4, tableStartRow, myNrOfTables, measMapping);
                    currentSheetInfo = Globals.PDCExcelAddIn.RegisterMeasurementTable(measurementPDCListObject, currentSheet, mainSheetInfo);
                    if (myInitialSheetId == null)
                    {
                        myInitialSheetId = currentSheetInfo.Identifier;
                    }
                    if (registerSheetInfo && !mySheets.ContainsKey(myNrOfTables))
                    {
                        mySheets.Add(myNrOfTables, currentSheetInfo);
                    }
                    string measurementPDCListRangeName = measurementPDCListObject.ListRangeName;
                    if (tableLinks.ContainsKey(measurementPDCListRangeName))
                    {
                        tableLinks.Remove(measurementPDCListRangeName);
                    }
                    tableLinks.Add(measurementPDCListRangeName, currentSheetInfo);
                    myNrOfTables++;
                    mySheets[myNrOfTables] = currentSheetInfo;
                    myMeasurementTableMap[myNrOfTables] = measurementPDCListObject;
                    string name = myMeasurementTableMap[myNrOfTables].ListRangeName;
                    int tmpBackLinkOffset = 1; // xxx =1
                    // megahack: wenn es nur eine Measurementspalte gibt, kommen die BackLinkInfos in die gleiche Zeile:
                    if (myFirstTable.ColumnCount == 1) tmpBackLinkOffset = 0;
                    SetBackLink(pdcList, tableStartRow + tmpBackLinkOffset, i, measurementPDCListObject, columnMapping);
                    Excel.Range sourceCell = (Excel.Range)(pdcList.Container.Cells[i, rangeColumn]);
                    pdcList.Container.Hyperlinks.Add(sourceCell, "", measurementPDCListObject.ListRangeName, missing, "Measurement");
                    PDCLogger.TheLogger.LogStoptime("Measurement.CreateTable", "Initialized new measurement table");
                    Globals.PDCExcelAddIn.SetStatusText(null);
                    tableStartRow += myFirstTable.ColumnCount + 1;
                }
                ExcelUtils.TheUtils.ProtectSheet(currentSheet, PROTECT_PASSWORD, GetSheetRange(true), ALLOW_EDIT_RANGE_NAME);
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
        /// Initializes the Measurement table columns
        /// </summary>
        /// <param name="measurementTable"></param>
        /// <param name="tableNr"></param>
        protected void InitTableColumns(PDCListObject measurementTable, int tableNr)
        {
            List<ListColumn> columns = new List<ListColumn>();
            foreach (KeyValuePair<int, Lib.TestVariable> variable in myTestDefinition.MeasurementVariables)
            {
                ListColumn column = PDCListObject.CreateColumn(variable.Value, tableNr);
                columns.Add(column);
            }
            columns.Sort(MultipleMeasurementTableHandler.CompareColumns);
            measurementTable.AddColumns(columns);
        }
        #endregion

        #region MaxColumns
        private int MaxColumns(List<Lib.TestVariableValue> measurements)
        {
            int max = 1;
            foreach (Lib.TestVariableValue variableValue in measurements)
            {
                if (variableValue.Position == null || variableValue.Position < 1)
                {
                    PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL,
                      "Got Measurement without position: for variable " + variableValue.VariableId);
                    continue;
                }
                if (variableValue.Position > max)
                {
                    max = variableValue.Position.Value;
                }
            }
            return max;
        }
        #endregion

        #region RemoveUnreferencedTables
        /// <summary>
        /// Searches for unreferenced measurement tables which will be deleted.
        /// Is called when rows in the main table were possibly deleted.
        /// </summary>
        /// <param name="mainList"></param>
        public virtual void RemoveUnreferencedTables(PDCListObject mainList, int columnIndex)
        {
            //Get all measurement hyperlinks in the table
            PDCLogger.TheLogger.LogStarttime("MeasurementHandler.RemoveUnreferencedTables", "Searching unreferenced tables");
            try
            {
                Excel.Range columnRange = mainList.ColumnRange(columnIndex, true);
                Dictionary<string, string> links = new Dictionary<string, string>();
                System.Collections.IEnumerator enumerator = columnRange.Hyperlinks.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    Excel.Hyperlink link = (Excel.Hyperlink)enumerator.Current;
                    string linkTarget = link.SubAddress;
                    if (!links.ContainsKey(linkTarget))
                    {
                        links.Add(linkTarget, linkTarget);
                    }
                }
                SheetInfo sheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(mainList.Container);

                Dictionary<Excel.Worksheet, Excel.Worksheet> unprotected = new Dictionary<Excel.Worksheet, Excel.Worksheet>();

                IDictionary<string, string> deleted = new Dictionary<string, string>();
                foreach (KeyValuePair<string, SheetInfo> pair in tableLinks)
                {
                    if (!links.ContainsKey(pair.Key))
                    {
                        PDCListObject list = pair.Value.MeasurementTables[pair.Key];
                        if (list != null)
                        {
                            if (!unprotected.ContainsKey(list.Container))
                            {
                                unprotected.Add(list.Container, list.Container);
                                ExcelUtils.TheUtils.UnprotectSheet(list.Container, PROTECT_PASSWORD);
                            }
                            DeleteMeasurementtable(pair, list);
                        }
                        deleted.Add(pair.Key, pair.Key);
                    }
                }
                foreach (Excel.Worksheet sheet in unprotected.Keys)
                {
                    ExcelUtils.TheUtils.ProtectSheet(sheet, PROTECT_PASSWORD, GetSheetRange(true), ALLOW_EDIT_RANGE_NAME);
                }
                Dictionary<int, int> keys = new Dictionary<int, int>();
                foreach (KeyValuePair<int, PDCListObject> pair in myMeasurementTableMap)
                {
                    if (pair.Value != null && deleted.ContainsKey(pair.Value.ListRangeName) && !keys.ContainsKey(pair.Key))
                    {
                        keys.Add(pair.Key, pair.Key);
                    }
                }
                foreach (int key in keys.Keys)
                {
                    myMeasurementTableMap.Remove(key);
                    mySheets.Remove(key);
                }

                foreach (string name in deleted.Keys)
                {
                    tableLinks.Remove(name);
                }
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("MeasurementHandler.RemoveUnreferencedTables", "Searching unreferenced tables");
            }
        }
        #endregion

        #region SetBackLink

        /// <summary>
        /// Sets via formular the the RowNumber and CompoundNo to the left of a MeasurementTable
        /// </summary>
        /// <param name="mainTable"></param>
        /// <param name="tableStartRow"></param>
        /// <param name="targetRow"></param>
        /// <param name="measurementTable"></param>
        /// <param name="columnMapping"></param>
        protected void SetBackLink(PDCListObject mainTable, int tableStartRow, int targetRow, PDCListObject measurementTable, Dictionary<ListColumn, int> columnMapping)
        {
            string targetSheetName = mainTable.Container.Name ?? "";
            int startColumn = mainTable.DataRange.Column;
            Excel.Range cells = ExcelUtils.TheUtils.GetRange(measurementTable.Container, tableStartRow, 1, tableStartRow, 2);
            object[] formula = new object[2];
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {
                if (pair.Key == mainTable.RowNumberColumn)
                {
                    formula[0] = "='" + targetSheetName + "'!R" + targetRow + "C" + (pair.Value + startColumn);
                }
                else if (pair.Key.Name == PDCExcelConstants.COMPOUNDNO)
                {
                    formula[1] = "='" + targetSheetName + "'!R" + targetRow + "C" + (pair.Value + startColumn);
                }
            }

            cells.FormulaR1C1 = formula;
        }
        #endregion

        #region SetTestValuesInMatix
        /// <summary>
        /// Initializes the measurement values for the specified test and column mapping. 
        /// The caller specifies offsets within the value matrix.
        /// </summary>
        /// <param name="testDefinition"></param>
        /// <param name="measurements"></param>
        /// <param name="columnMapping"></param>
        /// <param name="values"></param>
        /// <param name="yOffset"></param>
        /// <param name="xOffset"></param>
        protected void SetTestValuesInMatix(Lib.Testdefinition testDefinition, List<Lib.TestVariableValue> measurements,
          Dictionary<int, int> columnMapping, object[,] values, int yOffset, int xOffset)
        {
            foreach (Lib.TestVariableValue variableValue in measurements)
            {
                if (!variableValue.Position.HasValue)
                {
                    continue;
                }
                Lib.TestVariable testVariable = testDefinition.MeasurementVariables[variableValue.VariableId];
                int row = columnMapping[testVariable.VariableId];
                int column = variableValue.Position.Value - 1;
                if ((values.GetUpperBound(0) < row + yOffset) ||
                  (values.GetUpperBound(1) < column + xOffset))
                {
                    PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_EXCEL,
                        "More measurementvalues than columns or rows in the table: \n" +
                        "Tablesize:(" + values.GetUpperBound(0) + "," + values.GetUpperBound(1) + ")\n" +
                        "Requested:(" + (row + yOffset) + "," + (column + xOffset) + " )");
                    continue;
                }

                if (testVariable.IsNumeric())
                {
                    values[row + yOffset, column + xOffset] =
                     Lib.PDCConverter.Converter.NumericString2Double(variableValue.ValueChar, variableValue.Prefix, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                }
                else
                {
                    values[row + yOffset, column + xOffset] = variableValue.ValueChar;
                }
            }
        }
        #endregion

        #region SetValue
        /// <summary>
        /// Empty implementation since this handler sets all values at once
        /// </summary>
        /// <param name="pdcTable"></param>
        /// <param name="values"></param>
        /// <param name="row"></param>
        /// <param name="pos"></param>
        /// <param name="column"></param>
        /// <param name="experiment"></param>
        /// <param name="value"></param>
        public override void SetValue(PDCListObject pdcTable, object[,] values, int row, int pos, ListColumn column, Lib.ExperimentData experiment,
          Lib.TestVariableValue value)
        {
            //Empty, doing a bulk operation in SetValues instead
            //base.SetValue(aPdcTable, theValues, aRow, aPos, aColumn, anExperiment, aValue);
        }
        #endregion

        #region SetValues
        /// <summary>
        /// Initializes all measurement tables at once with the new values
        /// </summary>
        /// <param name="experimentTable"></param>
        /// <param name="values"></param>
        /// <param name="pos"></param>
        /// <param name="column"></param>
        /// <param name="testData"></param>
        public override void SetValues(PDCListObject experimentTable, object[,] values, int pos, ListColumn column, Lib.Testdata testData)
        {
            int row = 0;
            PDCLogger.TheLogger.LogStarttime("MCHSetValues", "MCHSetValues");
            Excel.Worksheet sheet = experimentTable.Container;
            Excel.Range dataRange = experimentTable.DataRange;
            int rowStart = dataRange.Row;
            int colStart = dataRange.Column;
            int maxPosition = GetMaxPosition(testData);
            if (myFirstTable == null)
            {
                return;
            }
            Excel.Range sheetRange = GetSheetRange(Math.Min(PDCListObject.max_column, maxPosition + myFirstTable.Rectangle.X));
            System.Drawing.Rectangle sheetArea = new System.Drawing.Rectangle(sheetRange.Column, sheetRange.Row, sheetRange.Columns.Count, sheetRange.Rows.Count);
            object[,] sheetValues = new object[sheetArea.Height, sheetArea.Width];
            Dictionary<int, int> columnMapping = null;
            foreach (Lib.ExperimentData experiment in testData.Experiments)
            {
                values[row, pos] = "Measurement";
                if (experiment is Lib.PlaceHolderExperiment)
                {
                    row++;
                    continue; //No data ignore
                }
                List<Lib.TestVariableValue> testValues = experiment.GetMeasurementValues();
                if (testValues == null || testValues.Count == 0)
                {
                    row++;
                    continue;
                }
                PDCListObject measurementTable = GetMeasurementTable(sheet, row + rowStart, pos + colStart);
                if (measurementTable == null)
                {
                    continue;
                }
                if (columnMapping == null)
                { //Use firstTable? all column mapping are now the same
                    columnMapping = ToDict(measurementTable.CurrentListColumnPlacements());
                }
                int y = measurementTable.Rectangle.Y - sheetArea.Y;
                int x = 0;
                SetTestValuesInMatix(measurementTable.Testdefinition, testValues, columnMapping, sheetValues, y, x);
                row++;
            }
            sheetRange.set_Value(Type.Missing, sheetValues);
            PDCLogger.TheLogger.LogStoptime("MCHSetValues", "MCHSetValues");
        }

        /// <summary>
        /// Returns the maximum position found in the measurement values
        /// </summary>
        /// <param name="testData"></param>
        /// <returns></returns>
        private int GetMaxPosition(Lib.Testdata testData)
        {
            int maxPos = 1; //To avoid a potentially 0-sized matrix
            foreach (ExperimentData experiment in testData.Experiments)
            {
                List<TestVariableValue> measurements = experiment.GetMeasurementValues();
                if (measurements == null)
                {
                    continue;
                }
                foreach (TestVariableValue measurement in measurements)
                {
                    if (measurement.Position != null && measurement.Position > maxPos)
                    {
                        maxPos = measurement.Position.Value;
                    }
                }
            }
            return maxPos;
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

        #region properties

        #region MeasurementTables
        public ICollection<PDCListObject> MeasurementTables
        {
            get
            {
                return myMeasurementTableMap.Values;
            }
        }
        #endregion

        public bool Deactivated
        {
            get { return myDeactivated; }
            set { myDeactivated = value; }
        }
        public IDictionary<int, PDCListObject> MeasurementTablesDictionary
        {
            get
            {
                return myMeasurementTableMap;
            }
        }
        #region SheetDataRange
        /// <summary>
        /// Returns the used sheet range of the measurement sheet.
        /// </summary>
        /// <param name="maxColumn">The returned range ends at this column</param>
        /// <returns></returns>
        private Excel.Range GetSheetRange(int maxColumn)
        {
            return GetSheetRange(true, maxColumn);
        }
        /// <summary>
        /// Returns the used sheet range of the measurement sheet.
        /// </summary>
        /// <param name="completeRows">The range will contain complete sheet rows, if true. 
        /// Empty columns are otherwise omitted. 
        /// </param>
        /// <returns></returns>
        private Excel.Range GetSheetRange(bool completeRows) {
            return GetSheetRange(completeRows, PDCListObject.max_column-1);
        }
        private Excel.Range GetSheetRange(bool completeRows, int endColumn)
        {
            if (myFirstTable == null)
            {
                return null;
            }
            System.Drawing.Rectangle rectangle = myFirstTable.Rectangle;
            Excel.Range usedRange = myFirstTable.Container.UsedRange;
            if (!completeRows)
            {
                endColumn = usedRange.Column + usedRange.Columns.Count;
                if (endColumn < 10)
                {
                    endColumn++;
                }
            } 
            return ExcelUtils.TheUtils.GetRange(myFirstTable.Container, rectangle.Bottom + 1, rectangle.X, usedRange.Rows.Count, endColumn);

        }
        /// <summary>
        /// Returns a range which contains the data ranges of all measurement tables.
        /// The range does not encompass the header ranges.
        /// Returns null if the operation is not supported (only sheets from the initial qa phase)
        /// </summary>
        internal virtual Excel.Range SheetDataRange
        {
            get
            {
                return GetSheetRange(false);
            }
        }
        #endregion

        #endregion

        internal void setInternalState(string aBaseSheetName, int aNrOfTables, IDictionary<int, PDCListObject> aMeasurementTableMap, IDictionary<int, SheetInfo> aSheets, IDictionary<string, SheetInfo> aTableLinks, MeasurementPDCListObject aFirstTable, object anInitialSheetId)
        {
            myBaseSheetName = aBaseSheetName;
            myNrOfTables = aNrOfTables;
            myMeasurementTableMap = aMeasurementTableMap;
            mySheets = aSheets;
            myFirstTable = aFirstTable;
            myInitialSheetId = anInitialSheetId;
        }
    }
}
