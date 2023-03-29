// --------------------------------------------------------------------------------------------------------------------
// <copyright file="PDCListObject.cs" company="">
//   
// </copyright>
// <summary>
//   PDC Abstraction of an area consisting of a header and data rows
// </summary>
// --------------------------------------------------------------------------------------------------------------------


namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Drawing;
    using System.Globalization;
    using System.Runtime.InteropServices;
    using System.Runtime.Serialization;
    using System.Windows.Forms;
    using System.Xml.Serialization;
    using Exceptions;
    using Predefined;
    using Properties;
    using Lib;
    using Lib.Util;
    using Microsoft.Office.Interop.Excel;
    using Rectangle = System.Drawing.Rectangle;

    /// <summary>
    /// PDC Abstraction of an area consisting of a header and data rows
    /// </summary>
    [Serializable]
    [ComVisible(false)]
    public class PDCListObject : ISerializable
    {
        /// <summary>
        /// If this variable is set to true, columns and rows are swapped, so that
        /// the column headings are row headings and the data is filled from left to right.
        /// </summary>
        private bool mySwapColumnsAndRows;

        /// <summary>
        /// The missing.
        /// </summary>
        internal readonly object missing = Type.Missing;

        /// <summary>
        /// The debu g_ names.
        /// </summary>
        private static bool DEBUG_NAMES = Settings.Default.debugNamedRanges;

        /// <summary>
        /// The initia l_ ro w_ count.
        /// </summary>
        private const int INITIAL_ROW_COUNT = 6;

        /// <summary>
        /// The ro w_ co l_ prefix.
        /// </summary>
        protected const string ROW_COL_PREFIX = "RowCol_";

        // Maximum no of rows and column in Excel
        /// <summary>
        /// The max_column.
        /// </summary>
        internal const int max_column = 16384;

        /// <summary>
        /// The max_row.
        /// </summary>
        internal const int max_row = 1048576;

        /// <summary>
        /// The my name.
        /// </summary>
        private string myName;

        /// <summary>
        /// The my test data adapter.
        /// </summary>
        private TestDataTableAdapter myTestDataAdapter;

        /// <summary>
        /// The my validation handler.
        /// </summary>
        private readonly ValidationHandler myValidationHandler;

        /// <summary>
        /// The my row number column.
        /// </summary>
        private ListColumn myRowNumberColumn;

        /// <summary>
        /// The my measurement column.
        /// </summary>
        private ListColumn myMeasurementColumn;

        /// <summary>
        /// The my date list columns.
        /// </summary>
        private Dictionary<string, ListColumn> myDateListColumns;

        /// <summary>
        /// The my unique experiment key handler.
        /// </summary>
        private UniqueExperimentKeyHandler myUniqueExperimentKeyHandler;


        /// <summary>
        /// Current size and location of the whole list area.
        /// </summary>
        private Rectangle myRectangle;

        /// <summary>
        /// Unique identifier for a list
        /// </summary>
        private readonly string myIdentifier;

        /// <summary>
        /// name of the header range which contains all header columns of the list.
        /// </summary>
        private string myHeaderRangeName;

        /// <summary>
        /// name of the data range, which contains all active data rows of the list. 
        /// Note: The list contains an additional insert row.
        /// </summary>
        private string myDataRangeName;

        /// <summary>
        /// The collection of all columns
        /// </summary>
        private List<ListColumn> myColumns = new List<ListColumn>();

        /// <summary>
        /// The collection of all PDC columns which were deleted by the user.
        /// </summary>
        private readonly List<ListColumn> myDeletedColumns = new List<ListColumn>();

        /// <summary>
        /// The worksheet on which the list is displayed. 
        /// Should be replaced by a reference to the SheetInfo
        /// </summary>
        private volatile Worksheet myContainer;

        /// <summary>
        /// An associated Testdefinition
        /// </summary>
        private Testdefinition myTestDefinition;

        /// <summary>
        /// The my is already uploaded.
        /// </summary>
        private bool myIsAlreadyUploaded;

        /// <summary>
        /// The my sheet info.
        /// </summary>
        private SheetInfo mySheetInfo;

        /// <summary>
        /// Performance attribute to avoid recalculation of listcolumn mapping to excel column position 
        /// </summary>
        [NonSerialized]
        protected Dictionary<ListColumn, int> currentColumnMapping;

        #region constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        public PDCListObject()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="anOther">
        /// The an other.
        /// </param>
        public PDCListObject(PDCListObject anOther)
        {
            mySwapColumnsAndRows = anOther.mySwapColumnsAndRows;

            myName = anOther.myName;

            myValidationHandler = new ValidationHandler(this);
            myTestDataAdapter = new TestDataTableAdapter(this);

            myRowNumberColumn = anOther.myRowNumberColumn;
            myMeasurementColumn = anOther.myMeasurementColumn;

            myDateListColumns = anOther.myDateListColumns;

            myUniqueExperimentKeyHandler = anOther.myUniqueExperimentKeyHandler;
            myRectangle = anOther.myRectangle;

            myIdentifier = anOther.myIdentifier;
            myHeaderRangeName = anOther.myHeaderRangeName;
            myDataRangeName = anOther.myDataRangeName;
            myColumns = anOther.myColumns;
            myDeletedColumns = anOther.myDeletedColumns;
            myContainer = anOther.myContainer;

            myTestDefinition = anOther.myTestDefinition;
            myIsAlreadyUploaded = anOther.myIsAlreadyUploaded;

            mySheetInfo = anOther.mySheetInfo;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        protected PDCListObject(string name)
        {
            Name = name;
            Guid guid = Guid.NewGuid();
            myIdentifier = guid.ToString().Replace('-', '_');
            myTestDataAdapter = new TestDataTableAdapter(this);
            myValidationHandler = new ValidationHandler(this);
            myDateListColumns = new Dictionary<string, ListColumn>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="sheet">
        /// The sheet.
        /// </param>
        /// <param name="startRow">
        /// The start row.
        /// </param>
        /// <param name="startColumn">
        /// The start column.
        /// </param>
        /// <param name="testDefinition">
        /// The test definition.
        /// </param>
        public PDCListObject(string name, Worksheet sheet, int startRow, int startColumn, Testdefinition testDefinition)
            : this(name, sheet, startRow, startColumn, testDefinition, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="sheet">
        /// The sheet.
        /// </param>
        /// <param name="startRow">
        /// The start row.
        /// </param>
        /// <param name="startColumn">
        /// The start column.
        /// </param>
        /// <param name="testDefinition">
        /// The test definition.
        /// </param>
        /// <param name="swapColumnAndRowsFlag">
        /// The swap column and rows flag.
        /// </param>
        public PDCListObject(string name, Worksheet sheet, int startRow, int startColumn, Testdefinition testDefinition,
                             bool swapColumnAndRowsFlag)
            : this(name, sheet, startRow, startColumn, testDefinition, swapColumnAndRowsFlag, INITIAL_ROW_COUNT, true)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="sheet">
        /// The sheet.
        /// </param>
        /// <param name="startRow">
        /// The start row.
        /// </param>
        /// <param name="startColumn">
        /// The start column.
        /// </param>
        /// <param name="testDefinition">
        /// The test definition.
        /// </param>
        /// <param name="swapColumnAndRowsFlag">
        /// The swap column and rows flag.
        /// </param>
        /// <param name="addRowNumberColumn">
        /// The add row number column.
        /// </param>
        public PDCListObject(string name, Worksheet sheet, int startRow, int startColumn, Testdefinition testDefinition,
                             bool swapColumnAndRowsFlag,
                             bool addRowNumberColumn)
            : this(
                name, sheet, startRow, startColumn, testDefinition, swapColumnAndRowsFlag, INITIAL_ROW_COUNT,
                addRowNumberColumn)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="name">
        /// The name.
        /// </param>
        /// <param name="sheet">
        /// The sheet.
        /// </param>
        /// <param name="startRow">
        /// The start row.
        /// </param>
        /// <param name="startColumn">
        /// The start column.
        /// </param>
        /// <param name="testDefinition">
        /// The test definition.
        /// </param>
        /// <param name="initialRowCount">
        /// The initial row count.
        /// </param>
        /// <param name="swapColumnAndRowsFlag">
        /// The swap column and rows flag.
        /// </param>
        /// <param name="addRowNumberColumn">
        /// The add row number column.
        /// </param>
        public PDCListObject(string name, Worksheet sheet, int startRow, int startColumn, Testdefinition testDefinition,
                             int initialRowCount,
                             bool swapColumnAndRowsFlag, bool addRowNumberColumn)
            : this(
                name, sheet, startRow, startColumn, testDefinition, swapColumnAndRowsFlag, initialRowCount,
                addRowNumberColumn)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class. 
        /// Creates a data entry list on the given worksheet for the specified test definition
        /// </summary>
        /// <param name="name">
        /// </param>
        /// <param name="sheet">
        /// </param>
        /// <param name="startRow">
        /// </param>
        /// <param name="startColumn">
        /// </param>
        /// <param name="testDefinition">
        /// </param>
        /// <param name="swapColumnAndRowsFlag">
        /// </param>
        /// <param name="initialRowCount">
        /// </param>
        /// <param name="addRowNumberColumn">
        /// </param>
        public PDCListObject(string name, Worksheet sheet, int startRow, int startColumn, Testdefinition testDefinition,
                             bool swapColumnAndRowsFlag,
                             int initialRowCount, bool addRowNumberColumn)
            : this(name)
        {
            myTestDefinition = testDefinition;
            mySwapColumnsAndRows = swapColumnAndRowsFlag;
            myContainer = sheet;
            Range headerRange = (Range)sheet.Cells[startRow, startColumn];
            DefineName(ListRangeName, headerRange, true);
            sheet.Names.Add(myHeaderRangeName, headerRange, true, missing, missing, missing, missing, missing, missing,
                            missing, missing);
            if (mySwapColumnsAndRows)
            {
                UpdateDataRange(startRow, startColumn + 1, 1, initialRowCount - 1);
            }
            else
            {
                UpdateDataRange(startRow + 1, startColumn, initialRowCount - 1, 1);
            }

            if (addRowNumberColumn)
            {
                if (!mySwapColumnsAndRows)
                {
                    ListColumn tmpRowNumberColumn = CreateRowNumberColumn();
                    AddColumn(tmpRowNumberColumn);
                }
            }
        }

        #endregion

        #region event handling

        #region CellsChanged

        /// <summary>
        /// Called by the PDCExcelAddin SheetChanged-Handler is the sheet changed 
        /// on the table area.
        /// </summary>
        /// <param name="range">
        /// The range.
        /// </param>
        /// <param name="sheetEvent">
        /// The sheet Event.
        /// </param>
        public virtual void CellsChanged(Range range, SheetEvent sheetEvent)
        {
            currentColumnMapping = null;
            try
            {
                if (CheckListAreaOnChange(range, sheetEvent))
                {
                    return;
                }

                RestoreValidations(range);
            }
            finally
            {
                UpdateRectangle();
            }
        }

        #endregion

        #region CheckListAreaOnChange

        /// <summary>
        /// Handles: 
        /// Deleted Columns
        /// Deleted Rows
        /// Added Rows
        /// </summary>
        /// <param name="cellChangedRange">
        /// The cell Changed Range.
        /// </param>
        /// <param name="sheetEvent">
        /// The sheet Event.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool CheckListAreaOnChange(Range cellChangedRange, SheetEvent sheetEvent)
        {
            bool stopEventHandling = false;
            if ((sheetEvent.entireColumn && !mySwapColumnsAndRows) || (sheetEvent.entireRow && mySwapColumnsAndRows))
            {
                // Most likely column movenment, deletion, addition
                stopEventHandling = CheckForRemovedColumns(stopEventHandling);
                return true;
            }

            if ((sheetEvent.entireColumn && mySwapColumnsAndRows) || (sheetEvent.entireRow && !mySwapColumnsAndRows))
            {
                // Most likely row movement, deletion, insertion
                stopEventHandling = HandleEntireRowEvent(cellChangedRange, sheetEvent);
            }

            if (stopEventHandling)
            {
                // Copy&Paste of Columns, ... : not safe for further event handling
                // return true;
            }

            Range dataRange = DataRange;
            if (mySwapColumnsAndRows)
            {
                int lastColumn = dataRange.Column + dataRange.Columns.Count - 1;
                int changeLastColumn = cellChangedRange.Column - 1 + cellChangedRange.Columns.Count;
                if (changeLastColumn >= lastColumn)
                {
                    ensureCapacity(changeLastColumn + 2 - dataRange.Column);
                }
            }
            else
            {
                int lastRow = dataRange.Row + dataRange.Rows.Count - 1;
                int changeLastRow = cellChangedRange.Row - 1 + cellChangedRange.Rows.Count;
                if (changeLastRow >= lastRow)
                {
                    ensureCapacity(changeLastRow + 2 - dataRange.Row);
                }
            }

            return false;
        }

        #endregion

        #region ColumnDeleted

        /// <summary>
        /// Called if the specified column was deleted. 
        /// Displays a warn message if the column is important
        /// </summary>
        /// <param name="column">
        /// A deleted column
        /// </param>
        /// <returns>
        /// Returns true if the eventhandling should stop here
        /// </returns>
        private bool ColumnDeleted(ListColumn column)
        {
            if (column.Removed)
            {
                return false;
            }

            column.Removed = true;
            if (column.TestVariable != null)
            {
                object defaultValaue = column.TestVariable.DefaultValue;
                if (defaultValaue == null && (column.TestVariable.IsMandatory || column.TestVariable.IsDifferentiating))
                {
                    MessageBox.Show(
                        string.Format(Resources.LIST_MANDATORY_DIFF_COLUMN_DELETED, column.TestVariable.VariableName),
                        Resources.MSG_INFO_TITLE);
                    return true;
                }

                if (column.TestVariable.IsExperimentLevel)
                {
                    MessageBox.Show(string.Format(Resources.LIST_NEEDED_PDC_COLUMN_DELETED, column.Name),
                                    Resources.MSG_INFO_TITLE);
                    return true;
                }
            }

            if (column.Name == PDCExcelConstants.UPLOAD_ID || column.Name == PDCExcelConstants.EXPERIMENT_NO ||
                column.Name == PDCExcelConstants.MEASUREMENTS)
            {
                MessageBox.Show(string.Format(Resources.LIST_NEEDED_PDC_COLUMN_DELETED, column.Name),
                                Resources.MSG_INFO_TITLE);
                return true;
            }

            if (myDateListColumns.ContainsKey(column.Name))
            {
                myDateListColumns.Remove(column.Name);
            }

            if (column.ParamHandler != null)
            {
                return column.ParamHandler.ColumnDeleted(this, column);
            }

            return false;
        }

        #endregion

        #region DeleteSMTMeasurementHyperLinks

        /// <summary>
        /// The delete smt measurement hyper links.
        /// </summary>
        public void DeleteSMTMeasurementHyperLinks()
        {
            if (myMeasurementColumn == null)
            {
                return;
            }

            try
            {
                PDCLogger.TheLogger.LogStarttime("Delete_Hyperlinks", "Deleting hyperlinks");
                Range tmpDataRange = DataRange;
                Range tmpRange = GetBodyRangeForColumn(myMeasurementColumn, 0, tmpDataRange.Rows.Count);
                tmpRange.Hyperlinks.Delete();
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Could not delete hyperlinks", e);
            }

            PDCLogger.TheLogger.LogStoptime("Delete_Hyperlinks", "Deleting hyperlinks");
        }

        #endregion

        #region HandleEntireRowEvent

        /// <summary>
        /// The handle entire row event.
        /// </summary>
        /// <param name="range">
        /// The range.
        /// </param>
        /// <param name="sheetEvent">
        /// The sheet event.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool HandleEntireRowEvent(Range range, SheetEvent sheetEvent)
        {
            if (myRowNumberColumn == null || myRowNumberColumn.Removed)
            {
                return false;
            }

            try
            {
                if (myDataRangeName.StartsWith("Data_Test_") && !DateRangeExists())
                {
                    UpdateDataRange(HeaderRange.Row + 1, 1, INITIAL_ROW_COUNT - 1, HeaderRange.Columns.Count);
                }


                Range dataRange = DataRange;

                Range rowColumn = Container.get_Range(myRowNumberColumn.Name, missing);
                object[,] rowNumbers = null;
                if (mySwapColumnsAndRows)
                {
                    rowNumbers = GetColumnValues(rowColumn.Row - dataRange.Row, true);
                }
                else
                {
                    rowNumbers = GetColumnValues(rowColumn.Column - dataRange.Column, true);
                }

                Dictionary<ListColumn, int> columnMapping = CurrentListColumnPlacements();
                if (!SwapColumnsAndRows)
                {
                    int j = dataRange.Row;
                    int startRow = j;
                    for (int i = rowNumbers.GetLowerBound(0); i <= rowNumbers.GetUpperBound(0); i++)
                    {
                        if (rowNumbers[i, rowNumbers.GetLowerBound(1)] == null)
                        {
                            foreach (ListColumn column in columnMapping.Keys)
                            {
                                InitializeNewColumnCells(column, j - startRow, j - startRow);
                            }
                        }

                        rowNumbers[i, rowNumbers.GetLowerBound(1)] = j;
                        j++;
                    }

                    SetColumnValues(rowColumn.Column - dataRange.Column, rowNumbers, true);
                }
                else
                {
                    int j = dataRange.Column;
                    int startColumn = j;
                    for (int i = rowNumbers.GetLowerBound(1); i <= rowNumbers.GetUpperBound(1); i++)
                    {
                        if (rowNumbers[rowNumbers.GetLowerBound(0), i] == null)
                        {
                            foreach (ListColumn column in columnMapping.Keys)
                            {
                                InitializeNewColumnCells(column, j - startColumn, j - startColumn);
                            }
                        }

                        rowNumbers[rowNumbers.GetLowerBound(1), i] = j;
                        j++;
                    }

                    SetColumnValues(rowColumn.Row - dataRange.Row, rowNumbers, true);
                }

                if (HasMeasurementParamHandler && myMeasurementColumn.HasMultiMeasurementTableHandler)
                {
                    myMeasurementColumn.MultiMeasurementTableHandler.RemoveUnreferencedTables(this,
                                                                                              GetColumnIndex(
                                                                                                  PDCExcelConstants
                                                                                                      .MEASUREMENTS)
                                                                                                  .Value);
                }
            }

#pragma warning disable 0168
            catch (Exception e)
            {
            }

#pragma warning restore 0168
            return true;
        }

        #endregion

        #endregion

        #region ISerializable Members

        /// <summary>
        /// The se r_ alread y_ uploaded.
        /// </summary>
        private const string SER_ALREADY_UPLOADED = "alreadyUploaded";

        /// <summary>
        /// The se r_ swap.
        /// </summary>
        private const string SER_SWAP = "swapColumns";

        /// <summary>
        /// The se r_ name.
        /// </summary>
        private const string SER_NAME = "name";

        /// <summary>
        /// The se r_ tableadapter.
        /// </summary>
        private const string SER_TABLEADAPTER = "tableAdapter";

        /// <summary>
        /// The se r_ rownumber.
        /// </summary>
        private const string SER_ROWNUMBER = "rowNumber";

        /// <summary>
        /// The se r_ measurementcolumn.
        /// </summary>
        private const string SER_MEASUREMENTCOLUMN = "measurementColumn";

        /// <summary>
        /// The se r_ rectangle.
        /// </summary>
        private const string SER_RECTANGLE = "rectangle";

        /// <summary>
        /// The se r_ identifier.
        /// </summary>
        private const string SER_IDENTIFIER = "identifier";

        /// <summary>
        /// The se r_ headerrangename.
        /// </summary>
        private const string SER_HEADERRANGENAME = "headerRangeName";

        /// <summary>
        /// The se r_ datarangename.
        /// </summary>
        private const string SER_DATARANGENAME = "dataRangeName";

        /// <summary>
        /// The se r_ columns.
        /// </summary>
        private const string SER_COLUMNS = "columns";

        /// <summary>
        /// The se r_ testdefinition.
        /// </summary>
        private const string SER_TESTDEFINITION = "testdefinition";

        /// <summary>
        /// The se r_ delete d_ columns.
        /// </summary>
        private const string SER_DELETED_COLUMNS = "deletedColumns";

        /// <summary>
        /// Initializes a new instance of the <see cref="PDCListObject"/> class.
        /// </summary>
        /// <param name="info">
        /// The info.
        /// </param>
        /// <param name="context">
        /// The context.
        /// </param>
        internal PDCListObject(SerializationInfo info, StreamingContext context)
        {
            myName = info.GetString(SER_NAME);
            mySwapColumnsAndRows = info.GetBoolean(SER_SWAP);
            myRowNumberColumn = (ListColumn)info.GetValue(SER_ROWNUMBER, typeof(ListColumn));
            myMeasurementColumn = (ListColumn)info.GetValue(SER_MEASUREMENTCOLUMN, typeof(ListColumn));
            myRectangle = (Rectangle)info.GetValue(SER_RECTANGLE, typeof(Rectangle));
            myIdentifier = info.GetString(SER_IDENTIFIER);
            myHeaderRangeName = info.GetString(SER_HEADERRANGENAME);
            myDataRangeName = info.GetString(SER_DATARANGENAME);
            myTestDefinition = (Testdefinition)info.GetValue(SER_TESTDEFINITION, typeof(Testdefinition));
            myDeletedColumns = (List<ListColumn>)info.GetValue(SER_DELETED_COLUMNS, typeof(List<ListColumn>));
            myColumns = (List<ListColumn>)info.GetValue(SER_COLUMNS, typeof(List<ListColumn>));
            myIsAlreadyUploaded = info.GetBoolean(SER_ALREADY_UPLOADED);
            myValidationHandler = new ValidationHandler(this);
            myTestDataAdapter = new TestDataTableAdapter(this);
            FillDateListColumns();
        }

        /// <summary>
        /// The get object data.
        /// </summary>
        /// <param name="info">
        /// The info.
        /// </param>
        /// <param name="context">
        /// The context.
        /// </param>
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue(SER_SWAP, mySwapColumnsAndRows);
            info.AddValue(SER_NAME, myName);
            info.AddValue(SER_ROWNUMBER, myRowNumberColumn);
            info.AddValue(SER_MEASUREMENTCOLUMN, myMeasurementColumn);
            info.AddValue(SER_RECTANGLE, myRectangle);
            info.AddValue(SER_IDENTIFIER, myIdentifier);
            info.AddValue(SER_HEADERRANGENAME, myHeaderRangeName);
            info.AddValue(SER_DATARANGENAME, myDataRangeName);
            info.AddValue(SER_TESTDEFINITION, myTestDefinition);
            info.AddValue(SER_DELETED_COLUMNS, myDeletedColumns);
            info.AddValue(SER_COLUMNS, myColumns);
            info.AddValue(SER_ALREADY_UPLOADED, myIsAlreadyUploaded);
        }

        #endregion

        #region methods

        #region AddColumn

        /// <summary>
        /// Convenience method to add a single column
        /// </summary>
        /// <param name="column">
        /// </param>
        public void AddColumn(ListColumn column)
        {
            List<ListColumn> columns = new List<ListColumn>();
            columns.Add(column);
            AddColumns(columns);
        }

        #endregion

        #region AddColumns

        /// <summary>
        /// Adds the specified columns to the sheet and formats them (eg color). The columns are placed at the end of the table
        /// </summary>
        /// <param name="columnList">
        /// </param>
        public void AddColumns(List<ListColumn> columnList)
        {
            bool enable = Globals.PDCExcelAddIn.Application.EnableEvents;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            object[] variableTypeHeaders = new object[columnList.Count];
            object[] pdcOnlyTypeHeaders = new object[columnList.Count];
            try
            {
                if (myTestDefinition != null)
                {
                    PicklistHandler.ThePicklistHandler(myContainer).UpdatePickListCache(myTestDefinition);
                }

                Range listRange = myContainer.get_Range(ListRangeName, missing);
                Range dataRange = DataRange;
                int tmpRow = listRange.Row;
                int i = listRange.Column + ColumnCount;
                int tmpStart = i;
                int rowCount = dataRange.Rows.Count;
                if (mySwapColumnsAndRows)
                {
                    i = listRange.Row + ColumnCount;
                    rowCount = dataRange.Columns.Count;
                    tmpRow = listRange.Column;
                }

                List<ListColumn> toInitialize = new List<ListColumn>();
                foreach (ListColumn column in columnList)
                {
                    Range range = null;
                    if (mySwapColumnsAndRows)
                    {
                        range = (Range)myContainer.Cells[i, tmpRow];
                    }
                    else
                    {
                        range = (Range)myContainer.Cells[tmpRow, i];
                        if (column.TestVariable != null)
                        {
                            if (column.TestVariable.IsPdcVariable)
                            {
                                pdcOnlyTypeHeaders[i - tmpStart] = "PDC only";
                            }
                            else
                            {
                                pdcOnlyTypeHeaders[i - tmpStart] = null;
                            }

                            if (!column.TestVariable.IsBinaryParameter())
                            {
                                variableTypeHeaders[i - tmpStart] = column.TestVariable.IsNumeric()
                                                                        ? "numeric"
                                                                        : "alpha-numeric";
                            }
                        }
                        else
                        {
                            pdcOnlyTypeHeaders[i - tmpStart] = null;
                            variableTypeHeaders[i - tmpStart] = null;
                        }
                    }

                    Name name = myContainer.Names.Add(column.Name, range, missing, missing, missing, missing, missing,
                                                      missing, missing, missing, missing);
                    range.Formula = column.Label;
                    range.Font.Bold = true;
                    range.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium,
                                       XlColorIndex.xlColorIndexAutomatic, missing);
                    range.Interior.ColorIndex = 0;
                    if (column.OleColor.HasValue)
                    {
                        range.Interior.Color = column.OleColor;
                    }

                    ExcelUtils.TheUtils.AddComment(range, column.Comment, true);
                    object widthO = range.ColumnWidth;
                    long width = 0;
                    if (widthO != null)
                    {
                        width = PDCConverter.Converter.ToLong(width) ?? 0;
                    }

                    if (column.Label == null)
                    {
                        range.ColumnWidth = width;
                    }
                    else
                    {
                        range.ColumnWidth = Math.Max(width, column.Label.Length + 2);
                    }

                    if (column.Hidden)
                    {
                        if (mySwapColumnsAndRows)
                        {
                            range.EntireRow.Hidden = true;
                        }
                        else
                        {
                            range.EntireColumn.Hidden = true;
                        }
                    }

                    if (column.ParamHandler is RowNumberHandler)
                    {
                        myRowNumberColumn = column;
                    }

                    if (column.ParamHandler is MultipleMeasurementTableHandler ||
                        column.ParamHandler is SingleMeasurementTableHandler)
                    {
                        myMeasurementColumn = column;
                    }

                    toInitialize.Add(column);
                    i++;
                    myColumns.Add(column);
                }

                // Invalidate stored column mapping before initializing the new columns
                int newColumnEnd = i - 1;
                if (mySwapColumnsAndRows)
                {
                    Range tmpHeaderRange = ExcelUtils.TheUtils.GetRange(myContainer,
                                                                     myContainer.Cells[listRange.Row, listRange.Column],
                                                                     myContainer.Cells[newColumnEnd, listRange.Column]);
                    Name tmpHeaderName = myContainer.Names.Add(myHeaderRangeName, tmpHeaderRange, true, missing, missing,
                                                               missing, missing, missing, missing, missing, missing);

                    UpdateDataRange(dataRange.Row, tmpHeaderRange.Column + 1, i - dataRange.Row, rowCount);
                }
                else
                {
                    int tmpRowNr = listRange.Row;
                    int tmpColumnNr = listRange.Column;
                    Range tmpHeaderRange = ExcelUtils.TheUtils.GetRange(myContainer,
                                                                     myContainer.Cells[tmpRowNr, tmpColumnNr],
                                                                     myContainer.Cells[tmpRowNr, newColumnEnd]);
                    Name tmpHeaderName = myContainer.Names.Add(myHeaderRangeName, tmpHeaderRange, true, missing, missing,
                                                               missing, missing, missing, missing, missing, missing);

                    WriteSingleLineRange(variableTypeHeaders, tmpStart, tmpRowNr - 2,
                                         Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_DATA_TYPE]);
                    WriteSingleLineRange(pdcOnlyTypeHeaders, tmpStart, tmpRowNr - 1,
                                         Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_PDC_ONLY]);
                    UpdateDataRange(dataRange.Row, tmpHeaderRange.Column, rowCount, i - tmpHeaderRange.Column);
                }

                // Initialize cells 
                ClearCurrentMapping();
                foreach (ListColumn column in toInitialize)
                {
                    InitializeNewColumnCells(column, 0, rowCount);
                }

                if (myMeasurementColumn != null && myMeasurementColumn.HasSingleMeasurementTableHandler)
                {
                    myMeasurementColumn.SingleMeasurementTableHandler.InitializeSheet(this);
                }
            }
            finally
            {
                ClearCurrentMapping();
                Globals.PDCExcelAddIn.Application.EnableEvents = enable;
            }
        }

        /// <summary>
        /// The write single line range.
        /// </summary>
        /// <param name="content">
        /// The content.
        /// </param>
        /// <param name="aStart">
        /// The a start.
        /// </param>
        /// <param name="aRowNr">
        /// The a row nr.
        /// </param>
        /// <param name="background">
        /// The background.
        /// </param>
        private void WriteSingleLineRange(object[] content, int aStart, int aRowNr, ClientConfiguration.Color background)
        {
            Range tmpRange;
            if (background != null)
            {
                int? tmpStart = null;
                for (int i = 0; i <= content.Length; i++)
                {
                    if (i == content.Length || content[i] == null)
                    {
                        if (tmpStart != null)
                        {
                            tmpRange = ExcelUtils.TheUtils.GetRange(myContainer, aRowNr, aStart + tmpStart.Value, aRowNr,
                                                                 aStart + i - 1);
                            tmpRange.Interior.Color = background.OleColor;
                            tmpStart = null;
                        }
                    }
                    else if (tmpStart == null)
                    {
                        tmpStart = i;
                    }
                }
            }

            tmpRange = ExcelUtils.TheUtils.GetRange(myContainer, aRowNr, aStart, aRowNr, aStart + content.Length - 1);
            object[,] tmpValues = (object[,])tmpRange.get_Value(missing);

            if (tmpValues != null)
            {
                int j = 0;
                int y = tmpValues.GetLowerBound(0);
                for (int i = tmpValues.GetLowerBound(1); i < tmpValues.GetUpperBound(1); i++)
                {
                    if (j >= content.Length)
                    {
                        break;
                    }

                    if (content[j] != null)
                    {
                        j++;
                        continue;
                    }

                    content[j] = tmpValues[y, i];
                    j++;
                }
            }

            tmpRange.Value = content;
        }

        #endregion

        #region CalculateIntersection

        /// <summary>
        /// Calculates the overlapping cell range between the rectangle and the pdc list.
        /// </summary>
        /// <param name="aRectangle">
        /// </param>
        /// <returns>
        /// The Range of overlapping cells
        /// </returns>
        public Range CalculateIntersection(Rectangle aRectangle)
        {
            Rectangle tmpListRectangle = Rectangle;
            Rectangle tmpListArea = new Rectangle(tmpListRectangle.Location, tmpListRectangle.Size);


            if (SwapColumnsAndRows)
            {
                tmpListArea.Width++;
            }
            else
            {
                tmpListArea.Height++;
            }

            if (tmpListArea.IntersectsWith(aRectangle))
            {
                tmpListArea.Intersect(aRectangle);
                if (tmpListArea.IsEmpty)
                {
                    return null;
                }

                Worksheet tmpSheet = Container;
                return ExcelUtils.TheUtils.GetRange(tmpSheet, tmpSheet.Cells[tmpListArea.Top, tmpListArea.Left],
                                                 tmpSheet.Cells[tmpListArea.Bottom, tmpListArea.Right]);
            }

            return null;
        }

        #endregion

        #region CheckForRemovedColumns

        /// <summary>
        /// Checks if columns were removed
        /// </summary>
        /// <param name="aStopEventHandling">
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool CheckForRemovedColumns(bool aStopEventHandling)
        {
            foreach (ListColumn tmpColumn in myColumns)
            {
                try
                {
                    Range tmpRange = myContainer.get_Range(tmpColumn.Name, missing);
                    tmpColumn.Removed = false;
                }

#pragma warning disable 0168
                catch (Exception e)
                {
                    aStopEventHandling |= ColumnDeleted(tmpColumn);
                }

#pragma warning restore 0168
            }

            return aStopEventHandling;
        }

        #endregion

        #region ClearContents

        /// <summary>
        /// Clears the contents of the pdc list
        /// </summary>
        public void ClearContents()
        {
            PDCLogger.TheLogger.LogStarttime(Name + ".ClearContents", "Clearing content");
            bool eventsEnabled = Globals.PDCExcelAddIn.EventsEnabled;
            Globals.PDCExcelAddIn.EventsEnabled = false;
            bool appEventsEnabled = Globals.PDCExcelAddIn.Application.EnableEvents;
            if (appEventsEnabled)
            {
                Globals.PDCExcelAddIn.Application.EnableEvents = false;
            }

            ShowAllData();

            try
            {
                Range dataRange = DataRange;
                object[,] values = new object[dataRange.Rows.Count, dataRange.Columns.Count];

                Dictionary<ListColumn, int> clumnMapping = CurrentListColumnPlacements();
                foreach (KeyValuePair<ListColumn, int> pair in clumnMapping)
                {
                    if (pair.Key.ParamHandler != null)
                    {
                        pair.Key.ParamHandler.ClearContents(this, pair, values);
                    }

                    if (pair.Key.ParamHandler2 != null)
                    {
                        pair.Key.ParamHandler2.ClearContents(this, pair, values);
                    }
                }

                dataRange.set_Value(missing, values);

                ClearHyperLinks();
                ValidationHandler.ClearAllValidationMessages();

                myIsAlreadyUploaded = false;
                ResetMeasurementHyperLinks();
                ExcelUtils.TheUtils.ResetRowHeights(dataRange);
                int rowy = dataRange.Rows.Count;
                if (rowy < INITIAL_ROW_COUNT - 10)
                {
                    if (HasMeasurementParamHandler && MeasurementColumn.HasMultiMeasurementTableHandler)
                    {
                        MeasurementColumn.MultiMeasurementTableHandler.RemoveUnreferencedTables(this,
                                                                                                GetColumnIndex(
                                                                                                    PDCExcelConstants
                                                                                                        .MEASUREMENTS)
                                                                                                    .Value);
                    }

                    UpdateDataRange(HeaderRange.Row + 1, 1, INITIAL_ROW_COUNT - 1, HeaderRange.Columns.Count);
                }

                ((Range)myContainer.Cells[HeaderRange.Row + 1, 2]).Select();
            }
            finally
            {
                Globals.PDCExcelAddIn.EventsEnabled = eventsEnabled;
                if (appEventsEnabled)
                {
                    Globals.PDCExcelAddIn.Application.EnableEvents = appEventsEnabled;
                }

                PDCLogger.TheLogger.LogStoptime(Name + ".ClearContents", "Clearing content");
            }
        }

        #endregion

        #region ClearCurrentMapping

        /// <summary>
        /// Clears the current column mapping if it could be changed in the meantime
        /// </summary>
        public void ClearCurrentMapping()
        {
            currentColumnMapping = null;
        }

        #endregion

        #region ClearHyperLinks

        /// <summary>
        /// Removes any hyperlink except the links to the measurement tables
        /// </summary>
        protected virtual void ClearHyperLinks()
        {
            Range dataRange = DataRange;

            if (myMeasurementColumn != null)
            {
                int measPos = -1;
                Dictionary<ListColumn, int> mapping = CurrentListColumnPlacements();
                foreach (KeyValuePair<ListColumn, int> pair in mapping)
                {
                    // todo Implement SingleMeasurementTableHandler
                    if (pair.Key.HasMultiMeasurementTableHandler || pair.Key.HasSingleMeasurementTableHandler)
                    {
                        measPos = pair.Value;
                        break;
                    }
                }

                // Do not remove measurement table hyperlink. Usually this is the last column of the table,
                // but in fact it can be at any position
                if (measPos >= 0)
                {
                    int column = mapping[myMeasurementColumn];
                    int firstRow = dataRange.Row;
                    int lastRow = firstRow + dataRange.Rows.Count - 1;
                    int firstColumn = dataRange.Column;
                    int lastColumn = dataRange.Columns.Count - 1;
                    if (column == 0)
                    {
                        dataRange = ExcelUtils.TheUtils.GetRange(Container, firstRow, firstColumn + 1, lastRow, firstColumn + lastColumn);
                    }
                    else if (column == lastColumn)
                    {
                        dataRange = ExcelUtils.TheUtils.GetRange(Container,
                                                              firstRow, firstColumn,
                                                              lastRow, firstColumn + lastColumn - 1);
                    }
                    else
                    {
                        dataRange = ExcelUtils.TheUtils.GetRange(Container,
                                                              firstRow, firstColumn,
                                                              lastRow, firstColumn + column - 1);
                        dataRange.Hyperlinks.Delete();

                        dataRange = ExcelUtils.TheUtils.GetRange(Container,
                                                              firstRow, column + 1,
                                                              lastRow, firstColumn + lastColumn - 1);
                    }
                }
            }

            dataRange.Hyperlinks.Delete();
        }

        #endregion

        #region ColumnHeader

        /// <summary>
        /// The column header.
        /// </summary>
        /// <param name="aColumnName">
        /// The a column name.
        /// </param>
        /// <returns>
        /// The <see cref="Range"/>.
        /// </returns>
        internal Range ColumnHeader(string aColumnName)
        {
            Range tmpColumn = myContainer.get_Range(aColumnName, missing);
            return (Range)myContainer.Cells[tmpColumn.Row, tmpColumn.Column];
        }

        #endregion

        #region ColumnIndex

        /// <summary>
        /// Returns the column index for the first ListColumn with the specified name.
        /// Returns -1 if no such column is known.
        /// </summary>
        /// <param name="aColumnIdentifier">
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int ColumnIndex(string aColumnIdentifier)
        {
            for (int i = 0; i < ColumnCount; i++)
            {
                if (myColumns[i].Name == aColumnIdentifier)
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Returns the column index for the specified ListColumn or -1 of the column could
        /// not be found.
        /// </summary>
        /// <param name="aColumn">
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int ColumnIndex(ListColumn aColumn)
        {
            for (int i = 0; i < ColumnCount; i++)
            {
                if (myColumns[i] == aColumn)
                {
                    return i;
                }
            }

            return -1;
        }

        #endregion

        #region ColumnRange

        /// <summary>
        /// The column range.
        /// </summary>
        /// <param name="i">
        /// The i.
        /// </param>
        /// <returns>
        /// The <see cref="Range"/>.
        /// </returns>
        private Range ColumnRange(int i)
        {
            return ColumnRange(i, false);
        }

        /// <summary>
        /// Returns the data range for the specfied column index
        /// </summary>
        /// <param name="aColumnIndex">
        /// The (zero-based) column no
        /// </param>
        /// <param name="includeInsertLine">
        /// Wether to include the insert line
        /// </param>
        /// <returns>
        /// The <see cref="Range"/>.
        /// </returns>
        public Range ColumnRange(int aColumnIndex, bool includeInsertLine)
        {
            Range tmpBodyRange = DataRange;
            int tmpColumn = tmpBodyRange.Column;

            int tmpRow = tmpBodyRange.Row;
            int tmpHeight = tmpBodyRange.Rows.Count;

            Range result;
            if (mySwapColumnsAndRows)
            {
                tmpHeight = tmpBodyRange.Columns.Count;
                if (includeInsertLine)
                {
                    tmpHeight++;
                }

                result = ExcelUtils.TheUtils.GetRange(myContainer,
                                                 myContainer.Cells[tmpRow + aColumnIndex, tmpColumn],
                                                 myContainer.Cells[tmpRow + aColumnIndex, tmpColumn + tmpHeight - 1]);
            }
            else
            {
                if (includeInsertLine)
                {
                    tmpHeight++;
                }

                result = ExcelUtils.TheUtils.GetRange(myContainer,
                                                 myContainer.Cells[tmpRow, tmpColumn + aColumnIndex],
                                                 myContainer.Cells[tmpRow + tmpHeight - 1, tmpColumn + aColumnIndex]);                
            }

            try
            {
                return result.SpecialCells(XlCellType.xlCellTypeVisible);
            }
            catch (COMException)
            {
                //Range have only hidden cells
                //if((Boolean)result.EntireColumn.Hidden || (Boolean)result.EntireRow.Hidden) return result; -> didn't work for ander cases
                return result;
            }

        }

        #endregion

        #region CreateColumn

        /// <summary>
        /// The create column.
        /// </summary>
        /// <param name="aVariable">
        /// The a variable.
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        public static ListColumn CreateColumn(TestVariable aVariable)
        {
            return CreateColumn(aVariable, null);
        }

        /// <summary>
        /// The create column.
        /// </summary>
        /// <param name="aVariable">
        /// The a variable.
        /// </param>
        /// <param name="aNameSuffix">
        /// The a name suffix.
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        public static ListColumn CreateColumn(TestVariable aVariable, int? aNameSuffix)
        {
            string tmpLabel = aVariable.Label;
            Color tmpColor = GetVariableColumnColor(aVariable);

            string tmpComment = aVariable.Comments ?? string.Empty;
            string tmpNewLine = aVariable.Comments == null ? string.Empty : "\n";
            if (aVariable.IsExperimentLevelReference)
            {
                tmpComment += tmpNewLine + "Experiment level Parameter";
                tmpNewLine = "\n";
            }

            if (aVariable.IsMandatory)
            {
                tmpComment += tmpNewLine + "Mandatory";
                tmpNewLine = "\n";
            }

            if (aVariable.IsDifferentiating)
            {
                tmpComment += tmpNewLine + "Differentiating";
                tmpNewLine = "\n";
            }

            if (aVariable.DefaultValue != null)
            {
                tmpComment += tmpNewLine + "Default:" + aVariable.DefaultValue;
            }

            if (tmpComment.Trim() == string.Empty)
            {
                tmpComment = null;
            }

            string tmpVarName = CreateColumnName(aVariable.VariableId, aNameSuffix);
            ListColumn tmpVarColumn = new ListColumn(tmpVarName, tmpLabel, aVariable.VariableId, tmpComment, tmpColor);
            tmpVarColumn.TestVariable = aVariable;
            aVariable.Tag = tmpVarColumn;
            return tmpVarColumn;
        }

        /// <summary>
        /// Returns the desired background color of the header cell.
        /// </summary>
        /// <param name="aVariable">
        /// </param>
        /// <returns>
        /// The <see cref="Color"/>.
        /// </returns>
        private static Color GetVariableColumnColor(TestVariable aVariable)
        {
            ClientConfiguration.Color tmpColor = null;

            if (aVariable.IsExperimentLevelReference)
            {
                tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_EXPERIMENT_LEVEL];
            }
            else if (aVariable.IsDifferentiating)
            {
                tmpColor =
                    Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_DIFFERENTIATING_PARAMETER];
            }
            else if (aVariable.IsMandatory)
            {
                tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_MANDATORY_PARAMETER];
            }
            else if (aVariable.IsCoreResult)
            {
                tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_CORE_RESULT];
            }

            if (tmpColor != null)
            {
                return tmpColor.SystemColor;
            }

            switch (aVariable.VariableClass)
            {
                case "D":
                    return
                        Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_DERIVED_RESULT].SystemColor;
                case "P":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_PARAMETER].SystemColor;
                case "R":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_RESULT].SystemColor;
                case "V":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_VARIABLE].SystemColor;
                case "B":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_BINARY].SystemColor;
                case "A":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_ANNOTATION].SystemColor;
                case "C":
                    return Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.HEADER_COMMENT].SystemColor;
            }

            return Color.White;

        }

        #endregion

        #region CreateColumnName

        /// <summary>
        /// The create column name.
        /// </summary>
        /// <param name="aVarName">
        /// The a var name.
        /// </param>
        /// <param name="aNameSuffix">
        /// The a name suffix.
        /// </param>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        protected static string CreateColumnName(int aVarName, int? aNameSuffix)
        {
            return "Var_" + aVarName + (aNameSuffix == null ? string.Empty : "_" + aNameSuffix.Value);
        }

        #endregion

        #region CreateMeasurementColumn

        /// <summary>
        /// The create measurement column.
        /// </summary>
        /// <param name="aSheet">
        /// The a sheet.
        /// </param>
        /// <param name="aTD">
        /// The a td.
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        public static ListColumn CreateMeasurementColumn(Worksheet aSheet, Testdefinition aTD)
        {
            ListColumn tmpColumn = new ListColumn(PDCExcelConstants.MEASUREMENTS, "Measurements",
                                                  PDCConstants.C_ID_MEASUREMENTNO,
                                                  Globals.PDCExcelAddIn.ClientConfiguration[
                                                      ClientConfiguration.HEADER_VARIABLE].SystemColor);

            // für den alten Measurement-Mechanismus verwnde man den MeasurementHandler
            PredefinedParameterHandler tmpHandler = null;
            if (Globals.PDCExcelAddIn.DoCreateBothMeasurementTables())
            {
                tmpHandler = new MultipleMeasurementTableHandler(aTD);
                tmpColumn.ParamHandler = tmpHandler;

                tmpHandler = new SingleMeasurementTableHandler(aTD);
                tmpColumn.ParamHandler2 = tmpHandler;
            }
            else
            {
                if (aTD.ShowSingleMeasurement)
                {
                    tmpHandler = new SingleMeasurementTableHandler(aTD);
                }
                else
                {
                    tmpHandler = new MultipleMeasurementTableHandler(aTD);
                }

                tmpColumn.ParamHandler = tmpHandler;
            }

            tmpColumn.ReadOnly = true;
            tmpColumn.IsHyperLink = true;

            return tmpColumn;
        }

        #endregion

        #region CreateRowNumberColumn

        /// <summary>
        /// The create row number column.
        /// </summary>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        private ListColumn CreateRowNumberColumn()
        {
            ListColumn tmpRowNumberColumn = new ListColumn(ROW_COL_PREFIX + myName, "Row number", -1, Color.Gray);
            tmpRowNumberColumn.Hidden = true;
            tmpRowNumberColumn.ReadOnly = true;
            tmpRowNumberColumn.ParamHandler = new RowNumberHandler();
            return tmpRowNumberColumn;
        }

        #endregion

        #region CurrentListColumnPlacements

        /// <summary>
        /// Returns the current mapping from ListColumn to column number (within the PDCListObject)
        /// </summary>
        public virtual Dictionary<ListColumn, int> CurrentListColumnPlacements()
        {
            if (currentColumnMapping != null)
            {
                return currentColumnMapping;
            }

            Dictionary<ListColumn, int> tmpMapping = new Dictionary<ListColumn, int>();
            Range tmpDataRange = DataRange;
            int tmpFirstColumn = mySwapColumnsAndRows ? tmpDataRange.Row : tmpDataRange.Column;
            int tmpLastColumn = tmpFirstColumn +
                                (mySwapColumnsAndRows ? tmpDataRange.Rows.Count : tmpDataRange.Columns.Count) - 1;

            foreach (ListColumn tmpColumn in myColumns)
            {
                try
                {
                    Range tmpColumnRange = myContainer.get_Range(tmpColumn.Name, missing);
                    int tmpColumnPlacement = mySwapColumnsAndRows ? tmpColumnRange.Row : tmpColumnRange.Column;
                    if (tmpColumnPlacement >= tmpFirstColumn && tmpColumnPlacement <= tmpLastColumn)
                    {
                        tmpMapping.Add(tmpColumn, tmpColumnPlacement - tmpFirstColumn);
                    }

                    // ignore ListColumn outside the DataRange
                }

#pragma warning disable 0168
                catch (Exception e)
                {
                    // Column probably not used.
                }

#pragma warning restore 0168
            }

            currentColumnMapping = tmpMapping;
            return tmpMapping;
        }

        #endregion

        #region ListColumn

        /// <summary>
        /// Returns the meta data for the specified column
        /// </summary>
        /// <param name="ColumnId">
        /// The Column Id.
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        public ListColumn ListColumnByColumnId(int ColumnId)
        {
            foreach (ListColumn tmpColumn in myColumns)
            {
                if (tmpColumn.Id.Value == ColumnId) return tmpColumn;
            }

            return null;
        }

        #endregion

        #region DateRangeExists

        /// <summary>
        /// The date range exists.
        /// </summary>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public virtual bool DateRangeExists()
        {
            try
            {
                myContainer.get_Range(myDataRangeName, missing);
                return true;
            }
            catch (Exception)
            {
            }

            return false;
        }

        #endregion

        #region DefineName

        /// <summary>
        /// </summary>
        /// <param name="aName">
        /// </param>
        /// <param name="aRange">
        /// </param>
        /// <param name="aVisibleFlag">
        /// </param>
        private void DefineName(string aName, Range aRange, bool aVisibleFlag)
        {
            Workbook tmpWB = (Workbook)myContainer.Parent;
            tmpWB.Names.Add(aName, aRange, aVisibleFlag, missing, missing, missing, missing, missing, missing, missing,
                            missing);
        }

        #endregion

        #region Delete

        /// <summary>
        /// Delete the PDCListObject and everything that belongs to it from Excel.
        /// </summary>
        /// <returns>The number of deleted lines</returns>
        public int Delete()
        {
            Container.AutoFilterMode = false;
            foreach (ListColumn tmpColumn in myColumns)
            {
                // Call predefined parameter handler
                if (tmpColumn.ParamHandler != null)
                {
                    tmpColumn.ParamHandler.Delete(tmpColumn, this, true);
                }

                if (tmpColumn.ParamHandler2 != null)
                {
                    tmpColumn.ParamHandler2.Delete(tmpColumn, this, true);
                }
            }

            // remove the saved datelist columns  (Upload Date & Test Date)
            myDateListColumns.Clear();

            Range tmpListRange = ListRangeByName;

            // DeleteNamedRanges();
            int tmpStart = tmpListRange.Row;
            if (UserConfiguration.TheConfiguration.GetBooleanProperty(UserConfiguration.PROP_DRAW_DATA_ENTRY_BORDER,
                                                                      false))
            {
                tmpListRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
                tmpListRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
                tmpListRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
                tmpListRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
            }

            int tmpEnd = tmpStart + tmpListRange.Rows.Count - 1;
            int tmpColumnEnd = ((Range)Container.Rows[1, missing]).Columns.Count;
            if (mySwapColumnsAndRows)
            {
                // try to minimize empty lines between tables
                Range tmpSpaceRange = (Range)Container.Rows[tmpEnd + 1, missing];
                int tmpSpaceEnd = tmpSpaceRange.get_End(XlDirection.xlToRight).Column;
                int tmpSpaceLineEnd = tmpSpaceRange.Columns.Count;
                if (tmpSpaceLineEnd == tmpSpaceEnd)
                {
                    tmpEnd = tmpEnd + 1;
                }
            }

            Range tmpDeleteRange = ExcelUtils.TheUtils.GetRange(Container, Container.Cells[tmpStart, 1],
                                                             Container.Cells[tmpEnd, tmpColumnEnd]);
            tmpDeleteRange.Delete(missing);

            SheetInfo tmpInfo = SheetInfo;
            tmpInfo.RemoveMeasurementTable(myIdentifier);
            return tmpEnd + 1 - tmpStart;
        }

        #endregion

        #region DeleteColumns

        /// <summary>
        /// </summary>
        /// <param name="aColumnList">
        /// </param>
        public void DeleteColumns(List<ListColumn> aColumnList)
        {
            // Todo checken ob das wirklich wegkann
            // ActivateAutoFilter(false);
            List<string> tmpNamedRanges = new List<string>();
            foreach (ListColumn tmpColumn in aColumnList)
            {
                tmpNamedRanges.Add(tmpColumn.Name);
                if (tmpColumn.ParamHandler != null)
                {
                    tmpColumn.ParamHandler.Delete(tmpColumn, this, false);
                }

                Range tmpRange = myContainer.get_Range(tmpColumn.Name, missing);
                tmpRange.EntireColumn.Delete(missing);
                myColumns.Remove(tmpColumn);
                if (tmpColumn == myMeasurementColumn)
                {
                    myMeasurementColumn = null;
                }

                if (tmpColumn == myRowNumberColumn)
                {
                    myRowNumberColumn = null;
                }
            }

            ExcelUtils.TheUtils.DeleteNames(myContainer, null, tmpNamedRanges.ToArray());
        }

        #endregion

        #region DrawBorder

        /// <summary>
        /// Draws a border around the area and removes any line fragment in the area itself.
        /// </summary>
        /// <param name="aRow">
        /// </param>
        /// <param name="aColumn">
        /// </param>
        /// <param name="aHeight">
        /// </param>
        /// <param name="aWidth">
        /// </param>
        private void DrawBorder(int aRow, int aColumn, int aHeight, int aWidth)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL,
                                                "DrawBorder(" + Name + ", " + aRow + "," + aColumn + "," + aHeight +
                                                " ," + aWidth + ")");
            if (aWidth < 1 || aHeight < 1)
            {
                return;
            }

            Range tmpRange = ExcelUtils.TheUtils.GetRange(myContainer, myContainer.Cells[aRow, aColumn],
                                                       myContainer.Cells[aRow + aHeight - 1, aColumn + aWidth - 1]);

            // tmpRange.Interior.Pattern = Excel.XlPattern.xlPatternGray16;
            tmpRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;

            // tmpRange.Interior.ColorIndex = 0;
            tmpRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThick, XlColorIndex.xlColorIndexAutomatic,
                                  missing);
        }

        #endregion

        #region DrawRectangle

        /// <summary>
        /// Draws a border around the data entry area if the configuration says so
        /// </summary>
        protected void DrawRectangle()
        {
            if (
                !UserConfiguration.TheConfiguration.GetBooleanProperty(UserConfiguration.PROP_DRAW_DATA_ENTRY_BORDER,
                                                                       false))
            {
                return;
            }

            if (mySwapColumnsAndRows)
            {
                DrawBorder(myRectangle.Y, myRectangle.X, myRectangle.Height, myRectangle.Width);
            }
            else
            {
                DrawBorder(myRectangle.Y, myRectangle.X + 1, myRectangle.Height, myRectangle.Width - 1);
            }
        }

        #endregion

        #region ensureCapacity

        /// <summary>
        /// Ensures that the list has at least aRowCount data rows
        /// </summary>
        /// <param name="aRowCount">
        /// </param>
        public virtual void ensureCapacity(int aRowCount)
        {
            PDCLogger.TheLogger.LogStarttime("EnsureCapacity." + Name, Name + ".EnsureCapacity " + aRowCount);
            try
            {
                Range tmpDataRange = DataRange;
                int tmpStartColumn = tmpDataRange.Column;
                int tmpColumnNr = tmpDataRange.Columns.Count;
                int tmpStartRow = tmpDataRange.Row;
                int tmpRowCount = tmpDataRange.Rows.Count;
                int tmpCurrentInsertRow = tmpStartRow + tmpRowCount;
                int tmpNewInsertRowEndRow = tmpStartRow + aRowCount;
                if (mySwapColumnsAndRows)
                {
                    tmpCurrentInsertRow = tmpStartColumn + tmpColumnNr;
                    tmpNewInsertRowEndRow = tmpStartColumn + aRowCount;
                }

                if (tmpCurrentInsertRow >= tmpNewInsertRowEndRow)
                {
                    return;
                }

                // TODO use list column mapping
                foreach (ListColumn tmpColumn in myColumns)
                {
                    if (tmpColumn.Removed)
                    {
                        continue;
                    }

                    if (mySwapColumnsAndRows)
                    {
                        InitializeNewColumnCells(tmpColumn, tmpColumnNr + 1, aRowCount);
                    }
                    else
                    {
                        InitializeNewColumnCells(tmpColumn, tmpRowCount + 1, aRowCount);
                    }
                }

                if (mySwapColumnsAndRows)
                {
                    UpdateDataRange(tmpStartRow, tmpStartColumn, tmpRowCount, aRowCount);
                }
                else
                {
                    UpdateDataRange(tmpStartRow, tmpStartColumn, aRowCount, tmpColumnNr);
                }
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("EnsureCapacity." + Name, Name + ".EnsureCapacity " + aRowCount);
            }
        }

        #endregion

        #region ExcelTableExists

        /// <summary>
        /// Checks if the named ranges for the table still exists.
        /// </summary>
        /// <returns>Returns true if all relevant named ranges still exist, false otherwise</returns>
        public bool ExcelTableExists()
        {
            try
            {
                Range tmpUnused = ListRangeByName;
                tmpUnused = DataRange;
                tmpUnused = HeaderRange;
                return true;
            }

#pragma warning disable 0168
            catch (Exception e)
            {
                return false;
            }

#pragma warning restore 0168
        }

        #endregion

        #region FindColumn

        /// <summary>
        /// The find column.
        /// </summary>
        /// <param name="aName">
        /// The a name.
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        private ListColumn FindColumn(string aName)
        {
            foreach (ListColumn tmpColumn in myColumns)
            {
                if (tmpColumn.Name == aName)
                {
                    return tmpColumn;
                }
            }

            return null;
        }

        #endregion

        #region GetBodyRangeForColumn

        /// <summary>
        /// The get body range for column.
        /// </summary>
        /// <param name="aColumn">
        /// The a column.
        /// </param>
        /// <param name="aStartRow">
        /// The a start row.
        /// </param>
        /// <param name="anEndRow">
        /// The an end row.
        /// </param>
        /// <returns>
        /// The <see cref="Range"/>.
        /// </returns>
        private Range GetBodyRangeForColumn(ListColumn aColumn, int aStartRow, int anEndRow)
        {
            Range tmpColumnRange = myContainer.get_Range(aColumn.Name, missing);

            // BBS.ST.BHC.BSP.PDC.ExcelClient.Debug.DebugPDCExcelAddin.GetDebugPDCExcelAddin().Dump(tmpColumnRange);
            Range tmpDataRange = DataRange;

            int tmpStartRow = tmpDataRange.Row + aStartRow;
            int tmpColumn = tmpColumnRange.Column;
            int tmpEndRow = tmpStartRow + (anEndRow - aStartRow);
            if (mySwapColumnsAndRows)
            {
                tmpStartRow = tmpColumnRange.Row;
                tmpColumn = tmpDataRange.Column + aStartRow;
                int tmpEndColumn = tmpDataRange.Column + anEndRow;
                return ExcelUtils.TheUtils.GetRange(myContainer, tmpStartRow, tmpColumn, tmpStartRow, tmpEndColumn);
            }

            return ExcelUtils.TheUtils.GetRange(myContainer, tmpStartRow, tmpColumn, tmpEndRow, tmpColumn);
        }

        #endregion

        #region GetColumnIndex

        /// <summary>
        /// Convenience method to get the current excel column (minus 1) of the specified list column
        /// </summary>
        /// <param name="aColumnName">
        /// </param>
        /// <returns>
        /// The <see cref="int?"/>.
        /// </returns>
        public int? GetColumnIndex(string aColumnName)
        {
            return GetColumnIndex(CurrentListColumnPlacements(), aColumnName);
        }

        /// <summary>
        /// Convenience method to get the current excel column of the specified list column
        /// </summary>
        /// <param name="aColumnMapping">
        /// </param>
        /// <param name="aColumnName">
        /// </param>
        /// <returns>
        /// The <see cref="int?"/>.
        /// </returns>
        private int? GetColumnIndex(Dictionary<ListColumn, int> aColumnMapping, string aColumnName)
        {
            foreach (KeyValuePair<ListColumn, int> tmpPair in aColumnMapping)
            {
                if (tmpPair.Key.Name == aColumnName)
                {
                    return tmpPair.Value;
                }
            }

            return null;
        }

        /// <summary>
        /// Convenience method to get the current excel column (minus 1) of the specified variableID
        /// </summary>
        /// <param name="variableID">
        /// </param>
        /// <returns>
        /// Excel column number
        /// </returns>
        public int? GetColumnIndex(int variableID)
        {
            foreach (KeyValuePair<ListColumn, int> tmpPair in CurrentListColumnPlacements())
            {
                if (tmpPair.Key.TestVariable == null)
                {
                    continue;
                }

                if (tmpPair.Key.TestVariable.VariableId == variableID)
                {
                    return tmpPair.Value;
                }
            }

            return null;
        }

        #endregion

        #region GetColumnValues

        /// <summary>
        /// Returns the values for the specified column within the DataRange
        /// </summary>
        /// <param name="aColumnIdx">
        /// </param>
        /// <returns>
        /// The <see cref="object[,]"/>.
        /// </returns>
        public object[,] GetColumnValues(int aColumnIdx)
        {
            return GetColumnValues(aColumnIdx, false);
        }

        /// <summary>
        /// Returns the column values for the specified column <see cref="CurrentListColumnPlacements"/>
        /// </summary>
        /// <param name="aColumnIdx">
        /// </param>
        /// <param name="completeList">
        /// Returns the values including the insert line if true
        /// </param>
        /// <returns>
        /// The <see cref="object[,]"/>.
        /// </returns>
        public object[,] GetColumnValues(int aColumnIdx, bool completeList)
        {
            Range tmpRange = ColumnRange(aColumnIdx, completeList);
            object tmpValues = tmpRange.get_Value(missing);
            if (tmpValues is object[,])
            {
                return (object[,])tmpValues;
            }

            return new[,] { { tmpValues } };
        }

        #endregion

        #region GetSelectedValues

        /// <summary>
        /// Returns the value matrix of the table, where all cellvalues of unselected rows are set to null.
        /// </summary>
        /// <returns>
        /// The <see cref="object[,]"/>.
        /// </returns>
        public object[,] GetSelectedValues()
        {
            object[,] tmpValues = Values;
            object tmpSelection = Globals.PDCExcelAddIn.Application.Selection;
            if (!(tmpSelection is Range))
            {
                return tmpValues;
            }

            Range tmpRange = (Range)tmpSelection;
            Range tmpDataRange = DataRange;
            int tmpY = tmpDataRange.Row;
            int tmpX = tmpDataRange.Column;
            int tmpLowY = tmpValues.GetLowerBound(0);
            int tmpLowX = tmpValues.GetLowerBound(1);
            int tmpHighY = tmpValues.GetUpperBound(0);
            int tmpHighX = tmpValues.GetUpperBound(1);
            Dictionary<int, int> tmpRows = ExcelUtils.TheUtils.ExtractRowNumbers(tmpRange);
            for (int y = tmpLowY; y <= tmpHighY; y++)
            {
                if (tmpRows.ContainsKey(y + tmpY - tmpLowY))
                {
                    continue;
                }

                for (int x = tmpLowX; x <= tmpHighX; x++)
                {
                    tmpValues[y, x] = null;
                }
            }

            return tmpValues;
        }

        #endregion

        #region GroupColumns

        /// <summary>
        /// The group columns.
        /// </summary>
        /// <param name="aStartColumn">
        /// The a start column.
        /// </param>
        /// <param name="anEndColumn">
        /// The an end column.
        /// </param>
        public void GroupColumns(int aStartColumn, int anEndColumn)
        {
            if (aStartColumn == anEndColumn)
            {
                return;
            }

            int tmpTableStart = HeaderRange.Column;
            Range tmpGroupRange = ExcelUtils.TheUtils.GetRange(myContainer,
                                                            myContainer.Cells[1, tmpTableStart + aStartColumn],
                                                            myContainer.Cells[1, tmpTableStart + anEndColumn - 1]);
            tmpGroupRange.Group(missing, missing, missing, missing);
            object tmpO = myContainer.GroupObjects(missing);
            object tmpW = myContainer.GroupBoxes(missing);
        }

        #endregion

        #region HasMeasurementParamHandler

        /// <summary>
        /// Gets a value indicating whether has measurement param handler.
        /// </summary>
        internal bool HasMeasurementParamHandler
        {
            get { return myMeasurementColumn != null; }
        }

        #endregion

        #region HiddenRows

        /// <summary>
        /// Returns the hidden status for all rows of the DataRange
        /// </summary>
        /// <returns>
        /// The <see cref="bool[]"/>.
        /// </returns>
        public bool[] HiddenRows()
        {
            Range tmpDataRange = DataRange;
            int tmpCount = SwapColumnsAndRows ? tmpDataRange.Columns.Count : tmpDataRange.Rows.Count;
            bool[] tmpHiddenFlags = new bool[tmpCount];
            if (!UserConfiguration.TheConfiguration.GetBooleanProperty(UserConfiguration.PROP_IGNORE_HIDDEN_ROWS, false))
            {
                return tmpHiddenFlags;
            }

            for (int i = 0; i < tmpCount; i++)
            {
                tmpHiddenFlags[i] = isHiddenRow(i, Rectangle.Y, Rectangle.X);
            }

            return tmpHiddenFlags;
        }

        #endregion

        #region InitCellFormat

        /// <summary>
        /// Initializes the column cells with the desired NumberFormat.
        /// Currently affects only the UploadDate and DateResult columns.
        /// </summary>
        /// <param name="columnRange">
        /// </param>
        private void InitCellFormatForDate(Range columnRange)
        {
            CultureInfo myCI = new CultureInfo(CultureInfo.CurrentCulture.LCID);
            columnRange.NumberFormat = myCI.DateTimeFormat.ShortDatePattern;
        }

        #endregion

        #region InitializeNewColumnCells

        /// <summary>
        /// Initializes new rows for the specified column
        /// </summary>
        /// <param name="aColumn">
        /// A column of the list
        /// </param>
        /// <param name="aStartRow">
        /// The first new row
        /// </param>
        /// <param name="anEndRow">
        /// The last new row
        /// </param>
        private void InitializeNewColumnCells(ListColumn aColumn, int aStartRow, int anEndRow)
        {
            PDCLogger.TheLogger.LogStarttime("InitializeNewColumnCells", "InitializeNewColumnCells");
            try
            {
                Range tmpColumnRange = null;
                bool tmpNeedsDateFormatting = NeedsFormatting(aColumn);
                bool tmpNeedsStringFormatting = !tmpNeedsDateFormatting && aColumn.TestVariable != null &&
                                                !aColumn.TestVariable.IsNumeric();

                // Special treatment with this 4 special kind if columns.
                if (aColumn.ReadOnly || aColumn.ParamHandler != null || tmpNeedsDateFormatting ||
                    tmpNeedsStringFormatting)
                {
                    // we need the range for column
                    tmpColumnRange = GetBodyRangeForColumn(aColumn, aStartRow, anEndRow);
                    if (tmpNeedsDateFormatting)
                    {
                        if (!myDateListColumns.ContainsKey(aColumn.Name))
                        {
                            myDateListColumns.Add(aColumn.Name, aColumn);
                        }

                        InitCellFormatForDate(tmpColumnRange);
                    }

                    if (tmpNeedsStringFormatting)
                    {
                        // this is the string format maker. everybody would guess so :-)
                        tmpColumnRange.NumberFormat = "@";
                    }

                    if (aColumn.ReadOnly)
                    {
                        ReadOnly(tmpColumnRange);
                    }

                    // Actually there are only two columns with a parameter handler. One for the RowNumber and one for the Measurements.
                    // So only when there is one of these handlers, the new cells of that paramhandler will be initialized.
                    if (aColumn.ParamHandler != null)
                    {
                        aColumn.ParamHandler.InitializeNewCells(tmpColumnRange, this);
                    }
                }

                initPicklist(aColumn, aStartRow, anEndRow, tmpColumnRange);
            }
            catch (TooManyMeasurementtables tm)
            {
                throw tm;
            }
            catch (Exception e)
            {
                ExceptionHandler.TheExceptionHandler.handleException(e, null);
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("InitializeNewColumnCells", "InitializeNewColumnCells");
            }
        }

        #endregion

        #region initPicklist

        /// <summary>
        /// The init picklist.
        /// </summary>
        /// <param name="aColumn">
        /// The a column.
        /// </param>
        /// <param name="aStartRow">
        /// The a start row.
        /// </param>
        /// <param name="anEndRow">
        /// The an end row.
        /// </param>
        /// <param name="aColumnRange">
        /// The a column range.
        /// </param>
        private void initPicklist(ListColumn aColumn, int aStartRow, int anEndRow, Range aColumnRange)
        {
            if (aColumn.TestVariable == null)
            {
                return;
            }

            PicklistHandler tmpPicklistHandler = PicklistHandler.ThePicklistHandler(myContainer);
            int? tmpPreferedSize = null;
            string tmpLink = tmpPicklistHandler.GetPicklistLink(aColumn.TestVariable, out tmpPreferedSize);
            if (tmpLink != null)
            {
                // Add Validation
                Range tmpColumnRange = aColumnRange;
                if (tmpColumnRange == null)
                {
                    tmpColumnRange = GetBodyRangeForColumn(aColumn, aStartRow, anEndRow);
                }

                if (tmpColumnRange.Validation != null)
                {
                    tmpColumnRange.Validation.Delete();
                }

                tmpColumnRange.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                                              XlFormatConditionOperator.xlBetween, "=" + tmpLink, missing);
                tmpColumnRange.Validation.IgnoreBlank = true;
                tmpColumnRange.Validation.InCellDropdown = true;
                tmpColumnRange.Validation.ShowError = false;
                tmpColumnRange.Validation.ShowInput = false;
                if (tmpPreferedSize != null && tmpPreferedSize > 0)
                {
                    Range tmpColumn = null;
                    if (!mySwapColumnsAndRows)
                    {
                        tmpColumn = ((Range)myContainer.Cells[tmpColumnRange.Row, tmpColumnRange.Column]).EntireColumn;
                        tmpColumn.ColumnWidth = tmpPreferedSize.Value;
                    }
                }
            }
            else
            {
                // Check for numerical lower/upper limit
                decimal? tmpLowerLimit = aColumn.TestVariable.LowerLimit;
                decimal? tmpUpperLimit = aColumn.TestVariable.UpperLimit;
                if (tmpLowerLimit == null)
                {
                    PredefinedParameter tmpPredefined = tmpPicklistHandler.GetPredefinedParameter(aColumn.TestVariable);
                    if (tmpPredefined != null)
                    {
                        tmpLowerLimit = tmpPredefined.LowerLimit;
                        tmpUpperLimit = tmpPredefined.UpperLimit;
                    }
                }

                if (tmpUpperLimit == null)
                {
                    tmpUpperLimit = tmpLowerLimit;
                }

                if (tmpLowerLimit != null)
                {
                    Range tmpColumnRange = GetBodyRangeForColumn(aColumn, aStartRow, anEndRow);
                    if (tmpColumnRange.Validation != null)
                    {
                        tmpColumnRange.Validation.Delete();
                    }

                    tmpColumnRange.Validation.Add(XlDVType.xlValidateDecimal, XlDVAlertStyle.xlValidAlertStop,
                                                  XlFormatConditionOperator.xlBetween, tmpLowerLimit.Value,
                                                  tmpUpperLimit.Value);
                }
            }
        }

        #endregion

        #region isHiddenRow

        /// <summary>
        /// The is hidden row.
        /// </summary>
        /// <param name="aRowNumber">
        /// The a row number.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool isHiddenRow(int aRowNumber)
        {
            return isHiddenRow(aRowNumber, Rectangle.Y, Rectangle.X);
        }

        /// <summary>
        /// Returns true if the specified table row is hidden, false otherwise
        /// </summary>
        /// <param name="aRowNumber">
        /// Table row (starting with 0)
        /// </param>
        /// <param name="aDataRangeRow">
        /// The a Data Range Row.
        /// </param>
        /// <param name="aDataRangeColumn">
        /// The a Data Range Column.
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        public bool isHiddenRow(int aRowNumber, int aDataRangeRow, int aDataRangeColumn)
        {
            PDCLogger.TheLogger.LogStarttime("IsHiddenRow", "IsHiddenRow(" + aRowNumber + ")");
            object tmpHidden = null;
            if (mySwapColumnsAndRows)
            {
                Range tmpRowRange = ((Range)myContainer.Cells[1, aDataRangeColumn + aRowNumber]).EntireColumn;
                tmpHidden = tmpRowRange.Hidden;
            }
            else
            {
                Range tmpRowRange = ((Range)myContainer.Cells[aDataRangeRow + aRowNumber, 2]).EntireRow;
                tmpHidden = tmpRowRange.Hidden;
            }

            PDCLogger.TheLogger.LogStoptime("IsHiddenRow", "IsHiddenRow(" + aRowNumber + ")");
            if (tmpHidden is bool)
            {
                return (bool)tmpHidden;
            }

            return false;
        }

        #endregion

        #region ListColumn

        /// <summary>
        /// Returns the meta data for the specified column
        /// </summary>
        /// <param name="anIndex">
        /// </param>
        /// <returns>
        /// The <see cref="ListColumn"/>.
        /// </returns>
        public ListColumn ListColumn(int anIndex)
        {
            return myColumns[anIndex];
        }

        #endregion

        #region ListRow

        /// <summary>
        /// The list row.
        /// </summary>
        /// <param name="i">
        /// The i.
        /// </param>
        /// <returns>
        /// The <see cref="Range"/>.
        /// </returns>
        private Range ListRow(int i)
        {
            Range tmpBodyRange = DataRange;
            int tmpColumn = tmpBodyRange.Column;
            int tmpWidth = tmpBodyRange.Columns.Count;

            int tmpRow = tmpBodyRange.Row;
            if (mySwapColumnsAndRows)
            {
                tmpWidth = tmpBodyRange.Rows.Count;
                return ExcelUtils.TheUtils.GetRange(myContainer,
                                                 myContainer.Cells[tmpRow, tmpColumn + i],
                                                 myContainer.Cells[tmpRow + tmpWidth - 1, tmpColumn + i]);
            }

            return ExcelUtils.TheUtils.GetRange(myContainer,
                                             myContainer.Cells[tmpRow + i, tmpColumn],
                                             myContainer.Cells[tmpRow + i, tmpColumn + tmpWidth - 1]);
        }

        #endregion

        #region NeedsFormatting

        /// <summary>
        /// Is a special NumberFormat necessary for the specified column?
        /// </summary>
        /// <param name="aColumn">
        /// </param>
        /// <returns>
        /// The <see cref="bool"/>.
        /// </returns>
        private bool NeedsFormatting(ListColumn aColumn)
        {
            return aColumn.Name == PDCExcelConstants.DATERESULT || aColumn.Name == PDCExcelConstants.UPLOADDATE;
        }

        #endregion

        #region ReadOnly

        /// <summary>
        /// Initializes the read-only validation for the specified row-range of a certain column
        /// </summary>
        /// <param name="columnRange">
        /// </param>
        private void ReadOnly(Range columnRange)
        {
            if (columnRange.Validation != null)
            {
                columnRange.Validation.Delete();
            }

            columnRange.Validation.Add(XlDVType.xlValidateCustom, XlDVAlertStyle.xlValidAlertStop,
                                       XlFormatConditionOperator.xlBetween, "IF($A$1=$A$1;False;False)", missing);
            columnRange.Validation.IgnoreBlank = false;
            columnRange.Validation.InputMessage = "This is a read-only parameter and must not be changed.";
            columnRange.Validation.ErrorMessage = "This is a read-only parameter and must not be changed.";
        }

        #endregion

        #region ReplaceWithHyperlinks

        /// <summary>
        /// Some columns (may) contain hyper link (binary files, measurements).
        /// If the table values are returned this method is called to replace the
        /// texts with arrays containing the display text AND the link
        /// </summary>
        /// <param name="aValueMatrix">
        /// Spreadsheet values
        /// </param>
        /// <param name="aStartRow">
        /// Starting at row
        /// </param>
        /// <param name="anEndRow">
        /// untit this row is reached
        /// </param>
        /// <returns>
        /// The <see cref="object[,]"/>.
        /// </returns>
        protected virtual object[,] ReplaceWithHyperlinks(object[,] aValueMatrix, int aStartRow, int anEndRow)
        {
            PDCLogger.TheLogger.LogStarttime("ReadingHyperlinks", "Reading Hyperlinks");
            try
            {
                Dictionary<ListColumn, int> tmpColumnMapping = CurrentListColumnPlacements();
                Range tmpDataRange = DataRange;
                int tmpRowOffset = tmpDataRange.Row + aStartRow;
                int tmpFirstColumn = tmpDataRange.Column;
                foreach (ListColumn tmpColumn in tmpColumnMapping.Keys)
                {
                    if (tmpColumn.IsHyperLink)
                    {
                        int tmpColumnNo = tmpColumnMapping[tmpColumn];

                        Range tmpColumnRange = ExcelUtils.TheUtils.GetRange(myContainer,
                                                                         myContainer.Cells[
                                                                             tmpRowOffset, tmpColumnNo + tmpFirstColumn],
                                                                         myContainer.Cells[
                                                                             tmpRowOffset + (anEndRow - aStartRow),
                                                                             tmpColumnNo + tmpFirstColumn]);

                        IEnumerator tmpHLEnum = tmpColumnRange.Hyperlinks.GetEnumerator();

                        while (tmpHLEnum.MoveNext())
                        {
                            Hyperlink tmpLink = (Hyperlink)tmpHLEnum.Current;
                            if (tmpLink == null)
                            {
                                continue;
                            }

                            int tmpRow = tmpLink.Range.Row - tmpRowOffset;
                            string tmpAddress = tmpLink.Address;
                            if (tmpAddress == null || tmpAddress.Trim() == string.Empty)
                            {
                                tmpAddress = tmpLink.SubAddress;
                            }

                            string[] tmpHyperLinkValue =
                                new[]
                                    {
                                        tmpAddress, string.Empty + aValueMatrix[tmpRow + aValueMatrix.GetLowerBound(0), 
                                                                      tmpColumnNo + aValueMatrix.GetLowerBound(1)]
                                    };
                            aValueMatrix[tmpRow + aValueMatrix.GetLowerBound(0),
                                         tmpColumnNo + aValueMatrix.GetLowerBound(1)] = tmpHyperLinkValue;
                        }
                    }
                }

                return aValueMatrix;
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("ReadingHyperlinks", "Reading Hyperlinks");
            }
        }

        #endregion

        #region ResetMeasurementHyperLinks

        /// <summary>
        /// The reset measurement hyper links.
        /// </summary>
        public void ResetMeasurementHyperLinks()
        {
            if (myMeasurementColumn == null)
            {
                return;
            }

            try
            {
                PDCLogger.TheLogger.LogStarttime("Reset_Hyperlinks", "Resetting hyperlinks");
                Range tmpDataRange = DataRange;
                Range tmpRange = GetBodyRangeForColumn(myMeasurementColumn, 0, tmpDataRange.Rows.Count);
                IEnumerator tmpLinkEnum = tmpRange.Hyperlinks.GetEnumerator();
                while (tmpLinkEnum.MoveNext())
                {
                    Hyperlink tmpHL = (Hyperlink)tmpLinkEnum.Current;
                    Range tmpCell = tmpHL.Range;
                    string tmpSubAdress = tmpHL.SubAddress ?? string.Empty;
                    string tmpAdress = tmpHL.Address ?? string.Empty;
                    string tmpText = tmpHL.TextToDisplay ?? string.Empty;
                    string tmpTip = tmpHL.ScreenTip ?? string.Empty;
                    tmpHL.Delete();
                    if (MeasurementColumn.ParamHandler is MultipleMeasurementTableHandler)
                    {
                        int tableNr = int.Parse(tmpSubAdress.Substring(tmpSubAdress.LastIndexOf("_") + 1)) + 1;
                        MultipleMeasurementTableHandler mmth =
                            (MultipleMeasurementTableHandler)MeasurementColumn.ParamHandler;
                        MeasurementPDCListObject measurementPDCListObject =
                            (MeasurementPDCListObject)mmth.MeasurementTablesDictionary[tableNr];
                        Container.Hyperlinks.Add(tmpCell, string.Empty, measurementPDCListObject.ListRangeName, missing,
                                                 "Measurement");
                    }
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Could not reset hyperlinks", e);
            }

            PDCLogger.TheLogger.LogStoptime("Reset_Hyperlinks", "Resetting hyperlinks");
        }

        #endregion

        #region RestoreValidations

        /// <summary>
        /// Restores the validations for the columns of the given range.
        /// </summary>
        /// <param name="aRange">
        /// </param>
        private void RestoreValidations(Range aRange)
        {
        }

        #endregion

        #region SetColumnValues

        /// <summary>
        /// Sets the values for the specified column in the data range
        /// </summary>
        /// <param name="aColumnIdx">
        /// Column index within the table starting with 0
        /// </param>
        /// <param name="theValues">
        /// The new column values
        /// </param>
        public void SetColumnValues(int aColumnIdx, object[,] theValues)
        {
            SetColumnValues(aColumnIdx, theValues, false);
        }

        /// <summary>
        /// Sets the values for the specified column
        /// </summary>
        /// <param name="aColumnIdx">
        /// Column index within the table starting with 0
        /// </param>
        /// <param name="theValues">
        /// The new column values
        /// </param>
        /// <param name="completeList">
        /// Includes the insert line if set to true
        /// </param>
        public void SetColumnValues(int aColumnIdx, object[,] theValues, bool completeList)
        {
            bool tmpEnabled = Globals.PDCExcelAddIn.Application.EnableEvents;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            try
            {
                Range tmpRange = ColumnRange(aColumnIdx, completeList);
                tmpRange.Value2 = theValues;
            }
            finally
            {
                Globals.PDCExcelAddIn.Application.EnableEvents = tmpEnabled;
            }
        }

        #endregion

        #region SetHyperLinkColumnValues

        /// <summary>
        /// For hyperlink special care is necessary: hyperlinks have a display text and a url. This method sets the
        /// hyperlinks from the key value pairs.
        /// </summary>
        /// <param name="aColumnIndex">
        /// </param>
        /// <param name="theLinksAndLabels">
        /// </param>
        /// <param name="completeList">
        /// </param>
        public void SetHyperLinkColumnValues(int aColumnIndex, KeyValuePair<string, string>?[] theLinksAndLabels,
                                             bool completeList)
        {
            // Set HyperLinks
            Range tmpRange = ColumnRange(aColumnIndex, completeList);
            int tmpX = tmpRange.Column;
            int tmpY = tmpRange.Row;
            if (theLinksAndLabels == null || theLinksAndLabels.Length == 0)
            {
                return; // Nothing to do
            }

            // Set Labels
            // SetColumnValues(aColumnIndex, theValues, completeList);
            int y = -1;
            foreach (KeyValuePair<string, string>? tmpPair in theLinksAndLabels)
            {
                y++;
                if (tmpPair == null)
                {
                    continue;
                }

                string tmpLink = tmpPair.Value.Key;
                string tmpLabel = tmpPair.Value.Value;
                if (tmpLink == null || tmpLink.Trim() == string.Empty)
                {
                    continue;
                }

                Range tmpCell = (Range)Container.Cells[tmpY + y, tmpX];
                try
                {
                    tmpCell.Hyperlinks.Delete();
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Delete Hyperlink", e);
                }

                Container.Hyperlinks.Add(tmpCell, tmpLink, missing, tmpLink, tmpLabel);
            }
        }

        #endregion

        #region ShowAllData

        /// <summary>
        /// The show all data.
        /// </summary>
        internal void ShowAllData()
        {
            try
            {
                if (myContainer.FilterMode)
                {
                    myContainer.ShowAllData();

                    // myContainer.AutoFilterMode = false;
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "ShowAllData", e);
            }
        }

        #endregion

        #region ToListRow

        /// <summary>
        /// Returns the no of the corresponding list row.
        /// Returns -1 if the specified row does not belong to the list.
        /// </summary>
        /// <param name="aSheetRow">
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int ToListRow(int aSheetRow)
        {
            Range tmpDataRange = DataRange;
            int tmpFirstListRow = tmpDataRange.Row;
            int tmpLastListRow = tmpFirstListRow + tmpDataRange.Rows.Count;
            if (tmpFirstListRow > aSheetRow || tmpLastListRow < aSheetRow)
            {
                // Out of Range
                return -1;
            }

            return aSheetRow - tmpFirstListRow;
        }

        #endregion

        #region ToSheetRow

        /// <summary>
        /// Returns the absolute position of the specified list row in the containing worksheet.
        /// </summary>
        /// <param name="aListRow">
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        public int ToSheetRow(int aListRow)
        {
            Range tmpDataRange = DataRange;
            return tmpDataRange.Row + aListRow;
        }

        #endregion

        #region UpdateColumns

        /// <summary>
        /// </summary>
        /// <param name="aColumnList">
        /// </param>
        internal void UpdateColumns(List<ListColumn> aColumnList)
        {
            foreach (ListColumn tmpColumn in aColumnList)
            {
                ListColumn tmpCurrentColumn = FindColumn(tmpColumn.Name);
                Range tmpColHeader = ColumnHeader(tmpColumn.Name);
                if (tmpColumn.Comment != tmpCurrentColumn.Comment)
                {
                    ExcelUtils.TheUtils.AddComment(tmpColHeader, tmpColumn.Comment, true);
                    tmpCurrentColumn.Comment = tmpColumn.Comment;
                }

                if (tmpColumn.Label != tmpCurrentColumn.Label)
                {
                    tmpColHeader.Formula = tmpColumn.Label;
                    tmpCurrentColumn.Label = tmpColumn.Label;
                }

                if (tmpColumn.OleColor != tmpCurrentColumn.OleColor)
                {
                    if (tmpColumn.OleColor.HasValue)
                    {
                        tmpColHeader.Interior.Color = tmpColumn.OleColor;
                    }
                    else
                    {
                        tmpColHeader.Interior.Color = null;
                    }

                    tmpCurrentColumn.OleColor = tmpColumn.OleColor;
                }

                if (tmpCurrentColumn.ParamHandler != null)
                {
                    tmpCurrentColumn.ParamHandler.UpdateColumn(this, tmpCurrentColumn, tmpColumn);
                }
            }
        }

        #endregion

        #region UpdateDataRange

        /// <summary>
        /// The update data range.
        /// </summary>
        /// <param name="aRow">
        /// The a row.
        /// </param>
        /// <param name="aColumn">
        /// The a column.
        /// </param>
        /// <param name="aHeight">
        /// The a height.
        /// </param>
        /// <param name="aWidth">
        /// The a width.
        /// </param>
        protected virtual void UpdateDataRange(int aRow, int aColumn, int aHeight, int aWidth)
        {
            PDCLogger.TheLogger.LogStarttime("UpdateDataRange" + Name, "Updating data range");
            Range tmpDataRange = null;
            Range tmpExtendedDataRange = null;

            Rectangle = new Rectangle(aColumn, aRow, aWidth, aHeight);

            Debug.WriteLine(Name + " :D:FD:" + Rectangle.ToString());
            if (mySwapColumnsAndRows)
            {
                tmpDataRange = ExcelUtils.TheUtils.GetRange(myContainer, myContainer.Cells[aRow, aColumn],
                                                         myContainer.Cells[aRow + aHeight - 1, aColumn - 1 + aWidth]);
                tmpExtendedDataRange = ExcelUtils.TheUtils.GetRange(myContainer, myContainer.Cells[aRow, aColumn],
                                                                 myContainer.Cells[aRow + aHeight - 1, aColumn + aWidth]);
            }
            else
            {
                tmpDataRange = ExcelUtils.TheUtils.GetRange(myContainer, myContainer.Cells[aRow, aColumn],
                                                         myContainer.Cells[aRow + aHeight - 1, aColumn - 1 + aWidth]);
                tmpExtendedDataRange = ExcelUtils.TheUtils.GetRange(myContainer, myContainer.Cells[aRow, aColumn],
                                                                 myContainer.Cells[aRow + aHeight, aColumn - 1 + aWidth]);
            }

            Name tmpDataName = myContainer.Names.Add(myDataRangeName, tmpDataRange, true, missing, missing, missing,
                                                     missing, missing, missing, missing, missing);
            Range tmpListRange = ListRange;

            DefineName(ListRangeName, tmpListRange, true);
            UpdateRectangle();
            PDCLogger.TheLogger.LogStoptime("UpdateDataRange" + Name, "Updated data range");
        }

        #endregion

        #region UpdateRectangle

        /// <summary>
        /// The update rectangle.
        /// </summary>
        internal void UpdateRectangle()
        {
            try
            {
                Range range = DataRange;
                myRectangle = new Rectangle(range.Column, range.Row, range.Columns.Count, range.Rows.Count);
                Debug.WriteLine(Name + " :UpdateRectangle: " + myRectangle.ToString());
                DrawRectangle();
            }

#pragma warning disable 0168
            catch (Exception e)
            {
            }

#pragma warning restore 0168
        }

        #endregion

        #region fill DateListColums

        /// <summary>
        /// The fill date list columns.
        /// </summary>
        private void FillDateListColumns()
        {
            if (myDateListColumns == null) myDateListColumns = new Dictionary<string, ListColumn>();
            foreach (var listColumn in myColumns)
            {
                bool tmpNeedsFormatting = NeedsFormatting(listColumn);
                if (listColumn.ReadOnly || listColumn.ParamHandler != null || tmpNeedsFormatting)
                {
                    if (tmpNeedsFormatting)
                    {
                        if (!myDateListColumns.ContainsKey(listColumn.Name))
                        {
                            myDateListColumns.Add(listColumn.Name, listColumn);
                        }
                    }
                }
            }
        }

        #endregion

        #region writeExcelDateFormat

        /// <summary>
        /// Writes Date Format (system DateFormat) to Date Columns
        /// </summary>
        public virtual void WriteExcelDateFormat()
        {
            int startRow = DataRange.Row;
            int endRow = DataRange.Rows.Count + DataRange.Row;

            FillDateListColumns();

            foreach (var dateListColumn in myDateListColumns.Values)
            {
                Range columnRange = GetBodyRangeForColumn(dateListColumn, 0, DataRange.Rows.Count);
                InitCellFormatForDate(columnRange);
            }
        }

        #endregion

        #endregion

        #region properties

        #region AlreadyUploaded

        /// <summary>
        /// Gets or sets a value indicating whether already uploaded.
        /// </summary>
        public bool AlreadyUploaded
        {
            get { return myIsAlreadyUploaded; }
            set { myIsAlreadyUploaded = value; }
        }

        #endregion

        #region ColumnCount

        /// <summary>
        /// Returns the number of columns in the list
        /// </summary>
        [XmlIgnore]
        public int ColumnCount
        {
            get { return myColumns.Count; }
        }

        #endregion

        #region Columns

        /// <summary>
        /// Gets or sets the columns.
        /// </summary>
        [XmlIgnore]
        public List<ListColumn> Columns
        {
            get { return myColumns; }
            set { myColumns = value; }
        }

        #endregion

        #region Container

        /// <summary>
        /// Gets or sets the container.
        /// </summary>
        [XmlIgnore]
        public Worksheet
            Container
        {
            get
            {
                if (myContainer != null)
                {
                    if (!ExcelUtils.TheUtils.IsSheetReferenceValid(myContainer) && SheetInfo != null)
                    {
                        myContainer = SheetInfo.ExcelSheet;
                    }
                }

                return myContainer;
            }

            set { myContainer = value; }
        }

        #endregion

        #region DataRange

        /// <summary>
        /// Returns an Excel.Range for the data body.
        /// </summary>
        [XmlIgnore]
        public virtual Range DataRange
        {
            get { return myContainer.get_Range(myDataRangeName, missing); }
        }

        #endregion

        #region HeaderRange

        /// <summary>
        /// Returns an Excel.Range for the header
        /// </summary>
        [XmlIgnore]
        public virtual Range HeaderRange
        {
            get { return myContainer.get_Range(myHeaderRangeName, missing); }
        }

        #endregion

        #region ItemCount

        /// <summary>
        /// Returns the number of entries in the list data body
        /// </summary>
        [XmlIgnore]
        public int ItemCount
        {
            get
            {
                if (!mySwapColumnsAndRows)
                {
                    return DataRange.Rows.Count;
                }

                return DataRange.Columns.Count;
            }
        }

        #endregion

        #region ListRange

        /// <summary>
        ///  Returns an Excel.Range for the complete list consisting of the header, data body and an insert row resp. col.
        /// </summary>
        [XmlIgnore]
        public Range ListRange
        {
            get
            {
                Range tmpDataRange = DataRange;
                Range tmpHeaderRange = HeaderRange;
                int tmpLastRow = tmpDataRange.Row + tmpDataRange.Rows.Count;
                int tmpColumn = tmpDataRange.Column + tmpDataRange.Columns.Count - 1;

                Debug.WriteLine("ListRange called:(" + tmpHeaderRange.Row + "," + tmpHeaderRange.Column + "," +
                                tmpLastRow + "," + tmpColumn + ")," + tmpDataRange.Row + "," + tmpDataRange.Rows.Count);

                // da die "insertrow" mit zurück geliefert werden soll, ist der Wert der unteren Zeile bei "normalen" Range 
                // (!swapColumnAndRows) einen höher als bei gekippten range (swapColumnAndRow). 
                // Sind Zeilen und Columns vertauscht, wird aus der InsertRow eine InsertColumn und daher ist für diesen 
                // Fall die letzte Spalte ein Wert höher als bei einem "normalen" Range
                // MIt anderen Worten: Bei gekippten Range wird die InsertRow weggenommen und eine InsertCol hinzugefügt.
                if (mySwapColumnsAndRows)
                {
                    return ExcelUtils.TheUtils.GetRange(myContainer,
                                                     myContainer.Cells[tmpHeaderRange.Row, tmpHeaderRange.Column],
                                                     myContainer.Cells[tmpLastRow - 1, tmpColumn + 1]);
                }

                return ExcelUtils.TheUtils.GetRange(myContainer,
                                                 myContainer.Cells[tmpHeaderRange.Row, tmpHeaderRange.Column],
                                                 myContainer.Cells[tmpLastRow, tmpColumn]);
            }
        }

        #endregion

        #region ListRangeByName

        /// <summary>
        /// Gets the list range by name.
        /// </summary>
        public Range ListRangeByName
        {
            get { return myContainer.get_Range(ListRangeName, missing); }
        }

        #endregion

        #region ListRangeByName

        /// <summary>
        /// Gets a value indicating whether measurement range exists.
        /// </summary>
        public bool MeasurementRangeExists
        {
            get
            {
                try
                {
                    Range testRange = myContainer.get_Range(PDCExcelConstants.MEASUREMENTS, missing);
                    return true;
                }
                catch (Exception e)
                {
                    return false;
                }
            }
        }

        #endregion

        #region ListRangeName

        /// <summary>
        /// name of the list range
        /// </summary>
        [XmlIgnore]
        public string ListRangeName
        {
            get { return "List_" + Name; }
        }

        #endregion

        #region MeasurementColumn

        /// <summary>
        /// Returns the column holding the measurement table links or null if no measurement column exists
        /// </summary>        
        [XmlIgnore]
        public ListColumn MeasurementColumn
        {
            get { return myMeasurementColumn; }
            set { myMeasurementColumn = value; }
        }

        #endregion

        #region Name

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name
        {
            get { return myName; }
            set
            {
                myName = value;
                myHeaderRangeName = "Header_" + Name;
                myDataRangeName = "Data_" + Name;
            }
        }

        #endregion

        #region Rectangle

        /// <summary>
        /// Gets or sets the rectangle.
        /// </summary>
        public Rectangle Rectangle
        {
            get { return myRectangle; }
            set { myRectangle = value; }
        }

        #endregion

        #region RowNumberColumn

        /// <summary>
        /// Returns the list column which contains the row number
        /// </summary>
        [XmlIgnore]
        public ListColumn RowNumberColumn
        {
            get { return myRowNumberColumn; }
            protected set { myRowNumberColumn = value; }
        }

        #endregion

        #region SheetInfo

        /// <summary>
        /// Gets or sets the sheet info.
        /// </summary>
        internal SheetInfo SheetInfo
        {
            get
            {
                if (mySheetInfo != null)
                {
                    return mySheetInfo;
                }

                mySheetInfo = Globals.PDCExcelAddIn.GetSheetInfo(Container);
                return mySheetInfo;
            }

            set { mySheetInfo = value; }
        }

        #endregion

        #region SwapColumnsAndRows

        /// <summary>
        /// Gets or sets a value indicating whether swap columns and rows.
        /// </summary>
        public bool SwapColumnsAndRows
        {
            get { return mySwapColumnsAndRows; }
            set { mySwapColumnsAndRows = value; }
        }

        #endregion

        #region TestDataAdapter

        /// <summary>
        /// Returns the Adapter for Testdata display/extraction
        /// </summary>
        [XmlIgnore]
        public TestDataTableAdapter TestDataAdapter
        {
            get { return myTestDataAdapter; }
            set { myTestDataAdapter = value; }
        }

        #endregion

        #region Testdefinition

        /// <summary>
        /// Accessor for the optional TestDefinition
        /// </summary>
        public Testdefinition Testdefinition
        {
            get { return myTestDefinition; }
            set { myTestDefinition = value; }
        }

        #endregion

        #region this

        /// <summary>
        /// Accessor for the list data rows.
        /// </summary>
        /// <param name="aRowIndex">
        /// </param>
        /// <returns>
        /// The <see cref="object[,]"/>.
        /// </returns>
        public object[,] this[int aRowIndex]
        {
            get
            {
                Range tmpDataRange = ListRow(aRowIndex);
                object tmpValues = tmpDataRange.get_Value(missing);
                if (tmpValues is object[,])
                {
                    return (object[,])tmpValues;
                }

                return new[,] { { tmpValues } };
            }

            set
            {
                Range tmpRowRange = ListRow(aRowIndex);
                tmpRowRange.set_Value(missing, value);
            }
        }

        #endregion

        #region UniqueExperimentKeyHandler

        /// <summary>
        /// Gets or sets the unique experiment key handler.
        /// </summary>
        internal UniqueExperimentKeyHandler UniqueExperimentKeyHandler
        {
            get { return myUniqueExperimentKeyHandler; }
            set { myUniqueExperimentKeyHandler = value; }
        }

        #endregion

        #region ValidationHandler

        /// <summary>
        /// Returns the ValidationHandler for this table
        /// </summary>
        [XmlIgnore]
        internal ValidationHandler ValidationHandler
        {
            get { return myValidationHandler; }
        }

        #endregion

        #region Values

        /// <summary>
        /// Accessor for the complete list data
        /// </summary>
        [XmlIgnore]
        public object[,] Values
        {
            get
            {
                Range tmpBodyRange = DataRange;
                object[,] tmpValues = (object[,])tmpBodyRange.get_Value(missing);
                return ReplaceWithHyperlinks(tmpValues, 0, tmpBodyRange.Rows.Count - 1);
            }

            set
            {
                int tmpYRange = value.GetUpperBound(0) - value.GetLowerBound(0);
                int tmpXRange = value.GetUpperBound(1) - value.GetLowerBound(1);

                ensureCapacity(tmpYRange);

                Range tmpBodyRange = DataRange;
                int tmpRow = tmpBodyRange.Row;
                int tmpColumn = tmpBodyRange.Column;
                Range tmpCopyRange = ExcelUtils.TheUtils.GetRange(myContainer,
                                                               myContainer.Cells[tmpRow, tmpColumn],
                                                               myContainer.Cells[
                                                                   tmpRow + tmpYRange, tmpColumn + tmpXRange]);
                tmpCopyRange.set_Value(missing, value);
            }
        }

        #endregion

        #region InitializeDefaultValues

        /// <summary>
        /// Initializes Default values and updates the pdc table
        /// </summary>
        /// <param name="values">
        /// The current values from the sheet
        /// </param>
        /// <param name="leaveFlags">
        /// </param>
        /// <returns>
        /// True if default values were set, false otherwise
        /// </returns>
        internal bool InitializeDefaultValues(object[,] values, bool[] leaveFlags)
        {
            bool sheetChanged = false;
            Dictionary<ListColumn, int> columnMapping = CurrentListColumnPlacements();
            int offset0 = values.GetLowerBound(0);
            int offset1 = values.GetLowerBound(1);
            Range dataRange = null; // Lazy initialization for better performance
            int x = 0;
            int y = 0;

            int uploadIdColumn = GetUploadIdColumn(columnMapping);
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {
                if (pair.Key.TestVariable == null || pair.Key.TestVariable.DefaultValue == null)
                {
                    continue;
                }

                for (int i = offset0; i <= values.GetUpperBound(0); i++)
                {
                    if (leaveFlags != null && leaveFlags.Length > (i - offset0) && leaveFlags[i - offset0])
                    {
                        // No default values for ignored rows.
                        continue;
                    }

                    if (ValidationHandler.IsEmptyRow(values, columnMapping, i))
                    {
                        continue;
                    }

                    if (uploadIdColumn >= 0 && values[i, uploadIdColumn + offset1] != null &&
                        !string.Empty.Equals(values[i, uploadIdColumn + offset1]))
                    {
                        continue;
                    }

                    if (values[i, pair.Value + offset1] == null || string.Empty.Equals(values[i, pair.Value + offset1]))
                    {
                        values[i, pair.Value + offset1] = pair.Key.TestVariable.DefaultValue;
                        if (dataRange == null)
                        {
                            dataRange = DataRange;
                            x = dataRange.Column;
                            y = dataRange.Row;
                        }

                        sheetChanged = true;
                        if (SwapColumnsAndRows)
                        {
                            Range tmpRange = (Range)Container.Cells[y + pair.Value, x + i - values.GetLowerBound(0)];
                            tmpRange.Value2 = values[i, pair.Value + offset1];
                            tmpRange.Font.Italic = true;
                        }
                        else
                        {
                            Range tmpRange = (Range)Container.Cells[y + i - values.GetLowerBound(0), x + pair.Value];
                            tmpRange.Value2 = values[i, pair.Value + offset1];
                            tmpRange.Font.Italic = true;
                        }
                    }
                }
            }

            return sheetChanged;
        }

        #endregion

        #region GetUploadIdColumn

        /// <summary>
        /// Returns the position of the upload id column or -1 if the column mapping does not contain the
        /// upload id column
        /// </summary>
        /// <param name="columnMapping">
        /// </param>
        /// <returns>
        /// The <see cref="int"/>.
        /// </returns>
        private int GetUploadIdColumn(Dictionary<ListColumn, int> columnMapping)
        {
            foreach (KeyValuePair<ListColumn, int> pair in columnMapping)
            {
                if (pair.Key.Name == PDCExcelConstants.UPLOAD_ID)
                {
                    return pair.Value;
                }
            }

            return -1;
        }

        #endregion

        #endregion
    }
}