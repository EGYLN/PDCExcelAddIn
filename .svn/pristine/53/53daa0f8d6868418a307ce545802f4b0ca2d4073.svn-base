using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// PDCListObject for measurement tables with a minimum of Names which relies on
    /// an immutable table structure
    /// </summary>
    [Serializable]
    internal class MeasurementPDCListObject : PDCListObject
    {
        #region constructors
        MeasurementPDCListObject()
            : base()
        {
        }

        MeasurementPDCListObject(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
        internal MeasurementPDCListObject(PDCListObject anOther)
            : base(anOther)
        {
        }
        internal MeasurementPDCListObject(string aName)
            : base(aName)
        {
        }
        public MeasurementPDCListObject(string aName, Excel.Worksheet aSheet, int aStartRow, int aStartColumn, Lib.Testdefinition aTD, bool aSwapColumnAndRowsFlag) :
            base(aName, aSheet, aStartRow, aStartColumn, aTD, aSwapColumnAndRowsFlag, 1, false)
        {
        }
        #endregion

        #region methods

        #region CellsChanged
        public override void CellsChanged(Microsoft.Office.Interop.Excel.Range aRange, SheetEvent anEvent)
        {
            //Do Nothing
        }
        #endregion

        #region CurrentListColumnPlacements
        /// <summary>
        /// Returns the fixed column placements
        /// </summary>
        /// <returns></returns>
        public override Dictionary<ListColumn, int> CurrentListColumnPlacements()
        {
            if (currentColumnMapping != null)
            {
                return currentColumnMapping;
            }
            Dictionary<ListColumn, int> tmpMapping = new Dictionary<ListColumn, int>();
            int i = 0;
            foreach (ListColumn tmpColumn in Columns)
            {
                tmpMapping.Add(tmpColumn, i);
                i++;
            }
            currentColumnMapping = tmpMapping;
            return tmpMapping;
        }
        #endregion

        #region CopyMeasurementTable
        /// <summary>
        /// Makes a copy of the receiver with a new name and a new position. 
        /// Intended for the creation of measurement tables by copying a template measurement table
        /// </summary>
        /// <param name="aNewName">The new name of the measurement table</param>
        /// <param name="aNewX">The new x position</param>
        /// <param name="aNewY">The new y position</param>
        /// <param name="aVariablenameSuffix">The new suffix for the variables</param>
        /// <param name="theColumnMappings">The current column mapping of the table</param>
        /// <returns></returns>
        internal MeasurementPDCListObject CopyMeasurementTable(string aNewName, int aNewX, int aNewY, int? aVariablenameSuffix, Dictionary<ListColumn, int> theColumnMappings)
        {
            PDCLogger.TheLogger.LogStarttime("CopyMeasurementTable", "Copying measurement table " + aNewName);
            MeasurementPDCListObject tmpCopy = new MeasurementPDCListObject(aNewName);
            tmpCopy.SwapColumnsAndRows = SwapColumnsAndRows;
            tmpCopy.AlreadyUploaded = AlreadyUploaded;
            tmpCopy.Columns = new List<ListColumn>();
            tmpCopy.Container = Container;

            tmpCopy.MeasurementColumn = MeasurementColumn;
            tmpCopy.Rectangle = new Rectangle(aNewX + 1, aNewY, Rectangle.Width, Rectangle.Height);

            tmpCopy.Testdefinition = Testdefinition;
            System.Drawing.Rectangle tmpRectangle = Rectangle;
            Excel.Range tmpListRange = ListRangeByName;
            int tmpOffY = aNewY - tmpRectangle.Y;
            int tmpOffX = aNewX - tmpRectangle.X;
            int tmpCountY = tmpRectangle.Height;

            int tmpCountX = aNewX + tmpRectangle.Width - 1;
            if (tmpCountX > max_column)
            {
                tmpCountX = max_column;
            }
            Excel.Range tmpNewPos = (Excel.Range)Container.Cells[aNewY, aNewX];
            Excel.Range tmpTargetRange = ExcelUtils.TheUtils.GetRange(Container, tmpNewPos, Container.Cells[aNewY + tmpCountY - 1, tmpCountX]);
            tmpListRange.Copy(tmpTargetRange);
            foreach (ListColumn tmpColumn in Columns)
            {
                ListColumn tmpNewColumn = tmpColumn.Copy();
                tmpCopy.Columns.Add(tmpNewColumn);
                if (tmpColumn == RowNumberColumn)
                {
                    tmpNewColumn.Name = ROW_COL_PREFIX + aNewName;
                    tmpNewColumn.ParamHandler = new Predefined.RowNumberHandler();
                    tmpCopy.RowNumberColumn = tmpNewColumn;
                }
                else
                {
                    tmpNewColumn.Name = CreateColumnName(tmpColumn.TestVariable.VariableId, aVariablenameSuffix);
                }
            }
            // Define Name Ranges
            PDCLogger.TheLogger.LogStarttime("DefineNames", "Defining Names");
            Excel.Workbook tmpWorkbook = (Excel.Workbook)Container.Parent;
            // ListRange         
            tmpWorkbook.Names.Add(tmpCopy.ListRangeName, tmpTargetRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
            PDCLogger.TheLogger.LogStoptime("DefineNames", "Defining Names");
            PDCLogger.TheLogger.LogStoptime("CopyMeasurementTable", "Copied measurement table " + aNewName);
            return tmpCopy;
        }
        #endregion


        #region ensureCapacity
        public override void ensureCapacity(int aRowCount)
        {
            //Has already maximum size
        }
        #endregion

        #region GetValues
        /// <summary>
        /// Returns the table values from the sheet matrix if available.
        /// Used in "bulk read" to minimize COM-calls to Excel.
        /// </summary>
        /// <param name="theSheetData"></param>
        /// <returns></returns>
        internal object[,] GetValues(MeasurementSheetData theSheetData)
    {
      Rectangle tmpRectangle = theSheetData.RangeArea;
      object[,] tmpSheetValues = theSheetData.Values;
      if (tmpSheetValues == null)
      {
        return Values;
      }
      int tmpYStart = (this.Rectangle.Y - tmpRectangle.Y) + tmpSheetValues.GetLowerBound(0);
      int tmpXStart = (this.Rectangle.X - tmpRectangle.X) + tmpSheetValues.GetLowerBound(1);
      int tmpHeight = SwapColumnsAndRows ? Rectangle.Height: tmpSheetValues.GetLength(0);
      int tmpWidth = SwapColumnsAndRows ? tmpSheetValues.GetLength(1): Rectangle.Width;
      object[,] tmpValues = new object[tmpHeight, tmpWidth];
      for (int i = tmpYStart; i < tmpYStart + tmpHeight; i++)
      {
        for (int j = tmpXStart; j < tmpXStart + tmpWidth; j++)
        {
          tmpValues[i - tmpYStart, j - tmpXStart] = tmpSheetValues[i, j];
        }
      }
      theSheetData[Name] = tmpValues;
      return tmpValues;
    }
        #endregion

        #region ReplaceWithHyperlinks
        /// <summary>
        /// MeasurementTables do not contain hyperlinks
        /// </summary>
        /// <param name="aValueMatrix"></param>
        /// <param name="aStartRow"></param>
        /// <param name="anEndRow"></param>
        /// <returns></returns>
        protected override object[,] ReplaceWithHyperlinks(object[,] aValueMatrix, int aStartRow, int anEndRow)
        {
            return aValueMatrix;
        }
        #endregion

        #region UpdateDataRange
        /// <summary>
        /// Simplified range update 
        /// </summary>
        /// <param name="aRow"></param>
        /// <param name="aColumn"></param>
        /// <param name="aHeight"></param>
        /// <param name="aWidth"></param>
        protected override void UpdateDataRange(int aRow, int aColumn, int aHeight, int aWidth)
        {
            PDCLogger.TheLogger.LogStarttime("UpdateDataRane" + Name, "Updating data range");
            Excel.Range tmpListRange = null;

            Rectangle = new System.Drawing.Rectangle(aColumn, aRow, 1, aHeight);
            tmpListRange = ExcelUtils.TheUtils.GetRange(Container, Container.Cells[aRow, aColumn - 1], Container.Cells[aRow + aHeight - 1, aColumn -1]);

            Container.Names.Add(ListRangeName, tmpListRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
            DrawRectangle();
            PDCLogger.TheLogger.LogStoptime("UpdateDataRane" + Name, "Updated data range");
        }
        #endregion

        #endregion

        #region properties

        #region DataRange
        public override Microsoft.Office.Interop.Excel.Range DataRange
        {
            get
            {
                return ExcelUtils.TheUtils.GetRange(Container, Rectangle.Y, Rectangle.X, Rectangle.Bottom - 1, Rectangle.Right);
            }
        }
        #endregion

        #region HeaderRange
        public override Microsoft.Office.Interop.Excel.Range HeaderRange
        {
            get
            {
                System.Diagnostics.Debug.WriteLine("HeaderRangeM:" + Name + ":" + Rectangle.ToString());
                return ExcelUtils.TheUtils.GetRange(Container,
                  Rectangle.Y, Rectangle.X - 1,
                  Rectangle.Bottom, Rectangle.X - 1);
            }
        }
        #endregion

        #endregion
    }
}
