using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /// <summary>
    /// The RowNumberHandler is responsible for the row number column, which is used
    /// to interpret the low-level sheet change event as row delete/add/move/modify
    /// </summary>
    [Serializable]
    class RowNumberHandler:PredefinedParameterHandler
    {
        public override void SetValue(PDCListObject pdcTable, object[,] tmpValues, int tmpRow, int tmpPos, ListColumn tmpColumn, BBS.ST.BHC.BSP.PDC.Lib.ExperimentData tmpExperiment, BBS.ST.BHC.BSP.PDC.Lib.TestVariableValue tmpValue)
        {
            tmpValues[tmpRow, tmpPos] = pdcTable.DataRange.Row + tmpRow;
        }
        public override void CellChanged(Microsoft.Office.Interop.Excel.Range aRange, ListColumn aColumn, PDCListObject aPDCList)
        {
            base.CellChanged(aRange, aColumn, aPDCList);
        }

        /// <summary>
        /// The Row number column is needed by pdc and must not be removed
        /// </summary>
        /// <param name="pDCListObject"></param>
        /// <param name="aColumn"></param>
        /// <returns></returns>
        public override bool ColumnDeleted(PDCListObject pDCListObject, ListColumn aColumn)
        {
          MessageBox.Show(string.Format(Properties.Resources.LIST_NEEDED_PDC_COLUMN_DELETED, aColumn.Name), Properties.Resources.MSG_INFO_TITLE);
            return true;
        }
        public override void ClearContents(PDCListObject pdcTable, KeyValuePair<ListColumn, int> aListColumn, object[,] theClearedValues)
        {
            int tmpStartRow = pdcTable.DataRange.Row;
            if (!pdcTable.SwapColumnsAndRows)
            {
                for (int i = theClearedValues.GetLowerBound(0); i <= theClearedValues.GetUpperBound(0); i++)
                {
                    theClearedValues[i, theClearedValues.GetLowerBound(1) + aListColumn.Value] = tmpStartRow + i - theClearedValues.GetLowerBound(0);
                }
            }
            else
            {
                tmpStartRow = pdcTable.DataRange.Column;
                for (int i = theClearedValues.GetLowerBound(1); i <= theClearedValues.GetUpperBound(1); i++)
                {
                    theClearedValues[theClearedValues.GetLowerBound(0) + aListColumn.Value, i] = tmpStartRow + i - theClearedValues.GetLowerBound(1);
                }

            }
        }
        /// <summary>
        ///   This method initialzes the cells in this column,. So each row number will be written in each cell.
        /// </summary>
        /// <param name="aRange"></param>
        /// <param name="aPDCList"></param>
        public override void InitializeNewCells(Microsoft.Office.Interop.Excel.Range aRange, PDCListObject aPDCList)
        {
            if (aPDCList.SwapColumnsAndRows)
            {
                int tmpStartRow = aRange.Column;
                int[,] tmpValues = new int[1, aRange.Columns.Count];
                for (int i = tmpStartRow; i < tmpStartRow + aRange.Columns.Count; i++)
                {
                    tmpValues[0, i - tmpStartRow] = i;
                }
                aRange.Value2 = tmpValues;
                return;
            }
            else
            {
                int tmpStartRow = aRange.Row;
                int[,] tmpValues = new int[aRange.Rows.Count, 1];
                for (int i = tmpStartRow; i < tmpStartRow + aRange.Rows.Count; i++)
                {
                    tmpValues[i - tmpStartRow, 0] = i;
                }
                aRange.Value2 = tmpValues;
            }
        }
    }
}
