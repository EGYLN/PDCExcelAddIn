using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.BHC.BSP.PDC.Lib;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Responsible for the client-side validation of a PDC Sheet and 
    /// its measurement tables.
    /// </summary>
    [ComVisible(false)]
    public class ValidationHandler
    {
        PDCListObject myPdcTable;

        #region constructor
        public ValidationHandler(PDCListObject aPDCList)
        {
            myPdcTable = aPDCList;
        }
        #endregion

        #region methods

        #region CalculateStartRowColumn
        private void CalculateStartRowColumn(object[,] theValues, ref int aStartRow, ref int aStartColumn)
        {
            Excel.Range tmpDataRange = myPdcTable.DataRange;
            aStartColumn = tmpDataRange.Column;
            aStartRow = tmpDataRange.Row;
            if (myPdcTable.SwapColumnsAndRows)
            {
                int tmpX = aStartRow;
                aStartRow = aStartColumn;
                aStartColumn = tmpX;
            }
            aStartRow -= theValues.GetLowerBound(0);
        }
        #endregion

        #region ClearAllValidationMessages
        /// <summary>
        /// Removes all validation messages
        /// </summary>
        internal void ClearAllValidationMessages()
        {
            Globals.PDCExcelAddIn.SetStatusText("Clearing Validation messages");
            ClearAllValidationMessages(myPdcTable);
            if (myPdcTable.MeasurementColumn != null)
            {
                if (myPdcTable.MeasurementColumn.HasMultiMeasurementTableHandler)
                {
                    Predefined.MultipleMeasurementTableHandler mmtHandler = myPdcTable.MeasurementColumn.MultiMeasurementTableHandler;
                    if (mmtHandler.SheetDataRange == null)
                    {
                        if (mmtHandler.MeasurementTables != null)
                        {
                            foreach (PDCListObject list in mmtHandler.MeasurementTables)
                            {
                                ClearAllValidationMessages(list);
                            }
                        }
                    }
                    ExcelUtils.TheUtils.ClearValidationAndDefaults(mmtHandler.SheetDataRange);
                }

                if (myPdcTable.MeasurementColumn.HasSingleMeasurementTableHandler)
                {
                    Predefined.SingleMeasurementTableHandler smtHandler = myPdcTable.MeasurementColumn.SingleMeasurementTableHandler;
                    if (smtHandler.SheetDataRange == null)
                    {
                        if (smtHandler.MeasurementTable != null)
                        {
                            ClearAllValidationMessages(smtHandler.MeasurementTable);
                        }
                    }
                    ExcelUtils.TheUtils.ClearValidationAndDefaults(smtHandler.SheetDataRange);
                }
            }
            Globals.PDCExcelAddIn.SetStatusText(null);
        }

        /// <summary>
        /// Clears all validation messages in the associated table and any measurements table, which may
        /// belong to it.
        /// </summary>
        private void ClearAllValidationMessages(PDCListObject aPDCList)
        {
            try
            {
                Excel.Range tmpDataRange = aPDCList.DataRange;
                ExcelUtils.TheUtils.ClearValidationAndDefaults(tmpDataRange);
                try
                {
                    RemovePrepnoLists(aPDCList);
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Clearing PrepnoLists", e);
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "ClearValidation", e);
            }
        }
        #endregion



        #region DisplayMeasurementValidationMessage
        /// <summary>
        /// Displays a validation for a measurement
        /// </summary>
        /// <param name="aMessage"></param>
        /// <returns>false, if the right place to visualize the message can not be found.</returns>
        public bool DisplayMeasurementValidationMessage(PDCMessage aMessage)
        {
            if (!aMessage.Position.HasValue || !aMessage.VariableNo.HasValue)
            {
                return false;
            }
            int tmpPosition = aMessage.Position.Value - 1; //Use 0-Based offset
            int tmpVariableId = aMessage.VariableNo.Value;

            Dictionary<int, int> tmpPosToRow = ExperimentIndexToRowNoMap();
            if (!tmpPosToRow.ContainsKey(tmpPosition))
            {
                return false;
            }
            Dictionary<ListColumn, int> tmpColumnMapping = myPdcTable.CurrentListColumnPlacements();
            foreach (KeyValuePair<ListColumn, int> tmpPair in tmpColumnMapping)
            {
                if (tmpPair.Key.TestVariable != null && tmpPair.Key.TestVariable.VariableId == tmpVariableId)
                {
                    Excel.Range tmpDataRange = myPdcTable.DataRange;
                    ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(aMessage);
                    DisplayValidationMessage(aMessage.Message, tmpPosToRow[tmpPosition], tmpPair.Value + tmpDataRange.Row, tmpColor);
                    return true;
                }
            }
            return false;
        }

        private bool DisplayMeasurementValidationMessage(PDCMessage tmpMessage, int aTableRow, Dictionary<ListColumn, int> aColumnMapping,
          Testdata aTestdata)
        {
            ListColumn tmpMeasurementColumn = myPdcTable.MeasurementColumn;
            if (tmpMeasurementColumn == null)
            {
                return false;
            }
            if (!aColumnMapping.ContainsKey(tmpMeasurementColumn))
            {
                return false; //Measurement tables deleted
            }
            object[,] tmpValues = ((ExperimentAndMeasurementValues)aTestdata.Tag).experimentValues;

            string tmpLink = "" + tmpValues[aTableRow + tmpValues.GetLowerBound(0), aColumnMapping[tmpMeasurementColumn] + tmpValues.GetLowerBound(1)];
            SheetInfo tmpSheetInfo = myPdcTable.SheetInfo;
            PDCListObject tmpMeasurementTable = tmpSheetInfo.GetMeasurementTable(tmpLink);
            if (tmpMeasurementTable == null)
            {
                return false;
            }

            return tmpMeasurementTable.ValidationHandler.DisplayMeasurementValidationMessage(tmpMessage);
        }
        #endregion

        #region DisplayMessageForUploadParameter
        //Displays a PDC message for an upload parameter.
        private void DisplayMessageForUploadParameter(int aStartColumn, List<PDCMessage> anUnspecificMessages, Dictionary<int, int> aRowMapping,
          Dictionary<ListColumn, int> tmpColumnMapping, PDCMessage tmpMessage, int tmpExperimentNo)
        {
            bool tmpFound = false;

            int? tmpOLEColor = null;
            ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(tmpMessage);
            if (tmpColor != null)
            {
                tmpOLEColor = tmpColor.OleColor;
            }
            foreach (KeyValuePair<ListColumn, int> tmpPair in tmpColumnMapping)
            {
                switch (tmpMessage.ParameterNo)
                {
                    case PDCConstants.C_ID_PREPARATIONNO:
                        tmpFound = tmpPair.Key.Name == PDCExcelConstants.PREPARATIONNO;
                        break;
                    case PDCConstants.C_ID_COMPOUNDIDENTIFIER:
                        tmpFound = tmpPair.Key.Name == PDCExcelConstants.COMPOUNDNO;
                        break;
                    case PDCConstants.C_ID_MCNO:
                        tmpFound = tmpPair.Key.Name == PDCExcelConstants.MCNO;
                        break;
                    default: tmpFound = false; break;
                }
                if (tmpFound)
                {
                    DisplayValidationMessage(tmpMessage.Message, aRowMapping[tmpExperimentNo], aStartColumn + tmpPair.Value, tmpOLEColor);
                    break;
                }
            }
            if (!tmpFound)
            {
                anUnspecificMessages.Add(tmpMessage);
            }
        }
        #endregion

        #region DisplayMessageForVariable
        private void DisplayMessageForVariable(int tmpStartColumn, List<PDCMessage> tmpUnspecificMessages, Dictionary<int, int> tmpRowMapping,
          Dictionary<ListColumn, int> tmpColumnMapping, PDCMessage tmpMessage, int tmpExperimentNo, int? aVariableId)
        {
            bool tmpFound = false;
            foreach (KeyValuePair<ListColumn, int> tmpPair in tmpColumnMapping)
            {
                if (aVariableId != null && tmpPair.Key.TestVariable != null && tmpPair.Key.TestVariable.VariableId == aVariableId.Value)
                {
                    tmpFound = true;
                    ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(tmpMessage);
                    DisplayValidationMessage(tmpMessage.Message, tmpRowMapping[tmpExperimentNo], tmpStartColumn + tmpPair.Value, tmpColor);
                }
            }
            if (!tmpFound)
            {
                tmpUnspecificMessages.Add(tmpMessage);
            }
        }
        #endregion

        #region DisplayValidationMessages
        /// <summary>
        /// Displays the validation messages. 
        /// If a message can be directly associated with a cell, it is visualized using 
        /// </summary>
        /// <param name="aMessageList"></param>
        /// <param name="aTestdata"></param>
        /// <param name="anInteractiveFlag"></param>
        /// <param name="aModal"></param>
        public void DisplayValidationMessages(List<PDCMessage> aMessageList, Testdata aTestdata, bool anInteractiveFlag, bool aModal)
        {
            Globals.PDCExcelAddIn.SetStatusText("Setting server messages in sheet");
            try
            {
                Excel.Range tmpDataRange = myPdcTable.DataRange;
                int tmpStartColumn = tmpDataRange.Column;
                int tmpStartRow = tmpDataRange.Row;
                List<PDCMessage> tmpUnspecificMessages = new List<PDCMessage>();
                Dictionary<int, int> tmpRowMapping = ExperimentIndexToRowNoMap();
                Dictionary<ListColumn, int> tmpColumnMapping = myPdcTable.CurrentListColumnPlacements();
                foreach (PDCMessage tmpMessage in aMessageList)
                {
                    int tmpExperimentNo = tmpMessage.ExperimentIndex;
                    int? tmpVariableNo = tmpMessage.VariableNo;
                    if (!tmpRowMapping.ContainsKey(tmpExperimentNo))
                    {
                        tmpUnspecificMessages.Add(tmpMessage);
                        continue;
                    }

                    if (tmpMessage.Position != null && tmpMessage.Position.Value > 0)
                    {
                        DisplayMeasurementValidationMessage(tmpMessage, tmpRowMapping[tmpExperimentNo] - tmpStartRow, tmpColumnMapping, aTestdata);
                    }
                    if (tmpVariableNo.HasValue)
                    {
                        DisplayMessageForVariable(tmpStartColumn, tmpUnspecificMessages, tmpRowMapping, tmpColumnMapping, tmpMessage, tmpExperimentNo, tmpVariableNo);
                    }
                    else //Use ParameterNo
                    {
                        DisplayMessageForUploadParameter(tmpStartColumn, tmpUnspecificMessages, tmpRowMapping, tmpColumnMapping, tmpMessage, tmpExperimentNo);
                    }
                }
                foreach (PDCMessage tmpPDCMessage in tmpUnspecificMessages)
                {
                    PDCLogger.TheLogger.LogWarning(string.Empty, tmpPDCMessage.Message);
                }

                if (tmpUnspecificMessages.Count > 0 && anInteractiveFlag)
                {
                    MessagesDialog.DisplayMessages(myPdcTable, tmpUnspecificMessages, "Validation", aModal);
                }
            }
            finally
            {
                Globals.PDCExcelAddIn.SetStatusText(null);
            }
        }
        #endregion

        #region DisplayValidationMessage
        private void DisplayValidationMessage(string tmpValidationMessage, int i, int tmpColumnNo, int? aColor)
        {
            Excel.Range tmpCell;
            if (myPdcTable.SwapColumnsAndRows)
            {
                tmpCell = (Excel.Range)myPdcTable.Container.Cells[tmpColumnNo, i];
            }
            else
            {
                tmpCell = (Excel.Range)myPdcTable.Container.Cells[i, tmpColumnNo];
            }
            ExcelUtils.TheUtils.AddComment(tmpCell, tmpValidationMessage, true);
            if (tmpValidationMessage != null && aColor != null)
            {
                tmpCell.Interior.Color = aColor.Value;
            }
        }

        private void DisplayValidationMessage(string tmpValidationMessage, int i, int tmpColumnNo, ClientConfiguration.Color aColor)
        {
            DisplayValidationMessage(tmpValidationMessage, i, tmpColumnNo, aColor == null ? (int?)null : aColor.OleColor);
        }
        #endregion

        #region ExperimentIndexToRowNoMap
        /// <summary>
        /// Hidden rows are ignored by PDC. Therefore we have to translate the experiment indices from
        /// the service to the visible rows.
        /// </summary>
        /// <returns></returns>
        private Dictionary<int, int> ExperimentIndexToRowNoMap()
        {
            Dictionary<int, int> tmpMap = new Dictionary<int, int>();
            int tmpIndex = 0;
            Excel.Range tmpDataRange = myPdcTable.DataRange;
            int tmpRowStart = myPdcTable.SwapColumnsAndRows ? tmpDataRange.Column : tmpDataRange.Row;
            int tmpRowEnd = tmpRowStart + (myPdcTable.SwapColumnsAndRows ? tmpDataRange.Columns.Count : tmpDataRange.Rows.Count);
            for (int i = tmpRowStart; i < tmpRowEnd; i++)
            {
                Excel.Range tmpRange = null;
                if (myPdcTable.SwapColumnsAndRows)
                {
                    tmpRange = ((Excel.Range)myPdcTable.Container.Cells[1, i]).EntireColumn;
                }
                else
                {
                    tmpRange = ((Excel.Range)myPdcTable.Container.Cells[i, 1]).EntireRow;
                }
                object tmpHidden = tmpRange.Hidden;
                if (tmpHidden is bool && ((bool)tmpHidden))
                {
                    tmpIndex++;
                    continue;
                }
                tmpMap.Add(tmpIndex, i);
                tmpIndex++;
            }
            return tmpMap;
        }
        #endregion

        #region RemovePrepnoLists
        /// <summary>
        /// Removes any list validation for the prepno column
        /// </summary>
        /// <param name="aList"></param>
        public void RemovePrepnoLists(PDCListObject aList)
        {
            int? tmpPrepnoColumn = aList.GetColumnIndex(PDCExcelConstants.PREPARATIONNO);
            if (tmpPrepnoColumn == null)
            {
                return;
            }
            Excel.Range tmpRange = aList.ColumnRange(tmpPrepnoColumn.Value, true);
           // tmpRange.Validation.Delete();
        }
        #endregion

        #region IsEmptyRow
        /// <summary>
        /// Returns true if a row does not contain data values.
        /// </summary>
        /// <param name="theValues"></param>
        /// <param name="theColumnMapping"></param>
        /// <param name="aRow"></param>
        /// <returns></returns>
        public bool IsEmptyRow(object[,] theValues, Dictionary<ListColumn, int> theColumnMapping, int aRow)
        {
            foreach (KeyValuePair<ListColumn, int> tmpPlacement in theColumnMapping)
            {
                // todo Implement SingleMeasurementTableHandler
                if (tmpPlacement.Key.ParamHandler is Predefined.MultipleMeasurementTableHandler)
                {
                    continue;
                }
                if (tmpPlacement.Key.ParamHandler is Predefined.RowNumberHandler)
                {
                    continue;
                }
                object tmpValue = theValues[aRow, tmpPlacement.Value + theValues.GetLowerBound(1)];
                string tmpString = tmpValue == null ? "" : ("" + tmpValue).Trim();
                if ("" != tmpString)
                {
                    return false;
                }
            }
            return true;
        }
        #endregion

        #region Transpone
        /// <summary>
        /// Swaps rows and columns of the specified matrix
        /// </summary>
        /// <param name="anOriginal"></param>
        /// <returns></returns>
        internal static object[,] Transpone(object[,] anOriginal)
        {
            object[,] tmpResult = new object[anOriginal.GetLength(1), anOriginal.GetLength(0)];
            for (int i = anOriginal.GetLowerBound(0); i <= anOriginal.GetUpperBound(0); i++)
            {
                for (int j = anOriginal.GetLowerBound(1); j <= anOriginal.GetUpperBound(1); j++)
                {
                    tmpResult[j - anOriginal.GetLowerBound(0), i - anOriginal.GetLowerBound(1)] = anOriginal[i, j];
                }
            }
            return tmpResult;
        }
        #endregion

        #region Validate
        /// <summary>
        /// Validates the values of the associated pdc list object
        /// </summary>
        /// <param name="theValues">The current values from the pdc list object</param>
        /// <param name="theLeaveFlags">Specifies the rows to be ignored. May be null</param>
        /// <returns>Returns true if the validation was successful</returns>
        internal bool Validate(ExperimentAndMeasurementValues theValues, bool[] theLeaveFlags)
        {
            bool tmpOk = true;
            bool tmpIsMeasururementTable = myPdcTable.SwapColumnsAndRows;
            object[,] tmpValues = null;
            try
            {
                if (tmpIsMeasururementTable)
                {
                    if (theValues.measurementValues != null)
                    {
                        tmpValues = theValues.measurementValues[myPdcTable.Name];
                    }
                    if (tmpValues == null)
                    {
                        tmpValues = myPdcTable.Values;
                    }
                }
                else
                {
                    Globals.PDCExcelAddIn.SetStatusText("Client-side validation");
                    PDCLogger.TheLogger.LogStarttime("ValidateTable", "Validating");
                    tmpValues = theValues.experimentValues;
                }
                int tmpStartRow = int.MinValue;
                int tmpStartColumn = int.MinValue;

                if (myPdcTable.SwapColumnsAndRows)
                {
                    tmpValues = Transpone(tmpValues);
                }

                Dictionary<ListColumn, int> tmpColumnMapping = myPdcTable.CurrentListColumnPlacements();
                for (int i = tmpValues.GetLowerBound(0); i <= tmpValues.GetUpperBound(0); i++)
                {
                    if (theLeaveFlags != null && theLeaveFlags.Length > (i - tmpValues.GetLowerBound(0)) && theLeaveFlags[i - tmpValues.GetLowerBound(0)])
                    {
                        continue;
                    }
                    if (IsEmptyRow(tmpValues, tmpColumnMapping, i))
                    {
                        continue;
                    }
                    bool tmpDerivedFound = false;
                    bool tmpCompoundInfoFound = false;
                    foreach (ListColumn tmpColumn in tmpColumnMapping.Keys)
                    {                        int tmpColumnNo = tmpColumnMapping[tmpColumn];
                        object tmpValue = tmpValues[i, tmpColumnNo + tmpValues.GetLowerBound(1)];
                        string tmpTableName;
                        string[] value = tmpValue as string[];
                        if (value != null)
                        {
                            tmpTableName = value[0];
                        }
                        else
                        {
                            tmpTableName = string.Empty + tmpValue;
                        }
                        // todo Implement SingleMeasurementTableHandler
                        if (tmpColumn.ParamHandler is Predefined.MultipleMeasurementTableHandler)
                        {
                            tmpOk &= ValidateMeasurementTable(theValues, string.Empty + tmpTableName);
                        }
                        if (tmpColumn.Validator == null)
                        {
                            continue;
                        }
                        if (tmpValue != null && tmpColumn.TestVariable != null && tmpColumn.TestVariable.IsDerivedResult())
                        {
                            tmpDerivedFound = true;
                        }
                        if (tmpValue != null && (tmpColumn.Name == PDCExcelConstants.COMPOUNDNO || tmpColumn.Name == PDCExcelConstants.PREPARATIONNO || tmpColumn.Name == PDCExcelConstants.MCNO))
                        {
                            tmpCompoundInfoFound = true;
                        }
                        string tmpValidationMessage = tmpColumn.Validator.Validate(tmpColumn, tmpValue, null);

                        if (tmpValidationMessage != null)
                        {
                            tmpOk = false;
                            if (tmpStartColumn == int.MinValue)
                            {
                                CalculateStartRowColumn(tmpValues, ref tmpStartRow, ref tmpStartColumn);
                            }
                            ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.MESSAGE_TYPE_ERROR];
                            DisplayValidationMessage(tmpValidationMessage, tmpStartRow + i, tmpStartColumn + tmpColumnNo, tmpColor);
                        }
                    }
                    if (!tmpIsMeasururementTable && !tmpDerivedFound)
                    {
                        tmpOk = false;
                        foreach (KeyValuePair<ListColumn, int> tmpPair in tmpColumnMapping)
                        {
                            if (tmpPair.Key.TestVariable != null && tmpPair.Key.TestVariable.IsDerivedResult())
                            {
                                if (tmpStartColumn == int.MinValue)
                                {
                                    CalculateStartRowColumn(tmpValues, ref tmpStartRow, ref tmpStartColumn);
                                }

                                ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.MESSAGE_TYPE_ERROR];
                                DisplayValidationMessage(Properties.Resources.VALIDATOR_DERIVED_MISSING, tmpStartRow + i, tmpStartColumn + tmpPair.Value,
                                  tmpColor);
                            }
                        }
                    }
                    if (!tmpIsMeasururementTable && !tmpCompoundInfoFound)
                    {
                        tmpOk = false;
                        foreach (KeyValuePair<ListColumn, int> tmpPair in tmpColumnMapping)
                        {
                            if (tmpPair.Key.Name == PDCExcelConstants.COMPOUNDNO)
                            {
                                if (tmpStartColumn == int.MinValue)
                                {
                                    CalculateStartRowColumn(tmpValues, ref tmpStartRow, ref tmpStartColumn);
                                }
                                DisplayValidationMessage(Properties.Resources.VALIDATOR_COMPOUND_INFO_MISSING, tmpStartRow + i, tmpStartColumn + tmpPair.Value, System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red));
                            }
                        }
                    }
                }
                return tmpOk;
            }
            finally
            {
                if (!tmpIsMeasururementTable)
                {
                    Globals.PDCExcelAddIn.SetStatusText(null);
                    PDCLogger.TheLogger.LogStoptime("ValidateTable", "Validating");
                }
            }
        }
        #endregion

        #region ValidateMeasurementTable
        /// <summary>
        /// Retrieves the specified measurement table and performs the validation for the table
        /// </summary>
        /// <param name="values"></param>
        /// <param name="tablename">The table name as it is used for the list range</param>
        /// <returns>Returns true if the validation succeeds or the table does not exist.</returns>
        private bool ValidateMeasurementTable(ExperimentAndMeasurementValues values, string tablename)
        {
            SheetInfo tmpSheetInfo = myPdcTable.SheetInfo;
            PDCListObject tmpMeasurementTable = tmpSheetInfo.FindMeasurementTable(tablename);
            if (tmpMeasurementTable == null)
            {
                return true;
            }
            return tmpMeasurementTable.ValidationHandler.Validate(values, null);
        }
        #endregion

        #endregion
    }
}
