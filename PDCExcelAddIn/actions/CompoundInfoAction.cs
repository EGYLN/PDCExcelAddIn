﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using BBS.ST.BHC.BSP.PDC.Lib;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions;
using BBS.ST.BHC.BSP.PDC.ExcelClient.actions;
using System.Reflection;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{

    /// <summary>
    /// Implements the lookup of the compound information and validation of compoundno/preparationno/mcno
    /// </summary>
    class CompoundInfoAction : PDCAction
    {
        public const string ACTION_TAG = "PDC_CompoundInfoAction";
        public const string MSG_UNKNOWN_COMPOUND = "unknown compound";
        
        private SheetInfo mySheetInfo;
        private readonly UserSettings myUserSettings;

        #region constructors
        public CompoundInfoAction(bool beginGroup, UserSettings userSettings)
            : base(Properties.Resources.Action_CompoundInfo_Caption, ACTION_TAG, beginGroup)
        {
            myUserSettings = userSettings;
        }

        public CompoundInfoAction(Office.CommandBarPopup aPopup, bool beginGroup, string aCaption, string aTag, UserSettings userSettings)
            : base(aCaption, aTag, beginGroup)
        {
            myUserSettings = userSettings;
            AddToMenu(aPopup, aTag);
        }
        #endregion

        #region methods

        #region ActionCompleted
        private void ActionCompleted(object aResult, ProgressDialog aWindowOwner, bool interactive)
        {
            Globals.PDCExcelAddIn.EnableExcel();
            if (aResult is Exception result && !aWindowOwner.IsDisposed)
            {
                ExceptionHandler.TheExceptionHandler.handleException(result, aWindowOwner);
            }
        }
        #endregion

        #region CheckPerformActionBySelection
        /// <summary>
        /// Validates if performing the CompoundInfoAction on a non-PDC Sheet is possible.
        /// Displays a error message if it is not possible.
        /// </summary>
        /// <returns>Returns if the Action execution on the active sheet is possible</returns>
        private bool CheckPerformActionBySelection()
        {
            if (Globals.PDCExcelAddIn.Application.ActiveSheet is Worksheet sheet) return CheckRangeSelection(sheet);
            MessageBox.Show(
                Properties.Resources.MSG_NO_WORKSHEET_TEXT, 
                Properties.Resources.MSG_ERROR_TITLE,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;

        }
        #endregion

        #region CheckRangeSelection
        /// <summary>
        /// Checks if the selection is a range and that the range is not empty.
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        protected bool CheckRangeSelection(object range)
        {
            if (Globals.PDCExcelAddIn.Application.Selection is Range tmpSelection && tmpSelection.Count > 0)
            {
                return true;
            }

            MessageBox.Show(
                Properties.Resources.MSG_NO_CELLS_SELECTED, 
                Properties.Resources.MSG_ERROR_TITLE,
                MessageBoxButtons.OK, 
                MessageBoxIcon.Error);

            return false;
        }
        #endregion

        #region CreateCompoundInfoAction
        /// <summary>
        /// Creates a CompoundInfoAction based on the specified parameters.
        /// </summary>
        public static CompoundInfoAction CreateCompoundInfoAction(CompoundInfoActionKind anActionKind, Office.CommandBarPopup aPopup, bool beginGroup, UserSettings userSettings)
        {
            string tmpActionTag = ACTION_TAG + "_" + anActionKind;
            string caption = null;
            bool formattingOnly = false;
            switch (anActionKind)
            {
                case CompoundInfoActionKind.All:
                case CompoundInfoActionKind.AllSelected:
                    caption = Properties.Resources.Action_CompoundInfo_Caption;
                    break;
                case CompoundInfoActionKind.FormatSelectedOnly:
                    caption = Properties.Resources.Action_CI_Format_Caption;
                    break;
                case CompoundInfoActionKind.PrepnoSelectedOnly:
                    caption = Properties.Resources.Action_CI_Prepno_Caption;
                    break;
                case CompoundInfoActionKind.FormulaSelectedOnly:
                    caption = Properties.Resources.Action_CI_Formula_Caption;
                    break;
                case CompoundInfoActionKind.StructureSelectedOnly:
                    caption = Properties.Resources.Action_CI_Structure_Caption;
                    break;
                case CompoundInfoActionKind.WeightSelectedOnly:
                    caption = Properties.Resources.Action_CI_WEIGHT_Caption;
                    break;
                case CompoundInfoActionKind.FormatSelectedBay:
                    caption = Properties.Resources.Action_CI_Format_BAY;
                    formattingOnly = true;
                    break;
                case CompoundInfoActionKind.FormatSelectedZk:
                    caption = Properties.Resources.Action_CI_Format_ZK;
                    formattingOnly = true; 
                    break;
                case CompoundInfoActionKind.FormatSelectedCos:
                    caption = Properties.Resources.Action_CI_Format_COS;
                    formattingOnly = true; 
                    break;
                case CompoundInfoActionKind.FormatSelectedCop:
                    caption = Properties.Resources.Action_CI_Format_COP;
                    formattingOnly = true; 
                    break;
            }

            Debug.Assert(!string.IsNullOrEmpty(caption));
            return formattingOnly?
                new FormattingCompoundInfoAction(aPopup,
                    beginGroup,
                    caption,
                    tmpActionTag,
                    userSettings) { Kind = anActionKind }
                : 
                new CompoundInfoAction(aPopup,
                    beginGroup,
                    caption,
                    tmpActionTag,
                    userSettings) { Kind = anActionKind };
        }
        #endregion

        #region DisplayError
        /// <summary>
        /// Examines the validation messages from the server and displays them in the appropriate cells.
        /// </summary>
        private void DisplayError(Worksheet aSheet, 
            TestStruct aCompoundInfo,
            int aCompoundRow, 
            int aCompoundColumn, 
            int aRowOff, 
            int aColumnOff, 
            bool aPrep, 
            bool aMcNo, 
            bool vertical)
        {

            if ((aCompoundInfo.msg != null && !aCompoundInfo.msg.Equals(string.Empty)) ||
                (aCompoundInfo.compoundno_msg != null && !aCompoundInfo.compoundno_msg.Equals(string.Empty)))
            {
                string tmpFailure = "";
                if (!string.IsNullOrEmpty(aCompoundInfo.msg))
                {
                    tmpFailure += aCompoundInfo.msg + "\n";
                }
                if (aCompoundInfo.compoundno_msg != null)
                {
                    tmpFailure += aCompoundInfo.compoundno_msg;
                }
                var tmpCompoundCell = (Range)aSheet.Cells[aCompoundRow, aCompoundColumn];
                ExcelUtils.TheUtils.AddComment(tmpCompoundCell, tmpFailure, false);
                var tmpMessage = new PDCMessage(aCompoundInfo.compoundno_msg, aCompoundInfo.compoundno_id.ToString())
                {
                    MessageCode = aCompoundInfo.compoundno_msg_code
                };
                ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(tmpMessage);
                if (tmpColor != null)
                {
                    tmpCompoundCell.Interior.Color = tmpColor.OleColor;
                }
            }
            int x = aCompoundColumn + aColumnOff;
            int y = aCompoundRow + aRowOff;
            if (aPrep)
            {
                if (!string.IsNullOrEmpty(aCompoundInfo.preparationno_msg))
                {
                    string message = aCompoundInfo.preparationno_msg;
                    Range tmpPreparationCell = (Range)aSheet.Cells[y, x];
                    ExcelUtils.TheUtils.AddComment(tmpPreparationCell, message, false);

                    var tmpMessage = new PDCMessage(message, aCompoundInfo.preparationno_id.ToString())
                    {
                        MessageCode = aCompoundInfo.preparationno_msg_code
                    };
                    ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(tmpMessage);
                    if (tmpColor != null)
                    {
                        tmpPreparationCell.Interior.Color = tmpColor.OleColor;
                    }
                }
            }
            if (vertical)
            {
                y++;
            }
            else
            {
                x++;
            }
            if (aMcNo)
            {
                if (!string.IsNullOrEmpty(aCompoundInfo.mcno_msg))
                {
                    var tmpMcnoCell = (Range)aSheet.Cells[y, x];
                    if (tmpMcnoCell == null)
                    {
                        return;
                    }
                    ExcelUtils.TheUtils.AddComment(tmpMcnoCell, aCompoundInfo.mcno_msg, false);
                    var tmpMessage = new PDCMessage(aCompoundInfo.mcno_msg, aCompoundInfo.mcno_id.ToString())
                    {
                        MessageCode = aCompoundInfo.mcno_msg_code
                    };
                    ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration.GetMessageColor(tmpMessage);
                    if (tmpColor != null)
                    {
                        tmpMcnoCell.Interior.Color = tmpColor.OleColor;
                    }
                }
            }
        }

        #endregion

        #region ExtractColumnMappings
        private static void ExtractColumnMappings(
            Dictionary<ListColumn, int> columnMapping, 
            ref KeyValuePair<ListColumn, int>? compoundPair, 
            ref KeyValuePair<ListColumn, int>? structureDrawingPair, 
            ref KeyValuePair<ListColumn, int>? preparationNoPair, 
            ref KeyValuePair<ListColumn, int>? structureMolWeightPair, 
            ref KeyValuePair<ListColumn, int>? structureMolFormulaPair, 
            ref KeyValuePair<ListColumn, int>? mcNoPair)
        {
            foreach (KeyValuePair<ListColumn, int> tmpPair in columnMapping)
            {
                switch (tmpPair.Key.Name)
                {
                    case PDCExcelConstants.COMPOUNDNO:
                        compoundPair = tmpPair;
                        break;
                    case PDCExcelConstants.STRUCTURE_DRAWING:
                        structureDrawingPair = tmpPair;
                        break;
                    case PDCExcelConstants.PREPARATIONNO:
                        preparationNoPair = tmpPair;
                        break;
                    case PDCExcelConstants.MOLECULAR_WEIGHT:
                        structureMolWeightPair = tmpPair;
                        break;
                    case PDCExcelConstants.FORMULA:
                        structureMolFormulaPair = tmpPair;
                        break;
                    case PDCExcelConstants.MCNO:
                        mcNoPair = tmpPair;
                        break;
                }
            }
        }
        #endregion


        #region GetRange
        private Range GetRange(Worksheet aSheet, int aY1, int aY2, int aX, KeyValuePair<ListColumn, int>? aColumnMapping)
        {
            if (aColumnMapping == null)
            {
                return null;
            }
            return ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aY1, aX + aColumnMapping.Value.Value],
              aSheet.Cells[aY2, aX + aColumnMapping.Value.Value]);
        }
        #endregion


        #region IsHidden
        private static bool IsHidden(Range aRange, bool vertical)
        {
            try
            {
                Range tmpColumn = vertical ? aRange.EntireColumn : aRange.EntireRow;
                return tmpColumn.Hidden != null && (bool)tmpColumn.Hidden;
#pragma warning disable 0168
            }
            catch (Exception)
            {
                return false;
            }
#pragma warning restore 0168
        }
        #endregion

        #region PerformAction
        /// <summary>
        /// Temporarily used
        /// </summary>
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            mySheetInfo = sheetInfo;

            if (Globals.PDCExcelAddIn.Application.ActiveWorkbook == null)
            {
                MessageBox.Show(
                    Properties.Resources.MSG_NO_WORKSHEET_TEXT, 
                    Properties.Resources.MSG_ERROR_TITLE, 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);

                return new ActionStatus();
            }

            if (mySheetInfo == null && !CheckPerformActionBySelection())
            {
                return new ActionStatus();
            }
            
            if (mySheetInfo != null)
            {
                CheckPDCSheet(mySheetInfo);
            }
            AddIn.Application.ScreenUpdating = false;
            AddIn.Application.EnableEvents = false;
            try
            {
                ProgressDialog.Show(ActionCompleted, PerformLookup, Properties.Resources.LABEL_LOOKUP_COMPOUND_INFO);
                return new ActionStatus();
            }
            finally
            {
                AddIn.EnableExcel();
            }
        }
        #endregion

        #region PerformActionBySelection
        /// <summary>
        /// Perform the CompoundInfoAction on the selected cell range.
        /// </summary>
        protected void PerformActionBySelection(Control owner, CompoundInfoActionKind anActionKind)
        {
            var tmpSheet = Globals.PDCExcelAddIn.Application.ActiveSheet as Worksheet;
            if (tmpSheet == null)
            {
                return;
            }

            var tmpSelection = Globals.PDCExcelAddIn.Application.Selection as Range;
            if (tmpSelection == null)
            {
                return;
            }

            Range tmpSelectedRange = tmpSelection;
            long? tmpValue = tmpSelectedRange.CountLarge as long?;
            if (tmpValue == null || tmpValue.Value != 1)
            {
                tmpSelectedRange = tmpSelection.SpecialCells(XlCellType.xlCellTypeVisible);                
            }
            if (tmpSelectedRange.Count == 0)
            {
                return;
            }

            if (Overlaps(tmpSelectedRange))
            {
                MessageBox.Show(Properties.Resources.MSG_OVERLAP, Properties.Resources.MSG_CompoundInfo_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<Ranges> allRanges = 
                (from Range tmpArea in tmpSelectedRange.Areas select CollectWriteRanges(anActionKind, tmpSheet, tmpArea)).ToList();
            if (!CheckEmptyRanges(allRanges.ToArray()))
            {
                return;
            }

            ClearRanges(allRanges.ToArray());
            foreach (var range in allRanges)
            {
                UpdateSheet(owner, anActionKind, mySheetInfo, range, tmpSheet);
            }
        }
        #endregion

        /// <summary>
        /// Collects all ranges which will be written to
        /// </summary>
        /// <param name="actionKind"></param>
        /// <param name="sheet"></param>
        /// <param name="selectedRange"></param>
        /// <returns></returns>
        private Ranges CollectWriteRanges(CompoundInfoActionKind actionKind, Worksheet sheet, Range selectedRange)
        {
            var ranges = new Ranges {OnPdcList = false};
            int tmpStartRow = selectedRange.Row;
            int tmpEndRow = selectedRange.Rows.Count + tmpStartRow - 1;
            int tmpStartCol = selectedRange.Column;
            int tmpEndCol = selectedRange.Columns.Count + tmpStartCol - 1;

            ranges.CompoundnoRange = actionKind == CompoundInfoActionKind.AllSelected ? 
                selectedRange : 
                ExcelUtils.TheUtils.GetRange(sheet, sheet.Cells[tmpStartRow, tmpStartCol], sheet.Cells[tmpEndRow, tmpStartCol]);
            CollectWriteRanges(actionKind, sheet, ranges, tmpStartCol, tmpEndCol, tmpStartRow, tmpEndRow);
            return ranges;
        }
        #region PerformActionForArea

        /// <summary>
        /// 
        /// </summary>
        /// <param name="selectedRange"></param>
        /// <returns></returns>
        protected bool Overlaps(Range selectedRange)
        {
            List<Range> ranges = selectedRange.Areas.Cast<Range>().ToList();
            return Overlaps(ranges);
        }

        private static bool Overlaps(List<System.Drawing.Rectangle> areas, System.Drawing.Rectangle newArea)
        {
            if (areas.Any(area => area.IntersectsWith(newArea)))
            {
                return true;
            }
            areas.Add(newArea);
            return false;
        }
        /// <summary>
        /// Checks for each displayed compound information, if the target area overlaps with any one of the 
        /// other displayed compound informations.
        /// </summary>
        /// <param name="ranges"></param>
        /// <returns></returns>
        protected bool Overlaps(List<Range> ranges)
        {
            var areas = new List<System.Drawing.Rectangle>();
            int infoCount = myUserSettings.DisplayPrepno ? 1 : 0;
            infoCount += myUserSettings.DisplayMolformula ? 1 : 0;
            infoCount += myUserSettings.DisplayMolweight ? 1 : 0;
            infoCount += myUserSettings.DisplayStructure ? 1 : 0;
            foreach (var range in ranges)
            {
                var area = new System.Drawing.Rectangle(range.Column, range.Row, range.Columns.Count, range.Rows.Count);
                if (Overlaps(areas, area))
                {
                    return true;
                }
                for (int i = 0; i < infoCount; i++)
                {
                    if (myUserSettings.Orientation == UserSettings.Direction.Vertical)
                    {
                        area = new System.Drawing.Rectangle(range.Column + myUserSettings.HorizontalOffset,
                            range.Row + myUserSettings.VerticalOffset + i,
                            range.Columns.Count,
                            range.Rows.Count);
                    }
                    else
                    {
                        area = new System.Drawing.Rectangle(range.Column + myUserSettings.HorizontalOffset + i,
                            range.Row + myUserSettings.VerticalOffset,
                            range.Columns.Count,
                            range.Rows.Count);
                    }
                    if (Overlaps(areas, area))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// Checks if the argument ranges are all empty. Otherwise asks the user if the ranges may be overridden.
        /// Returns true if the action can go on.
        /// </summary>
        /// <param name="ranges"></param>
        /// <returns></returns>
        protected bool CheckEmptyRanges(params Ranges[] ranges)
        {
            bool isEmptyRange = true;
            foreach (var range in ranges)
            {
                if (range.OnPdcList || (range.PrepnoRange != null && !ExcelUtils.TheUtils.IsEmptyRange(range.PrepnoRange)))
                {
                    continue;
                }

                if (range.McNoRange != null && !ExcelUtils.TheUtils.IsEmptyRange(range.McNoRange))
                {
                    isEmptyRange = false;
                }
                else if (range.StructureColumn != null && !ExcelUtils.TheUtils.IsEmptyRange(range.StructureColumn))
                {
                    isEmptyRange = false;
                }
                else if (range.FormulaRange != null && !ExcelUtils.TheUtils.IsEmptyRange(range.FormulaRange))
                {
                    isEmptyRange = false;
                }
                else if (range.WeightRange != null && !ExcelUtils.TheUtils.IsEmptyRange(range.WeightRange))
                {
                    isEmptyRange = false;
                }

                if (!isEmptyRange)
                {
                    return DialogResult.Yes == MessageBox.Show(Properties.Resources.MSG_CompoundInfoOverride_Text, Properties.Resources.MSG_CompoundInfo_TITLE, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                }
            }
            return true;
        }

        /// <summary>
        /// Collects the additional ranges in the ranges object which will be written to.
        /// </summary>
        /// <param name="anActionKind"></param>
        /// <param name="aSheet"></param>
        /// <param name="ranges"></param>
        /// <param name="aStartCol"></param>
        /// <param name="anEndCol"></param>
        /// <param name="aStartRow"></param>
        /// <param name="anEndRow"></param>
        private void CollectWriteRanges(CompoundInfoActionKind anActionKind, Worksheet aSheet, Ranges ranges, int aStartCol, int anEndCol, int aStartRow, int anEndRow)
        {
            ranges.Vertical = false;
            switch (anActionKind)
            {
                case CompoundInfoActionKind.All:
                    ranges.PrepnoRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 1], aSheet.Cells[anEndRow, aStartCol + 1]);
                    ranges.WeightRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 3], aSheet.Cells[anEndRow, aStartCol + 3]);
                    ranges.FormulaRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 4], aSheet.Cells[anEndRow, aStartCol + 4]);
                    ranges.StructureColumn = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 2], aSheet.Cells[anEndRow, aStartCol + 2]);
                    break;
                case CompoundInfoActionKind.AllSelected:
                    int tmpXIncr;
                    int tmpYIncr;
                    int tmpEndY = anEndRow;
                    int tmpEndX = anEndCol;
                    if (myUserSettings.Orientation == UserSettings.Direction.Horizontal)
                    {
                        tmpXIncr = 1; tmpYIncr = 0;
                        ranges.Vertical = false;
                    }
                    else
                    {
                        tmpXIncr = 0; tmpYIncr = 1;
                        ranges.Vertical = true;
                    }

                    int tmpXOff = myUserSettings.HorizontalOffset;
                    int tmpYOff = myUserSettings.VerticalOffset;
                    // add offsets
                    if (myUserSettings.DisplayPrepno)
                    {
                        ranges.PrepnoRange = GetRange(aSheet, aStartRow + tmpYOff, aStartCol + tmpXOff, tmpEndY + tmpYOff, tmpEndX + tmpXOff);
                        tmpXOff += tmpXIncr;
                        tmpYOff += tmpYIncr;
                    }
                    //check sheet borders
                    if (myUserSettings.DisplayStructure)
                    {
                        ranges.StructureColumn = GetRange(aSheet, aStartRow + tmpYOff, aStartCol + tmpXOff, tmpEndY + tmpYOff, tmpEndX + tmpXOff);
                        tmpXOff += tmpXIncr;
                        tmpYOff += tmpYIncr;
                    }
                    if (myUserSettings.DisplayMolweight)
                    {
                        ranges.WeightRange = GetRange(aSheet, aStartRow + tmpYOff, aStartCol + tmpXOff, tmpEndY + tmpYOff, tmpEndX + tmpXOff);
                        tmpXOff += tmpXIncr;
                        tmpYOff += tmpYIncr;
                    }
                    if (myUserSettings.DisplayMolformula)
                    {
                        ranges.FormulaRange = GetRange(aSheet, aStartRow + tmpYOff, aStartCol + tmpXOff, tmpEndY + tmpYOff, tmpEndX + tmpXOff);
                    }
                    break;
                case CompoundInfoActionKind.PrepnoSelectedOnly:
                    ranges.PrepnoRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 1], aSheet.Cells[anEndRow, aStartCol + 1]);
                    break;
                case CompoundInfoActionKind.FormulaSelectedOnly:
                    ranges.FormulaRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 1], aSheet.Cells[anEndRow, aStartCol + 1]);
                    break;
                case CompoundInfoActionKind.StructureSelectedOnly:
                    ranges.StructureColumn = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 1], aSheet.Cells[anEndRow, aStartCol + 1]);
                    break;
                case CompoundInfoActionKind.WeightSelectedOnly:
                    ranges.WeightRange = ExcelUtils.TheUtils.GetRange(aSheet, aSheet.Cells[aStartRow, aStartCol + 1], aSheet.Cells[anEndRow, aStartCol + 1]);
                    break;
            }
        }

        private static Range GetRange(Worksheet aSheet, int startY, int startX, int endY, int endX)
        {
            if (!startY.InRange(0, PDCListObject.max_row) ||
                !startX.InRange(0, PDCListObject.max_column))
            {
                throw new InvalidRangeException();
            }

            return aSheet.Range[aSheet.Cells[startY, startX], aSheet.Cells[endY, endX]];
        }
        #endregion

        #region PerformActionOnPDCSheet
        private void PerformActionOnPDCSheet(Control owner, CompoundInfoActionKind anActionKind)
        {
            PDCListObject tmpList = mySheetInfo.MainTable;
            Dictionary<ListColumn, int> tmpColumnMapping = tmpList.CurrentListColumnPlacements();
            KeyValuePair<ListColumn, int>? tmpCompoundPair = null;
            KeyValuePair<ListColumn, int>? tmpStructureDrawingPair = null;
            KeyValuePair<ListColumn, int>? tmpPreparationNoPair = null;
            KeyValuePair<ListColumn, int>? tmpStructureMolWeightPair = null;
            KeyValuePair<ListColumn, int>? tmpStructureMolFormulaPair = null;
            KeyValuePair<ListColumn, int>? tmpMcNoPair = null;

            // Get ColumnMappings
            ExtractColumnMappings(
                tmpColumnMapping, 
                ref tmpCompoundPair, 
                ref tmpStructureDrawingPair, 
                ref tmpPreparationNoPair, 
                ref tmpStructureMolWeightPair, 
                ref tmpStructureMolFormulaPair, 
                ref tmpMcNoPair);

            if (!tmpCompoundPair.HasValue)
            {
                PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_EXCEL, "The compoundno column is missing");
            }

            ExcelUtils.TheUtils.DeleteShapes(mySheetInfo);
            if (anActionKind != CompoundInfoActionKind.All)
            {
                PerformActionOnPDCSheetBySelection(
                    owner,
                    anActionKind, 
                    tmpList,
                    tmpCompoundPair,
                    tmpPreparationNoPair,
                    tmpMcNoPair,
                    tmpStructureDrawingPair,
                    tmpStructureMolFormulaPair,
                    tmpStructureMolWeightPair);
                return;
            }

            Debug.Assert(tmpCompoundPair != null);
            Debug.Assert(tmpPreparationNoPair != null);

            var writeRange = new Ranges
            {
                Vertical = false,
                OnPdcList = true,
                CompoundnoRange = tmpList.ColumnRange(tmpCompoundPair.Value.Value, false),
                PrepnoRange = tmpList.ColumnRange(tmpPreparationNoPair.Value.Value, false)
            };

            Worksheet tmpSheet = tmpList.Container;
            // definition of the column offset for Structure Drawing column
            if (tmpMcNoPair != null)
            {
                writeRange.McNoRange = tmpList.ColumnRange(tmpMcNoPair.Value.Value, false);
            }
            //Set values
            if (tmpStructureMolFormulaPair != null)
            {
                writeRange.FormulaRange = tmpList.ColumnRange(tmpStructureMolFormulaPair.Value.Value, false);
            }
            if (tmpStructureMolWeightPair != null)
            {
                writeRange.WeightRange = tmpList.ColumnRange(tmpStructureMolWeightPair.Value.Value, false);
            }
            if (tmpStructureDrawingPair != null)
            {
                writeRange.StructureColumn = tmpList.ColumnRange(tmpStructureDrawingPair.Value.Value, false);
            }
            ClearRanges(writeRange);
            UpdateSheet(owner, anActionKind, mySheetInfo, writeRange, tmpSheet);
        }
        #endregion

        #region PerformActionOnPDCSheetBySelection

        /// <summary>
        /// Performs the compound info action on the PDC sheet for all action kinds that are
        /// based on the current selection instead of the data entry area
        /// </summary>
        /// <param name="owner"></param>
        /// <param name="anActionKind">Kind of Compoundinfo</param>
        /// <param name="aList">The PDCList on the sheet</param>
        /// <param name="aCompoundNoPair">The column mapping for the compound no</param>
        /// <param name="aPrepNoPair">The column mapping for the prepno</param>
        /// <param name="aMcNoPair">The column mapping for the mc no</param>
        /// <param name="aStructureDrawingPair">The column mapping for the structure drawing</param>
        /// <param name="aFormulaPair">The column mapping for the formula</param>
        /// <param name="aWeightPair">The column mapping for the weight</param>
        private void PerformActionOnPDCSheetBySelection(Control owner,
          CompoundInfoActionKind anActionKind,
          PDCListObject aList,
          KeyValuePair<ListColumn, int>? aCompoundNoPair,
          KeyValuePair<ListColumn, int>? aPrepNoPair,
          KeyValuePair<ListColumn, int>? aMcNoPair,
          KeyValuePair<ListColumn, int>? aStructureDrawingPair,
          KeyValuePair<ListColumn, int>? aFormulaPair,
          KeyValuePair<ListColumn, int>? aWeightPair
          )
        {

            if (!CheckRangeSelection(aList.Container))
            {
                return;
            }
            var ranges = new List<Ranges>();
            Worksheet tmpSheet = aList.Container;

            var currentSelection = (Range) Globals.PDCExcelAddIn.Application.Selection;
            if (currentSelection.Areas.Count > 1 || currentSelection.Rows.Count > 1)
            {
                currentSelection = currentSelection.SpecialCells(XlCellType.xlCellTypeVisible);
            }

            Range tmpListRange = aList.ListRangeByName;
            System.Drawing.Rectangle tmpListRectangle = aList.Rectangle;
            int tmpX = tmpListRectangle.X;
            var nonPdcListRanges = new List<Range>();
            foreach (Range tmpArea in currentSelection.Areas)
            {

                //check if there is an intersection with the PDCList
                Range tmpIntersection = PDCExcelAddIn.TheSingleton().Intersect(tmpArea, tmpListRange);
                if (tmpIntersection == null || tmpIntersection.Count == 0)
                {
                    //If no intersection could be found forget about the pdc sheet
                    nonPdcListRanges.Add(tmpArea);
                }
                else
                {
                    var range = new Ranges
                    {
                        OnPdcList = true,
                        Vertical = false
                    };
                    //range intersects with the pdc list. use the columns from the list.
                    int tmpY1 = tmpArea.Row;
                    int tmpY2 = tmpY1 + (tmpArea.Rows.Count - 1);
                    switch (anActionKind)
                    {
                        case CompoundInfoActionKind.AllSelected:
                            range.StructureColumn = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aStructureDrawingPair);
                            range.WeightRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aWeightPair);
                            range.PrepnoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aPrepNoPair);
                            range.McNoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aMcNoPair);
                            range.FormulaRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aFormulaPair);
                            break;
                        case CompoundInfoActionKind.FormatSelectedOnly:
                            range.PrepnoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aPrepNoPair);
                            range.McNoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aMcNoPair);
                            break;
                        case CompoundInfoActionKind.FormulaSelectedOnly:
                            range.FormulaRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aFormulaPair);
                            break;
                        case CompoundInfoActionKind.PrepnoSelectedOnly:
                            range.PrepnoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aPrepNoPair);
                            break;
                        case CompoundInfoActionKind.StructureSelectedOnly:
                            range.StructureColumn = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aStructureDrawingPair);
                            break;
                        case CompoundInfoActionKind.WeightSelectedOnly:
                            range.WeightRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aWeightPair);
                            break;
                        default:
                            return;
                    }
                    range.CompoundnoRange = GetRange(tmpSheet, tmpY1, tmpY2, tmpX, aCompoundNoPair);
                    ranges.Add(range);
                }
            }
            if (nonPdcListRanges.Count > 0)
            {
                if (Overlaps(nonPdcListRanges))
                {
                    MessageBox.Show(Properties.Resources.MSG_OVERLAP, Properties.Resources.MSG_CompoundInfo_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                ranges.AddRange(nonPdcListRanges.Select(range => CollectWriteRanges(anActionKind, tmpSheet, range)));
            }
            if (!CheckEmptyRanges(ranges.ToArray()))
            {
                return;
            }

            ClearRanges(ranges.ToArray());
            foreach (var nextRange in ranges)
            {
                UpdateSheet(owner, anActionKind, mySheetInfo, nextRange, tmpSheet);
            }
        }
        #endregion

        /// <summary>
        /// Clears the ranges which will be overwritten
        /// </summary>
        /// <param name="ranges"></param>
        private void ClearRanges(params Ranges[] ranges)
        {
            var excelRanges = new List<Range>();
            Worksheet container = null;
            foreach (var range in ranges)
            {
                if (range.CompoundnoRange != null)
                {
                    excelRanges.Add(range.CompoundnoRange);
                }
                if (range.PrepnoRange != null)
                {
                    excelRanges.Add(range.PrepnoRange);
                }
                if (range.McNoRange != null)
                {
                    excelRanges.Add(range.McNoRange);
                }
                if (range.FormulaRange != null)
                {
                    excelRanges.Add(range.FormulaRange);
                }
                if (range.WeightRange != null)
                {
                    excelRanges.Add(range.WeightRange);
                }
                if (range.StructureColumn != null)
                {
                    excelRanges.Add(range.StructureColumn);
                    container = (Worksheet)range.StructureColumn.Parent;
                }
            }
            if (container != null)
            {
                if (mySheetInfo != null)
                {
                    ExcelUtils.TheUtils.DeleteShapes(mySheetInfo);
                }
            }
            ExcelUtils.TheUtils.ClearValidationAndDefaults(excelRanges.ToArray());
        }

        #region PerformLookup
        private void PerformLookup(ProgressDialog aWindowOwner)
        {
            Exception tmpResult = null;
            try
            {
                Globals.PDCExcelAddIn.Application.Cursor = XlMousePointer.xlWait;
                AddIn.Application.EnableEvents = false;
                AddIn.Application.ScreenUpdating = false;
                //Get CompoundNos
                if (mySheetInfo == null)
                {
                    if (Kind == CompoundInfoActionKind.All)
                    {
                        Kind = CompoundInfoActionKind.AllSelected;
                    }
                    PerformActionBySelection(aWindowOwner, Kind);
                    return;
                }
                PerformActionOnPDCSheet(aWindowOwner, Kind);
            }
            catch (Exception e)
            {
                tmpResult = e;
            }
            finally
            {
                aWindowOwner.StatusCallback(tmpResult);
            }
        }
        #endregion


        #region UpdateSheet

        /// <summary>
        /// Retrieves the compound infos from the webservices and places the retrieved data in
        /// the appropriate columns
        /// </summary>
        /// <param name="anActionKind"></param>
        /// <param name="aSheetInfo"></param>
        /// <param name="writeRanges"></param>
        /// <param name="sheet"></param>
        /// <param name="owner"></param>
        protected virtual void UpdateSheet(Control owner, CompoundInfoActionKind anActionKind, SheetInfo aSheetInfo, Ranges writeRanges, Worksheet sheet)
        {
            PDCLogger.TheLogger.LogStarttime("CompoundInfo", "CompoundInfoAction");
            PremarationNumber.PrepartionNo = new List<string>();
            for (int i = 1; i <= writeRanges.CompoundnoRange.Areas.Count; i++)
            {
                var ranges = writeRanges.GetRangesFromArea(i);
                int tmpX;
                int tmpY;
                int tmpColumn;
                int tmpRow;

                bool tmpShowStructureDrawings = ranges.StructureColumn != null && !IsHidden(ranges.StructureColumn, true);
                object[,] tmpPreparationValues = ExcelUtils.TheUtils.RangeToMatrix(ranges.PrepnoRange);
                object[,] tmpMcNoValues = ExcelUtils.TheUtils.RangeToMatrix(ranges.McNoRange);
                object[,] tmpWeights = ExcelUtils.TheUtils.RangeToMatrix(ranges.WeightRange);
                object[,] tmpFormulas = ExcelUtils.TheUtils.RangeToMatrix(ranges.FormulaRange);
                bool useMdl = ExcelShapeUtils.TheUtils.UseIsisOrMdl();
                //compoundno
                var tmpCompoundValues = ExcelUtils.TheUtils.RangeToMatrix(ranges.CompoundnoRange);

                //Get Information from Service
                var tmpCompoundInfos = new TestStruct[ranges.CompoundnoRange.Rows.Count, ranges.CompoundnoRange.Columns.Count];

                PDCLogger.TheLogger.LogStarttime("GettingCompoundInfo", "Getting");
                var tmpInputTestStruct = new TestStruct();
                for (tmpY = 0; tmpY < ranges.CompoundnoRange.Rows.Count; tmpY++)
                {
                    tmpRow = tmpY;
                    for (tmpX = 0; tmpX < ranges.CompoundnoRange.Columns.Count; tmpX++)
                    {
                        tmpColumn = tmpX;
                        PDCLogger.TheLogger.LogStarttime("GettingCompoundInfo_Data LogUpdatedByAnshu", "Getting Data");
                        string tmpCompoundNo = ("" + tmpCompoundValues[tmpRow, tmpColumn]).Trim();
                        string tmpPreparationNo = tmpPreparationValues == null ? null : ("" + tmpPreparationValues[tmpRow, tmpColumn]).Trim();
                        string tmpMcNo = tmpMcNoValues == null ? null : ("" + tmpMcNoValues[tmpRow, tmpColumn]).Trim();
                        if (tmpCompoundNo.Trim() == "" && (tmpPreparationNo == null || tmpPreparationNo.Trim() == "") && (tmpMcNo == null || tmpMcNo.Trim() == ""))
                        {
                            continue; //skip empty rows
                        }

                        tmpInputTestStruct.compoundno = tmpCompoundNo;
                        tmpInputTestStruct.preparationno = tmpPreparationNo;
                        tmpInputTestStruct.mcno = tmpMcNo;
                        tmpInputTestStruct.hydrogendisplaymode = myUserSettings.HydrogenDisplayMode.ToString();

                        tmpInputTestStruct.fileformat = (ExcelShapeUtils.TheUtils.UseIsisOrMdl()) ? "MOL" : "BMP";
                        tmpInputTestStruct.username = Globals.PDCExcelAddIn.PdcService.UserInfo.Cwid;

                        TestStruct tmpOutputTestStruct = PDCService.ThePDCService.GetCompoundInformation(tmpInputTestStruct);

                        tmpCompoundInfos[tmpRow, tmpColumn] = tmpOutputTestStruct;
                        PDCLogger.TheLogger.LogStoptime("GettingCompoundInfo_Data", "Getting Data");
                    }
                }

                PDCLogger.TheLogger.LogStoptime("GettingCompoundInfo LogUpdatedByAnshu", "Getting");
                bool tmpEnabled = Globals.PDCExcelAddIn.Application.EnableEvents;
                PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Set value of sheet visible=" + sheet.Visible);
                var oldValue = sheet.Visible;
                Worksheet temporarySheet = null;
                try
                {
                    Globals.PDCExcelAddIn.Application.EnableEvents = false;
                    if (Application.ActiveWorkbook.Sheets.Count == 1)
                    {
                        //Otherwise it will not be possible to hide the current sheet.
                        temporarySheet = (Worksheet) Application.ActiveWorkbook.Sheets.Add(After: sheet);
                    }
                    PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Before xlSheetHidden");
                    sheet.Visible = XlSheetVisibility.xlSheetHidden;
                    PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "After xlSheetHidden");

                    PDCLogger.TheLogger.LogStoptime("SettingCompoundInfo", "Setting");
                    for (tmpY = 0; tmpY < ranges.CompoundnoRange.Rows.Count; tmpY++)
                    {
                        tmpRow = tmpY;                        
                        for (tmpX = 0; tmpX < ranges.CompoundnoRange.Columns.Count; tmpX++)
                        {
                            bool last = tmpY == ranges.CompoundnoRange.Rows.Count - 1 && tmpX == ranges.CompoundnoRange.Columns.Count-1;
                            tmpColumn = tmpX;
                            if (tmpCompoundInfos[tmpRow, tmpColumn] != null)
                            {
                                PDCLogger.TheLogger.LogStarttime("SettingCompoundInfo_Data", "Setting Data");
                                // place Image im column with Offset for StructureDrawingPair ???!!! 
                                tmpCompoundValues[tmpRow, tmpColumn] = tmpCompoundInfos[tmpRow, tmpColumn].compoundno;
                                if (tmpCompoundInfos[tmpRow, tmpColumn].msg == null || tmpCompoundInfos[tmpRow, tmpColumn].msg.Trim() == "")
                                {
                                    if (tmpWeights != null)
                                    {
                                        // this is  regarding dummy prep changes
                                        //if (IsUnknownCompound(tmpCompoundInfos[tmpRow, tmpColumn]))
                                        //{
                                        //    tmpWeights[tmpRow, tmpColumn] = MSG_UNKNOWN_COMPOUND;
                                        //}
                                        //else
                                        if (Math.Abs(tmpCompoundInfos[tmpRow, tmpColumn].molweight) > 0)
                                        {
                                            tmpWeights[tmpRow, tmpColumn] = tmpCompoundInfos[tmpRow, tmpColumn].molweight;
                                        }
                                    }
                                    if (tmpFormulas != null)
                                    {
                                        if (IsUnknownCompound(tmpCompoundInfos[tmpRow, tmpColumn]))
                                        {
                                            //This is  regarding dummy prep changes
                                            // tmpFormulas[tmpRow, tmpColumn] = MSG_UNKNOWN_COMPOUND;
                                        }
                                        else
                                        {

                                            tmpFormulas[tmpRow, tmpColumn] = tmpCompoundInfos[tmpRow, tmpColumn].molformula;
                                        }
                                    }
                                }

                                TestStruct tmpCis = tmpCompoundInfos[tmpRow, tmpColumn];
                                int xOffset = myUserSettings.HorizontalOffset;
                                int yOffset = myUserSettings.VerticalOffset;
                                if (writeRanges.OnPdcList || anActionKind != CompoundInfoActionKind.AllSelected)
                                {
                                    yOffset = 0;
                                    xOffset = 1;
                                }

                                DisplayError(
                                    sheet,
                                    tmpCis,
                                    ranges.CompoundnoRange.Row + tmpRow,
                                    ranges.CompoundnoRange.Column + tmpColumn,
                                    yOffset,
                                    xOffset,
                                    writeRanges.PrepnoRange != null,
                                    writeRanges.McNoRange != null,
                                    writeRanges.Vertical);

                                // in tmpCompoundInfos[tmpCompoundNoArray[j]].CompoundNo steht jetzt hilfsweise der Link auf die Datei

                                // no MDL/Draw installation

                                if (tmpShowStructureDrawings)
                                {
                                    string shapeName = null;
                                    var column = tmpColumn;
                                    var row = tmpRow;
                                    owner.Invoke(new MethodInvoker(() =>
                                    {
                                        shapeName = ExcelShapeUtils.TheUtils.InsertStructureDrawing(
                                            sheet,
                                            tmpCis,
                                            ranges.StructureColumn.Column + column,
                                            ranges.StructureColumn.Row + row,
                                            myUserSettings,
                                            Properties.Settings.Default.MolfileServletPath,
                                            useMdl, last);
                                    }));

                                    if (aSheetInfo != null && shapeName != null) aSheetInfo.ShapeFileNames.Add(shapeName);
                                }
                                if (ranges.PrepnoRange != null)
                                {
                                    var compRow = ranges.CompoundnoRange.Row + tmpRow;
                                    var compCol = ranges.CompoundnoRange.Column + tmpColumn;
                                    var compoundNoCellRange = (Range)sheet.Cells[compRow, compCol];
                                    // place PreparationNo im column with Offset for PreparationNo
                                    // if more than one preparation number associated with the compoundNo
                                    // then put the ProparationNo's list as validation list to the cells
                                    // otherwise place the 
                                    var prepRow = ranges.PrepnoRange.Row + tmpRow;
                                    var prepCol = ranges.PrepnoRange.Column + tmpColumn;
                                    var tmpCellRangePreparationNo = (Range)sheet.Cells[prepRow, prepCol];
                                    string prepNoList = tmpCis.preparationno ?? "";
                                    // it is related to dummy prep number change(jira  : pdc-926)

                                    if (prepNoList.EndsWith(";") || string.IsNullOrEmpty(prepNoList))
                                        prepNoList += "UNKNOWN;";

                                    else if (!prepNoList.EndsWith(";")) // dummy prep 25/08/2022
                                          prepNoList += ";UNKNOWN;";

                                    if (prepNoList.EndsWith(";"))
                                    {
                                        prepNoList = prepNoList.Substring(0, prepNoList.Length - 1);
                                    }

                                    string[] preno = prepNoList.Split(';');
  
                                        if (preno.Length > 2)
                                        {
                                            prepNoList += ";NOT SPECIFIED"; //"UNKNOWN";
                                        }

                                    PremarationNumber.PrepartionNo.Add(prepNoList);

                                    if (prepNoList.Contains(";"))
                                    {
                                        if (prepNoList.Count(f => (f == ';')) > 2)
                                        {
                                            ClientConfiguration.Color tmpColor = Globals.PDCExcelAddIn.ClientConfiguration[ClientConfiguration.MESSAGE_MULTIPLE_PREPNOS];
                                            if (tmpColor != null)
                                            {
                                                tmpCellRangePreparationNo.Interior.Color = tmpColor.OleColor;
                                            }
                                        }
                                        else if(prepNoList.Count(f => (f == ';')) == 1)
                                        {
                                            tmpCellRangePreparationNo.Value = prepNoList.Split(';')[0]; // single prep number change PDC-926
                                                                                                        
                                        }

                                            try
                                        {
                                            PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Assign oldVisibleValue");
                                            var oldVisibleValue = sheet.Visible;
                                            PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Set sheet.Visible");

                                            sheet.Visible = XlSheetVisibility.xlSheetVisible;
                                            sheet.Activate();

                                            tmpCellRangePreparationNo.Select();
                                            PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Assign oldVisibleValue to sheet.Visible");

                                            sheet.Visible = oldVisibleValue;
                                        }                                        

                                         catch (Exception ex)
                                        {
                                            PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Visible issue caught in catch in main " + ex.Message);
                                        }

                                        //workaround for issue PDC-777
                                        AddValidation(
                                            tmpCellRangePreparationNo,
                                            tmpCis.compoundno,
                                            prepNoList,
                                            compoundNoCellRange.Address.Split(Convert.ToChar("$"))[1]);
                                    }
                                    else if (prepNoList.Length > 0)
                                    {
                                        tmpCellRangePreparationNo.Value[missing] = prepNoList;
                                    }
                                    else if (IsUnknownCompound(tmpCompoundInfos[tmpRow, tmpColumn]))
                                    {
                                      //this changes has been made regarding dummy prep changes
                                      // tmpCellRangePreparationNo.Value[missing] = MSG_UNKNOWN_COMPOUND;
                                    }
                                }
                            }
                        }
                        PDCLogger.TheLogger.LogStoptime("SettingCompoundInfo_Data", "Setting Data");
                    }
                    writeRanges.CompoundnoRange.Areas[i].Value[missing] = tmpCompoundValues;
                    if (ranges.WeightRange != null)
                    {
                        ranges.WeightRange.Value[missing] = tmpWeights;
                    }
                    if (ranges.FormulaRange != null)
                    {
                        ranges.FormulaRange.Value[missing] = tmpFormulas;
                    }

                    PDCLogger.TheLogger.LogStoptime("SettingCompoundInfo", "Setting");
                }
                finally
                {
                    try
                    {
                        Globals.PDCExcelAddIn.Application.EnableEvents = tmpEnabled;
                        PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Assign sheet.Visible in finally");
                        sheet.Visible = oldValue;
                        if (sheet.Visible == XlSheetVisibility.xlSheetVisible)
                            sheet.Activate();

                        temporarySheet?.Delete();
                        PDCLogger.TheLogger.LogStoptime("CompoundInfo", "CompoundInfoAction");
                    }
                   

                     catch (Exception ex)
                    {
                
                        PDCLogger.TheLogger.LogDebugMessage("Excel Issue", "Visible issue caught in catch in finally" + ex.Message);
            
                    }
        }                
            }
        }


        private void AddValidation(Range range, string compoundNo, string values, string compoundNoColumnLetters)
        {
            Debug.Assert(range != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));
            Debug.Assert(!string.IsNullOrEmpty(values));
            Debug.Assert(!string.IsNullOrEmpty(compoundNoColumnLetters));

            var currentActivesheet = Application.ActiveSheet as Worksheet;
            var valueList = values.Split(Convert.ToChar(";"));
            var validationDataSheet = GetValidationDataSheet();
            var row = AddValidationDataToSheet(validationDataSheet, compoundNo, valueList);

            currentActivesheet?.Select();

            AddNamedRange(Application.ActiveWorkbook, validationDataSheet, compoundNo, row, valueList.Length + 1);

            AddValidationToRange(range, compoundNoColumnLetters);
        }

        private Worksheet GetValidationDataSheet()
        {
            const string sheetName = "BAYNOValidationData";
            var validationDataSheet = FindSheet(Application.ActiveWorkbook, sheetName);
            if (validationDataSheet == null)
            {
                validationDataSheet = (Worksheet)Application.Worksheets.Add(After: (Worksheet)Application.Worksheets[Application.Worksheets.Count]);
                validationDataSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                validationDataSheet.Name = sheetName;
            }

            return validationDataSheet;
        }

        private static Worksheet FindSheet(Workbook workbook, string name)
        {
            Debug.Assert(workbook != null);
            Debug.Assert(!string.IsNullOrEmpty(name));

            System.Collections.IEnumerator tmpEnum = workbook.Worksheets.GetEnumerator();
            while (tmpEnum.MoveNext())
            {
                if (tmpEnum.Current is Worksheet tmpSheet && tmpSheet.Name == name)
                {
                    return tmpSheet;
                }
            }
            return null;
        }

        private int FindFirstCompoundNoInValidationDataSheet(Worksheet sheet, string compoundNo)
        {
            Debug.Assert(sheet != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));

            Range tmpRange = sheet.UsedRange;
            int firstUsedRowNumber = tmpRange.Row;
            int usedRowsCount = tmpRange.Rows.Count;

            for (int i = firstUsedRowNumber; i < usedRowsCount + firstUsedRowNumber; i++)
            {
                var cell = (Range) sheet.Cells[i, 1];
                if ((string)cell.Value2 == compoundNo)
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindFirstEmptyRow(Worksheet sheet)
        {
            Debug.Assert(sheet != null);

            Range tmpRange = sheet.UsedRange;
            int firstUsedRowNumber = tmpRange.Row;
            int usedRowsCount = tmpRange.Rows.Count;

            if (firstUsedRowNumber > 1)
            {
                return 1;
            }

            for (int i = firstUsedRowNumber; i < usedRowsCount + firstUsedRowNumber; i++)
            {
                var cell = (Range)sheet.Cells[i, 1];
                if (cell.Value2 == null)
                {
                    return i;
                }
            }

            return firstUsedRowNumber + usedRowsCount;
        }

        private int AddValidationDataToSheet(Worksheet sheet, string compoundNo, string[] values)
        {
            Debug.Assert(sheet != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));
            Debug.Assert(values.Any());

            var row = FindFirstCompoundNoInValidationDataSheet(sheet, compoundNo);
            if (row >= 0)
            {
                ((Range)sheet.Rows[row]).Delete();
            }

            row = FindFirstEmptyRow(sheet);
            WriteValidationDataToRow(sheet, row, compoundNo, values);
            return row;
        }

        private void WriteValidationDataToRow(Worksheet sheet, int rowIndex, string compoundNo, string[] values)
        {
            Debug.Assert(sheet != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));
            Debug.Assert(values.Any());
            Debug.Assert(rowIndex > 0);

            ((Range)sheet.Cells[rowIndex, 1]).Value2 = compoundNo;

            int column = 2;
            foreach (var value in values)
            {
                ((Range)sheet.Cells[rowIndex, column]).Value2 = value;
                column++;
            }
        }

        private void AddNamedRange(Workbook workbook, Worksheet sheet, string compoundNo, int row, int lastColumn)
        {
            Debug.Assert(workbook != null);
            Debug.Assert(sheet != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));
            Debug.Assert(row > 0);
            Debug.Assert(lastColumn > 0);

            string namedRangeName = CreateNamedRangeName(compoundNo);
            var existNamedRange = FindNamedRange(workbook, namedRangeName);
            existNamedRange?.Delete();
            var range = sheet.Range[sheet.Cells[row, 2], sheet.Cells[row, lastColumn]];
            workbook.Names.Add(namedRangeName, range);
        }

        private static string CreateNamedRangeName(string compoundNo)
        {
            Debug.Assert(!string.IsNullOrEmpty(compoundNo));

            return "NR_" + compoundNo.Replace(" ", string.Empty);
        }

        private static Name FindNamedRange(Workbook workbook, string name)
        {
            Debug.Assert(workbook != null);
            Debug.Assert(!string.IsNullOrEmpty(name));

            System.Collections.IEnumerator tmpEnum = workbook.Names.GetEnumerator();
            while (tmpEnum.MoveNext())
            {
                if (tmpEnum.Current is Name tmpName && tmpName.Name == name)
                {
                    return tmpName;
                }
            }

            return null;
        }

        private void AddValidationToRange(Range range, string compoundNoColumnLetters)
        {
            Debug.Assert(range != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNoColumnLetters));

            var formula = CreateFormula(range.Worksheet, compoundNoColumnLetters);

            object oldCompoundNo = null;
            object oldPrepNo = null;

            int col = range.Column;
            int xOffset = Kind == CompoundInfoActionKind.AllSelected ? myUserSettings.HorizontalOffset : 1;
            int yOffset = Kind == CompoundInfoActionKind.AllSelected ? myUserSettings.VerticalOffset : 0;
            Range columnRange = range.Worksheet.Columns[col] as Range;

            if (myUserSettings.VerticalOffset != 0 || myUserSettings.Orientation == UserSettings.Direction.Vertical)
            {
                columnRange = range;
            }
            Debug.Assert(columnRange != null);

            try
            {
                columnRange.Validation.ShowError = false;
                return;
            }
            catch (Exception e)
            {
                //no validation
                if (e is StackOverflowException || e is OutOfMemoryException)
                    throw;
            }

            try
            {
                int startRow = range.Row;
                if (startRow > 1)
                {
                    oldCompoundNo = ((Range)range.Worksheet.Cells[1, col - xOffset]).Value2;
                    oldPrepNo = ((Range)range.Worksheet.Cells[1, col]).Value2;
                    ((Range)range.Worksheet.Cells[1, col - xOffset]).Value2 = ((Range)range.Worksheet.Cells[startRow-yOffset, col - xOffset]).Value2;
                    ((Range)range.Worksheet.Cells[1, col]).Value2 = ((Range)range.Worksheet.Cells[startRow, col]).Value2;
                }

                columnRange.Validation.Delete();

                columnRange.Validation.Add(XlDVType.xlValidateList,
                    XlDVAlertStyle.xlValidAlertInformation,
                    XlFormatConditionOperator.xlBetween,
                    formula,
                    missing);
                columnRange.Validation.ShowError = false;
            }
            finally
            {
                ((Range) range.Worksheet.Cells[1, col - xOffset]).Value2 = null;
                ((Range)range.Worksheet.Cells[1, col]).Value2 = null;
                if (oldCompoundNo != null)
                {
                    ((Range)range.Worksheet.Cells[1, col - xOffset]).Value2 = oldCompoundNo;
                }
                if (oldPrepNo != null)
                {
                    ((Range)range.Worksheet.Cells[1, col]).Value2 = oldPrepNo;
                }  
            }
        }

        private object CreateFormula(Worksheet sheet, string compoundNoColumnLetters)
        {
            Debug.Assert(sheet != null);
            Debug.Assert(!string.IsNullOrEmpty(compoundNoColumnLetters));
            int yOffset = Kind == CompoundInfoActionKind.AllSelected ? myUserSettings.VerticalOffset : 0;
            string rowOffset = "ROW()";
            if (yOffset > 0)
            {
                rowOffset = "(ROW()-" + yOffset +")";
            } else if (yOffset < 0)
            {
                rowOffset = "(ROW()+" + Math.Abs(yOffset)+")";
            }
            Range temp = sheet.Range["a1"];
            dynamic tempValue = temp.Value2;
            string innerFormula =  string.Format("INDIRECT(\"NR_\"&SUBSTITUTE(INDIRECT(\"{0}\"&" + rowOffset +"),\" \",\"\"))", compoundNoColumnLetters);

            temp.Formula = "=" + innerFormula;
            var formula = temp.FormulaLocal;
            temp.Formula = "";
            temp.Value2 = tempValue;
            return formula;
        }

        /// <summary>
        /// Simple check if the compound no is unknown
        /// </summary>
        /// <param name="testStruct"></param>
        /// <returns></returns>
        private bool IsUnknownCompound(TestStruct testStruct)
        {
            return !string.IsNullOrEmpty(testStruct.compoundno_msg);
        }

        #endregion

        #endregion

        #region properties

        #region Kind
        protected CompoundInfoActionKind Kind { get; private set; }

        #endregion

        #endregion
    }
}
