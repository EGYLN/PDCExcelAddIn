using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Globalization;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Microsoft.Office.Interop.Excel;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// This singleton class provides a number of convenience method for common Excel tasks.
    /// </summary>
    class ExcelUtils
    {
        /// <summary>
        /// Shapename used by IVY Chemistry to insert a shape.
        /// </summary>
        private const string IvyChemistryShapename = "IvyChemistry";
        private static readonly ExcelUtils utils = new ExcelUtils();
        protected object Missing = Type.Missing;
        #region methods

        #region AddComment
        /// <summary>
        /// Adds a comment to the specified range
        /// </summary>
        /// <param name="aRange"></param>
        /// <param name="aComment"></param>
        /// <param name="clearComment"></param>
        /// <returns></returns>
        internal Comment AddComment(Range aRange, string aComment, bool clearComment)
        {
            if (clearComment && aRange.Comment != null)
            {
                aRange.Comment.Delete();
            }
            if (aComment == null || aComment.Trim() == "")
            {
                return null;
            }
            Comment tmpComment = aRange.AddComment(aComment);
            try
            {
                tmpComment.Shape.TextFrame.AutoSize = true;
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Autosize comment failed", e);
            }
            return tmpComment;
        }
        #endregion

        #region IsAnyFilterON
        public bool IsAnyFilterOn(Worksheet aWorkSheet)
        {
            if (aWorkSheet.AutoFilter == null)
            {
                return false;
            }
            foreach (Filter filter in aWorkSheet.AutoFilter.Filters)
            {
                if (filter.On)
                {
                    return true;
                }
            }
            return false;
        }
        #endregion

        public ExcelFilterStatus CollectExcelFilters(PDCListObject aPdcListObject)
        {
            ExcelFilterStatus excelFilterStatus =
                new ExcelFilterStatus {AutoFilter = aPdcListObject.Container.AutoFilterMode};
            if (excelFilterStatus.AutoFilter)
            {
                int tmpCurrentFieldNr = 0;
                foreach (Filter tmpFilter in aPdcListObject.Container.AutoFilter.Filters)
                {
                    tmpCurrentFieldNr++;
                    if (tmpFilter.On)
                    {
                        ExcelFilter tmpExcelFilter = new ExcelFilter
                        {
                            Field = tmpCurrentFieldNr,
                            Criteria1 = tmpFilter.Criteria1
                        };
                        if (tmpFilter.Count == 2)
                        {
                            tmpExcelFilter.Criteria2 = tmpFilter.Criteria2;
                        }
                        if (tmpFilter.Count > 1)
                        {
                            tmpExcelFilter.Criteria_operator = tmpFilter.Operator;
                        }
                        excelFilterStatus.ExcelFilters.Add(tmpExcelFilter);
                    }
                }
                // set Empty Filters, if no filter value is to be set
                if (excelFilterStatus.ExcelFilters.Count == 0)
                {
                    ExcelFilter tmpExcelFilterSaver = new ExcelFilter
                    {
                        Field = 1,
                        Criteria_operator = XlAutoFilterOperator.xlAnd
                    };
                    excelFilterStatus.ExcelFilters.Add(tmpExcelFilterSaver);
                }


            }
            return excelFilterStatus;
        }

        public void SetExcelFilters(PDCListObject aPdcListObject, ExcelFilterStatus aExcelFilterStatus)
        {
            if (aExcelFilterStatus == null || !aExcelFilterStatus.AutoFilter) return;
            foreach (ExcelFilter tmpExcelFilter in aExcelFilterStatus.ExcelFilters)
            {
                aPdcListObject.HeaderRange.AutoFilter(tmpExcelFilter.Field, tmpExcelFilter.Criteria1, tmpExcelFilter.Criteria_operator, tmpExcelFilter.Criteria2, true);
            }
        }
        #region CheckSheetExists
        /// <summary>
        /// Checks if the specified workbook already contains a sheet with the given name.
        /// Returns null if no sheet with the name was found, otherwise the sheet type name is returned.
        /// </summary>
        /// <param name="aWorkbook"></param>
        /// <param name="aName"></param>
        /// <returns></returns>
        public string CheckSheetExists(Workbook aWorkbook, string aName)
        {
            System.Collections.IEnumerator tmpEnum = aWorkbook.Worksheets.GetEnumerator();
            while (tmpEnum.MoveNext())
            {
                object tmpSheetCand = tmpEnum.Current;
                Worksheet worksheet = tmpSheetCand as Worksheet;
                if (worksheet != null)
                {
                    Worksheet tmpSheet = worksheet;
                    if (tmpSheet.Name == aName)
                    {
                        return "Worksheet";
                    }
                }
                else if (tmpSheetCand is Chart)
                {
                    Chart tmpChart = (Chart)tmpSheetCand;
                    if (tmpChart.Name == aName)
                    {
                        return "Chart";
                    }
                }
                else if (tmpSheetCand is DialogSheet)
                {
                    DialogSheet tmpDialogSheet = (DialogSheet)tmpSheetCand;
                    if (tmpDialogSheet.Name == aName)
                    {
                        return "Dialogsheet";
                    }
                }
            }
            return null;
        }
        #endregion
        public string GetAddressLocal(Worksheet sheet, int row, int column)
        {
            Range range = (Range)sheet.Cells[row, column];
            return range.get_AddressLocal(Type.Missing, Type.Missing, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
        }
        public Range GetRange(Worksheet aSheet, int aStartRow, int aStartColumn, int anEndRow, int anEndColumn)
        {
            try
            {
                return GetRange(aSheet, (Range)aSheet.Cells[aStartRow, aStartColumn], (Range)aSheet.Cells[anEndRow, anEndColumn]);
            }
            catch (Exception)
            {
                PDCLogger.TheLogger.LogError(PDCLogger.LOG_NAME_COM, string.Format("Failed to get range ({0}, {1}), ({2}, {3})", aStartRow, aStartColumn, anEndRow, anEndColumn));
                throw;
            }
        }
        public Range GetRange(Worksheet aSheet, object aStart, object anEnde)
        {
            return aSheet.get_Range(aStart, anEnde);
        }

        #region CreateNewSheet
        /// <summary>
        /// Creates a new Excel worksheet for the specified testdefinition
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="testdefinition"></param>
        /// <returns></returns>
        public Worksheet CreateNewSheet(Workbook workbook, Lib.Testdefinition testdefinition)
        {
            string sheetName = GenerateSheetName(testdefinition);
            Worksheet sheet = CreateNewSheet(workbook, sheetName, testdefinition);
            return sheet;
        }

        /// <summary>
        /// Creates a new Excel worksheet for use with the specified test definition with the given name.
        /// If the name is already used for an existing worksheet a SheetNameAlreadyExistsException is thrown.
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="name"></param>
        /// <param name="testdefinition"></param>
        /// <returns></returns>
        public Worksheet CreateNewSheet(Workbook workbook, string name, Lib.Testdefinition testdefinition)
        {
            EnsureSheetnameNotUsed(workbook, name);
            Worksheet newSheet = (Worksheet)workbook.Worksheets.Add(Type.Missing, Type.Missing, 1, XlSheetType.xlWorksheet);
            newSheet.Name = name;
            if (testdefinition != null)
            {
                try
                {
                    newSheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_TESTNAME, testdefinition.TestName);
                    newSheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_TESTNO, testdefinition.TestNo);
                    newSheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_TESTVERSION, testdefinition.Version);
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Exception while setting testdefinition custom properties", e);
                }
            }

            return newSheet;
        }
        #endregion

        #region DeleteNames

        /// <summary>
        /// Removes all specified NamedRanges from the given worksheet
        /// </summary>
        /// <param name="sheet">An Excel worksheet</param>
        /// <param name="workbook">Workbook, which contains the sheet</param>
        /// <param name="names">A set of names. All found names are removed from the worksheet.
        /// Names which are not defined on the worksheet are ignored.
        /// </param>
        public void DeleteNames(Worksheet sheet, Workbook workbook, params string[] names)
        {
            Dictionary<string, string> namesToDelete = new Dictionary<string, string>();
            string sheetName = sheet == null ? "" : "'" + sheet.Name + "'!";
            foreach (string rangeName in names)
            {
                namesToDelete.Add(sheetName + rangeName, rangeName);
            }
            if (sheet != null)
            {
                System.Collections.IEnumerator tmpNames = sheet.Names.GetEnumerator();
                DeleteNames(namesToDelete, tmpNames);
            }
            if (workbook != null)
            {
                System.Collections.IEnumerator tmpNames = workbook.Names.GetEnumerator();
                DeleteNames(namesToDelete, tmpNames);
            }
        }

        /// <summary>
        /// Deletes the named ranges
        /// </summary>
        /// <param name="theNamesToDelete">The set of named range identifiers</param>
        /// <param name="theNames">Collection of Excel.Names</param>
        private void DeleteNames(Dictionary<string, string> theNamesToDelete, System.Collections.IEnumerator theNames)
        {
            while (theNames.MoveNext())
            {
                Name name = (Name)theNames.Current;
                if (name != null && theNamesToDelete.ContainsKey(name.Name))
                {
                    name.Delete();
                }
            }
        }
        #endregion

        #region DeleteRows
        /// <summary>
        /// Deletes the specified sheet rows.
        /// </summary>
        /// <param name="theSelectedCells"></param>
        internal void DeleteRows(Range theSelectedCells)
        {
            Range tmpRange = theSelectedCells.EntireRow;
            tmpRange.Delete(Type.Missing);
        }
        #endregion

        #region DeleteShapes
        /// <summary>
        /// Deletes all Shapes on the sheet which start with a constant prefix for ivy structure drawings
        /// </summary>
        /// <param name="sheet"></param>
        public void DeleteIvyShapes(Worksheet sheet)
        {

            System.Collections.IEnumerator shapes = sheet.Shapes.GetEnumerator();
            while (shapes.MoveNext())
            {
                try
                {
                    Shape shape = (Shape)shapes.Current;
                    string shapeName = shape != null ? shape.Name : null;
                    if (shapeName != null && shapeName.StartsWith(IvyChemistryShapename))
                    {
                        shape.Delete();
                    }
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Cannot delete image ", e);
                }
            }
        }
        /// <summary>
        /// Removes the pdc generated shapes from the sheet
        /// </summary>
        /// <param name="sheetInfo"></param>
        public void DeleteShapes(SheetInfo sheetInfo)
        {
            try
            {
                List<string> filenames = new List<string>(sheetInfo.ShapeFileNames);
                Worksheet sheet = sheetInfo.ExcelSheet;
                if (filenames.Count > 0)
                {
                    System.Collections.IEnumerator shapes = sheet.Shapes.GetEnumerator();
                    while (shapes.MoveNext())
                    {
                        Shape shape = (Shape)shapes.Current;
                        if (shape == null)
                        {
                            continue;
                        }
                        string shapeName = shape.Name;
                        if (filenames.Contains(shapeName))
                        {
                            try
                            {
                                shape.Delete();
                                sheetInfo.ShapeFileNames.Remove(shapeName);
                            }
                            catch (Exception e)
                            {
                                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Cannot delete image ", e);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Cannot delete images", e);
            }
        }
        #endregion

        #region EnsureSheetnameNotUsed
        /// <summary>
        /// Throws an exception if the workbook contains a sheet with the specified name.
        /// </summary>
        /// <param name="aWorkbook"></param>
        /// <param name="aName"></param>
        private void EnsureSheetnameNotUsed(Workbook aWorkbook, string aName)
        {
            string tmpExistingSheetType = CheckSheetExists(aWorkbook, aName);
            if (tmpExistingSheetType != null)
            {
                throw new Exceptions.SheetNameAlreadyExistsException(aName, tmpExistingSheetType);
            }
        }
        #endregion

        #region ExitCellEditing
        /// <summary>
        /// Exits the cell editing mode
        /// </summary>
        public void ExitCellEditing()
        {
            SendKeys.SendWait("{TAB}");
        }
        #endregion

        #region ExtractRowNumbers
        /// <summary>
        /// Extracts the row numbers of the selected cells
        /// </summary>
        /// <param name="aRange"></param>
        /// <returns></returns>
        internal Dictionary<int, int> ExtractRowNumbers(Range aRange)
        {
            Dictionary<int, int> tmpRows = new Dictionary<int, int>();
            if (aRange.Count == 0)
            {
                return tmpRows;
            }
            foreach (Range tmpRange in aRange.Areas)
            {
                int tmpRowStart = tmpRange.Row;
                int tmpRowCount = tmpRange.Rows.Count;
                for (int i = tmpRowStart; i < tmpRowStart + tmpRowCount; i++)
                {
                    if (!tmpRows.ContainsKey(i))
                    {
                        tmpRows.Add(i, i);
                    }
                }
            }
            return tmpRows;
        }
        #endregion

        #region GenerateSheetName
        /// <summary>
        /// Generates the PDC sheet name for the specified testdefinition
        /// </summary>
        private static string GenerateSheetName(Lib.Testdefinition testdefinition)
        {
            string sheetName = testdefinition.TestName;
            if (sheetName.Length > 20)
            {
                sheetName = sheetName.Substring(0, 20);
            }

            sheetName = GetSheetNameWithoutSpecialCharacter(sheetName);

            sheetName += "(" + testdefinition.TestNo + "_" + testdefinition.Version + ")";
            return sheetName;
        }
        #endregion

        #region GetExcelNumberSeparators
        /// <summary>
        /// Returns the current settings for Decimal and Groupseparator saved within NumberFormatInfo
        /// </summary>
        /// <returns>NumberFormatInfo with current settings </returns>
        internal NumberFormatInfo GetExcelNumberSeparators()
        {
            NumberFormatInfo nfi = new NumberFormatInfo();
            if (Globals.PDCExcelAddIn.Application.UseSystemSeparators)
            {

                nfi.NumberDecimalSeparator = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator;
                nfi.NumberGroupSeparator = NumberFormatInfo.CurrentInfo.NumberGroupSeparator;
            }
            else
            {
                nfi.NumberDecimalSeparator = Globals.PDCExcelAddIn.Application.DecimalSeparator;
                nfi.NumberGroupSeparator = Globals.PDCExcelAddIn.Application.ThousandsSeparator;
            }
            return nfi;
        }
        #endregion

        #region GetKey
        /// <summary>
        /// Returns the guid for the sheet from the custom property or null
        /// if the associated custom property is not set.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public object GetKey(Worksheet sheet)
        {
            return GetKey(sheet, false);
        }

        /// <summary>
        /// Returns the guid for the sheet from the custom property.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="createKey">If set to true a guid will be created
        /// if it cannot be found</param>
        /// <returns></returns>
        public object GetKey(Worksheet sheet, bool createKey)
        {
            System.Collections.IEnumerator properties = sheet.CustomProperties.GetEnumerator();
            while (properties.MoveNext())
            {
                CustomProperty property = (CustomProperty)properties.Current;
                if (property != null && property.Name == PDCExcelConstants.PROPERTY_GUID)
                {
                    return property.Value;
                }
            }
            if (createKey)
            {
                string newGUID = Guid.NewGuid().ToString();
                sheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_GUID, newGUID);
                return newGUID;
            }

            return null;
        }
        #endregion

        #region GetSheetNameWithoutSpecialCharacter
        /// <summary>
        ///   Replaces all the characters '\', '/', '?', '*', '[' and ']' by a blnk.
        /// </summary>
        /// <param name="sheetName">
        ///   The actual name of the sheet.
        /// </param>
        /// <returns>
        ///   The new name of the sheet.
        /// </returns>
        internal static string GetSheetNameWithoutSpecialCharacter(string sheetName)
        {
            string newName = sheetName.Replace("\\", " ");
            newName = newName.Replace("/", " ");
            newName = newName.Replace("?", " ");
            newName = newName.Replace("*", " ");
            newName = newName.Replace("[", " ");
            newName = newName.Replace("]", " ");

            return newName.Trim();
        }
        #endregion

        #region IsEmptySheet
        /// <summary>
        /// Checks if the specified worksheet can be considered as empty.
        /// </summary>
        /// <param name="aSheet">Excel worksheet (must not be null)</param>
        /// <returns>true, if the specified sheet can be considered empty</returns>
        public bool IsEmptySheet(Worksheet aSheet)
        {
            Range tmpRange = aSheet.UsedRange;
            int tmpRowStart = tmpRange.Row;
            int tmpRowCount = tmpRange.Rows.Count;
            int tmpColumnStart = tmpRange.Column;
            int tmpColumnCount = tmpRange.Columns.Count;
            return tmpRowStart == 1 && tmpRowCount == 1 && tmpColumnStart == 1 && tmpColumnCount == 1;
        }
        #endregion
        #region IsEmptyRange
        public bool IsEmptyRange(Range aRange)
        {
            if (aRange == null)
            {
                return true;
            }
            object[,] tmpValue = RangeToMatrix(aRange);
            if (tmpValue == null)
            {
                return true;
            }
            for (int i = tmpValue.GetLowerBound(0); i <= tmpValue.GetUpperBound(0); i++)
                for (int j = tmpValue.GetLowerBound(1); j <= tmpValue.GetUpperBound(1); j++)
                    if (tmpValue[i, j] != null && !string.IsNullOrEmpty(""+tmpValue[i, j]))
                        return false;
            return true;
        }
        #endregion

        #region IsInCellEditingMode
        /// <summary>
        /// Checks if Excel is currently in cell editing mode. In this case it is not save
        /// to change the sheet.
        /// </summary>
        /// <returns></returns>
        public bool IsInCellEditingMode()
        {
            bool tmpChanged = false;
            if (Globals.PDCExcelAddIn.Application.Interactive == false)
            {
                return false;
            }
            try
            {
                Globals.PDCExcelAddIn.Application.Interactive = false;
                tmpChanged = true;
                return false;
            }
#pragma warning disable 0168
            catch (Exception e)
            {
                return true;
            }
#pragma warning restore 0168
            finally
            {
                try
                {
                    Globals.PDCExcelAddIn.Application.Interactive = true;
                }
                catch (Exception e)
                {
                    if (tmpChanged)
                    {
                        PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Failed to switch on interactive mode", e);
                    }
                }
            }
        }
        #endregion

        #region IsSheetReferenceValid
        /// <summary>
        /// Checks if the specified worksheet reference is still valid.
        /// </summary>
        /// <param name="aSheet"></param>
        /// <returns></returns>
        public bool IsSheetReferenceValid(Worksheet aSheet)
        {
            try
            {
                object unused = aSheet.Parent;
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

        #region MoreThanOneVisibleSheet
        /// <summary>
        /// The last visible worksheet cannot be removed from a workbook.
        /// This method checks if there is at least one removable worksheet.
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        internal bool MoreThanOneVisibleSheet(Workbook workbook)
        {
            int count = 0;
            System.Collections.IEnumerator sheets = workbook.Worksheets.GetEnumerator();
            while (sheets.MoveNext())
            {
                Worksheet current = sheets.Current as Worksheet;
                if (current != null)
                {
                    Worksheet sheet = current;
                    if (count == 1)
                    {
                        return true;
                    }
                    XlSheetVisibility visible = sheet.Visible;
                    if (visible == XlSheetVisibility.xlSheetVisible)
                    {
                        count++;
                    }
                }
            }
            return false;
        }
        #endregion

        #region NeutralizeSheet
        /// <summary>
        /// Clears the worksheet so that it is in a neutral state.
        /// </summary>
        /// <param name="excelSheet"></param>
        internal void NeutralizeSheet(Worksheet excelSheet)
        {
            bool tmpEventsenabled = Globals.PDCExcelAddIn.Application.EnableEvents;
            Globals.PDCExcelAddIn.EventsEnabled = false;
            Globals.PDCExcelAddIn.Application.EnableEvents = false;
            try
            {
                Range tmpUsedRange = excelSheet.UsedRange;
                tmpUsedRange.Delete(Type.Missing);
                excelSheet.Name = "_";
            }
            finally
            {
                Globals.PDCExcelAddIn.EventsEnabled = true;
                Globals.PDCExcelAddIn.Application.EnableEvents = tmpEventsenabled;
            }
        }
        #endregion

        #region ProtectSheet
        /// <summary>
        /// Protects the specified sheet. Cell Formatting and hyperlinks are still allowed. Optionally
        /// a range of edit cells can be specified.
        /// </summary>
        /// <param name="aSheet">The sheet to protect</param>
        /// <param name="aPassword">The password used to protect the sheet</param>
        /// <param name="anEditRange">Optional range of editable cells</param>
        /// <param name="aRangeName">Range name for editable cells.</param>
        internal void ProtectSheet(Worksheet aSheet, string aPassword, Range anEditRange, string aRangeName)
        {
            if (anEditRange != null)
            {
                AllowEditRange tmpEditRange = null;

                try
                {//fast check if the desired edit range already exists, avoids use of enumeration
                    tmpEditRange = aSheet.Protection.AllowEditRanges.get_Item(aRangeName);
                }
#pragma warning disable 0168
                catch (Exception e)
                {
                    //ok, allowEditRange does not exist -> create a new
                }
#pragma warning restore 0168
                if (tmpEditRange == null)
                {
                    aSheet.Protection.AllowEditRanges.Add(aRangeName, anEditRange);
                }
                else
                {
                    tmpEditRange.Range = anEditRange;
                }
            }
            aSheet.Protect(aPassword, false, true, false, false, true, true, true, false, false, true, false, false, false, false, false);
        }
        #endregion

        #region RenameSheetForTestdefinition
        /// <summary>
        /// Renames the specified worksheet for the given testdefinition
        /// </summary>
        public void RenameSheetForTestdefinition(Workbook workbook, Worksheet sheet, Lib.Testdefinition testdefinition)
        {
            string currentName = sheet.Name;
            string newName = GenerateSheetName(testdefinition);
            if (currentName == newName)
            {
                return;
            }
            EnsureSheetnameNotUsed(workbook, newName);
            sheet.Name = newName;
        }
        #endregion

        #region SameCells
        /// <summary>
        /// Returns true if both objects are Excel.Range objects describing the same area, false otherwise
        /// </summary>
        /// <param name="aCellRange1">A candidate for a cell range</param>
        /// <param name="aCellRange2">Another candidate for a cell range</param>
        /// <returns></returns>
        internal bool SameCells(object aCellRange1, object aCellRange2)
        {
            if (!(aCellRange1 is Range))
            {
                return false;
            }
            if (!(aCellRange2 is Range))
            {
                return false;
            }
            Range tmpRange1 = (Range)aCellRange1;
            Range tmpRange2 = (Range)aCellRange2;
            int tmpCount1 = tmpRange1.Count;
            int tmpCount2 = tmpRange2.Count;
            if (tmpCount1 == 0 || tmpCount1 != tmpCount2)
            {
                return false;
            }
            Range tmpIntersect = Globals.PDCExcelAddIn.Intersect(tmpRange1, tmpRange2);
            return tmpCount1 == tmpIntersect.Count;
        }
        #endregion

        #region SearchSheet
        /// <summary>
        /// Tries to find the sheet with the specified PDC key. Returns null if
        /// no associated sheet were found.
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public Worksheet SearchSheet(object guid)
        {
            System.Collections.IEnumerator allSheets = Globals.PDCExcelAddIn.Application.Worksheets.GetEnumerator();
            while (allSheets.MoveNext())
            {
                object currentSheet = allSheets.Current;
                Worksheet cand = currentSheet as Worksheet;
                if (cand != null)
                {
                    Worksheet sheet = cand;
                    if (guid == GetKey(sheet))
                    {
                        return sheet;
                    }
                }
            }
            return null;
        }
        #endregion

        #region SelectedRows
        /// <summary>
        /// Extracts the distinct row numbers of the current selection
        /// </summary>
        /// <param name="aRange">The range of selected cells</param>
        /// <returns>A set of selected row number</returns>
        public Dictionary<int, int> SelectedRows(Range aRange)
        {
            Dictionary<int, int> tmpSelectedRows = new Dictionary<int, int>();
            foreach (Range tmpArea in aRange.Areas)
            {
                int tmpStartRow = tmpArea.Row;
                int tmpEndRow = tmpStartRow + tmpArea.Rows.Count;
                for (int i = tmpStartRow; i < tmpEndRow; i++)
                {
                    tmpSelectedRows.Add(i, i);
                }
            }
            return tmpSelectedRows;
        }
        #endregion

        #region UnprotectSheet
        /// <summary>
        /// Removes the protection from the sheet with the specified password
        /// </summary>
        /// <param name="aSheet">The protected sheet</param>
        /// <param name="aPassword">The protection password</param>
        internal void UnprotectSheet(Worksheet aSheet, string aPassword)
        {
            if (aSheet.ProtectContents)
            {
                aSheet.Unprotect(aPassword);
            }
        }
        #endregion

        #region ClearValidationCommentsAndColors
        /// <summary>
        /// Clears the validation comments and colors
        /// </summary>
        internal void ClearValidationAndDefaults(params Range[] ranges)
        {
            foreach (Range range in ranges)
            {
                if (range == null)
                {
                    continue;
                }
                try
                {
                    range.ClearComments();
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "ClearValidation-Comments", e);
                }
                try
                {
                    range.Interior.ColorIndex = XlColorIndex.xlColorIndexNone;
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "ClearValidation-Color", e);
                }
                try
                {
                    range.Font.Italic = false;
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "ClearValidation-Font", e);
                }
            }
        }
        #endregion
        #region RangeToMatrix

        /// <summary>
        /// Returns an object matrix for the given range. Returns null, if the range is null.
        /// (A) If the get_Value() method does not return a matrix by itself, it is assumed, that the range is a single table cell.
        /// </summary>
        /// <param name="range">Input range, may be null</param>
        /// <returns></returns>
        internal object[,] RangeToMatrix(Range range)
        {
            if (range == null)
            {
                return null;
            }
            object rangeValue = range.get_Value(Missing);
            object[,] values = rangeValue as object[,];
            if (values != null)
            {
                object[,] fromExcel = values;
                //Copies the values into a zero-offset matrix to make live easier
                object[,] copy = new object[fromExcel.GetLength(0), fromExcel.GetLength(1)];
                for (int i = 0; i < fromExcel.GetLength(0); i++)
                {
                    for (int j = 0; j < fromExcel.GetLength(1); j++)
                    {
                        copy[i, j] = fromExcel[i + 1, j + 1];
                    }
                }
                return copy;
            }

            if (rangeValue == null) { return new object[,] {{null}}; }
            if (rangeValue is string) { return new object[,] { { ((string)rangeValue).Trim() } }; }                
            if (rangeValue is double) { return new object[,] {{((double) rangeValue)}};}

            return null;
        }

        #endregion
        #region ResetRowHeights
        /// <summary>
        /// Resets the row heights of the range to the standard size of the containing worksheet
        /// </summary>
        /// <param name="range">The Datarange</param>
        public void ResetRowHeights(Range range)
        {
            object parent = range.Parent;
            Worksheet worksheet = parent as Worksheet;
            if (worksheet != null)
            {
                Worksheet sheet = worksheet;
                range.RowHeight = sheet.StandardHeight;
            }
        }
        #endregion

        #endregion

        #region properties

        #region ActiveSheet
        /// <summary>
        /// Returns the active sheet if it is a worksheet and null otherwise
        /// </summary>
        public Worksheet ActiveSheet
        {
            get
            {
                object activeSheet = Globals.PDCExcelAddIn.Application.ActiveSheet;
                Worksheet worksheet = activeSheet as Worksheet;
                return worksheet;
            }
        }
        #endregion

        #region ActiveSheetInfo
        /// <summary>
        /// Returns the sheet info for the active sheet or null if no sheet info exists.
        /// </summary>
        public SheetInfo ActiveSheetInfo
        {
            get
            {
                Worksheet tmpSheet = ActiveSheet;
                if (tmpSheet == null)
                {
                    return null;
                }
                return Globals.PDCExcelAddIn.GetSheetInfo(tmpSheet);
            }
        }
        #endregion

        #region TheUtils
        /// <summary>
        /// Returns the ExcelUtils singleton instance.
        /// </summary>
        public static ExcelUtils TheUtils
        {
            get
            {
                return utils;
            }
        }
        #endregion

        #endregion
    }
}
