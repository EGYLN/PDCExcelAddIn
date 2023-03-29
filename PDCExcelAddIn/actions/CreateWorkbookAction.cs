using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Properties;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Lets the user browse for a test definition and 
    /// creates a new PDC worksheet based on the selected test definition 
    /// or initializes the active worksheet for the selected test definition.
    /// </summary>
    class CreateWorkbookAction : PDCAction
    {

        public const string ACTION_TAG = "PDC_CreateSheetAction";
        private const int HEADER_X = 3;
        private const int HEADER_Y = 1;

        private bool myMeasurementHandlerCreated;
        private bool mySheetCreationFailed;


        #region constructor
        public CreateWorkbookAction(bool beginGroup)
            : base(Properties.Resources.Action_CreateWorkbook_Caption, ACTION_TAG, beginGroup)
        {
        }
        #endregion

        #region methods

        #region AddTestVariable
        private void AddTestVariable(Lib.Testdefinition aTestdefinition, Excel.Worksheet aSheet, List<ListColumn> aColumnList, Lib.TestVariable aVariable, bool aInitParamHandlerFlag)
        {
            ListColumn tmpVarColumn;

            if (aVariable.IsMeasurementLevel)
            {
                if (!myMeasurementHandlerCreated)
                {
                    tmpVarColumn = PDCListObject.CreateMeasurementColumn(aSheet, aTestdefinition);
                    myMeasurementHandlerCreated = true;
                    aColumnList.Add(tmpVarColumn);
                }
                return;
            }
            if (aVariable.IsExperimentLevelReferenceForSMT) return;


            // Only for non measurement variables
            tmpVarColumn = PDCListObject.CreateColumn(aVariable);
            if (aInitParamHandlerFlag)
            {
                aColumnList.Add(tmpVarColumn);
            }
            else
            {
                aColumnList.Add(tmpVarColumn);
            }
            if (aVariable.IsBinaryParameter())
            {
                tmpVarColumn.IsHyperLink = true;
            }
        }
        #endregion

        #region CreateClassGroupings
        /// <summary>
        /// Groups test variable columns with the same variable class together
        /// </summary>
        /// <param name="aList"></param>
        private void CreateClassGroupings(PDCListObject aList)
        {
            if (!Properties.Settings.Default.useGroups)
            {
                return;
            }
            int tmpGroupStart = -1;
            int tmpGroupEnd = -1;
            string tmpClass = "";
            for (int i = 0; i < aList.ColumnCount; i++)
            {
                ListColumn tmpColumn = aList.ListColumn(i);
                if (tmpColumn.TestVariable == null)
                {
                    tmpGroupStart = tmpGroupEnd = -1;
                    continue;
                }
                if (tmpGroupStart == -1)
                {
                    tmpGroupStart = i;
                    tmpClass = tmpColumn.TestVariable.VariableClass;
                    continue;
                }
                if (tmpClass == tmpColumn.TestVariable.VariableClass)
                {
                    tmpGroupEnd = i;
                    continue;
                }
                aList.GroupColumns(tmpGroupStart, tmpGroupEnd);
                tmpGroupStart = tmpColumn.TestVariable == null ? -1 : i;
                tmpGroupEnd = tmpGroupStart;
                tmpClass = tmpColumn.TestVariable == null ? "" : tmpColumn.TestVariable.VariableClass;
            }
            if (tmpGroupStart != -1)
            {
                aList.GroupColumns(tmpGroupStart, aList.ColumnCount - 1);
            }
        }
        #endregion

        #region CreateWorkbook
        /// <summary>
        /// Creates a new Excel Workbook and activates it
        /// </summary>
        /// <param name="aTitle">The title of the workbook</param>
        /// <returns></returns>
        private Excel.Workbook CreateWorkbook(string aTitle)
        {
            int tmpDefaultSheets = Application.SheetsInNewWorkbook;
            try
            {
                Application.SheetsInNewWorkbook = 1;
                Excel.Workbook tmpWb = Application.Workbooks.Add(missing);
                tmpWb.Title = aTitle;
                tmpWb.Activate();
                return tmpWb;
            }
            finally
            {
                Application.SheetsInNewWorkbook = tmpDefaultSheets;
            }
        }
        #endregion

        #region InitializeWorkbook
        /// <summary>
        /// Initializes the active or a new workbook with the selected Testdefinition depending on the
        /// user choice and current Excel state.
        /// </summary>
        private void InitializeWorkbook(Lib.Testdefinition aTestdefinition, bool noMeasurements)
        {
            if (aTestdefinition == null)
            {
                return;
            }
            try
            {
                Application.ScreenUpdating = false;
                Application.EnableEvents = false;
                Excel.Worksheet tmpSheet;
                Excel.Workbook tmpWb = Application.ActiveWorkbook;
                if (tmpWb == null)
                {
                    tmpWb = CreateWorkbook(aTestdefinition.TestName);


                    tmpSheet = (Excel.Worksheet)Application.ActiveSheet;
                    if (tmpSheet != null)
                    {
                        ExcelUtils.TheUtils.RenameSheetForTestdefinition(tmpWb, tmpSheet, aTestdefinition);
                    }
                }
                else
                {
                    tmpSheet = ExcelUtils.TheUtils.CreateNewSheet(tmpWb, aTestdefinition);
                }
                ReinitializeSheet(aTestdefinition, tmpSheet, noMeasurements);
            }
            catch (Exception e)
            {
                mySheetCreationFailed = true;
                ExceptionHandler.TheExceptionHandler.handleException(e, null);
            }
            finally
            {
                myMeasurementHandlerCreated = false;
                Globals.PDCExcelAddIn.EnableExcel();
                Application.Calculate();
            }
        }

        #endregion

        #region PerformAction
        /// <summary>
        /// Displays the test definition browser. 
        /// </summary>
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            mySheetCreationFailed = false;
            BrowseTestDefinitionForm tmpBrowser = new BrowseTestDefinitionForm();
            tmpBrowser.callback += InitializeWorkbook;
            Lib.Testdefinition tmpTestdefinition = tmpBrowser.BrowseTestDefinitions();
            if (tmpTestdefinition == null || mySheetCreationFailed)
            {
                return new ActionStatus();
            }
            //createWorkbook(tmpTestdefinition, tmpBrowser.CreateNewSheetSelected);
            Excel.Worksheet tmpSheet = ExcelUtils.TheUtils.ActiveSheet;
            SheetInfo tmpSheetInfo = GetSheetInfo(tmpSheet);
            if (tmpSheetInfo == null)
            {
                AddIn.RegisterTestdefinition(tmpTestdefinition, tmpSheet);
            }
            else
            {
                tmpSheetInfo.ExcelSheet = tmpSheet;
            }

            Excel.Workbook tmpWb = (Excel.Workbook)tmpSheet.Parent;
            if (!Globals.PDCExcelAddIn.WorkbookMap.ContainsKey(tmpWb))
            {
                Globals.PDCExcelAddIn.WorkbookMap.Add(tmpWb, tmpWb);
                tmpWb.SheetChange += Globals.PDCExcelAddIn.SheetChanged;
                AddIn.InitVBACode(tmpWb);
            }

            tmpSheet.EnableCalculation = true;
            Application.EnableEvents = true;
            tmpSheet.Activate();
            return new ActionStatus();
        }
        #endregion

        #region ReinitializeSheet
        /// <summary>
        /// Initializes the given sheet for usage with the specified test definition
        /// </summary>
        private void ReinitializeSheet(Lib.Testdefinition testdefinition, Excel.Worksheet sheet, bool noMeasurements)
        {

            Excel.Worksheet tmpSheet = sheet;
            Lib.Testdefinition tmpTestdefinition = testdefinition;
            tmpSheet.EnableCalculation = false;
            ExcelUtils.TheUtils.GetKey(tmpSheet, true);
            SetHeaderInfo(testdefinition, sheet);
            //Dis-Disabled confidential note for now, since the discussion turned.
            //May be we have to removed it finally.
            AddConfidentialNote(sheet);
            Lib.ClientConfiguration tmpScheme = Globals.PDCExcelAddIn.ClientConfiguration;
            PDCListObject tmpPDCList = new PDCListObject("Test_" + tmpTestdefinition.TestNo + "_" + tmpTestdefinition.Version, tmpSheet, 4, 1, tmpTestdefinition);

            List<ListColumn> tmpColumns = new List<ListColumn>();
            tmpColumns.Add(new ListColumn(PDCExcelConstants.COMPOUNDNO, "Compound No", Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
            tmpColumns.Add(new ListColumn(PDCExcelConstants.PREPARATIONNO, "Preparation No", Lib.PDCConstants.C_ID_PREPARATIONNO,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
            tmpColumns.Add(new ListColumn(PDCExcelConstants.MCNO, "MC No", Lib.PDCConstants.C_ID_MCNO,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor));
            tmpColumns.Add(new ListColumn(PDCExcelConstants.STRUCTURE_DRAWING, "Structure Drawing", Lib.PDCConstants.C_ID_STRUCTURE_DRAWING,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor, true));
            tmpColumns.Add(new ListColumn(PDCExcelConstants.MOLECULAR_WEIGHT, "Molecular Weight", Lib.PDCConstants.C_ID_WEIGHT,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor, true));
            tmpColumns.Add(new ListColumn(PDCExcelConstants.FORMULA, "Formula", Lib.PDCConstants.C_ID_FORMULA,
              tmpScheme[Lib.ClientConfiguration.HEADER_COMPOUND_INFO].SystemColor, true));

            tmpColumns.Add(new ListColumn(PDCExcelConstants.RESULTSTATUS, "Result Status", Lib.PDCConstants.C_ID_RESULT_STATUS,
              tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, true));
            ListColumn tmpExperimentNoColumn = new ListColumn(
                PDCExcelConstants.EXPERIMENT_NO, 
                "Experiment No",
                Lib.PDCConstants.C_ID_EXPERIMENTNO,
                tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, 
                true)
            {
                Hidden = true
            };
            tmpColumns.Add(tmpExperimentNoColumn);



            tmpColumns.Add(new ListColumn(PDCExcelConstants.UPLOAD_ID, "Upload Id", Lib.PDCConstants.C_ID_UPLOAD_ID,
              tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, true));

            ListColumn tmpUploadDateColumn = new ListColumn(
                PDCExcelConstants.UPLOADDATE, 
                "Upload Date",
                Lib.PDCConstants.C_ID_UPLOADDATE,
                tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, 
                false)
            {
                Comment = PDCExcelConstants.COMMENT_SEARCH_HELP_FOR_DATE
            };
            tmpColumns.Add(tmpUploadDateColumn);

            ListColumn tmpDateResult = new ListColumn(
                PDCExcelConstants.DATERESULT,
                Properties.Resources.COL_HEADER_DATE_RESULT, 
                Lib.PDCConstants.C_ID_DATE_RESULT,
                tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, 
                false)
            {
                Comment = PDCExcelConstants.COMMENT_SEARCH_HELP_FOR_DATE
            };
            tmpColumns.Add(tmpDateResult);

            tmpColumns.Add(new ListColumn(PDCExcelConstants.PERSONID, "Uploader Name", Lib.PDCConstants.C_ID_PERSONID,
               tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, true));

            tmpColumns.Add(new ListColumn(PDCExcelConstants.REPORT_TO_PIX, "Report to PIx", Lib.PDCConstants.C_ID_PDC_ONLY_DATA,
               tmpScheme[Lib.ClientConfiguration.HEADER_PDC_INFO].SystemColor, false));

            foreach (Lib.TestVariable tmpVar in tmpTestdefinition.Variables)
            {
                // if no measurements data should be read (checkbox on testselection form) and current variable is a measurement variable:
                // do NOT add variables to TestVariables
                if (noMeasurements && tmpVar.IsMeasurementLevel)
                {
                    continue;
                }

                AddTestVariable(tmpTestdefinition, tmpSheet, tmpColumns, tmpVar, true);
            }
            SheetInfo tmpInfo = AddIn.RegisterTestdefinition(testdefinition, sheet);

            tmpInfo.MainTable = tmpPDCList;

            tmpPDCList.AddColumns(tmpColumns);
            CreateClassGroupings(tmpPDCList);
            tmpTestdefinition.Tag = tmpPDCList;
            tmpSheet.Activate();
            Globals.PDCExcelAddIn.EnableExcel();
        }

        private void AddConfidentialNote(Excel.Worksheet sheet)
        {
            try
            {
                Uri tmpUri = new Uri(GetType().Assembly.CodeBase);

                string tmpCodeBase = Path.GetDirectoryName(tmpUri.LocalPath);
                string tmpFileName = Path.Combine(tmpCodeBase, "Resources/confidential.png");
                if (!File.Exists(tmpFileName))
                {
                    Lib.Util.PDCLogger.TheLogger.LogError(Lib.Util.PDCLogger.LOG_NAME_EXCEL, "Could not find confidential note picture.");
                    return;
                }

                Excel.Range tmpCellRange = ExcelUtils.TheUtils.GetRange(sheet, HEADER_Y, HEADER_X + 4, HEADER_Y + 2, HEADER_X + 4 + 3);
                double tmpLeft = (double) tmpCellRange.Left;
                double tmpTop = (double) tmpCellRange.Top;
                double tmpWidth = (double) tmpCellRange.Width;
                double tmpHeight = (double) tmpCellRange.Height;

                Excel.Shape shape = sheet.Shapes.AddPicture(tmpFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoCTrue, (float) tmpLeft, (float) tmpTop,
                    (float) tmpWidth, (float) tmpHeight);
                shape.Locked = true;
                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
            }
            catch (Exception e)
            {
                Lib.Util.PDCLogger.TheLogger.LogException(Lib.Util.PDCLogger.LOG_NAME_EXCEL, "Could not add confidential note picture", e);
                MessageBox.Show(Resources.MSG_FAILED_TO_ADD_SECRET_PIC_TEXT, Resources.MSG_FAILED_TO_ADD_SECRET_PIC_TITLE);
            }

        }

        #endregion

        #region SetHeaderInfo
        /// <summary>
        /// Displays the Testno, Version and Testname on the sheet
        /// </summary>
        /// <param name="testdefinition">A test definition</param>
        /// <param name="sheet">The main worksheet of the test definition</param>
        private void SetHeaderInfo(Lib.Testdefinition testdefinition, Excel.Worksheet sheet)
        {
            try
            {
                Excel.Range tmpRange = (Excel.Range)sheet.Cells[HEADER_Y, HEADER_X];
                tmpRange.Value2 = Properties.Resources.Header_Testno;
                tmpRange.Font.Bold = true;
                tmpRange = (Excel.Range)sheet.Cells[HEADER_Y + 1, HEADER_X];
                tmpRange.Value2 = testdefinition.TestNo;
                sheet.Names.Add(PDCExcelConstants.RANGE_TESTNO, tmpRange, true, missing, missing, missing, missing, missing, missing, missing, missing);

                tmpRange = (Excel.Range)sheet.Cells[HEADER_Y, HEADER_X + 1];
                tmpRange.Value2 = Properties.Resources.Header_Version;
                tmpRange.Font.Bold = true;
                tmpRange = (Excel.Range)sheet.Cells[HEADER_Y + 1, HEADER_X + 1];
                tmpRange.Value2 = testdefinition.Version;
                sheet.Names.Add(PDCExcelConstants.RANGE_VERSION, tmpRange, true, missing, missing, missing, missing, missing, missing, missing, missing);

                tmpRange = (Excel.Range)sheet.Cells[HEADER_Y, HEADER_X + 2];
                tmpRange.Value2 = Properties.Resources.Header_Testname;
                tmpRange.Font.Bold = true;
                tmpRange = (Excel.Range)sheet.Cells[HEADER_Y + 1, HEADER_X + 2];
                tmpRange.Value2 = testdefinition.TestName;
                sheet.Names.Add(PDCExcelConstants.RANGE_TESTNAME, tmpRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
                object tempHeaderStart = sheet.Cells[HEADER_Y, HEADER_X];
                object tempHeaderEnde = sheet.Cells[HEADER_Y + 1, HEADER_X + 2];
                tmpRange = sheet.Range[tempHeaderStart, tempHeaderEnde];
                tmpRange.Interior.ColorIndex = 0;
                tmpRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, missing);
            }
            catch (Exception e)
            {
                Lib.Util.PDCLogger.TheLogger.LogException(Lib.Util.PDCLogger.LOG_NAME_EXCEL, "Setting HeaderInformation", e);
            }
        }
        #endregion

        #endregion

        #region properties

        #endregion
    }
}
