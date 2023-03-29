using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /// <summary>
    /// PredefinedParameterHandler for the predefined parameter project.
    /// </summary>
    public class ProjectParameterHandler:PredefinedParameterHandler
    {
        private string name;
        private Excel.Worksheet projectSheet;
        public ProjectParameterHandler(Excel.Worksheet aSheet, string aName)
        {
            name = aName;
            Excel.Workbook tmpWB = (Excel.Workbook)aSheet.Parent;
            projectSheet = (Excel.Worksheet)tmpWB.Sheets.Add(missing, missing, 1, Excel.XlSheetType.xlWorksheet);
            projectSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden;
            projectSheet.Name = name;
            List<string> tmpProjectNames = Globals.PDCExcelAddIn.PdcService.GetProjectNames();
            string[,] tmpAllowed = new string[tmpProjectNames.Count, 1];
            initializeProjectList(aSheet, tmpProjectNames, tmpAllowed);
        }

        /// <summary>
        /// Updates the Project column. Replaces the current project list with a new list from the server.
        /// </summary>
        /// <param name="aPDCList"></param>
        /// <param name="aCurrentColumn"></param>
        /// <param name="anUpdateColumn"></param>
        public override void UpdateColumn(PDCListObject aPDCList, ListColumn aCurrentColumn, ListColumn anUpdateColumn)
        {
            Excel.Range tmpListRange = projectSheet.get_Range(name + "Values", missing);
            List<string> tmpProjectNames = Globals.PDCExcelAddIn.PdcService.GetProjectNames();
            int tmpSize = Math.Max(tmpListRange.Rows.Count, tmpProjectNames.Count);
            string[,] tmpAllowed = new string[tmpSize, 1];
            initializeProjectList(aPDCList.Container, tmpProjectNames, tmpAllowed);
        }

        private void initializeProjectList(Excel.Worksheet aSheet, List<string> tmpProjectNames, string[,] tmpAllowed)
        {
            for (int i = 0; i < tmpProjectNames.Count; i++)
            {
                tmpAllowed[i, 0] = tmpProjectNames[i];
            }
            Excel.Range tmpListRange = projectSheet.get_Range(
                (Excel.Range)projectSheet.Cells[1, 1],
            ((Excel.Range)projectSheet.Cells[tmpAllowed.Length, 1]));
            tmpListRange.Formula = tmpAllowed;
            Excel.Name tmpName = aSheet.Names.Add(name + "Values", tmpListRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
            projectSheet.Names.Add(name + "Values", tmpListRange, true, missing, missing, missing, missing, missing, missing, missing, missing);
        }

        public override void InitializeNewCells(Microsoft.Office.Interop.Excel.Range aRange, PDCListObject aPDCList)
        {
            Excel.Worksheet tmpSheet = (Excel.Worksheet)aRange.Parent;
            aRange.Validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, "="+name+"Values",missing);
        }
        public override void Delete(ListColumn aColumn, PDCListObject aPDCList, bool completeList)
        {
            Globals.PDCExcelAddIn.Application.DisplayAlerts = false;
            ExcelUtils.TheUtils.DeleteNames(aPDCList.Container, null, name + "Values");
            projectSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            projectSheet.Delete();
            Globals.PDCExcelAddIn.Application.DisplayAlerts = true;
        }
    }
}
