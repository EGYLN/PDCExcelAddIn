using System;

using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    using System.Windows.Forms;

    /// <summary>
    /// PDC Action for client-side formatting of compoundnos.
    /// </summary>
    class FormattingCompoundInfoAction : CompoundInfoAction
    {
        public FormattingCompoundInfoAction(
            Office.CommandBarPopup aPopup, 
            bool beginGroup, 
            string aCaption, 
            string aTag,
            UserSettings userSettings ) :
            base(aPopup, beginGroup, aCaption, aTag,userSettings)

        {
        }

        /// <summary>
        /// Adds leading zeros if the compoundno string until the specified length is reached
        /// </summary>
        /// <param name="aCompoundNo"></param>
        /// <param name="aLength"></param>
        /// <returns></returns>
        private string AddLeadingZeros(string aCompoundNo, int aLength)
        {
            while (aCompoundNo.Length < aLength)
            {
                aCompoundNo = '0' + aCompoundNo;
            }
            return aCompoundNo;
        }
        /// <summary>
        /// Formats CompoundNo as a BAY No
        /// </summary>
        /// <param name="aCompoundNo"></param>
        /// <returns></returns>
        private string FormatBAY(string aCompoundNo)
        {   
            string tmpCompoundNo = aCompoundNo.ToUpper();
            string tmpSeperator = " "; // Seperator (default: " " -> "BAY 123456" or for old baynos a letter i.e. "A"  BAYA001040

            if (tmpCompoundNo.StartsWith("BAY"))
            {
                tmpCompoundNo = tmpCompoundNo.Substring(3, tmpCompoundNo.Length - 3).Trim();
            }
            //Baynos can start with letters like I123456. The Bayno is then formatted as BAYI123456
            //Therefore we only append the BAY if the bayno does not start with a digit.
            // if A1040 -> BAYA001040            
            if (tmpCompoundNo[0] >= 'A' && tmpCompoundNo[0] <= 'Z')
            {
              tmpSeperator = tmpCompoundNo.Substring(0,1);
              tmpCompoundNo = tmpCompoundNo.Substring(1,tmpCompoundNo.Length-1).Trim();
            }
            if (tmpCompoundNo == string.Empty) {
                return aCompoundNo;
            }
            return "BAY" + tmpSeperator + AddLeadingZeros(tmpCompoundNo, 6);
        }

        /// <summary>
        /// Formats CompoundNo as a ZK No
        /// </summary>
        /// <param name="aCompoundNo"></param>
        /// <param name="aPrefix"></param>
        /// <param name="aLength"></param>
        /// <returns></returns>
        private string FormatCompoundNo(string aCompoundNo, string aPrefix, int aLength)
        {
            string tmpCompoundNo = aCompoundNo.ToUpper();

            if (tmpCompoundNo.ToUpper().StartsWith(aPrefix))
            {
                tmpCompoundNo = tmpCompoundNo.Substring(aPrefix.Length, tmpCompoundNo.Length - aPrefix.Length).Trim();
                if (tmpCompoundNo != string.Empty && char.IsDigit(tmpCompoundNo, 0))
                {
                    tmpCompoundNo = AddLeadingZeros(tmpCompoundNo, aLength);
                }
                return aPrefix + " " + tmpCompoundNo;
            }

            if (aCompoundNo != string.Empty && char.IsDigit(aCompoundNo, 0)) 
            {
                return aPrefix + " " + AddLeadingZeros(aCompoundNo, aLength);
            } //Possibly another compound type or something completely different
            return aCompoundNo;
        }

        /// <summary>
        /// Always takes the selection and disables the special pdc data entry sheet handling of the super class.
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="interactive"></param>
        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            AddIn.Application.ScreenUpdating = false;
            AddIn.Application.EnableEvents = false;
            try
            {
                PerformActionBySelection(null, Kind);
                return new ActionStatus();
            }
            finally
            {
                AddIn.EnableExcel();
            }
        }

        /// <summary>
        /// Formats the compound no in the cell area for the specified compound type
        /// </summary>
        /// <param name="owner"></param>
        /// <param name="anActionKind">Specifies the desired compound type through dedicated action kinds</param>
        /// <param name="aSheetInfo"></param>
        /// <param name="writeRanges"></param>
        /// <param name="tmpSheet"></param>
        //TODO -> UpdateSheet
        protected override void UpdateSheet(
            Control owner, 
            CompoundInfoActionKind anActionKind, 
            SheetInfo aSheetInfo, 
            Ranges writeRanges, 
            Excel.Worksheet tmpSheet)
        {
            // get current compound nos from range
            var tmpCompoundNosObject = writeRanges.CompoundnoRange.Value[Excel.XlRangeValueDataType.xlRangeValueDefault];
            var tmpCompoundNos = tmpCompoundNosObject as object[,] ?? new [,] { { tmpCompoundNosObject } };
            // Format compound nos
            for (int y = tmpCompoundNos.GetLowerBound(0); y <= tmpCompoundNos.GetUpperBound(0); y++)
            {
                for (int x = tmpCompoundNos.GetLowerBound(1); x <= tmpCompoundNos.GetUpperBound(1); x++)
                {
                    string tmpContent = "" + tmpCompoundNos[y, x];
                    tmpContent = tmpContent.Trim();
                    tmpContent = System.Text.RegularExpressions.Regex.Replace(tmpContent, "( ){2,}"," ");
                    if (tmpContent == string.Empty)
                    {
                        continue;
                    }
                    switch (anActionKind)
                    {
                        case CompoundInfoActionKind.FormatSelectedCop:
                            tmpContent = FormatCompoundNo(tmpContent, "COP", 7);
                            break;
                        case CompoundInfoActionKind.FormatSelectedCos:
                            tmpContent = FormatCompoundNo(tmpContent, "COS", 7);
                            break;
                        case CompoundInfoActionKind.FormatSelectedZk:
                            tmpContent = FormatCompoundNo(tmpContent, "ZK", 7);
                            break;
                        default:
                            tmpContent = FormatBAY(tmpContent);
                            break;
                    }
                    tmpCompoundNos[y, x] = tmpContent;
                }
            }
            // write back formatted compound nos
            writeRanges.CompoundnoRange.Value2 = tmpCompoundNos;
        }
    }
}
