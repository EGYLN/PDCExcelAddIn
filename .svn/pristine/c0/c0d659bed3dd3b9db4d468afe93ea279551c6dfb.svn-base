using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /*
     * SB: This class should be merged with its super class. The superclass still exists only because it
     * may be used in older workbooks, which can not be loaded/deserialized easily if the super class is removed. 
     * Most probably this is not really a problem anymore.
     */
    /// <summary>
    /// Alternative/Experimental implementation of the MeasurementHandler with the
    /// target to increase the performance of the Measurementtable creation
    /// by copying from a table prototype
    /// </summary>
    [Serializable]
    public class MeasurementCopyHandler:MeasurementHandler
    {
        public MeasurementCopyHandler(Excel.Worksheet aSheet, Lib.Testdefinition aTD)
            : base(aSheet, aTD)
        {
        }        

        internal MeasurementPDCListObject firstTable;
        internal object initialSheetId;

     
    }
}
