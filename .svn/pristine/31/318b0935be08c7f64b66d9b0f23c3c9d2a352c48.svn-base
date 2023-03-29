using System.Runtime.InteropServices;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    /// <summary>
    /// Is thrown if an operation on the active sheet can not be executed since the active sheet
    /// is not a PDC sheet.
    /// </summary>
    [ComVisible(false)]
  public class NoPDCSheetException:PDCExcelAddInFault
  {
    public NoPDCSheetException() : base(PDCExcelAddInFaultMessage.ACTIVE_SHEET_NO_PDC_SHEET)
    {
    }
  }
}
