using System.Runtime.InteropServices;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    [ComVisible(false)]
  public class SubordinatePDCSheetException:PDCExcelAddInFault
  {
    public SubordinatePDCSheetException():base(PDCExcelAddInFaultMessage.SUBORDINATE_PDC_SHEET)
    {
    }
  }
}
