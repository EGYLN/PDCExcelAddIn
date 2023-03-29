namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    public class InvalidRangeException : PDCExcelAddInFault
  {
    public InvalidRangeException()
      : base(PDCExcelAddInFaultMessage.INVALID_RANGE)
    {
    }
  }
}
