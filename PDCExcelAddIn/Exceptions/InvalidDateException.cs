namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    class InvalidDateException: PDCExcelAddInFault
  {
        public InvalidDateException(string dateObject, string id)
      : base(PDCExcelAddInFaultMessage.INVALID_DATE, new object[] { dateObject, id})
    {
    }
  }
}
