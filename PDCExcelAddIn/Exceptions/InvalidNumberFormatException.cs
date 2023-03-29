namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    public class InvalidNumberFormatException : PDCExcelAddInFault
  {
    public InvalidNumberFormatException(string varName,object value, string valueType,string sep1,string sep2)
      : base(PDCExcelAddInFaultMessage.INVALID_NUMBER_FORMAT, new object[] { varName, value,valueType, sep1,sep2 })
    {
    }
  }
}
