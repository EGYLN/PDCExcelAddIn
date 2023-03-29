namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    class SheetNameAlreadyExistsException : PDCExcelAddInFault
    {
        public SheetNameAlreadyExistsException(string aSheetName, string aSheetType)
            : base(PDCExcelAddInFaultMessage.SHEET_NAME_ALREADY_EXISTS,new object[] {aSheetName, aSheetType})
        {
        }
    }
}
