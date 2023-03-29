namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    class AmbitiousExperimentsException : PDCExcelAddInFault
  {
    public AmbitiousExperimentsException(long row1, long row2)
      : base(PDCExcelAddInFaultMessage.AMBITIOUS_EXPERIMENTS, new object[] { row1, row2 })
    {
    }
  }
}
