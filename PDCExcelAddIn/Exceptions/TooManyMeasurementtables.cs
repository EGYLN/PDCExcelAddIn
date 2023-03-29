namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    /// <summary>
    /// Thrown if the measurement sheet has no room for more tables
    /// </summary>
    class TooManyMeasurementtables:PDCExcelAddInFault
    {
        public TooManyMeasurementtables()
            : base(PDCExcelAddInFaultMessage.TOO_MANY_MEASUREMENT_TABLES)
        {
        }
    }

}
