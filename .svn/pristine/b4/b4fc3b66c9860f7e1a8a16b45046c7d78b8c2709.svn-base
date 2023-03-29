using System.Runtime.InteropServices;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    /// <summary>
    /// Message codes for PDCExcelAddIn exception
    /// </summary>
    [ComVisible(false)]
  public enum PDCExcelAddInFaultMessage
  {
    /// <summary>
    /// The active sheet is not a PDC sheet, but a PDC sheet is needed
    /// </summary>
    ACTIVE_SHEET_NO_PDC_SHEET,
    /// <summary>
    /// The active sheet is a subordinate PDC Sheet, and the selected action is not allowed here
    /// </summary>
    SUBORDINATE_PDC_SHEET,
    /// <summary>
    /// A sheet should be created with a name that is already in use
    /// </summary>
    SHEET_NAME_ALREADY_EXISTS,
    /// <summary>
    /// Too many measurement tables were created for the measurement sheet
    /// </summary>
    TOO_MANY_MEASUREMENT_TABLES,
    /// <summary>
    /// Two rows with the same tupel of CompoudnNo, Experimentno and ExperimentlevelVariablevalues
    /// </summary>
    AMBITIOUS_EXPERIMENTS,
    /// <summary>
    /// No Experiment can be found for the tupel on the measurement sheet 
    /// </summary>
    NO_EXPERIMENT_FOUND_FOR_MEASUREMENT,
    /// <summary>
    /// An invalid Excel range was specified.
    /// </summary>
    INVALID_RANGE,
    /// <summary>
    /// The user specified an invalid date (range)
    /// </summary>
    INVALID_DATE,
    /// <summary>
    /// The user entered a number which cannot be interpreted 
    /// </summary>
    INVALID_NUMBER_FORMAT
  }
}
