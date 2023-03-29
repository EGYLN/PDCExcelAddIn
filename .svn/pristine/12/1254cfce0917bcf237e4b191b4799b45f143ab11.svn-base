using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.VBA
{
    [ComVisible(true)]
  [Guid("33C7064B-F4D5-4854-9B2D-51CB0C6A21F4")]
  [InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface IVBAInterface
  {
    [Description("Starts the ClearData action. The pdc menu and the upload menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus ClearData(Excel.Range range);
    [Description("Starts the Delete action. The pdc menu and the upload menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus Delete(Excel.Range range);
    [Description("Returns the cwid of the loggedin user or null if not logged in")]
    string GetPDCUsername(Excel.Range range);
    [Description("Retrieves the measurement data and writes them to the single measurement table sheet of the given range object's workbook.\n Returns:\n   -999 when an unknown error occured.\n   -10 when the action button is not enabled.\n   -9 when the action button shortcut is not found.\n   -8 when the PDC action button is not found.\n   -7 when the PDC menu is not present or visible.\n   -6 when no range is given.\n   -5 when the sheet is no PDC sheet.\n   -4 when the user is not logged in.\n    0 when the action executed successfully.")]
    VBAActionStatus RetrieveMeasurementLevelData(Excel.Range range);
    [Description("Starts the Update action. The pdc menu and the upload menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus Update(Excel.Range range);
    [Description("Starts the Upload action. The pdc menu and the upload menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus Upload(Excel.Range range);
    [Description("Starts the Retrieves action. The pdc menu and Search Testdata menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus SearchTestData(Excel.Range range);
    [Description("Starts the Validate action. The pdc menu and Validate menu button must be available. Returns one of the defined readonly values as status code")]
    VBAActionStatus Validate(Excel.Range range);

  }

  [ClassInterface(ClassInterfaceType.None)]
  [ComVisible(true)]
  public class VBAInterface:IVBAInterface
  {
    #region ClearData
    public VBAActionStatus  ClearData(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton = tmpSingleton = GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }

      return tmpSingleton.StartClearData();
    }
    #endregion

    #region Delete
    public VBAActionStatus Delete(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton = tmpSingleton = GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }

      return tmpSingleton.StartDelete();
    }
    #endregion

    #region GetCOMSingleton
    private ICOMSingleton GetCOMSingleton(Excel.Range aRange)
    {
      try
      {
        Excel.Application tmpApplication = aRange.Application;
        object tmpAddInName = "PDCExcelAddIn";
        Office.COMAddIn tmpAddIn = tmpApplication.COMAddIns.Item(ref tmpAddInName);
        if (tmpAddIn == null)
        {
          return null;
        }
        return (ICOMSingleton)tmpAddIn.Object;
      }
      catch (Exception e)
      {
        PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_COM, "COM call failed", e);
        return null;
      }
    }
    #endregion

    #region GetPDCUsername
    public string GetPDCUsername(Excel.Range aRange)
    {
      return GetCOMSingleton(aRange).GetPDCUser();
    }
    #endregion

    #region Retrieve
    public VBAActionStatus SearchTestData(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton = tmpSingleton = GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }

      return tmpSingleton.StartSearchTestData();
    }
    #endregion

    #region RetrieveMeasurementLevelData
    /// <summary>
    ///   Retrieves the measurement data and writes them to the single measurement table sheet.
    /// </summary>
    /// <param name="range">
    ///   A range on the Excel sheet.
    /// </param>
    /// <returns>
    ///   The status of the operation:
    ///    -999 when an unknown error occured.
    ///     -10 when the action button is not enabled.
    ///      -9 when the action button shortcut is not found.
    ///      -8 when the PDC action button is not found.
    ///      -7 when the PDC menu is not present or visible.
    ///      -6 when no range is given.
    ///      -5 when the sheet is no PDC sheet.
    ///      -4 when the user is not logged in.
    ///       0 when the action executed successfully.
    /// </returns>
    public VBAActionStatus RetrieveMeasurementLevelData(Excel.Range range)
    {
      ICOMSingleton singleton = GetCOMSingleton(range);
      if (singleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }
      return singleton.StartRetrieveMeasurementLevelData();
    }
    #endregion

    #region Upload
    public VBAActionStatus Upload(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton = tmpSingleton = GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }
      return tmpSingleton.StartUpload();
    }
    #endregion

    #region Update
    public VBAActionStatus Update(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton = tmpSingleton = GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }

      return tmpSingleton.StartUpdate();
    }
    #endregion
    #region Validate
    public VBAActionStatus Validate(Excel.Range aRange)
    {
      ICOMSingleton tmpSingleton= GetCOMSingleton(aRange);
      if (tmpSingleton == null)
      {
        VBAActionStatus tmpStatus = new VBAActionStatus();
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_FATAL;
        tmpStatus.Messages = new string[] { "COM Call failed" };
      }

      return tmpSingleton.StartValidate();
    }
    #endregion

  }
}
