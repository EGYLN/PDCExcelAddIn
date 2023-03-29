using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Actions;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.VBA
{
    //This class could be used in later version of PDC as
    // a true Excel-wide Singleton which could be used by
    // VBA/Excel UdF/OpenLib to directly access the PDC add in functionality
    // The PDCExcelAddIn.RequestComAddInAutomationService must be implemented to
    // return the singleton instance.
    [ComVisible(true)]
  [Guid("5976AA35-6225-4e61-8943-0510760B3EB3")]
  public interface ICOMSingleton
  {
    string GetPDCUser();
    VBAActionStatus StartClearData();
    VBAActionStatus StartDelete();
    VBAActionStatus StartRetrieveMeasurementLevelData();
    VBAActionStatus RetrieveMeasurementLevelData();
    VBAActionStatus StartUpdate();
    VBAActionStatus StartUpload();
    VBAActionStatus StartSearchTestData();
    VBAActionStatus StartValidate();
  }
  
  [ClassInterface(ClassInterfaceType.None)]
  [ComVisible(true)]
  public class COMSingleton:ICOMSingleton
  {
    private static COMSingleton mySingleton = new COMSingleton();

    #region methods

    #region ExecuteAction
    private VBAActionStatus ExecuteAction(PDCAction anAction, string anActionName)
    {
      return ExecuteAction(anAction, anActionName, false);
    }
    /// <summary>
    /// Calls the specified PDCActions and returns the execution status.
    /// </summary>
    /// <param name="anAction"></param>
    /// <param name="anActionName">Name of operation, for error message</param>
    /// <returns></returns>
    private VBAActionStatus ExecuteAction(Actions.PDCAction anAction, string anActionName, bool interactive)
    {
      ActionStatus tmpStatus = null;
      if (anAction == null)
      {
        tmpStatus = new ActionStatus(new Lib.PDCMessage(string.Format(Properties.Resources.MSG_ACTION_NOT_FOUND, anActionName), Lib.PDCMessage.TYPE_FATAL));
        return ToVBA(tmpStatus);
      }
      if (Globals.PDCExcelAddIn.PdcService.UserInfo == null)
      {
        tmpStatus = new ActionStatus(new Lib.PDCMessage(string.Format(Properties.Resources.MSG_NOT_LOGGED_IN, anActionName), Lib.PDCMessage.TYPE_FATAL));
        return ToVBA(tmpStatus);
      }
      SheetInfo tmpSheetInfo = anAction.GetActiveSheetInfo();
      if (tmpSheetInfo == null || !tmpSheetInfo.IsMainSheet)
      {
        tmpStatus = new ActionStatus(new Lib.PDCMessage(string.Format(Properties.Resources.MSG_NO_PDC_SHEET, anActionName), Lib.PDCMessage.TYPE_FATAL));
        return ToVBA(tmpStatus);
      }
      if (!anAction.Enabled || !anAction.Visible)
      {
        tmpStatus = new ActionStatus(new Lib.PDCMessage(string.Format(Properties.Resources.MSG_ACTION_NOT_ENABLED, anActionName), Lib.PDCMessage.TYPE_FATAL));
        return ToVBA(tmpStatus);
      }
      tmpStatus = anAction.PerformAction(interactive);
      return ToVBA(tmpStatus);
    }
    #endregion

    #region GetPDCUser
    /// <summary>
    /// Returns the name of the logged in user or null
    /// </summary>
    /// <returns></returns>
    public string GetPDCUser()
    {
        PDCLogger.TheLogger.LogDebugMessage(nameof(GetPDCUser), "Called via COMSingleton.");
      Lib.UserInfo tmpUser  = Globals.PDCExcelAddIn.PdcService.UserInfo;
      return tmpUser == null ? null : tmpUser.Cwid;
    }
    #endregion

    #region StartClearData
    public VBAActionStatus StartClearData()
    {
      Actions.PDCAction tmpAction = Globals.PDCExcelAddIn.ClearDataAction;
      return ExecuteAction(tmpAction, "ClearData");
    }
    #endregion

    #region StartDelete
    public VBAActionStatus StartDelete()
    {
      return ExecuteAction(Globals.PDCExcelAddIn.DeleteAction, "Delete");
    }
    #endregion

    #region SearchTestData
    /// <summary>
    ///   Retrieves testdata
    /// </summary>
    /// <returns>
    ///   The VBAActionStatus object with all messages.
    /// </returns>
    public VBAActionStatus StartSearchTestData()
    {
      return ExecuteAction(Globals.PDCExcelAddIn.SearchTestdataAction, "SearchTestData");
    }
    #endregion

    #region StartRetrieveMeasurementLevelData
    /// <summary>
    ///   Retrieves the measurement data and writes them to the single measurement table sheet.
    /// </summary>
    /// <returns>
    ///   The VBAActionStatus object with all messages.
    /// </returns>
    public VBAActionStatus StartRetrieveMeasurementLevelData()
    {
      return ExecuteAction(Globals.PDCExcelAddIn.RetrieveMeasurementLevelDataAction, "RetrieveMeasurementLevelData");
    }

    /// <summary>
    ///   Retrieves the measurement data and writes them to the single measurement table sheet.
    ///   This method is the target of the appropriate shortcut key and runs in interactive mode.
    /// </summary>
    /// <returns>
    ///   The VBAActionStatus object with all messages.
    /// </returns>
    public VBAActionStatus RetrieveMeasurementLevelData()
    {
      return ExecuteAction(Globals.PDCExcelAddIn.RetrieveMeasurementLevelDataAction, "RetrieveMeasurementLevelData", true);
    }
    #endregion

    #region StartUpdate
    public VBAActionStatus StartUpdate()
    {
      return ExecuteAction(Globals.PDCExcelAddIn.UpdateDataAction, "Update");
    }
    #endregion

    #region StartUpload
    public VBAActionStatus StartUpload()
    {
      Actions.PDCAction tmpAction = Globals.PDCExcelAddIn.UploadDataAction;
      return ExecuteAction(tmpAction, "Upload");
    }
    #endregion

    #region StartValidate
    public VBAActionStatus StartValidate()
    {
      Actions.PDCAction tmpAction = Globals.PDCExcelAddIn.ValidateAction;
      return ExecuteAction(tmpAction, "Upload");
    }
    #endregion
    #region ToVBA
    /// <summary>
    /// Transforms an instance of ActionStatus into the simpler VBAActionStatus struct
    /// </summary>
    /// <param name="aStatus"></param>
    /// <returns></returns>
    private VBAActionStatus ToVBA(ActionStatus aStatus)
    {
      VBAActionStatus tmpStatus = new VBAActionStatus();
      tmpStatus.Status = 0;
      List<string> tmpResult = new List<string>();
      if (aStatus == null)
      {
        return tmpStatus;
      }
      if (aStatus.Failure != null)
      {
        tmpResult.Add(aStatus.Failure.Message);
        tmpResult.Add(aStatus.Failure.StackTrace);
        tmpStatus.Status = Lib.PDCConstants.C_LOG_LEVEL_ERROR;
      }
      if (aStatus.Messages != null)
      {
        foreach (Lib.PDCMessage tmpMessage in aStatus.Messages)
        {
          tmpResult.Add(tmpMessage.MessageTypeText + ":" + tmpMessage.Message);
          int tmpMessageType = tmpMessage.LogLevel;
          if (tmpMessageType > tmpStatus.Status)
          {
            tmpStatus.Status = tmpMessageType;
          }
        }
      }
      tmpStatus.Messages = tmpResult.ToArray();
      //tmpStatus.failure = aStatus.Failure == null?null:aStatus.Failure.Message + "\n" + aStatus.Failure.StackTrace;
      if (aStatus.Failure != null)
      {
        tmpStatus.ExceptionMessage = new string[] { aStatus.Failure.Message, aStatus.Failure.StackTrace };
      }
      return tmpStatus;
    }
    #endregion

    #endregion

    #region properties

    #region Singleton
    public static COMSingleton Singleton
    {
      get
      {
        return mySingleton;
      }
    }
    #endregion

    #endregion
  }
}
