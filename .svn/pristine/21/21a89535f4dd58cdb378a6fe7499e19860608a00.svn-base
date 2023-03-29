using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Security.Permissions;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using Microsoft.Win32;
using log4net;

using BBS.ST.BHC.AutoUpdater;
using BBS.ST.BHC.AutoUpdater.Business;
using BBS.ST.BHC.AutoUpdater.Data;
using BBS.ST.BHC.AutoUpdater.Services;
using BBS.ST.BHC.AutoUpdater.Dialogs;
using BBS.ST.BHC.PDC.AutoUpdater.Data;
using BBS.ST.BHC.AppUpdater.Dialogs;

namespace BBS.ST.BHC.PDC.AutoUpdater.Services
{
  /// <summary>
  ///   PdcUpdateService can update the PDC application.
  /// </summary>
  //[PermissionSetAttribute(SecurityAction.Demand, Name = "FullTrust")]
  public class PdcUpdateService : AppUpdateService
  {
    private static ILog MyLogger = LogManager.GetLogger(typeof(PdcUpdateService));

    // .ctor
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region .ctor

    /// <summary>
    ///   Creates a new instance of the AppInfoService class.
    /// </summary>
    /// <param name="logger">
    ///   The logger to use for writing logging informations.
    /// </param>
    /// <param name="updateSource">
    ///   The update source to use for this updater.
    /// </param>
    public PdcUpdateService(UpdateInfo info, IUpdateSource updateSource) : base(info, updateSource)
    {
    }

    #endregion


    // Methods
    // ////////////////////////////////////////////////////////////////////////////////////////////

    private String GetPdcLanguage()
    {
      String tmpLanguage = PDCLanguage.English;
      if (this.UpdateInfo.SettingDictionary.ContainsKey(BBS.ST.BHC.PDC.AutoUpdater.Data.SettingName.Language))
        tmpLanguage = this.UpdateInfo.SettingDictionary[BBS.ST.BHC.PDC.AutoUpdater.Data.SettingName.Language];

      return tmpLanguage;
    }

    #region CheckIsUpdateRequired

    /// <summary>
    ///   Checks if the update is required.
    /// </summary>
    /// <returns>
    ///   true if the specified update is required; otherwise false.
    /// </returns>
    public Boolean CheckIsUpdateRequired()
    {
      String      xmlText;
      XmlDocument xmlDoc;
      XmlNodeList xmlNodeList;


      try
      {
        xmlText = this.UpdateSource.GetInformation(this.UpdateInfo.ApplicationID);

        if (String.IsNullOrEmpty(xmlText)) return false;


        xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(xmlText);


        xmlNodeList = xmlDoc.DocumentElement.GetElementsByTagName("isrequired");

        if (xmlNodeList.Count != 1)
        {
          throw new Exception("Incorrect XML format found on server! Exactly one \"isrequired\" tag " +
            "is expected representing a flag that determines if the update is required.");
        }

        return xmlNodeList[0].InnerText.Trim().ToLower() == "true";
      }
      catch (Exception ex)
      {
        MyLogger.Error("Error occured. Value 'isrequired' is set to true.", ex);

        return true;
      }
    }

    #endregion

    #region CloseExcel

    /// <summary>
    ///   Promt the user to close all running excel applications.
    /// </summary>
    /// <returns>
    ///   true if the user closes all running excel applications; otherwise false.
    /// </returns>
    private Boolean CloseExcel()
    {
      Process[] processes;
      Boolean foundExcelProcess;
      Boolean isUpdateRequired;
      Int32 counter = 1;

      isUpdateRequired = this.CheckIsUpdateRequired();


      while (true)
      {
        try
        {
          processes = Process.GetProcesses();

          foundExcelProcess = false;
          foreach (Process tmpProcess in processes)
          {
            if (tmpProcess.ProcessName.ToUpper().Equals("EXCEL"))
            {

              foundExcelProcess = true;
              break;
            }
          }

          if (!foundExcelProcess) return true;

          String tmpMsg = "A new PDC version has been found. Please close all excel applications and klick the OK button.";
          if (GetPdcLanguage().Equals(PDCLanguage.German))
            tmpMsg = "Eine neue PDC Version wurde gefunden. Bitte schliessen Sie die laufenden Excel Programme und klicken Sie OK.";

          String tmpTitle = "Update available";
          if (GetPdcLanguage().Equals(PDCLanguage.German))
            tmpTitle = "Update verfügbar";

          String tmpTerminate = "Terminate Excel!";
          if (GetPdcLanguage().Equals(PDCLanguage.German))
            tmpTerminate = "Excel beenden";

          String tmpCancle = "Cancel";
          if (GetPdcLanguage().Equals(PDCLanguage.German))
            tmpCancle = "Abbrechen";

          MessageBoxButtons tmpButtons = MessageBoxButtons.OK;
          if (!isUpdateRequired)
            tmpButtons = MessageBoxButtons.OKCancel;

          DialogResult tmpResult = ExcelUpdateMessageBox.Show(tmpButtons, tmpTitle, tmpMsg, "Ok", tmpCancle, tmpTerminate, counter >= 2);

          if (tmpResult == DialogResult.Cancel) 
            return false;

          if (tmpResult == DialogResult.Retry)
          {
            if (this.KillExcel())
              return true;
          }

          counter++;
        }
        catch (Exception ex)
        {
          MyLogger.Error("Error in CloseExcel", ex);
          if (GetPdcLanguage().Equals(PDCLanguage.German))
          {
            MessageBox.Show("Excel konnte nicht beendet werden. Wahrscheinlich ist noch ein anderer Benutzer auf dem Computer eingeloggt und arbeitet mit Excel.\nEine Aktualisierung ist nicht möglich!", "Fehler");
          }
          else
          {
            MessageBox.Show("Error terminating Excel.\nProbably another user is also logged on this computer and is working with Excel.\nNo Update possible!", "Error");
          }
          return false;
        }
      }
    }

    private Boolean KillExcel()
    {
      Process[] processes;

      String tmpMsg = "Are you sure to terminate all Excel instances? It may result in data loss.";
      if (GetPdcLanguage().Equals(PDCLanguage.German))
        tmpMsg = "Sind Sie sicher, dass Sie alle Excelinstanzen beenden wollen? Nicht gespeicherte Daten können verloren gehen.";

      String tmpTitle = "Warning!";
      if (GetPdcLanguage().Equals(PDCLanguage.German))
        tmpTitle = "Warnung!";
      String tmpCancel = "Cancel";
      if (GetPdcLanguage().Equals(PDCLanguage.German))
        tmpCancel = "Abbrechen";
      

      MessageBoxButtons tmpButtons = MessageBoxButtons.OKCancel;
      DialogResult tmpResult = UpdateMessageBox.Show(tmpTitle, tmpMsg, tmpButtons, "Ok", tmpCancel);
      
      if (tmpResult != DialogResult.OK) return false;
    

      processes = Process.GetProcesses();

      for (Int32 i = 0; i < processes.Length; i++)
      {
        if (processes[i].ProcessName.ToUpper().Equals("EXCEL"))
        {
          processes[i].Kill();
          processes[i].WaitForExit();
        } 
      }

      return true;
    }

    #endregion


    #region Prepare

    /// <summary>
    ///   Promt the user to close all existing excel sheets.
    /// </summary>
    /// <returns>
    ///   true if the user closes all open excel sheets; otherwise false.
    /// </returns>
    public override bool Prepare()
    {
      this.OnStatusChanged(10, "Prepare installation...");

      return this.CloseExcel();
    }

    #endregion


    // Properties
    // ////////////////////////////////////////////////////////////////////////////////////////////

  }
}
