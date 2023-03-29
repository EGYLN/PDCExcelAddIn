using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;

using log4net;

using BBS.ST.BHC.AutoUpdater.Business;
using BBS.ST.BHC.AutoUpdater.Data;
using BBS.ST.BHC.AutoUpdater.Services;
using BBS.ST.BHC.AutoUpdater;

using BBS.ST.BHC.PDC.AutoUpdater.Security;
using BBS.ST.BHC.PDC.AutoUpdater.Services;
using BBS.ST.BHC.PDC.AutoUpdater.IO;
using BBS.ST.BHC.PDC.AutoUpdater.Data;

namespace BBS.ST.BHC.PDC.AutoUpdater.Business
{
  using PdcSettingName = BBS.ST.BHC.PDC.AutoUpdater.Data.SettingName;
  using BBS.ST.BHC.AppUpdater.Dialogs;
  using System.Windows.Forms;

  /// <summary>
  ///   The PdcUpdater extends the default Updater for pdc specific purposes.
  /// </summary>
  public class PdcUpdater : Updater
  {
    private static ILog MyLogger = LogManager.GetLogger(typeof(PdcUpdater));

    private SettingsFile mySettingsFile;
    private Dictionary<String, String> myPdcConfigSettings;

    private static String MySettingsFile = "pdcconfig.properties";
    private static Char MySettingsSeparator = '=';


    // Nested Types
    // ////////////////////////////////////////////////////////////////////////////////////////////




    // .ctor
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region .ctor

    /// <summary>
    ///   Creates a new instance of the Updater class.
    /// </summary>
    public PdcUpdater()
    {
    }

    #endregion


    // Methods
    // ////////////////////////////////////////////////////////////////////////////////////////////    

    private String GetPdcLanguage()
    {
      String tmpLanguage = PDCLanguage.English;
      if (this.PdcConfigSettings.ContainsKey(BBS.ST.BHC.PDC.AutoUpdater.Data.SettingName.Language))
        tmpLanguage = this.PdcConfigSettings[BBS.ST.BHC.PDC.AutoUpdater.Data.SettingName.Language];
      return tmpLanguage;
    }

    protected override void AfterInstall()
    {
      String msg;
      String title;

      base.AfterInstall();

      msg = "The PDC update has been installed successfully. You can now continue working with Excel.";
      if (GetPdcLanguage().Equals(PDCLanguage.German))
        msg = "Das PDC-Update wurde erfolgreich installiert. Sie können nun mit Excel weiter arbeiten.";

      title = "Update succeeded";
      if (GetPdcLanguage().Equals(PDCLanguage.German))
        title = "Update erfolgreich";

      UpdateMessageBox.Show(title, msg, MessageBoxButtons.OK, "OK", String.Empty);
    }

    #region CreateUpdateService

    protected override AppUpdateService CreateUpdateService(IUpdateSource updateSource, UpdateInfo updateInfo)
    {
      return new PdcUpdateService(updateInfo, updateSource);
    }

    #endregion

    #region CreateUpdateSource

    /// <summary>
    ///   Creates the update source.
    /// </summary>
    /// <param name="logger">
    ///   The logger.
    /// </param>
    /// <returns>
    ///   The creates update source.
    /// </returns>
    protected override IUpdateSource CreateUpdateSource()
    {
      String                 url = null;
      String                 username = null;
      String                 license = null;

      if (!this.PdcConfigSettings.TryGetValue(PdcSettingName.PDCUpdateURL, out url))
      {
        MyLogger.Error("Setting '" + PdcSettingName.PDCUpdateURL + "' is not available.");

        return null;
      }

      MyLogger.InfoFormat("Webserver URL: {0}", url);

      if (!this.PdcConfigSettings.TryGetValue(PdcSettingName.PDCUpdateUser, out username) ||
          !this.PdcConfigSettings.TryGetValue(PdcSettingName.PDCUpdateLicense, out license) ||
          String.IsNullOrEmpty(username) ||
          String.IsNullOrEmpty(license))
      {
        return new WebServiceUpdateSource(url);
      }

      MyLogger.InfoFormat("Webserver Username: {0}", username);

      try
      {
        license = AesEncoder.Decrypt(license, MyKey);
      }
      catch
      {
        MyLogger.Error("An error occurred while decrypting the license key.");

        return null;
      }

      return new WebServiceUpdateSource(url, username, license);      
    }

    #endregion

    protected override UpdateInfo CreateUpdateInfo()
    {
      UpdateInfo tmpInfo = base.CreateUpdateInfo();
      String tmpLanguage = PDCLanguage.English;
      if (this.PdcConfigSettings.ContainsKey(PdcSettingName.Language))
        tmpLanguage = this.PdcConfigSettings[PdcSettingName.Language];
      tmpInfo.SettingDictionary.Add(PdcSettingName.Language, tmpLanguage);
      return tmpInfo;
    }

    private Dictionary<String, String> ReadPdcConfigSettings()
    {
      Dictionary<String, String> tmpSettings = this.SettingsFile.GetSettings(); // read pdcconfig.properties
      if (tmpSettings == null)
      {
        String tmpMsg = String.Format("Settings file {0} could not be found.", MySettingsFile);
        MyLogger.Error(tmpMsg);
        throw new Exception(tmpMsg);
      }
      return tmpSettings;
    }

    // Properties
    // ////////////////////////////////////////////////////////////////////////////////////////////

    private SettingsFile SettingsFile
    {
      get
      {
        if (this.mySettingsFile != null) return this.mySettingsFile;

        this.mySettingsFile = new SettingsFile(MySettingsFile, MySettingsSeparator);

        return this.mySettingsFile;
      }
    }

    public Dictionary<String, String> PdcConfigSettings
    {
      get 
      { 
        if (myPdcConfigSettings == null)
          myPdcConfigSettings = ReadPdcConfigSettings();
        return myPdcConfigSettings; 
      }
    }
  }
}
