using System;
using System.Collections.Generic;
using System.IO;

namespace BBS.ST.BHC.BSP.PDC.Lib.Util
{
  /// <summary>
  /// Singleton which reads in a simple property file which may contain user-specific configuration settings
  /// </summary>
  public class UserConfiguration
  {
    /// <summary>
    /// Enables the automatic cancellation of cell editing mode
    /// </summary>
    public const string PROP_ENABLE_EXIT_CELLEDIT = "EnableAutomaticExitCellEdit";

    /// <summary>
    /// URL to the PDC Webservice
    /// </summary>
    public const string PROP_WS_PDC_URL = "PDCServiceURL";

    /// <summary>
    /// URL to the Customer Management service
    /// </summary>
    public const string PROP_WS_CUSTOMERMANAGEMENT_URL = "WSCustomerManagementServiceURL";

    /// <summary>
    /// URL to the compound information service
    /// </summary>
    public const string PROP_WS_COMPOUNDINFORMATIONSERVICE_URL = "WSCompoundInformationServiceURL";

    /// <summary>
    /// URL to the PDC Portal
    /// </summary>
    public const string PROP_PORTAL_PDC_URL = "PDCPortalURL";

    /// <summary>
    /// Encoded licence key for accessing the update werb service
    /// </summary>
    public const string PROP_UPDATE_PDC_URL = "PDCUpdateURL";

    /// <summary>
    /// Encoded licence key for accessing the update werb service
    /// </summary>
    public const string PROP_UPDATE_PDC_LICENSE = "PDCUpdateLicense";

    /// <summary>
    /// Username for accessing the update webservice
    /// </summary>
    public const string PROP_UPDATE_PDC_USER = "PDCUpdateUser";

    /// <summary>
    /// Whether to draw a border around the data entry area
    /// </summary>
    public const string PROP_DRAW_DATA_ENTRY_BORDER = "DrawDataEntryBorder";

    /// <summary>
    /// Whether to use the QA settings
    /// </summary>
    public const string PROP_PROD_QA_MODE = "QAMode";

    /// <summary>
    /// Wether to use the menu bar instead of a toolbar
    /// </summary>
    public const string PROP_USE_MENU = "MenuInsteadOfToolbar";

    /// <summary>
    /// Whether to ignore hidden rows during upload, validation
    /// </summary>
    public const string PROP_IGNORE_HIDDEN_ROWS = "IgnoreHiddenRows";


    /// <summary>
    /// The shortcut which is assigned to the RetrieveMeasurementLevelData action
    /// </summary>
    public const string PROP_RETRIEVE_MEASUREMENTS_SHORTCUT = "RetrieveMeasurementsShortcut";

    public const string PROP_RETRIEVE_MEASUREMENTS_SHORTCUT_TEXT = "RetrieveMeasurementsShortcutText";

    /// <summary>
    /// Can set the TimeOutsinMillis
    /// </summary>
    public const string PROP_WEBSERVICE_TIMEOUT_MILLIS = "WebserviceTimeoutMillies";

    private static UserConfiguration myConfiguration;

    private readonly Dictionary<string,string> myProperties = new Dictionary<string,string>();

    private static readonly object LOCK = new object();

    #region methods

    #region GetBooleanProperty
    /// <summary>
    /// Returns a boolean Property value
    /// </summary>
    /// <param name="aPropertyName">The name of the property</param>
    /// <param name="aDefault">The default value of the property</param>
    /// <returns>The value from the properties file or the default value</returns>
    public bool GetBooleanProperty(string aPropertyName, bool aDefault)
    {
      string tmpValue = GetProperty(aPropertyName, aDefault.ToString());
      return (tmpValue.ToLower() == "true");
    }
    #endregion

    #region GetIntProperty
    /// <summary>
    /// Returns the int value for the specified property or the default value, if
    /// the property is not set or if it cannot be converted to an int.
    /// </summary>
    /// <param name="aPropertyName"></param>
    /// <param name="aDefaultValue"></param>
    /// <returns></returns>
    public int GetIntProperty(string aPropertyName, int aDefaultValue)
    {
      string tmpValue = GetProperty(aPropertyName, null);
      if (tmpValue == null || tmpValue.Trim() == "")
      {
        return aDefaultValue;
      }
      try
      {
        return int.Parse(tmpValue);
      }
      catch (Exception e)
      {
        PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_LIB, "GetProperty(" + aPropertyName + ")", e);
        return aDefaultValue;
      }
    }
    #endregion

    #region GetProperty
    /// <summary>
    /// Returns the value of the specified property or null if the either the property file does not 
    /// exist or does not contain the specified property.
    /// </summary>
    /// <param name="aPropertyName"></param>
    /// <returns></returns>
    public string GetProperty(string aPropertyName)
    {
      return GetProperty(aPropertyName, null);
    }

    /// <summary>
    /// Returns the property value for the specified property name. If the property file either
    /// does not exist or does not contain the property the default value is returned instead.
    /// </summary>
    /// <param name="aPropertyName"></param>
    /// <param name="aDefault"></param>
    /// <returns></returns>
    public string GetProperty(string aPropertyName, string aDefault)
    {
      if (myProperties.ContainsKey(aPropertyName))
      {
        return myProperties[aPropertyName];
      }
      return aDefault;
    }
    #endregion

    #region UserConfiguration
    private UserConfiguration()
    {
      string tmpCodeBase = Path.GetDirectoryName(GetType().Assembly.CodeBase);
      Uri tmpUri = new Uri(GetType().Assembly.CodeBase);

      tmpCodeBase = Path.GetDirectoryName(tmpUri.LocalPath);
      string tmpFileName = Path.Combine(tmpCodeBase, "pdcconfig.properties");

      if (!File.Exists(tmpFileName))
      {
        PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_LIB, "No pdcconfig.properties found");
        return;
      }
      using (StreamReader tmpReader = new StreamReader(tmpFileName))
      {
        string tmpLine;
        while ((tmpLine = tmpReader.ReadLine()) != null)
        {
          tmpLine = tmpLine.Trim();
          if (tmpLine.StartsWith("#"))
          {
            continue;
          }
          int tmpIdx = tmpLine.IndexOf('=');
          if (tmpIdx > 0)
          {
            string tmpPropertyName = tmpLine.Substring(0, tmpIdx);
            string tmpPropertyValue = tmpLine.Substring(tmpIdx + 1);
            tmpPropertyName = tmpPropertyName.Trim();
            tmpPropertyValue = tmpPropertyValue.Trim();
            if (tmpPropertyName == "")
            {
              continue;
            }
            if (myProperties.ContainsKey(tmpPropertyName))
            {
              myProperties[tmpPropertyName] = tmpPropertyValue;
              PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_LIB, "Overriding previously defined property " + tmpPropertyName + " with value " + tmpPropertyValue);
            }
            else
            {
              myProperties.Add(tmpPropertyName, tmpPropertyValue);
              PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_LIB, "Found property " + tmpPropertyName + " with value " + tmpPropertyValue);
            }
          }
        }
      }
    }
    #endregion

    #endregion

    #region properties

    #region TheConfiguration
    /// <summary>
    /// Returns the initialized UserConfiguration singleton
    /// </summary>
    public static UserConfiguration TheConfiguration
    {
      get
      {
        lock(LOCK)
        {
          if (myConfiguration == null)
          {
            myConfiguration = new UserConfiguration();
          }
          return myConfiguration;
        }
      }
    }
    #endregion

    #endregion
  }
}
