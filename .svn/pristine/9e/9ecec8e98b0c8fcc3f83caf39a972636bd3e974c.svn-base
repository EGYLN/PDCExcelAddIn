using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;

namespace BBS.ST.BHC.PDC.AutoUpdater.IO
{
  /// <summary>
  ///   The SettingsFile provides the functionallity for reading the settings of a settings file.
  /// </summary>
  public class SettingsFile
  {
    private String  myFile;
    private Char    mySeparator;


    private static String MySettingsFile = "BBS.AutoUpdater.Settings.ini";
    private static SettingsFile MyDefaultSettings = null;

    // .ctor
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region .ctor

    /// <summary>
    ///   Initializes a new instance of the settings file.
    /// </summary>
    /// <param name="file">
    ///   The name of the settings file.
    /// </param>
    /// <param name="separator">
    ///   The separator.
    /// </param>
    public SettingsFile(String file, Char separator)
    {
      this.myFile = file;
      this.mySeparator = separator;
    }

    #endregion


    // Methods
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region GetKey

    /// <summary>
    ///   Gets the key of the specified line that contains the key-value pair.
    /// </summary>
    /// <param name="line">
    ///   The line that contains the key-value pair.
    /// </param>
    /// <returns>
    ///   The key of the specified line that contains the key-value pair.
    /// </returns>
    private String GetKey(String line)
    {
      Int32 index;


      index = line.IndexOf(this.mySeparator);

      return line.Substring(0, index).Trim().ToLower();
    }

    #endregion

    #region GetSettings

    /// <summary>
    ///   Gets a dictionary that contains the settings of the specified settings file or null
    ///   if the specified settings file does not exist.
    /// </summary>
    /// <param name="file">
    ///   The file that contains the settings to read.
    /// </param>
    /// <returns>
    ///   The settings of the specified settings file or null
    ///   if the specified settings file does not exist.
    /// </returns>
    /// <remarks>
    ///   The keys will be convertedt to lower case.
    /// </remarks>
    public Dictionary<String, String> GetSettings()
    {
      String[] settings;
      String value;
      String key;
      Dictionary<String, String> result;


      if (!File.Exists(this.Location)) return null;

      
      settings = File.ReadAllLines(this.Location);

      result = new Dictionary<String,String>(32);


      foreach (String line in settings)
      {
        if (SettingsFile.IsComment(line)) continue;

        if (!line.Contains(this.mySeparator.ToString())) continue;


        key = this.GetKey(line);

        value = this.GetValue(line);


        if (String.IsNullOrEmpty(key) ||
            String.IsNullOrEmpty(value)) continue;


        if (result.ContainsKey(key))
        {
          result[key] = value;

          continue;
        }

        result.Add(key, value);
      }
      
      return result;
    }

    #endregion

    #region GetValue

    /// <summary>
    ///   Gets the value of the specified line that contains the key-value pair.
    /// </summary>
    /// <param name="line">
    ///   The line that contains the key-value pair.
    /// </param>
    /// <returns>
    ///   The value of the specified line that contains the key-value pair.
    /// </returns>
    private String GetValue(String line)
    {
      Int32 index;


      index = line.IndexOf(this.mySeparator);

      return line.Substring(index + 1).Trim();
    }

    #endregion


    #region IsComment

    /// <summary>
    ///   Determines if the specified line is a comment line.
    /// </summary>
    /// <param name="line">
    ///   The line to check.
    /// </param>
    /// <returns>
    ///   True if the specified line is a comment line (starts with '#'); otherwise false.
    /// </returns>
    private static Boolean IsComment(String line)
    {
      return line.TrimStart().StartsWith("#");
    }

    #endregion


    // Properties
    // ////////////////////////////////////////////////////////////////////////////////////////////

    #region Default

    /// <summary>
    ///   Gets the settings file that contains the default settings.
    /// </summary>
    public static SettingsFile Default
    {
      get
      {
        if (MyDefaultSettings != null) return MyDefaultSettings;

        SettingsFile.MyDefaultSettings = new SettingsFile(MySettingsFile, ':');

        return SettingsFile.MyDefaultSettings;
      }
    }

    #endregion


    #region Location

    /// <summary>
    ///   Gets the location of the settings file, including path and file name.
    /// </summary>
    public String Location
    {
      get
      {
        String location;


        location = Path.GetDirectoryName(typeof(SettingsFile).Assembly.Location);

        location = Path.Combine(location, this.myFile);


        return location;
      }
    }

    #endregion


    #region Name

    /// <summary>
    ///   Gets the name of the settings file.
    /// </summary>
    public String Name
    {
      get
      {
        return myFile;
      }
    }

    #endregion
  }
}
