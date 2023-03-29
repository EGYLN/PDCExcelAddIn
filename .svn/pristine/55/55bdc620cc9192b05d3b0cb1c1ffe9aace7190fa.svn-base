using System;
using System.Collections;
using System.Collections.Generic;
using BBS.ST.BHC.BSP.PDC.Lib.Properties;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    [Serializable]
  /// <summary>
  /// Contains information about the logged in user
  /// </summary>
  public class UserInfo
  {
    string myCwid;

    #region methods

    #region IsAdmin
    /// <summary>
    /// Returns true if the user is in the PDC Admin group
    /// </summary>
    /// <returns></returns>
    public bool IsAdmin()
    {
      return IsUserInRole(Settings.Default.ROLE_PDC_ADMIN);
    }
    #endregion

    #region IsAuthor
    /// <summary>
    /// Returns true if the user is in the PDC TD Author group
    /// </summary>
    /// <returns></returns>
    public bool IsAuthor()
    {
      return IsUserInRole(Settings.Default.ROLE_PDC_AUTHOR);
    }
    #endregion

    #region IsCurator
    /// <summary>
    /// Returns true if the user is in the PDC Curator group
    /// </summary>
    /// <returns></returns>
    public bool IsCurator()
    {
      return IsUserInRole(Settings.Default.ROLE_PDC_CURATOR);
    }
    #endregion

    #region IsDataProvider
    /// <summary>
    /// Returns true if the user is in the PDC Data Provder group
    /// </summary>
    /// <returns></returns>
    public bool IsDataProvider()
    {
      return IsUserInRole(Settings.Default.ROLE_PDC_DATAPROVIDER);
    }
    #endregion
    /// <summary>
    /// Returns true if the user is in the PDC Data Provder group
    /// </summary>
    /// <returns></returns>
    public bool IsPDCUser => IsUserInRole(Settings.Default.ROLE_PDC_USER);
    public bool IsIcbUser => IsUserInRole(Settings.Default.ROLE_PDC_USER_ICB);

    /// <summary>
    /// Returns true if the user is in the specified role
    /// </summary>
    /// <param name="aRole"></param>
    /// <returns></returns>
    public bool IsUserInRole(string aRole)
    {
      return aRole != null && Roles.Contains(aRole.ToUpper());
    }

    #endregion

    public bool IsIcbOnlyUser => Roles.Contains(Settings.Default.ROLE_PDC_USER_ICB) && !Roles.Contains(Settings.Default.ROLE_PDC_USER);
    #region properties

    #region Cwid
    /// <summary>
    /// Property for the user id
    /// </summary>
    public string Cwid
    {
      get => myCwid;
      set
      {
        myCwid = value;
        PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_LIB, "Setting CWID to " + value);
        if (myCwid != null)
        {
          //check for domain info and remove it
          int tmpIndex = -1;
          for (int i = myCwid.Length - 1; i >= 0; i--)
          {
            if (!Char.IsLetterOrDigit(myCwid[i]))
            {
              tmpIndex = i;
              break;
            }
          }
          if (tmpIndex >= 0 && tmpIndex < myCwid.Length - 1)
          {
            myCwid = myCwid.Substring(tmpIndex + 1);
          }
        }
      }
    }
    #endregion

    #region Roles

    public int NrOfRoles => Roles?.Count ?? 0;
    public List<string> Roles { get; set; } = new List<string>();

    #endregion

    public bool HasRoles()
    {
      return Roles.Count > 0;
    }

    #endregion
  }
}
