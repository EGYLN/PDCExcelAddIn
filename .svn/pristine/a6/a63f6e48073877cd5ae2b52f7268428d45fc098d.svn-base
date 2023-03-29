using System;
using System.DirectoryServices.Protocols;
using System.Net;
using System.Security.Principal;
using System.Web.Services.Protocols;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
using BBS.ST.BHC.BSP.PDC.Lib.Properties;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using UM = BBS.ST.BHC.BSP.PDC.Lib.WsCustomerManagement;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// Handles the login/logout procedure for the PDC client/library
  /// </summary>
  class LoginService
  {
    private UserInfo myUserInfo;

    #region constructor
    public LoginService()
    {
      PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_LIB, "LoginService created");
    }
    #endregion

    #region methods

    #region Login
    public void Login(string aUsername, string aPassword, Boolean useWinLogin)
    {
            
      UserInfo = new UserInfo();
      PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Logging into the system");
      if (useWinLogin)
      {
        WindowsIdentity tmpIdentity = WindowsIdentity.GetCurrent();
        UserInfo.Cwid = tmpIdentity.Name;
      }
      else
      {
                Console.WriteLine(System.Net.ServicePointManager.SecurityProtocol);
                UserInfo.Cwid = aUsername;
      }

      if (Settings.Default.UsePHUser)
      {
        LoginPHUser(aUsername, aPassword, useWinLogin);
      }
      else
      {
        LoginLDAP(aUsername, aPassword, useWinLogin);
      }

      if (UserInfo.Roles.Count == 0)
      {
      }
      else if (PDCLogger.TheLogger.IsDebugEnabled) 
      {
        foreach (string tmpRole in UserInfo.Roles)
        {
          PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, UserInfo.Cwid + " has role " + tmpRole);
        }
      }
    }
    #endregion

    #region LoginLDAP
    private void LoginLDAP(string aUsername, string aPassword, Boolean useWinLogin)
    {
      try
      {
        LdapConnection tmpConnection = new LdapConnection(Settings.Default.LDAPServer);
        NetworkCredential tmpCred;
        if (!useWinLogin)
        {//authentication
          tmpCred = new NetworkCredential(aUsername, aPassword);
          UserInfo.Cwid = aUsername;
          tmpConnection.Bind(tmpCred);
        }
        //authorization
        tmpCred = new NetworkCredential("testbinduser", "Initial1");
        tmpConnection.Bind(tmpCred);
        DirectoryRequest tmpRequest = new SearchRequest("CN=" + UserInfo.Cwid + "," + Settings.Default.LDAP_OU_PEOPLE, "(objectClass=*)", SearchScope.Base, "memberOf");
        try
        {
          SearchResponse tmpResponse = (SearchResponse)tmpConnection.SendRequest(tmpRequest);
          foreach (SearchResultEntry tmpEntry in tmpResponse.Entries)
          {
            foreach (DirectoryAttribute tmpAttribute in tmpEntry.Attributes.Values)
            {
              if (tmpAttribute.Name != "memberOf")
              {
                continue;
              }
              object[] tmpValues = tmpAttribute.GetValues(typeof(string));
              if (tmpValues != null)
              {
                for (int i = tmpValues.GetLowerBound(0); i <= tmpValues.GetUpperBound(0); i++)
                {
                  string tmpGroup = "" + tmpValues[i];
                  if (tmpGroup.Equals(Settings.Default.LDAP_PDC_USER_GROUP, StringComparison.InvariantCultureIgnoreCase))
                  {
                    UserInfo.Roles.Add(Settings.Default.ROLE_PDC_USER);
                  }
                  else if (tmpGroup.Equals(Settings.Default.LDAP_PDC_DATAPROVIDER_GROUP, StringComparison.InvariantCultureIgnoreCase))
                  {
                    UserInfo.Roles.Add(Settings.Default.ROLE_PDC_DATAPROVIDER);
                  }
                  else if (tmpGroup.Equals(Settings.Default.LDAP_PDC_CURATOR_GROUP, StringComparison.InvariantCultureIgnoreCase))
                  {
                    UserInfo.Roles.Add(Settings.Default.ROLE_PDC_CURATOR);
                  }
                  else if (tmpGroup.Equals(Settings.Default.LDAP_PDC_AUTHOR_GROUP, StringComparison.InvariantCultureIgnoreCase))
                  {
                    UserInfo.Roles.Add(Settings.Default.ROLE_PDC_AUTHOR);
                  }
                  else if (tmpGroup.Equals(Settings.Default.LDAP_PDC_ADMIN_GROUP, StringComparison.InvariantCultureIgnoreCase))
                  {
                    UserInfo.Roles.Add(Settings.Default.ROLE_PDC_ADMIN);
                  }
                }
              }
            }
          }
        }
        catch (DirectoryOperationException doE)
        {
          PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_LIB, "LDAP exception",doE);
        }
      }
      catch (LdapException anException)
      {
        PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_LIB, "LDAP exception", anException);
      }

    }
    #endregion

    #region LoginPHUser
    private void LoginPHUser(string aUsername, string aPassword, Boolean useWinLogin)
    {
      try
      {
        UM.WsCustomerManagementService tmpService = new UM.WsCustomerManagementService();
        WebServiceClientInitializer tmpInitializer = new WebServiceClientInitializer();
        string tmpUrl = Settings.Default.WsCustomerManagement_URL;
        tmpUrl = UserConfiguration.TheConfiguration.GetProperty(UserConfiguration.PROP_WS_CUSTOMERMANAGEMENT_URL, tmpUrl);
        tmpInitializer.InitializeWebServiceClient(tmpService, tmpUrl, Settings.Default.WsCustomerManagement_PDCUser,
          Settings.Default.WsCustomerManagement_License);
        if (!useWinLogin)
        {
          try
          {
            UM.Principal[] tmpPrincipal = tmpService.getAuthentication(aUsername, aPassword);
            UserInfo.Cwid = aUsername;
          }
          catch (SoapException sE) {
            PDCLogger.TheLogger.LogException("", "Login Failed", sE);
            throw (sE);
          }
          catch (Exception e)
          {
            PDCLogger.TheLogger.LogException("", "Login Failed", e);
            throw(e);
          }
        }
        if (UserInfo.Cwid == null)
        {
          return;
        }
        UM.Customer tmpCustomerInfo = tmpService.getCustomerInfo(UserInfo.Cwid, new[] { "PDC" });
        UM.CustomerRole[] tmpRoles = tmpCustomerInfo.roles;
        if (tmpRoles == null)
        {
          return;
        }
        foreach (UM.CustomerRole tmpRole in tmpRoles) {
          if (tmpRole.name != null)
          {
            UserInfo.Roles.Add(tmpRole.name.ToUpper());
          }
        }
      }
      catch (Exception e)
      {
          //Zu Testzwecken, um nicht selbst für PDC registriert sein zu müssen.
          //UserInfo.Roles.Add("PDC_ADMIN", "PDC_ADMIN");
          //UserInfo.Roles.Add("PDC_TD_AUTHOR", "PDC_TD_AUTHOR");
          //UserInfo.Roles.Add("PDC_DP", "PDC_DP");
          //UserInfo.Roles.Add("PDC_USER", "PDC_USER");

        LoginException.CreateFromWSUM(e);
      }
    }
    #endregion

    #region Logout
    public void Logout()
    {
      UserInfo = null;
    }
    #endregion

    #endregion

    #region properties

    #region UserInfo
    public UserInfo UserInfo
    {
      get
      {
        return myUserInfo;
      }
      internal set
      {
        myUserInfo = value;
      }
    }
   
    #endregion

    #endregion
  }
}
