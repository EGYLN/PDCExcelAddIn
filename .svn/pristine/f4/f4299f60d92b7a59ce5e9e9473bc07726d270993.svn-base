using System;
using System.Collections;
using System.DirectoryServices.Protocols;
using System.Web.Services.Protocols;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using EX=BBS.ST.Base.STException;
namespace BBS.ST.BHC.BSP.PDC.Lib.Exceptions
{
  /// <summary>
  /// This exception is thrown if the user login fails. 
  /// </summary>
  public class LoginException:PDCLibFault
  {
    private int myErrorCode;
    private IDictionary myData;
    private string myMessage;

    #region constructor
    /// <summary>
    /// Initializes a LoginExcepion for a specific failure.
    /// </summary>
    /// <param name="aFaultMessage">Specifies the error message</param>
    /// <param name="anErrorCode">Error code of the original exception</param>
    /// <param name="aMessage">Message of the original exception</param>
    /// <param name="anInfoMap">May contain additional informations</param>
    public LoginException(PDCFaultMessage aFaultMessage, int anErrorCode, string aMessage, IDictionary anInfoMap) :
      base(aFaultMessage, new object[] {anErrorCode, aMessage})
    {
      myErrorCode = anErrorCode;
      myMessage = aMessage;
      myData = anInfoMap;
    }
    #endregion

    #region methods

    #region CreateFromLDAP
    /// <summary>
    /// Called if the LDAP bind fails. Always throws a LoginException.
    /// </summary>
    /// <param name="anExeption"></param>
    public static void CreateFromLDAP(LdapException anExeption)
    {
      string tmpServerMessage = anExeption.ServerErrorMessage.ToUpper();
      PDCFaultMessage tmpMessage;
      if (tmpServerMessage.Contains("DATA 52E"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_INVALID_CREDENTIALS;
      }
      else if (tmpServerMessage.Contains("DATA 530"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_CURRENTLY_NOT_PERMITTED;
      }
      else if (tmpServerMessage.Contains("DATA 532"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_PASSWORD_EXPIRED;
      }
      else if (tmpServerMessage.Contains("DATA 533"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_ACCOUNT_DISABLED;
      }
      else if (tmpServerMessage.Contains("DATA 701"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_ACCOUNT_EXPIRED;
      }
      else if (tmpServerMessage.Contains("DATA 773"))
      {
        tmpMessage = PDCFaultMessage.LOGIN_MUST_CHANGE_PASSWORD;
      }
      else
      {
        tmpMessage = PDCFaultMessage.LOGIN_UNKNOWN_FAULT;
      }
      throw new LoginException(tmpMessage, anExeption.ErrorCode, anExeption.Message, anExeption.Data);
    }
    #endregion

    #region CreateFromWSUM
    internal static void CreateFromWSUM(Exception e)
    {
      PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, e.Message, e);
      if (e is SoapException)
      {
        SoapException tmpSE = (SoapException) e;
        if (tmpSE.Detail != null && tmpSE.Detail.InnerText != null && tmpSE.Detail.InnerText.IndexOf("ApplicationFault", 1) > 0)
        {
          LoginException tmpLE = new LoginException(PDCFaultMessage.LOGIN_FAILED, 99, e.Message, null);
          tmpLE.HelpLink = tmpSE.HelpLink;
          throw tmpLE;
        }
      }
      throw e;
    }
    #endregion

    #endregion
  }
}
