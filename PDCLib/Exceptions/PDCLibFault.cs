using System;
using BBS.ST.Base.STException;
using BBS.ST.Base.Translation;

namespace BBS.ST.BHC.BSP.PDC.Lib.Exceptions
{
  /// <summary>
  /// Base class for PDC library exceptions
  /// </summary>
  public class PDCLibFault: ProgramFault
  {
    #region constructors
    static PDCLibFault()
    {
      SetTranslator(new ResxTranslator(typeof(PDCFaultMessage).FullName, typeof(PDCFaultMessage).Assembly));
    }

    /// <summary>
    /// Initializes the exception for the specified message
    /// </summary>
    /// <param name="aMessage"></param>
    public PDCLibFault(PDCFaultMessage aMessage) : base(aMessage.ToString())
    {
    }

    /// <summary>
    /// Initializes the exception for the specified message and arguments
    /// </summary>
    /// <param name="aMessage">A Well-known exception type</param>
    /// <param name="anArgumentList">Optional arguments which will be added to the message text</param>
    public PDCLibFault(PDCFaultMessage aMessage, object[] anArgumentList) : base(aMessage.ToString(), anArgumentList)
    {
    }
    #endregion

    #region methods
    /// <summary>
    /// Call this method from a unit test to ensure all message codes 
    /// are defined in the translation source.
    /// </summary>
    public static void CheckMessageCodes()
    {
      string[] tmpCodeNames = Enum.GetNames(typeof(PDCFaultMessage));
      CheckMessageCodes(tmpCodeNames);
    }
    #endregion
  }
}
