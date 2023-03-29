using System;
using System.Runtime.InteropServices;
using BBS.ST.Base.STException;
using BBS.ST.Base.Translation;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Exceptions
{
    [ComVisible(false)]
  public class PDCExcelAddInFault : ProgramFault
  {
    static PDCExcelAddInFault()
    {
      ProgramFault.SetTranslator(new ResxTranslator(typeof(PDCExcelAddInFaultMessage).FullName, typeof(PDCExcelAddInFaultMessage).Assembly));
    }

    /// <summary>
    /// Initializes the exception for the specified message
    /// </summary>
    /// <param name="aMessage"></param>
    public PDCExcelAddInFault(PDCExcelAddInFaultMessage aMessage) : base(aMessage.ToString())
    {
    }

    /// <summary>
    /// Initializes the exception for the specified message and arguments
    /// </summary>
    /// <param name="aMessage">A Well-known exception type</param>
    /// <param name="anArgumentList">Optional arguments which will be added to the message text</param>
    public PDCExcelAddInFault(PDCExcelAddInFaultMessage aMessage, object[] anArgumentList) : base(aMessage.ToString(), anArgumentList)
    {
    }

    /// <summary>
    /// Call this method from a unit test to ensure all message codes 
    /// are defined in the translation source.
    /// </summary>
    public static void CheckMessageCodes()
    {
      string[] tmpCodeNames = Enum.GetNames(typeof(PDCExcelAddInFaultMessage));
      ProgramFault.CheckMessageCodes(tmpCodeNames);
    }
  }
}
