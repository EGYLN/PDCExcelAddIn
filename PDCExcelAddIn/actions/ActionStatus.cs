using System;
using System.Runtime.InteropServices;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// The ActionStatus is used to collect information about the execution status.
    /// The status is evaluated by PDCAction in non-interactive mode
    /// </summary>
    [ComVisible(false)]
  public class ActionStatus
  {
    Lib.PDCMessage[] myMessages;
    Exception myFailure;

    internal ActionStatus()
    {
      myFailure = null;
      myMessages = null;
    }

    internal ActionStatus(Lib.PDCMessage aMessage)
    {
      myFailure = null;
      myMessages = new Lib.PDCMessage[] { aMessage };
    }

    internal ActionStatus(Exception anException, Lib.PDCMessage[] theMessages)
    {
      myFailure = anException;
      myMessages = theMessages;
    }

    /// <summary>
    /// Creates a status for an exception
    /// </summary>
    /// <param name="anException"></param>
    internal ActionStatus(Exception anException)
    {
      myFailure = anException;
      myMessages = null;
    }

    /// <summary>
    /// Creates a status with the specified messages.
    /// </summary>
    /// <param name="theMessages"></param>
    internal ActionStatus(Lib.PDCMessage[] theMessages)
    {
      myFailure = null;
      myMessages = theMessages;
    }

    public Lib.PDCMessage[] Messages
    {
      get
      {
        return myMessages;
      }
      internal set
      {
        myMessages = value;
      }
    }
    public Exception Failure
    {
      get
      {
        return myFailure;
      }
      internal set
      {
        myFailure = value;
      }
    }
  }
}
