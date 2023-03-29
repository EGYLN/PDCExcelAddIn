using WS=BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;

namespace BBS.ST.BHC.BSP.PDC.Lib.Exceptions
{
  /// <summary>
  /// A server error from the db represented by a PDC message
  /// </summary>
  public class ServerFailure:PDCLibFault
  {
    #region constructor
    /// <summary>
    /// Server side failure
    /// </summary>
    /// <param name="aPDCMessage">The server message from the db</param>
    public ServerFailure(WS.PDCMessage aPDCMessage) : base(PDCFaultMessage.SERVER_FAILURE, new object[] {aPDCMessage.message,aPDCMessage.messageType.name})
    {
    }

    public ServerFailure(PDCFaultMessage faulMessage) :base(faulMessage) {}
    #endregion
  }
}
