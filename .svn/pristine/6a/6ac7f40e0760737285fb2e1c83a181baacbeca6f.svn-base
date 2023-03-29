using System;
using System.Windows.Forms;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Small IWin32Window around the Hwnd of the Excel Application.
    /// so that we can possibly use it for Windows.Forms calls as a proper
    /// parent window for a modal dialogs.
    /// </summary>
    class ExcelHwndWrapper : IWin32Window
  {
    #region IWin32Window Members

    public IntPtr  Handle
    {
      get
      {
        return new IntPtr(Globals.PDCExcelAddIn.Application.Hwnd);
      }
    }

    #endregion
  }
}
