using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
  [ComVisible(false)]
  public partial class PDCLoginForm : Form
  {
    private bool cancelled = true;

    public PDCLoginForm()
    {
      InitializeComponent();
    }

    public bool ShowLogin()
    {
      CenterToParent();
      ShowDialog();
      return cancelled;
    }

    private void Login(object sender, EventArgs args)
    {
      string tmpUsername = cwidField.Text;
      string tmpPassword = passwordField.Text;
      try
      {
          if (tmpUsername == null || tmpPassword == null || string.Empty == tmpUsername || string.Empty == tmpPassword)
          {
              MessageBox.Show(this, Properties.Resources.MSG_CWID_PWD_MISSING, Properties.Resources.MSG_LOGIN_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Error);
              return;
          }
        Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
        tmpService.Login(tmpUsername, tmpPassword, false);
        if (!tmpService.UserInfo.HasRoles()) 
         throw new LoginException(PDCFaultMessage.LOGIN_FAILED, 99, null, null);
        
        
      }
      catch (Exception e)
      {
        ExceptionHandler.TheExceptionHandler.handleException(e, this);
        return;
      }
      cancelled = false;
      Dispose();
    }

    private void Cancelled(object sender, EventArgs e)
    {
      cancelled = true;
      Dispose();
    }
  }
}
