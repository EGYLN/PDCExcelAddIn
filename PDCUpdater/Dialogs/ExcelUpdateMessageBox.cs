using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace BBS.ST.BHC.AppUpdater.Dialogs
{
  public partial class ExcelUpdateMessageBox : Form
  {
    public ExcelUpdateMessageBox()
    {
      InitializeComponent();
    }

    public static DialogResult Show(MessageBoxButtons buttons, String title, String message, String okButtonText, String cancelButtonText, String killButtonText, Boolean displayKillButton)
    {
      ExcelUpdateMessageBox dlg = new ExcelUpdateMessageBox();

      dlg.myBtnOk.Text = okButtonText;
      dlg.myBtnCancel.Text = cancelButtonText;
      dlg.myBtnKillExcels.Text = killButtonText;
      dlg.myLblMessage.Text = message;
      dlg.Text = title;

      switch (buttons)
      {
        case MessageBoxButtons.OK:
          dlg.AssignOkButton();
          break;

        case MessageBoxButtons.OKCancel:
          //dlg.AssignOkCancelButton();
          break;

        default:
          throw new NotSupportedException("Only \"OK\" and \"OkCancel\" buttons are supported.");
      }

      dlg.AssignTeminateExcelButton(displayKillButton);

      return dlg.ShowDialog();
    }

    private void AssignTeminateExcelButton(Boolean visible)
    {
      this.myBtnKillExcels.Visible = visible;
    }

    private void AssignOkButton()
    {
      Int32 width;
      Int32 x;

      
      this.myBtnOk.Visible = true;
      this.myBtnCancel.Visible = false;

      width = this.myBtnOk.Width;

      x = this.Width - width - 6;

      this.myBtnOk.Left = x;
    }

    private void AssignOkCancelButton()
    {
      Int32 width;
      Int32 x;

      
      this.myBtnOk.Visible = true;
      this.myBtnCancel.Visible = true;

      width = this.myBtnOk.Width;

      this.myBtnCancel.Width = width;

      //x = (this.Width - width * 2 + 6) / 2;

      x = this.Width - width - 6;

      this.myBtnCancel.Left = x;
      
      x = x - width - 6;

      this.myBtnOk.Left = x;
    }

    private void myBtnOk_Click(object sender, EventArgs e)
    {
      this.DialogResult = DialogResult.OK;
    }

    private void myBtnCancel_Click(object sender, EventArgs e)
    {
      this.DialogResult = DialogResult.Cancel;
    }

    private void myBtnKillExcels_Click(object sender, EventArgs e)
    {
      this.DialogResult = DialogResult.Retry;
    }
  }
}
