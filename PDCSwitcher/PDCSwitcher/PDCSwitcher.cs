using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Threading;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
namespace PDCSwitcher
{
  public partial class PDCSwitcher : Form
  {
    private bool disableEvents = false;
    private const String LOAD_BEHAVIOUR_PATH_CU = "HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Office\\Excel\\Addins\\PDCExcelAddIn";
    private const String LOAD_BEHAVIOUR_PATH_LM = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Office\\Excel\\Addins\\PDCExcelAddIn";
    public PDCSwitcher()
    {
      InitializeComponent();
     
    }
    private void SetStatus(bool changePDCAddinSetting)
    {
      if (disableEvents) return;
      Cursor = Cursors.WaitCursor;
      this.Enabled = false;
      try
      {
        object loadBehavior = 1;
        loadBehavior = Microsoft.Win32.Registry.GetValue(LOAD_BEHAVIOUR_PATH_CU, "LoadBehavior", 1);
        
        if (loadBehavior == null) {
          loadBehavior =  Microsoft.Win32.Registry.GetValue(LOAD_BEHAVIOUR_PATH_LM, "LoadBehavior", 1);
          Microsoft.Win32.Registry.SetValue(LOAD_BEHAVIOUR_PATH_CU, "LoadBehavior", loadBehavior); 
        }
        if (changePDCAddinSetting)
        {
          Microsoft.Win32.Registry.SetValue(LOAD_BEHAVIOUR_PATH_CU, "LoadBehavior", myRBEnabled.Checked ? 3 : 1); 
        } else {
          disableEvents = true;
          myRBEnabled.Checked = (Int32) loadBehavior == 3;
          myRbDisabled.Checked = !myRBEnabled.Checked;
          disableEvents = false;
        }
      } catch (Exception e)      {
          MessageBox.Show("Error setting status:\n" + e.ToString() + "\n" + e.Message  , "Error");
      } finally {
        Cursor = Cursors.Default;
        this.Enabled = true;
        Application.DoEvents();
      }
          
         
    }

  private static bool findPDCExcelAddin()
  {
    Excel.Application myexcelApp = null;
    IEnumerator tmpComAddIns = null;
    Office.COMAddIn tmpAddIn = null;
    bool tmpPDCExcelAddinFound = false;
    try {
      myexcelApp = new Excel.Application();
      tmpComAddIns = myexcelApp.COMAddIns.GetEnumerator();
      while (!tmpPDCExcelAddinFound && tmpComAddIns.MoveNext())
      {
        tmpAddIn = (Office.COMAddIn)tmpComAddIns.Current;
        if (tmpAddIn.ProgId.Equals("PDCExcelAddIn"))
        {
          tmpPDCExcelAddinFound = true;
        }
        Marshal.ReleaseComObject(tmpAddIn);
        tmpAddIn = null;
      }
    } finally {
      if (tmpAddIn != null) {    
        Marshal.ReleaseComObject(tmpAddIn);
        tmpAddIn = null;
      }    
      tmpComAddIns = null;
      Marshal.ReleaseComObject(myexcelApp);
      myexcelApp = null;
    }
    return tmpPDCExcelAddinFound;
  }
 
    private void button2_Click(object sender, EventArgs e)
    {
      Application.Exit();
    }


    private void myRbDisabled_CheckedChanged(object sender, EventArgs e)
    {
      lblStatus.Text = "Changing Status..."; 
      Application.DoEvents();
      SetStatus(true);
      lblStatus.Text = "";
      Application.DoEvents();
    }

    private void PDCSwitcher_Load(object sender, EventArgs e)
    {
      this.Show();
      lblStatus.Text = "Retrieving Status...";
      Application.DoEvents();
      if (!findPDCExcelAddin())
      {
        MessageBox.Show("PDCExcelAddin has not been installed properly", "Warning");
        Application.Exit();

      }
      SetStatus(false);
      lblStatus.Text = "";
      Application.DoEvents();
      
    }
  }
}
