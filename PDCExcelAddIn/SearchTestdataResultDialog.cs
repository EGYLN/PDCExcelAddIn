using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    [ComVisible(false)]
  public partial class SearchTestdataResultDialog : Form
  {
    private ResultStatus resultStatus = ResultStatus.CANCEL;

    [ComVisible(false)]
    public enum ResultStatus
    {
      CANCEL,
      REPLACE,
      REPLACEWITHOUTMEASUREMENTS
    }

    #region constructor
    public SearchTestdataResultDialog(bool showLoadMeasurementsCheckbox, bool showOmitMeasurementsCheckbox)
    {
      InitializeComponent();
      this.chkLoadMeasurements.Visible = showLoadMeasurementsCheckbox;
      this.createMeasurementTables.Visible = showOmitMeasurementsCheckbox;
    }
    #endregion

    #region methods

    #region Show
    public static ResultStatus Show(Form aWindowOwner, int aResultSize, bool loadMeasurements, bool omitMeasurements)
    {
      SearchTestdataResultDialog tmpDialog = new SearchTestdataResultDialog(loadMeasurements,omitMeasurements);
      tmpDialog.messageLabel.Text = string.Format(Properties.Resources.MSG_SEARCH_RESULT_TEXT, aResultSize);
      tmpDialog.ShowDialog(aWindowOwner);
      return tmpDialog.resultStatus;
    }
    #endregion

    #endregion

    #region events

    #region cancelButton_Click
    private void cancelButton_Click(object sender, EventArgs e)
    {
      resultStatus = ResultStatus.CANCEL;
      Dispose();
    }
    #endregion

    #region okButton_Click
    private void okButton_Click(object sender, EventArgs e)
    {
        if (this.chkLoadMeasurements.Visible && !this.chkLoadMeasurements.Checked)
        {
            resultStatus = ResultStatus.REPLACEWITHOUTMEASUREMENTS;
        }
        else if (createMeasurementTables.Visible && createMeasurementTables.Checked)
        {
            resultStatus = ResultStatus.REPLACEWITHOUTMEASUREMENTS;
        }
        else
        {
            resultStatus = ResultStatus.REPLACE;
        }
      Dispose();
    }
    #endregion

    private void SearchTestdataResultDialogLoad(object sender, EventArgs e)
    {
        ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SearchTestdataResultDialog));
        ToolTip toolTip = new ToolTip();
        toolTip.AutoPopDelay = 5000;
        toolTip.InitialDelay = 1000;
        toolTip.ReshowDelay = 500;
        toolTip.ShowAlways = true;
        toolTip.SetToolTip(createMeasurementTables, resources.GetString("createMeasurementTables.Tooltip"));
    }

    #endregion

  }
}
