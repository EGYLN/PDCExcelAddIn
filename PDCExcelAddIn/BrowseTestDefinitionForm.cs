using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using BBS.ST.BHC.BSP.PDC.Lib;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    [ComVisible(false)]
  public partial class BrowseTestDefinitionForm : Form
  {
    private ListViewItem mySelectedRow = null;
    private Lib.Testdefinition mySelectedTd = null;
    private TestdefinitionComparer myComparer = new TestdefinitionComparer();
    private List<ListViewItem> myResult;
    private Thread mySearchThread;

    /// <summary>
    /// Status of the worker thread
    /// </summary>
    [ComVisible(false)]
    public enum Status
    {
      /// <summary>
      /// Status after the search for test definition finished
      /// </summary>
      SEARCH_DONE,
      /// <summary>
      /// Status if the dialog is about to be closed
      /// </summary>
      CLOSE
    }

    /// <summary>
    /// Action callback after a test definition was selected
    /// </summary>
    /// <param name="aTD">The selected testdefinition</param>
    /// <param name="noMeasurements">Wether to ignore eventual measurement variables</param>
    [ComVisible(false)]
    public delegate void Callback(Lib.Testdefinition aTD, bool noMeasurements);
    public event Callback callback;

    private Lib.Testdefinition mySearchTemplate;
    /// <summary>
    /// Called from a background thread if it is about to finish.
    /// </summary>
    /// <param name="status">May be a Status or an exception in case of failure</param>
    private delegate void WSStatus(object status);
    private event WSStatus Finished;
    List<Lib.Testdefinition> myTestDefinitionList;

    #region classes

    #region TestdefinitionComparer
    private class TestdefinitionComparer : System.Collections.IComparer
    {
      private int sortColumn = 1;
      private int multiplicator = 1;

      public void sortTable(ListView aListView, int aColumn)
      {
        SortOrder tmpPrev = aListView.Sorting;
        
        if (aColumn == sortColumn)
        {
          if (tmpPrev == SortOrder.None)
          {
            multiplicator = 1;
            aListView.Sorting = SortOrder.Ascending;
          }
          else if (tmpPrev == SortOrder.Ascending)
          {
            multiplicator = -1;
            aListView.Sorting = SortOrder.Descending;
          }
          else
          {
            multiplicator = 1;
            aListView.Sorting = SortOrder.Ascending;
          }
        }
        else
        {
          SortColumn = aColumn;
          multiplicator = 1;
          aListView.Sorting = SortOrder.Ascending;
        }
      }

      public int Compare(object a, object b)
      {
        ListViewItem tmpA = (ListViewItem)a;
        ListViewItem tmpb = (ListViewItem)b;
        Lib.Testdefinition tmpDefA = (Lib.Testdefinition)tmpA.Tag;
        Lib.Testdefinition tmpDefB = (Lib.Testdefinition)tmpb.Tag;
        switch (sortColumn)
        {
          case 0:
            return multiplicator * (tmpDefA.TestNo < tmpDefB.TestNo ? -1 : tmpDefA.TestNo == tmpDefB.TestNo ? 0 : 1);
          case 1:
            return multiplicator *(tmpDefA.TestName.CompareTo(tmpDefB.TestName));
          case 2:
            return multiplicator * (tmpDefA.Version < tmpDefB.Version ? -1 : tmpDefA.Version == tmpDefB.Version ? 0 : 1);
          default:
            return multiplicator * (tmpDefA.TestName.CompareTo(tmpDefB.TestName));
        }
      }

      public int SortColumn
      {
        get
        {
          return sortColumn;
        }
        set
        {
          sortColumn = value;
        }
      }
    }
    #endregion

    #endregion

    #region constructor
    public BrowseTestDefinitionForm()
    {
      InitializeComponent();
      this.InitToolTips();
    }

    private void InitToolTips()
    {
      ToolTip toolTip = new ToolTip();
      toolTip.AutoPopDelay = 100000;
      toolTip.SetToolTip(testNameLabel,Tooltips.TestName);
      toolTip.SetToolTip(testName, Tooltips.TestName);

      toolTip.SetToolTip(selectButton, Tooltips.SelectButton);
      toolTip.SetToolTip(accessibleTestsField, Tooltips.MyAccessibleTests);
      toolTip.SetToolTip(rdbMultipleMeasurement, Tooltips.MultipleMeasurementTable);
      toolTip.SetToolTip(rdbSingleMeasurement, Tooltips.SingleMeasurementTable);

    }

    #endregion

    #region events

    #region AccessibleChanged
    private void AccessibleChanged(object sender, EventArgs e)
    {
      DisplayResult();
    }
    #endregion

    #region BrowseTestDefinitionForm_Load
    private void BrowseTestDefinitionForm_Load(object sender, EventArgs e)
    {
      Finished += new WSStatus(WorkerThreadFinished);
      resultView.ListViewItemSorter = myComparer;
      resultView.Sorting = SortOrder.Ascending;
    }
    #endregion

    #region BrowseTestDefinitionForm_Shown
    private void BrowseTestDefinitionForm_Shown(object sender, EventArgs e)
    {
      this.testName.Focus();
    }
    #endregion

    #region KeyPressed
    private void KeyPressed(object sender, KeyPressEventArgs e)
    {
      if (!char.IsDigit(e.KeyChar)&& !char.IsControl(e.KeyChar))
      {
        e.Handled = true;
      }
    }
    #endregion

    #region noMeasurementsCheckbox_CheckedChanged
    /// <summary>
    ///   Enables or disables the radio buttons for single and multiple measurements.
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void noMeasurementsCheckbox_CheckedChanged(object sender, EventArgs e)
    {
      if (noMeasurementsCheckbox.Checked)
      {
        rdbMultipleMeasurement.Enabled = false;
        rdbSingleMeasurement.Enabled = false;
      }
      else
      {
        rdbMultipleMeasurement.Enabled = true;
        rdbSingleMeasurement.Enabled = true;
      }
    }
    #endregion

    #region SortBy
    private void SortBy(object sender, ColumnClickEventArgs e)
    {
      myComparer.sortTable(resultView, e.Column);
    }
    #endregion

    #region validateNumberField
    private void validateNumberField(object sender, KeyPressEventArgs e)
    {
      if ((e.KeyChar < '0' || e.KeyChar > '9') && !char.IsControl(e.KeyChar))
      {
        e.Handled = true;
      }
    }
    #endregion

    #endregion

    #region methods

    #region BrowseTestDefinitions
    public Lib.Testdefinition BrowseTestDefinitions()
    {
      this.CenterToScreen();
      InitializeCreateWorksheet();
      ShowDialog();
      return mySelectedTd;
    }
    #endregion

    #region Cancel
    /// <summary>
    /// Cancel dialog and possibly active search thread
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void Cancel(object sender, EventArgs e)
    {
      mySelectedRow = null;
      mySelectedTd = null;
      if (mySearchThread != null && mySearchThread.IsAlive)
      {
        try
        {
          mySearchThread.Abort();
        }
        catch (Exception ee)
        {
          Lib.Util.PDCLogger.TheLogger.LogException(Lib.Util.PDCLogger.LOG_NAME_EXCEL, "Aborting Search thread", ee);
        }
      }
      Dispose();
    }
    #endregion

    #region CanUpload
    /// <summary>
    /// Has the user upload privileges for the test definition?
    /// </summary>
    /// <param name="aTD"></param>
    /// <returns></returns>
    private bool CanUpload(Lib.Testdefinition aTD)
    {
      return aTD.HasUploadPrivileges(Globals.PDCExcelAddIn.PdcService.UserInfo);
    }
    #endregion

    #region DisplayResult
    /// <summary>
    /// Displays the search result from the server, including the 
    /// information if the user has upload privileges
    /// </summary>
    private void DisplayResult()
    {
      resultView.Items.Clear();
      if (myResult == null)
      {
        return;
      }
      bool tmpDAEOnly = accessibleTestsField.Checked;
      List<ListViewItem> tmpNewItems = new List<ListViewItem>();
      foreach (ListViewItem tmpItem in myResult)
      {
        Lib.Testdefinition tmpTD = (Lib.Testdefinition)tmpItem.Tag;
        if (!tmpDAEOnly || CanUpload(tmpTD))
        {
          tmpNewItems.Add(tmpItem);
        }
      }
      resultView.Items.AddRange(tmpNewItems.ToArray());
    }
    #endregion

    #region EnableGUI
    /// <summary>
    /// Enables the GUI after a worker thread finished
    /// </summary>
    private void EnableGUI(bool anEnableFlag)
    {
      searchButton.Enabled = anEnableFlag;
      cancelButton.Enabled = anEnableFlag;
      selectButton.Enabled = anEnableFlag;
      this.Cursor = anEnableFlag ? DefaultCursor : Cursors.WaitCursor;
      this.UseWaitCursor = !anEnableFlag;
    }
    #endregion

    #region HandleException
    /// <summary>
    /// Handles the exception from a worker thread
    /// </summary>
    /// <param name="e"></param>
    private void HandleException(Exception e)
    {
      if (!IsDisposed)
      {
        BeginInvoke(new MethodInvoker(delegate()
        {
          ExceptionHandler.TheExceptionHandler.handleException(e, this);
        }));
      }
    }
    #endregion

    #region InitializeCreateWorksheet
    /// <summary>
    /// newWorksheetField is initially unchecked if
    /// the active worksheet is empty
    /// newWorksheetField is disabled if there is 
    /// no active workbook/active worksheet (checked) or
    /// active worksheet is empty (unchecked
    /// </summary>
    private void InitializeCreateWorksheet()
    {
      Excel.Application tmpApp = Globals.PDCExcelAddIn.Application;
    }
    #endregion

    #region InitializeTestDefinition
    /// <summary>
    /// Gets the variables, picklists, ... for the selected test definition.
    /// </summary>
    private void InitializeTestDefinition()
    {
      object tmpStatus = null;
      if (mySelectedTd == null)
      {
        return;
      }
      try
      {
        Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
        tmpService.InitializeTestdefinition(mySelectedTd);
        if (callback != null && callback.GetInvocationList() != null && callback.GetInvocationList().Length > 0)
        {
          if (!noMeasurementsCheckbox.Checked && rdbSingleMeasurement.Checked) mySelectedTd.ShowSingleMeasurement = true;
          callback(mySelectedTd, noMeasurementsCheckbox.Checked);
        }
        tmpStatus = Status.CLOSE;
      }
#pragma warning disable 0168
      catch (Exception e)
      {                
        tmpStatus = e;
      }
      // As the search runs in a worker thread and the result must be sent to the UI, BeginInvoke is necessary, because
      // worker threads must not access UIs. 
      BeginInvoke(new WSStatus(WorkerThreadFinished), new object[] { tmpStatus });
#pragma warning restore 0168
    }
    #endregion

    #region PerformSearch
    private void PerformSearch()
    {
      object tmpStatus = null;

      try
      {
        Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
        myTestDefinitionList = tmpService.FindTestdefinitions(mySearchTemplate);
        tmpStatus = Status.SEARCH_DONE;                
      }
      catch (Exception e)
      {
        tmpStatus = e;
      }
      if (!IsDisposed)
      {
        BeginInvoke(new WSStatus(WorkerThreadFinished), new object[] { tmpStatus });
      }
    }
    #endregion

    #region Search
    private void Search(object sender, EventArgs e)
    {
      //Extract search condition 
      string tmpTestName = testName.Text.Trim();
      string tmpTestNo = testNo.Text.Trim();
      string tmpVersion = versionNo.Text.Trim();

      if ("".Equals(tmpTestName) && "".Equals(tmpTestNo))
      {
        return; //Do not search with empty conditions
      }
      int? tmpNo = null;
      int? tmpVersionNo = null;
      int tmpNoParse = 0;
      int tmpVersionNoParse = 0;

      bool tmpSuccess = int.TryParse(tmpTestNo, out tmpNoParse);
      if (tmpSuccess)
      {
        tmpNo = tmpNoParse;
      }
      tmpSuccess = int.TryParse(tmpVersion, out tmpVersionNoParse);
      if (tmpSuccess)
      {
        tmpVersionNo = tmpVersionNoParse;
      }

      //Perform background search
      mySearchTemplate = new Lib.Testdefinition(null, tmpTestName, tmpNo, tmpVersionNo);
      mySearchThread = new Thread(new ThreadStart(PerformSearch));
      mySearchThread.SetApartmentState(ApartmentState.STA);
      EnableGUI(false);
      mySearchThread.Start();
    }
    #endregion

    #region select
    /// <summary>
    /// The user has selected a testdefinition. 
    /// Checks if the testdefinition can be opened in the active workbook and 
    /// creates the sheet(s).
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    private void select(object sender, EventArgs e)
    {
      if (resultView.SelectedItems.Count > 0)
      {
        // This ennumeration returns only ONE TestDefinition
        System.Collections.IEnumerator tmpSelected = resultView.SelectedItems.GetEnumerator();
        if (tmpSelected.MoveNext())
        {
          mySelectedRow = ((ListViewItem)tmpSelected.Current);
          mySelectedTd = (Lib.Testdefinition)mySelectedRow.Tag;
          if (VetoToOpen(mySelectedTd, Globals.PDCExcelAddIn.Application.ActiveWorkbook))
          {
            MessageBox.Show(Properties.Resources.MSG_TEST_ALREADY_OPEN_TEXT, Properties.Resources.MSG_TEST_ALREADY_OPEN_TITLE,
              MessageBoxButtons.OK, MessageBoxIcon.Stop);
            return;
          }
          Thread tmpThread = new Thread(new ThreadStart(InitializeTestDefinition));
          EnableGUI(false);
          tmpThread.Start();
        }                
      }
    }
    #endregion

    #region VetoToOpen
    /// <summary>
    /// Returns true if the Testdefinition should not be opened, since a version of it
    /// is already openend on another sheet.
    /// </summary>
    /// <param name="selectedTd"></param>
    /// <returns></returns>
    private bool VetoToOpen(Lib.Testdefinition selectedTd, Excel.Workbook aWorkbook)
    {
      return Globals.PDCExcelAddIn.TestOpen(aWorkbook, selectedTd);
    }
    #endregion

    #region WorkerThreadFinished
    /// <summary>
    /// Displays the status or results after a background operation
    /// </summary>
    /// <param name="aStatusOrException"></param>
    public void WorkerThreadFinished(object aStatusOrException)
    {
      if (aStatusOrException is Exception)
      {
        ExceptionHandler.TheExceptionHandler.handleException((Exception)aStatusOrException, this);
        EnableGUI(true);
        return;
      } 
      Status aStatus = (Status)aStatusOrException;
      if (aStatus == Status.SEARCH_DONE)
      {
        myResult = new List<ListViewItem>();
        Lib.UserInfo tmpUserInfo = Globals.PDCExcelAddIn.PdcService.UserInfo;
        foreach (Lib.Testdefinition tmpDefinition in myTestDefinitionList)
        {
          ListViewItem tmpRow = new ListViewItem(new string[] {
            "" + tmpDefinition.TestNo,
            tmpDefinition.TestName,
            "" + tmpDefinition.Version,
            tmpDefinition.HasUploadPrivileges(tmpUserInfo)?
                Properties.Resources.PRIVILEGE_UPLOAD:
                Properties.Resources.PRIVILEGE_READ_ONLY}) ;
          tmpRow.Tag = tmpDefinition;
          myResult.Add(tmpRow);
        }
        DisplayResult();
      }
      EnableGUI(true);
      if (aStatus == Status.CLOSE)
      {
        Dispose();
      }
    }
    #endregion


    #endregion

    #region properties

    #endregion
  }
}
