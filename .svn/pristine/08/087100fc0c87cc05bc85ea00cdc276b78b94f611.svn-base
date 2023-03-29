using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Progress dialog for time consuming operations.
    /// The operation is executed in a background worker thread.
    /// Optionally the user can cancel the operation.
    /// </summary>
    [ComVisible(false)]
  public partial class ProgressDialog : Form
  {
    public const string CANCELLED = "Cancelled";

    Status statusDelegate;
    Executor executor;
    Thread executingThread;
    private bool myCanCancel = true;
    private bool myInteractive = true;
    private bool myIsCancelled = false;
    private object myResult = null;

    [ComVisible(false)]
    public delegate void Status(object result, ProgressDialog windowOwner, bool interactiveMode);

    /// <summary>
    /// This delegate implements a possibly time consuming operation.
    /// The delegate has to call the Status callback delegate to finish the operation and 
    /// close the progess dialog.
    /// </summary>
    /// <param name="windowOwner"></param>
    [ComVisible(false)]
    public delegate void Executor(ProgressDialog windowOwner);

    #region constructor

    public ProgressDialog(Status aDelegate, Executor anExecutor, string aLabel)
    {
      statusDelegate = aDelegate;
      executor = anExecutor;
      InitializeComponent();
      textLabel.Text = aLabel;
      textLabel.TextAlign = ContentAlignment.MiddleCenter;
    }

    #endregion

    #region events

    #region cancelButton_Click
    private void CancelButton_Click(object sender, EventArgs e)
    {
      myIsCancelled = true;
      if (executingThread != null)
      {
        executingThread.Abort();
      }
      if (!IsDisposed)
      {
        Dispose();
      }
      executingThread = null;
      myResult = null;
      statusDelegate(CANCELLED, this, myInteractive);
    }
    #endregion

    #region ProgressDialog_Shown
    private void ProgressDialog_Shown(object sender, EventArgs e)
    {
      if (executingThread == null)
      {
        executingThread = new Thread(new ThreadStart(Execute));
        executingThread.CurrentCulture = Thread.CurrentThread.CurrentCulture;
        executingThread.CurrentUICulture = Thread.CurrentThread.CurrentUICulture;
        executingThread.Start();
      }
    }
    #endregion

    #endregion

    #region methods

    #region Execute
    private void Execute()
    {
      executor(this);
    }
    #endregion

    #region SearchFinished
    private void SearchFinished(object result, ProgressDialog windowOwner, bool interactiveMode)
    {
      try
      {
        statusDelegate(result, windowOwner, interactiveMode);
        myResult = result;
      }
      catch (Exception e)
      {
        if (!Cancelled && !IsDisposed)
        {
          if (myInteractive)
          {
            ExceptionHandler.TheExceptionHandler.handleException(e, this);
          }
          else
          {
            myResult = e;
          }
        }
      }
      finally
      {
        Globals.PDCExcelAddIn.EnableExcel(); //TODO necessary for non-interactive mode?
        if (!IsDisposed)
        {
          Dispose();
        }
      }
    }
    #endregion

    #region Show
    public static object Show(Status aDelegate, Executor executor, string label)
    {
      return Show(aDelegate, executor, label, true, true);
    }

    public static object Show(Status aDelegate, Executor executor, string label, bool isCancelPossible, bool isInteractive)
    {
      ProgressDialog tmpDialog = new ProgressDialog(aDelegate, executor, label);
      tmpDialog.myInteractive = isInteractive;
      tmpDialog.CanCancel = isCancelPossible && isInteractive;
      tmpDialog.ShowDialog(new ExcelHwndWrapper());
      return tmpDialog.myResult;
    }
    #endregion

    #region StatusCallback
    public void StatusCallback(object result)
    {
      if (!myIsCancelled)
      {
        BeginInvoke(new Status(SearchFinished), new object[] { result, this, myInteractive});
      }
    }
    #endregion

    #endregion

    #region properties

    #region CanCancel
    public bool CanCancel
    {
      get
      {
        return myCanCancel;
      }
      set
      {
        myCanCancel = value;
        cancelButton.Enabled = myCanCancel;
      }
    }
    #endregion

    #region Cancelled
    public bool Cancelled
    {
      get
      {
        return myIsCancelled;
      }
    }
    #endregion

    #region Message
    /// <summary>
    /// Property for the display message
    /// </summary>
    public string Message
    {
      get
      {
        return textLabel.Text;
      }
      set
      {
        textLabel.Text = value;
      }
    }
    #endregion

    #region Result
    /// <summary>
    /// Returns the result of the executed action or an exception
    /// </summary>
    public object Result
    {
      get
      {
        return myResult;
      }
    }
    #endregion

    #endregion
  }
}
