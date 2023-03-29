using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Dialog with a table presenting validation and error messages.
    /// </summary>
    [ComVisible(false)]
  public partial class MessagesDialog : Form
  {
      /// <summary>
      /// Cons.
      /// </summary>
    public MessagesDialog()
    {
      InitializeComponent();
    }

    private void closeButton_Click(object sender, EventArgs e)
    {
      Dispose();
    }

    /// <summary>
    /// Displays a dialog with the validation/error messages within a table.
    /// </summary>
    /// <param name="aListObject"></param>
    /// <param name="theMessages"></param>
    /// <param name="aTitle"></param>
    /// <param name="aModalFlag"></param>
    public static void DisplayMessages(PDCListObject aListObject, List<Lib.PDCMessage> theMessages, string aTitle, bool aModalFlag)
    {
        MessagesDialog tmpDialog = new MessagesDialog {Text = aTitle};
        //Fill list view.
      List<ListViewItem> tmpViewItems = new List<ListViewItem>();
      int tmpStartRow = aListObject.ToSheetRow(0);
      foreach (Lib.PDCMessage tmpMessage in theMessages)
      {
        string tmpColumn = tmpMessage.ParameterName;
        if (tmpMessage.VariableName != null)
        {
          tmpColumn = tmpMessage.VariableName;
        }
        ListViewItem tmpItem = new ListViewItem(new [] {string.Empty +(tmpMessage.ExperimentIndex+tmpStartRow), tmpColumn, tmpMessage.Message, tmpMessage.MessageType, });
        tmpItem.Tag = tmpMessage;
        tmpViewItems.Add(tmpItem);
      }
      tmpDialog.messageTable.Items.AddRange(tmpViewItems.ToArray());
      if (aModalFlag)
      {
        tmpDialog.ShowDialog();
      }
      else
      {
        tmpDialog.Show();
      }
    }

    private void MessagesDialog_Load(object sender, EventArgs e)
    {
    }
  }
}
