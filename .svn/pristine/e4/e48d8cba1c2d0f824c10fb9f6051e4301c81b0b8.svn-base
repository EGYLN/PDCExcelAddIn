using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Not used anymore
    /// </summary>
    [ComVisible(false)]
    public partial class SearchTestdataDialog : Form
    {
        private Lib.Testdefinition testDefinition;
        private Lib.Testdata testData;
        private Lib.TestdataSearchCriteria searchCriteria;

        [ComVisible(false)]
        public delegate void FinishedCallback(Lib.Testdata aResult);
        public event FinishedCallback Finished;

        public SearchTestdataDialog()
        {
            InitializeComponent();
        }
        private void performSearch()
        {
            Lib.PDCService tmpService = Globals.PDCExcelAddIn.PdcService;
            List<Lib.TestdataSearchCriteria> tmpCriterias = new List<BBS.ST.BHC.BSP.PDC.Lib.TestdataSearchCriteria>();
            tmpCriterias.Add(searchCriteria);
            testData = tmpService.FindTestdata(testDefinition, tmpCriterias);
            BeginInvoke(new MethodInvoker(delegate()
            {
                SearchCompleted();
            }));

        }
        private void Callback()
        {
            Finished(testData);
            BeginInvoke(new MethodInvoker(delegate()
            {
                CallbackCompleted();
            }));
        }
        private void CallbackCompleted()
        {
            this.UseWaitCursor = false;
            searchButton.Enabled = true;
            cancelButton.Enabled = true;
            Dispose();
        }
        private void SearchCompleted()
        {
            DialogResult tmpResult;
            this.UseWaitCursor = false;
            searchButton.Enabled = true;
            cancelButton.Enabled = true;

            if (testData == null)
            {
              
                tmpResult = MessageBox.Show(Properties.Resources.MSG_NO_ENTRIES_FOUND_TEXT, Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.YesNo);
                if (DialogResult.Yes == tmpResult)
                {
                    return;
                }
                Dispose();
            }
            tmpResult = MessageBox.Show(string.Format(Properties.Resources.MSG_SEARCH_RESULT_TEXT, testData.Experiments.Count ), Properties.Resources.MSG_SEARCH_TITLE, MessageBoxButtons.YesNoCancel);
            if (DialogResult.Yes == tmpResult)
            {
                if (testData != null && Finished != null && Finished.GetInvocationList() != null && Finished.GetInvocationList().Length > 0)
                {
                    Thread tmpThread = new Thread(Callback);
                    UseWaitCursor = true;
                    searchButton.Enabled = false;
                    cancelButton.Enabled = false;
                    tmpThread.Start();
                    return;
                } 

                Dispose();
                return;
            }
            else if (DialogResult.No == tmpResult)
            {
                testData = null;
                return;
            }
            testData = null;
            Dispose();
        }
        private void searchButton_Click(object sender, EventArgs e)
        {
            searchCriteria = new Lib.TestdataSearchCriteria();
            searchCriteria.TestDefinition = testDefinition;
            searchCriteria[Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER, false] = new Lib.TestVariableValue(Lib.PDCConstants.C_ID_COMPOUNDIDENTIFIER,compoundNoField.Text);
            searchCriteria[Lib.PDCConstants.C_ID_PREPARATIONNO, false] = new Lib.TestVariableValue(Lib.PDCConstants.C_ID_PREPARATIONNO, preparationNoField.Text);

            Thread tmpThread = new Thread(new ThreadStart(performSearch));
            this.UseWaitCursor = true;
            this.UseWaitCursor = true;
            searchButton.Enabled = false;
            cancelButton.Enabled = false;
            tmpThread.Start();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            testData = null;
            Dispose();
        }

        public Lib.Testdata ShowDialog(Lib.Testdefinition aTD)
        {
            testDefinition = aTD;
            Text = Text + " " + aTD.TestNo;
            testField.Text = aTD.TestName;
            versionField.Text = aTD.Version == null ? "" : "" + aTD.Version;
            uploadDateFrom.Value = new DateTime(1900, 1, 1);
            uploadDateTo.Value = DateTime.Now;
            this.CenterToScreen();
            ShowDialog();
            return testData;
        }
    }
}
