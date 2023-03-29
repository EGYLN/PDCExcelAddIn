using System.Collections.Generic;
using System.Diagnostics;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Actions;
using Microsoft.Office.Tools.Ribbon;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Implementation of the PDC ribbon.
    /// </summary>
    public partial class PdcDesignedRibbon
    {
        private readonly IDictionary<RibbonButton, PDCAction> myActions = new Dictionary<RibbonButton, PDCAction>();

        internal void RegisterActions()
        {
            PDCExcelAddIn addin = PDCExcelAddIn.TheSingleton();
            myActions[loginWindowsButton] = addin.WindowsLoginAction;
            myActions[loginOtherButton] = addin.LoginAction;
            myActions[logoutButton] = addin.LogoutAction;
            myActions[newWorksheetButton] = addin.CreateWorkbookAction;
            myActions[uploadButton] = addin.UploadDataAction;
            myActions[validateButton] = addin.ValidateAction;
            myActions[searchButton] = addin.SearchTestdataAction;
            myActions[clearButton] = addin.ClearDataAction;
            myActions[uploadChangesButton] = addin.UpdateDataAction;
            myActions[deleteButton] = addin.DeleteAction;
            myActions[retrieveMeasurementsButton] = addin.RetrieveMeasurementLevelDataAction;
            myActions[chemicalDataButton] = addin.CompoundDataAction;
            myActions[formatCompoundButton] = addin.StructureFormatAction;
            myActions[versionInfoButton] = addin.VersionAction;
        }

        internal void UpdateEnablement()
        {
            foreach (var pair in myActions)
            {
                pair.Key.Enabled = pair.Value.Enabled;
                pair.Key.Visible = pair.Value.Visible;
            }
        }

        private void OpeninBrowser(object sender, RibbonControlEventArgs e)
        {
            string url = sender == contactButton
                ? Properties.Settings.Default.URL_Support
                : Properties.Settings.Default.URL_Documenation;
            ProcessStartInfo process = new ProcessStartInfo(url);
            Process.Start(process);
        }
        private void ButtonClicked(object sender, RibbonControlEventArgs e)
        {
            RibbonButton button = sender as RibbonButton;
            PDCAction action;
            if (button != null && myActions.TryGetValue(button, out action))
            {
                action.PerformAction(true);
            }
        }
    }
}
