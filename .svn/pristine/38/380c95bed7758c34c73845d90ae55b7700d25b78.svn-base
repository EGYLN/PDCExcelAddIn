using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using VBE = Microsoft.Vbe.Interop;
using Coll = System.Collections.Generic;
using System.Threading;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Formatters;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Actions;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using System.Drawing;
using Microsoft.Win32;

using System.Runtime.InteropServices;
using System.Globalization;
using System.Diagnostics;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Util;
using BBS.ST.BHC.BSP.PDC.Lib;
using BBS.ST.BHC.BSP.PDC.Lib.Properties;
using log4net.Core;
using Microsoft.Office.Core;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    public partial class PDCExcelAddIn
    {
        // Declares Win32 function used for low-level event handling
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        private const string CUSTOM_PROPERTY_VERSION = "PDC_WorkbookVersion";

        private const string CUSTOM_PROPERTY_STORE = "PDC_WorkbookState";

        //Tags for Excel menu entries. Can be used to search specific menus
        private const string TAG_CI_CONTEXT_MENU = "PDC_CI_Context";
        private const string TAG_CI_LIST_CONTEXT_MENU = "PDC_CI_ListContext";


        //Menu actions
        private PDCAction loginAction;
        private PDCAction winLoginAction;
        private PDCAction logoutAction;
        private PDCAction validateAction;
        private PDCAction searchTestdataAction;
        private PDCAction pdcCreateWorkbookAction;
        private PDCAction uploadAction;
        private PDCAction deleteAction;
        private PDCAction updateAction;
        private PDCAction compoundInfoAction;
        private PDCAction structureFormatAction;
        private PDCAction clearDataAction;
        private PDCAction versionInfoAction;
        private PDCAction retrieveMeasurementLevelDataAction;
        private bool retrieveMeasurementLevelEnabled;

        //actions for the compound information context menu
        private PDCAction retrieveMeasurementsContextMenuAction;
        private CompoundInfoAction compoundInfoSelectedAction;
        private CompoundInfoAction prepnoAction;
        private CompoundInfoAction formatCompoundsAction;
        private CompoundInfoAction formulaAction;
        private CompoundInfoAction weightAction;
        private CompoundInfoAction structureAction;
        private CompoundInfoAction formatZKAction;
        private CompoundInfoAction formatBAYAction;
        private CompoundInfoAction formatCOSAction;
        private CompoundInfoAction formatCOPAction;

        private bool myCreateBothMeasurementTables;

        private readonly object LOCK = new object();
        private Lib.ClientConfiguration myClientConfiguration;
        private readonly Dictionary<Excel.Worksheet, SheetInfo> mySheetMap = new Dictionary<Excel.Worksheet, SheetInfo>();

        /// <summary>
        /// Workaround double workbook open event during macro trust check 
        /// </summary>
        private readonly HashSet<string> myOpenedWorkbookNames = new HashSet<string>();

        /// <summary>
        /// Mapping from sheet key to SheetInfo
        /// </summary>
        private readonly Dictionary<object, SheetInfo> mySheetInfos = new Dictionary<object, SheetInfo>();

        private volatile IDictionary<Excel.Workbook, Excel.Workbook> myWorkbooks = new Coll.Dictionary<Excel.Workbook, Excel.Workbook>();

        private bool myAreEventsEnabled = true;
        private bool myIsShuttingDown;
        private bool myIsWorbookClosing;


        private Lib.PDCService myPdcService;

        private readonly UserSettings myUserSettings = new UserSettings();

        private bool myIsInSheetChanged;


        /// <summary>
        /// Returns the client configuration (color, binary data constraints
        /// </summary>
        public Lib.ClientConfiguration ClientConfiguration
        {
            get
            {
                if (myClientConfiguration == null)
                {
                    myClientConfiguration = PdcService.ClientConfiguration();
                }
                return myClientConfiguration;
            }
        }

        /// <summary>
        /// Convenience method.
        /// </summary>
        /// <returns></returns>
        public static PDCExcelAddIn TheSingleton()
        {
            return Globals.PDCExcelAddIn;
        }

        /// <summary>
        /// Returns the SheetInfo for the specified sheet key
        /// </summary>
        /// <param name="aSheetKey"></param>
        /// <returns></returns>
        internal SheetInfo GetSheetInfo(object aSheetKey)
        {
            if (aSheetKey == null)
            {
                return null;
            }
            if (mySheetInfos.ContainsKey(aSheetKey))
            {
                return mySheetInfos[aSheetKey];
            }
            return null;
        }

        /// <summary>
        /// Adds the specified SheetInfo to the list of sheet infos if it is not 
        /// already known.
        /// </summary>
        /// <param name="aSheetInfo"></param>
        internal void AddSheetInfo(SheetInfo aSheetInfo)
        {
            if (!mySheetInfos.ContainsKey(aSheetInfo.Identifier))
            {
                mySheetInfos.Add(aSheetInfo.Identifier, aSheetInfo);
            }
        }

        /// <summary>
        /// Returns the SheetInfo for the specified Excel sheet or null if it has none
        /// </summary>
        /// <param name="aSheet"></param>
        /// <returns></returns>
        internal SheetInfo GetSheetInfo(Excel.Worksheet aSheet)
        {
            if (aSheet == null)
            {
                return null;
            }
            if (mySheetMap.ContainsKey(aSheet))
            {
                return mySheetMap[aSheet];
            }
            SheetInfo tmpSheetInfo = GetSheetInfo(ExcelUtils.TheUtils.GetKey(aSheet));
            if (tmpSheetInfo != null)
            {
                mySheetMap.Add(aSheet, tmpSheetInfo);
            }
            return tmpSheetInfo;
        }

        /// <summary>
        /// Removes the specified SheetInfo from the list of sheet infos.
        /// </summary>
        /// <param name="aSheetInfo"></param>
        internal void RemoveSheetInfo(SheetInfo aSheetInfo)
        {
            if (mySheetInfos.ContainsKey(aSheetInfo.Identifier))
            {
                mySheetInfos.Remove(aSheetInfo.Identifier);
            }
            aSheetInfo.Cleanup();
        }

        /// <summary>
        /// The need references to the pdc workbook. Otherwise our registered eventhandler
        /// will get lost.
        /// </summary>
        public IDictionary<Excel.Workbook, Excel.Workbook> WorkbookMap
        {
            get
            {
                return myWorkbooks;
            }
        }

        /// <summary>
        /// Registers a Testdefinition for use with the specified worksheet.
        /// Returns a SheetInfo for the sheet.
        /// </summary>
        /// <param name="aTD"></param>
        /// <param name="aSheet"></param>
        /// <returns></returns>
        internal SheetInfo RegisterTestdefinition(Lib.Testdefinition aTD, Excel.Worksheet aSheet)
        {
            object tmpKey = ExcelUtils.TheUtils.GetKey(aSheet, true);
            SheetInfo tmpInfo = GetSheetInfo(tmpKey);
            if (tmpInfo == null)
            {
                tmpInfo = new SheetInfo();
                tmpInfo.Identifier = tmpKey;
                tmpInfo.ExcelSheet = aSheet;
                AddSheetInfo(tmpInfo);
                try
                {
                    aSheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_ENTRY_SHEET, "true");
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "customproperty for A5", e);
                }
            }
            tmpInfo.TestDefinition = aTD;

            //object tmpKey = Guid.NewGuid().ToString();
            //Excel.CustomProperty tmpProperty = aSheet.CustomProperties.Add("PDC", tmpKey);

            return tmpInfo;
        }

        /// <summary>
        /// Registers a measurement table on the specified worksheet.
        /// Returns the SheetInfo for the containing worksheet.
        /// </summary>
        /// <param name="aMeasurementTable"></param>
        /// <param name="aSheet"></param>
        /// <param name="aMainSheet"></param>
        /// <returns></returns>
        internal SheetInfo RegisterMeasurementTable(PDCListObject aMeasurementTable, Excel.Worksheet aSheet, SheetInfo aMainSheet)
        {
            PDCLogger.TheLogger.LogStarttime("RegisterMeasurementTable", "Registering table " + aMeasurementTable.Name);
            SheetInfo tmpExistingSheetInfo = null;
            object aKey = null;
            if (mySheetMap.ContainsKey(aSheet))
            {
                tmpExistingSheetInfo = mySheetMap[aSheet];
                aKey = tmpExistingSheetInfo.Identifier;
            }
            else
            {
                aKey = ExcelUtils.TheUtils.GetKey(aSheet, true);
                tmpExistingSheetInfo = GetSheetInfo(aKey);
                if (tmpExistingSheetInfo != null)
                {
                    mySheetMap.Add(aSheet, tmpExistingSheetInfo);
                }
            }
            if (tmpExistingSheetInfo == null)
            {
                tmpExistingSheetInfo = new SheetInfo();
                tmpExistingSheetInfo.ExcelSheet = aSheet;
                tmpExistingSheetInfo.Identifier = aKey;
                tmpExistingSheetInfo.TestDefinition = aMainSheet.TestDefinition;
                tmpExistingSheetInfo.MainSheetInfo = aMainSheet;
                AddSheetInfo(tmpExistingSheetInfo);
                aMainSheet.AddAdditionalSheet(tmpExistingSheetInfo);
                try
                {
                    aSheet.CustomProperties.Add(PDCExcelConstants.PROPERTY_MEASUREMENT, "true");
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "customproperty for A5", e);
                }
            }
            string tmpListName = aMeasurementTable.ListRangeName;
            tmpExistingSheetInfo.AddMeasurementTable(tmpListName, aMeasurementTable);
            PDCLogger.TheLogger.LogStoptime("RegisterMeasurementTable", "Registered table " + aMeasurementTable.Name);
            return tmpExistingSheetInfo;
        }

        public string LoggedInCwid => LoggedInUser?.Cwid ?? string.Empty;
        /// <summary>
        /// Property for the login state, which is centralized in the PDCLIB
        /// </summary>
        public UserInfo LoggedInUser
        {
            get => Lib.RegistryUtil.LoggedInUser;
            set
            {
                Lib.RegistryUtil.LoggedInUser = value;
                EnablePdcActions();
            }
        }

        /// <summary>
        /// Property for the login state
        /// </summary>
        public bool IsLoggedIn => !string.IsNullOrEmpty(LoggedInCwid);

        /// <summary>
        /// Property for the login state
        /// </summary>
        public void ResetLoggedIn()
        {
            LoggedInUser = null;
        }

        /// <summary>
        /// Convenience method to create the intersection of two ranges.
        /// </summary>
        /// <param name="aRange1">First cell region</param>
        /// <param name="aRange2">Second cell region</param>
        /// <returns>The intersection of both regions</returns>
        public Excel.Range Intersect(Excel.Range aRange1, Excel.Range aRange2)
        {
            return Application.Intersect(aRange1, aRange2, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
        }

        /// <summary>
        /// Convenience method to create the union of two ranges
        /// </summary>
        /// <param name="aRange1">First cell region</param>
        /// <param name="aRange2">Second cell region</param>
        /// <returns>The union of both regions</returns>
        public Excel.Range Union(Excel.Range aRange1, Excel.Range aRange2)
        {
            return Application.Union(aRange1, aRange2, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
        }

        /// <summary>
        /// Property to enable/disable the event processing. Used to avoid
        /// recursive event processing/optimization.
        /// </summary>
        public bool EventsEnabled
        {
            get
            {
                return myAreEventsEnabled;
            }
            set
            {
                myAreEventsEnabled = value;
            }
        }

        public Lib.PDCService PdcService
        {
            get
            {
                lock (LOCK)
                {
                    if (myPdcService == null)
                    {
                        myPdcService = Lib.PDCService.ThePDCService;
                    }
                    return myPdcService;
                }
            }
        }

        internal bool DoCreateBothMeasurementTables()
        {
            return myCreateBothMeasurementTables;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs anEvent)
        {
            PDCLogger.TheLogger.LogSevere(PDCLogger.LOG_NAME_EXCEL, "Startup PDC");
            bool failed = false;
            Lib.RegistryUtil.LoggedInUser = null;
            Application.EnableEvents = true;

            myCreateBothMeasurementTables = ((int)Registry.GetValue(Lib.PDCClientConstants.PDC_REGISTRY_KEY, "MMT&SMT", 0) == 1);

            try
            {
                RegisterOpenLibAddIn(true);
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Exception while registering open library", e);
            }
            try
            {
                AddExcelHooks();
            }
            catch (Exception e1)
            {
                failed = true;
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Exception while registering event handler. Skipping Menu creation.", e1);
                ExceptionHandler.TheExceptionHandler.handleException(e1, null);
            }
            if (!failed)
            {
                try
                {
                    InitializeActions();
                    ReadSettings();
                    if (!VersionInfoAction.GetInstallPath().Equals("") && !File.Exists(Path.Combine(VersionInfoAction.GetInstallPath(), "pdc-log.txt")))
                    {
                        MessageBox.Show(string.Format(Properties.Resources.MSG_PATH_NO_WRITABLE, VersionInfoAction.GetInstallPath()), Properties.Resources.MSG_ERROR_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                    }

                }
                catch (Exception e2)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Exception while setting up menu.", e2);
                    ExceptionHandler.TheExceptionHandler.handleException(e2, null);
                }
            }
        }

        private void ReadSettings()
        {
            myUserSettings.ReadSettings();
        }


        /// <summary>
        /// Convenience method: Tries to register the OpenLib as Excel Add-In, so that
        /// it is not necessary to do it manually using the automation button.
        /// </summary>
        /// <param name="aReregisterFlag">a registered OpenLib Addin is only reregistered if the flag is true</param>
        private void RegisterOpenLibAddIn(bool aReregisterFlag)
        {
            bool tmpInstalled = false;
            //Workaround to get Excel to use the same AppDomain for calling our UDFs
            try
            {
                IEnumerator tmpAddIns = Application.AddIns.GetEnumerator();
                while (tmpAddIns.MoveNext())
                {
                    Excel.AddIn tmpAddIn = (Excel.AddIn)tmpAddIns.Current;
                    string tmpId = tmpAddIn.progID;
                    string tmpName = "";
                    try
                    {
                        tmpName = tmpAddIn.Name;
                    }
                    catch (Exception)
                    {
                    }

                    PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_COM, "Found AddIn " + tmpName + "(ProdId:" + tmpId + ")");
                    if (tmpId == "PDCOpenLibrary.PDCOpenLib" || tmpName == "PDCOpenLibrary.PDCOpenLib")
                    {
                        if (tmpAddIn.Installed)
                        {
                            if (!aReregisterFlag)
                            {
                                return;
                            }
                            tmpAddIn.Installed = false;
                        }
                        tmpAddIn.Installed = true;
                        tmpInstalled = true;
                        break;
                    }
                }
                if (!tmpInstalled)
                {
                    Excel.AddIn tmpOpenLibAddIn = Application.AddIns.Add("PDCOpenLibrary.PDCOpenLib", missing);
                    tmpOpenLibAddIn.Installed = true;
                }
            }
            catch (Exception ee)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_COM, "Exception during OpenLib-Addin registration", ee);
            }
        }

        /// <summary>
        /// Registers the diverse event handlers in Excel.
        /// </summary>
        private void AddExcelHooks()
        {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Adding event handlers");
            Application.WorkbookActivate += WorkbookActivated;
            Application.WorkbookDeactivate += WorkbookDeactivated;
            Application.WorkbookOpen += WorkbookOpened;
            Application.WorkbookBeforeClose += WorkbookClosing;
            Application.SheetActivate += WorksheetActivated;
            Application.SheetDeactivate += WorksheetDeactivated;
            Application.WorkbookBeforeSave += WorkbookSaving;
            Application.SheetBeforeRightClick += SheetBeforeRightClick;

            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Initialization of PDC Excel AddIn Version: " + VersionInfoAction.GetVersionText());
        }

        /// <summary>
        /// Event delegate called when Excel is about to save a workbook.
        /// User may cancel the save operation after the delegate was called!
        /// </summary>
        /// <param name="Wb">The workbook which is possibly saved on disk</param>
        /// <param name="SaveAsUI"></param>
        /// <param name="Cancel"></param>
        void WorkbookSaving(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                saveWorkbookState(Wb);
            }
            catch (Exception e)
            {
                ExceptionHandler.TheExceptionHandler.handleException(e, null);
            }
        }

        /// <summary>
        /// Do some menu enabling
        /// </summary>
        void WorksheetDeactivated(object sh)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "WorkbookDeactivated");
            EnablePdcActions();
        }

        /// <summary>
        /// Do some menu enabling
        /// </summary>
        void WorksheetActivated(object sh)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "WorksheetActivated");
            EnablePdcActions();
        }

        /// <summary>
        /// Sets Language with the Config-File 'pdcconfig.properties'. Attention: 
        /// During debugging (running within IDE) the config pdcconfig.properties is NOT being read
        /// as the installer is responsible for putting the config-File into the correct directory
        /// </summary>
        internal static void SetupLanguage()
        {
            var cuii = CultureInfo.CurrentUICulture;

            //only if the property "Language" is "english"
            if (UserConfiguration.TheConfiguration.GetProperty("Language", "german").Equals("english", StringComparison.OrdinalIgnoreCase))
            {
                CultureInfo enCulture = new CultureInfo("en-US")
                {
                    NumberFormat = cuii.NumberFormat
                };
                Thread.CurrentThread.CurrentUICulture = enCulture;
            }
            else
            {
                CultureInfo enCulture = new CultureInfo("de-DE") {NumberFormat = cuii.NumberFormat};
                Thread.CurrentThread.CurrentUICulture = enCulture;
            }
        }

        /// <summary>
        /// Initializes the availabe PDC actions
        /// </summary>
        private void InitializeActions()
        {
            // Sets Language:
            SetupLanguage();
            //Define the Shortcut key for RetrieveMeasurementLevelData 
            Application.OnKey(RetrieveMeasurementLevelDataShortcut(), "PDC_LoadButton_Click");
            //Login/Logout
            loginAction = new LoginAction(false);

            winLoginAction = new WinLoginAction(false);

            logoutAction = new LogoutAction(false);
            logoutAction.Visible = false;

            //Create PDC Workbook
            pdcCreateWorkbookAction = new CreateWorkbookAction(true);

            //Upload/Update/Delete
            uploadAction = new UploadWorkbookAction(false);
            updateAction = new UpdateAction(false);
            deleteAction = new DeleteAction(false);


            //Clear
            clearDataAction = new ClearDataAction(false);

            //Validation
            validateAction = new ValidationAction(false);

            //retrieve measurement data
            retrieveMeasurementLevelDataAction = new RetrieveMeasurementLevelDataAction(true);
            retrieveMeasurementLevelDataAction.Shortcut = RetrieveMeasurementLevelDataShortcutText();

            //Search
            searchTestdataAction = new SearchTestdataAction(true);

            //Miscellaneous
            compoundInfoAction = new CompoundInfoAction(false, myUserSettings);
            compoundInfoAction.Enabled = true;

            structureFormatAction = new StructureFormatAction(false, myUserSettings);
            structureFormatAction.Enabled = true;

            versionInfoAction = new VersionInfoAction(false);
            versionInfoAction.Enabled = true;

            Globals.Ribbons.PdcDesignedRibbon.RegisterActions();
            EnablePdcActions();
        }

        private Office.CommandBarPopup GetPopup(CommandBar commandBar, string tag) => (CommandBarPopup) commandBar.FindControl(MsoControlType.msoControlPopup,null,tag, false, false); 

        private void RemovePopup(CommandBarPopup popup)
        {
            if (popup == null)
            {
                return;
            }
            foreach(var control in popup.Controls)
            {
                if (control is CommandBarButton button)
                {
                    button.Delete();
                }
            }
            popup.Delete();
        }
        /// <summary>
        /// Initializes the context menu for the compound informations
        /// Is called after a mouse right-click 
        /// </summary>
        private void SetupCompoundInfoContextMenu(CommandBar aContext, CommandBar aListContext)
        {

            // if Contextmenu has been build already, leave the function
            CommandBarPopup ciContextMenu = GetPopup(aContext, TAG_CI_CONTEXT_MENU);
            CommandBarPopup ciListContextMenu = GetPopup(aListContext, TAG_CI_LIST_CONTEXT_MENU);


            bool contextMenuEnabled = IsLoggedIn;

            if (PdcService.UserInfo != null && PdcService.UserInfo.IsIcbOnlyUser)
            {
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, $"{nameof(SheetBeforeRightClick)} - ICB user - PDC context menu not available.");
                RemovePopup(ciListContextMenu);
                RemovePopup(ciContextMenu);
                retrieveMeasurementsContextMenuAction = null;
                return;
            }

            if (ciContextMenu == null)
            {
                
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Initializing pdc context menu.");

                ciContextMenu = (Office.CommandBarPopup)aContext.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                ciContextMenu.Caption = Properties.Resources.CI_Contextmenu_Caption;
                ciContextMenu.Tag = TAG_CI_CONTEXT_MENU;
                compoundInfoSelectedAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.AllSelected, ciContextMenu, false, myUserSettings);

                prepnoAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.PrepnoSelectedOnly, ciContextMenu, true, myUserSettings);
                structureAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.StructureSelectedOnly, ciContextMenu, false, myUserSettings);
                weightAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.WeightSelectedOnly, ciContextMenu, false, myUserSettings);
                formulaAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormulaSelectedOnly, ciContextMenu, false, myUserSettings);

                formatCompoundsAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormatSelectedOnly, ciContextMenu, true, myUserSettings);

                formatBAYAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormatSelectedBay, ciContextMenu, true, myUserSettings);
                formatZKAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormatSelectedZk, ciContextMenu, false, myUserSettings);
                formatCOSAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormatSelectedCos, ciContextMenu, false, myUserSettings);
                formatCOPAction = CompoundInfoAction.CreateCompoundInfoAction(CompoundInfoActionKind.FormatSelectedCop, ciContextMenu, false, myUserSettings);

                retrieveMeasurementsContextMenuAction = new RetrieveMeasurementLevelDataAction(true);
                retrieveMeasurementsContextMenuAction.Shortcut = RetrieveMeasurementLevelDataShortcutText();
            
                retrieveMeasurementsContextMenuAction.AddToMenu(ciContextMenu, TAG_CI_CONTEXT_MENU);
            }
            ciContextMenu.Enabled = contextMenuEnabled;
            if (ciListContextMenu == null)
            {
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Initializing pdc list context menu.");
                ciListContextMenu = (Office.CommandBarPopup)aListContext.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, missing, true);
                ciListContextMenu.Caption = Properties.Resources.CI_Contextmenu_Caption;
                ciListContextMenu.Tag = TAG_CI_LIST_CONTEXT_MENU;
                compoundInfoSelectedAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);

                prepnoAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                structureAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                weightAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                formulaAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);

                formatCompoundsAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);

                formatBAYAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                formatZKAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                formatCOSAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                formatCOPAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
                retrieveMeasurementsContextMenuAction.AddToMenu(ciListContextMenu, TAG_CI_LIST_CONTEXT_MENU);
            }
            ciListContextMenu.Enabled = contextMenuEnabled;
            retrieveMeasurementsContextMenuAction.Enabled = retrieveMeasurementLevelEnabled;
        }

        /// <summary>
        /// A workbook was opened. Examines if it is a PDC workbook and reinitializes the
        /// respective internal state.
        /// </summary>
        /// <param name="aWorkbook"></param>
        private void WorkbookOpened(Excel.Workbook aWorkbook)
        {
            if (AlreadyOpened(aWorkbook))
            {
                return;
            }
            try
            {
                string tmpWBName = aWorkbook.Name;
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Opening Workbook " + tmpWBName);

                Excel.Worksheet tmpSheet = HiddenPDCSheet(aWorkbook, tmpWBName);
                //Alternative we could try to reconstruct the SheetInfo from the named ranges and get the
                //test definition from the webservice. This would afford some matching effort, 
                //but we wouldn't need the serialization anymore
                if (tmpSheet != null)
                {
                    PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Reading pdc data for " + tmpWBName);
                    ResolveEventHandler loadComponentAssembly = LoadComponentAssembly;
                    AppDomain.CurrentDomain.AssemblyResolve += loadComponentAssembly;
                    try
                    {
                        Excel.CustomProperty tmpVersionProperty = FindCustomProperty(tmpSheet, CUSTOM_PROPERTY_VERSION);
                        string tmpVersion = null;
                        if (tmpVersionProperty != null)
                        {
                            tmpVersion = (string) tmpVersionProperty.Value;
                        }

                        byte[] tmpBytes = ReadEncodedCustomProperty(tmpSheet, CUSTOM_PROPERTY_STORE);
                        if (tmpBytes != null)
                        {
                            MemoryStream tmpStream = new MemoryStream(tmpBytes);
                            BinaryFormatter tmpFormatter = new BinaryFormatter();
                            tmpFormatter.AssemblyFormat = FormatterAssemblyStyle.Simple;
                            object tmpResult = null;
                            try
                            {
                                tmpResult = tmpFormatter.UnsafeDeserialize(tmpStream, null);
                            }
                            catch (Exception ee)
                            {
                                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Deserialize failure", ee);
                                DialogResult result = MessageBox.Show(string.Format(Properties.Resources.MSG_DESERIALIZATION_WORKBOOK_TEXT, aWorkbook.Name),
                                    Properties.Resources.MSG_DESERIALIZATION_WORKBOOK_TITLE, MessageBoxButtons.YesNoCancel);
                                switch (result)
                                {
                                    case DialogResult.Yes:
                                        RemoveCustomProperty(tmpSheet, CUSTOM_PROPERTY_STORE);
                                        aWorkbook.Save();
                                        break;
                                    case DialogResult.Cancel:
                                        return;
                                    case DialogResult.No:
                                        break;
                                }
                            }

                            //Remove any outdated state information, since it may overlap with
                            //the loaded workbook
                            RemoveNonExistingSheets();
                            if (tmpResult is WorkbookState)
                            {
                                Reinitialize((WorkbookState) tmpResult, aWorkbook, tmpSheet, tmpVersion);
                                UpdateWorkbookVersion(aWorkbook, tmpSheet, true);
                                if (WorkbookMap.ContainsKey(aWorkbook))
                                {
                                    WorkbookMap[aWorkbook] = aWorkbook;
                                }
                                else
                                {
                                    WorkbookMap.Add(aWorkbook, aWorkbook);
                                }
                            }
                            else
                            {
                                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, tmpWBName + " has no valid pdc state data");
                            }
                        }
                    }
                    finally
                    {
                        AppDomain.CurrentDomain.AssemblyResolve -= loadComponentAssembly;
                    }
                }
            }
            catch (Exception e)
            {
                ExceptionHandler.TheExceptionHandler.handleException(e, null);
            }
            EnablePdcActions();
        }

        /// <summary>
        /// Is the workbook already opened?
        /// </summary>
        /// <param name="aWorkbook"></param>
        /// <returns></returns>
        private bool AlreadyOpened(Excel.Workbook aWorkbook)
        {
            try
            {
                string fullName = aWorkbook.FullName;
                if (myOpenedWorkbookNames.Contains(fullName))
                {
                    return true;
                }
                myOpenedWorkbookNames.Add(fullName);
                return false;
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "AlreadyOpenend?", e);
            }
            return false;
        }

        /// <summary>
        /// checks if workbook has PDC sheets. (All PDC Workbookx do have the Excel.Name Property called PicklistHandler.NAMED_RANGE_PICKLIST_ID
        /// </summary>
        /// <param name="workbook">Excelworkbook to check</param>
        /// <returns>true/false</returns>
        private bool IsPdcWorkBook(Excel.Workbook workbook)
        {
            try
            {
                return (workbook.Names.Item(PicklistHandler.NAMED_RANGE_PICKLIST_ID, Type.Missing, Type.Missing) is Excel.Name);
            }
            catch
            {
                return false;
            }

        }
        /// <summary>
        /// Returns the hidden pdc sheet from the specified workbook or null if no hidden pdc sheet was found.
        /// </summary>
        private Excel.Worksheet HiddenPDCSheet(Excel.Workbook aWorkbook, string workbookName)
        {
            Excel.Range tmpRange = null;
            Excel.Name tmpRangeCand = null;
            try
            {
                tmpRangeCand = aWorkbook.Names.Item(PicklistHandler.NAMED_RANGE_PICKLIST_ID, Type.Missing, Type.Missing);
            }
#pragma warning disable 0168
            catch (Exception)
            {
                //Dont care. Just a ShortCut for iterating over all names to find the relevant one.
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, workbookName + " hat no stored pdc data");
            }
#pragma warning restore 0168

            if (tmpRangeCand != null)
            {
                Excel.Name tmpName = tmpRangeCand;

                string value = tmpName.Value;
                // Invalid reference... no sheet available!
                if (tmpName.Value.StartsWith("=#REF!"))
                {
                    tmpName.Delete();
                    MessageBox.Show(Properties.Resources.MSG_INVALID_WORKBOOK_DEFINED_NAME_DELTED,
                    Properties.Resources.MSG_INFO_TITLE, MessageBoxButtons.OK);
                    return null;
                }
                tmpRange = tmpName.RefersToRange;
            }
            if (tmpRange != null)
            {
                return (Excel.Worksheet)tmpRange.Parent;
            }
            return null;
        }


        /// <summary>
        /// Sets the PDC version of the workbook to the current pdc version.
        /// </summary>
        private void UpdateWorkbookVersion(Excel.Workbook workbook, Excel.Worksheet worksheet, bool onLoad)
        {
            Excel.CustomProperty tmpProperty = FindCustomProperty(worksheet, CUSTOM_PROPERTY_VERSION);
            if (tmpProperty == null)
            {
                worksheet.CustomProperties.Add(CUSTOM_PROPERTY_VERSION, VersionInfoAction.GetVersionText());
                if (onLoad)
                {
                    InitVBACode(workbook);
                }

            }
            else if (!onLoad)
            {
                tmpProperty.Value = VersionInfoAction.GetVersionText();
            }
        }

        /// <summary>
        /// Cleanup resources
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cancel"></param>
        private void WorkbookClosing(Excel.Workbook workbook, ref bool cancel)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "WorkbookClosing");
            string fullName = workbook.FullName;
            if (myOpenedWorkbookNames.Contains(fullName))
            {
                myOpenedWorkbookNames.Remove(fullName);
            }
            if (Application.DisplayAlerts && Application.Visible && !workbook.Saved && !workbook.IsAddin)
            {
                string dialogText = string.Format(Properties.Resources.MSG_SAVE_WORKBOOK_TEXT, workbook.Name);
                if (!IsPdcWorkBook(workbook)) dialogText = dialogText.Substring(5);
                DialogResult result = MessageBox.Show(dialogText, Properties.Resources.MSG_SAVE_WORKBOOK_TITLE, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

                switch (result)
                {
                    case DialogResult.Yes:
                        myIsWorbookClosing = true;
                        workbook.Save();
                        break;
                    case DialogResult.Cancel:
                        myIsWorbookClosing = false;
                        cancel = true;
                        break;
                    case DialogResult.No:
                        myIsWorbookClosing = true;
                        workbook.Saved = true;
                        break;
                }
            }
            else
            {
                myIsWorbookClosing = true;
            }
            EnablePdcActions();
            myIsWorbookClosing = false;
        }


        private void WorkbookActivated(Excel.Workbook aWorkbook)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "WorkbookActivated");
            myIsWorbookClosing = false;
            RegisterOpenLibAddIn(false);
            EnablePdcActions();
        }

        private void WorkbookDeactivated(Excel.Workbook aWorkbook)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "WorkbookDeactivated");
            myIsWorbookClosing = true;
            //Enablement
            EnablePdcActions();
            myIsWorbookClosing = false;
        }

        /// <summary>
        /// Called if cells in a sheet were changed. The event is propageted to the 
        /// appropriate PDCListObjects
        /// </summary>
        /// <param name="e"></param>
        /// <param name="range"></param>
        public void SheetChanged(object e, Excel.Range range)
        {
            if (myIsShuttingDown || range == null)
            {
                return;
            }
            if (!myAreEventsEnabled || myIsInSheetChanged)
            {
                return;
            }
            bool tmpEventsEnabled = Application.EnableEvents;
            Excel.Worksheet tmpSheet = (Excel.Worksheet)range.Parent;
            SheetInfo tmpInfo = GetSheetInfo(tmpSheet);
            if (tmpInfo == null)
            {
                return;
            }

            bool tmpCalcEnabled = tmpSheet.EnableCalculation;
            bool tmpDrawEnabled = Application.ScreenUpdating;
            tmpSheet.EnableCalculation = false;
            //            Application.ScreenUpdating = false;
            myIsInSheetChanged = true;
            Application.EnableEvents = false;
            try
            {
                List<PDCListObject> tmpLists = tmpInfo.ListsOnSheet;
                //iterate over all rectangular areas enclosed in the changed cell range
                foreach (Excel.Range tmpRange in range.Areas)
                {
                    try
                    {
                        AreaChanged(tmpLists, tmpSheet, tmpRange);
                    }
                    catch (Exception ee)
                    {
                        PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "SheetChanged", ee);
                    }
                }
            }
            finally
            {
                myIsInSheetChanged = false;
                Application.EnableEvents = tmpEventsEnabled;
                tmpSheet.EnableCalculation = tmpCalcEnabled;
                Application.ScreenUpdating = tmpDrawEnabled;
            }
        }

        /// <summary>
        /// Handles the change event of a rectangular area. All PDCListObjects on the sheet are examined,
        /// if they are affected.
        /// </summary>
        /// <param name="theLists">Possibly affected PDCLists</param>
        /// <param name="aSheet">The worksheet where the event occured</param>
        /// <param name="range">A rectangular range describing the changed cells</param>
        private void AreaChanged(List<PDCListObject> theLists, Excel.Worksheet aSheet, Excel.Range range)
        {
            int tmpRow = range.Row;
            int tmpColumn = range.Column;
            Rectangle tmpRectangle = new Rectangle(tmpColumn, tmpRow, range.Columns.Count, range.Rows.Count);

            SheetEvent tmpSheetEvent = SheetEvent.AsSheetEvent(aSheet, tmpRectangle);
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, tmpSheetEvent.ToString());
            foreach (PDCListObject tmpList in theLists)
            {
                if (tmpList is MeasurementPDCListObject)
                {
                    //MeasurementPDCListObjects are optimized PDCListObject with
                    //the assumption that their structure is immutable and that they
                    //are already full-sized. Therefore we can skip the event handling
                    continue;
                }
                if (AreaChanged(tmpList, aSheet, range, tmpRectangle, tmpSheetEvent))
                {
                    break;
                }
            }
        }


        /// <summary>
        /// Examines if the given PDCListObject is affected by the sheet change and delegates the event
        /// to the list if it is.
        /// </summary>
        /// <param name="aPDCList">A possibly affected PDCListObject</param>
        /// <param name="aSheet">The sheet where the change event occured</param>
        /// <param name="range">A rectangular Excel range with the changed cells</param>
        /// <param name="aRectangle"></param>
        /// <returns></returns>
        private bool AreaChanged(PDCListObject aPDCList, Excel.Worksheet aSheet, Excel.Range range, Rectangle aRectangle, SheetEvent anEvent)
        {
            if (anEvent.entireRow && aRectangle.Y < aPDCList.Rectangle.Y)
            {
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Added or Remove row above pdc list");
                aPDCList.UpdateRectangle();
            }
            Excel.Range tmpChangeRange = aPDCList.CalculateIntersection(aRectangle);
            if (tmpChangeRange == null)
            {
                return false;
            }
            if (tmpChangeRange.Rows.Count > 0 && tmpChangeRange.Columns.Count > 0)
            {
                aPDCList.CellsChanged(range, anEvent);
                return false;
            }
            return false;
        }

        /// <summary>
        /// Performs the enabling/disabling of the PDC menu entries.
        /// </summary>
        public void EnablePdcActions()
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "EnablePdcActions: WorkbookClosing:" + myIsWorbookClosing);
            object tmpActiveSheet = Application.ActiveSheet;
            if (tmpActiveSheet is Excel.Worksheet)
            {
                EnablePdcActions(GetSheetInfo((Excel.Worksheet)tmpActiveSheet), IsLoggedIn, ActiveSheetIsPDCSheet);
            }
            else
            {
                PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "EnablePdcActions: Not a worksheet");
                EnablePdcActions(null, IsLoggedIn, ActiveSheetIsPDCSheet);
            }
        }

        /// <summary>
        /// Performs the enabling/disabling of the PDC menu entries
        /// </summary>
        /// <param name="sheetInfo">SheetInfo of a PDC related sheet or null</param>
        /// <param name="isLoggedIn">Wether the user is logged in</param>
        /// <param name="isPDCSheet">Wether the active sheet is a main data entry sheet</param>
        void EnablePdcActions(SheetInfo sheetInfo, bool isLoggedIn, bool isPDCSheet)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "myIsWorkbookClosing=" + myIsWorbookClosing);
            if (myIsShuttingDown)
            {
                return;
            }

            Lib.UserInfo userInfo = myPdcService == null ? null : PdcService.UserInfo;
            bool tmpIsLoggedIn = isLoggedIn;
            bool isMainSheet = sheetInfo != null && sheetInfo.MainTable != null && sheetInfo.MainSheetInfo == null;
            bool mayUpload = isMainSheet && tmpIsLoggedIn && sheetInfo.TestDefinition.HasUploadPrivileges(userInfo) && userInfo.IsDataProvider();
            bool maySelect = isMainSheet && tmpIsLoggedIn && userInfo != null && userInfo.IsPDCUser;
            if (sheetInfo != null && sheetInfo.TestDefinition.Sourcesystem == "ICB")
            {
                maySelect = isMainSheet && tmpIsLoggedIn && userInfo != null && (userInfo.IsUserInRole(Settings.Default.ROLE_PDC_USER_ICB));
            }
            bool isNewSheet = isMainSheet && (!sheetInfo.MainTable.AlreadyUploaded);
            structureFormatAction.Enabled = tmpIsLoggedIn;
            versionInfoAction.Enabled = true;
            versionInfoAction.Visible = true;
            loginAction.Visible = !tmpIsLoggedIn;
            winLoginAction.Visible = !tmpIsLoggedIn;
            logoutAction.Visible = tmpIsLoggedIn;
            pdcCreateWorkbookAction.Enabled = tmpIsLoggedIn;

            uploadAction.Enabled = tmpIsLoggedIn && mayUpload && isNewSheet && !myIsWorbookClosing;
            updateAction.Enabled = tmpIsLoggedIn && mayUpload && !isNewSheet && !myIsWorbookClosing;
            deleteAction.Enabled = tmpIsLoggedIn && mayUpload && !isNewSheet && !myIsWorbookClosing;

            clearDataAction.Enabled = tmpIsLoggedIn && isMainSheet && !myIsWorbookClosing;
            validateAction.Enabled = tmpIsLoggedIn && isMainSheet && !myIsWorbookClosing;
            searchTestdataAction.Enabled = maySelect && !myIsWorbookClosing;
            compoundInfoAction.Enabled = tmpIsLoggedIn;
            retrieveMeasurementLevelEnabled = tmpIsLoggedIn && IsSingleMeasurementTableLoaded(sheetInfo) && sheetInfo != null &&
              !sheetInfo.AreMeasurementsLoaded && !isNewSheet && !myIsWorbookClosing;
            retrieveMeasurementLevelDataAction.Enabled = retrieveMeasurementLevelEnabled;
            if (retrieveMeasurementsContextMenuAction != null)
            {
                retrieveMeasurementsContextMenuAction.Enabled = retrieveMeasurementLevelEnabled;
            }
            
            SetStatusText(null);
            Globals.Ribbons.PdcDesignedRibbon.UpdateEnablement();
        }

        /// <summary>
        ///   Checks for a loaded single measurement table sheet.
        /// </summary>
        /// <param name="sheetInfo">
        ///   The sheet info object of the selected sheet.
        /// </param>
        /// <returns>
        ///   True, if the single measurement table sheet is loaded.
        ///   False, otherwise.
        /// </returns>
        private bool IsSingleMeasurementTableLoaded(SheetInfo sheetInfo)
        {
            if (sheetInfo == null) return false;
            if (sheetInfo.IsMainSheet)
            {
                if (sheetInfo.MainTable.MeasurementColumn == null) return false;
                return sheetInfo.MainTable.MeasurementColumn.HasSingleMeasurementTableHandler;
            }
            return false;
        }

        /// <summary>
        /// Returns true if the active sheet is a main data entry pdc sheet and false otherwise
        /// </summary>
        /// <returns></returns>
        public bool ActiveSheetIsPDCSheet
        {
            get
            {
                object tmpSheet = Application.ActiveSheet;
                if (tmpSheet is Excel.Worksheet)
                {
                    SheetInfo tmpSheetInfo = GetSheetInfo((Excel.Worksheet)tmpSheet);
                    return tmpSheetInfo != null && tmpSheetInfo.MainTable != null && tmpSheetInfo.MainSheetInfo == null;
                }
                return false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //ClearMenus();
            myIsShuttingDown = true;
            bool tmpFoundException = false;
            //Remove Eventhandlers
            if (myWorkbooks != null)
            {
                foreach (Excel.Workbook tmpWB in myWorkbooks.Keys)
                {
                    try
                    {
                        tmpWB.SheetChange -= SheetChanged;
                    }
                    catch (Exception ee)
                    {
                        if (!tmpFoundException)
                        {
                            tmpFoundException = true;
                            PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Shutdown: Exception while removing workbook event handlers", ee);
                        }
                    }
                }
            }
            try
            {

                Application.WorkbookActivate -= WorkbookActivated;
                Application.WorkbookDeactivate -= WorkbookDeactivated;
                Application.WorkbookOpen -= WorkbookOpened;
                Application.WorkbookBeforeClose -= WorkbookClosing;
                Application.SheetActivate -= WorksheetActivated;
                Application.SheetDeactivate -= WorksheetDeactivated;
                Application.WorkbookBeforeSave -= WorkbookSaving;
                Application.SheetBeforeRightClick -= SheetBeforeRightClick;
            }
            catch (Exception eee)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Shutdown: Exception while removing application event handlers", eee);
            }
            finally
            {
                ResetLoggedIn();

                if (mySheetMap != null)
                {
                    mySheetMap.Clear();
                }
                if (mySheetInfos != null)
                {
                    mySheetInfos.Clear();
                }

            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion

        /// <summary>
        /// Clears the given sheet, so that it can be transformed to a PDC sheet.
        /// </summary>
        /// <param name="tmpSheet"></param>
        internal void ClearSheet(Excel.Worksheet tmpSheet)
        {
            SheetInfo tmpInfo = GetSheetInfo(tmpSheet);
            if (tmpInfo != null && tmpInfo.MainTable != null && tmpInfo.TestDefinition != null)
            {
                Lib.Testdefinition tmpTD = tmpInfo.TestDefinition;
                PDCListObject tmpList = tmpInfo.MainTable;
                tmpList.Delete();
                RemoveSheetInfo(tmpInfo);
            }
            else if (tmpInfo != null && tmpInfo.MainSheetInfo != null)
            {
                throw new Exceptions.SubordinatePDCSheetException();
            }
            else
            {
                Excel.Range tmpUsedRange = tmpSheet.UsedRange;
                tmpUsedRange.Formula = "";
            }
        }

        /// <summary>
        /// Does the worksheet belong the specified workbook?
        /// </summary>
        /// <param name="aWorksheet"></param>
        /// <param name="aWorkbook"></param>
        /// <returns></returns>
        private bool BelongsToWorkbook(Excel.Worksheet aWorksheet, Excel.Workbook aWorkbook)
        {
            if (aWorksheet == null || aWorkbook == null)
            {
                return false;
            }
            try
            {
                object tmpSheetParent = aWorksheet.Parent;
                if (tmpSheetParent == aWorkbook)
                {
                    return true;
                }
                //We cannot rely that we always get the same VSTO wrapper for the same sheet/workbook
                //Therefore a name loop comparison is needed if the reference comparison fails to be sure.
                if (tmpSheetParent is Excel.Workbook)
                {
                    Excel.Workbook tmpSheetWb = (Excel.Workbook)tmpSheetParent;
                    string tmpFullName = aWorkbook.FullName;
                    if (tmpSheetWb.FullName == tmpFullName)
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Checking worksheet-workbook correlation", e);
                return true; //being conservative
            }
        }

        /// <summary>
        /// Tests if the specified Testdefinition is currently open in the same workbook
        /// </summary>
        /// <param name="aWorkbook">The workbook in which the Testdefinition shall be opened</param>
        /// <param name="selectedTd">The Testdefinition which shall be opened</param>
        /// <returns>True if the Testdefinition (version) is already opened in the same workbook</returns>
        internal bool TestOpen(Excel.Workbook aWorkbook, Lib.Testdefinition selectedTd)
        {
            if (selectedTd == null || selectedTd.TestNo == null)
            {
                return false;
            }
            RemoveNonExistingSheets();
            foreach (SheetInfo tmpInfo in mySheetInfos.Values)
            {
                if (tmpInfo.TestDefinition == null)
                {
                    continue;
                }
                if (selectedTd.TestNo == tmpInfo.TestDefinition.TestNo && selectedTd.Version == tmpInfo.TestDefinition.Version)
                { //same test definition version
                    if (BelongsToWorkbook(tmpInfo.ExcelSheet, aWorkbook))
                    { // same workbook
                        return true;
                    } //go on
                }
            }
            return false;
        }

        /// <summary>
        ///    Checks, if all sheets which belongs to one test are available
        /// </summary>
        /// <returns>
        ///    True, if all sheets are in the workbook. Otherwise false.
        /// </returns>
        internal bool AreAllSheetsForTheSelectedSheetAvailable()
        {
            // check first, if the current active Sheet is an PDC Sheet...
            SheetInfo activeSheetInfo = ExcelClient.ExcelUtils.TheUtils.ActiveSheetInfo;
            if (activeSheetInfo == null) return false;

            // cget all sheets belonging to the active sheet (referenced by the sheetinfo) 
            List<SheetInfo> additionalSheetInfos = activeSheetInfo.AdditionalSheets;

            // as far I as known, there is only one additional sheet (aka Measurementsheet) available! (to be checked!!!)
            // TODO
            if (additionalSheetInfos.Count > 1) return false;
            foreach (SheetInfo sheetInfo in additionalSheetInfos)
            {
                if (sheetInfo.IsSheetMissing())
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Removes any PDC sheets and SheetInfos which are not in use anymore.
        /// This explicit clean-up operation is necessary since Excel does not 
        /// provide related events which we could use.
        /// </summary>
        internal void RemoveNonExistingSheets()
        {
            if (mySheetInfos == null)
            {
                return;
            }

            PDCLogger.TheLogger.LogStarttime("CheckRemovedSheets", "Checking for removed PDC worksheets");
            bool tmpFinished = false;
            int tmpSheetInfoCount = mySheetInfos.Count;
            while (!tmpFinished)
            {
                tmpFinished = true;
                foreach (SheetInfo tmpSheetInfo in mySheetInfos.Values)
                {
                    if (!tmpSheetInfo.CheckStillExists())
                    {
                        tmpFinished = false;
                        break;
                    }
                }
                int tmpNewCount = mySheetInfos.Count;
                if (tmpNewCount >= tmpSheetInfoCount)
                {
                    break;
                }
            }
            //Now clean sheetMap
            Dictionary<Excel.Worksheet, SheetInfo> tmpCopy = new Dictionary<Excel.Worksheet, SheetInfo>(mySheetMap);
            foreach (KeyValuePair<Excel.Worksheet, SheetInfo> tmpPair in tmpCopy)
            {
                if (!mySheetInfos.ContainsValue(tmpPair.Value))
                {
                    mySheetMap.Remove(tmpPair.Key);
                    continue;
                }
                try
                {
                    string tmpUnused = tmpPair.Key.Name;
                }
#pragma warning disable 0168
                catch (Exception e)
                {
                    mySheetMap.Remove(tmpPair.Key);
                }
#pragma warning restore 0168
            }
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Found " + (tmpCopy.Count - mySheetMap.Count) + " PDC sheets which were removed");
            //Now clean sheet infos and Measurement handler
            foreach (SheetInfo tmpSheetInfo in mySheetInfos.Values)
            {
                // todo Implement SingleMeasurementTableHandler
                if (tmpSheetInfo.IsMainSheet && tmpSheetInfo.MainTable.MeasurementColumn != null
                 && tmpSheetInfo.MainTable.MeasurementColumn.ParamHandler is Predefined.MultipleMeasurementTableHandler)
                {
                    Predefined.MultipleMeasurementTableHandler tmpHandler = (Predefined.MultipleMeasurementTableHandler)tmpSheetInfo.MainTable.MeasurementColumn.ParamHandler;
                    tmpHandler.CleanupReferences();
                }
            }
            PDCLogger.TheLogger.LogStoptime("CheckRemovedSheets", "Checking for removed PDC worksheets");
        }

        /// <summary>
        /// Create a state object for the specified workbook.
        /// </summary>
        /// <param name="aWorkbook"></param>
        /// <returns></returns>
        private WorkbookState CreateStateObject(Excel.Workbook aWorkbook)
        {
            WorkbookState tmpState = new WorkbookState();
            List<SheetInfo> tmpSheetInfos = new List<SheetInfo>();
            RemoveNonExistingSheets();
            foreach (SheetInfo tmpSheetInfo in mySheetInfos.Values)
            {
                if (tmpSheetInfo.ExcelSheet == null)
                {
                    continue;
                }
                try
                {
                    Excel.Workbook tmpParent = (Excel.Workbook)tmpSheetInfo.ExcelSheet.Parent;
                    if (tmpParent.Name == aWorkbook.Name)
                    {
                        tmpSheetInfos.Add(tmpSheetInfo);
                    }
                }
                catch (Exception e)
                {
                    tmpSheetInfo.Delete();
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "PDC Sheet probably removed Id:" + tmpSheetInfo.Identifier, e);
                }
            }
            if (tmpSheetInfos.Count == 0)
            {
                return null;
            }
            tmpState.PicklistHandler = PicklistHandler.ThePicklistHandler(aWorkbook, true);
            tmpState.SheetInfos = tmpSheetInfos;
            return tmpState;
        }

        /// <summary>
        /// Saves the PDC workbook state into the workbook
        /// </summary>
        /// <param name="aWorkbook"></param>
        public void saveWorkbookState(Excel.Workbook aWorkbook)
        {

            //Get PDC data related to the workbook
            WorkbookState tmpState = CreateStateObject(aWorkbook);
            if (tmpState == null || tmpState.SheetInfos == null || tmpState.SheetInfos.Count == 0)
            {
                return;
            }
            //Serialize PDC data
            BinaryFormatter tmpSerializer = new BinaryFormatter();
            tmpSerializer.AssemblyFormat = FormatterAssemblyStyle.Simple;
            MemoryStream tmpStream = new MemoryStream();
            tmpSerializer.Serialize(tmpStream, tmpState);
            PicklistHandler tmpHandler = PicklistHandler.ThePicklistHandler(aWorkbook, true);
            Excel.Worksheet tmpSheet = tmpHandler.Worksheet;
            UpdateWorkbookVersion(aWorkbook, tmpSheet, false);
            byte[] tmpBytes = tmpStream.ToArray();
            //Save serialized data
            StoreEncodedCustomProperty(tmpSheet, CUSTOM_PROPERTY_STORE, tmpBytes);
        }

        public bool checkVersion(string version)
        {
            return PdcService.CheckVersion(version);
        }

        public void updatePrivileges()
        {
            foreach (SheetInfo tmpSheetInfo in mySheetInfos.Values)
            {
                // update privileges
                PdcService.SetPrivilegesForTestdefinition(tmpSheetInfo.TestDefinition);
            }
        }
        /// <summary>
        /// Searches the given worksheet for the specified property.
        /// </summary>
        /// <param name="aSheet">The sheet to examine</param>
        /// <param name="aPropertyName">The property to search for</param>
        /// <returns>The named CustomProperty or null if it could not be found</returns>
        private Excel.CustomProperty FindCustomProperty(Excel.Worksheet aSheet, string aPropertyName)
        {
            try
            {
                IEnumerator tmpCustomProps = aSheet.CustomProperties.GetEnumerator();
                while (tmpCustomProps.MoveNext())
                {
                    Excel.CustomProperty tmpProperty = (Excel.CustomProperty)tmpCustomProps.Current;
                    if (tmpProperty.Name == aPropertyName)
                    {
                        return tmpProperty;
                    }
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Searching PDC CustomProperty", e);
            }
            return null;
        }

        /// <summary>
        /// remove Custom property
        /// </summary>
        /// <param name="aSheet">The sheet to examine</param>
        /// <param name="aPropertyName">The property to search for</param>
        /// <returns>The named CustomProperty or null if it could not be found</returns>
        private void RemoveCustomProperty(Excel.Worksheet aSheet, string aPropertyName)
        {
            try
            {
                IEnumerator tmpCustomProps = aSheet.CustomProperties.GetEnumerator();
                while (tmpCustomProps.MoveNext())
                {
                    Excel.CustomProperty tmpProperty = (Excel.CustomProperty)tmpCustomProps.Current;
                    if (tmpProperty.Name == aPropertyName)
                    {
                        tmpProperty.Delete();

                    }
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Searching PDC CustomProperty", e);
            }

        }

        /// <summary>
        /// Reads the Base64 encoded custom property and returns it as a byte[]
        /// </summary>
        /// <param name="aSheet"></param>
        /// <param name="aPropertyName"></param>
        /// <returns></returns>
        private byte[] ReadEncodedCustomProperty(Excel.Worksheet aSheet, string aPropertyName)
        {
            Excel.CustomProperty tmpProperty = FindCustomProperty(aSheet, aPropertyName);
            if (tmpProperty == null)
            {
                return null;
            }
            string tmpValue = "" + (tmpProperty.Value ?? "");
            return Convert.FromBase64String(tmpValue);
        }

        void Application_SheetBeforeDoubleClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            Cancel = false;
            try
            {
                Target.Value2 = Target.Value2;
            }
            catch (Exception)
            {
                Cancel = true;
            }
        }

        /// <summary>
        /// builds up the context menu
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        /// <param name="Cancel"></param>
        void SheetBeforeRightClick(object Sh, Microsoft.Office.Interop.Excel.Range Target, ref bool Cancel)
        {
            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, $"{nameof(SheetBeforeRightClick)} - Setting up PDC context menu.");
            Cancel = false;

            Office.CommandBar tmpContext = Application.CommandBars["Cell"];
            Office.CommandBar tmpListContext = Application.CommandBars["List Range Popup"];
            SetupCompoundInfoContextMenu(tmpContext, tmpListContext);
        }

        /// <summary>
        /// Stores the byte[] as Base64 encoded Custom property in the sheet.
        /// </summary>
        /// <param name="aSheet"></param>
        /// <param name="aPropertyName"></param>
        /// <param name="aValue"></param>
        private void StoreEncodedCustomProperty(Excel.Worksheet aSheet, string aPropertyName, byte[] aValue)
        {
            Excel.CustomProperty tmpProperty = FindCustomProperty(aSheet, aPropertyName);
            if (tmpProperty != null)
            {
                tmpProperty.Delete();
            }
            string tmpEncoded = null;
            tmpEncoded = Convert.ToBase64String(aValue);
            Excel.CustomProperties tmpProperties = aSheet.CustomProperties;
            tmpProperties.Add(aPropertyName, tmpEncoded);
        }

        /// <summary>
        /// Reinitializes the PDC internal structure for the specified workbook when it was loaded from disk.
        /// </summary>
        /// <param name="workbookState"></param>
        /// <param name="aWorkbook"></param>
        /// <param name="aPicklistSheet"></param>
        private void Reinitialize(WorkbookState workbookState, Excel.Workbook aWorkbook, Excel.Worksheet aPicklistSheet, string aVersion)
        {
            PicklistHandler.Deserialize(workbookState.PicklistHandler, aWorkbook, aPicklistSheet);
            foreach (SheetInfo tmpSheetInfo in workbookState.SheetInfos)
            {
                // update privileges
                PdcService.SetPrivilegesForTestdefinition(tmpSheetInfo.TestDefinition);

                if (mySheetInfos.ContainsKey(tmpSheetInfo.Identifier))
                {
                    mySheetInfos[tmpSheetInfo.Identifier] = tmpSheetInfo;
                }
                else
                {
                    mySheetInfos.Add(tmpSheetInfo.Identifier, tmpSheetInfo);
                }
            }
            //Init connection to sheet
            IEnumerator tmpSheets = aWorkbook.Worksheets.GetEnumerator();
            while (tmpSheets.MoveNext())
            {
                object tmpSheetCand = tmpSheets.Current;
                if (!(tmpSheetCand is Excel.Worksheet))
                {
                    continue;
                }
                Excel.Worksheet tmpSheet = (Excel.Worksheet)tmpSheetCand;
                SheetInfo tmpSheetInfo = GetSheetInfo(tmpSheet);
                tmpSheetInfo?.InitSheet(tmpSheet, aVersion);
            }
            // Now delete sheet info which excel sheets were deleted in the meantime
            foreach (SheetInfo tmpAdded in workbookState.SheetInfos)
            {
                if (tmpAdded.ExcelSheet == null && tmpAdded.Identifier != null && mySheetInfos.ContainsKey(tmpAdded.Identifier))
                {
                    mySheetInfos.Remove(tmpAdded.Identifier);
                }
            }
            aWorkbook.SheetChange += SheetChanged;
        }

        /// <summary>
        /// This delegate method is for some odd reasons necessary to be able
        /// to deserialize the object graph caused by strong naming of the assemblies.
        /// The CLR will otherwise refuse to load workbooks from old add-in version due 
        /// to version mismatch.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        static System.Reflection.Assembly LoadComponentAssembly(object sender, ResolveEventArgs args)
        {
            string simpleName = args.Name.Substring(0, args.Name.IndexOf(','));
            PDCLogger.TheLogger.LogDebugMessage($"{nameof(LoadComponentAssembly)}",$"loading {simpleName}");
            if (simpleName.ToUpper().StartsWith("PDCEXCEL"))
            {
                return typeof(PDCExcelAddIn).Assembly;
            }
            if (simpleName.ToUpper().StartsWith("PDCLIB"))
            {
                return typeof(Lib.PDCService).Assembly;
            }
            System.Reflection.Assembly tmpAssembly = System.Reflection.Assembly.LoadFrom(args.Name);
            return tmpAssembly;
        }

        /// <summary>
        /// Sets a status message on the status bar
        /// </summary>
        /// <param name="aStatusMessage"></param>
        internal void SetStatusText(string aStatusMessage)
        {
            try
            {
                if (Application.DisplayStatusBar)
                {
                    Application.StatusBar = aStatusMessage;
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Failed to set message on status bar", e);
            }
        }

        /// <summary>
        /// Enables Excel event handling, screen updating, ... and resets the cursor
        /// </summary>
        internal void EnableExcel()
        {
            try
            {
                EventsEnabled = true;
                Application.EnableEvents = true;
                Application.Cursor = Excel.XlMousePointer.xlDefault;
                Application.ScreenUpdating = true;
                if (Application.ActiveSheet is Excel.Worksheet)
                {
                    ((Excel.Worksheet)Application.ActiveSheet).EnableCalculation = true;
                }
                SetStatusText(null);
            }
            catch (Exception e) //Paranoid
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Enable Excel", e);
            }
        }

        /// <summary>
        /// Initializes a VBA module. The module is necessary since the only way to define menu shortcuts is to
        /// have a VBA procedure calling the menu.
        /// </summary>
        /// <param name="workbook"></param>
        public void InitVBACode(Excel.Workbook workbook)
        {
            try
            {
                // If you  cant believe that this makes any sense have a look at the discussion at
                // http://blogs.msdn.com/andreww/archive/2007/01/15/vsto-add-ins-comaddins-and-requestcomaddinautomationservice.aspx
                //Sample Code to add VBA makro code to a workbook
                VBE._VBComponent tmpModule = workbook.VBProject.VBComponents.Add(Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                Type tmpType = typeof(VBA.VBAInterface);
                int tmpMajor = tmpType.Assembly.GetName().Version.Major;
                int tmpMinor = tmpType.Assembly.GetName().Version.Minor;

                object[] tmpGuids = tmpType.Assembly.GetCustomAttributes(typeof(GuidAttribute), true);
                workbook.VBProject.References.AddFromGuid("{" + ((GuidAttribute)tmpGuids[0]).Value + "}", tmpMajor, tmpMinor);
                tmpModule.CodeModule.AddFromString(
                    "Private Sub PDC_LoadButton_Click() \n" +
                    "Dim addin As Office.COMAddIn\n" +
                    "Dim aO As Object\n" +
                    "Set addin = Application.COMAddIns(\"PDCExcelAddIn\")\n" +
                    "Set aO = addin.Object\n" +
                    "aO.RetrieveMeasurementLevelData\n" +
                    "End Sub\n");
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "VBA Shortcut Module", e);
            }
        }

        /// <summary>
        /// Retrieves the shortcut for the RetrieveMeasurementLevelData action from the specified worksheet or
        /// the default value from the user configuration.
        /// "^d" is the default value, unless reconfigured by the user.
        /// </summary>
        internal string RetrieveMeasurementLevelDataShortcut()
        {
            return UserConfiguration.TheConfiguration.GetProperty(UserConfiguration.PROP_RETRIEVE_MEASUREMENTS_SHORTCUT, "^d");
        }
        /// <summary>
        /// Retrieves the shortcut for the RetrieveMeasurementLevelData action from the specified worksheet or
        /// the default value from the user configuration.
        /// "CTRL-D" is the default value, unless reconfigured by the user.
        /// </summary>
        internal string RetrieveMeasurementLevelDataShortcutText()
        {
            return UserConfiguration.TheConfiguration.GetProperty(UserConfiguration.PROP_RETRIEVE_MEASUREMENTS_SHORTCUT_TEXT, "CTRL-D");
        }
        /// <summary>
        /// May be used in later versions to register a real COM-enabled Singleton within Excel.
        /// So that VSTO-code, VBA and Excel UDFs use the same instances. AddIn assembly needs to be 
        /// COM enabled and COMSingleton needs the appropriate implementation.
        /// </summary>
        /// <returns></returns>
        protected override object RequestComAddInAutomationService()
        {
            return VBA.COMSingleton.Singleton;
        }

        public void StartAutoUpdater()
        {
            // Check for updates
            string installPath = VersionInfoAction.GetInstallPath();
            string path = Path.Combine(installPath, "AutoUpdater.exe");

            if (File.Exists(path))
            {
                PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL,
                string.Format("Starting '{0}' in working dir '{1}'", path, installPath), Level.Info);

                ProcessStartInfo startInfo = new ProcessStartInfo(path);
                startInfo.WorkingDirectory = installPath;

                // install update if available
                Process.Start(startInfo);
            }
        }

        #region Action accessors

        internal PDCAction WindowsLoginAction
        {
            get { return winLoginAction; }
        }

        internal PDCAction LoginAction
        {
            get { return loginAction; }
        }

        internal PDCAction LogoutAction
        {
            get { return logoutAction; }
        }

        internal PDCAction CreateWorkbookAction
        {
            get { return pdcCreateWorkbookAction; }
        }
        internal PDCAction ClearDataAction
        {
            get
            {
                return clearDataAction;
            }
        }

        internal PDCAction UpdateDataAction
        {
            get
            {
                return updateAction;
            }
        }

        internal PDCAction UploadDataAction
        {
            get
            {
                return uploadAction;
            }
        }

        internal PDCAction DeleteAction
        {
            get
            {
                return deleteAction;
            }
        }

        internal PDCAction RetrieveMeasurementLevelDataAction
        {
            get
            {
                return retrieveMeasurementLevelDataAction;
            }
        }

        internal PDCAction CompoundDataAction
        {
            get { return compoundInfoAction; }
        }

        internal PDCAction StructureFormatAction
        {
            get { return structureFormatAction; }
        }

        internal PDCAction VersionAction
        {
            get { return versionInfoAction; }
        }
        internal PDCAction SearchTestdataAction
        {
            get
            {
                return searchTestdataAction;
            }
        }

        internal PDCAction ValidateAction
        {
            get
            {
                return validateAction;
            }
        }
        #endregion
    }
}
