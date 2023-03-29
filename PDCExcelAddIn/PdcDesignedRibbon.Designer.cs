using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    //partial class PdcDesignedRibbon : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    partial class PdcDesignedRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Default Constructor
        /// </summary>
        public PdcDesignedRibbon(): base(Globals.Factory.GetRibbonFactory())
        {
            if (!DesignMode)
            {
                PDCExcelAddIn.SetupLanguage();
            }
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PdcDesignedRibbon));
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.pdctab2 = this.Factory.CreateRibbonTab();
            this.loginGroup = this.Factory.CreateRibbonGroup();
            this.loginWindowsButton = this.Factory.CreateRibbonButton();
            this.loginOtherButton = this.Factory.CreateRibbonButton();
            this.logoutButton = this.Factory.CreateRibbonButton();
            this.workbookGroup = this.Factory.CreateRibbonGroup();
            this.newWorksheetButton = this.Factory.CreateRibbonButton();
            this.uploadButton = this.Factory.CreateRibbonButton();
            this.validateButton = this.Factory.CreateRibbonButton();
            this.searchButton = this.Factory.CreateRibbonButton();
            this.clearButton = this.Factory.CreateRibbonButton();
            this.uploadChangesButton = this.Factory.CreateRibbonButton();
            this.placeTaker = this.Factory.CreateRibbonLabel();
            this.deleteButton = this.Factory.CreateRibbonButton();
            this.retrieveMeasurementsButton = this.Factory.CreateRibbonButton();
            this.chemicalDataButton = this.Factory.CreateRibbonButton();
            this.formatCompoundButton = this.Factory.CreateRibbonButton();
            this.infoGroup = this.Factory.CreateRibbonGroup();
            this.versionInfoButton = this.Factory.CreateRibbonButton();
            this.docuButton = this.Factory.CreateRibbonButton();
            this.contactButton = this.Factory.CreateRibbonButton();
            this.pdctab2.SuspendLayout();
            this.loginGroup.SuspendLayout();
            this.workbookGroup.SuspendLayout();
            this.infoGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // pdctab2
            // 
            this.pdctab2.Groups.Add(this.loginGroup);
            this.pdctab2.Groups.Add(this.workbookGroup);
            this.pdctab2.Groups.Add(this.infoGroup);
            resources.ApplyResources(this.pdctab2, "pdctab2");
            this.pdctab2.Name = "pdctab2";
            this.pdctab2.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabInsert");
            // 
            // loginGroup
            // 
            this.loginGroup.Items.Add(this.loginWindowsButton);
            this.loginGroup.Items.Add(this.loginOtherButton);
            this.loginGroup.Items.Add(this.logoutButton);
            resources.ApplyResources(this.loginGroup, "loginGroup");
            this.loginGroup.Name = "loginGroup";
            // 
            // loginWindowsButton
            // 
            this.loginWindowsButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.loginWindowsButton, "loginWindowsButton");
            this.loginWindowsButton.Name = "loginWindowsButton";
            this.loginWindowsButton.OfficeImageId = "GroupPagePermissionsActions";
            this.loginWindowsButton.ShowImage = true;
            this.loginWindowsButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.LoginWindows;
            this.loginWindowsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // loginOtherButton
            // 
            this.loginOtherButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.loginOtherButton, "loginOtherButton");
            this.loginOtherButton.Name = "loginOtherButton";
            this.loginOtherButton.OfficeImageId = "GroupPermissionGroupsEdit";
            this.loginOtherButton.ShowImage = true;
            this.loginOtherButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.LoginOther;
            this.loginOtherButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // logoutButton
            // 
            this.logoutButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.logoutButton, "logoutButton");
            this.logoutButton.Name = "logoutButton";
            this.logoutButton.OfficeImageId = "ContactPictureMenu";
            this.logoutButton.ShowImage = true;
            this.logoutButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.Logout;
            this.logoutButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // workbookGroup
            // 
            this.workbookGroup.Items.Add(this.newWorksheetButton);
            this.workbookGroup.Items.Add(this.uploadButton);
            this.workbookGroup.Items.Add(this.validateButton);
            this.workbookGroup.Items.Add(this.searchButton);
            this.workbookGroup.Items.Add(this.clearButton);
            this.workbookGroup.Items.Add(this.uploadChangesButton);
            this.workbookGroup.Items.Add(this.placeTaker);
            this.workbookGroup.Items.Add(this.deleteButton);
            this.workbookGroup.Items.Add(this.retrieveMeasurementsButton);
            this.workbookGroup.Items.Add(this.chemicalDataButton);
            this.workbookGroup.Items.Add(this.formatCompoundButton);
            this.workbookGroup.Name = "workbookGroup";
            // 
            // newWorksheetButton
            // 
            this.newWorksheetButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.newWorksheetButton, "newWorksheetButton");
            this.newWorksheetButton.Name = "newWorksheetButton";
            this.newWorksheetButton.OfficeImageId = "CreateTable";
            this.newWorksheetButton.ShowImage = true;
            this.newWorksheetButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.NewPdcWorksheet;
            this.newWorksheetButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // uploadButton
            // 
            this.uploadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.uploadButton, "uploadButton");
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.OfficeImageId = "DatabaseSqlServer";
            this.uploadButton.ShowImage = true;
            this.uploadButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.UploadTestData;
            this.uploadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // validateButton
            // 
            this.validateButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.validateButton, "validateButton");
            this.validateButton.Name = "validateButton";
            this.validateButton.OfficeImageId = "DataValidation";
            this.validateButton.ShowImage = true;
            this.validateButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.ValidateWorksheet;
            this.validateButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // searchButton
            // 
            this.searchButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.searchButton, "searchButton");
            this.searchButton.Name = "searchButton";
            this.searchButton.OfficeImageId = "DatabaseModelingReverse";
            this.searchButton.ShowImage = true;
            this.searchButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.SearchTestData;
            this.searchButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // clearButton
            // 
            this.clearButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.clearButton, "clearButton");
            this.clearButton.Name = "clearButton";
            this.clearButton.OfficeImageId = "TableEraser";
            this.clearButton.ShowImage = true;
            this.clearButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.ClearWorksheet;
            this.clearButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // uploadChangesButton
            // 
            resources.ApplyResources(this.uploadChangesButton, "uploadChangesButton");
            this.uploadChangesButton.Name = "uploadChangesButton";
            this.uploadChangesButton.OfficeImageId = "GanttDataImport";
            this.uploadChangesButton.ShowImage = true;
            this.uploadChangesButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.UploadChanges;
            this.uploadChangesButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // placeTaker
            // 
            resources.ApplyResources(this.placeTaker, "placeTaker");
            this.placeTaker.Name = "placeTaker";
            // 
            // deleteButton
            // 
            resources.ApplyResources(this.deleteButton, "deleteButton");
            this.deleteButton.Name = "deleteButton";
            this.deleteButton.OfficeImageId = "DeleteRows";
            this.deleteButton.ShowImage = true;
            this.deleteButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.DeleteData;
            this.deleteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // retrieveMeasurementsButton
            // 
            resources.ApplyResources(this.retrieveMeasurementsButton, "retrieveMeasurementsButton");
            this.retrieveMeasurementsButton.Name = "retrieveMeasurementsButton";
            this.retrieveMeasurementsButton.OfficeImageId = "ChartTrendline";
            this.retrieveMeasurementsButton.ShowImage = true;
            this.retrieveMeasurementsButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.RetrieveMeasurements;
            this.retrieveMeasurementsButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // chemicalDataButton
            // 
            resources.ApplyResources(this.chemicalDataButton, "chemicalDataButton");
            this.chemicalDataButton.Name = "chemicalDataButton";
            this.chemicalDataButton.OfficeImageId = "SelectColumn";
            this.chemicalDataButton.ShowImage = true;
            this.chemicalDataButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.RetrieveChemicalData;
            this.chemicalDataButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // formatCompoundButton
            // 
            resources.ApplyResources(this.formatCompoundButton, "formatCompoundButton");
            this.formatCompoundButton.Name = "formatCompoundButton";
            this.formatCompoundButton.OfficeImageId = "FormatColumns";
            this.formatCompoundButton.ShowImage = true;
            this.formatCompoundButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.FormatChemicalDataSettings;
            this.formatCompoundButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // infoGroup
            // 
            this.infoGroup.Items.Add(this.versionInfoButton);
            this.infoGroup.Items.Add(this.docuButton);
            this.infoGroup.Items.Add(this.contactButton);
            this.infoGroup.Name = "infoGroup";
            // 
            // versionInfoButton
            // 
            resources.ApplyResources(this.versionInfoButton, "versionInfoButton");
            this.versionInfoButton.Name = "versionInfoButton";
            this.versionInfoButton.OfficeImageId = "Info";
            this.versionInfoButton.ShowImage = true;
            this.versionInfoButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.PdcVersionInfo;
            this.versionInfoButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // docuButton
            // 
            resources.ApplyResources(this.docuButton, "docuButton");
            this.docuButton.Name = "docuButton";
            this.docuButton.OfficeImageId = "DeveloperReference";
            this.docuButton.ShowImage = true;
            this.docuButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.Documentation;
            this.docuButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // contactButton
            // 
            resources.ApplyResources(this.contactButton, "contactButton");
            this.contactButton.Name = "contactButton";
            this.contactButton.OfficeImageId = "Call";
            this.contactButton.ShowImage = true;
            this.contactButton.SuperTip = global::BBS.ST.BHC.BSP.PDC.ExcelClient.Tooltips.ContactSupport;
            this.contactButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ButtonClicked);
            // 
            // PdcDesignedRibbon
            // 
            this.Name = "PdcDesignedRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.pdctab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PdcDesignedRibbon_Load);
            this.pdctab2.ResumeLayout(false);
            this.pdctab2.PerformLayout();
            this.loginGroup.ResumeLayout(false);
            this.loginGroup.PerformLayout();
            this.workbookGroup.ResumeLayout(false);
            this.workbookGroup.PerformLayout();
            this.infoGroup.ResumeLayout(false);
            this.infoGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        private void PdcDesignedRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab pdctab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup loginGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginWindowsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton loginOtherButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton logoutButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup workbookGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton searchButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton clearButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton uploadChangesButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton deleteButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton retrieveMeasurementsButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton formatCompoundButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton chemicalDataButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton newWorksheetButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton uploadButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton validateButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel placeTaker;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup infoGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton versionInfoButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton docuButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton contactButton;
        private ToolTip toolTip1;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal PdcDesignedRibbon PdcDesignedRibbon
        {
            get { return this.GetRibbon<PdcDesignedRibbon>(); }
        }
    }

}
