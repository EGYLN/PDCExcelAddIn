using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    /// <summary>
    /// Base class of all pdc commands
    /// </summary>
    [ComVisible(false)]
    public abstract class PDCAction
    {
        protected object missing = Type.Missing;
        private bool myEnabled = true;
        private bool myVisible = true;
        protected Dictionary<string, int> myActionTags = new Dictionary<string, int>();
        protected bool myBeginsGroup = false;
        protected Dictionary<object, Office.CommandBarButton> myMenuMap =
          new Dictionary<object, Microsoft.Office.Core.CommandBarButton>();
        protected string myButtonText;
        protected string myCommandBarText;

        protected string shortcut;

        /// <summary>
        /// Tag which identifies the action in Excels menusystem
        /// </summary>
        protected string myBaseActionTag;

        /// <summary>
        /// Must be implemented by the concrete subclasses with the necessary business logic
        /// </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="interactive"></param>
        /// <returns></returns>
        internal abstract ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive);

        #region constructor
        protected PDCAction(string aButtonText, string anActionTag, bool beginGroup)
        {
            myBeginsGroup = beginGroup;
            myBaseActionTag = anActionTag;
            myButtonText = aButtonText;
        }
        #endregion
        #region Properties
        public string CommandBarText
        {
            get
            {
                return myCommandBarText;
            }
            protected set
            {
                myCommandBarText = value;
            }
        }
        #endregion
        #region methods

        #region AddToMenu
        public void AddToMenu(object aCommandBar, string aMenuTag)
        {
            Office.CommandBarButton tmpButton = CreateButton(aCommandBar, myBeginsGroup, aMenuTag + "_" + myBaseActionTag);
            Initialize(tmpButton);
            string caption = "";
            if (myButtonText != null)
            {
                caption = myButtonText;
            }
            if (Shortcut != null)
            {
                caption += " (" + Shortcut + ")";
            }
            tmpButton.Caption = caption;
            myMenuMap.Add(aCommandBar, tmpButton);
        }
        #endregion

        #region CanPerformAction
        /// <summary>
        /// checks wether the Action can be performed. 
        /// </summary>
        virtual protected bool CanPerformAction(out ActionStatus actionStatus, bool interactive)
        {
            actionStatus = null;
            return true;
        }
        #endregion

        #region CheckPDCSheet
        /// <summary>
        /// Checks if the specified SheetInfo is associated with a PDC data entry sheet
        /// </summary>
        /// <param name="aSheetInfo">May be null</param>
        internal void CheckPDCSheet(SheetInfo aSheetInfo)
        {
            if (aSheetInfo == null || aSheetInfo.MainTable == null || aSheetInfo.MainSheetInfo != null)
            {
                throw new Exceptions.NoPDCSheetException();
            }
            if (aSheetInfo.IsMainSheet)
            {
                if (aSheetInfo.MainTable.MeasurementColumn != null && aSheetInfo.MainTable.MeasurementColumn.HasSingleMeasurementTableHandler)
                {
                    if (!aSheetInfo.MainTable.MeasurementRangeExists)
                    {
                        throw new Exceptions.NoPDCSheetException();
                    }
                }
            }
        }
        #endregion

        #region CreateButton
        private Office.CommandBarButton CreateButton(object aParent, bool beginGroup, string aTag)
        {
            Office.CommandBarButton tmpButton = null;
            Office.CommandBarControls controls = null;
            if (aParent is Office.CommandBarPopup)
            {
                controls = ((Office.CommandBarPopup)aParent).Controls;
            }
            else
            {
                controls = ((Office.CommandBar)aParent).Controls;
            }
            tmpButton = (Office.CommandBarButton)controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);

            tmpButton.BeginGroup = beginGroup;
            if (aTag != null)
            {
                tmpButton.Tag = aTag;
                myActionTags.Add(aTag, tmpButton.Id);
            }
            tmpButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(HandleActionEvent);
            return tmpButton;
        }
        #endregion

        #region GetActiveSheetInfo
        internal SheetInfo GetActiveSheetInfo()
        {
            object tmpActiveSheet = AddIn.Application.ActiveSheet;
            if (tmpActiveSheet is Excel.Worksheet)
            {
                return AddIn.GetSheetInfo((Excel.Worksheet)tmpActiveSheet);
            }
            return null;
        }
        #endregion

        #region GetSheetInfo
        internal SheetInfo GetSheetInfo(Excel.Worksheet aSheet)
        {
            return AddIn.GetSheetInfo(aSheet);
        }
        #endregion

        #region Initialize
        /// <summary>
        /// Called after a new button was created for the action. Subclasses may add their own initializations
        /// </summary>
        /// <param name="aButton">A new button for the action</param>
        protected virtual void Initialize(Office.CommandBarButton aButton)
        {
        }
        #endregion

        #region ResultToActionStatus
        /// <summary>
        /// Convenience method to handle the most common result types of a PDCAction.
        /// Handles (localized) string, PDCMessage, PDCMessage[], Exception and null
        /// </summary>
        /// <param name="aResult"></param>
        /// <param name="aMessageTitle">Used as the dialog title, if a message box is shown.</param>
        /// <param name="interactive">A (localized) string is displayed in a message dialog, if this parameter is set to true</param>
        /// <returns></returns>
        protected virtual ActionStatus ResultToActionStatus(object aResult, string aMessageTitle, bool interactive)
        {
            if (aResult is string)
            {
                if (interactive)
                {
                    MessageBox.Show((string)aResult, aMessageTitle, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                return new ActionStatus(new Lib.PDCMessage((string)aResult, Lib.PDCMessage.TYPE_ERROR));
            }
            if (aResult is Exception)
            {
                throw (Exception)aResult;
            }
            if (aResult is Lib.PDCMessage[])
            {
                return new ActionStatus((Lib.PDCMessage[])aResult);
            }
            if (aResult is List<Lib.PDCMessage>)
            {
                return new ActionStatus(((List<Lib.PDCMessage>)aResult).ToArray());
            }
            return new ActionStatus();
        }
        #endregion

        #region SafeForCellEditingMode
        /// <summary>
        /// Returns true if the action can be executing even if Excel is in cell editing mode.
        /// </summary>
        /// <returns></returns>
        protected virtual bool SafeForCellEditingMode()
        {
            return false;
        }
        #endregion

        #endregion

        #region properties
        #region Shortcut
        /// <summary>
        /// Menu shortcut (text) which should be displayed on the menu button. Beware that it is necessary to
        /// register the shortcut on Excel application with an appropriate VBA procedure to make the shortcut really work.
        /// </summary>
        internal string Shortcut
        {
            get
            {
                return shortcut;
            }
            set
            {
                shortcut = value;
                foreach (Office.CommandBarButton button in myMenuMap.Values)
                {
                    string caption = button.Caption;
                    if (caption != null && caption.StartsWith(myButtonText))
                    {
                        caption = myButtonText;
                        if (shortcut != null)
                        {
                            caption += " (" + Shortcut + ")";
                        }
                        button.Caption = caption;
                    }
                }
            }
        }
        #endregion
        #region AddIn
        protected PDCExcelAddIn AddIn
        {
            get
            {
                return Globals.PDCExcelAddIn;
            }
        }
        #endregion

        #region Application
        protected Excel.Application Application
        {
            get
            {
                return Globals.PDCExcelAddIn.Application;
            }
        }
        #endregion

        #region Enabled
        public bool Enabled
        {
            get
            {
                return myEnabled;
            }
            set
            {
                myEnabled = value;
                try
                {
                    foreach (Office.CommandBarButton tmpButton in myMenuMap.Values)
                    {
                        tmpButton.Enabled = value;
                    }
                }
                catch (COMException e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_COM, string.Format("Failed to set enablement of {0} to {1}", myButtonText, value), e);
                    //throw;
                }
            }
        }
        #endregion

        #region PDCService
        protected Lib.PDCService PDCService
        {
            get
            {
                return Globals.PDCExcelAddIn.PdcService;
            }
        }
        #endregion

        #region PerformAction
        internal ActionStatus PerformAction(bool interactive)
        {
            if (interactive)
            {
                Application.EnableEvents = true;
            }


            // checks wether the user is editing a excel cell.
            if (!SafeForCellEditingMode() && ExcelUtils.TheUtils.IsInCellEditingMode())
            {
                bool tmpCancel = true;
                if (UserConfiguration.TheConfiguration.GetBooleanProperty(UserConfiguration.PROP_ENABLE_EXIT_CELLEDIT, false))
                {
                    ExcelUtils.TheUtils.ExitCellEditing();
                }
                tmpCancel = ExcelUtils.TheUtils.IsInCellEditingMode();
                if (tmpCancel)
                {
                    string tmpText = Properties.Resources.MSG_CELLEDITINGMODE_TEXT;
                    string tmpTitle = Properties.Resources.MSG_CELLEDITINGMODE_TITLE;
                    if (interactive)
                    {
                        MessageBox.Show(tmpText, tmpTitle);
                    }
                    return new ActionStatus(new[] { new Lib.PDCMessage(tmpTitle, Lib.PDCMessage.TYPE_ERROR), new Lib.PDCMessage(tmpText, Lib.PDCMessage.TYPE_ERROR) });
                }
            }
            try
            {
                PDCLogger.TheLogger.LogStarttime("PDCAction", "Executing Action: " + this.GetType().Name);
                ActionStatus actionStatus = null;
                if (!CanPerformAction(out actionStatus, interactive))
                    return actionStatus;
                else
                    return PerformAction(GetActiveSheetInfo(), interactive);
            }
            catch (Exception e)
            {
                if (interactive)
                {
                    ExceptionHandler.TheExceptionHandler.handleException(e, null);
                }
                return new ActionStatus(e);
            }
            finally
            {
                Lib.Util.PDCLogger.TheLogger.LogStoptime("PDCAction", "Executing Action: " + this.GetType().Name);
                Globals.PDCExcelAddIn.EnablePdcActions();
            }
        }
        #endregion

        #region Visible
        public bool Visible
        {
            get
            {
                return myVisible;
            }
            set
            {
                myVisible = value;
                foreach (Office.CommandBarButton tmpButton in myMenuMap.Values)
                {
                    tmpButton.Visible = value;
                }
            }
        }
        #endregion

        #endregion

        #region events

        #region HandleActionEvent
        private void HandleActionEvent(Office.CommandBarButton aButton, ref bool cancelDefault)
        {
            PerformAction(true);
        }
        #endregion

        #endregion
    }
}
