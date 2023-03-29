using System;
using System.Collections;

using Microsoft.Win32;
using System.Windows.Forms;
using Office=Microsoft.Office.Core;
using System.Diagnostics;
using System.IO;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    class VersionInfoAction:PDCAction 
    {
        public VersionInfoAction(bool beginGroup)
            : base(Properties.Resources.Action_VersionInfo_Caption, ACTION_TAG, beginGroup)
        {
        }
        public const string ACTION_TAG = "PDC_VersionInfoAction";
        public static string myInstallPath = null;
        protected override bool SafeForCellEditingMode()
        {
            return true;
        }

        internal override ActionStatus PerformAction(SheetInfo sheetInfo, bool interactive)
        {
            string tempVersionInfo = GetVersionText();
            string tmpServerInfo = GetServerInfo();
            string tmpInstallDir = GetInstallDirectoryInfo();
            MessageBox.Show(new ExcelHwndWrapper(), tempVersionInfo +"\n" +tmpServerInfo + "\n" + tmpInstallDir, Properties.Resources.MSG_VERSIONINFO_TITLE);
            return new ActionStatus();
        }

        private string GetInstallDirectoryInfo()
        {
            string path;


            path = VersionInfoAction.GetInstallPath();

            if (string.IsNullOrEmpty(path)) return string.Empty;

            return "Installpath: " + path;
        }

        public static string GetInstallPath()
        {
          if (myInstallPath != null)
            return myInstallPath;
          
          try
          {
            IEnumerator tmpComAddIns = Globals.PDCExcelAddIn.Application.COMAddIns.GetEnumerator();
            while (tmpComAddIns.MoveNext())
            {
              Office.COMAddIn tmpAddIn = (Office.COMAddIn)tmpComAddIns.Current;
              if (tmpAddIn.ProgId.Equals("PDCExcelAddIn"))
              {
                RegistryKey myKey = Registry.ClassesRoot.OpenSubKey("CLSID\\{C3943293-C5E0-4271-B1BF-CDD46A39BE06}\\InprocServer32");
                if (myKey == null)
                {
                  //string tmpCodeBase = System.IO.Path.GetDirectoryName((new ExcelUtils()).GetType().Assembly.CodeBase);
                  Uri tmpUri = new Uri((new ExcelUtils()).GetType().Assembly.CodeBase);
                  myInstallPath = System.IO.Path.GetDirectoryName(tmpUri.LocalPath);

                }
                else
                {
                  myInstallPath = myKey.GetValue("ManifestLocation").ToString();
                }
                return myInstallPath;
              }
            }
          }
          catch
          {
            // ignore
          }

          return string.Empty;
        }

        public static string GetServerInfo()
        {
            string tmpServerURL = Lib.Properties.Settings.Default.PDC_Server;
//            string tmpCISURL = Lib.Properties.Settings.Default.PDCLib_CompoundInformationService_CompoundInformationService;
            tmpServerURL = Lib.Util.UserConfiguration.TheConfiguration.GetProperty(Lib.Util.UserConfiguration.PROP_PORTAL_PDC_URL, tmpServerURL);

//            tmpCISURL = Lib.Util.UserConfiguration.TheConfiguration.GetProperty(Lib.Util.UserConfiguration.PROP_WS_COMPOUNDINFORMATIONSERVICE_URL, tmpCISURL);
            return "PDC Server URL: " + tmpServerURL;
        }

        public static string GetVersionNo()
        {
            string file = GetInstallPath() + "\\PDCVersion.dll";
          try
          {
            FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(file);
            string tmpVersion = fvi.FileVersion;
            return tmpVersion;
          }
          catch (FileNotFoundException e)
          {
            PDCLogger.TheLogger.LogException("FileNotFoundException", "Error getting Version", e);
            throw new Exception("PDCVersion.dll not found to determine Client Version",e);
//            return "Error:" + file + " not found";
          }
#pragma warning disable 0168
          catch (Exception e)
          {
            PDCLogger.TheLogger.LogException("Exception", "Error getting Version", e);
            throw new Exception("Unknown Error occured while determining Client Version with PDCVersion.dll", e);
          }
#pragma warning restore 0168

        }

        public static string GetVersionText()
        {
          return "PDC Client Version: " + GetVersionNo() + " (Win7-Win10)";
        }

    }
}
