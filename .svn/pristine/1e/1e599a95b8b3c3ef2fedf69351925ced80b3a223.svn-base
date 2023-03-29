using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;

using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Lib = BBS.ST.BHC.BSP.PDC.Lib;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using BBS.ST.IVY.Chemistry.Util;


namespace PDCOpenLibrary
{
    #region classes

    #region TestStruct
    /// <summary>
    /// structure that might be used for information transfer
    /// </summary>
    [Guid("503DC9E1-780F-4cde-BECD-D4C9A1E642D2")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class TestStruct
    {
#pragma warning disable 1591

        public string compoundno = string.Empty;
        public string preparationno = string.Empty;
        public string mcno = string.Empty;
        public string compoundno_msg = string.Empty;
        public string preparationno_msg = string.Empty;
        public string mcno_msg = string.Empty;
        public int compoundno_id = 0;
        public int preparationno_id = 0;
        public int mcno_id = 0;
        public string[] ErrInfo;
        public int result;
#pragma warning restore 1591
    }
    #endregion

    #region PDCOpenLib
    /// <summary>
    /// the OpenPDC public Class
    /// that servres all the public functions for use in Excel as UDF or in VBA
    /// </summary>
    [Guid("D2A9F659-F7DA-4e06-9582-A21F2B9727B3")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class PDCOpenLib
    {
        // Define some constants as return values for VBA
        // Use readonly instead of const so that they are visible in VBA
        private const int COMPOUNDNO_AND_PREPNO_VALID = 1;
        private const int COMPOUNDNO_PREPNO_PAIR_INVALID = 0;
        private const int PREPNO_INVALID = -1;
        private const int COMPOUNDNO_INVALID = 2;
        private const int COMPOUNDNO_AND_PREPNO_INVALID = -3;
        private const int UNKNOWN_ERROR = -999;
        private const int NOT_LOGGED_IN = -4;
        private const int NOT_A_PDC_SHEET = -5;
        private const int NO_RANGE = -6;
        private const int PDC_MENU_NOT_PRESENT_OR_VISIBLE = -7;
        private const int PDC_ACTION_BUTTON_NOT_FOUND = -8;
        private const int OK = 0;
        private const int BUTTON_SHORTCUT_NOT_FOUND = -9;
        private const int BUTTON_NOT_ENABLED = -10;
        private const int SYMYX_NOT_INSTALLED = -2;
        private const int MOLFILE_NOT_FOUND = -11;
        private const int CUSTOM_ATOM_COLOR_FILE_NOT_EXISTS = -12;
        private const int INVALID_FONT = -13;


        private const string FONT_NAME_ARIAL = "Arial";
        private const string FONT_NAME_COURIER = "Courier";
        private const string MSG_NOT_LOGGED_IN = "not logged in";
        private const string MSG_UNKNOWN_COMPOUND = "unknown compound";

        //private static Thread myThread;
        #region constructor
        /// <summary>
        /// constructor of PDCOpenLib class
        /// </summary>
        public PDCOpenLib()
        {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_COM, "PDCOpenLib init");
        }
        #endregion

        #region methods
        #region IsEmpty
        /// <summary>
        /// Convenience method to check for a null or empty string
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        private bool IsEmpty(string arg)
        {
            return arg == null || string.Empty.Equals(arg.Trim());
        }
        #endregion
        #region EnsureCapacity
        /// <summary>
        /// Sets the size of corresponding measurement to the specified size
        /// </summary>
        /// <param name="anyRange">Provides a link to the Excel object model</param>
        /// <param name="aListRangeName">The list range name pointing to a measurement table</param>
        /// <param name="aSize">The desired size of the table</param>
        public void EnsureCapacity(Excel.Range anyRange, string aListRangeName, int aSize)
        {
            if (!IsLoggedIn) return;
            PDCLogger.TheLogger.LogStarttime("OpenLib.EnsureCapacity", "Extending Range'" + aListRangeName + "' to " + aSize);
            bool tmpEvents = anyRange.Application.EnableEvents;
            try
            {
                anyRange.Application.EnableEvents = false;
                string tmpDataRangeName = aListRangeName.Replace("List", "Data");
                Excel.Worksheet tmpSheet = (Excel.Worksheet)anyRange.Parent;
                Excel.Workbook tmpWB = (Excel.Workbook)tmpSheet.Parent;

                Excel.Range tmpDataRange = tmpSheet.get_Range(tmpDataRangeName, Type.Missing);

                int tmpY = tmpDataRange.Row;
                int tmpX = tmpDataRange.Column - 1;
                int tmpYCount = tmpDataRange.Rows.Count;
                int tmpXCount = aSize + 3;
                Excel.Range tmpNewListRange = range(tmpSheet, tmpSheet.Cells[tmpY, tmpX], tmpSheet.Cells[tmpY + tmpYCount - 1, tmpX + tmpXCount]);

                tmpWB.Names.Add(aListRangeName, tmpNewListRange, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                Excel.Range tmpNewDataRange = tmpSheet.get_Range(tmpSheet.Cells[tmpY, tmpX + 1], tmpSheet.Cells[tmpY + tmpYCount - 1, tmpX + tmpXCount - 1]);
                tmpSheet.Names.Add(tmpDataRangeName, tmpNewDataRange, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Range tmpEventRange = (Excel.Range)tmpSheet.Cells[tmpY, tmpX + 1];
                anyRange.Application.EnableEvents = true;
                tmpEventRange.set_Value(Type.Missing, tmpEventRange.get_Value(Type.Missing));
            }
            finally
            {
                anyRange.Application.EnableEvents = tmpEvents;
                PDCLogger.TheLogger.LogStoptime("OpenLib.EnsureCapacity", "Extended Range'" + aListRangeName + "'");
            }
        }
        #endregion

        private Excel.Range range(Excel.Worksheet aSheet, object aStart, object anEnd)
        {
            return aSheet.get_Range(aStart, anEnd);
        }
        #region GetButton
        /// <summary>
        /// Returns the button with the specified tag in the specified menu.
        /// </summary>
        /// <param name="aTag"></param>
        /// <param name="tmpPDCMenu"></param>
        /// <returns></returns>
        private Office.CommandBarButton GetButton(string aTag, Office.CommandBarPopup tmpPDCMenu)
        {
            System.Collections.IEnumerator tmpButtons = tmpPDCMenu.Controls.GetEnumerator();
            while (tmpButtons.MoveNext())
            {
                object tmpEntry = tmpButtons.Current;
                if (tmpEntry is Office.CommandBarButton)
                {
                    Office.CommandBarButton tmpButton = (Office.CommandBarButton)tmpEntry;
                    if (tmpButton.Tag == aTag)
                    {
                        return tmpButton;
                    }
                }
            }
            return null;
        }
        #endregion

        #region GetCompoundInformation
        /// <summary>
        /// retrieves from CoumpoundInformationService a CompoundInfoStructure
        /// including Picture and Molfile
        /// </summary>
        /// <param name="CI">
        ///   
        /// </param>
        /// <returns>
        ///    1 = compound ID and prep pair valid
        ///    0 = compound ID and prep are both valid, pair is invalid (don't go together)
        ///   -1 = compound ID is valid, prep is invalid   
        ///   -2 = compound ID is invalid, prep is valid 
        ///   -3 = compound ID is invalid, prep is invalid
        ///   -4 = user not logged in 
        /// -999 = undetermined error
        ///</returns>
        [Description("retrieves from CoumpoundInformationService a CompoundInfoStructure including Picture and Molfile")]
        public int GetCompoundInformation(ref TestStruct CI)
        {
            if (!IsLoggedIn) return NOT_LOGGED_IN;

            if (CI == null) return OK;

            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_COM, "GetCompoundInformation start");
            CI.result = -999;
            BBS.ST.BHC.BSP.PDC.Lib.TestStruct tmpCI = new BBS.ST.BHC.BSP.PDC.Lib.TestStruct();
            tmpCI.compoundno = CI.compoundno;
            tmpCI.preparationno = CI.preparationno;
            tmpCI.mcno = CI.mcno;
            try
            {
                BBS.ST.BHC.BSP.PDC.Lib.TestStruct retCI = Lib.PDCService.ThePDCService.GetCompoundInformation(tmpCI);
                CI.compoundno = retCI.compoundno;
                CI.compoundno_id = retCI.compoundno_id;
                CI.compoundno_msg = retCI.compoundno_msg;
                CI.preparationno = retCI.preparationno;
                CI.preparationno_id = retCI.preparationno_id;
                CI.preparationno_msg = retCI.preparationno_msg;
                CI.mcno = retCI.mcno;
                CI.mcno_id = retCI.mcno_id;
                CI.mcno_msg = retCI.mcno_msg;
                if (CI.compoundno_id < 2 && CI.preparationno_id < 2)
                {
                    // compound ID and prep pair valid
                    CI.result = COMPOUNDNO_AND_PREPNO_VALID;
                }
                else if (CI.compoundno_id < 2 && CI.preparationno_id > 1)
                {
                    // compound ID is valid, prep is invalid
                    CI.result = PREPNO_INVALID;
                    List<string> MyList = new List<string>();

                    MyList.Add(CI.compoundno_msg);
                    MyList.Add(CI.preparationno_msg);
                    MyList.Add(CI.mcno_msg);
                    CI.ErrInfo = MyList.ToArray();
                }
                else if (CI.preparationno_id < 2 && CI.compoundno_id > 1)
                {
                    // compound ID is invalid, prep is valid
                    CI.result = COMPOUNDNO_INVALID;
                    List<string> MyList = new List<string>();

                    MyList.Add(CI.compoundno_msg);
                    MyList.Add(CI.preparationno_msg);
                    MyList.Add(CI.mcno_msg);
                    CI.ErrInfo = MyList.ToArray();
                }
                else
                {
                    // compound ID is invalid, prep is invalid
                    CI.result = COMPOUNDNO_AND_PREPNO_INVALID;
                    // compound ID and prep are both valid, pair is invalid (don't go together)
                    CI.result = COMPOUNDNO_PREPNO_PAIR_INVALID;
                    List<string> MyList = new List<string>();

                    MyList.Add(CI.compoundno_msg);
                    MyList.Add(CI.preparationno_msg);
                    MyList.Add(CI.mcno_msg);
                    CI.ErrInfo = MyList.ToArray();
                }
                return OK;
            }
            finally
            {
                PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_COM, "GetCompoundInformation Dispose");
                tmpCI = null;
            }
        }
        #endregion

        #region GetMolFormula
        /// <summary>
        /// retrieves for a CompondNo the Mol Formula
        /// </summary>
        /// <param name="compoundNo">the compound number, e.g. BAY 101079</param>
        /// <returns>the Mol Formula, or errorinformation</returns>
        [Description("retrieves for a CompondNo the Mol Formula")]
        public string GetMolFormula(string compoundNo)
        {
            PDCLogger.TheLogger.LogStarttime("Openlib.GetMolFormula", "GetMolFormula - Method start");
            if (!IsLoggedIn) return MSG_NOT_LOGGED_IN;
            if (compoundNo == null) return "";

            string strRetval;

            Lib.TestStruct tmpInfo = new Lib.TestStruct();
            try
            {
                tmpInfo.compoundno = compoundNo;
                tmpInfo = Lib.PDCService.ThePDCService.GetCompoundInformation(tmpInfo);

                if (tmpInfo.msg != null && !tmpInfo.msg.Equals(string.Empty))
                {
                    strRetval = tmpInfo.msg;
                }
                else
                {
                    if (!IsEmpty(tmpInfo.compoundno_msg) && IsEmpty(tmpInfo.molfile))
                    {
                        return MSG_UNKNOWN_COMPOUND;
                    }
                    if (tmpInfo.molformula == null)
                    {
                        return "";
                    }
                    strRetval = tmpInfo.molformula;
                }
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("Openlib.GetMolFormula", "GetMolFormula - Method end");
            }
            return strRetval;
        }
        #endregion

        #region GetMolWeight
        /// <summary>
        /// retrieves for a CompondNo the Mol weight
        /// </summary>
        /// <param name="CompoundNo">the compounf number</param>
        /// <returns>the mol weight</returns>
        [Description("retrieves for a CompondNo the Mol Weight")]
        public string GetMolWeight(string CompoundNo)
        {
            if (!IsLoggedIn) return MSG_NOT_LOGGED_IN;
            if (CompoundNo == null) return "";
            PDCLogger.TheLogger.LogStarttime("Openlib.GetMolWeight", "GetMolWeight - Method start");
            string strRetval = string.Empty;
            BBS.ST.BHC.BSP.PDC.Lib.TestStruct tmpInfo = new BBS.ST.BHC.BSP.PDC.Lib.TestStruct();
            try
            {
                tmpInfo.compoundno = CompoundNo;

                tmpInfo = Lib.PDCService.ThePDCService.GetCompoundInformation(tmpInfo);

                if (tmpInfo.msg != null && !tmpInfo.msg.Equals(string.Empty))
                {
                    strRetval = tmpInfo.msg;
                }
                else
                {
                    CompoundNo = tmpInfo.compoundno;
                    if (!IsEmpty(tmpInfo.compoundno_msg))
                    {
                        return MSG_UNKNOWN_COMPOUND;
                    }

                    return System.Convert.ToString(tmpInfo.molweight, CultureInfo.CurrentUICulture);
                }
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(nameof(OpenLib) + "." + nameof(GetMolWeight), "", e);
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("Openlib.GetMolWeight", "GetMolWeight - Method end");
            }
            return strRetval;
        }
        #endregion
        #region GetStructureDrawing
        /// <summary>
        ///  retrieves the Molfile for a given compound number,
        ///  and place it to the given range
        /// </summary>
        /// <param name="compoundNo">the compound number</param>
        /// <param name="targetRange">the target range</param>
        /// <param name="cellWidth">fix cell width for the cell. structure will be resized</param>
        /// <param name="cellHeight">fix cell height for the cell. structure will be resized. 
        ///             IF width AND height is zero, the cell is resized the size of the structure</param>
        /// <param name="hydroges">as literal one off (ALL, NONE, HETERO, TERMINAL, HETEROORTERMINAL)</param>
        /// <returns></returns>
        [Description("retrieves the Molfile for a given compound number,and place it to the given range")]
        public string GetStructureDrawing(string compoundNo, Excel.Range targetRange, int cellWidth, int cellHeight, [Optional] object hydroges)
        {

            if (!IsLoggedIn) return MSG_NOT_LOGGED_IN;



            if (compoundNo == null || targetRange == null || compoundNo == string.Empty) return "";
            PDCLogger.TheLogger.LogStarttime("Openlib.GetStructureDrawing", "GetStructureDrawing - Method start");

            try
            {

                UserSettings userSettings = new UserSettings();
                userSettings.HydrogenDisplayMode = GetHydrogenDisplayMode(hydroges);
                userSettings.ResizeMode = UserSettings.ResizeModes.Maximum;
                userSettings.ColumnWidth = cellWidth;
                userSettings.RowHeight = cellHeight;

                Lib.TestStruct tmpInputTestStruct = new Lib.TestStruct();
                tmpInputTestStruct.compoundno = compoundNo;
                tmpInputTestStruct.hydrogendisplaymode = userSettings.HydrogenDisplayMode.ToString();
                tmpInputTestStruct.fileformat = "BMP";
                tmpInputTestStruct.username = LoggedInUser;

                Lib.TestStruct compoundInfo = Lib.PDCService.ThePDCService.GetCompoundInformation(tmpInputTestStruct);

                System.Diagnostics.Debug.WriteLine(compoundInfo.ErrInfo == null || compoundInfo.ErrInfo.Length == 0
                    ? "null"
                    : compoundInfo.ErrInfo[0]);

                if (compoundInfo.ErrInfo != null && compoundInfo.ErrInfo.Length > 0)
                {
                    return compoundInfo.ErrInfo[0];
                }
                //  1 = compound ID and prep pair valid
                if (compoundInfo.molimagearray != null)
                {
                    Excel.Worksheet sheet = (Excel.Worksheet) targetRange.Parent;
                    ExcelShapeUtils.TheUtils.InsertStructureDrawing(sheet, compoundInfo, targetRange.Column,
                        targetRange.Row, userSettings, null, false, true);


                }
                return "";
            }
            catch (Exception e)
            {
                return e.Message;
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("Openlib.GetStructureDrawing", "GetStructureDrawing - Method end");
            }
        }
        #endregion

        private HydrogenDisplayMode GetHydrogenDisplayMode(object hydroges)
        {

            string strHydroges = string.Empty;
            if (!(hydroges is System.Reflection.Missing))
            {
                string givenhydrogen;
                if (hydroges is Excel.Range)
                {
                    // Optinal parameters are passed as Ranges !!!!
                    Excel.Range rr = hydroges as Excel.Range;
                    givenhydrogen = (string)rr.Text;
                }
                else
                {
                    givenhydrogen = (string)hydroges;
                }

                switch (givenhydrogen.ToUpper())
                {
                    case "ALL":
                        return HydrogenDisplayMode.All;
                    case "NONE":
                        return HydrogenDisplayMode.None;
                    case "HETERO":
                        return HydrogenDisplayMode.Hetero;
                    case "TERMINAL":
                        return HydrogenDisplayMode.Terminal;
                    case "HETEROORTERMINAL":
                        return HydrogenDisplayMode.HeteroOrTerminal;
                    default:
                        return HydrogenDisplayMode.None;
                }
            }
            // ALL, NONE, HETERO, TERMINAL, HETEROORTERMINAL
            //Hydroges = "NONE";
            return HydrogenDisplayMode.None;
        }
        #region GetSubKeyName
        private static string GetSubKeyName(Type type, string subKeyName)
        {
            System.Text.StringBuilder s = new System.Text.StringBuilder();
            s.Append(@"CLSID\{");
            s.Append(type.GUID.ToString().ToUpper());
            s.Append(@"}\");
            s.Append(subKeyName);
            return s.ToString();
        }
        #endregion

        #region GetSystemFont
        /// <summary>
        ///   Returns a system font object with the given name and size.
        /// </summary>
        /// <param name="fontName">
        ///   The name of the font.
        /// </param>
        /// <param name="fontSize">
        ///   The size of the font.
        /// </param>
        /// <returns>
        ///   A system font object with the given name and size.
        /// </returns>
        private Font GetSystemFont(string fontName, float fontSize)
        {
            switch (fontName)
            {
                case FONT_NAME_ARIAL:
                    return new Font("Arial", fontSize);
                case FONT_NAME_COURIER:
                    return new Font("Courier", fontSize);
                default:
                    return new Font("Arial", 11);
            }
        }
        #endregion

        #region GetVersion
        /// <summary>
        /// Returns the version of the open library
        /// </summary>
        /// <returns></returns>
        [Description("Returns the version of the open library")]
        public string GetVersion()
        {
            return OpenLibRevision();
        }
        #endregion

        #region IsLoggedIn
        /// <summary>
        /// Property for the login state
        /// </summary>
        public String LoggedInUser => Lib.RegistryUtil.LoggedInUser?.Cwid ?? string.Empty;

        /// <summary>
        /// Property for the login state
        /// </summary>
        public bool IsLoggedIn => !string.IsNullOrEmpty(LoggedInUser);

        #endregion



        #region IsSymyxComponentInstalled
        /// <summary>
        ///   Returns true, when there is ISIS/Draw, MDL Draw or Symyx Draw installed.
        ///   Otherwise false.
        /// </summary>
        /// <returns>
        ///   True, when there is ISIS/Draw, MDL Draw or Symyx Draw installed.
        ///   False otherwise.
        /// </returns>
        private bool IsSymyxComponentInstalled()
        {
            return ExcelShapeUtils.TheUtils.UseIsisOrMdl();
        }
        #endregion

        #region OpenLibRevision
        /// <summary>
        /// Returns the revision information for the open library
        /// </summary>
        /// <returns></returns>
        public static string OpenLibRevision()
        {
            try
            {
                Version tmpVersion = typeof(PDCOpenLib).Assembly.GetName().Version;
                if (tmpVersion == null)
                {
                    return "Unknown";
                }
                return tmpVersion.ToString();
            }
#pragma warning disable 0168
            catch (Exception e)
            {
                return "Unknown";
            }
#pragma warning restore 0168
        }
        #endregion

        #region PaintStructure
        /// <summary>
        ///   Prints the image of the structure for the given compound no to the given target range.
        /// </summary>
        /// <param name="compoundNo">
        ///   The compound no of the structure.
        /// </param>
        /// <param name="targetRange">
        ///   The cell where the image shall be written.
        /// </param>
        /// <returns>
        ///    1 when it was successfull.
        ///    0 when an unknown error happened.
        ///   -1 when the user is not logged in.
        ///   -2 when the Symyx components are not installed.
        ///   -3 when a necessary parameter is null.
        ///   -4 when there exist no structure information for the given compound no.
        /// </returns>
        [Description("Prints the image of the structure for the given compound no to the given target range.\n Returns:\n   1 when it was successfull.\n   0 when an unknown error happened.\n  -1 when the user is not logged in.\n  -2 when the Symyx components are not installed.\n  -3 when a necessary parameter is null.\n  -4 when there exist no structure information for the given compound no.")]
        public int PaintStructure(string compoundNo, Excel.Range targetRange)
        {
            if (compoundNo == null || targetRange == null) return OK;

            if (!this.IsLoggedIn) return NOT_LOGGED_IN;

            if (!this.IsSymyxComponentInstalled()) return SYMYX_NOT_INSTALLED;
            Lib.Util.UserSettings usersettings = new UserSettings();
            string hydrogenDisplayMode;
            try
            {
                usersettings.ReadSettings();
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException("", "PaintStructure: usersetting.Readsettings() failed", e);
                return UNKNOWN_ERROR;
            }


            switch (usersettings.HydrogenDisplayMode)
            {
                case HydrogenDisplayMode.All:
                    hydrogenDisplayMode = "All";
                    break;
                case HydrogenDisplayMode.Hetero:
                    hydrogenDisplayMode = "Hetero";
                    break;
                case HydrogenDisplayMode.HeteroOrTerminal:
                    hydrogenDisplayMode = "HeteroOrTerminal";
                    break;
                case HydrogenDisplayMode.Terminal:
                    hydrogenDisplayMode = "Terminal";
                    break;
                default:
                    hydrogenDisplayMode = "None";
                    break;
            }


            /// here comes the reading of Cell width and cellheight from USERSETTINGS!!!
            int cellWidth = Convert.ToInt32(usersettings.ColumnWidth);
            int cellHeight = Convert.ToInt32(usersettings.RowHeight);

            return this.PaintStructureSpecial(compoundNo, targetRange, cellWidth, cellHeight, usersettings.ChemLabelFont.Name, usersettings.ChemLabelFont.Size,
              usersettings.DisplayCarbonLabels, hydrogenDisplayMode, usersettings.TextFont.Name, usersettings.TextFont.Size, usersettings.TransparentBackground,
              usersettings.BondLength, usersettings.AtomColor);
        }
        #endregion

        #region PaintStructureSpecial
        /// <summary>
        ///   Prints the image of the structure for the given compound no to the given target range.
        ///   The output format will be set by the different settings.
        /// </summary>
        /// <param name="compoundNo">
        ///   The compound no a of the structure.
        /// </param>
        /// <param name="targetRange">
        ///   The cell where the image shall be written.
        /// </param>
        /// <param name="cellWidth">
        ///   fix cell width for the cell. structure will be resized
        /// </param>
        /// <param name="cellHeight">
        ///   fix cell height for the cell. structure will be resized. 
        ///   IF width AND height is zero, the cell is resized the size of the structure
        /// </param>
        /// <param name="chemLabelFont">
        ///   The name of the font for the chemical labels.
        /// </param>
        /// <param name="chemLabelFontSize">
        ///   The size of the font for the chemical labels.
        /// </param>
        /// <param name="displayCarbonLabels">
        ///   A flag, whether the carbon labels shall be printed or not.
        /// </param>
        /// <param name="hydrogenDisplayMode">
        ///   A string containing the hydrogen display mode. Values are:
        ///   All, Hetero, HeteroOrTerminal, None or Terminal
        /// </param>
        /// <param name="textFont">
        ///   The name of the font for the text labels.
        /// </param>
        /// <param name="textFontSize">
        ///   The size of the font for the text labels.
        /// </param>
        /// <param name="transparentBackground">
        ///   A flag, whether the background shall be transparent or not.
        /// </param>
        /// <param name="bondLength">
        /// The length of the bonds.
        /// </param>
        /// <param name="atomColor">
        ///   A flag, whether the atom shouild be colored or not
        /// </param>
        /// <returns>
        ///    1 when it was successfull.
        ///    0 when an unknown error happened.
        ///   -1 when the user is not logged in.
        ///   -2 when the Symyx components are not installed.
        ///   -3 when a necessary parameter is null.
        ///   -4 when there exist no structure information for the given compound no.
        /// </returns>
        [Description("Prints the image of the structure for the given compound no to the given target range. The output format will be set by the different settings.\n Parameters:\n   compoundNo => The compound no of the structure.\n   targetRange => The cell where the image shall be written.\n   chemLabelFont => The name of the font for the chemical labels.\n   chemLabelFontSize => The size of the font for the chemical labels.\n   displayCarbonLabels => A flag, whether the carbon labels shall be printed or not.\n   hydrogenDisplayMode => A string containing the hydrogen display mode. Values are: All, Hetero, HeteroOrTerminal, None or Terminal\n   textFont => The name of the font for the text labels.\n   textFontSize => The size of the font for the text labels.\n   transparentBackground => A flag, whether the background shall be transparent or not.\n Returns:\n   1 when it was successfull.\n   0 when an unknown error happened.\n  -1 when the user is not logged in.\n  -2 when the Symyx components are not installed.\n  -3 when a necessary parameter is null.\n  -4 when there exist no structure information for the given compound no.")]
        public int PaintStructureSpecial(string compoundNo, Excel.Range targetRange, int cellWidth, int cellHeight, string chemLabelFont, float chemLabelFontSize, bool displayCarbonLabels,
          string hydrogenDisplayMode, string textFont, float textFontSize, bool transparentBackground, float bondLength, bool atomColor)
        {

            if (compoundNo == null || targetRange == null || chemLabelFont == null || hydrogenDisplayMode == null || textFont == null) return -3;
            if (!this.IsSymyxComponentInstalled())
            {
                PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_COM, $"{nameof(PaintStructureSpecial)}: This version of PaintStructure requires an installed version of Symyx or similar.");
                return SYMYX_NOT_INSTALLED;
            }
            if (!this.IsLoggedIn) return NOT_LOGGED_IN;
            PDCLogger.TheLogger.LogStarttime("OpenLib.PaintStructureSpecial", "PaintStructureSpecial - Method start");
            try
            {
                PDCLogger.TheLogger.LogStarttime("OpenLib.PaintStructureSpecial.Settings", "PaintStructureSpecial.Settings - Initializing settings");

                UserSettings settings = new UserSettings();
                settings.HydrogenDisplayMode = GetHydrogenDisplayMode(hydrogenDisplayMode);
                settings.DisplayCarbonLabels = displayCarbonLabels;
                //settings.TextFont = this.GetSystemFont(textFont, textFontSize);
                settings.TextFont = new Font(textFont, textFontSize);

                settings.TransparentBackground = transparentBackground;
                try
                {
                    settings.AtomColor = atomColor;
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException("", "PaintStructureSpecial: set AtomColor failed", e);
                    return UNKNOWN_ERROR;
                }

                //settings.ChemLabelFont = this.GetSystemFont(chemLabelFont, chemLabelFontSize);
                settings.ChemLabelFont = new Font(chemLabelFont, chemLabelFontSize);

                if (!settings.TextFont.OriginalFontName.Equals(settings.TextFont.Name))
                {
                    PDCLogger.TheLogger.LogWarning("Warning",
                        "Text Font has changed from " + settings.TextFont.OriginalFontName + " to " +
                        settings.TextFont.Name);
                }
                if (!settings.ChemLabelFont.OriginalFontName.Equals(settings.ChemLabelFont.Name))
                {
                    PDCLogger.TheLogger.LogWarning("Warning",
                        "ChemLabel Font has changed from " + settings.ChemLabelFont.OriginalFontName + " to " +
                        settings.ChemLabelFont.Name);
                }

                settings.ResizeMode = UserSettings.ResizeModes.Maximum;
                settings.RowHeight = cellHeight;
                settings.ColumnWidth = cellWidth;
                settings.BondLength = bondLength;

                PDCLogger.TheLogger.LogStoptime("OpenLib.PaintStructureSpecial.Settings", "PaintStructureSpecial.Settings - Initialized settings");

                Excel.Worksheet sheet = (Excel.Worksheet) targetRange.Parent;
                try
                {
                    Lib.TestStruct tmpInputTestStruct = new Lib.TestStruct();
                    tmpInputTestStruct.compoundno = compoundNo;
                    tmpInputTestStruct.hydrogendisplaymode = settings.HydrogenDisplayMode.ToString();
                    tmpInputTestStruct.fileformat = "MOL";
                    tmpInputTestStruct.username = LoggedInUser;

                    Lib.TestStruct compoundInfo =
                        Lib.PDCService.ThePDCService.GetCompoundInformation(tmpInputTestStruct);

                    if (compoundInfo.molfile == null || compoundInfo.molfile.Trim().Equals(""))
                        return MOLFILE_NOT_FOUND;

                    ExcelShapeUtils.TheUtils.InsertISISObject(sheet, targetRange.Column, targetRange.Row,
                        compoundInfo.molfile, settings, true);

                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_COM, "PaintStructureSpecial failed", e);
                    return UNKNOWN_ERROR;
                }
                return OK;
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("OpenLib.PaintStructureSpecial", "PaintStructureSpecial - Method end");

            }
        }
        #endregion

        #region RegisterFunction
        /// <summary>
        /// registers the COM functions of this class
        /// </summary>
        /// <param name="type"></param>
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type, "Programmable"));
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(GetSubKeyName(type, "InprocServer32"), true);
            key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
        }
        #endregion

        #region UnregisterFunction
        /// <summary>
        /// deregisters the COM functions of this class
        /// </summary>
        /// <param name="type"></param>
        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type, "Programmable"), false);
        }
        #endregion

        #endregion

        #region properties

        #region FontNameArial
        /// <summary>
        ///   Gets the name of the arial font.
        /// </summary>
        public string FontNameArial
        {
            get
            {
                return FONT_NAME_ARIAL;
            }
        }
        #endregion

        #region FontNameCourier
        /// <summary>
        ///   Gets the name of the courier font.
        /// </summary>
        public string FontNameCourier
        {
            get
            {
                return FONT_NAME_COURIER;
            }
        }
        #endregion

        #endregion
    }

    #endregion

    #endregion
}
