using System;
using System.Drawing;
using System.Globalization;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using BBS.ST.IVY.Chemistry.Util;
using Microsoft.Win32;

namespace BBS.ST.BHC.BSP.PDC.Lib.Util
{
    /// <summary>
    ///    This class implements the user settings.
    /// </summary>
    [Serializable]
    public class UserSettings
    {
        public const float MAX_ROW_HEIGHT = 400;
        public const float MIN_ROW_HEIGHT = 15;
        public const float DEFAULT_ROW_HEIGHT = 80;

        public const float MAX_COLUMN_WIDTH = 250;
        public const float MIN_COLUMN_WIDTH = 5;
        public const float DEFAULT_COLUMN_WIDTH = 20;

        public const float MIN_BOND_LENGTH = 0.02f;
        public const float DEFAULT_BOND_LENGTH = 0.4f;
        public const float MAX_BOND_LENGTH = 25;

        #region Attributes

        private float myBondLength;
        private float myColumnWidth;
        private float myRowHeight;
        private Font myChemLabelFont;
        private bool myDisplayAtomNumbers;
        private bool myDisplayCarbonLabels;
        private HydrogenDisplayMode myHydrogenDisplayMode;
        private Font myTextFont;
        private bool myTransparentBackground;
        private bool myAtomColor;
        private ResizeModes myResizeMode;
        private bool myDisplayPrepno;
        private bool myDisplayStructure;
        private bool myDisplayMolweight;
        private bool myDisplayMolformula;
        private int myHorizontalOffset;
        private int myVerticalOffset;
        private Direction myOrientation;

        #endregion
        public enum ResizeModes
        {
            StructureDefault,
            FixedWidth,
            FixedHeight,
            /// <summary>
            /// Both width and height are specified and maximum is taken.
            /// </summary>
            Maximum
        }

        /// <summary>
        /// The compound infos are spread in vertical or horizontal direction
        /// </summary>
        public enum Direction
        {
            Horizontal,
            Vertical
        }

        #region constructor
        /// <summary>
        ///    The UserSettings constructor.
        /// </summary>
        public UserSettings()
        {
            myChemLabelFont = new Font("Arial", 10, FontStyle.Regular);
            myDisplayAtomNumbers = false;
            myDisplayCarbonLabels = false;
            myHydrogenDisplayMode = HydrogenDisplayMode.None;
            myTextFont = new Font("Times New Roman", 10, FontStyle.Regular);
            myTransparentBackground = false;
            AtomColor = false;
            myResizeMode = ResizeModes.StructureDefault;
        }
        #endregion

        #region methods

        #region GetFontFromRegistryValue
        /// <summary>
        ///    Returns a font for the given registry value.
        /// </summary>
        /// <param name="registryValue">
        ///    The value from the registry.
        /// </param>
        /// <returns>
        ///    A font containing the given registry value.
        /// </returns>
        private Font GetFontFromRegistryValue(string registryValue)
        {
            if (registryValue == null)
            {
                return FallbackFont;
            }
            string[] strings = registryValue.Split(';');
            if (strings.Length < 3)
            {
                return FallbackFont;
            }
            try
            {
                switch (strings[2].Trim())
                {
                    case "Bold":
                        return new Font(strings[0], Convert.ToSingle(strings[1]), FontStyle.Bold);
                    case "Italic":
                        return new Font(strings[0], Convert.ToSingle(strings[1]), FontStyle.Italic);
                    case "Regular":
                        return new Font(strings[0], Convert.ToSingle(strings[1]), FontStyle.Regular);
                    case "Strikeout":
                        return new Font(strings[0], Convert.ToSingle(strings[1]), FontStyle.Strikeout);
                    case "Underline":
                        return new Font(strings[0], Convert.ToSingle(strings[1]), FontStyle.Underline);
                    default:
                        return FallbackFont;
                }
            }
            catch
            {
                return FallbackFont;
            }

        }

        private static Font FallbackFont
        {
            get { return new Font("Arial", 10, FontStyle.Regular); }
        }

        #endregion

        #region ReadSettings

        private bool errorOccured = false;
        private float GetRegistryValue(string key, float defaultValue)
        {
            CultureInfo info = CultureInfo.InvariantCulture;
            string regValue = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, key, defaultValue)?.ToString();
            if (regValue != null && regValue.Contains(","))
            {
                info = CultureInfo.GetCultureInfo("de-DE");
            }
            if (float.TryParse(regValue, NumberStyles.Float, info, out var retValue))
            {
                return retValue;
            }
            errorOccured = true;
            return defaultValue;
        }
        private int GetRegistryValue(string key, int defaultValue)
        {
            object regValue = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, key, defaultValue);
            int retValue;
            if (int.TryParse(regValue.ToString(), out retValue))
            {
                return retValue;
            }
            errorOccured = true;
            return defaultValue;
        }
        /// <summary>
        ///    This method reads the user settings from the registry.
        ///    If there are no registry settings, the default settings will be written.
        /// </summary>
        public void ReadSettings()
        {
            myDisplayPrepno = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayPrepno", "True").ToString() == "True";
            myDisplayStructure = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayStructure", "True").ToString() == "True";
            myDisplayMolformula = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayMolformula", "True").ToString() == "True";
            myDisplayMolweight = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayMolweight", "True").ToString() == "True";

            myHorizontalOffset = GetRegistryValue("HorizontalOffset", 1);
            myVerticalOffset = GetRegistryValue("VerticalOffset", 0);

            switch ((string)Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "Orientation", Direction.Horizontal.ToString()))
            {
                case "Vertical":
                    myOrientation = Direction.Vertical; break;
                case "Horizontal":
                    myOrientation = Direction.Horizontal; break;
            }

            myBondLength = GetRegistryValue("BondLength", DEFAULT_BOND_LENGTH);
            if (myBondLength > MAX_BOND_LENGTH || myBondLength < MIN_BOND_LENGTH)
            {
                myBondLength = DEFAULT_BOND_LENGTH;
                errorOccured = true;
            }
            myColumnWidth = GetRegistryValue("ColumnWidth", DEFAULT_COLUMN_WIDTH);
            if (myColumnWidth > MAX_COLUMN_WIDTH || myColumnWidth < MIN_COLUMN_WIDTH)
            {
                myColumnWidth = DEFAULT_COLUMN_WIDTH;
                errorOccured = true;
            }
            myRowHeight = GetRegistryValue("RowHeight", DEFAULT_ROW_HEIGHT);
            if (myRowHeight > MAX_ROW_HEIGHT || myRowHeight < MIN_ROW_HEIGHT)
            {
                myRowHeight = DEFAULT_ROW_HEIGHT;
                errorOccured = true;
            }

            if (errorOccured || Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ChemLabelFont", "").Equals(""))
            {
                if (errorOccured)
                {
                    PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_LIB, "Found some invalid registry settings. Replacing them with default values.");
                }
                errorOccured = false;
                WriteSettings();
            }

      		myChemLabelFont = GetFontFromRegistryValue(Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ChemLabelFont", "").ToString());

            myDisplayAtomNumbers = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayAtomNumbers", "False").ToString() == "True";

            myDisplayCarbonLabels = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayCarbonLabels", "False").ToString() == "True";

            switch ((string)Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "HydrogenDisplayMode", "None"))
            {
                case "All":
                    myHydrogenDisplayMode = HydrogenDisplayMode.All;
                    break;
                case "Hetero":
                    myHydrogenDisplayMode = HydrogenDisplayMode.Hetero;
                    break;
                case "HeteroOrTerminal":
                    myHydrogenDisplayMode = HydrogenDisplayMode.HeteroOrTerminal;
                    break;
                case "None":
                    myHydrogenDisplayMode = HydrogenDisplayMode.None;
                    break;
                case "Terminal":
                    myHydrogenDisplayMode = HydrogenDisplayMode.Terminal;
                    break;
            }
            switch ((string)Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ResizeMode", "StructureDefault"))
            {
                case "StructureDefault":
                    myResizeMode = ResizeModes.StructureDefault;
                    break;
                case "FixedHeight":
                    myResizeMode = ResizeModes.FixedHeight;
                    break;
                case "FixedWidth":
                    myResizeMode = ResizeModes.FixedWidth;
                    break;
                default:
                    myResizeMode = ResizeModes.StructureDefault;
                    break;
            }


            myTextFont = GetFontFromRegistryValue(Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "TextFont", "").ToString());
            myTransparentBackground = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "TransparentBackground", "False").ToString() == "True";
            AtomColor = Registry.GetValue(PDCClientConstants.PDC_REGISTRY_KEY, "AtomColor", "False").ToString() == "True";
        }
        #endregion

        #region WriteSettings
        /// <summary>
        ///    This method writes the user settings to the registry.
        /// </summary>
        public void WriteSettings()
        {
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayPrepno", myDisplayPrepno);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayStructure", myDisplayStructure);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayMolweight", myDisplayMolweight);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayMolformula", myDisplayMolformula);

            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "HorizontalOffset", myHorizontalOffset.ToString());
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "VerticalOffset", myVerticalOffset.ToString());
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "Orientation", myOrientation);

            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ChemLabelFont",
              myChemLabelFont.Name + "; " + myChemLabelFont.Size + "; " + myChemLabelFont.Style);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayAtomNumbers", myDisplayAtomNumbers);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "DisplayCarbonLabels", myDisplayCarbonLabels);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "HydrogenDisplayMode", myHydrogenDisplayMode);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ResizeMode", myResizeMode);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "BondLength", myBondLength.ToString(CultureInfo.InvariantCulture));
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "RowHeight", myRowHeight.ToString(CultureInfo.InvariantCulture));
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "ColumnWidth", myColumnWidth.ToString(CultureInfo.InvariantCulture));

            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "TextFont", myTextFont.Name + "; " + myTextFont.Size + "; " + myTextFont.Style);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "TransparentBackground", myTransparentBackground);
            Registry.SetValue(PDCClientConstants.PDC_REGISTRY_KEY, "AtomColor", myAtomColor);
        }
        #endregion

        #endregion

        #region properties

        #region BondLength
        /// <summary>
        ///   The length of the bonds.
        /// </summary>
        public float BondLength
        {
            get
            {
                return myBondLength;
            }
            set
            {
                myBondLength = Math.Max(value, MIN_BOND_LENGTH);
            }
        }
        #endregion

        #region ColumnWidth
        public float ColumnWidth
        {
            get
            {
                return myColumnWidth;
            }
            set
            {
                myColumnWidth = Math.Min(Math.Max(value, MIN_COLUMN_WIDTH), MAX_COLUMN_WIDTH);
            }
        }
        #endregion
        #region RowHeight
        public float RowHeight
        {
            get
            {
                return myRowHeight;
            }
            set
            {
                myRowHeight = Math.Min(Math.Max(value, MIN_ROW_HEIGHT), MAX_ROW_HEIGHT);
            }
        }
        #endregion
        #region ChemLabelFont
        /// <summary>
        ///    The font for the chemical labels.
        /// </summary>
        [XmlIgnore]
        public Font ChemLabelFont
        {
            get
            {
                if (myChemLabelFont == null && ChemLabelFontname != null && ChemLabelFontsize != null)
                {
                    myChemLabelFont = new Font(ChemLabelFontname, ChemLabelFontsize.Value);
                }
                return myChemLabelFont;
            }
            set
            {
                myChemLabelFont = value;
                ChemLabelFontname = value?.Name;
                ChemLabelFontsize = value?.Size;
            }
        }

        public string ChemLabelFontname { get; set; }
        public float? ChemLabelFontsize { get; set; }
        #endregion

        #region DisplayAtomNumbers
        /// <summary>
        ///    Property to display atom numbers or not.
        /// </summary>
        public bool DisplayAtomNumbers
        {
            get
            {
                return myDisplayAtomNumbers;
            }
            set
            {
                myDisplayAtomNumbers = value;
            }
        }
        #endregion

        #region DisplayCarbonLabels
        /// <summary>
        ///    Property to display carbon labels or not.
        /// </summary>
        public bool DisplayCarbonLabels
        {
            get
            {
                return myDisplayCarbonLabels;
            }
            set
            {
                myDisplayCarbonLabels = value;
            }
        }
        #endregion

        #region DisplayPrepno
        public bool DisplayPrepno
        {
            get
            {
                return myDisplayPrepno;
            }
            set
            {
                myDisplayPrepno = value;
            }
        }
        #endregion
        #region DisplayStructure
        public bool DisplayStructure
        {
            get
            {
                return myDisplayStructure;
            }
            set
            {
                myDisplayStructure = value;
            }
        }
        #endregion
        #region DisplayMolweight
        public bool DisplayMolweight
        {
            get
            {
                return myDisplayMolweight;
            }
            set
            {
                myDisplayMolweight = value;
            }
        }
        #endregion
        #region DisplayMolformula
        public bool DisplayMolformula
        {
            get
            {
                return myDisplayMolformula;
            }
            set
            {
                myDisplayMolformula = value;
            }
        }
        #endregion
        #region Orientation
        public Direction Orientation
        {
            get
            {
                return myOrientation;
            }
            set
            {
                myOrientation = value;
            }
        }
        #endregion
        #region Offsets
        public int HorizontalOffset
        {
            get
            {
                return myHorizontalOffset;
            }
            set
            {
                myHorizontalOffset = value;
            }
        }
        public int VerticalOffset
        {
            get
            {
                return myVerticalOffset;
            }
            set
            {
                myVerticalOffset = value;
            }
        }
        #region HydrogenDisplayMode
        /// <summary>
        ///    The hydrogen display mode.
        /// </summary>
        public HydrogenDisplayMode HydrogenDisplayMode
        {
            get
            {
                return myHydrogenDisplayMode;
            }
            set
            {
                myHydrogenDisplayMode = value;
            }
        }
        #endregion

        #region ResizeMode
        /// <summary>
        ///    The Resize Mode .
        /// </summary>
        public ResizeModes ResizeMode
        {
            get
            {
                return myResizeMode;
            }
            set
            {
                myResizeMode = value;
            }
        }
        #endregion
        #region TextFont
        /// <summary>
        ///    The font for the text labels.
        /// </summary>
        [XmlIgnore]
        public Font TextFont
        {
            get
            {
                if (myTextFont == null && TextFontname != null && TextFontsize != null)
                {
                    myTextFont = new Font(TextFontname, TextFontsize.Value);
                }
                return myTextFont;
            }
            set
            {
                myTextFont = value;
                TextFontname = value?.Name;
                TextFontsize = value?.Size;
            }
        }
        public string TextFontname { get; set; }
        public float? TextFontsize { get; set; }
        #endregion

        #region TransparentBackground
        /// <summary>
        ///    Property whether the background should be transparent or not.
        /// </summary>
        public bool TransparentBackground
        {
            get
            {
                return myTransparentBackground;
            }
            set
            {
                myTransparentBackground = value;
            }
        }
        #endregion

        #region AtomColor
        /// <summary>
        ///    Property whether the atoms should be coloring or not.
        /// </summary>
        public bool AtomColor
        {
            get
            {
                return myAtomColor;
            }
            set
            {
                myAtomColor = value;
            }
        }
        #endregion

        #endregion
        #endregion
    }
}
