using System;
using System.Collections.Generic;
using System.Drawing;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using WS = BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// Contains configuration settings as color scheme and binary data constraints
    /// </summary>
    public class ClientConfiguration
    {
        Dictionary<string, Color> myColors;
        /// <summary>
        /// Background color for CompoundInfo headers
        /// </summary>
        public const string HEADER_COMPOUND_INFO = "HeaderCompoundInfo";
        /// <summary>
        /// Background color for PDCInfo headers
        /// </summary>
        public const string HEADER_PDC_INFO = "HeaderPDCInfo";
        /// <summary>
        /// Background color for derived result header
        /// </summary>
        public const string HEADER_DERIVED_RESULT = "HeaderDerivedResult";

        /// <summary>
        /// Background color for core result headers
        /// </summary>
        public const string HEADER_CORE_RESULT = "HeaderCoreResult";
        /// <summary>
        /// Background color for test parameter header
        /// </summary>
        public const string HEADER_PARAMETER = "HeaderParameter";

        public const string HEADER_MANDATORY_PARAMETER = "HeaderMandatoryParameter";
        public const string HEADER_DIFFERENTIATING_PARAMETER = "HeaderDifferentiatingParameter";
        /// <summary>
        /// Background color for binary upload headers
        /// </summary>
        public const string HEADER_BINARY = "HeaderBinary";
        /// <summary>
        /// Background color for comment headers
        /// </summary>
        public const string HEADER_COMMENT = "HeaderComment";
        /// <summary>
        /// Background color for result headers in measurement tables
        /// </summary>
        public const string HEADER_RESULT = "HeaderMeasurementResult";
        /// <summary>
        /// Background color for variable headers in measurement tables
        /// </summary>
        public const string HEADER_VARIABLE = "HeaderMeasurementVariable";
        /// <summary>
        /// Background color for annotation headers
        /// </summary>
        public const string HEADER_ANNOTATION = "HeaderAnnotation";

        public const string MESSAGE_TYPE_ERROR = "MessageTypeError";
        public const string MESSAGE_TYPE_FATAL = "MessageTypeFatal";
        public const string MESSAGE_TYPE_WARN = "MessageTypeWarn";
        public const string MESSAGE_TYPE_INFO_DEBUG = "MessageTypeInfoDebug";

        public const string MESSAGE_MULTIPLE_PREPNOS = "MessageMultiplePrepNos";
        public const string HEADER_PDC_ONLY = "HeaderPdcOnly";
        public const string HEADER_DATA_TYPE = "HeaderDataType";
        public const string HEADER_EXPERIMENT_LEVEL = "HeaderExperimentLevel";
        /// <summary>
        /// A Color consists of a name and an rgb value
        /// </summary>
        private struct ColorParam
        {
            internal string name;
            internal int r;
            internal int g;
            internal int b;
        }

        private Dictionary<string, long> mySupportedTypes = new Dictionary<string, long>();
        private long myMaxByteSize = 512 * 1024;

        #region classes

        #region Color
        /// <summary>
        /// Wrapper around System.Drawing.Color used by .Net and OleColor used by Excel.
        /// </summary>
        public class Color
        {
            System.Drawing.Color color;
            int oleColor;
            /// <summary>
            /// Initializes a PDC Color from the corresponding system color
            /// </summary>
            /// <param name="aColor"></param>
            public Color(System.Drawing.Color aColor)
            {
                color = aColor;
                oleColor = ColorTranslator.ToOle(color);
            }
            public Color(int anOleColor)
            {
                oleColor = anOleColor;
                color = ColorTranslator.FromOle(oleColor);
            }

            /// <summary>
            /// 
            /// </summary>
            public int OleColor
            {
                get
                {
                    return oleColor;
                }
            }

            /// <summary>
            /// 
            /// </summary>
            public System.Drawing.Color SystemColor
            {
                get
                {
                    return color;
                }
            }
        }
        #endregion

        #endregion

        #region constructor
        internal ClientConfiguration(PDCService aService)
        {
            myColors = new Dictionary<string, Color>();
            // Add some default values first
            myColors.Add(HEADER_COMPOUND_INFO, new Color(System.Drawing.Color.Gray));
            myColors.Add(HEADER_PDC_INFO, new Color(System.Drawing.Color.LightYellow));
            myColors.Add(HEADER_DERIVED_RESULT, new Color(System.Drawing.Color.Red));
            Color tmpColor = new Color(System.Drawing.Color.Yellow);
            myColors.Add(HEADER_PARAMETER, tmpColor);
            myColors.Add(HEADER_BINARY, tmpColor);
            myColors.Add(HEADER_COMMENT, tmpColor);
            myColors.Add(HEADER_VARIABLE, tmpColor);
            myColors.Add(HEADER_DIFFERENTIATING_PARAMETER, new Color(System.Drawing.Color.Gold));
            myColors.Add(HEADER_MANDATORY_PARAMETER, new Color(System.Drawing.Color.Cyan));
            myColors.Add(HEADER_CORE_RESULT, new Color(System.Drawing.Color.LawnGreen));
            myColors.Add(HEADER_RESULT, new Color(System.Drawing.Color.Blue));
            myColors.Add(HEADER_ANNOTATION, new Color(System.Drawing.Color.LightBlue));
            myColors.Add(MESSAGE_TYPE_ERROR, new Color(System.Drawing.Color.Red));
            myColors.Add(MESSAGE_TYPE_FATAL, new Color(System.Drawing.Color.Red));
            myColors.Add(MESSAGE_TYPE_WARN, new Color(System.Drawing.Color.Yellow));
            myColors.Add(MESSAGE_TYPE_INFO_DEBUG, new Color(System.Drawing.Color.Green));
            myColors.Add(PDCConstants.C_MSG_UNKNOWN_PREPNO, new Color(System.Drawing.Color.Salmon));
            myColors.Add(PDCConstants.C_MSG_MISSING_PREPNO, new Color(System.Drawing.Color.LightGray));
            myColors.Add(PDCConstants.C_MSG_PDC_PREPARATION_DEFERRED, new Color(System.Drawing.Color.LightGreen));
            myColors.Add(PDCConstants.C_MSG_UNKNOWN_MCNO, new Color(System.Drawing.Color.Salmon));
            myColors.Add(MESSAGE_MULTIPLE_PREPNOS, new Color(System.Drawing.Color.Gold));
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Getting the color configuration from the server");
            try
            {
                InitFromServer(aService);
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Could not connect to get color configuration", e);
            }
            if (mySupportedTypes == null || mySupportedTypes.Count == 0)
            { //fallback in case of server failure
                mySupportedTypes = new Dictionary<string, long>();
                mySupportedTypes.Add("pdf", myMaxByteSize);
                mySupportedTypes.Add("png", myMaxByteSize);
                mySupportedTypes.Add("gif", myMaxByteSize);
                mySupportedTypes.Add("jpg", myMaxByteSize);
            }
        }
        #endregion

        #region methods

        #region InitFromServer
        /// <summary>
        /// Get the configuration data from the PDC WS.
        /// </summary>
        /// <param name="aService"></param>
        private void InitFromServer(PDCService aService)
        {
            WS.Output tmpOutput = aService.CallService(PDCConstants.C_OP_GET_COLOR_CONFIGURATION, null);
            if (tmpOutput == null || tmpOutput.output == null)
            {
                return;
            }
            Dictionary<int, ColorParam> tmpColors = new Dictionary<int, ColorParam>();
            mySupportedTypes = new Dictionary<string, long>();
            foreach (WS.PDCParameter tmpParam in tmpOutput.output)
            {
                if (!tmpParam.position.HasValue)
                {
                    PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_LIB, "Color config without position");
                    continue;
                }
                ColorParam tmpColorParam;
                int tmpPosition = tmpParam.position.Value;
                if (tmpColors.ContainsKey(tmpPosition))
                {
                    tmpColorParam = tmpColors[tmpPosition];
                }
                else
                {
                    tmpColorParam = new ColorParam();
                    tmpColors.Add(tmpPosition, tmpColorParam);
                }
                switch (tmpParam.id)
                {
                    case PDCConstants.C_ID_COLOR_ITEM:
                        tmpColorParam.name = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_COLOR_RGB_R:
                        tmpColorParam.r = (int)(tmpParam.valueNum ?? 0);
                        break;
                    case PDCConstants.C_ID_COLOR_RGB_G:
                        tmpColorParam.g = (int)(tmpParam.valueNum ?? 0);
                        break;
                    case PDCConstants.C_ID_COLOR_RGB_B:
                        tmpColorParam.b = (int)(tmpParam.valueNum ?? 0);
                        break;
                    case PDCConstants.C_ID_BINARY_DATA_MAXSIZE:
                        MaxByteSize = (int)(tmpParam.valueNum ?? MaxByteSize);
                        break;
                    case PDCConstants.C_ID_BINARY_DATA_VALID_FORMAT:
                        if (tmpParam.valueChar != null && !mySupportedTypes.ContainsKey(tmpParam.valueChar.ToLower()) && tmpParam.position != null && tmpParam.position.Value > 0)
                        {
                            mySupportedTypes.Add(tmpParam.valueChar.ToLower(), tmpParam.position.Value);
                        }
                        break;
                }
                tmpColors[tmpPosition] = tmpColorParam;
            }
            foreach (ColorParam tmpColParam in tmpColors.Values)
            {
                if (tmpColParam.name != null)
                {
                    if (myColors.ContainsKey(tmpColParam.name))
                    {
                        myColors[tmpColParam.name] = new Color(System.Drawing.Color.FromArgb(tmpColParam.r, tmpColParam.g, tmpColParam.b));
                    }
                    else
                    {
                        myColors.Add(tmpColParam.name, new Color(System.Drawing.Color.FromArgb(tmpColParam.r, tmpColParam.g, tmpColParam.b)));
                    }
                }
            }
        }
        #endregion

        #endregion

        #region properties

        #region MaxByteSize
        /// <summary>
        /// Returns the maximum size of binary upload data in bytes
        /// </summary>
        public long MaxByteSize
        {
            get
            {
                return myMaxByteSize;
            }
            set
            {
                myMaxByteSize = value;
            }
        }
        #endregion

        #region SupportedTypes
        /// <summary>
        /// Returns the supported file types for binary uploads
        /// </summary>
        public Dictionary<string, long> SupportedTypes
        {
            get
            {
                return mySupportedTypes;
            }
            set
            {
                mySupportedTypes = value;
            }
        }
        #endregion

        #region SupportedTypesAsString
        /// <summary>
        /// Returns a string representing the supported file types for binary uploads.
        /// (Display only)
        /// </summary>
        public string SupportedTypesAsString
        {
            get
            {
                if (mySupportedTypes == null)
                {
                    return "";
                }
                string tmpReturn = "";
                string tmpDelim = "";
                foreach (string tmpType in mySupportedTypes.Keys)
                {
                    tmpReturn += tmpDelim + tmpType;
                    tmpDelim = ",";
                }
                return tmpReturn;
            }
        }
        #endregion

        #region this
        /// <summary>
        /// Returns the color for the specified key or null if the key is unknown.
        /// </summary>
        /// <param name="aKey"></param>
        /// <returns></returns>
        public Color this[string aKey]
        {
            get
            {
                return myColors.ContainsKey(aKey) ? myColors[aKey] : null;
            }
        }
        #endregion

        #region PDCMessage
        public Color GetMessageColor(PDCMessage message)
        {
            if (message.MessageCode != null)
            {
                Color tmpColor = null;
                if (myColors.TryGetValue(message.MessageCode, out tmpColor))
                {
                    return tmpColor;
                }
            }
            switch (message.MessageType)
            {
                case PDCMessage.TYPE_ERROR:
                    return myColors[MESSAGE_TYPE_ERROR];
                case PDCMessage.TYPE_FATAL:
                    return myColors[MESSAGE_TYPE_FATAL];
                case PDCMessage.TYPE_WARN:
                    return myColors[MESSAGE_TYPE_WARN];
                case PDCMessage.TYPE_INFO_DEBUG:
                    return myColors[MESSAGE_TYPE_INFO_DEBUG];
                default:
                    return null;
            }

            return null;
        }

        #endregion
        #endregion
    }
}
