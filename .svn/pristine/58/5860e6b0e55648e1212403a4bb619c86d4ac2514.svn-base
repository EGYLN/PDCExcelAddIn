using System;
using System.Collections.Generic;
using System.Text;
using WS = BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// 
    /// </summary>
    public class ColorScheme
    {
        /// <summary>
        /// 
        /// </summary>
        public class Color {
            System.Drawing.Color color;
            int oleColor;
            /// <summary>
            /// Initializes a PDC Color from the corresponding system color
            /// </summary>
            /// <param name="aColor"></param>
            public Color(System.Drawing.Color aColor) {
                color = aColor;
                oleColor = System.Drawing.ColorTranslator.ToOle(color);
            }
            /// <summary>
            /// 
            /// </summary>
            public int OleColor {
                get
                {
                    return oleColor;
                }
            }
            /// <summary>
            /// 
            /// </summary>
            public System.Drawing.Color SystemColor {
                get {
                    return color;
                }
            }
        }

        private struct ColorParam
        {
            internal string name;
            internal int r;
            internal int g;
            internal int b;
        }

        Dictionary<string, Color> colors;
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_COMPOUND_INFO = "HeaderCompoundInfo";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_PDC_INFO = "HeaderPDCInfo";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_DERIVED_RESULT = "HeaderDerivedResult";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_PARAMETER = "HeaderParameter";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_BINARY = "HeaderBinary";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_COMMENT = "HeaderComment";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_RESULT = "HeaderMeasurementResult";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_VARIABLE = "HeaderMeasurementVariable";
        /// <summary>
        /// 
        /// </summary>
        public const string HEADER_ANNOTATION = "HeaderAnnotation";
        /// <summary>
        /// 
        /// </summary>
        public const string BACKGROUND_MANDATORY = "BackgroundMandatory";

        internal ColorScheme(Lib.PDCService aService)
        {
            colors = new Dictionary<string, Color>();
            // Add some default values first
            colors.Add(HEADER_COMPOUND_INFO, new Color(System.Drawing.Color.Gray));
            colors.Add(HEADER_PDC_INFO, new Color(System.Drawing.Color.LightYellow));
            colors.Add(HEADER_DERIVED_RESULT, new Color(System.Drawing.Color.Red));
            Color tmpColor = new Color(System.Drawing.Color.Yellow);
            colors.Add(HEADER_PARAMETER, tmpColor);
            colors.Add(HEADER_BINARY, tmpColor);
            colors.Add(HEADER_COMMENT, tmpColor);
            colors.Add(HEADER_VARIABLE, tmpColor);
            colors.Add(HEADER_RESULT, new Color(System.Drawing.Color.Blue));
            colors.Add(HEADER_ANNOTATION, new Color(System.Drawing.Color.LightBlue));
            colors.Add(BACKGROUND_MANDATORY, new Color(System.Drawing.Color.IndianRed));
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Getting the color configuration from the server");
            try
            {
                //initFromServer(aService);
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Could not connect to get color configuration", e);
            }
        }

        private void initFromServer(Lib.PDCService aService)
        {
            WS.Output tmpOutput = aService.callService(PDCConstants.C_OP_GET_COLOR_CONFIGURATION, null);
            if (tmpOutput == null || tmpOutput.output == null)
            {
                return;
            }
            Dictionary<int, ColorParam> tmpColors = new Dictionary<int, ColorParam>();
            foreach (WS.PDCParameter tmpParam in tmpOutput.output)
            {
                if (!tmpParam.position.HasValue)
                {
                    Util.PDCLogger.TheLogger.LogWarning(Util.PDCLogger.LOG_NAME_LIB, "Color config without position");
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
                }
                tmpColors[tmpPosition] = tmpColorParam;
            }
            foreach (ColorParam tmpColParam in tmpColors.Values)
            {
                if (tmpColParam.name != null)
                {
                    if (colors.ContainsKey(tmpColParam.name))
                    {
                        colors[tmpColParam.name] = new Color(System.Drawing.Color.FromArgb(
                            tmpColParam.r, tmpColParam.g, tmpColParam.b));
                    }
                    else
                    {
                        colors.Add(tmpColParam.name, new Color(System.Drawing.Color.FromArgb(
                            tmpColParam.r, tmpColParam.g, tmpColParam.b)));
                    }
                }
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="aKey"></param>
        /// <returns></returns>
        public Color this[string aKey] {
            get
            {
                return colors.ContainsKey(aKey) ? colors[aKey] : null;
            }
        }
    }
}
