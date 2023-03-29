using System;
using System.Runtime.InteropServices;
using LibUtil = BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Default validator for table cell input.
    /// </summary>
    [ComVisible(false)]
    public partial class CellValidator
    {
        private object missing = global::System.Type.Missing;

        
        /// <summary>
        /// Validates a column value against the test definition.
        /// </summary>
        /// <param name="aColumn">Specifies the column with its properties</param>
        /// <param name="aValue">The current value, which is validated against the column properties</param>
        /// <param name="anOriginalValue">The original value if it exists</param>
        /// <returns>Returns null if the validation was successful. Otherwise the localized validation message is returned</returns>
        public string Validate(ListColumn aColumn, object aValue, object anOriginalValue)
        {
            
            if (aColumn.TestVariable == null)
            {
                return null;
            }
            Lib.TestVariable tmpVar = aColumn.TestVariable;
            if (aValue == null)
            {
                if (aColumn.TestVariable.DefaultValue != null)
                {
                    return null;
                }
                if (aColumn.TestVariable.IsMandatory)
                {
                    return Properties.Resources.VALIDATOR_MANDATORY;
                }
                if (aColumn.TestVariable.IsDifferentiating)
                {
                    return Properties.Resources.VALIDATOR_DIFFERENTIATING;
                }
                return null;
            }
            if ("".Equals(aValue))
            {
                return null;
            }
            if (aColumn.IsHyperLink && tmpVar.IsBinaryParameter())
            {
                return CheckFileProperties(aValue);
            }
            if (tmpVar.IsNumeric())
            {
                if (aValue is string)
                {
                    string tmpPrefix = null;
                    object tmpValueObject = Lib.PDCConverter.Converter.RemoveWellKnownPrefix((string)aValue, out tmpPrefix, Globals.PDCExcelAddIn.PdcService.Prefixes());
                    Lib.TestVariableValue tmpTestVariableValue = new Lib.TestVariableValue(tmpVar.VariableId);
                    if (!Lib.PDCConverter.Converter.DoubleToString(tmpValueObject,ExcelUtils.TheUtils.GetExcelNumberSeparators(),tmpTestVariableValue))
                    {
                      return string.Format(Properties.Resources.VALIDATOR_NUMERIC_CONVERSION, aValue);
                    }
                        
                }
#pragma warning restore 0168
            }
            return null;
        }

        private string CheckFileProperties(object aValue)
        {
            Lib.ClientConfiguration tmpConfig = Globals.PDCExcelAddIn.ClientConfiguration;
            string tmpFileName = "";
            string tmpUrl = "";
            if (aValue is string[]) {
                string[] tmpPair = (string[]) aValue;
                tmpUrl = tmpPair[0];
                tmpFileName = tmpPair[1];
            } else {
                tmpFileName = tmpFileName + aValue;
                tmpUrl = tmpFileName;
            }
            if (tmpUrl.StartsWith(Properties.Settings.Default.ImageServletPath))
            {
                return null;
            }
            int tmpLastDotPos = tmpFileName.LastIndexOf('.');
            if (tmpLastDotPos <= 0)
            {
                return string.Format(Properties.Resources.VALIDATOR_UNKNOWN_FILE_TYPE, new object[] {
                    tmpFileName, tmpConfig.SupportedTypesAsString});
            }
            if (tmpLastDotPos >= tmpFileName.Length - 2)
            {
                return string.Format(Properties.Resources.VALIDATOR_UNKNOWN_FILE_TYPE, new object[] {
                    tmpFileName, tmpConfig.SupportedTypesAsString});
            }
            string tmpExtension = tmpFileName.Substring(tmpLastDotPos + 1, tmpFileName.Length - (tmpLastDotPos + 1));
            tmpExtension = tmpExtension.ToLower();
            if (!tmpConfig.SupportedTypes.ContainsKey(tmpExtension))
            {
                return string.Format(Properties.Resources.VALIDATOR_UNKNOWN_FILE_TYPE, new object[] {
                    tmpFileName, tmpConfig.SupportedTypesAsString});
            }
            long tmpLength = -1;
            try
            {
                tmpLength = LibUtil.StreamUtil.GetSize(tmpUrl);
            }
#pragma warning disable 0168
            catch (Exception e)
            {
                return string.Format(Properties.Resources.VALIDATOR_BINARY_NOT_FOUND, new object[] {
                    tmpFileName});
            }
#pragma warning restore 0168
            if (tmpLength > tmpConfig.SupportedTypes[tmpExtension])
            {
                return string.Format(Properties.Resources.VALIDATOR_BINARY_TOO_LARGE,
                    new object[] { tmpFileName, tmpLength, "" + tmpConfig.SupportedTypes[tmpExtension]});
            }
            return null;
        }
    }
}
