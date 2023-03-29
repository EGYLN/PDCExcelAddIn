using System;
using System.Collections.Generic;
using System.Globalization;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// Central converter for type conversions from PDCParameters value*s to dotnet types and vice versa.
  /// </summary>
  public class PDCConverter
  {
    /// <summary>
    /// Format used for string formatting of timestamps
    /// </summary>
    public const string TIMESTAMP_FORMAT = "yyyy-MM-dd-HH-mm-ss";
    /// <summary>
    /// Format used for string formatting of dates.
    /// </summary>
    public const string DATE_FORMAT = "yyyy-MM-dd";
    public const string INPUT_DATE_FORMAT = "d.M.yyyy";

    private static string formatNonScientificNumber = ""; // this will have sothing like #.#############################....####
    private static PDCConverter mySingleton;
    private static NumberFormatInfo webServiceNfi = new NumberFormatInfo();  // this is the NumberFormat for the Webservice

    #region methods

    #region ConvertFromDate
    /// <summary>
    /// Parses a string with the date format
    /// </summary>
    /// <param name="aDateTime"></param>
    /// <param name="aFormat"></param>
    /// <returns></returns>
    public string ConvertFromDate(object aDateTime, string aFormat)
    {
      if (aDateTime == null || aDateTime is string)
      {
        return (string)aDateTime;
      }
      if (aDateTime is DateTime)
      {
        return ((DateTime)aDateTime).ToString(aFormat);
      }
      return "" + aDateTime;
    }
    #endregion

    #region FromDate
    /// <summary>
    /// Converts the specified object into a date string.
    /// </summary>
    /// <param name="aDateTime"></param>
    /// <returns></returns>
    public string FromDate(object aDateTime)
    {
      return ConvertFromDate(aDateTime, DATE_FORMAT);
    }
    #endregion

    #region FromDecimal
    /// <summary>
    /// Converts the specified decimal value to a string.
    /// </summary>
    /// <param name="aDecimalValue"></param>
    /// <param name="nfi">NumberFormatInfo with the current Excel-Setting (comma and groupseparator)</param>
    /// <returns>String value of the aDecimalvalue</returns>
    public string FromDecimal(decimal aDecimalValue,NumberFormatInfo nfi)
    {
      return Convert.ToString(aDecimalValue,nfi);
    }
    /// <summary>
    /// Converts the specified decimal value to a string.
    /// </summary>
    /// <param name="aDecimalValue"></param>
    /// <param name="nfi">NumberFormatInfo with the current Excel-Setting (comma and groupseparator)</param>
    /// <returns>String value of the aDecimalvalue</returns>
    //public double FromDecimal(string aDecimalValue, NumberFormatInfo nfi)
    //{

    //  return System.Convert.ToString(Double.Parse(aDecimalValue,webServiceNfi), nfi);
    //}
    #region MergeWithPrefix
    /// <summary>
    /// Adds the prefix to the string value.
    /// </summary>
    /// <param name="prefix"></param>
    /// <param name="value"></param>
    /// <returns></returns>
    public object MergeWithPrefix(string prefix, string value)
    {
      string preFix = prefix.Trim();
      if (preFix == "=")
      {
        preFix = "'=";
      }
      return preFix + (value == null ? "" : " " + value);
    }
    #endregion

    public object NumericString2Double(string aStringNumericValue, string aPrefix, NumberFormatInfo nfi)
    {
      double tmpDoubleValue = Double.Parse(aStringNumericValue, webServiceNfi);
      if (aPrefix != null && aPrefix.Trim() != "")
      {
        string tmpStingValue = Convert.ToString(tmpDoubleValue,nfi);
        return MergeWithPrefix(aPrefix, tmpStingValue);
      }
      return tmpDoubleValue;
    }

    
    #endregion

    #region FromTimestamp
    /// <summary>
    /// Converts the specified object into a timestamp string
    /// </summary>
    /// <param name="aDateTime"></param>
    /// <returns></returns>
    public string FromTimestamp(object aDateTime)
    {
      return ConvertFromDate(aDateTime, TIMESTAMP_FORMAT);
    }
    #endregion

    #region ParseDate
    /// <summary>
    /// Parses a string for a date/timestamp with the given format
    /// </summary>
    /// <param name="aString"></param>
    /// <param name="aFormat"></param>
    /// <returns></returns>
    public DateTime? ParseDate(string aString, string aFormat)
    {
      DateTime tmpDate;
      if (aString == null)
      {
        return null;
      }
      if (DateTime.TryParseExact(aString, aFormat, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out tmpDate))
      {
        return tmpDate;
      }
        return null;
    }
    #endregion

    #region RemoveWellKnownPrefix
    /// <summary>
    /// Removes an optional prefix from the string aValue and returns it in aPrefix
    /// </summary>
    /// <param name="aValue"></param>
    /// <param name="aPrefix"></param>
    /// <param name="thePrefixes"></param>
    /// <returns></returns>
    public string RemoveWellKnownPrefix(string aValue, out string aPrefix, List<string> thePrefixes)
    {
      foreach (string tmpPrefix in thePrefixes)
      {
        if (aValue.StartsWith(tmpPrefix) || aValue.StartsWith("'" + tmpPrefix))
        {
          aPrefix = tmpPrefix;
          int tmpToStrip = tmpPrefix.Length;
          if (aValue.StartsWith("'")) //'= prevents formula interpretation of = by Excel
          {
            tmpToStrip += 1;
          }
          string tmpNewValue = aValue.Remove(0, tmpToStrip).Trim();
          return tmpNewValue;
        }
      }
      aPrefix = null;
      return aValue;
    }
    #endregion

    #region ToBool
    /// <summary>
    /// Returns true if the specified string has the value "Y" or "y" and false otherwise
    /// </summary>
    /// <param name="aValue"></param>
    /// <returns></returns>
    public bool ToBool(string aValue)
    {
      return aValue != null && aValue.ToUpper() == "Y";
    }
    #endregion

    #region ToDate
    /// <summary>
    /// Tries for convert the given object to a date
    /// </summary>
    /// <param name="aDateRepr"></param>
    /// <returns></returns>
    public DateTime? ToDate(object aDateRepr)
    {
      if (aDateRepr == null)
      {
        return null;
      }
      if (aDateRepr is DateTime)
      {
        return (DateTime)aDateRepr;
      }
      return ToDate("" + aDateRepr);
    }

    /// <summary>
    /// Converts a date string from the server to a DataTime
    /// </summary>
    /// <param name="aString">A string which represents a date</param>
    /// <returns></returns>
    public DateTime? ToDate(string aString)
    {
      return ParseDate(aString, DATE_FORMAT);
    }
    #endregion

    #region ToDecimal
    /// <summary>
    /// Converts the specified value to a decimal. Returns null if conversion is not possible.
    /// </summary>
    /// <param name="aValue"></param>
    /// <param name="nfi">NumberFormatInfo with the current Excel-Setting (comma and groupseparator)</param>
    /// <returns>decimal value of aValue</returns>
    public decimal? ToDecimal(object aValue,NumberFormatInfo nfi)
    {
      if (aValue is decimal)
      {
        return (decimal)aValue;
      }
      if (aValue is long)
      {
        return (long)aValue;
      }
      if (aValue is double)
      {
        double tmpDouble = (double)aValue;
        decimal tmpDecimal = (decimal)tmpDouble;

        if (tmpDouble != 0 && tmpDecimal == 0) 
        {
          throw new ApplicationException("Invalid ToDecimal conversion of doublevalue : '" + tmpDouble + "'");
        }
        return tmpDecimal;
      }
      if (aValue is string)
      {
        try
        {
          return decimal.Parse((string)aValue,nfi);
        }
#pragma warning disable 0168
        catch (Exception e)
        {
          return null;
        }
#pragma warning restore 0168
      }
      return null;
    }
    /// <summary>
    /// Writes the value to testVariable.valueChar while making sure, it is a number, other the function returns false 
    /// </summary>
    /// <param name="aValue">value to be converted to a number</param>
    /// <param name="nfi">NumberFormatInfo with the current Excel-Setting (comma and groupseparator)</param>
    /// <returns>true if value was a number and successfull converted to </returns>
    public bool DoubleToString(object aValue,NumberFormatInfo nfi,TestVariableValue testVariableValue)
    {
      try
      {
          if (aValue is double)
        {
          testVariableValue.ValueChar = String.Format(webServiceNfi, formatNonScientificNumber, aValue);
          return true;
        }
          if (aValue is string)
          {
          
              double tmpValue = Double.Parse((string) aValue,nfi);
              testVariableValue.ValueChar = String.Format(webServiceNfi, formatNonScientificNumber, tmpValue);
              return true;
          }
      } catch (Exception)
      {      
        return false;
      }
      return false;
    }

    #endregion

    #region ToLong
    /// <summary>
    /// Converts the specified value to a long. Returns null if conversion is not possible.
    /// </summary>
    /// <param name="aValue"></param>
    /// <returns></returns>
    public long? ToLong(object aValue)
    {
      if (aValue is decimal)
      {
        return (long)(decimal)aValue;
      }
      if (aValue is long)
      {
        return (long)aValue;
      }
      if (aValue is double)
      {
        return (long)(double)aValue;
      }
      if (aValue is string)
      {
        long tmpResult = 0;
        if (long.TryParse((string) aValue, out tmpResult))
        {
          return tmpResult;
        }
      }
      return null;
    }
    #endregion

    #region ToString
    /// <summary>
    /// Returns a "Y" or a "N"
    /// </summary>
    /// <param name="aValue"></param>
    /// <returns></returns>
    public string ToString(bool aValue)
    {
      return aValue ? "Y" : "N";
    }
    #endregion

    #region ToTimestamp
    /// <summary>
    /// Parses a string with the timestamp format
    /// </summary>
    /// <param name="aString"></param>
    /// <returns></returns>
    public DateTime? ToTimestamp(string aString)
    {
      return ParseDate(aString, TIMESTAMP_FORMAT);
    }
    #endregion

    #endregion

    #region properties

    #region Converter
    /// <summary>
    /// Returns the PDCConverter singleton
    /// </summary>
    public static PDCConverter Converter
    {
      get
      {
        if (mySingleton==null) 
        {
          mySingleton =new PDCConverter();
          formatNonScientificNumber = "{0:0." + new string('#', 300) + "}";
          webServiceNfi.NumberDecimalSeparator = ".";
          webServiceNfi.NumberGroupSeparator = ",";
        }
        return mySingleton;
      }
    }
    #endregion
    
    #endregion
  }
}
