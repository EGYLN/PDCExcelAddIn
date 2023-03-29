using System;
using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
  /// <summary>
  /// Holds the value arrays for experiment and measurements values as they can be read from and set to Excel.
  /// </summary>
  [Serializable]
  internal struct ExperimentAndMeasurementValues
  {
    public object[,] experimentValues;
    public MeasurementSheetData measurementValues;
    public ExperimentAndMeasurementValues(object[,] theExperimentValues)
    {
      experimentValues = theExperimentValues;
      measurementValues = null;
    }
    public ExperimentAndMeasurementValues(object[,] theExperimentValue, MeasurementSheetData theMeasurementData)
    {
      experimentValues = theExperimentValue;
      measurementValues = theMeasurementData;
    }
  }
  
  /// <summary>
  /// Helper class used for read/write-operations on measurement sheet.
  /// </summary>
  [Serializable]
  internal class MeasurementSheetData
  {
    private Rectangle rangeArea;
    private Excel.Worksheet sheet;
    private Excel.Range range;
    private object[,] values;
    Dictionary<string, object[,]> tableValues = new Dictionary<string,object[,]>();

    #region properties

    #region Range
    /// <summary>
    /// The used data range.
    /// </summary>
    internal Excel.Range Range
    {
      get
      {
        return range;
      }
      set
      {        
        range = value;
        tableValues.Clear();
        if (range == null)
        {
          sheet = null;
          rangeArea = new Rectangle(0, 0, 0, 0);
          values = null;
        }
        else
        {
          sheet = (Excel.Worksheet) range.Parent;
          rangeArea = new Rectangle(range.Column, range.Row, range.Columns.Count, range.Rows.Count);
          PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "MeasurementSheetArea: " + rangeArea.ToString());
          values = (object[,]) range.get_Value(Type.Missing);
        }
      }
    }
    #endregion

    #region RangeArea
    /// <summary>
    /// Describes the location of the range within the worksheet
    /// </summary>
    internal Rectangle RangeArea
    {
      get
      {
        return rangeArea;
      }
    }
    #endregion

    #region Sheet
    /// <summary>
    /// Reference to the associated worksheet
    /// </summary>
    internal Excel.Worksheet Sheet
    {
      get
      {
        return sheet;
      }
    }
    #endregion

    #region this
    /// <summary>
    /// Accessor for the table values.
    /// </summary>
    /// <param name="aTablename">The name of a (measurement) table</param>
    /// <returns>The value matrix for the specified table</returns>
    internal object[,] this[string aTablename]
    {
      get
      {
        if (tableValues.ContainsKey(aTablename))
        {
          return tableValues[aTablename];
        }
        return null;
      }
      set
      {
        if (tableValues.ContainsKey(aTablename))
        {
          tableValues[aTablename] = value;
        }
        else
        {
          tableValues.Add(aTablename, value);
        }
      }
    }
    #endregion

    #region Values
    /// <summary>
    /// The value matrix of the used range.
    /// </summary>
    internal object[,] Values
    {
      get
      {
        return values;
      }
      set
      {
        values = value;
      }
    }
    #endregion

    #endregion
  }
}
