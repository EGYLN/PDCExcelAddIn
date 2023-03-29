using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /// <summary>
    /// da in allen Measurementtabellen derselbe Name für die gleichen Variablen verwendet werden.
    /// -> Tabellennummer muss eingeplegt werden.
    /// </summary>
    [Serializable]
  public class MeasurementHandler : MultipleMeasurementTableHandler
  {
    protected const string LIST_PREFIX = "Meas_";
    protected const string PROTECT_PASSWORD = "pdcPRotectME";
    protected const string ALLOW_EDIT_RANGE_NAME = "Mdata";
    internal string baseSheetName;
    internal int nrOfTables;
    internal int nrOfSheets;
    internal Lib.Testdefinition testDefinition;

    /// <summary>
    /// Mapping from measurement table no to PDCListObject
    /// </summary>
    internal IDictionary<int, PDCListObject> measurementTableMap = new Dictionary<int, PDCListObject>();
    internal IDictionary<int, SheetInfo> sheets = new Dictionary<int, SheetInfo>();

    /// <summary>
    /// Measurement list range name to sheet
    /// </summary>
    internal IDictionary<string, SheetInfo> tableLinks = new Dictionary<string, SheetInfo>();

    public ICollection<SheetInfo> GetUsedSheetInfos()
    {
      return sheets == null ? null : sheets.Values;
    }

    public ICollection<PDCListObject> MeasurementTables
    {
      get
      {
        return measurementTableMap.Values;
      }
    }

    public MeasurementHandler(Excel.Worksheet aSheet, Lib.Testdefinition aTD):base(aTD)
    {
      testDefinition = aTD;
      nrOfTables = 0;
      nrOfSheets = 0;
      baseSheetName = LIST_PREFIX + aTD.TestName;
      if (baseSheetName.Length > 15)
      {
        baseSheetName = baseSheetName.Substring(0, 15);
      }
      baseSheetName += "(" + aTD.TestNo + "_" + aTD.Version + ")";
    }

  }

}
