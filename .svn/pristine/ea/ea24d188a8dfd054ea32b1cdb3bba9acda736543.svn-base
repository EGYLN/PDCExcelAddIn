using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// The PicklistHandler manages the sheet(s) with the reference data
    /// For each PDC workbook one sheet with picklist may exist.
    /// PDC worksheets have to use the picklist sheet from the same workbook.
    /// </summary>
    [Serializable]
  [ComVisible(false)]
  public class PicklistHandler
  {
    public const string NAMED_RANGE_PICKLIST_ID = "PDC_PicklistId";
    private const string PICKLIST_SHEETNAME = "PDC_PicklistProviderSheet";
    private static readonly object LOCK = new object();
    //Temporarily set in ThePicklistHandler
    [NonSerialized]
    private Excel.Worksheet picklistSheet;
    [NonSerialized]
    private Excel.Workbook workbook;

    private string pickListGUID;

    private IDictionary<decimal, Lib.Picklist> picklistsById = new Dictionary<decimal, Lib.Picklist>();
    private IDictionary<int, Lib.PredefinedParameter> predefinedByVariableId = new Dictionary<int, Lib.PredefinedParameter>();
    private IDictionary<int, bool> predefinedInitialized = new Dictionary<int, bool>();
    private static IDictionary<string,PicklistHandler> picklistHandlers = new Dictionary<string, PicklistHandler>();
    private static IDictionary<Excel.Worksheet, PicklistHandler> sheetToPicklistHandler = new Dictionary<Excel.Worksheet, PicklistHandler>();

    /// <summary>
    /// Deletes the picklist handler for the specified workbook
    /// </summary>
    /// <param name="aWorkbook"></param>
    public static void DeletePicklistHandler(Excel.Workbook aWorkbook)
    {
      PicklistHandler tmpHandler = ThePicklistHandler(aWorkbook, false);
      if (tmpHandler == null)
      {
        return;
      }
      if (picklistHandlers.ContainsKey(tmpHandler.pickListGUID))
      {
        picklistHandlers.Remove(tmpHandler.pickListGUID);
      }
      //Clear from sheet lookup map
      List<Excel.Worksheet> tmpLookupSheets = new List<Excel.Worksheet>();
      foreach (KeyValuePair<Excel.Worksheet, PicklistHandler> tmpPair in sheetToPicklistHandler)
      {
        if (tmpPair.Value.pickListGUID == tmpHandler.pickListGUID)
        {
          tmpLookupSheets.Add(tmpPair.Key);
        }
      }
      foreach (Excel.Worksheet tmpSheet in tmpLookupSheets)
      {
        if (sheetToPicklistHandler.ContainsKey(tmpSheet))
        {
          sheetToPicklistHandler.Remove(tmpSheet);
        }
      }
      tmpLookupSheets.Clear();
      tmpHandler.picklistSheet = null;
      if (tmpHandler.predefinedInitialized != null)
      {
        tmpHandler.predefinedInitialized.Clear();
      }
      if (tmpHandler.predefinedByVariableId != null)
      {
        tmpHandler.predefinedByVariableId.Clear();
      }
    }

    public static PicklistHandler ThePicklistHandler(Excel.Worksheet aSheet)
    {
      if (sheetToPicklistHandler.ContainsKey(aSheet))
      {
        return sheetToPicklistHandler[aSheet];
      }
      return ThePicklistHandler((Excel.Workbook)aSheet.Parent, true);
    }

    /// <summary>
    /// Returns the appropriate PicklistHandler instance for the specified workbook
    /// </summary>
    /// <param name="aWB">The workbook of the desired PicklistHandler</param>
    /// <param name="aCreateFlag">If set to true a PickListHandler will be generated on the fly,
    /// if the workbook does not have one
    /// </param>
    /// <returns>The PicklistHandler for the workbook or null</returns>
    public static PicklistHandler ThePicklistHandler(Excel.Workbook aWB, bool aCreateFlag)
    {
      //One Picklist handler per Workbook            
      Excel.Worksheet tmpSheet = null;
      PicklistHandler tmpPicklistHandler = null;
      try
      {
        object tmpRangeCand = aWB.Names.Item(NAMED_RANGE_PICKLIST_ID, Type.Missing, Type.Missing);
        if (tmpRangeCand is Excel.Name)
        {
          tmpRangeCand = ((Excel.Name)tmpRangeCand).RefersToRange;
        }
        if (tmpRangeCand is Excel.Range)
        {
          Excel.Range tmpRange = (Excel.Range)tmpRangeCand;
          tmpSheet = (Excel.Worksheet)((Excel.Range) tmpRange).Parent;
          string tmpKey = "" + tmpRange.Text;
          if (picklistHandlers.ContainsKey(tmpKey))
          {
            tmpPicklistHandler = picklistHandlers[tmpKey];
            tmpPicklistHandler.picklistSheet = tmpSheet;
            tmpPicklistHandler.workbook = aWB;
          }
        }                
      }
#pragma warning disable 0168
      catch (Exception e)
      {
        tmpPicklistHandler = null;
      }
#pragma warning restore 0168
      if (tmpPicklistHandler == null && aCreateFlag)
      {
        tmpPicklistHandler = CreatePickListHandler(aWB);
      }
      return tmpPicklistHandler;
    }

    /// <summary>
    /// Creates a PicklistHandler for the specified workbook.
    /// The PicklistHandler uses a hidden excel sheet, where it places 
    /// the picklist arrays.
    /// </summary>
    /// <param name="aWorkbook"></param>
    /// <returns></returns>
    private static PicklistHandler CreatePickListHandler(Excel.Workbook aWorkbook)
    {
      PicklistHandler tmpHandler = new PicklistHandler();
      Excel.Worksheet tmpSheet = ExcelUtils.TheUtils.CreateNewSheet(aWorkbook, PICKLIST_SHEETNAME, null); 
      tmpSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVeryHidden;            
      string tmpGUID = Guid.NewGuid().ToString();
      tmpHandler.pickListGUID = tmpGUID;
      Excel.Range tmpCell = (Excel.Range) tmpSheet.Cells[1,1];
      tmpCell.Value2 = tmpGUID;
      Excel.Name tmpName = aWorkbook.Names.Add(NAMED_RANGE_PICKLIST_ID, tmpCell, true, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
      tmpSheet.AutoFilterMode = false;
      tmpHandler.picklistSheet = tmpSheet;
      tmpHandler.workbook = aWorkbook;
      string tmpNameName = tmpName.Name;
      string tmpNameNameLocal = tmpName.NameLocal;
      object tmpNameReferTo = tmpName.RefersTo;
      picklistHandlers.Add(tmpGUID, tmpHandler);
      tmpHandler.predefinedByVariableId = Globals.PDCExcelAddIn.PdcService.PredefinedParameter();
      if (tmpHandler.predefinedByVariableId == null)
      {
        tmpHandler.predefinedByVariableId = new Dictionary<int, Lib.PredefinedParameter>();
      }
      foreach (int tmpId in tmpHandler.predefinedByVariableId.Keys)
      {
        tmpHandler.predefinedInitialized.Add(tmpId, false);
      }
      return tmpHandler;
    }

    /// <summary>
    /// Returns the worksheet which holds the picklist infos
    /// </summary>
    public Excel.Worksheet Worksheet
    {
      get
      {
        return picklistSheet;
      }
    }

    /// <summary>
    /// Gets the reference data for the specified predefined parameter and puts
    /// it into an Excel list.
    /// </summary>
    /// <param name="aParameter"></param>
    private void LoadPredefinedValues(Lib.PredefinedParameter aParameter)
    {
      if (predefinedInitialized.ContainsKey(aParameter.VariableId) && predefinedInitialized[aParameter.VariableId])
      {
        return;
      }
      if ((aParameter.Tablename != null && aParameter.Tablename.Trim() != "") || (aParameter.Servicename != null && aParameter.Servicename.Trim() != ""))
      {
        if (!aParameter.DataLoaded)
        {
          aParameter.PicklistValues = Globals.PDCExcelAddIn.PdcService.GetReferenceData(aParameter.Servicename, aParameter.Tablename);
          aParameter.DataLoaded = true;
        }
        string tmpName = null;
        if (aParameter.Servicename != null && aParameter.Servicename.Trim() != "")
        {
          tmpName = "S_" + aParameter.Servicename.Trim();
        }
        else
        {
          tmpName = "T_" + aParameter.Tablename.Trim();
        }
        if (aParameter.PicklistValues != null && aParameter.PicklistValues.Count > 0)
        {
          aParameter.Tag = "Picklist_" + tmpName;
          CreateExcelList(aParameter.PicklistValues, aParameter.Tag);
        }
      }
    }

    /// <summary>
    /// Inserts the specified picklist into an excel list.
    /// </summary>
    /// <param name="aPickList"></param>
    private void InsertPicklist(Lib.Picklist aPickList)
    {
      if (picklistsById.ContainsKey(aPickList.PicklistId))
      {
        return;
      }
      picklistsById.Add(aPickList.PicklistId, aPickList);
      aPickList.Tag = "Picklist_" + aPickList.PicklistId;
      CreateExcelList(aPickList.Values, aPickList.Tag);
    }

    /// <summary>
    /// Creates a list of values in a new column of the picklist sheet and gives it
    /// the specified list range name
    /// </summary>
    /// <param name="theValues">The list values </param>
    /// <param name="aListRangeName">The list name</param>
    private void CreateExcelList(List<object> theValues, string aListRangeName)
    {
      //shift away any existing lists
      Excel.Range tmpThirdColumn = ((Excel.Range)picklistSheet.Cells[1, 3]).EntireColumn;
      tmpThirdColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
      Excel.Range tmpListRange = ExcelUtils.TheUtils.GetRange(picklistSheet, picklistSheet.Cells[2, 3], picklistSheet.Cells[2 + theValues.Count - 1, 3]);
      object[,] tmpPicklistValues = new object[theValues.Count, 1];
      for (int i = 0; i < theValues.Count; i++)
      {
        tmpPicklistValues[i, 0] = theValues[i];
      }
      tmpListRange.set_Value(Type.Missing, tmpPicklistValues);
      workbook.Names.Add(aListRangeName, tmpListRange, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);            
    }

    /// <summary>
    /// Updates the pick lists for the specified test definition
    /// Can be used to optimize the number of server calls before
    /// GetPicklistLink is called.
    /// </summary>
    /// <param name="aTestdefinition"></param>
    public void UpdatePickListCache(Lib.Testdefinition aTestdefinition)
    {
      Dictionary<decimal, decimal> tmpPickListIds = new Dictionary<decimal, decimal>();
      IDictionary<int, Lib.PredefinedParameter> tmpPredefined = Globals.PDCExcelAddIn.PdcService.PredefinedParameter();

      foreach (Lib.TestVariable tmpVariable in aTestdefinition.Variables)
      {
        if (tmpVariable.PicklistId != null && !tmpPickListIds.ContainsKey(tmpVariable.PicklistId.Value))
        {
          tmpPickListIds.Add(tmpVariable.PicklistId.Value, tmpVariable.PicklistId.Value);
          continue;
        }
        if (tmpPredefined.ContainsKey(tmpVariable.VariableId))
        {
          Lib.PredefinedParameter tmpParameter = tmpPredefined[tmpVariable.VariableId];
          if (tmpParameter.PicklistId != null && !tmpPickListIds.ContainsKey(tmpParameter.PicklistId.Value))
          {
            tmpPickListIds.Add(tmpParameter.PicklistId.Value, tmpParameter.PicklistId.Value);
          }
        }
      }
      List<decimal> tmpPickListsToGet = new List<decimal>();
      foreach (decimal tmpPicklistId in tmpPickListIds.Keys)
      {
        if (!picklistsById.ContainsKey(tmpPicklistId))
        {
          tmpPickListsToGet.Add(tmpPicklistId);
        }
      }
      if (tmpPickListsToGet.Count == 0)
      {
        return;
      }
      // Get Picklists
      Dictionary<decimal, Lib.Picklist> tmpLoaded = Globals.PDCExcelAddIn.PdcService.LoadPicklists(tmpPickListsToGet);
      // Insert each picklist
      foreach (Lib.Picklist tmpPicklist in tmpLoaded.Values)
      {
        InsertPicklist(tmpPicklist);
      }
    }

    /// <summary>
    /// Returns the Predefined Parameter for the specified variable or null if the variable is not 
    /// a predefined parameter.
    /// </summary>
    /// <param name="aVariable"></param>
    /// <returns></returns>
    public Lib.PredefinedParameter GetPredefinedParameter(Lib.TestVariable aVariable) 
    {
      if (aVariable != null && predefinedByVariableId.ContainsKey(aVariable.VariableId))
      {
        return predefinedByVariableId[aVariable.VariableId];
      }
      return null;
    }
    /// <summary>
    /// Returns a link to the Excel pick list. The Picklist cache should be updated 
    /// by calling UpdatePickListCache before calling this method
    /// </summary>
    /// <param name="aVariable"></param>
    /// <returns></returns>
    public string GetPicklistLink(Lib.TestVariable aVariable, out int? preferredSize)
    {
      preferredSize = null;
      if (aVariable.PicklistId != null && picklistsById.ContainsKey(aVariable.PicklistId.Value))
      {
        Lib.Picklist tmpPickList = picklistsById[aVariable.PicklistId.Value];
        preferredSize = CalcPreferedSize(tmpPickList.Values);
        return tmpPickList.Tag;
      }
      if (predefinedByVariableId != null && predefinedByVariableId.ContainsKey(aVariable.VariableId))
      {
        Lib.PredefinedParameter tmpPredefined = predefinedByVariableId[aVariable.VariableId];
        LoadPredefinedValues(tmpPredefined);
        preferredSize = CalcPreferedSize(tmpPredefined.PicklistValues);
        return tmpPredefined.Tag;
      }
      return null;
    }

    /// <summary>
    /// Calculates the preferred width for the combo box given by the largest string
    /// </summary>
    /// <param name="list"></param>
    /// <returns></returns>
    private int? CalcPreferedSize(List<object> list)
    {
      if (list == null || list.Count == 0)
      {
        return null;
      }
      int tmpSize = 0;
      foreach (object tmpObject in list)
      {
        if (tmpObject != null)
        {
          tmpSize = Math.Max(tmpSize, tmpObject.ToString().Length);
        }
      }
      return tmpSize;
    }

    internal static void Deserialize(PicklistHandler picklistHandler, Excel.Workbook aWorkbook, Excel.Worksheet aPicklistSheet)
    {
      if (picklistHandler == null)
      {
        return;
      }
      if (picklistHandlers.ContainsKey(picklistHandler.pickListGUID))
      {
        picklistHandlers[picklistHandler.pickListGUID] = picklistHandler;
      }
      else
      {
        picklistHandlers.Add(picklistHandler.pickListGUID, picklistHandler);
      }
    }
  }
}
