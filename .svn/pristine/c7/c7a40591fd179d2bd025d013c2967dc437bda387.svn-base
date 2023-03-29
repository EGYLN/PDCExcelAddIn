using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using Excel = Microsoft.Office.Interop.Excel;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using System.Xml.Serialization;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Contains all PDC related information about a worksheet
    /// </summary>
    [Serializable]
  [ComVisible(false)]
  public class SheetInfo:ISerializable
  {
    private const string SER_ADDITIONALSHEETS = "additionalSheets";
    private const string SER_AREMEASUREMENTSLOADED = "areMeasurementsLoaded";
    private const string SER_IDENTIFIER = "identifier";
    private const string SER_MAINSHEET= "mainSheet";
    private const string SER_MAINTABLE = "mainTable";
    private const string SER_MEASUREMENTTABLES = "measurementTables";
    private const string SER_SHAPEFILENAMES = "shapeFileNames";
    private const string SER_TESTDEFINITION = "testDefinition";

    List<SheetInfo>                     myAdditionalSheets = new List<SheetInfo>();
    private bool                        myAreMeasurementsLoaded;
    Excel.Worksheet                     myExcelSheet;
    private bool                        myExcelSheetRemoved = false;
    object                              myIdentifier;
    SheetInfo                           myMainSheet;
    PDCListObject                       myMainTable;
    IDictionary<string, PDCListObject>  myMeasurementTables = new Dictionary<string, PDCListObject>();
    List<string>                        myShapeFileNames = new List<string>();
    Lib.Testdefinition                  myTestDefinition;

    #region constructors
    public SheetInfo()
    {
    }

    public SheetInfo(SerializationInfo info, StreamingContext context)
    {
      myIdentifier = info.GetValue(SER_IDENTIFIER,typeof(object));
      myMainSheet = (SheetInfo) info.GetValue(SER_MAINSHEET, typeof(SheetInfo));
      myMainTable = (PDCListObject) info.GetValue(SER_MAINTABLE, typeof(PDCListObject));
      myTestDefinition = (Lib.Testdefinition) info.GetValue(SER_TESTDEFINITION, typeof(Lib.Testdefinition));
      myAdditionalSheets = (List<SheetInfo>)info.GetValue(SER_ADDITIONALSHEETS, typeof(List<SheetInfo>));
      myMeasurementTables = (IDictionary<string, PDCListObject>)info.GetValue(SER_MEASUREMENTTABLES, typeof(IDictionary<string, PDCListObject>));
      myShapeFileNames = (List<string>)info.GetValue(SER_SHAPEFILENAMES, typeof(List<string>));
      myAreMeasurementsLoaded = true;
      try
      {
        myAreMeasurementsLoaded = info.GetBoolean(SER_AREMEASUREMENTSLOADED);
      }
      catch (Exception ee)
      {
        PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Deserialisation 'SheetInfo(SerializationInfo info, StreamingContext context)' " + ee.Message);
      }
    }
    #endregion

    #region methods

    #region AddMeasurementTable
    public void AddMeasurementTable(string anIdentifier, PDCListObject aTable)
    {
      if (!myMeasurementTables.ContainsKey(anIdentifier))
      {
        myMeasurementTables.Add(anIdentifier, aTable);
      }
      else
      {
        myMeasurementTables[anIdentifier] = aTable;
      }
    }
    #endregion

    #region AddAdditionalSheet
    /// <summary>
    /// Adds an additional worksheet to the test definition
    /// </summary>
    /// <param name="aSheetInfo"></param>
    public void AddAdditionalSheet(SheetInfo aSheetInfo)
    {
      if (!myAdditionalSheets.Contains(aSheetInfo))
      {
        myAdditionalSheets.Add(aSheetInfo);
      }
    }
    #endregion

    #region CheckStillExists
    /// <summary>
    /// Checks if the Excel worksheet still exists and unregisters itself if it does not.
    /// </summary>
    public bool CheckStillExists()
    {
      if (myExcelSheetRemoved)
      {
        return true;
      }
      try
      {
        object tmpParent = myExcelSheet.Parent;
        return true;
      }
#pragma warning disable 0168
      catch (Exception e)
      {
        try
        {
          PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Invalid reference to excel sheet: " + myIdentifier);
          myExcelSheet = ExcelUtils.TheUtils.SearchSheet(myIdentifier);
          object tmpUnused = myExcelSheet.Parent;
          PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Invalid reference to excel sheet " + myIdentifier + " replaced");
          return true;
        }
        catch (Exception ee)
        {
          PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "No replacement found for invalid reference to excel sheet " + myIdentifier + ". Removing PDC data for sheet");
          myExcelSheetRemoved = true;
          DeleteMeasurementSheets();
          Globals.PDCExcelAddIn.RemoveSheetInfo(this);
          return false;
        }
      }
#pragma warning restore 0168
    }
    #endregion

    #region Cleanup
    /// <summary>
    /// Clears the internal data structure and removes possibly remaining Excel.Names
    /// </summary>
    internal void Cleanup()
    {
      

      Excel.Workbook tmpWB = null;
      List<string> tmpNames = new List<string>();
      if (myMainTable != null)
      {
        tmpNames.Add(myMainTable.ListRangeName);
      }
      if (myMeasurementTables != null)
      {
        foreach (PDCListObject tmpList in myMeasurementTables.Values)
        {
          tmpNames.Add(tmpList.ListRangeName);
        }
      }
      if (tmpNames.Count > 0 && tmpWB != null)
      {
        ExcelUtils.TheUtils.DeleteNames(null, tmpWB, tmpNames.ToArray());
      }
      myMeasurementTables = null;
      myAdditionalSheets = null;
      myMainSheet = null;
      myMainTable = null;
      myTestDefinition = null;
      myShapeFileNames = null;
      myExcelSheet = null;
      myExcelSheetRemoved = true;
    }
    #endregion

    #region ClearAdditionalSheets
    /// <summary>
    /// Clears the list of additional sheets
    /// </summary>
    public void ClearAdditionalSheets()
    {
      myAdditionalSheets.Clear();
    }
    #endregion

    #region ClearMeasurementTables
    public void ClearMeasurementTables()
    {
      myMeasurementTables.Clear();
    }
    #endregion

    /// <summary>
    /// This function is to check wether the ExcelSheet is actually valid
    /// </summary>
    /// <returns></returns>
    private string GetExcelSheetName()
    {
      string name = null;
      try
      {
        name = myExcelSheet.Name;
      }
      catch (Exception e)
      {
        // sweet nothing
      }
      finally
      {
      }
      return name;
    }
    #region Delete
    /// <summary>
    /// Delete the associated worksheet and any dependant sheets/SheetInfo.
    /// </summary>
    public void Delete()
    {
      bool tmpDisplayAlerts = Globals.PDCExcelAddIn.Application.DisplayAlerts;
      try
      {
        if (GetExcelSheetName() != null)
        {
          Excel.Workbook tmpWB = (Excel.Workbook)myExcelSheet.Parent;
          if (!ExcelUtils.TheUtils.MoreThanOneVisibleSheet(tmpWB))
          {
            ExcelUtils.TheUtils.NeutralizeSheet(myExcelSheet);
          }
          else
          {
            Globals.PDCExcelAddIn.Application.DisplayAlerts = false;
            myExcelSheet.Delete();
          }
        }
      }
      finally
      {
        Globals.PDCExcelAddIn.Application.DisplayAlerts = tmpDisplayAlerts;
        DeleteMeasurementSheets();
        Globals.PDCExcelAddIn.RemoveSheetInfo(this);
      }            
    }
    #endregion

    #region DeleteMeasurementSheets
    private void DeleteMeasurementSheets()
    {
      if (myAdditionalSheets == null)
      {
        return;
      }
      foreach (SheetInfo tmpMeasInfo in AdditionalSheets)
      {
        try
        {
          tmpMeasInfo.Delete();
        }
        catch (Exception e)
        {
          PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Removing measurement sheets", e);
        }
      }
      myAdditionalSheets = new List<SheetInfo>();
    }
    #endregion

    #region FindMeasurementTable
    /// <summary>
    /// Searches for a measurement table with the specified name. Also searches on
    /// the additional worksheets if there are any.
    /// </summary>
    /// <param name="anIdentifier"></param>
    /// <returns></returns>
    public PDCListObject FindMeasurementTable(string anIdentifier)
    {
      PDCListObject tmpList = GetMeasurementTable(anIdentifier);
      if (tmpList != null || AdditionalSheets == null)
      {
        return tmpList;
      }
      foreach (SheetInfo tmpInfo in AdditionalSheets)
      {
        tmpList = tmpInfo.FindMeasurementTable(anIdentifier);
        if (tmpList != null)
        {
          return tmpList;
        }
      }
      return tmpList;
    }
    #endregion

    #region GetMeasurementTable
    public PDCListObject GetMeasurementTable(string anIdentifier)
    {
      if (anIdentifier == null || myMeasurementTables == null)
      {
        return null;
      }
      return myMeasurementTables.ContainsKey(anIdentifier) ? myMeasurementTables[anIdentifier] : null;
    }
    #endregion

    #region GetObjectData
    public void GetObjectData(SerializationInfo info, StreamingContext context)
    {
      info.AddValue(SER_IDENTIFIER, myIdentifier);
      info.AddValue(SER_MAINSHEET, myMainSheet);
      info.AddValue(SER_MAINTABLE, myMainTable);
      info.AddValue(SER_TESTDEFINITION, myTestDefinition);
      info.AddValue(SER_ADDITIONALSHEETS, myAdditionalSheets);
      info.AddValue(SER_MEASUREMENTTABLES, myMeasurementTables);
      info.AddValue(SER_SHAPEFILENAMES, myShapeFileNames);
      info.AddValue(SER_AREMEASUREMENTSLOADED, myAreMeasurementsLoaded);
    }
    #endregion

    #region InitSheet
    /// <summary>
    /// Initializes the connection to the excel sheet.
    /// </summary>
    /// <param name="aWorksheet"></param>
    internal void InitSheet(Excel.Worksheet aWorksheet, string aVersion)
    {
      myExcelSheet = aWorksheet;
      if (myMainTable != null)
      {
        myMainTable.Container = aWorksheet;
        myMainTable.SheetInfo = this;
        foreach (ListColumn tmpColumn in myMainTable.Columns)
        {
          tmpColumn.MigrateVersion(aVersion);
        }
      }
      if (myMeasurementTables != null)
      {
        foreach (PDCListObject tmpList in myMeasurementTables.Values)
        {
          tmpList.Container = aWorksheet;
          tmpList.SheetInfo = this;
        }
      }
    }
    #endregion

    #region IsSheetMissing
    /// <summary>
    ///    Checks whether the sheet is missing.
    /// </summary>
    /// <returns>
    ///    Returns true, when the sheet is missing. Otherwise false.
    /// </returns>
    public bool IsSheetMissing()
    {
      try
      {
        object parent = myExcelSheet.Parent;
        return false;
      }
      catch (Exception)
      {
        return true;
      }
    }
    #endregion

    #region RemoveAdditionalSheet
    /// <summary>
    /// Removes an additional worksheet from the test definition
    /// </summary>
    /// <param name="aSheetInfo"></param>
    public void RemoveAdditionalSheet(SheetInfo aSheetInfo)
    {
      if (myAdditionalSheets.Contains(aSheetInfo))
      {
        myAdditionalSheets.Remove(aSheetInfo);
      }
    }
    #endregion

    #region RemoveInvalidLists
    /// <summary>
    /// Removes all PDCListObjects which are not usable anymore since
    /// their excel ranges were deleted by the user
    /// </summary>
    internal void RemoveInvalidLists()
    {
      if (myMeasurementTables != null)
      {
        Dictionary<string, PDCListObject> tmpCopy = new Dictionary<string, PDCListObject>(myMeasurementTables);
        foreach (KeyValuePair<string, PDCListObject> tmpPair in tmpCopy)
        {
          if (!tmpPair.Value.ExcelTableExists())
          {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, tmpPair.Value.Name + " not found. Removing pdc info");
            myMeasurementTables.Remove(tmpPair.Key);
          }
        }
      }
    }
    #endregion

    #region RemoveMeasurementTable
    public void RemoveMeasurementTable(string anIdentifier)
    {
      if (myMeasurementTables.ContainsKey(anIdentifier))
      {
        myMeasurementTables.Remove(anIdentifier);
      }
    }
    #endregion

    #endregion

    #region properties

    #region AdditionalSheets
    /// <summary>
    /// Returns SheetInfos for any additional worksheets belonging to the test definition
    /// </summary>
    public List<SheetInfo> AdditionalSheets
    {
      get
      {
        return myAdditionalSheets;
      }
    }
    #endregion

    #region AreMeasurementsLoaded
    /// <summary>
    ///   Gets or sets whether the measurements are loaded.
    /// </summary>
    public bool AreMeasurementsLoaded
    {
      get
      {
        return myAreMeasurementsLoaded;
      }
      set
      {
        myAreMeasurementsLoaded = value;
      }
    }
    #endregion

    #region ExcelSheet
    /// <summary>
    /// Property for the associated excel sheet
    /// </summary>
    [XmlIgnore]
    public Excel.Worksheet ExcelSheet
    {
      get
      {
        //Is our reference still valid?
        if (myExcelSheet != null && !ExcelUtils.TheUtils.IsSheetReferenceValid(myExcelSheet))
        {
          PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Invalid reference to sheet " + myIdentifier + " found");
          myExcelSheet = ExcelUtils.TheUtils.SearchSheet(myIdentifier);
          if (myExcelSheet == null)
          {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "No replacement for invalid reference to sheet " + myIdentifier + " found");
          }
          else
          {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_EXCEL, "Replaced invalid reference to sheet " + myIdentifier);
          }
        }
        return myExcelSheet;
      }
      set
      {
        myExcelSheet = value;
      }
    }
    #endregion

    #region Identifier
    /// <summary>
    /// A global identifier for a pdc sheet
    /// </summary>
    public object Identifier
    {
      get
      {
        return myIdentifier;
      }
      set
      {
        myIdentifier = value;
      }
    }
    #endregion

    #region IsMainSheet
    public bool IsMainSheet
    {
      get
      {
        return (myMainSheet == null || myMainSheet == this) && MainTable != null && myTestDefinition != null ;
      }
    }
    #endregion

    #region ListsOnSheet
    /// <summary>
    /// Returns all PDCListObject which are directly placed on the associated Excel sheet
    /// </summary>
    public List<PDCListObject> ListsOnSheet
    {
      get
      {
        List<PDCListObject> tmpLists = new List<PDCListObject>();
        if (myMainSheet == null && myMainTable != null)
        {
          tmpLists.Add(myMainTable);
        }
        tmpLists.AddRange(myMeasurementTables.Values);
        return tmpLists;
      }
    }
    #endregion

    #region MainSheetInfo
    /// <summary>
    /// Returns the main sheet of a sub ordinate sheet
    /// </summary>
    public SheetInfo MainSheetInfo
    {
      get
      {
        return myMainSheet;
      }
      set
      {
        myMainSheet = value;
      }
    }
    #endregion

    #region MainTable
    /// <summary>
    /// The main PDCListObject for the test definition
    /// </summary>
    public PDCListObject MainTable
    {
      get
      {
        if (myMainSheet != null)
        {
          return myMainSheet.MainTable;
        }
        else
        {
          return myMainTable;
        }
      }
      set
      {
        myMainTable = value;
      }
    }
    #endregion

    #region MeasurementTables
    /// <summary>
    /// Returns the dictionary of Measurement tables
    /// </summary>
    public IDictionary<string, PDCListObject> MeasurementTables
    {
      get
      {
        return myMeasurementTables;
      }
    }
    #endregion

    #region ShapeFileNames
    /// <summary>
    /// Property for the names of the shapes which were produced by the CompoundLookup.
    /// The names are used to delete the old shapes before new shapes are inserted 
    /// </summary>
    public List<string> ShapeFileNames
    {
      get
      {
        return myShapeFileNames;
      }
      set
      {
        myShapeFileNames = value;
      }
    }
    #endregion

    #region TestDefinition
    /// <summary>
    /// Property for the testdefinition to which the sheet belongs
    /// </summary>
    public Lib.Testdefinition TestDefinition
    {
      get
      {
        return myTestDefinition;
      }
      set
      {
        myTestDefinition = value;
      }
    }
    #endregion

    #endregion
  }
}
