using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined
{
    /// <summary>
    /// A PredefinedParameterHandler specifies the PDC Client behaviour for predefined parameters.
    /// This involves the initialization of excel ranges (validation, formatting), data retrieval,
    /// validation, conversions.
    /// </summary>
    [Serializable]
  [ComVisible(false)]
  public class PredefinedParameterHandler
  {
    protected object missing = global::System.Type.Missing;

    /// <summary>
    /// Can be overriden by subclasses to perform some initialization of table cells like
    /// validation, formatting, ...
    /// </summary>
    /// <param name="aRange"></param>
    /// <param name="aPDCList"></param>
    public virtual void InitializeNewCells(Excel.Range aRange, PDCListObject aPDCList)
    {
    }

    /// <summary>
    /// Can be overriden by subclasses to perform some initialization of the sheet
    /// validation, formatting, ...
    /// </summary>
    /// <param name="aPDCList"></param>
    public virtual void InitializeSheet(PDCListObject aPDCList)
    {
    }

    /// <summary>
    /// Can be overriden by subclasses to provide additional logic when a table cell was changed. 
    /// </summary>
    /// <param name="aRange"></param>
    /// <param name="aColumn"></param>
    /// <param name="aPDCList"></param>
    public virtual void CellChanged(Excel.Range aRange, ListColumn aColumn, PDCListObject aPDCList, object originalValue)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="aRange"></param>
    /// <param name="aColumn"></param>
    /// <param name="aPDCList"></param>
    public virtual void CellChanged(Excel.Range aRange, ListColumn aColumn, PDCListObject aPDCList)
    {
    }

    /// <summary>
    /// Can be overridden by subclasses to add special handling if the associated column or list is deleted.
    /// </summary>
    /// <param name="aColumn"></param>
    /// <param name="aPDCList"></param>
    /// <param name="completeList"></param>
    public virtual void Delete(ListColumn aColumn, PDCListObject aPDCList, bool completeList)
    {
    }

    /// <summary>
    /// Can be overriden by subclasses to add special handling if the associated column needs an update.
    /// </summary>
    /// <param name="aPDCList">The list, which holds the associated list column</param>
    /// <param name="aCurrentColumn">The list column which is updated</param>
    /// <param name="anUpdateColumn">A template list column containing the updated information.
    /// (Does not replace the current list column!)
    /// </param>
    public virtual void UpdateColumn(PDCListObject aPDCList, ListColumn aCurrentColumn, ListColumn anUpdateColumn)
    {            
    }

    /// <summary>
    /// Sets the value at the specified place.
    /// </summary>
    /// <param name="pdcTable">The PDC Listobject on which the method is executed</param>
    /// <param name="theValues">The value matrix which will be set on the table</param>
    /// <param name="aRow">The current row</param>
    /// <param name="aPos">The current column position</param>
    /// <param name="aColumn">The affected ListColumn </param>
    /// <param name="anExperiment">The experiment data for the current row</param>
    /// <param name="aValue">The concrete value</param>
    public virtual void SetValue(PDCListObject pdcTable, object[,] theValues, int aRow, int aPos, ListColumn aColumn, Lib.ExperimentData anExperiment, Lib.TestVariableValue aValue)
    {
    }

    /// <summary>
    /// Can be overriden to sets the values at the specified place all at once.
    /// </summary>
    /// <param name="pdcTable">The PDC Listobject on which the method is executed</param>
    /// <param name="theValues">The value matrix which will be set on the table</param>
    /// <param name="aPos">The current column position</param>
    /// <param name="aColumn">The affected ListColumn </param>
    /// <param name="theTestdata">The complete testdata</param>
    public virtual void SetValues(PDCListObject pdcTable, object[,] theValues, int aPos, ListColumn aColumn, Lib.Testdata theTestdata)
    {
    }

    /// <summary>
    /// Called if the contents of the specified PDCListObject is cleared.
    /// </summary>
    /// <param name="pdcTable"></param>
    /// <param name="aListColumn"></param>
    /// <param name="theClearedValues"></param>
    public virtual void ClearContents(PDCListObject pdcTable, KeyValuePair<ListColumn, int> aListColumn, object[,] theClearedValues)
    {
    }

    /// <summary>
    /// Notifies the Predefined Parameter handler that its column was deleted.
    /// </summary>
    /// <param name="aPDCListObject">The list object holding the deleted column</param>
    /// <param name="aColumn">The deleted column</param>
    /// <returns>Returns true if the event processing should be stopped.</returns>
    public virtual bool ColumnDeleted(PDCListObject aPDCListObject, ListColumn aColumn)
    {
      return true;
    }
  }
}
