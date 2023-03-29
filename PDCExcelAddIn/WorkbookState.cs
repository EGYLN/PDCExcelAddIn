using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Container for the PDC related workbook state which is stored in a saved workbook
    /// </summary>
    [Serializable]
  [ComVisible(false)]
  public class WorkbookState
  {
    List<SheetInfo> sheetInfos;
    PicklistHandler picklistHandler;

    #region properties

    #region PicklistHandler
    public PicklistHandler PicklistHandler
    {
      get
      {
        return picklistHandler;
      }
      set
      {
        picklistHandler = value;
      }
    }
    #endregion

    #region SheetInfos
    public List<SheetInfo> SheetInfos
    {
      get
      {
        return sheetInfos;
      }
      set
      {
        sheetInfos = value;
      }
    }
    #endregion

    #endregion
  }
}
