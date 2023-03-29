using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{

    class ExcelFilterStatus
  {
    public bool AutoFilter { get; set; }
    public List<ExcelFilter> ExcelFilters { get; set; }  

    public ExcelFilterStatus()
    {
      AutoFilter = false;
      ExcelFilters = new List<ExcelFilter>();
    }
  }
  class ExcelFilter
  {

    public int Field = 0;
    public object Criteria1 = Type.Missing;
    public object Criteria2 = Type.Missing;
    public XlAutoFilterOperator Criteria_operator = XlAutoFilterOperator.xlOr;
  }
}
