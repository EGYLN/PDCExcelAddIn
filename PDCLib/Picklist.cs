using System;
using System.Collections.Generic;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// A Picklists contains a set of values from which the value of a test variable can be selected.
  /// </summary>
  [Serializable]
  public class Picklist
  {
    private List<object> values;
    private string tag;
    private decimal picklistId;

    #region constructor
    /// <summary>
    /// Initializes a new instance with the specified identifier
    /// </summary>
    /// <param name="aPicklistId"></param>
    public Picklist(decimal aPicklistId)
    {
      picklistId = aPicklistId;
      values = new List<object>();
    }
    #endregion

    #region properties

    #region PicklistId
    /// <summary>
    /// Technical identifier of the picklist
    /// </summary>
    public decimal PicklistId
    {
      get
      {
        return picklistId;
      }
    }
    #endregion

    #region Tag
    /// <summary>
    /// An arbitrary tag which may be used by the client for its own purposes
    /// </summary>
    public string Tag
    {
      get
      {
        return tag;
      }
      set
      {
        tag = value;
      }
    }
    #endregion

    #region Values
    /// <summary>
    /// Property holding the enumeration values
    /// </summary>
    public List<object> Values
    {
      get
      {
        return values;
      }
      internal set
      {
        values = value;
      }
    }
    #endregion

    #endregion
  }
}
