using System;
using System.Collections.Generic;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// A Predefined Parameter is a variable with predefined behavior/value restrictions
  /// </summary>
  [Serializable]
  public class PredefinedParameter
  {
    int variableid;
    decimal? picklistid;
    string tablename;
    string servicename;
    string description;
    decimal? lowerLimit;
    decimal? upperLimit;
    List<object> picklistValues;
    bool dataLoaded;
    string tag;

    #region properties

    #region DataLoaded
    /// <summary>
    /// Property indicating wether the data was already loaded. 
    /// Used to prevent multiple servicecalls when no data was returned.
    /// </summary>
    public bool DataLoaded
    {
      get
      {
        return dataLoaded;
      }
      set
      {
        dataLoaded = value;
      }
    }
    #endregion

    #region Description
    /// <summary>
    /// Description of the predefined parameter
    /// </summary>
    public string Description
    {
      get
      {
        return description;
      }
      internal set
      {
        description = value;
      }
    }
    #endregion

    #region LowerLimit
    /// <summary>
    /// Lower limit if the predefined parameter has a value range
    /// </summary>
    public decimal? LowerLimit
    {
      get
      {
        return lowerLimit;
      }
      internal set
      {
        lowerLimit = value;
      }
    }
    #endregion

    #region PicklistId
    /// <summary>
    /// Identifier of an optional picklist
    /// </summary>
    public decimal? PicklistId
    {
      get
      {
        return picklistid;
      }
      internal set
      {
        picklistid = value;
      }
    }
    #endregion

    #region PicklistValues
    /// <summary>
    /// Allowed values if the predefined parameter is associated with an enumeration
    /// </summary>
    public List<object> PicklistValues
    {
      get
      {
        return picklistValues;
      }
      set
      {
        picklistValues = value;
      }
    }
    #endregion

    #region Servicename
    /// <summary>
    /// Optional name of a reference data service
    /// </summary>
    public string Servicename
    {
      get
      {
        return servicename;
      }
      internal set
      {
        servicename = value;
      }
    }
    #endregion

    #region Tablename
    /// <summary>
    /// Optional name of a reference data table
    /// </summary>
    public string Tablename
    {
      get
      {
        return tablename;
      }
      internal set
      {
        tablename = value;
      }
    }
    #endregion

    #region Tag
    /// <summary>
    /// Arbitrary property which may be used by the client
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

    #region UpperLimit
    /// <summary>
    /// Upper limit if the predefined parameter has a value range
    /// </summary>
    public decimal? UpperLimit
    {
      get
      {
        return upperLimit;
      }
      internal set
      {
        upperLimit = value;
      }
    }
    #endregion

    #region VariableId
    /// <summary>
    /// Identifier of the associated variable
    /// </summary>
    public int VariableId
    {
      get
      {
        return variableid;
      }
      internal set
      {
        variableid = value;
      }
    }
    #endregion

    #endregion
  }
}
