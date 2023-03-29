using System;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// Wrapper for the value of a test variable. The value may be a string, a numeric value or
  /// binary data.
  /// </summary>
  [Serializable]
  public class TestVariableValue
  {
    private bool isNummeric;
    private string valueChar;
    private byte[] valueBlob;
    private string fileName;
    private string fileFormat;
    private string url;
    private int? position;
    private int variableId;
    private string prefix;

    // For the search
    private string valueCharUpperLimit;
    private string valueCharLowerLimit;

        #region constructors
    /// <summary>
    /// Initializes a TestVariableValue for the specified variable
    /// </summary>
    /// <param name="aVariableId">The identification of a test variable</param>
    public TestVariableValue(int aVariableId)
    {
      variableId = aVariableId;
    }
    /// <summary>
    /// Initializes a TestVariableValue for the specified variable with the given string value.
    /// </summary>
    /// <param name="aVariableId">Idenfities a test variable</param>
    /// <param name="aValue">The value of the test variable</param>
    public TestVariableValue(int aVariableId, string aValue):this(aVariableId)
    {
      valueChar = aValue;
    }
    /// <summary>
    /// Initializes a TestVariableValue for the specified variable with the given decimal value
    /// </summary>
    /// <param name="aVariableId"></param>
    /// <param name="aValue"></param>
    public TestVariableValue(int aVariableId, decimal aValue):this(aVariableId)
    {
      valueChar = aValue.ToString();
      isNummeric = true;
    }
    #endregion

    #region properties

    #region Fileformat
    /// <summary>
    /// File format (file extension) property for test variables with variable class binary
    /// </summary>
    public string Fileformat
    {
      get
      {
        return fileFormat;
      }
      set
      {
        fileFormat = value;
      }
    }
    #endregion

    #region Filename
    /// <summary>
    /// Filename property for test variables with variable class binary
    /// </summary>
    public string Filename
    {
      get
      {
        return fileName;
      }
      set
      {
        fileName = value;
      }
    }
    #endregion

    #region Position
    /// <summary>
    /// The position of a measurement value
    /// </summary>
    public int? Position
    {
      get
      {
        return position;
      }
      set
      {
        position = value;
      }
    }
    #endregion

    #region Prefix
    /// <summary>
    /// Prefix for numeric values
    /// </summary>
    public string Prefix
    {
      get
      {
        return prefix;
      }
      set
      {
        prefix = value;
      }
    }
    #endregion
    #region 
    /// <summary>
    /// Is value numeric 
    /// </summary>
    public bool IsNummeric
    {
      get
      {
        return isNummeric;
      }
      set
      {
        isNummeric = value;
      }
    }
    public String VariableType
    {
      get
      {
        return (IsNummeric) ? "N" : "C";
      }
    }

    #endregion
    #region Url
    /// <summary>
    /// Url property for test variables with variable class binary
    /// </summary>
    public string Url
    {
      get
      {
        return url;
      }
      set
      {
        url = value;
      }
    }
    #endregion

    #region ValueBlob
    /// <summary>
    /// Property for binary data
    /// </summary>
    public byte[] ValueBlob
    {
      get
      {
        return valueBlob;
      }
      set
      {
        valueBlob = value;
      }
    }
    #endregion

    #region ValueChar
    /// <summary>
    /// Property for the string value
    /// </summary>
    public string ValueChar
    {
      get
      {
        return valueChar;
      }
      set
      {
        valueChar = value;
      }
    }
    #endregion
    #region ValueCharLowerLimit
    public string ValueCharLowerLimit
    {
        get
        {
            return valueCharLowerLimit;
        }
        set
        {
            valueCharLowerLimit = value;
        }
    }
    #endregion
    #region ValueCharUpperLimit
    /// <summary>
    /// This property contains the upper limit of a search if a range is specified.
    /// The property valueChar then contains the lower limit
    /// </summary>
    public string ValueCharUpperLimit
    {
      //Currently only used for the upload date column
      get
      {
        return valueCharUpperLimit;
      }
      set
      {
        valueCharUpperLimit = value;
      }
    }
    #endregion


    #region VariableId
    /// <summary>
    /// The variable no of the associated test variable
    /// </summary>
    public int VariableId
    {
      get
      {
        return variableId;
      }
      set
      {
        variableId = value;
      }
    }
    #endregion

    #endregion
  }
}
