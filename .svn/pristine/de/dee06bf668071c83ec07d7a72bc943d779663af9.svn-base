using System;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using BBS.ST.BHC.BSP.PDC.ExcelClient.Predefined;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    /// <summary>
    /// Describes a column in a PDCListObject.
    /// </summary>
    [Serializable]
  [ComVisible(false)]
  public class ListColumn
  {
    string name;
    string comment;
    bool readOnly;
    string label;
    int? oleColor;
    int? id;
    bool hidden;
    bool removed;
    bool userDefined;
    bool isHyperLink;
    
    [NonSerialized]
    CellValidator validator = new CellValidator();

    /// <summary>
    /// If special handling is needed
    /// </summary>
    PredefinedParameterHandler paramHandler;

    /// <summary>
    /// second handler for implementation phase for horizontal measurements
    /// </summary>
    PredefinedParameterHandler paramHandler2;

    /// <summary>
    /// Link to the associated test variable or null
    /// </summary>
    Lib.TestVariable testVariable;

    #region Constructors
    public ListColumn()
    {
      validator = new CellValidator();
    }

    public ListColumn(string aName):this()
    {
      name = aName;
    }

    public ListColumn(string aName, string aLabel): this(aName)
    {
      label = aLabel;
    }

    public ListColumn(string aName, string aLabel, int anId, System.Drawing.Color aColor):this(aName,aLabel)
    {
      oleColor = System.Drawing.ColorTranslator.ToOle(aColor);
      id = anId;
    }

    public ListColumn(string aName, string aLabel, int anId, System.Drawing.Color aColor, bool aReadOnlyFlag) : this(aName, aLabel, anId, aColor)
    {
      readOnly = aReadOnlyFlag;
    }

    public ListColumn(string aName, string aLabel, int anId, string aComment, System.Drawing.Color aColor) : this(aName, aLabel, anId, aColor)
    {
      comment = aComment;
    }
    #endregion

    #region Accessors

    [XmlIgnore]
    public PredefinedParameterHandler ParamHandler
    {
      get
      {
        return paramHandler;
      }
      set
      {
        paramHandler = value;
      }
    }

    [XmlIgnore]
    public PredefinedParameterHandler ParamHandler2
    {
      get
      {
        return paramHandler2;
      }
      set
      {
        paramHandler2 = value;
      }
    }

    public bool IsHyperLink
    {
      get
      {
        return isHyperLink;
      }
      set
      {
        isHyperLink = value;
      }
    }

    public bool Removed
    {
      get
      {
          return removed;
      }
      set
      {
          removed = value;
      }
    }

    public bool UserDefined
    {
      get
      {
        return userDefined;
      }
      set
      {
        userDefined = value;
      }
    }

    public bool Hidden
    {
      get
      {
        return hidden;
      }
      set
      {
        hidden = value;
      }
    }

    public string Name
    {
      get
      {
        return name;
      }
      set
      {
        name = value;
      }
    }

    public Lib.TestVariable TestVariable
    {
      get
      {
        return testVariable;
      }
      set
      {
        if (testVariable != null)
        {
          testVariable.Tag = null;
        }
        testVariable = value;
        if (testVariable != null)
        {
          testVariable.Tag = this;
        }
      }
    }

    public string Comment
    {
      get
      {
        return comment;
      }
      set
      {
        comment = value;
      }
    }

    public int? Id
    {
      get
      {
        return id;
      }
      set
      {
        id = value;
      }
    }

    public bool ReadOnly
    {
      get
      {
        return readOnly;
      }
      set
      {
        readOnly = value;
      }
    }

    public string Label
    {
      get
      {
        return label;
      }
      set
      {
        label = value;
      }
    }

    public int? OleColor
    {
      get
      {
        return oleColor;
      }
      set
      {
        oleColor = value;
      }
    }

    public bool HasSingleMeasurementTableHandler
    {
      get
      {
        return paramHandler is Predefined.SingleMeasurementTableHandler || paramHandler2 is Predefined.SingleMeasurementTableHandler;
      }
    }

    public Predefined.SingleMeasurementTableHandler SingleMeasurementTableHandler
    {
      get
      {
        if (paramHandler is Predefined.SingleMeasurementTableHandler)
        {
          return (Predefined.SingleMeasurementTableHandler)paramHandler;
        }
        if (paramHandler2 is Predefined.SingleMeasurementTableHandler)
        {
          return (Predefined.SingleMeasurementTableHandler)paramHandler2;
        }
        return null;
      }
    }

    public bool HasMultiMeasurementTableHandler
    {
      get
      {
        return paramHandler is Predefined.MultipleMeasurementTableHandler;
      }
    }

    public Predefined.MultipleMeasurementTableHandler MultiMeasurementTableHandler
    {
      get
      {
        if (paramHandler is Predefined.MultipleMeasurementTableHandler)
        {
          return (Predefined.MultipleMeasurementTableHandler)paramHandler;
        }
        return null;
      }
    }

    [XmlIgnore]
    public CellValidator Validator
    {
      get
      {
        if (validator == null)
        {
          validator = new CellValidator();
        }
        return validator;
      }
      set
      {
        validator = value;
      }
    }
    #endregion

    public string Validate(object aValue, object anOriginalValue)
    {
      if (validator == null)
      {
        return null;
      }
      return validator.Validate(this, aValue, anOriginalValue);
    }

    internal ListColumn Copy()
    {
      ListColumn tmpCopy = new ListColumn();
      tmpCopy.Name = name;
      tmpCopy.comment = comment;
      tmpCopy.readOnly = readOnly;
      tmpCopy.label = label;
      tmpCopy.oleColor = oleColor;
      tmpCopy.id = id;
      tmpCopy.hidden = hidden;
      tmpCopy.Removed = removed;
      tmpCopy.userDefined = userDefined;
      tmpCopy.isHyperLink = isHyperLink;
      tmpCopy.validator = new CellValidator();
      tmpCopy.paramHandler = paramHandler;
      tmpCopy.testVariable = testVariable;
      return tmpCopy;
    }

    internal void MigrateVersion(string anOldVersion)
    {

      Util.VersionMigrator tmpMigrator = new Util.VersionMigrator(anOldVersion);
      paramHandler = tmpMigrator.Migrate(paramHandler);
    }
  }
}
