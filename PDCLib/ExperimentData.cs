using System;
using System.Collections.Generic;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// Contains the data of one experiment (usually a row in excel). ExperimentData always belongs
  /// to certain Testdefinition version.
  /// The data of an experiment consists of general information about an experiment, the state of its data and
  /// experiment values. The structure of the experiment values is given by the test definition.
  /// </summary>
  [Serializable]
  public class ExperimentData
  {
    /// <summary>
    /// What to do with the experiment data when committing the test data
    /// </summary>
    public enum OperationState
    {
      /// <summary>
      /// No changes happened
      /// </summary>
      NONE,
      /// <summary>
      /// New experiment
      /// </summary>
      INSERT,
      /// <summary>
      /// Existing but changed experiment
      /// </summary>
      UPDATE,
      /// <summary>
      /// Experiment was deleted.
      /// </summary>
      DELETE
    }
    
    private Testdefinition myTestDefinition;
  
    private List<TestVariableValue> myExperimentValues = new List<TestVariableValue>();
    private List<TestVariableValue> myMeasurementValues = new List<TestVariableValue>();
    private List<TestVariableValue> myExperimtentValues4SMT = new List<TestVariableValue>();


    private long? myRunNo;
    private string myCompoundNo;
    private string myPreparationNo;
    private string myReference;
    private long? myUploadId;
    private long? myExperimentNo;
    private string myAlternatePrepNo;
    private string myOrigin;
    private DateTime? myDateResult;
    private DateTime? myUploadDate;
    private DateTime? myReplicatedAt;
    private DateTime? myScheduledDate;
    private string myAssayReference;
    private decimal? myResultStatus;
    private string myPersonId;
    private int? myPersonIdType;
    private string myMcNo;
    private string myReportToPIx;


    private OperationState myOperation;
    private bool myMeasurementsLoaded = true;

    #region constructor

    /// <summary>
    /// Creates an experiment data structure for a given test definition
    /// </summary>
    /// <param name="aTD"></param>
    public ExperimentData(Testdefinition aTD)
    {
      myTestDefinition = aTD;
      myReportToPIx = "Yes";
    }

    #endregion

    #region methods

    #region GetExperimentLevelVariableValues
    /// <summary>
    /// Returns the measurement level values for the specified (measurement level) variable.
    /// Returns null if there is no entry for this variable in the measurement table.
    /// </summary>
    /// <returns>All measurement values of the experiment</returns>
    public List<TestVariableValue> GetExperimentLevelVariableValues()
    {
      return myExperimtentValues4SMT;
    }
    #endregion

    #region GetExperimentValue
    /// <summary>
    /// Return the experiment value for the specifies variable id
    /// </summary>
    /// <param name="aVariableId"></param>
    /// <returns></returns>
    public TestVariableValue GetExperimentValue(int aVariableId)
    {
      List<TestVariableValue> tmpValues = GetExperimentValues();
      if (tmpValues == null)
      {
        return null;
      }
      foreach (TestVariableValue tmpValue in tmpValues)
      {
        if (tmpValue.VariableId == aVariableId)
        {
          return tmpValue;
        }
      }
      return null;
    }
    #endregion

    #region GetExperimentValues
    /// <summary>
    /// Returns the experiment level test variable values.
    /// </summary>
    /// <returns></returns>
    public List<TestVariableValue> GetExperimentValues()
    {
      return myExperimentValues;
    }
    #endregion

    #region GetMeasurementValues
    /// <summary>
    /// Returns the measurement level values for the specified (measurement level) variable.
    /// Returns null if there is no entry for this variable in the measurement table.
    /// </summary>
    /// <returns>All measurement values of the experiment</returns>
    public List<TestVariableValue> GetMeasurementValues()
    {
      return myMeasurementValues;
    }
    #endregion

    #region SetExperimtentValues4SMT
    /// <summary>
    /// Sets the list of all measurement values for the experiment
    /// </summary>
    /// <param name="aValueList">The list of measurement level values</param>
    public void SetExperimtentValues4SMT(List<TestVariableValue> aValueList)
    {
      myExperimtentValues4SMT = aValueList;
    }
    #endregion

    #region SetMeasurementValues
    /// <summary>
    /// Sets the list of all measurement values for the experiment
    /// </summary>
    /// <param name="aValueList">The list of measurement level values</param>
    public void SetMeasurementValues(List<TestVariableValue> aValueList)
    {
      myMeasurementValues = aValueList;
    }
    #endregion

    #endregion

    #region properties

    #region AlternatePrepno
    /// <summary>
    /// Property for the alternate preparation no
    /// </summary>
    public string AlternatePrepno
    {
      get
      {
        return myAlternatePrepNo;
      }
      set
      {
        myAlternatePrepNo = value;
      }
    }
    #endregion

    #region AssayReference
    /// <summary>
    /// Property for the assay reference
    /// </summary>
    public string AssayReference
    {
      get
      {
        return myAssayReference;
      }
      set
      {
        myAssayReference = value;
      }
    }
    #endregion

    #region CompoundNo
    /// <summary>
    /// The compound no of the experiment
    /// </summary>
    public string CompoundNo
    {
      get
      {
        return myCompoundNo;
      }
      set
      {
        myCompoundNo = value;
      }
    }
    #endregion

    #region DateResult
    /// <summary>
    /// Property for the result date
    /// </summary>
    public DateTime? DateResult
    {
      get
      {
        return myDateResult;
      }
      set
      {
        myDateResult = value;
      }
    }
    #endregion

    #region ExperimentNo
    /// <summary>
    /// ExperimentNo of the experiment
    /// </summary>
    public long? ExperimentNo
    {
      get
      {
        return myExperimentNo;
      }
      set
      {
        myExperimentNo = value;
      }
    }
    #endregion

    #region MaxNumberOfMeasurementValues
    public int MaxNumberOfMeasurementValues
    {
      get
      {
        int retVal = 0;
        foreach (TestVariableValue variableValue in GetMeasurementValues())
        {
          if (!variableValue.Position.HasValue)
          {
            continue;
          }
          retVal = Math.Max(variableValue.Position.Value, retVal);
        }
        return retVal;
      }
    }
    #endregion

    #region MCNo
    /// <summary>
    /// Property for the MC number
    /// </summary>
    public string MCNo
    {
      get
      {
        return myMcNo;
      }
      set
      {
        myMcNo = value;
      }
    }
    #endregion

    #region Operation
    /// <summary>
    /// Returns the operation state of the experiment data
    /// </summary>
    public OperationState Operation
    {
      get
      {
        return myOperation;
      }
      set
      {
        myOperation = value;
      }
    }
    #endregion

    #region Origin
    /// <summary>
    /// The origin of the test data
    /// </summary>
    public string Origin
    {
      get
      {
        return myOrigin;
      }
      set
      {
        myOrigin = value;
      }
    }
    #endregion

    #region PersonId
    /// <summary>
    /// Property for the person id
    /// </summary>
    public string PersonId
    {
      get
      {
        return myPersonId;
      }
      set
      {
        myPersonId = value;
      }
    }
    #endregion

    #region PersonIdType
    /// <summary>
    /// Property for the person type
    /// </summary>
    public int? PersonIdType
    {
      get
      {
        return myPersonIdType;
      }
      set
      {
        myPersonIdType = value;
      }
    }
    #endregion

    #region PreparationNo
    /// <summary>
    /// The preparation no of the experiment
    /// </summary>
    public string PreparationNo
    {
      get
      {
        return myPreparationNo;
      }
      set
      {
        myPreparationNo = value;
      }
    }
    #endregion

    #region Reference
    /// <summary>
    /// Property for the experiment reference
    /// </summary>
    public string Reference
    {
      get
      {
        return myReference;
      }
      set
      {
        myReference = value;
      }
    }
    #endregion

    #region ReplicatedAt
    /// <summary>
    /// The time stamp of the replication to pix if applicable
    /// </summary>
    public DateTime? ReplicatedAt
    {
      get
      {
        return myReplicatedAt;
      }
      set
      {
        myReplicatedAt = value;
      }
    }
    #endregion

    #region ResultStatus
    /// <summary>
    /// Property for the result status
    /// </summary>
    public decimal? ResultStatus
    {
      get
      {
        return myResultStatus;
      }
      set
      {
        myResultStatus = value;
      }
    }
    #endregion

    #region Runno
    /// <summary>
    /// Property for the runno upload parameter
    /// </summary>
    public long? Runno
    {//TODO Really needed?
      get
      {
        return myRunNo;
      }
      set
      {
        myRunNo = value;
      }
    }
    #endregion

    #region ScheduledDate
    /// <summary>
    /// Property for the scheduled date
    /// </summary>
    public DateTime? ScheduledDate
    {
      get
      {
        return myScheduledDate;
      }
      set
      {
        myScheduledDate = value;
      }
    }
    #endregion

    #region TestVersion
    /// <summary>
    /// Returns the test definition to which the data belongs
    /// </summary>
    public Testdefinition TestVersion
    {
      get
      {
        return myTestDefinition;
      }
    }
    #endregion

    #region UploadDate
    /// <summary>
    /// Property for the upload date
    /// </summary>
    public DateTime? UploadDate
    {
      get
      {
        return myUploadDate;
      }
      set
      {
        myUploadDate = value;
      }
    }
    #endregion

    #region UploadId
    /// <summary>
    /// Property for the upload id (after upload)
    /// </summary>
    public long? UploadId
    {
      get
      {
        return myUploadId;
      }
      set
      {
        myUploadId = value;
      }
    }
    #endregion
    #region MeasurementsLoaded
    public bool MeasurementsLoaded
    {
        get
        {
            return myMeasurementsLoaded;
        }
        set
        {
            myMeasurementsLoaded = value;
        }
    }
      #endregion
    #region ReportToPix
    public string ReportToPix
    {
      get {
        return myReportToPIx;
      }
      set { 
        myReportToPIx = value; 
      }
    }
    #endregion
    #endregion
  }

  /// <summary>
  /// PlaceHolderExperiments are used for skipped rows (eg hidden rows, update) when testdata is transfered to
  /// the server. This is necessary to keep the correct mapping from Excel sheet rows and 
  /// transfered experiment data.
  /// </summary>
  [Serializable]
  public class PlaceHolderExperiment : ExperimentData
  {
    /// <summary>
    /// Cons for a skipped row place holder
    /// </summary>
    /// <param name="aTD"></param>
    public PlaceHolderExperiment(Testdefinition aTD)
        : base(aTD)
    {
    }
  }
}
