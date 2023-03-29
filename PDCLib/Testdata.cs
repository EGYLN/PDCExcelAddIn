using System;
using System.Collections.Generic;
using BBS.ST.BHC.BSP.PDC.Lib.Util;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// Container for test data for a specific test version
  /// </summary>
  [Serializable]
  public class Testdata
  {
    private Testdefinition testVersion;
    private List<TestdataSearchCriteria> searchConditions;
    private List<ExperimentData> experiments = new List<ExperimentData>();
    private object tag;

    /// <summary>
    /// Flag indicating that the test data is not known on the server
    /// </summary>
    private bool newData = true;

    /// <summary>
    /// Temporary flag indicating that some fields may have changed after upload.
    /// </summary>
    private bool changedDuringUpload;

    #region constructor
    /// <summary>
    /// Binds the testdata to the specified test definition
    /// </summary>
    /// <param name="aTestdefinition"></param>
    public Testdata(Testdefinition aTestdefinition)
    {
      testVersion = aTestdefinition;
    }

    /// <summary>
    /// Binds the test data as a search result to the specified test definition and search criteria
    /// </summary>
    /// <param name="testDefinition"></param>
    /// <param name="searchCriteria"></param>
    public Testdata(Testdefinition testDefinition, List<TestdataSearchCriteria> searchCriteria)
        : this(testDefinition)
    {
        searchConditions = searchCriteria;
    }
    #endregion

    #region methods

    #region Add
    /// <summary>
    /// Adds a line of experiment data
    /// </summary>
    /// <param name="anExperiment">The experiment data</param>
    public void Add(ExperimentData anExperiment)
    {
      experiments.Add(anExperiment);
    }
    #endregion

    #region Compare
    /// <summary>
    /// Used for sorting the experiment data by compoundno, preparation no and date result
    /// </summary>
    /// <param name="anExperiment1">first Experiment</param>
    /// <param name="anExperiment2">second Experiment</param>
    /// <returns>less than zero:anExperiment1 is sorted first,
    /// positive value: anExperiment2 is sorted first,
    /// zero: equality</returns>
    private int Compare(ExperimentData anExperiment1, ExperimentData anExperiment2)
    {
      int tmpOrder = 0;
      if (anExperiment1.CompoundNo != null && anExperiment2.CompoundNo != null)
      {
        tmpOrder = anExperiment1.CompoundNo.CompareTo(anExperiment2.CompoundNo);
      }
      if (tmpOrder == 0 && anExperiment1.PreparationNo != null && anExperiment2.PreparationNo != null)
      {
        tmpOrder = anExperiment1.PreparationNo.CompareTo(anExperiment2.PreparationNo);
      }
      if (tmpOrder == 0 && anExperiment1.DateResult.HasValue && anExperiment2.DateResult.HasValue)
      {
        tmpOrder = - anExperiment1.DateResult.Value.CompareTo(anExperiment2.DateResult.Value);
      }
      return tmpOrder;
    }
    #endregion

    #region CountNew
    /// <summary>
    /// Returns the number of "Experiments" without experimentno
    /// </summary>
    /// <returns></returns>
    public int CountNew()
    {
      int tmpCount = 0;
      if (experiments == null)
      {
        return 0;
      }
      foreach (ExperimentData tmpExperiment in experiments)
      {
        if (!(tmpExperiment is PlaceHolderExperiment) && tmpExperiment.ExperimentNo == null)
        {
          tmpCount++;
        }
      }
      return tmpCount;
    }
    #endregion

    #region CountUploaded
    /// <summary>
    /// Returns the number of "Experiments" with experimentno
    /// </summary>
    /// <returns></returns>
    public int CountUploaded()
    {
      int tmpCount = 0;
      if (experiments == null)
      {
        return 0;
      }
      foreach (ExperimentData tmpExperiment in experiments)
      {
        if (!(tmpExperiment is PlaceHolderExperiment) && tmpExperiment.ExperimentNo != null)
        {
          tmpCount++;
        }
      }
      return tmpCount;
    }
    #endregion

    #region IsEmpty
    /// <summary>
    /// Returns true if the Testdata set does not contains valid experimentdatas.
    /// </summary>
    /// <returns></returns>
    public bool IsEmpty()
    {
      if (experiments == null || experiments.Count == 0)
      {
        return true;
      }
      foreach (ExperimentData tmpExperiment in experiments)
      {
        if (!(tmpExperiment is PlaceHolderExperiment))
        {
          return false;
        }
      }
      return true;
    }
    #endregion


    #region SortExperiments
    /// <summary>
    /// Sorts the experiment by compoundno and UploadDate. Sorting is only possible, if
    /// the experiment list does not contain PlaceHolderExperiments
    /// </summary>
    public void SortExperiments()
    {
      if (experiments == null || experiments.Count == 0)
      {
        return;
      }
      foreach (ExperimentData tmpExperiment in experiments)
      {
        if (tmpExperiment is PlaceHolderExperiment)
        {
            PDCLogger.TheLogger.LogMessage(PDCLogger.LOG_NAME_LIB, "Tried to sort experiment list with place holder experiments");                
            return;
        }
      }
      experiments.Sort(Compare);
    }
    #endregion

    #endregion

    #region properties

    #region Experiments
    /// <summary>
    /// Property holding the experiments of this test
    /// </summary>
    public List<ExperimentData> Experiments
    {
      get
      {
        return experiments;
      }
      set
      {
        experiments = value;
      }
    }
    #endregion

    #region NewData
    /// <summary>
    /// Specifies if the test data is new or already known to pdc
    /// </summary>
    public bool NewData
    {
      get
      {
        return newData;
      }
      set
      {
        newData = value;
      }
    }
    #endregion

    #region NumberOfAllMeasurementValues
    public int NumberOfAllMeasurementValues
    {
      get
      {
        int retVal = 0;
        foreach (ExperimentData experiment in Experiments)
        {
          if (experiment is PlaceHolderExperiment)
          {
            continue; //No data ignore
          }
          retVal += experiment.MaxNumberOfMeasurementValues;
        }
        return retVal;
      }
    }
    #endregion

    #region SearchConditions
    /// <summary>
    /// This property holds the search criteria
    /// </summary>
    public List<TestdataSearchCriteria> SearchConditions
    {
      get
      {
        return searchConditions;
      }
      set
      {
        searchConditions = value;
      }
    }
    #endregion

    #region Tag
    /// <summary>
    /// Arbitrary object which can to associate the test data with client specifc information
    /// </summary>
    public object Tag
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

    #region TestVersion
    /// <summary>
    /// Returns the associated test definition version
    /// </summary>
    public Testdefinition TestVersion
    {
      get
      {
        return testVersion;
      }
    }
    #endregion

    #region this
    /// <summary>
    /// Accessor for experiment data
    /// </summary>
    /// <param name="i">Index of the experiment data</param>
    /// <returns></returns>
    public ExperimentData this[int i]
    {
      get
      {
        if (experiments.Count <= i)
        {
          return null;
        }
        return experiments[i];
      }
      set
      {
        if (experiments.Capacity <= i)
        {
          experiments.Capacity = (int) (i * 1.5);
        }
        experiments[i] = value;
      }
    }
    #endregion

    #region UploadChangeFlag
    /// <summary>
    /// Flag which indicates wether the test data was changed during a upload process.
    /// </summary>
    public bool UploadChangeFlag
    {
      get
      {
        return changedDuringUpload;
      }
      set
      {
        changedDuringUpload = value;
      }
    }
    #endregion

    #endregion
  }
}
