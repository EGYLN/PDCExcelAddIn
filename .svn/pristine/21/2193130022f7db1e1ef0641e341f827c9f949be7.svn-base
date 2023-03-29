using System.Collections.Generic;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
  /// <summary>
  /// This class is a container for the search conditions when
  /// searching for Testdata.
  /// </summary>
  public class TestdataSearchCriteria
  {
    Testdefinition myTestDefinition;
    /// <summary>
    /// Contains general(non-test) search parameters, for which an PDCConstants.C_ID_ constant exists.
    /// </summary>
    IDictionary<int, TestVariableValue> myUploadParameterTemplate = new Dictionary<int, TestVariableValue>();

    /// <summary>
    /// Contains search values for test variables (specific to a test definition).
    /// </summary>
    IDictionary<int, TestVariableValue> myTestParameterTemplate = new Dictionary<int,TestVariableValue>();

    #region methods

    #region ContainsValueFor
    /// <summary>
    /// Returns true if the search criteria contains a value for the specified 
    /// pdc parameter or variable
    /// </summary>
    /// <param name="aVariableOrPDCParameter">The id of a pdc parameter or a variable number</param>
    /// <param name="aVariable">Specifies wether the first argument is a pdc parameter id or a variable number</param>
    /// <returns></returns>
    public bool ContainsValueFor(int aVariableOrPDCParameter, bool aVariable)
    {
      IDictionary<int, TestVariableValue> tmpMap = aVariable ? myTestParameterTemplate : myUploadParameterTemplate;
      return tmpMap.ContainsKey(aVariableOrPDCParameter);
    }
    #endregion

    #endregion

    #region properties

    #region TestDefinition
    /// <summary>
    /// Identifies the test for which data is searched
    /// </summary>
    public Testdefinition TestDefinition
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

    #region this
    /// <summary>
    /// Value of the specified test variable or pdc parameter
    /// </summary>
    /// <param name="aVariableOrPDCParameter">The is of a test variable or pdc parameter</param>
    /// <param name="isVariable">Chooses between a test variable and a pdc parameter</param>
    /// <returns></returns>
    public TestVariableValue this[int aVariableOrPDCParameter, bool isVariable]
    {
      get
      {
        IDictionary<int, TestVariableValue> tmpTemplate = isVariable ? myTestParameterTemplate : myUploadParameterTemplate;
        if (tmpTemplate.ContainsKey(aVariableOrPDCParameter))
        {
          return tmpTemplate[aVariableOrPDCParameter];
        }
        return null;
      }
      set
      {
        IDictionary<int, TestVariableValue> tmpTemplate = isVariable ? myTestParameterTemplate : myUploadParameterTemplate;
        if (tmpTemplate.ContainsKey(aVariableOrPDCParameter))
        {
          tmpTemplate[aVariableOrPDCParameter] = value;
        }
        else
        {
          tmpTemplate.Add(aVariableOrPDCParameter, value);
        }
      }
    }
    #endregion

    #region UploadParameters
    /// <summary>
    /// Property for the contained upload parameters
    /// </summary>
    public IDictionary<int, TestVariableValue> UploadParameters
    {
      get
      {
        return myUploadParameterTemplate;
      }
    }
    #endregion

    #region Variables
    /// <summary>
    /// Returns the mapping from variable no to variable value for contained variables
    /// </summary>
    public IDictionary<int, TestVariableValue> Variables 
    {
      get
      {
        return myTestParameterTemplate;
      }
    }
    #endregion

    #endregion
  }
}
