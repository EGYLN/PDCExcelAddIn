using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using BBS.ST.BHC.BSP.PDC.Lib.Properties;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// Test variables always belong to a specific test version.
    /// </summary>
    [Serializable]
    public class TestVariable : ISerializable
    {
        /// <summary>
        /// Constant for variable class Annotation (M-Level)
        /// </summary>
        public const string VAR_CLASS_ANNOTATION = "A";
        /// <summary>
        /// Constant for variable class Binary (E-Level)
        /// </summary>
        public const string VAR_CLASS_BINARY = "B";
        /// <summary>
        /// Constant for variable class Comment (E-Level)
        /// </summary>
        public const string VAR_CLASS_COMMENT = "C";
        /// <summary>
        /// Constant for variable class Derived Result (E-Level)
        /// </summary>
        public const string VAR_CLASS_DERIVED_RESULT = "D";
        /// <summary>
        /// Constant for variable class Parameter (E-Level)
        /// </summary>
        public const string VAR_CLASS_PARAMETER = "P";
        /// <summary>
        /// Constant for variable class Result (M-Level)
        /// </summary>
        public const string VAR_CLASS_RESULT = "R";
        /// <summary>
        /// Constant for variable class Variable (M-Level)
        /// </summary>
        public const string VAR_CLASS_VARIABLE = "V";

        private int? myVariableNo;
        private string myVariableName;
        private int myVariableId;
        private string myVariableType;
        private string myVariableClass;
        private string myUnit;
        private bool myIsCoreResult;
        private string myComments;
        private object myTag;
        private List<string> myPickList;
        private bool myIsDifferentiating;
        private bool myIsMandatory;
        private bool myIsPdcVariable;
        private bool myIsExperimentLevelReference;
        private bool myIsExperimentLevelReferenceForSMT;
        private object myDefaultValue;
        private decimal? myPicklistId;
        private decimal? myLowerLimit;
        private decimal? myUpperLimit;

        #region methods
        public TestVariable()
        {
            // asweet nothing
        }

        public TestVariable(SerializationInfo info, StreamingContext context)
        {
            myVariableNo = (int?)info.GetValue("variableNo", typeof(int?));
            myVariableName = info.GetString("variableName");
            myVariableId = info.GetInt16("variableId");
            myVariableType = info.GetString("variableType");
            myVariableClass = info.GetString("variableClass");
            myUnit = info.GetString("unit");
            myIsCoreResult = info.GetBoolean("isCoreResult");
            myComments = info.GetString("comments");
            myTag = info.GetValue("tag", typeof(object));
            myPickList = (List<string>)info.GetValue("pickList", typeof(List<string>));
            myIsDifferentiating = info.GetBoolean("differentiating");
            myIsMandatory = info.GetBoolean("mandatory");
            myIsPdcVariable = info.GetBoolean("pdcVariable");
            try
            {
                myIsExperimentLevelReference = info.GetBoolean("experimentLevelReference");
            }
            catch (Exception) { };
            try
            {
                myIsExperimentLevelReferenceForSMT = info.GetBoolean("experimentLevelReferenceForSMT");
            }
            catch (Exception) { };

            myDefaultValue = info.GetValue("defaultValue", typeof(object));
            myPicklistId = (decimal?)info.GetValue("picklistId", typeof(decimal?));
            myLowerLimit = (decimal?)info.GetValue("lowerLimit", typeof(decimal?)); ;
            myUpperLimit = (decimal?)info.GetValue("upperLimit", typeof(decimal?)); ;
        }


        #region Clone
        public TestVariable Clone()
        {
            TestVariable testVariable = new TestVariable();
            testVariable.Comments = Comments;
            testVariable.DefaultValue = DefaultValue;
            testVariable.IsDifferentiating = IsDifferentiating;
            testVariable.IsExperimentLevelReference = IsExperimentLevelReference;
            testVariable.IsCoreResult = IsCoreResult;
            testVariable.LowerLimit = LowerLimit;
            testVariable.IsPdcVariable = IsPdcVariable;
            // todo
            // Deep copy has to be made (maybe):
            testVariable.myPickList = myPickList;
            testVariable.PicklistId = PicklistId;
            testVariable.Tag = Tag;
            testVariable.UpperLimit = UpperLimit;
            testVariable.VariableClass = VariableClass;
            testVariable.VariableId = VariableId;
            testVariable.VariableName = VariableName;
            testVariable.VariableNo = VariableNo;
            testVariable.VariableType = VariableType;
            return testVariable;
        }
        #endregion

        #region Differs
        /// <summary>
        /// Tries to detect changes. Returns false if no changes were detected
        /// </summary>
        /// <param name="aVar"></param>
        /// <returns></returns>
        public bool Differs(TestVariable aVar)
        {
            if (myPickList != null)
            {
                if (aVar.myPickList == null)
                {
                    return true;
                }
                if (myPickList.Count != aVar.myPickList.Count)
                {
                    return true;
                }
                foreach (string tmpS in myPickList)
                {
                    if (!aVar.myPickList.Contains(tmpS))
                    {
                        return true;
                    }
                }
            }
            return myVariableId != aVar.VariableId && myVariableNo != aVar.VariableNo && myVariableName != aVar.VariableName && myVariableType != aVar.VariableType &&
              myUnit != aVar.Unit && myIsMandatory != aVar.myIsMandatory && myIsDifferentiating != aVar.myIsDifferentiating && myIsPdcVariable != aVar.myIsPdcVariable &&
              myDefaultValue != aVar.myDefaultValue && myIsCoreResult != aVar.myIsCoreResult && myComments != aVar.Comments;
        }
        #endregion

        #region IsBinaryParameter
        /// <summary>
        /// Returns true if the test variable is a binary parameter
        /// </summary>
        /// <returns></returns>
        public bool IsBinaryParameter()
        {
            return myVariableClass == "B";
        }
        #endregion

        #region IsDerivedResult
        /// <summary>
        /// Returns true if the test variable is a derived result parameter
        /// </summary>
        /// <returns></returns>
        public bool IsDerivedResult()
        {
            return VariableClass == VAR_CLASS_DERIVED_RESULT;
        }
        #endregion

        #region IsNumeric
        /// <summary>
        /// Returns true if the test variable has a numeric value type
        /// </summary>
        /// <returns></returns>
        public bool IsNumeric()
        {
            return myVariableType == "N";
        }
        #endregion

        #endregion

        #region properties

        #region Label

        public string Label
        {
            get
            {
                string tmpLabel = VariableName;
                if (IsBinaryParameter())
                {
                    tmpLabel += "\n[" + Resources.UNIT_FILE + "]";
                }
                else if (Unit != null)
                {
                    tmpLabel += "\n[" + Unit + "]";
                }
                return tmpLabel;
            }
        }
        #endregion
        #region Comments
        /// <summary>
        /// Property for comments on variables
        /// </summary>
        public string Comments
        {
            get
            {
                return myComments;
            }
            set
            {
                myComments = value;
            }
        }
        #endregion

        #region DefaultValue
        /// <summary>
        /// Property for the default value for the test variable.
        /// The default value is either a string or a decimal
        /// </summary>
        public object DefaultValue
        {
            get
            {
                return myDefaultValue;
            }
            set
            {
                if (value == null)
                {
                    myDefaultValue = null;
                    return;
                }
                if (IsNumeric() && !(value is decimal))
                {
                    //TODO throw Exception
                }
                else
                {
                    myDefaultValue = value;
                }
            }
        }
        #endregion

        #region IsCoreResult
        /// <summary>
        /// Property of the is core result status of the variable
        /// </summary>
        public bool IsCoreResult
        {
            get
            {
                return myIsCoreResult;
            }
            set
            {
                myIsCoreResult = value;
            }
        }
        #endregion

        #region IsDifferentiating
        /// <summary>
        /// Differentiating Property
        /// </summary>
        public bool IsDifferentiating
        {
            get
            {
                return myIsDifferentiating;
            }
            set
            {
                myIsDifferentiating = value;
            }
        }
        #endregion

        #region IsExperimentLevel
        /// <summary>
        /// Returns true if the test variable is an experiment level variable
        /// </summary>
        public bool IsExperimentLevel
        {
            get
            {
                return !IsMeasurementLevel;
            }
        }
        #endregion

        #region IsExperimentLevelReference
        /// <summary>
        /// experimentLevelReference Property. This shows that the Experiment Variable is used for referencing measurements values in the
        /// single measurement table sheet
        /// </summary>
        public bool IsExperimentLevelReference
        {
            get
            {
                return myIsExperimentLevelReference;
            }
            set
            {
                myIsExperimentLevelReference = value;
            }
        }
        #endregion

        #region IsExperimentLevelReferenceForSMT
        /// <summary>
        /// IsExpVar4SMT: this variable belongs to the measurement sheet and is a clone from the referencing variable on the PDC Main sheet 
        /// with the corresponding IsExperimentLevelReference = true
        /// </summary>
        public bool IsExperimentLevelReferenceForSMT
        {
            get
            {
                return myIsExperimentLevelReferenceForSMT;
            }
            set
            {
                myIsExperimentLevelReferenceForSMT = value;
            }
        }
        #endregion

        #region IsMandatory
        /// <summary>
        /// Mandatory Property
        /// </summary>
        public bool IsMandatory
        {
            get
            {
                return myIsMandatory;
            }
            set
            {
                myIsMandatory = value;
            }
        }
        #endregion

        #region IsMeasurementLevel
        /// <summary>
        /// Returns true if the test variable is a measurement level variable
        /// </summary>
        [XmlIgnore]
        public bool IsMeasurementLevel
        {
            get
            {
                return myVariableClass == VAR_CLASS_ANNOTATION || myVariableClass == VAR_CLASS_RESULT || myVariableClass == VAR_CLASS_VARIABLE;
            }
        }
        #endregion

        #region IsPdcVariable
        /// <summary>
        /// Property specifying if the test variable is pdc private
        /// </summary>
        public bool IsPdcVariable
        {
            get
            {
                return myIsPdcVariable;
            }
            set
            {
                myIsPdcVariable = value;
            }
        }
        #endregion

        #region LowerLimit
        /// <summary>
        /// The lower numeric value of the variable
        /// </summary>
        public decimal? LowerLimit
        {
            get
            {
                return myLowerLimit;
            }
            set
            {
                myLowerLimit = value;
            }
        }
        #endregion

        #region Picklist
        /// <summary>
        /// Returns a list of possible values for the test variable. Returns null or an empty list,
        /// if any value is allowed for input.
        /// </summary>
        public List<string> Picklist
        {
            get
            {
                return myPickList;
            }
            set
            {
                myPickList = value;
            }
        }
        #endregion

        #region PicklistId
        /// <summary>
        /// Property for the technical id of an optional picklist.
        /// </summary>
        public decimal? PicklistId
        {
            get
            {
                return myPicklistId;
            }
            set
            {
                myPicklistId = value;
            }
        }
        #endregion

        #region Tag
        /// <summary>
        /// An arbitrary object associated with this TestVariable
        /// </summary>
        [XmlIgnore]
        public object Tag
        {
            get
            {
                return myTag;
            }
            set
            {
                myTag = value;
            }
        }
        #endregion

        #region Unit
        /// <summary>
        /// Property for the variable unit
        /// </summary>
        public string Unit
        {
            get
            {
                return myUnit;
            }
            set
            {
                myUnit = value;
            }
        }
        #endregion

        #region UpperLimit
        /// <summary>
        /// The upper numeric value of the variable
        /// </summary>
        public decimal? UpperLimit
        {
            get
            {
                return myUpperLimit;
            }
            set
            {
                myUpperLimit = value;
            }
        }
        #endregion

        #region VariableClass
        /// <summary>
        /// Property for the variable class
        /// </summary>
        public string VariableClass
        {
            get
            {
                return myVariableClass;
            }
            set
            {
                myVariableClass = value;
            }
        }
        #endregion

        #region VariableId
        /// <summary>
        /// Property for the technical id of a test variable
        /// </summary>
        public int VariableId
        {
            get
            {
                return myVariableId;
            }
            set
            {
                myVariableId = value;
            }
        }
        #endregion

        #region VariableName
        /// <summary>
        /// Property for the variable name
        /// </summary>
        public string VariableName
        {
            get
            {
                return myVariableName;
            }
            set
            {
                myVariableName = value;
            }
        }
        #endregion

        #region VariableNo
        /// <summary>
        /// Property for the (PIx) variable no
        /// </summary>
        public int? VariableNo
        {
            get
            {
                return myVariableNo;
            }
            set
            {
                myVariableNo = value;
            }
        }
        #endregion

        #region VariableType
        /// <summary>
        /// Property for the variable type
        /// </summary>
        public string VariableType
        {
            get
            {
                return myVariableType;
            }
            set
            {
                myVariableType = value;
            }
        }
        #endregion

        #endregion


        #region ISerializable Members

        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("variableNo", myVariableNo);
            info.AddValue("variableName", myVariableName);
            info.AddValue("variableId", myVariableId);
            info.AddValue("variableType", myVariableType);
            info.AddValue("variableClass", myVariableClass);
            info.AddValue("unit", myUnit);
            info.AddValue("isCoreResult", myIsCoreResult);
            info.AddValue("comments", myComments);
            info.AddValue("tag", myTag);
            info.AddValue("pickList", myPickList);
            info.AddValue("differentiating", myIsDifferentiating);
            info.AddValue("mandatory", myIsMandatory);
            info.AddValue("pdcVariable", myIsPdcVariable);
            info.AddValue("experimentLevelReference", myIsExperimentLevelReference);
            info.AddValue("experimentLevelReferenceForSMT", myIsExperimentLevelReferenceForSMT);
            info.AddValue("defaultValue", myDefaultValue);
            info.AddValue("picklistId", myPicklistId);
            info.AddValue("lowerLimit", myLowerLimit);
            info.AddValue("upperLimit", myUpperLimit);
        }

        #endregion
    }
}
