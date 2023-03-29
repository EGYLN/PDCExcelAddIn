using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Xml.Serialization;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// Describes a test. Mandatory fields are nullable so that instances can be used as search templates
    /// </summary>
    [Serializable]
    public class Testdefinition : ISerializable
    {
        private string myAuthor;
        private string myDateChange;
        private string myDescription;
        private IDictionary<int, TestVariable> myExperimentVariables;
        private IDictionary<int, TestVariable> myExperimentLevelVariables = new Dictionary<int, TestVariable>();
        private int? myIdentifier;
        private IDictionary<int, TestVariable> myMeasurementVariables;
        private Dictionary<string, string> myPrivileges;
        private bool myShowSingleMeasurement;
        private string mySourcesystem;
        private string myTestName;
        private int? myTestNo;
        private int? myVersion;
        private object myTag;
        private bool myCompoundNoExpLevel;
        private bool myPrepNoExpLevel;
        private bool myMcNoExpLevel;

        #region constructor
        /// <summary>
        /// For serialization only
        /// </summary>
        public Testdefinition()
        {
            myShowSingleMeasurement = false;
        }

        /// <summary>
        /// Initializes a test definition which has a name, a number and a version
        /// </summary>
        /// <param name="anIdentifier">Technical id of a test definition</param>
        /// <param name="aName">Human readable name</param>
        /// <param name="aNo">Unique identifier for a test</param>
        /// <param name="aVersion">Specifies the version of a test</param>
        public Testdefinition(int? anIdentifier, string aName, int? aNo, int? aVersion)
        {
            myIdentifier = anIdentifier;
            myShowSingleMeasurement = false;
            myTestName = aName;
            myTestNo = aNo;
            myVersion = aVersion;
        }
        public Testdefinition(SerializationInfo info, StreamingContext context)
        {
            myAuthor = info.GetString("author");
            myDateChange = info.GetString("dateChange");
            myDescription = info.GetString("description");
            myExperimentVariables = (IDictionary<int, TestVariable>)info.GetValue("experimentVariables", typeof(IDictionary<int, TestVariable>));
            try
            {
                myExperimentLevelVariables = (IDictionary<int, TestVariable>)info.GetValue("experimentLevelVariables", typeof(IDictionary<int, TestVariable>));
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch (Exception) { }
            // ReSharper restore EmptyGeneralCatchClause
            myIdentifier = (int?)info.GetValue("identifier", typeof(int?));
            myMeasurementVariables = (IDictionary<int, TestVariable>)info.GetValue("measurementVariables", typeof(IDictionary<int, TestVariable>));
            myPrivileges = (Dictionary<string, string>)info.GetValue("privileges", typeof(Dictionary<string, string>));
            try
            {
                myShowSingleMeasurement = info.GetBoolean("showSingleMeasurement");
            }
// ReSharper disable EmptyGeneralCatchClause
            catch (Exception) { }
// ReSharper restore EmptyGeneralCatchClause

            try
            {
                myCompoundNoExpLevel = info.GetBoolean("compoundNoExpLevel");
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch (Exception) { }
            // ReSharper restore EmptyGeneralCatchClause
            try
            {
                myPrepNoExpLevel = info.GetBoolean("prepNoExpLevel");
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch (Exception) { }
            // ReSharper restore EmptyGeneralCatchClause
            try
            {
                myMcNoExpLevel = info.GetBoolean("mcNoExpLevel");
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch (Exception) { }
            // ReSharper restore EmptyGeneralCatchClause

            mySourcesystem = info.GetString("sourcesystem");
            myTestName = info.GetString("testName");
            myTestNo = (int?)info.GetValue("testNo", typeof(int?));
            myVersion = (int?)info.GetValue("version", typeof(int?));
            myTag = info.GetValue("tag", typeof(object));
        }
        void ISerializable.GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("author", myAuthor);
            info.AddValue("dateChange", myDateChange);
            info.AddValue("description", myDescription);
            info.AddValue("experimentVariables", myExperimentVariables);
            info.AddValue("experimentLevelVariables", myExperimentLevelVariables);
            info.AddValue("identifier", myIdentifier);
            info.AddValue("measurementVariables", myMeasurementVariables);
            info.AddValue("privileges", myPrivileges);
            info.AddValue("showSingleMeasurement", myShowSingleMeasurement);
            info.AddValue("compoundNoExpLevel", myCompoundNoExpLevel);
            info.AddValue("prepNoExpLevel", myPrepNoExpLevel);
            info.AddValue("sourcesystem", mySourcesystem);
            info.AddValue("testName", myTestName);
            info.AddValue("testNo", myTestNo);
            info.AddValue("version", myVersion);
            info.AddValue("tag", myTag);
            info.AddValue("mcNoExpLevel", myMcNoExpLevel);
        }
        #endregion

        #region methods

        #region HasPrivilege
        /// <summary>
        /// Returns true if the current user has data entry privileges for this test definition
        /// </summary>
        /// <returns></returns>
        public bool HasPrivilege(string cwid, string privilege)
        {
            return myPrivileges != null && myPrivileges.ContainsKey(privilege);
        }
        #endregion

        #region HasUploadPrivileges
        /// <summary>
        /// Returns true if the user may upload test data to the server 
        /// </summary>
        /// <param name="aUserInfo"></param>
        /// <returns></returns>
        public bool HasUploadPrivileges(UserInfo aUserInfo)
        {
            if (aUserInfo?.Cwid == null)
            {
                return false;
            }
            return HasPrivilege(aUserInfo.Cwid.ToUpper(), "DATAENTRY") || aUserInfo.IsAdmin() || aUserInfo.IsCurator();
        }
        #endregion

        #endregion

        #region properties

        #region Author
        /// <summary>
        /// Property for the author of the test definition
        /// </summary>
        public string Author
        {
            get => myAuthor;
            set => myAuthor = value;
        }
        #endregion

        #region DateChange
        /// <summary>
        /// Point in time of the last test definition change
        /// </summary>
        public string DateChange
        {
            get => myDateChange;
            set => myDateChange = value;
        }
        #endregion

        #region Description
        /// <summary>
        /// Property for the description of the test definition
        /// </summary>
        public string Description
        {
            get => myDescription;
            set => myDescription = value;
        }
        #endregion

        #region ExperimentVariables
        /// <summary>
        /// Property for the experiment level variables of a test definition
        /// </summary>
        [XmlIgnore]
        public IDictionary<int, TestVariable> ExperimentVariables
        {
            get => myExperimentVariables;
            set => myExperimentVariables = value;
        }
        #endregion

        #region ExperimentVariablesForSingleMeasurementtable
        /// <summary>
        /// Property holding the experiment level variables of the test definition
        /// </summary>
        [XmlIgnore]
        public IDictionary<int, TestVariable> ExperimentLevelVariables
        {
            get
            {
                if (myExperimentLevelVariables == null)
                {
                    return new Dictionary<int, TestVariable>();
                }
                return myExperimentLevelVariables;
            }
            set => myExperimentLevelVariables = value;
        }
        #endregion

        #region HasMeasurementVariables
        /// <summary>
        /// Returns true if the test definition contains measurement variables
        /// </summary>
        [XmlIgnore]
        public bool HasMeasurementVariables => myMeasurementVariables != null && myMeasurementVariables.Count > 0;

        #endregion

        #region HasExperimentLevelVariables
        /// <summary>
        /// Returns true if the test definition contains experimentlevele variables
        /// </summary>
        [XmlIgnore]
        public bool HasExperimentLevelVariables => myExperimentLevelVariables != null && myExperimentLevelVariables.Count > 0;

        #endregion

        #region Identifier
        /// <summary>
        /// Property for the technical pdc identifier of a test definition
        /// </summary>
        public int? Identifier
        {
            get => myIdentifier;
            set => myIdentifier = value;
        }
        #endregion

        #region MeasurementVariables
        /// <summary>
        /// Property holding the measurement level variables of the test definition
        /// </summary>
        [XmlIgnore]
        public IDictionary<int, TestVariable> MeasurementVariables
        {
            get => myMeasurementVariables;
            set => myMeasurementVariables = value;
        }


        #endregion

        #region Privileges
        /// <summary>
        /// Property containing the privileges of the logged-in user for this test definition
        /// </summary>
        [XmlIgnore]
        public Dictionary<string, string> Privileges
        {
            get => myPrivileges;
            internal set => myPrivileges = value;
        }
        #endregion


        #region ShowSingleMeasurement
        /// <summary>
        /// Gets or sets whether the measurements should be shown in single rows. Only when measurements are displayed.
        /// </summary>
        public bool ShowSingleMeasurement
        {
            get => myShowSingleMeasurement;
            set => myShowSingleMeasurement = value;
        }
        #endregion


        #region Sourcesystem
        /// <summary>
        /// Source system to use with the test definition
        /// </summary>
        public string Sourcesystem
        {
            get => mySourcesystem ?? "PDC";
            set => mySourcesystem = value;
        }
        #endregion

        #region Tag
        /// <summary>
        /// Arbitrary data associated with this Testdefinition
        /// </summary>
        public object Tag
        {
            get => myTag;
            set => myTag = value;
        }
        #endregion

        #region TestName
        /// <summary>
        /// Property for the (display) name of a test
        /// </summary>
        public string TestName
        {
            get => myTestName;
            set => myTestName = value;
        }
        #endregion

        #region TestNo
        /// <summary>
        /// The identification of the (published) test. This property is nullable, so 
        /// that a Testdefinition can be used a search template.
        /// </summary>
        public int? TestNo
        {
            get => myTestNo;
            set => myTestNo = value;
        }
        #endregion

        #region Variables
        /// <summary>
        /// Returns a collection of all variables in this test definition.
        /// </summary>
        public List<TestVariable> Variables
        {
            get
            {
                List<TestVariable> vars = new List<TestVariable>();
                if (myExperimentVariables != null)
                {
                    vars.AddRange(myExperimentVariables.Values);
                }
                if (myMeasurementVariables != null)
                {
                    vars.AddRange(myMeasurementVariables.Values);
                }
                return vars;
            }
            set
            {
                Dictionary<int, TestVariable> expVars = new Dictionary<int, TestVariable>();
                Dictionary<int, TestVariable> measVars = new Dictionary<int, TestVariable>();
                Dictionary<int, TestVariable> experimentLevelVariables = new Dictionary<int, TestVariable>();

                foreach (TestVariable testVar in value)
                {
                    if (testVar.IsExperimentLevel)
                    {
                        if (expVars.ContainsKey(testVar.VariableId) || measVars.ContainsKey(testVar.VariableId))
                        {
                            continue;
                            //throw new Exception("Variable " + testVar.VariableName + " ( " + testVar.VariableId + ") is defined twice in the test. Check the Test in the Portal");
                        }
                        expVars.Add(testVar.VariableId, testVar);
                        if (testVar.IsExperimentLevelReference)
                        {
                            if (experimentLevelVariables.ContainsKey(testVar.VariableId))
                            {
                                continue;
                            }
                            TestVariable testVariable4SMT = testVar.Clone();
                            testVariable4SMT.IsExperimentLevelReferenceForSMT = true;
                            experimentLevelVariables.Add(testVariable4SMT.VariableId, testVariable4SMT);
                        }
                    }
                    else
                    {
                        if (measVars.ContainsKey(testVar.VariableId) || expVars.ContainsKey(testVar.VariableId))
                        {
                            continue;
                        }
                        measVars.Add(testVar.VariableId, testVar);
                    }
                }

                myExperimentVariables = expVars;
                myMeasurementVariables = measVars;
                myExperimentLevelVariables = experimentLevelVariables;

            }
        }



        #endregion

        #region VariableMap
        /// <summary>
        /// Returns a mapping of variable nos to test variables for variables in this test definition.
        /// </summary>
        [XmlIgnore]
        public Dictionary<int, TestVariable> VariableMap
        {
            get
            {
                Dictionary<int, TestVariable> tmpVars = new Dictionary<int, TestVariable>();
                if (myExperimentVariables != null)
                {
                    foreach (KeyValuePair<int, TestVariable> tmpKV in myExperimentVariables)
                    {
                        tmpVars.Add(tmpKV.Key, tmpKV.Value);
                    }
                }
                if (myMeasurementVariables != null)
                {
                    foreach (KeyValuePair<int, TestVariable> tmpKVM in myMeasurementVariables)
                    {
                        tmpVars.Add(tmpKVM.Key, tmpKVM.Value);
                    }
                }

                return tmpVars;
            }
            set
            {
                myExperimentVariables = new Dictionary<int, TestVariable>();
                myMeasurementVariables = new Dictionary<int, TestVariable>();
                foreach (TestVariable tmpVar in value.Values)
                {
                    if (tmpVar.IsExperimentLevel)
                    {
                        myExperimentVariables.Add(tmpVar.VariableId, tmpVar);
                    }
                    else
                    {
                        myMeasurementVariables.Add(tmpVar.VariableId, tmpVar);
                    }
                }
            }
        }
        #endregion

        #region Version
        /// <summary>
        /// The version property of the test definition
        /// </summary>
        public int? Version
        {
            get => myVersion;
            set => myVersion = value;
        }
        #endregion

        #region McNoExpLevel
        /// <summary>
        /// Does PrepNo belong to ExperimentLevel (SM Table)
        /// </summary>
        public bool IsMcNoExpLevel
        {
            get => myMcNoExpLevel;
            set => myMcNoExpLevel = value;
        }
        #endregion

        #region PrepNoExpLevel
        /// <summary>
        /// Does PrepNo belong to ExperimentLevel (SM Table)
        /// </summary>
        public bool IsPrepNoExpLevel
        {
            get => myPrepNoExpLevel;
            set => myPrepNoExpLevel = value;
        }
        #endregion
        #region CompoundNoExpLevel
        /// <summary>
        /// Does CompoundNo belong to ExperimentLevel (SM Table)
        /// </summary>
        public bool IsCompoundNoExpLevel
        {
            get => myCompoundNoExpLevel;
            set => myCompoundNoExpLevel = value;
        }
        #endregion

        #endregion
    }
}
