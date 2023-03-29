using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using WS = BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;
using Settings = BBS.ST.BHC.BSP.PDC.Lib.Properties;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    public class TestStruct
    {
        public string compoundno = string.Empty;
        public string preparationno = string.Empty;
        public string mcno = string.Empty;
        public string compoundno_msg = string.Empty;
        public string compoundno_msg_code = string.Empty;
        public string preparationno_msg = string.Empty;
        public string preparationno_msg_code = string.Empty;
        public string mcno_msg = string.Empty;
        public string mcno_msg_code = string.Empty;
        public int compoundno_id;
        public int preparationno_id;
        public int mcno_id;
        public string[] ErrInfo;
        public int result;
        public double molweight;
        public string molformula;
        public string molfile;
        public string hydrogendisplaymode;
        public byte[] molimagearray;
        public string filename;
        public string username;
        public string fileformat;

        public string msgid;
        public string msg;
        public string msgtype;
        public string msglevel;

    }

    /// <summary>
    /// Central singleton for back end operations.
    /// </summary>
    public class PDCService
    {
        /// <summary>
        /// List of all prefixes which are known by PDC
        /// </summary>
        private List<string> myPrefixes;
        /// <summary>
        /// List of all predefined parameters which are known by PDC
        /// </summary>
        private Dictionary<int, PredefinedParameter> myPredefinedParameters;

        /// <summary>
        /// Handles login/logout related stuff
        /// </summary>
        private readonly LoginService myLoginService = new LoginService();
        /// <summary>
        /// Handles upload,update,validation, delete
        /// </summary>
        private readonly UploadService myUploadService;
        /// <summary>
        /// Handles test data search
        /// </summary>
        private readonly SearchTestdataService mySearchService;
        private static readonly object LOCK = new object();

        private static PDCService myPdcService;

        #region constructor
        /// <summary>
        /// Create a service object
        /// </summary>
        private PDCService()
        {
            myUploadService = new UploadService(this);
            mySearchService = new SearchTestdataService(this);
        }
        #endregion

        #region methods

        #region CallService
        /// <summary>
        /// Calls the named service.
        /// </summary>
        /// <param name="aServiceName">Typically the name of a pl-sql procedure</param>
        /// <param name="anInput">Input specified as pdc parameters</param>
        /// <returns></returns>
        public WS.Output CallService(string aServiceName, WS.Input anInput)
        {
            PDCLogger.TheLogger.LogStarttime(PDCLogger.LOG_NAME_LIB + "_" + aServiceName, "Connnect to service");

            WS.PDCService tmpService = Connect();
            PDCLogger.TheLogger.LogStoptime(PDCLogger.LOG_NAME_LIB + "_" + aServiceName, "Connnected");

            using (tmpService)
            {
                PDCLogger.TheLogger.LogStarttime("callService", "Calling service method '" + aServiceName + "'");
                try
                {
                    return tmpService.executeNamedOperation(aServiceName, anInput);
                }
                catch (Exception e)
                {
                    PDCLogger.TheLogger.LogError("callService", "Calling service method '" + aServiceName + "'\n" + e.Message);
                    throw;
                }
                finally
                {
                    PDCLogger.TheLogger.LogStoptime("callService", "Calling service method '" + aServiceName + "'");

                }
            }
        }
        #endregion

        #region CheckVersion
        public bool CheckVersion(String version)
        {
            PDCLogger.TheLogger.LogDebugMessage("PDC", "Check version: '" + version + "'");
            WS.Input tmpInput = new WS.Input();
            WS.PDCParameter tmpParameter = new WS.PDCParameter();
            tmpParameter.idSpecified = true;
            tmpParameter.id = PDCConstants.C_ID_PDCVERSIONID;
            tmpParameter.valueChar = version;
            tmpInput.input = new[] { tmpParameter };

            WS.Output tmpOutput = CallService(PDCConstants.C_OP_CHECK_VERSION, tmpInput);
            if (tmpOutput != null && tmpOutput.message != null)
            {
                return false;
            }
            return true;
        }
        #endregion

        #region ClientConfiguration
        /// <summary>
        /// Returns the client configuration which should be used by the client
        /// </summary>
        /// <returns></returns>
        public ClientConfiguration ClientConfiguration()
        {
            return new ClientConfiguration(this);
        }
        #endregion

        #region Connect
        /// <summary>
        /// Creates a connection to the PDC Webservice. 
        /// </summary>
        /// <returns></returns>
        public WS.PDCService Connect()
        {
            WS.PDCService tmpService = new WS.PDCService();
            WebServiceClientInitializer tmpInitializer = new WebServiceClientInitializer();
            string tmpPdcServiceName = Settings.Settings.Default.PDCLib_pdcWS_PDCService;
            tmpPdcServiceName = UserConfiguration.TheConfiguration.GetProperty(UserConfiguration.PROP_WS_PDC_URL, tmpPdcServiceName);
            PDCLogger.TheLogger.LogMessage("PDC", "Connecting to webservice: " + tmpPdcServiceName);
            tmpInitializer.InitializeWebServiceClient(tmpService, tmpPdcServiceName, "sampleuser", "ENxlhin3iSX4ZmfXCAjaXnLBBDbg31CZ2K/f94gYJQZUbpxoVqOP2eW56svqZH4QjuVwZ2nAisg=");
            tmpService.Timeout = UserConfiguration.TheConfiguration.GetIntProperty(UserConfiguration.PROP_WEBSERVICE_TIMEOUT_MILLIS, tmpService.Timeout);
            return tmpService;
        }
        #endregion

        #region DeleteTestdata
        /// <summary>
        /// Delegates the delete operation to the upload service
        /// </summary>
        /// <param name="aTD">The Testdefinition for the experimentnos</param>
        /// <param name="theExperimentNos">The experimentnos which have to be deleted.</param>
        public void DeleteTestdata(Testdefinition aTD, List<decimal> theExperimentNos)
        {
            myUploadService.DeleteTestData(aTD, theExperimentNos);
        }
        #endregion

        #region ExtractMessages
        /// <summary>
        /// Extracts the pdc messages from the annotated input vector
        /// </summary>
        /// <param name="anOutput">The result from the server</param>
        /// <param name="aMessageList">The message list which is filled with messages from the output</param>
        public void ExtractMessages(WS.Output anOutput, List<PDCMessage> aMessageList)
        {
            int tmpExperimentIndex = 0;
            WS.PDCMessage tmpError;
            PDCMessage tmpMessage;

            foreach (WS.PDCParameter tmpParam in anOutput.annotatedInput)
            {
                if (tmpParam.id == PDCConstants.C_ID_SEPERATOR)
                {
                    tmpExperimentIndex++;
                    continue;
                }
                if (tmpParam.error == null)
                {
                    continue;
                }
                tmpError = tmpParam.error;
                tmpMessage = new PDCMessage();
                tmpMessage.ExperimentIndex = tmpExperimentIndex;
                tmpMessage.Message = tmpError.message;
                tmpMessage.MessageType = tmpError.messageType == null ? null : tmpError.messageType.logLevel;
                if (tmpMessage.MessageType == null && tmpError.id != null)
                {
                    tmpMessage.MessageType = PDCMessage.GetType(tmpError.id);
                    tmpMessage.MessageCode = tmpError.id;
                }
                tmpMessage.ParameterName = tmpParam.name;
                tmpMessage.ParameterNo = tmpParam.id ?? 0;
                tmpMessage.Position = tmpParam.position;
                tmpMessage.VariableName = tmpParam.variableName;
                tmpMessage.VariableNo = tmpParam.variableId;
                aMessageList.Add(tmpMessage);
            }
            if (anOutput.message != null && Int32.Parse(anOutput.message.messageType.logLevel) > PDCConstants.C_LOG_LEVEL_INFO)
            {
                tmpError = anOutput.message;
                tmpMessage = new PDCMessage();
                tmpMessage.Message = tmpError.message;
                tmpMessage.MessageType = tmpError.messageType == null ? null : tmpError.messageType.logLevel;
                aMessageList.Add(tmpMessage);
            }
        }
        #endregion

        #region ExtractTestdefinitions
        /// <summary>
        /// Extracts an overview of testdefinitions and associated user rights from the ws output
        /// </summary>
        /// <param name="anOutput">The output from the webservice</param>
        /// <returns>a List of Testdefinition overviews</returns>
        private List<Testdefinition> ExtractTestdefinitions(WS.PDCParameter[] anOutput)
        {
            List<Testdefinition> tmpResult = new List<Testdefinition>();
            if (anOutput == null || anOutput.Length == 0)
            {
                return tmpResult;
            }
            string tmpName = null;
            int? tmpVersion = null;
            int? tmpTestNo = null;
            int? tmpPosition = null;
            int? tmpVersionId = null;
            string tmpDescription = null;
            string tmpDatechange = null;
            string tmpCurrentCWID = (Credentials ?? "").ToUpper();
            string tmpCwid = null;
            string tmpPriv = null;
            string tmpAuthor = null;
            string tmpSourceSystem = null;
            Testdefinition tmpTD = null;
            Dictionary<string, string> tmpPrivileges = new Dictionary<string, string>();
            bool tmpSepIsLast = false;
            foreach (WS.PDCParameter tmpParam in anOutput)
            {
                if (tmpParam.position != tmpPosition)
                {
                    if (tmpCwid != null && tmpPriv != null && tmpCwid.ToUpper() == tmpCurrentCWID && !tmpPrivileges.ContainsKey(tmpPriv))
                    {
                        tmpPrivileges.Add(tmpPriv, "TRUE");
                    }
                    tmpPosition = tmpParam.position;
                    tmpPriv = null;
                    tmpCwid = null;
                }
                switch (tmpParam.id)
                {
                    case PDCConstants.C_ID_TEST_NAME:
                        tmpSepIsLast = false;
                        tmpName = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_SOURCESYSTEM:
                        tmpSourceSystem = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_TESTNO:
                        tmpSepIsLast = false;
                        tmpTestNo = (int?)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_VERSION:
                        tmpSepIsLast = false;
                        tmpVersion = (int)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_TESTVERSION_ID:
                        tmpSepIsLast = false;
                        tmpVersionId = (int)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_CWID:
                        tmpCwid = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_TD_AUTHOR:
                        tmpAuthor = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_PRIVILEGE_NAME:
                        tmpPriv = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_SEPERATOR:
                        if (tmpCwid != null && tmpPriv != null && tmpCwid.ToUpper() == tmpCurrentCWID)
                        {
                            tmpPrivileges.Add(tmpPriv, "TRUE");
                        }
                        if (tmpTestNo != null)
                        {
                            tmpTD = new Testdefinition(tmpVersionId, tmpName, tmpTestNo, tmpVersion);
                            tmpTD.Sourcesystem = tmpSourceSystem;
                            tmpTD.DateChange = tmpDatechange;
                            tmpTD.Privileges = tmpPrivileges;
                            tmpTD.Description = tmpDescription;
                            tmpTD.Author = tmpAuthor;
                            tmpResult.Add(tmpTD);
                            tmpTD = null;
                            tmpDatechange = null;
                            tmpAuthor = null;
                            tmpDescription = null;
                            tmpVersionId = null;
                            tmpName = null;
                            tmpTestNo = null;
                            tmpVersion = null;
                            tmpPrivileges = new Dictionary<string, string>();
                            tmpSepIsLast = true;
                        }
                        else
                        {
                            PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Got test without testno:" + tmpVersionId);
                        }
                        break;
                }
            }
            if (!tmpSepIsLast)
            {
                if (tmpTestNo != null)
                {
                    if (tmpCwid != null && tmpPriv != null && tmpCwid.ToUpper() == tmpCurrentCWID)
                    {
                        tmpPrivileges.Add(tmpPriv, "TRUE");
                    }
                    tmpTD = new Testdefinition(tmpVersionId, tmpName, tmpTestNo, tmpVersion);
                    tmpTD.Sourcesystem = tmpSourceSystem;
                    tmpTD.DateChange = tmpDatechange;
                    tmpTD.Privileges = tmpPrivileges;
                    tmpTD.Description = tmpDescription;
                    tmpTD.Author = tmpAuthor;
                    tmpResult.Add(tmpTD);
                }
                else
                {
                    PDCLogger.TheLogger.LogDebugMessage(PDCLogger.LOG_NAME_EXCEL, "Got test without testno:" + tmpVersionId);
                }
            }
            return tmpResult;
        }
        #endregion

        #region FindTestdata
        /// <summary>
        /// Searches test data for the given search criteria
        /// </summary>
        /// <param name="aTestdefinition">The testdefinition to which the requested data belongs</param>
        /// <param name="aSearchCriteria">The search conditions</param>
        /// <returns></returns>
        public Testdata FindTestdata(Testdefinition aTestdefinition, List<TestdataSearchCriteria> aSearchCriteria)
        {
            return mySearchService.FindTestdata(aTestdefinition, aSearchCriteria);
        }
        #endregion

        #region FindTestdefinitions
        /// <summary>
        /// Search for test definition by a test definition template.
        /// </summary>
        /// <param name="aTemplate">Contains the search criteria</param>
        /// <returns>All matching curated test definition version</returns>
        public List<Testdefinition> FindTestdefinitions(Testdefinition aTemplate)
        {
            List<Testdefinition> testDefinitions = new List<Testdefinition>();
            if (UserInfo.IsPDCUser)
            {
                testDefinitions.AddRange(FindTestdefinitions(aTemplate, false));
            }
            if (UserInfo.IsIcbUser)
            {
                testDefinitions.AddRange(FindTestdefinitions(aTemplate, true));
            }
            return testDefinitions;
        }

        private List<Testdefinition> FindTestdefinitions(Testdefinition aTemplate, bool icb) {
            WS.Input tmpInput = new WS.Input();
            WS.PDCParameter tmpName = new WS.PDCParameter();
            tmpName.id = PDCConstants.C_ID_TEST_NAME;
            tmpName.idSpecified = true;
            tmpName.valueChar = aTemplate.TestName;
            WS.PDCParameter tmpNo = new WS.PDCParameter();
            tmpNo.id = PDCConstants.C_ID_TESTNO;
            tmpNo.idSpecified = true;
            WS.PDCParameter tmpCwid = new WS.PDCParameter();
            tmpCwid.id = PDCConstants.C_ID_CWID;
            tmpCwid.idSpecified = true;
            tmpCwid.valueChar = GetCwid();

            if (aTemplate.TestNo.HasValue)
            {
                tmpNo.valueNum = (decimal)aTemplate.TestNo;
                tmpNo.valueNumSpecified = true;
            }

            WS.PDCParameter tmpVersion = new WS.PDCParameter();
            tmpVersion.id = PDCConstants.C_ID_VERSION;
            tmpVersion.idSpecified = true;
            if (aTemplate.Version.HasValue)
            {
                tmpVersion.valueNum = (decimal)aTemplate.Version;
                tmpVersion.valueNumSpecified = true;
            }

            if (icb)
            {
                WS.PDCParameter tmpSourceSystem = new WS.PDCParameter();
                tmpSourceSystem.id = PDCConstants.C_ID_SOURCESYSTEM;
                tmpSourceSystem.idSpecified = true;
                tmpSourceSystem.valueChar = "ICB";
                tmpInput.input = new[] { tmpName, tmpNo, tmpVersion, tmpCwid, tmpSourceSystem };
            }
            else
            {
                tmpInput.input = new[] { tmpName, tmpNo, tmpVersion, tmpCwid };
            }

            WS.PDCService tmpService = Connect();
            using (tmpService)
            {
                WS.Output tmpOutput = tmpService.findTestdefinition(tmpInput);
                return ExtractTestdefinitions(tmpOutput.output);
            }
        }
        #endregion

        #region GetCompoundInformation
        public TestStruct GetCompoundInformation(TestStruct CI)
        {
            PDCLogger.TheLogger.LogStarttime("PDCLib.GetCompoundInformation", "GetCompoundInformation - Method start");
            try
            {

                TestStruct ret = new TestStruct();
                List<WS.PDCParameter> tmpParameters = new List<WS.PDCParameter>();
                WS.PDCParameter cnoParameter = new WS.PDCParameter();
                cnoParameter.id = PDCConstants.C_ID_COMPOUNDIDENTIFIER;
                cnoParameter.idSpecified = true;
                cnoParameter.valueChar = CI.compoundno;
                tmpParameters.Add(cnoParameter);

                WS.PDCParameter pnoParameter = new WS.PDCParameter();
                pnoParameter.id = PDCConstants.C_ID_PREPARATIONNO;
                pnoParameter.idSpecified = true;
                pnoParameter.valueChar = CI.preparationno;
                tmpParameters.Add(pnoParameter);

                WS.PDCParameter mcnoParameter = new WS.PDCParameter();
                mcnoParameter.id = PDCConstants.C_ID_MCNO;
                mcnoParameter.idSpecified = true;
                mcnoParameter.valueChar = CI.mcno;
                tmpParameters.Add(mcnoParameter);


                WS.PDCParameter person = new WS.PDCParameter();
                person.id = PDCConstants.C_ID_PERSONID;
                person.idSpecified = true;
                person.valueChar = CI.username;
                tmpParameters.Add(person);


                WS.PDCParameter fileFormat = new WS.PDCParameter();
                fileFormat.id = PDCConstants.C_ID_FILE_FORMAT_BY_ID;
                fileFormat.idSpecified = true;
                fileFormat.valueChar = CI.fileformat;
                tmpParameters.Add(fileFormat);

                WS.PDCParameter hydrogen = new WS.PDCParameter();
                hydrogen.id = PDCConstants.C_ID_HYDROGEN;
                hydrogen.idSpecified = true;
                hydrogen.valueChar = CI.hydrogendisplaymode;
                tmpParameters.Add(hydrogen);



                WS.Input tmpInput = new WS.Input();
                tmpInput.input = tmpParameters.ToArray();
                string method = UserInfo.IsIcbOnlyUser
                    ? "CHECK_COMPOUND_ICB"
                    : "CHECK_COMPOUND_ALL";
                WS.Output tmpOutput = CallService(method, tmpInput);
                if (tmpOutput == null || tmpOutput.output == null)
                {
                    return ret;
                }
                foreach (WS.PDCParameter tmpOutParam in tmpOutput.output)
                {
                    if (tmpOutParam.id == PDCConstants.C_ID_COMPOUNDIDENTIFIER)
                    {
                        ret.compoundno = tmpOutParam.valueChar;
                        if (tmpOutParam.error != null)
                        {
                            ret.compoundno_msg = tmpOutParam.error.message;
                            ret.compoundno_msg_code = tmpOutParam.error.id;
                            if (tmpOutParam.error.messageType != null && tmpOutParam.error.messageType.logLevel != null)
                                ret.compoundno_id = Int32.Parse(tmpOutParam.error.messageType.logLevel);
                            else if (tmpOutParam.error.id != null)
                            {
                                ret.compoundno_id = Int32.Parse(PDCMessage.GetType(tmpOutParam.error.id));
                            }
                        }
                    }
                    if (tmpOutParam.id == PDCConstants.C_ID_PREPARATIONNO)
                    {
                        ret.preparationno = tmpOutParam.valueChar;
                        if (tmpOutParam.error != null)
                        {
                            ret.preparationno_msg = tmpOutParam.error.message;
                            ret.preparationno_msg_code = tmpOutParam.error.id;
                            if (tmpOutParam.error.messageType != null && tmpOutParam.error.messageType.logLevel != null)
                                ret.preparationno_id = Int32.Parse(tmpOutParam.error.messageType.logLevel);
                            else if (tmpOutParam.error.id != null)
                            {
                                ret.preparationno_id = Int32.Parse(PDCMessage.GetType(tmpOutParam.error.id));
                            }
                        }
                    }
                    if (tmpOutParam.id == PDCConstants.C_ID_MCNO)
                    {
                        ret.mcno = tmpOutParam.valueChar;
                        if (tmpOutParam.error != null)
                        {
                            ret.mcno_msg = tmpOutParam.error.message;
                            ret.mcno_msg_code = tmpOutParam.error.id;
                            if (tmpOutParam.error.messageType != null && tmpOutParam.error.messageType.logLevel != null)
                                ret.mcno_id = Int32.Parse(tmpOutParam.error.messageType.logLevel);
                            else if (tmpOutParam.error.id != null)
                            {
                                ret.mcno_id = Int32.Parse(PDCMessage.GetType(tmpOutParam.error.id));
                            }
                        }
                    }
                    if (tmpOutParam.id == PDCConstants.C_ID_WEIGHT && tmpOutParam.valueNum != null)
                    {
                        ret.molweight = (double) tmpOutParam.valueNum;
                    }
                    if (tmpOutParam.id == PDCConstants.C_ID_FORMULA && tmpOutParam.valueChar != null)
                    {
                        ret.molformula = tmpOutParam.valueChar;
                    }
                    if (tmpOutParam.id == PDCConstants.C_ID_STRUCTURE_DRAWING && tmpOutParam.valueBlob != null &&
                        tmpOutParam.valueBlob.Length > 0)
                    {
                        if (CI.fileformat.Equals("MOL"))
                        {
                            ret.molfile = Encoding.Default.GetString(tmpOutParam.valueBlob);
                        }
                        else
                        {
                            ret.molimagearray = tmpOutParam.valueBlob;
                        }


                    }
                }
                if (tmpOutput.message != null)
                {
                    ret.msgid = tmpOutput.message.id;
                    ret.msg = tmpOutput.message.message;
                }
                if (tmpOutput.message != null && tmpOutput.message.messageType != null)
                {
                    ret.msgtype = tmpOutput.message.messageType.name;
                    ret.msglevel = tmpOutput.message.messageType.logLevel;
                }
                return ret;
            }
            finally
            {
                PDCLogger.TheLogger.LogStoptime("PDCLib.GetCompoundInformation", "GetCompoundInformation - Method end");
            }
        }
        #endregion

        #region GetCwid
        /// <summary>
        /// Returns the cwid of the logged in user or null
        /// </summary>
        /// <returns></returns>
        private string GetCwid() => UserInfo?.Cwid;

        #endregion

        #region GetReferenceData
        /// <summary>
        /// Returns the Reference data for the specified table or service
        /// </summary>
        /// <param name="aServiceName"></param>
        /// <param name="aTableName"></param>
        /// <returns></returns>
        public List<object> GetReferenceData(string aServiceName, string aTableName)
        {
            List<object> tmpValues = new List<object>();
            WS.Input tmpInput = new WS.Input();
            WS.PDCParameter tmpParameter = new WS.PDCParameter();
            WS.Output tmpOutput = null;
            tmpParameter.idSpecified = true;
            tmpInput.input = new[] { tmpParameter };
            if (aServiceName != null && aServiceName.Trim() != "")
            {
                tmpParameter.id = PDCConstants.C_ID_PREDEF_SERVICE;
                tmpParameter.valueChar = aServiceName;
                tmpOutput = CallService(aServiceName, tmpInput);
            }
            else if (aTableName != null && aTableName.Trim() != "")
            {
                tmpParameter.id = PDCConstants.C_ID_PREDEF_TABLE;
                tmpParameter.valueChar = aTableName;
                tmpOutput = CallService(PDCConstants.C_OP_GET_REFERENCE_DATA, tmpInput);
            }
            if (tmpOutput == null || tmpOutput.output == null)
            {
                return tmpValues;
            }
            foreach (WS.PDCParameter tmpParam in tmpOutput.output)
            {
                if (tmpParam.valueNum != null)
                {
                    tmpValues.Add(tmpParam.valueNum);
                }
                if (tmpParam.valueChar != null && tmpParam.valueChar.Trim() != "")
                {
                    tmpValues.Add(tmpParam.valueChar.Trim());
                }
            }
            return tmpValues;
        }
        #endregion

        #region InitializeTestdefinition
        /// <summary>
        /// Initializes the given Testdefinition with variables, ...
        /// </summary>
        /// <param name="aTestdefinition"></param>
        /// <returns></returns>
        public Testdefinition InitializeTestdefinition(Testdefinition aTestdefinition)
        {
            WS.Input tmpInput = new WS.Input();
            WS.PDCParameter tmpVersion = new WS.PDCParameter();
            tmpVersion.id = PDCConstants.C_ID_TESTVERSION_ID;
            tmpVersion.idSpecified = true;
            tmpVersion.valueNum = (decimal)aTestdefinition.Identifier;
            tmpVersion.valueNumSpecified = true;
            WS.PDCParameter tmpUserId = new WS.PDCParameter();
            tmpUserId.id = PDCConstants.C_ID_CWID;
            tmpUserId.idSpecified = true;
            tmpUserId.valueChar = GetCwid();
            tmpInput.input = new[] { tmpVersion, tmpUserId };

            WS.PDCService tmpService = Connect();
            using (tmpService)
            {
                WS.Output tmpVariables = tmpService.executeNamedOperation("get_variables", tmpInput);
                if (tmpVariables.message != null)
                {
                    throw new ServerFailure(tmpVariables.message);
                }
                InitializeTestdefinition(aTestdefinition, tmpVariables);
            }
            return aTestdefinition;
        }

        /// <summary>
        /// Initializes the specified testdefinition with variables,...
        /// </summary>
        /// <param name="aTestdefinition">The testdefinition to initialize</param>
        /// <param name="variables">The webservice output containing the variables</param>
        private void InitializeTestdefinition(Testdefinition aTestdefinition, WS.Output variables)
        {
            List<TestVariable> tmpVariables = new List<TestVariable>();
            if (variables.output == null || variables.output.Length == 0)
            {
                return;
            }
            int? tmpTestVersionId = null;
            int? tmpPosition = variables.output[0].position;

            TestVariable tmpTestVariable = new TestVariable();
            aTestdefinition.IsCompoundNoExpLevel = true;
            aTestdefinition.IsPrepNoExpLevel = true;
            foreach (WS.PDCParameter tmpParam in variables.output)
            {
                if (tmpParam.id == PDCConstants.C_ID_COMPOUND_NO_EXP_LEVEL)
                {
                    aTestdefinition.IsCompoundNoExpLevel = PDCConverter.Converter.ToBool(tmpParam.valueChar);
                    continue;
                }
                if (tmpParam.id == PDCConstants.C_ID_PREPARATION_NO_EXP_LEVEL)
                {
                    aTestdefinition.IsPrepNoExpLevel = PDCConverter.Converter.ToBool(tmpParam.valueChar);
                    continue;
                }
                if (tmpParam.id == PDCConstants.C_ID_MCNO_NO_EXP_LEVEL)
                {
                    aTestdefinition.IsMcNoExpLevel = PDCConverter.Converter.ToBool(tmpParam.valueChar);
                    continue;
                }
                if (tmpParam.position != tmpPosition)
                {
                    if (tmpTestVersionId == aTestdefinition.Identifier)
                    {
                        tmpVariables.Add(tmpTestVariable);
                    }
                    tmpPosition = tmpParam.position;
                    tmpTestVariable = new TestVariable();
                }
                switch (tmpParam.id)
                {
                    case PDCConstants.C_ID_PICKLIST_IDENTIFIER:
                        tmpTestVariable.PicklistId = tmpParam.valueNum;
                        break;
                    case PDCConstants.C_ID_DEFAULT_VALUE:
                        if (tmpParam.valueNum.HasValue)
                        {
                            tmpTestVariable.DefaultValue = tmpParam.valueNum;
                        }
                        else
                        {
                            tmpTestVariable.DefaultValue = tmpParam.valueChar;
                        }
                        break;
                    case PDCConstants.C_ID_TESTVERSION_ID:
                        tmpTestVersionId = (int?)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_VARIABLENO:
                        tmpTestVariable.VariableNo = (int?)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_VARIABLENAME:
                        tmpTestVariable.VariableName = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_VARIABLETYPE:
                        tmpTestVariable.VariableType = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_VARIABLECLASS:
                        tmpTestVariable.VariableClass = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_UNIT:
                        tmpTestVariable.Unit = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_VARIABLE_ID:
                        tmpTestVariable.VariableId = (int)tmpParam.valueNum; break;
                    case PDCConstants.C_ID_PICKLIST_LOW_LIMIT:
                        tmpTestVariable.LowerLimit = tmpParam.valueNum; break;
                    case PDCConstants.C_ID_PICKLIST_HIGH_LIMIT:
                        tmpTestVariable.UpperLimit = tmpParam.valueNum; break;
                    case PDCConstants.C_ID_DIFFERENTIATING:
                        tmpTestVariable.IsDifferentiating = PDCConverter.Converter.ToBool(tmpParam.valueChar); break;
                    case PDCConstants.C_ID_MANDATORY:
                        tmpTestVariable.IsMandatory = PDCConverter.Converter.ToBool(tmpParam.valueChar); break;
                    case PDCConstants.C_ID_ISCORERESULT:
                        tmpTestVariable.IsCoreResult = PDCConverter.Converter.ToBool(tmpParam.valueChar); break;
                    case PDCConstants.C_ID_COMMENT:
                        tmpTestVariable.Comments = tmpParam.valueChar; break;
                    case PDCConstants.C_ID_EXPERIMENTLEVEL:
                        tmpTestVariable.IsExperimentLevelReference = PDCConverter.Converter.ToBool(tmpParam.valueChar); break;
                    case PDCConstants.C_ID_PDC_ONLY_DATA:
                        tmpTestVariable.IsPdcVariable = PDCConverter.Converter.ToBool(tmpParam.valueChar); break;
                }
            }
            if (tmpTestVersionId == aTestdefinition.Identifier)
            {
                tmpVariables.Add(tmpTestVariable);
            }
            tmpVariables.Sort(delegate(TestVariable a, TestVariable b)
                {
                    int tmpCompare = OrderByClass(a).CompareTo(OrderByClass(b));
                    if (tmpCompare != 0)
                    {
                        return tmpCompare;
                    }
                    return string.Compare(a.Label, b.Label, StringComparison.InvariantCultureIgnoreCase);
                });
            aTestdefinition.Variables = tmpVariables;
        }
        #endregion

        #region LoadPicklists
        /// <summary>
        /// Loads the specified set of picklists.
        /// </summary>
        /// <param name="aListOfPicklistIds">Contains the identifiers of the picklists to load.</param>
        /// <returns></returns>
        public Dictionary<decimal, Picklist> LoadPicklists(List<decimal> aListOfPicklistIds)
        {
            Dictionary<decimal, Picklist> tmpPicklists = new Dictionary<decimal, Picklist>();
            WS.Input tmpInput = new WS.Input();
            List<WS.PDCParameter> tmpParameters = new List<WS.PDCParameter>();
            foreach (decimal tmpPicklistId in aListOfPicklistIds)
            {
                WS.PDCParameter tmpParameter = new WS.PDCParameter();
                tmpParameter.id = PDCConstants.C_ID_PICKLIST_IDENTIFIER;
                tmpParameter.idSpecified = true;
                tmpParameter.valueNum = tmpPicklistId;
                tmpParameter.valueNumSpecified = true;
                tmpParameters.Add(tmpParameter);
            }
            tmpInput.input = tmpParameters.ToArray();
            WS.Output tmpOutput = CallService(PDCConstants.C_OP_GET_PICKLISTS, tmpInput);
            if (tmpOutput == null || tmpOutput.output == null)
            {
                return tmpPicklists;
            }
            foreach (WS.PDCParameter tmpOutParam in tmpOutput.output)
            {
                if (tmpOutParam.position == null)
                {
                    continue;
                }
                decimal tmpPicklistId = tmpOutParam.position.Value;
                Picklist tmpPicklist;
                if (tmpPicklists.ContainsKey(tmpPicklistId))
                {
                    tmpPicklist = tmpPicklists[tmpPicklistId];
                }
                else
                {
                    tmpPicklist = new Picklist(tmpPicklistId);
                    tmpPicklists.Add(tmpPicklistId, tmpPicklist);
                }
                if (tmpOutParam.valueNum != null)
                {
                    tmpPicklist.Values.Add(tmpOutParam.valueNum);
                }
                if (tmpOutParam.valueChar != null && tmpOutParam.valueChar.Trim() != "")
                {
                    tmpPicklist.Values.Add(tmpOutParam.valueChar);
                }
            }
            return tmpPicklists;
        }
        #endregion

        #region LoadPredefinedParameters
        /// <summary>
        /// Loads the information about predefined parameters from the WS.
        /// </summary>
        private void LoadPredefinedParameters()
        {
            Dictionary<int, PredefinedParameter> tmpParameters = new Dictionary<int, PredefinedParameter>();
            try
            {
                WS.Output tmpOutput = CallService(PDCConstants.C_OP_GET_PREDEFINED_PARAMETERS, null);
                if (tmpOutput == null || tmpOutput.output == null)
                {
                    myPredefinedParameters = tmpParameters;
                    return;
                }
                foreach (WS.PDCParameter tmpParam in tmpOutput.output)
                {
                    if (tmpParam.position == null)
                    {
                        continue;
                    }
                    PredefinedParameter tmpPredefined;
                    if (tmpParameters.ContainsKey(tmpParam.position.Value))
                    {
                        tmpPredefined = tmpParameters[tmpParam.position.Value];
                    }
                    else
                    {
                        tmpPredefined = new PredefinedParameter();
                        tmpPredefined.VariableId = tmpParam.position.Value;
                        tmpParameters.Add(tmpPredefined.VariableId, tmpPredefined);
                    }
                    switch (tmpParam.id)
                    {
                        case PDCConstants.C_ID_PREDEF_DESCRIPTION:
                            tmpPredefined.Description = tmpParam.valueChar;
                            break;
                        case PDCConstants.C_ID_PREDEF_SERVICE:
                            tmpPredefined.Servicename = tmpParam.valueChar;
                            break;
                        case PDCConstants.C_ID_PREDEF_TABLE:
                            tmpPredefined.Tablename = tmpParam.valueChar;
                            break;
                        case PDCConstants.C_ID_PICKLIST_HIGH_LIMIT:
                            tmpPredefined.UpperLimit = tmpParam.valueNum;
                            break;
                        case PDCConstants.C_ID_PICKLIST_LOW_LIMIT:
                            tmpPredefined.LowerLimit = tmpParam.valueNum;
                            break;
                        case PDCConstants.C_ID_PICKLIST_IDENTIFIER:
                            tmpPredefined.PicklistId = tmpParam.valueNum;
                            break;
                    }
                }
                myPredefinedParameters = tmpParameters;
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Failed to load Predefined Parameters", e);
            }
        }
        #endregion

        #region LoadPrefixes
        /// <summary>
        /// Loads the information about known prefixes from the WS.
        /// </summary>
        private void LoadPrefixes()
        {
            List<string> tmpPrefixes = new List<string>();
            try
            {
                WS.Output tmpOutput = CallService(PDCConstants.C_OP_GET_PREFIXES, null);
                if (tmpOutput == null || tmpOutput.output == null)
                {
                    return;
                }
                foreach (WS.PDCParameter tmpParam in tmpOutput.output)
                {
                    if (tmpParam.id == PDCConstants.C_ID_PREFIX && tmpParam.valueChar != null)
                    {
                        tmpPrefixes.Add(tmpParam.valueChar);
                    }
                }
                myPrefixes = tmpPrefixes;
            }
            catch (Exception e)
            {
                PDCLogger.TheLogger.LogException(PDCLogger.LOG_NAME_EXCEL, "Failed to load prefixes", e);
            }
        }
        #endregion

        #region Login
        /// <summary>
        /// Handles the login process of the PDC Client
        /// </summary>
        /// <param name="aUsername">The cwid of user</param>
        /// <param name="aPassword">The appropriate password</param>
        /// <param name="useWinLogin">If this flag is true, the current windows login is used.
        /// Username and password should be null in this case and only the AD group information is 
        /// catched in this case.</param>
        public void Login(string aUsername, string aPassword, Boolean useWinLogin)
        {
            myLoginService.Login(aUsername, aPassword, useWinLogin);
        }
        #endregion

        #region Logout
        /// <summary>
        /// Logs the user out.
        /// </summary>
        public void Logout()
        {
            myLoginService.Logout();
        }
        #endregion

        #region LogPDCMessage
        /// <summary>
        /// Logs the specified PDC Messge to the log file
        /// </summary>
        /// <param name="aMessage"></param>
        public void LogPDCMessage(WS.PDCMessage aMessage)
        {
            if (aMessage == null)
            {
                return;
            }
            string tmpType = aMessage.messageType == null ? "" : " of type " + aMessage.messageType.name + " level " + aMessage.messageType.logLevel;
            PDCLogger.TheLogger.LogError(PDCLogger.LOG_NAME_LIB, aMessage.message + tmpType);
        }
        #endregion

        #region OrderByClass
        /// <summary>
        /// Returns the sort order of the test variable given by the variable class.
        /// </summary>
        /// <param name="aVariable">a Testvariable</param>
        /// <returns>The numerical order given by the variable class</returns>
        private int OrderByClass(TestVariable aVariable)
        {
            if (aVariable.VariableClass == null)
            {
                return 0;
            }
            switch (aVariable.VariableClass)
            {
                case "P": return 1;
                case "D": return 2;
                case "C": return 3;
                case "V": return 4;
                case "R": return 5;
                case "A": return 6;
                default: return 7;
            }
        }
        #endregion

        #region PredefinedParameter
        /// <summary>
        /// Returns the List of Predefined parameters. The parameters are are lazily loaded
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, PredefinedParameter> PredefinedParameter()
        {
            if (myPredefinedParameters == null)
            {
                LoadPredefinedParameters();
            }
            return myPredefinedParameters;
        }
        #endregion

        #region Prefixes
        /// <summary>
        /// Returns a list of well known prefixes
        /// </summary>
        /// <returns></returns>
        public List<string> Prefixes()
        {
            if (myPrefixes == null)
            {
                LoadPrefixes();
            }
            return myPrefixes;
        }
        #endregion

        #region SetPrivilegesForTestdefinition
        /// <summary>
        /// Update TestDefinitions privilege list with rights for cwid
        /// </summary>
        /// <param name="aTemplate">Contains the test definition</param>
        public void SetPrivilegesForTestdefinition(Testdefinition aTemplate)
        {
            List<Testdefinition> result = FindTestdefinitions(aTemplate, UserInfo?.IsIcbOnlyUser??false);
            aTemplate.Privileges = result?.FirstOrDefault()?.Privileges;
        }
        #endregion

        #region UploadChanges

        /// <summary>
        /// Upload of updates. Delegated to the upload service
        /// </summary>
        /// <param name="aTestdata">The changed upload data from the workbook</param>
        /// <param name="experimentNosToDelete">ExperimentNos which have to be deleted</param>
        /// <returns>A list of validation/error messages</returns>
        public List<PDCMessage> UploadChanges(Testdata aTestdata, HashSet<decimal> experimentNosToDelete)
        {
            return myUploadService.UploadChanges(aTestdata, experimentNosToDelete);
        }
        #endregion

        #region Update Checks

        public List<PDCMessage> CheckDuplicateExperiments(Testdata testdata)
        {
            return myUploadService.CheckDuplicateExperiments(testdata);
        }
        public bool CheckExperimentNos(Testdata testdata)
        {
            return myUploadService.CheckExistenceExperimentNos(testdata);
        }
        #endregion
        #region UploadTestdata
        /// <summary>
        /// Uploads the test data and returns a list of validation messages. Delegated to
        /// the upload service.
        /// </summary>
        /// <param name="aTestdata">The upload data from the workbook</param>
        /// <returns>A list of validation/error messages</returns>
        public List<PDCMessage> UploadTestdata(Testdata aTestdata)
        {
            return myUploadService.UploadTestdata(aTestdata);
        }
        #endregion

        #region Autoupdate
        /// <summary>
        /// Checks for existing data for compound,preparation,exp level combinations and creates new Testdata with experimentnos for existing data
        /// </summary>
        /// <param name="aTestdata">The upload data from the workbook</param>
        /// <param name="newTestdata">The upload data from the workbook which will be filled with the experimentnos</param>
        /// <returns>A list of validation/error messages</returns>
        public List<PDCMessage> Autoupdate(Testdata aTestdata, Testdata newTestdata)
        {
            return myUploadService.Autoupdate(aTestdata, newTestdata);
        }
        #endregion


        #region ValidateTestdata
        /// <summary>
        /// Validates the testdata against the server and returns a list of validation messages.
        /// Delegated to the upload service.
        /// </summary>
        /// <param name="aTestdata">The upload data from the workbook</param>
        /// <returns>A list of validation/error messages</returns>
        public List<PDCMessage> ValidateTestdata(Testdata aTestdata)
        {
            return myUploadService.ValidateTestdata(aTestdata);
        }
        #endregion

        #endregion

        #region properties

        #region Credentials
        /// <summary>
        /// Returns the user id
        /// </summary>
        public string Credentials
        {
            get
            {
                return myLoginService.UserInfo == null ? null : myLoginService.UserInfo.Cwid;
            }
        }
        #endregion

        #region LoggedIn
        /// <summary>
        /// Returns true if the user is logged in, false otherwise
        /// </summary>
        public bool LoggedIn
        {
            get
            {
                return myLoginService.UserInfo != null;
            }
        }
        #endregion

        #region ServerURL
        /// <summary>
        /// Returns the server url for the portal
        /// </summary>
        public string ServerURL
        {
            get
            {
                string tmpServerUrl = Settings.Settings.Default.PDC_Server;
                tmpServerUrl = UserConfiguration.TheConfiguration.GetProperty(UserConfiguration.PROP_PORTAL_PDC_URL, tmpServerUrl);
                return tmpServerUrl;
            }
        }
        #endregion

        #region ThePDCService
        /// <summary>
        /// Returns the PDC Service singleton for the specified sourcename 
        /// </summary>
        /// <returns>The PDCService singleton</returns>        
        public static PDCService ThePDCService
        {
            get
            {
                lock (LOCK)
                {
                    if (myPdcService == null)
                    {
                        myPdcService = new PDCService();
                    }
                    return myPdcService;
                }
            }
        }
        #endregion

        #region UserInfo
        /// <summary>
        /// Returns information about the logged in user or null
        /// </summary>
        public UserInfo UserInfo => myLoginService.UserInfo ?? RegistryUtil.LoggedInUser;

        #endregion

        #endregion
    }
}