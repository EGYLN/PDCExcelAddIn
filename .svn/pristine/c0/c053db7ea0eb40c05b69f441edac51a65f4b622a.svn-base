using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
using BBS.ST.BHC.BSP.PDC.Lib.Util;
using WS = BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;

namespace BBS.ST.BHC.BSP.PDC.Lib
{
    /// <summary>
    /// Handles the upload, update, validate and delete operations
    /// </summary>
    class UploadService
    {
        /// <summary>
        /// Enumerates the kinds of UploadActions which are handled by this class
        /// </summary>
        enum UploadAction
        {
            Validate,
            Upload,
            Update
        }

        readonly PDCService myService;

        #region constructor
        public UploadService(PDCService service)
        {
            myService = service;
        }
        #endregion

        #region methods

        #region AddFileParameters
        /// <summary>
        /// Adds the webservice parameters for binary data
        /// </summary>
        /// <param name="parameters">The list of PDC Parameters</param>
        /// <param name="variable">The binary test variable</param>
        /// <param name="variableValue">The value for the test variable</param>
        private void AddFileParameters(List<WS.PDCParameter> parameters, TestVariable variable, TestVariableValue variableValue)
        {
            string filename = variableValue.Filename;
            int lastDotPos = filename.LastIndexOf('.');
            if (lastDotPos <= 0)
            {
                //Handled by GUI
                return;
            }
            if (lastDotPos == filename.Length - 1)
            {
                //Handled by GUI
                return;
            }
            string url = variableValue.Url ?? filename;
            string fileFormat = filename.Substring(lastDotPos + 1, filename.Length - (lastDotPos + 1));

            byte[] tmpContents = StreamUtil.GetContents(url);
            WS.PDCParameter parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_FILE_CONTENTS_BY_ID,
                idSpecified = true,
                valueBlob = tmpContents,
                variableId = variable.VariableId,
                variableIdSpecified = true
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_FILE_NAME_BY_ID,
                idSpecified = true,
                variableId = variable.VariableId,
                variableIdSpecified = true,
                valueChar = StreamUtil.GetShortFileName(filename)
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_FILE_FORMAT_BY_ID,
                idSpecified = true,
                variableId = variable.VariableId,
                variableIdSpecified = true,
                valueChar = fileFormat
            };
            parameters.Add(parameter);
        }
        #endregion

        #region AddSeperatorsForPlaceHolders
        /// <summary>
        /// For a correct assignment of messages to Excel cells it is necessary 
        /// to take the skipped rows into account. Seperators are therefore added
        /// for PlaceHolderExperiments to simplify message processing later on.
        /// </summary>
        /// <param name="aTestdata">The testdata from the workbook</param>
        /// <param name="theParameters">The PDC Parameter returned back in the annotated input</param>
        /// <returns>Annotated input plus seperators for PlaceHolderExperiments</returns>
        private WS.PDCParameter[] AddSeperatorsForPlaceHolders(Testdata aTestdata, WS.PDCParameter[] theParameters)
        {
            List<WS.PDCParameter> tmpParameters = new List<WS.PDCParameter>(theParameters);
            int tmpPos = 0;
            if (aTestdata.Experiments == null || aTestdata.Experiments.Count == 0 || theParameters.Length == 0)
            {
                return theParameters;
            }
            WS.PDCParameter tmpSeperator = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SEPERATOR,
                idSpecified = true
            };

            foreach (ExperimentData tmpExperiment in aTestdata.Experiments)
            {
                if (tmpPos >= tmpParameters.Count)
                {
                    break;
                }
                if (tmpExperiment is PlaceHolderExperiment)
                {
                    tmpParameters.Insert(tmpPos, tmpSeperator);
                    tmpPos++;
                }
                else
                {
                    while (tmpPos < tmpParameters.Count)
                    {
                        if (tmpParameters[tmpPos] != null && tmpParameters[tmpPos].id == PDCConstants.C_ID_SEPERATOR)
                        {
                            tmpPos++;
                            break;
                        }
                        tmpPos++;
                    }
                }
            }
            return tmpParameters.ToArray();
        }
        #endregion

        #region AddTestVariableValue
        /// <summary>
        /// Adds the specified test variable value to the PDC Parameter list.
        /// </summary>
        /// <param name="testdefinition">The testdefinition the test variable values belong to</param>
        /// <param name="parameters">The list of PDC Parameters</param>
        /// <param name="variableValue">A testvariable value. May be a string, numeric value or binary data</param>
        /// <param name="uploadAction">Binary data is only handled for upload and update.</param>
        /// <param name="isMeasurementValue">Is the specified value an measurement level value?</param>
        private void AddTestVariableValue(Testdefinition testdefinition, List<WS.PDCParameter> parameters, TestVariableValue variableValue, UploadAction uploadAction, bool isMeasurementValue)
        {
            if (variableValue == null)
            {
                return;
            }
            TestVariable variable = testdefinition.VariableMap[variableValue.VariableId];
            if (variableValue.Filename != null && !"".Equals(variableValue.Filename.Trim()))
            {
                if (uploadAction == UploadAction.Validate && !(variable.IsMandatory || variable.IsDifferentiating || variable.IsDerivedResult()))
                {
                    return;
                }
                AddFileParameters(parameters, variable, variableValue);
                return;
            }

            if (variableValue.ValueChar == null)
            {
                return;
            }
            var parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTPARAMETER_BY_ID,
                idSpecified = true,
                variableType = variable.VariableType,
                variableId = variable.VariableId,
                variableIdSpecified = true,
                unit = variable.Unit,
                valueChar = variableValue.ValueChar,
                praefix = variableValue.Prefix
            };
            parameter.variableType = variableValue.VariableType;
            if (isMeasurementValue)
            {
                parameter.position = variableValue.Position;
                parameter.positionSpecified = true;
            }
            else
            {
                parameter.position = null;
            }

            parameters.Add(parameter);
        }
        #endregion
        #region AddDeleteExperimentParameters

        /// <summary>
        /// Adds the necessary pdc parameters to the parameter list for the specified experiment data.
        /// </summary>
        /// <param name="testdefinition">Testdefinition to which the experiment data belongs</param>
        /// <param name="experimentNos"></param>
        /// <param name="parameters">The input vector for the webservice</param>
        private void AddDeleteExperimentParameters(Testdefinition testdefinition, HashSet<decimal> experimentNos, List<WS.PDCParameter> parameters)
        {
            if (experimentNos == null || !experimentNos.Any())
            {
                return;
            }
            WS.PDCParameter parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SOURCESYSTEM,
                idSpecified = true,
                valueChar = testdefinition.Sourcesystem
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTNO,
                idSpecified = true,
                valueNum = testdefinition.TestNo,
                valueNumSpecified = true
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_VERSION,
                idSpecified = true,
                valueNumSpecified = true,
                valueNum = testdefinition.Version
            };
            parameters.Add(parameter);

            foreach (decimal experimentNo in experimentNos)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_DELETED_EXPERIMENTNO,
                    idSpecified = true,
                    valueNum = experimentNo,
                    valueNumSpecified = true
                };
                parameters.Add(parameter);                
            }
            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SEPERATOR,
                idSpecified = true
            };
            parameters.Add(parameter);

        }
        #endregion

        #region AddUploadExperimentParameters
        /// <summary>
        /// Adds the necessary pdc parameters to the parameter list for the specified experiment data.
        /// </summary>
        /// <param name="testdefinition">Testdefinition to which the experiment data belongs</param>
        /// <param name="experiment">The testdata</param>
        /// <param name="parameters">The input vector for the webservice</param>
        /// <param name="uploadAction">The upload action</param>
        private void AddUploadExperimentParameters(Testdefinition testdefinition, ExperimentData experiment, List<WS.PDCParameter> parameters, UploadAction uploadAction)
        {
            WS.PDCParameter parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SOURCESYSTEM,
                idSpecified = true,
                valueChar = testdefinition.Sourcesystem
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTNO,
                idSpecified = true,
                valueNum = testdefinition.TestNo,
                valueNumSpecified = true
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_VERSION,
                idSpecified = true,
                valueNumSpecified = true,
                valueNum = testdefinition.Version
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_COMPOUNDIDENTIFIER,
                idSpecified = true,
                valueChar = experiment.CompoundNo
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_PREPARATIONNO,
                idSpecified = true,
                valueChar = experiment.PreparationNo
            };
            parameters.Add(parameter);

            if (experiment.MCNo != null)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_MCNO,
                    idSpecified = true,
                    valueChar = experiment.MCNo
                };
                parameters.Add(parameter);
            }
            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_PERSONID,
                idSpecified = true,
                valueChar = experiment.PersonId.ToUpper()
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_PDC_ONLY_DATA,
                idSpecified = true,
                valueChar = (experiment.ReportToPix.ToUpper().StartsWith("N")) ? "Y" : "N"
            };
            parameters.Add(parameter);
            if (experiment.ExperimentNo != null)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_EXPERIMENTNO,
                    idSpecified = true,
                    valueNum = experiment.ExperimentNo.Value,
                    valueNumSpecified = true
                };
                parameters.Add(parameter);
            }
            if (experiment.UploadId != null)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_UPLOAD_ID,
                    idSpecified = true,
                    valueNum = experiment.UploadId.Value,
                    valueNumSpecified = true
                };
                parameters.Add(parameter);
            }
            if (experiment.DateResult != null)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_DATE_RESULT,
                    idSpecified = true,
                    valueChar = PDCConverter.Converter.FromDate(experiment.DateResult)
                };
                parameters.Add(parameter);
            }

            foreach (TestVariableValue variableValue in experiment.GetExperimentValues())
            {
                AddTestVariableValue(testdefinition, parameters, variableValue, uploadAction, false);
            }

            if (!experiment.MeasurementsLoaded && uploadAction.Equals(UploadAction.Update))
            {
                // Unknown Parameter from db
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_KEEPMEASUREMENTS,
                    valueChar = "Y",
                    idSpecified = true
                };
                parameters.Add(parameter);

            }
            else
            {
                foreach (TestVariableValue variableValue in experiment.GetMeasurementValues())
                {
                    AddTestVariableValue(testdefinition, parameters, variableValue, uploadAction, true);
                }
            }
        }
        #endregion

        #region addAutoupdateParameters

        /// <summary>
        /// Adds the necessary pdc parameters to the parameter list for the specified experiment data.
        /// </summary>
        /// <param name="testdefinition">Testdefinition to which the experiment data belongs</param>
        /// <param name="experiment">The testdata</param>
        /// <param name="parameterList">The input vector for the webservice</param>
        /// <param name="position"></param>
        private void AddAutoupdateParameters(Testdefinition testdefinition, ExperimentData experiment, List<WS.PDCParameter> parameterList, int position)
        {
            WS.PDCParameter parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SOURCESYSTEM,
                idSpecified = true,
                valueChar = testdefinition.Sourcesystem
            };
            parameterList.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTNO,
                idSpecified = true,
                valueNum = testdefinition.TestNo,
                valueNumSpecified = true
            };
            parameterList.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_VERSION,
                idSpecified = true,
                valueNumSpecified = true,
                valueNum = testdefinition.Version
            };
            parameterList.Add(parameter);

            if (testdefinition.IsCompoundNoExpLevel)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_COMPOUNDIDENTIFIER,
                    idSpecified = true,
                    position = position,
                    positionSpecified = true,
                    valueChar = experiment.CompoundNo
                };
                parameterList.Add(parameter);
            }
            if (testdefinition.IsPrepNoExpLevel)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_PREPARATIONNO,
                    idSpecified = true,
                    valueChar = experiment.PreparationNo,
                    position = position,
                    positionSpecified = true
                };
                parameterList.Add(parameter);
            }
            if (testdefinition.IsMcNoExpLevel)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_MCNO,
                    idSpecified = true,
                    valueChar = experiment.MCNo,
                    position = position,
                    positionSpecified = true
                };
                parameterList.Add(parameter);
            }

            foreach (TestVariableValue tmpValue in experiment.GetExperimentValues())
            {
                if (testdefinition.VariableMap[tmpValue.VariableId].IsExperimentLevelReference)
                {
                    parameter = new WS.PDCParameter
                    {
                        id = PDCConstants.C_ID_TESTPARAMETER_BY_ID,
                        idSpecified = true,
                        variableId = tmpValue.VariableId,
                        variableIdSpecified = true,
                        position = position,
                        positionSpecified = true,
                        valueChar = tmpValue.ValueChar,
                        variableType = tmpValue.VariableType,
                        praefix = tmpValue.Prefix
                    };
                    parameterList.Add(parameter);
                }
            }
            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_PARTIAL_COMPARE,
                idSpecified = true,
                valueNum = 1,
                valueNumSpecified = true
            };
            parameterList.Add(parameter);

        }
        #endregion

        #region CreateTestdataInput

        /// <summary>
        /// Creates the PDCParameter input vector from the specified testdata under consideration of the upload action
        /// </summary>
        /// <param name="aTestdata">Contains the test data values from the workbook</param>
        /// <param name="anAction">Binary upload parameters are not considered for validation</param>
        /// <param name="experimentNosToDelete"></param>
        /// <returns>Array of input parameters describing the testdata</returns>
        private WS.Input CreateTestdataInput(Testdata aTestdata, UploadAction anAction, HashSet<decimal> experimentNosToDelete)
        {
            List<WS.PDCParameter> tmpParameters = new List<WS.PDCParameter>();
            bool tmpFirst = true;
            AddDeleteExperimentParameters(aTestdata.TestVersion, experimentNosToDelete, tmpParameters);
            foreach (ExperimentData tmpExperiment in aTestdata.Experiments)
            {
                if (tmpExperiment is PlaceHolderExperiment)
                {
                    continue;
                }
                if (!tmpFirst)
                {
                    WS.PDCParameter tmpParam = new WS.PDCParameter
                    {
                        id = PDCConstants.C_ID_SEPERATOR,
                        idSpecified = true
                    };
                    tmpParameters.Add(tmpParam);
                }
                tmpFirst = false;
                AddUploadExperimentParameters(aTestdata.TestVersion, tmpExperiment, tmpParameters, anAction);
            }
            WS.Input tmpInput = new WS.Input {input = tmpParameters.ToArray()};
            return tmpInput;
        }
        #endregion

        #region CreateAutoupdateInput
        private WS.Input CreateAutoupdateInput(Testdata testdata)
        {
            List<WS.PDCParameter> parameters = new List<WS.PDCParameter>();
            bool addSeperator = true;
            int i = 0;
            foreach (ExperimentData experiment in testdata.Experiments)
            {
                if (experiment is PlaceHolderExperiment)
                {
                    continue;
                }
                if (!addSeperator)
                {
                    WS.PDCParameter parameter = new WS.PDCParameter
                    {
                        id = PDCConstants.C_ID_SEPERATOR,
                        idSpecified = true
                    };
                    parameters.Add(parameter);
                }
                addSeperator = false;
                i++;
                AddAutoupdateParameters(testdata.TestVersion, experiment, parameters, i);
            }
            WS.Input input = new WS.Input {input = parameters.ToArray()};
            return input;
        }
        #endregion

        #region DeleteTestData
        /// <summary>
        /// Deletes the specified testdata given their experimentnos
        /// </summary>
        /// <param name="testdefinition">The testdefinition to which the experimentnos belong</param>
        /// <param name="theExperimentNos">The list of experimentnos specify the entries to delete</param>
        public void DeleteTestData(Testdefinition testdefinition, List<decimal> theExperimentNos)
        {
            WS.Input input = new WS.Input();
            List<WS.PDCParameter> parameters = new List<WS.PDCParameter>();

            WS.PDCParameter parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SOURCESYSTEM,
                idSpecified = true,
                valueChar = testdefinition.Sourcesystem
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTNO,
                idSpecified = true,
                valueNum = testdefinition.TestNo,
                valueNumSpecified = true
            };
            parameters.Add(parameter);

            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_VERSION,
                idSpecified = true,
                valueNum = testdefinition.Version,
                valueNumSpecified = true
            };
            parameters.Add(parameter);

            foreach (decimal experimentNo in theExperimentNos)
            {
                parameter = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_EXPERIMENTNO,
                    idSpecified = true,
                    valueNum = experimentNo,
                    valueNumSpecified = true
                };
                parameters.Add(parameter);
            }
            input.input = parameters.ToArray();
            WS.Output output = myService.CallService(PDCConstants.C_OP_DELETE_EXPERIMENTS, input);
            if (output.message != null)
            {
                throw new ServerFailure(output.message);
            }
        }
        #endregion

        #region ProcessOutput
        /// <summary>
        /// Postprocessing after validation/upload. The WS may have updated some fields.
        /// </summary>
        /// <param name="thePDCParameter"></param>
        /// <param name="aTestdata"></param>
        private void ProcessOutput(WS.PDCParameter[] thePDCParameter, Testdata aTestdata)
        {
            List<ExperimentData> tmpExperiments = aTestdata.Experiments;
            int i = 0;
            while (i < tmpExperiments.Count && tmpExperiments[i] is PlaceHolderExperiment)
            { //skip placeholders
                i++;
            }
            foreach (WS.PDCParameter tmpParam in thePDCParameter)
            {
                if (i == tmpExperiments.Count)
                {// eod
                    break;
                }
                if (tmpParam == null)
                {
                    continue;
                }
                switch (tmpParam.id)
                {
                    case PDCConstants.C_ID_UPLOADDATE:
                        tmpExperiments[i].UploadDate = PDCConverter.Converter.ToDate(tmpParam.valueChar);
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_EXPERIMENTNO:
                        tmpExperiments[i].ExperimentNo = (long?)tmpParam.valueNum;
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_UPLOAD_ID:
                        tmpExperiments[i].UploadId = (long?)tmpParam.valueNum;
                        aTestdata.UploadChangeFlag = true;
                        aTestdata.NewData = false;
                        break;
                    case PDCConstants.C_ID_PERSONID:
                        tmpExperiments[i].PersonId = tmpParam.valueChar;
                        aTestdata.UploadChangeFlag = true;
                        aTestdata.NewData = false;
                        break;
                    case PDCConstants.C_ID_PDC_ONLY_DATA:
                        tmpExperiments[i].ReportToPix = tmpParam.valueChar.Equals("Y") ? "No" : "Yes";
                        aTestdata.UploadChangeFlag = true;
                        aTestdata.NewData = false;
                        break;
                    case PDCConstants.C_ID_MCNO:
                        tmpExperiments[i].MCNo = tmpParam.valueChar;
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_PREPARATIONNO:
                        tmpExperiments[i].PreparationNo = tmpParam.valueChar;
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_COMPOUNDIDENTIFIER:
                        tmpExperiments[i].CompoundNo = tmpParam.valueChar;
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_RESULT_STATUS:
                        tmpExperiments[i].ResultStatus = tmpParam.valueNum;
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_DATE_RESULT:
                        tmpExperiments[i].DateResult = PDCConverter.Converter.ToDate(tmpParam.valueChar);
                        aTestdata.UploadChangeFlag = true;
                        break;
                    case PDCConstants.C_ID_SEPERATOR:
                        i++; //next experiment
                        while (i < tmpExperiments.Count && tmpExperiments[i] is PlaceHolderExperiment)
                        { //skip PlaceHolders
                            i++;
                        }
                        break;
                }
            }
        }
        #endregion

        #region UploadChanges

        /// <summary>
        /// Performs the update operation
        /// </summary>
        /// <param name="aTestdata">The testdata from the workbook</param>
        /// <param name="experimentNosToDelete"></param>
        /// <returns></returns>
        public List<PDCMessage> UploadChanges(Testdata aTestdata, HashSet<decimal> experimentNosToDelete)
        {
            return ValidateAndUpload(aTestdata, UploadAction.Update, experimentNosToDelete);
        }
        #endregion

        #region UploadTestdata

        /// <summary>
        /// Uploads the test data to the server and returns any failure messages
        /// </summary>
        /// <param name="aTestdata"></param>
        /// <returns></returns>
        public List<PDCMessage> UploadTestdata(Testdata aTestdata)
        {
            return ValidateAndUpload(aTestdata, UploadAction.Upload, null);
        }
        #endregion

        #region ValidateAndUpload

        /// <summary>
        /// Performs the upload, update and validate operations
        /// </summary>
        /// <param name="testdata">The testdata from the workbook</param>
        /// <param name="uploadAction">The kind of operation</param>
        /// <param name="experimentNosToDelete"></param>
        /// <returns></returns>
        private List<PDCMessage> ValidateAndUpload(Testdata testdata, UploadAction uploadAction, HashSet<decimal> experimentNosToDelete)
        {
            List<PDCMessage> messages = new List<PDCMessage>();
            if (testdata.Experiments == null || testdata.Experiments.Count == 0)
            {
                return messages;
            }
            WS.Input input = CreateTestdataInput(testdata, uploadAction, experimentNosToDelete);
            WS.PDCService pdcService = myService.Connect();
            using (pdcService)
            {
                WS.Output output;
                if (uploadAction == UploadAction.Validate)
                {//Validate
                    output = pdcService.executeNamedOperation(PDCConstants.C_OP_VALIDATE_TABLE, input);
                }
                else if (uploadAction == UploadAction.Upload)
                { //No explicit validation needed since uploadTable implicitly validates the data
                    output = pdcService.uploadTable(input);
                }
                else
                {//Update
                    output = pdcService.executeNamedOperation(PDCConstants.C_OP_UPLOAD_CHANGES, input);
                }
                if (output == null)
                {
                    return messages;
                }
                if (output.output != null && output.output.Length > 0)
                {
                    ProcessOutput(output.output, testdata);
                }
                if (output.message != null)
                {
                    Debug.WriteLine("Message:" + output.message.message);
                    if (output.annotatedInput != null)
                    {
                        ProcessOutput(output.annotatedInput, testdata);
                        output.annotatedInput = AddSeperatorsForPlaceHolders(testdata, output.annotatedInput);
                    }
                    myService.ExtractMessages(output, messages);
                }
            }

            return messages;
        }
        #endregion

        #region Autoupdate

        public bool CheckExistenceExperimentNos(Testdata testdata)
        {
            if (testdata.Experiments == null)
            {
                return true;
            }
            List<ExperimentData> experiments = testdata.Experiments.Where(e => !(e is PlaceHolderExperiment) && e.ExperimentNo != null).ToList();
            if (!experiments.Any())
            {
                return true;
            }
            if (experiments.Count > Properties.Settings.Default.Limit_ExperimentNoSearch)
            {
                
                throw new TooManyParametersException();
            }
            // ReSharper disable once PossibleInvalidOperationException
            IDictionary<long, ExperimentData> experimentMap = experiments.ToDictionary(e => e.ExperimentNo.Value);
            WS.Input input = CreateFindByExperimentNoInput(testdata.TestVersion, experiments);
            using (WS.PDCService service = myService.Connect())
            {
                WS.Output output = service.findUploadData(input);
                if (output.message != null && Int32.Parse(output.message.messageType.logLevel) > PDCConstants.C_LOG_LEVEL_INFO)
                {
                    if (output.message.id == PDCConstants.C_MSG_SEARCHPARAM_LIMIT)
                    {
                        throw new TooManyParametersException(output.message);
                    }

                    throw new ServerFailure(output.message);
                }
                if (output.output == null)
                {
                    return false;
                }
                foreach (WS.PDCParameter outParam in output.output.Where(p=>p.id == PDCConstants.C_ID_EXPERIMENTNO))
                {
                    if (outParam.valueNumSpecified && outParam.valueNum != null)
                    {
                        long eno = ((long?) outParam.valueNum).Value;
                        experimentMap.Remove(eno);
                    }
                }
            }
            return !experimentMap.Any();
        }

        private WS.Input CreateFindByExperimentNoInput(Testdefinition testdefinition, List<ExperimentData> experiments)
        {
            List<WS.PDCParameter> parameters = new List<WS.PDCParameter>();
            bool first = true;
            int i = 0;

            WS.PDCParameter param = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_SOURCESYSTEM,
                idSpecified = true,
                valueChar = testdefinition.Sourcesystem
            };
            parameters.Add(param);

            param = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_TESTNO,
                idSpecified = true,
                valueNum = testdefinition.TestNo,
                valueNumSpecified = true
            };
            parameters.Add(param);

            param = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_VERSION,
                idSpecified = true,
                valueNumSpecified = true,
                valueNum = testdefinition.Version
            };
            parameters.Add(param);

            param = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_OVERVIEW_ONLY,
                idSpecified = true,
                valueChar = "Y"
            };
            parameters.Add(param);

            foreach (ExperimentData experiment in experiments)
            {
                if (experiment is PlaceHolderExperiment)
                {
                    continue;
                }
                if (!first)
                {
                    param = new WS.PDCParameter
                    {
                        id = PDCConstants.C_ID_SEPERATOR,
                        idSpecified = true
                    };
                    parameters.Add(param);
                }
                first = false;
                i++;


                param = new WS.PDCParameter
                {
                    id = PDCConstants.C_ID_EXPERIMENTNO,
                    idSpecified = true,
                    valueNumSpecified = true,
                    valueNum = experiment.ExperimentNo,
                    position = i,
                    positionSpecified = true
                };
                parameters.Add(param);
            }
            WS.Input tmpInput = new WS.Input {input = parameters.ToArray()};
            return tmpInput;
        }

        public List<PDCMessage> CheckDuplicateExperiments(Testdata testdata)
        {
            if (testdata.Experiments == null || testdata.Experiments.Count == 0)
            {
                return null;
            }
            List<PDCMessage> messages = new List<PDCMessage>();
            WS.Input input = CreateAutoupdateInput(testdata);
            WS.PDCService pdcService = myService.Connect();
            using (pdcService)
            {
                WS.Output output = pdcService.findUploadData(input);
                if (output == null)
                {
                    return null;
                }
                HashSet<long> experimentNos = new HashSet<long>();
                foreach (ExperimentData experiment in testdata.Experiments)
                {
                    if (!(experiment is PlaceHolderExperiment) && experiment.ExperimentNo != null)
                    {
                        experimentNos.Add(experiment.ExperimentNo.Value);
                    }
                }
                if (
                    output.output.Any(param =>
                            param.id == PDCConstants.C_ID_EXPERIMENTNO && 
                            param.valueNum != null && !experimentNos.Contains((long) param.valueNum.Value)))
                {
                    messages = FillExperimentNos(testdata, output);
                }
            }
            return messages;

        }
        /// <summary>
        /// Checks for existing data for compound,preparation,exp level parameters and fills in experimentnos in newTestdata
        /// </summary>
        /// <param name="testdata">The testdata from the workbook</param>
        /// <param name="newTestdata">The testdata from the workbook, experimentnos are filled in</param>
        /// <returns>List of messages</returns>
        public List<PDCMessage> Autoupdate(Testdata testdata, Testdata newTestdata)
        {
            List<PDCMessage> messages = new List<PDCMessage>();
            if (testdata.Experiments == null || testdata.Experiments.Count == 0)
            {
                return messages;
            }
            WS.Input input = CreateAutoupdateInput(testdata);
            WS.PDCService pdcService = myService.Connect();
            using (pdcService)
            {
                WS.Output output = pdcService.findUploadData(input);

                if (output != null)
                {
                    messages = FillExperimentNos(newTestdata, output);
                }
                else
                {
                    // new data, nothing to do
                    return null;
                }
            }
            return messages;
        }
        #endregion

        #region fillExperimentNos
        private void SetIdentifierFromTestdata(Testdata newTestdata, ExperimentData experiment, List<PDCMessage> messages)
        {
            // find experiment in newTestdata and set experimentno
            foreach (ExperimentData newExperiment in newTestdata.Experiments)
            {
                if (newExperiment is PlaceHolderExperiment)
                {
                    continue;
                }
                Boolean experimentAlreadyFound = true;
                PDCMessage pdcMessage = new PDCMessage();
                String logMessage = "";
                if (newTestdata.TestVersion.IsCompoundNoExpLevel)
                {
                    experimentAlreadyFound = (newExperiment.CompoundNo == null && (experiment.CompoundNo == null || experiment.CompoundNo.Equals("BAY MISSING"))) ||
                                                (newExperiment.CompoundNo != null && experiment.CompoundNo != null &&
                                                 newExperiment.CompoundNo.ToUpper().Trim().Equals(experiment.CompoundNo));
                    logMessage += ",CompoundNo [" + (newExperiment.CompoundNo ?? "NULL") + "]";
                    pdcMessage.Message += ",CompoundNo";
                }

                if (newTestdata.TestVersion.IsPrepNoExpLevel)
                {
                    if (experimentAlreadyFound)
                    {
                        experimentAlreadyFound = (newExperiment.PreparationNo == null && experiment.PreparationNo == null) ||
                                                    (newExperiment.PreparationNo != null && experiment.PreparationNo != null &&
                                                     newExperiment.PreparationNo.ToUpper().Trim().Equals(experiment.PreparationNo));
                    }
                    logMessage += ",PrepNo [" + (newExperiment.PreparationNo ?? "NULL") + "]";
                    pdcMessage.Message += ",PrepNo";
                }
                if (newTestdata.TestVersion.IsMcNoExpLevel)
                {
                    if (experimentAlreadyFound)
                    {
                        experimentAlreadyFound = (newExperiment.MCNo == null && experiment.MCNo == null) ||
                                                    (newExperiment.MCNo != null && experiment.MCNo != null &&
                                                     newExperiment.MCNo.ToUpper().Trim().Equals(experiment.MCNo));
                    }
                    logMessage += ",McNo [" + (newExperiment.MCNo ?? "NULL") + "]";
                    pdcMessage.Message += ",McNo";

                }

                // test for each exp level variable if values are both null or the same
                foreach (int varid in newTestdata.TestVersion.ExperimentLevelVariables.Keys)
                {
                    TestVariableValue newvar = newExperiment.GetExperimentValue(varid);
                    TestVariableValue tmpvar = experiment.GetExperimentValue(varid);
                    if (experimentAlreadyFound)
                    {
                        experimentAlreadyFound = (newvar == null && tmpvar == null) ||
                                                    (newvar != null && tmpvar != null && newvar.ValueChar != null && tmpvar.ValueChar != null &&
                                                     newvar.ValueChar.Equals(tmpvar.ValueChar));
                    }
                    logMessage += "," + newTestdata.TestVersion.ExperimentLevelVariables[varid].VariableName + " [" + ((newvar == null || newvar.ValueChar == null) ? "NULL" : newvar.ValueChar) + "]";
                    pdcMessage.Message += "," + newTestdata.TestVersion.ExperimentLevelVariables[varid].VariableName;
                }
                if (experimentAlreadyFound)
                {
                    newExperiment.ExperimentNo = experiment.ExperimentNo;
                    newExperiment.UploadId = experiment.UploadId;
                    pdcMessage.Message = pdcMessage.Message.Remove(0, 1); //Remove leading ','
                    logMessage = logMessage.Remove(0, 1);
                    PDCLogger.TheLogger.LogWarning(PDCLogger.LOG_NAME_EXCEL, "Following ELP are already in PDC: " + logMessage);
                    if (messages.Count == 0)
                    {
                        messages.Add(pdcMessage);
                    }
                }
            }
        }

        private List<PDCMessage> FillExperimentNos(Testdata newTestdata, WS.Output output)
        {
            List<PDCMessage> tmpMessages = new List<PDCMessage>();
            if (output == null || output.output == null || output.output.Length == 0)
            {
                return tmpMessages;
            }

            if (output.output.Length == 1 && output.output[0] == null)
            {
                //Workaround: Axis2 sends an 1-element array with a null value, where
                //one would expect an empty array
                return tmpMessages;
            }
            newTestdata.NewData = false;
            ExperimentData tmpExperiment = new ExperimentData(newTestdata.TestVersion);
            Dictionary<int, TestVariableValue> tmpValues = new Dictionary<int, TestVariableValue>();

            foreach (WS.PDCParameter tmpParam in output.output)
            {
                switch (tmpParam.id)
                {
                    case PDCConstants.C_ID_RUNNO:
                        tmpExperiment.Runno = (long?)tmpParam.valueNum;
                        break;
                    case PDCConstants.C_ID_COMPOUNDIDENTIFIER:
                        tmpExperiment.CompoundNo = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_PREPARATIONNO:
                        tmpExperiment.PreparationNo = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_EXPERIMENTNO:
                        tmpExperiment.ExperimentNo = (long?)tmpParam.valueNum;
                        break;
                    case PDCConstants.C_ID_UPLOAD_ID:
                        tmpExperiment.UploadId = (long?)tmpParam.valueNum;
                        break;
                    case PDCConstants.C_ID_PERSONID:
                        tmpExperiment.PersonId = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_PDC_ONLY_DATA:
                        tmpExperiment.ReportToPix = tmpParam.valueChar.Equals("Y") ? "No" : "Yes";
                        break;

                    case PDCConstants.C_ID_ORIGIN:
                        tmpExperiment.Origin = tmpParam.valueChar;
                        break;
                    case PDCConstants.C_ID_TESTPARAMETER_BY_ID:
                        TestVariableValue variableValue = GetOrCreateValue(tmpValues, tmpParam);
                        if (variableValue == null)
                            continue;

                        bool isVariableIdValid = newTestdata.TestVersion.VariableMap.ContainsKey(variableValue.VariableId);

                        if (!isVariableIdValid)
                            continue;

                        if (!newTestdata.TestVersion.VariableMap[variableValue.VariableId].IsExperimentLevelReference ||
                             newTestdata.TestVersion.VariableMap[variableValue.VariableId].IsMeasurementLevel)
                        {
                            continue;
                        }
                        variableValue.ValueChar = tmpParam.valueChar;
                        variableValue.Prefix = tmpParam.praefix;
                        tmpExperiment.GetExperimentValues().Add(variableValue);
                        break;
                    case PDCConstants.C_ID_SEPERATOR:
                        // find experiment in newTestdata and set experimentno
                        SetIdentifierFromTestdata(newTestdata, tmpExperiment, tmpMessages);
                        tmpExperiment = new ExperimentData(newTestdata.TestVersion);
                        tmpValues.Clear();
                        break;
                }
            }
            // check last row without separator
            if (tmpExperiment.ExperimentNo != null && tmpExperiment.ExperimentNo > 0)
            {
                SetIdentifierFromTestdata(newTestdata, tmpExperiment, tmpMessages);
            }
            return tmpMessages;
        }
        #endregion

        #region GetOrCreateValue
        private TestVariableValue GetOrCreateValue(Dictionary<int, TestVariableValue> theValues, WS.PDCParameter parameter)
        {
            TestVariableValue variableValue;
            if (parameter.variableId == null)
            {
                return null;
            }
            int variableId = (int)parameter.variableId;
            if ((parameter.position ?? 0) > 0)
            {
                TestVariableValue tmpValue = new TestVariableValue(variableId) {Position = parameter.position};
                return tmpValue;
            }
            if (!theValues.ContainsKey(variableId))
            {
                variableValue = new TestVariableValue(variableId);
                theValues[variableId] = variableValue;
            }
            else
            {
                variableValue = theValues[variableId];
            }
            return variableValue;
        }
        #endregion

        #region ValidateTestdata
        /// <summary>
        /// Validates the test data against the server and returns the validation messages
        /// </summary>
        /// <param name="aTestdata"></param>
        /// <returns></returns>
        public List<PDCMessage> ValidateTestdata(Testdata aTestdata)
        {
            return ValidateAndUpload(aTestdata, UploadAction.Validate, null);
        }
        #endregion

        #endregion
    }
}
