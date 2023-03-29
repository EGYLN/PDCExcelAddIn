using System;
using System.Collections.Generic;
using BBS.ST.BHC.BSP.PDC.Lib.Exceptions;
using WS = BBS.ST.BHC.BSP.PDC.Lib.PDCWebservice;
namespace BBS.ST.BHC.BSP.PDC.Lib
{
    class SearchTestdataService
    {
        private readonly PDCService myService;

        #region constructor
        public SearchTestdataService(PDCService service)
        {
            myService = service;
        }
        #endregion

        #region methods

        #region ExtractTestdata
        private Testdata ExtractTestdata(Testdefinition testdefinition, List<TestdataSearchCriteria> searchCriterias, WS.Output output)
        {
            bool duplicate = false;
            if (output == null || output.output == null || output.output.Length == 0)
            {
                return null;
            }

            if (output.output.Length == 1 && output.output[0] == null)
            {
                //Workaround: Axis2 sends an 1-element array with a null value, where
                //one would expect an empty array
                return null;
            }
            Testdata testdata = new Testdata(testdefinition, searchCriterias) {NewData = false};
            ExperimentData experiment = new ExperimentData(testdefinition);
            Dictionary<int, TestVariableValue> values = new Dictionary<int, TestVariableValue>();

            foreach (WS.PDCParameter parameter in output.output)
            {
                switch (parameter.id)
                {
                    case PDCConstants.C_ID_RUNNO:
                        experiment.Runno = (long?)parameter.valueNum;
                        break;
                    case PDCConstants.C_ID_COMPOUNDIDENTIFIER:
                        experiment.CompoundNo = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_PREPARATIONNO:
                        experiment.PreparationNo = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_MCNO:
                        experiment.MCNo = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_REFERENCE:
                        experiment.Reference = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_EXPERIMENTNO:
                        experiment.ExperimentNo = (long?)parameter.valueNum;
                        break;
                    case PDCConstants.C_ID_UPLOAD_ID:
                        experiment.UploadId = (long?)parameter.valueNum;
                        break;
                    case PDCConstants.C_ID_ALTERNATE_PREPNO:
                        experiment.AlternatePrepno = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_ORIGIN:
                        experiment.Origin = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_DATE_RESULT:
                        experiment.DateResult = PDCConverter.Converter.ToDate(parameter.valueChar);
                        break;
                    case PDCConstants.C_ID_UPLOADDATE:
                        experiment.UploadDate = PDCConverter.Converter.ToDate(parameter.valueChar);
                        break;
                    case PDCConstants.C_ID_REPLICATED_AT:
                        experiment.ReplicatedAt = PDCConverter.Converter.ToDate(parameter.valueChar);
                        break;
                    case PDCConstants.C_ID_SCHEDULEDDATE:
                        experiment.ScheduledDate = PDCConverter.Converter.ToDate(parameter.valueChar);
                        break;
                    case PDCConstants.C_ID_ASSAY_REFERENCE:
                        experiment.AssayReference = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_RESULT_STATUS:
                        experiment.ResultStatus = parameter.valueNum;
                        break;
                    case PDCConstants.C_ID_PERSONID:
                        experiment.PersonId = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_PERSONID_TYPE:
                        experiment.PersonIdType = (int?)parameter.valueNum;
                        break;
                    case PDCConstants.C_ID_PDC_ONLY_DATA:
                        experiment.ReportToPix = parameter.valueChar.Equals("Y") ? "No" : "Yes";
                        break;

                    case PDCConstants.C_ID_TESTPARAMETER_BY_ID:
                        TestVariableValue variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        if (!testdefinition.VariableMap.ContainsKey(variableValue.VariableId))
                        {
                            continue;
                        }
                        variableValue.ValueBlob = parameter.valueBlob;
                        variableValue.ValueChar = parameter.valueChar;
                        variableValue.Prefix = parameter.praefix;
                        // add the experimentLevelVariables  to the collection
                        if (IsExperimentLevelVariable(testdefinition, variableValue.VariableId))
                        {
                            if (!experiment.GetExperimentLevelVariableValues().Contains(variableValue))
                            {
                                experiment.GetExperimentLevelVariableValues().Add(variableValue);
                            }
                        }
                        if (IsMeasurementLevel(testdefinition, variableValue.VariableId))
                        {
                            if (!experiment.GetMeasurementValues().Contains(variableValue))
                            {
                                experiment.GetMeasurementValues().Add(variableValue);
                            }
                        }

                        else
                        {
                            if (!experiment.GetExperimentValues().Contains(variableValue))
                            {
                                experiment.GetExperimentValues().Add(variableValue);
                            }
                        }

                        break;
                    case PDCConstants.C_ID_FILE_NAME:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.Filename = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_FILE_FORMAT:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.Fileformat = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_FILE_CONTENTS:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.ValueBlob = parameter.valueBlob;
                        break;
                    case PDCConstants.C_ID_FILE_NAME_BY_ID:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.Filename = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_FILE_FORMAT_BY_ID:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.Fileformat = parameter.valueChar;
                        break;
                    case PDCConstants.C_ID_FILE_CONTENTS_BY_ID:
                        variableValue = GetOrCreateValue(values, parameter, ref duplicate);
                        if (variableValue == null)
                        {
                            continue;
                        }
                        variableValue.ValueBlob = parameter.valueBlob;
                        break;
                    case PDCConstants.C_ID_SEPERATOR:
                        testdata.Experiments.Add(experiment);
                        experiment = new ExperimentData(testdefinition);
                        values.Clear();
                        break;
                }
            }

            if (output.output[output.output.Length - 1].id != PDCConstants.C_ID_SEPERATOR)
            {
                testdata.Experiments.Add(experiment);
            }
            testdata.SortExperiments();
            return testdata;
        }
        #endregion

        #region FindTestdata
        public Testdata FindTestdata(Testdefinition testdefinition, List<TestdataSearchCriteria> searchCriterias)
        {
            if (searchCriterias == null || searchCriterias.Count == 0)
            {
                return null;
            }
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
            int i = 0;
            foreach (TestdataSearchCriteria criteria in searchCriterias)
            {
                i++;
                foreach (int parameterId in criteria.UploadParameters.Keys)
                {
                    parameter = new WS.PDCParameter
                    {
                        id = parameterId,
                        idSpecified = true,
                        position = i,
                        positionSpecified = true
                    };
                    TestVariableValue variableValue = criteria[parameterId, false];
                    parameter.valueChar = variableValue.ValueChar;
                    parameter.variableType = variableValue.VariableType;
                    parameters.Add(parameter);
                }

                foreach (int variableId in criteria.Variables.Keys)
                {
                    parameter = new WS.PDCParameter
                    {
                        id = PDCConstants.C_ID_TESTPARAMETER_BY_ID,
                        idSpecified = true,
                        variableId = variableId,
                        variableIdSpecified = true,
                        position = i,
                        positionSpecified = true
                    };

                    TestVariableValue variableValue = criteria[variableId, true];
                    parameter.valueChar = variableValue.ValueChar;
                    parameter.variableType = variableValue.VariableType;
                    parameter.praefix = variableValue.Prefix;
                    parameters.Add(parameter);
                }
            }
            parameter = new WS.PDCParameter
            {
                id = PDCConstants.C_ID_PARTIAL_COMPARE,
                idSpecified = true,
                valueNum = 1,
                valueNumSpecified = true
            };
            parameters.Add(parameter);
            WS.Input input = new WS.Input {input = parameters.ToArray()};

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
                return ExtractTestdata(testdefinition, searchCriterias, output);
            }
        }
        #endregion

        #region GetOrCreateValue
        private TestVariableValue GetOrCreateValue(Dictionary<int, TestVariableValue> theValues, WS.PDCParameter parameter, ref bool isDuplicate)
        {
            TestVariableValue variableValue;
            if (parameter.variableId == null)
            {
                return null;
            }
            int variableId = (int)parameter.variableId;
            if ((parameter.position ?? 0) > 0)
            {
                TestVariableValue mearurementValue = new TestVariableValue(variableId) {Position = parameter.position};
                return mearurementValue;
            }
            if (!theValues.ContainsKey(variableId))
            {
                variableValue = new TestVariableValue(variableId);
                theValues[variableId] = variableValue;
            }
            else
            {
                variableValue = theValues[variableId];
                isDuplicate = true;
            }
            return variableValue;
        }
        #endregion

        #region IsExperimentLevelVariable
        private bool IsExperimentLevelVariable(Testdefinition testdefinition, int aVariableId)
        {
            //TODO: ->testdefinition.ExperimentLevelVariables.ContainsKey(aVariableId)
            foreach (int variableId in testdefinition.ExperimentLevelVariables.Keys)
            {
                if (variableId == aVariableId) return true;
            }
            return false;
        }
        #endregion

        #region IsMeasurementLevel
        private bool IsMeasurementLevel(Testdefinition testdefinition, int aVariableId)
        {
            return testdefinition.VariableMap[aVariableId].IsMeasurementLevel;
        }
        #endregion

        #endregion
    }
}
