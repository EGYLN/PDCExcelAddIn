namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{

    // handles the plain ExperimentData and Measurementdata for the SMT within a two dimensional array.
    class ValueMatrixForSMT
    {
        object[,] myValues;
        int myRow = 0;

        internal object missing = global::System.Type.Missing;

        #region constructor
        public ValueMatrixForSMT(int rows, int columns)
        {
            myValues = new object[rows, columns];
            myRow = 0;
        }
        #endregion

        #region methods

        #region WriteExperimentLevelValues
        /// <summary>
        /// writes all experimentvalues of one experiment into the twodimensional array (used for writing the range in one go)
        /// (values which mark to which experiment the data belong)
        /// </summary>
        /// <param name="experiment">current experiment with the values to </param>
        /// <param name="pdcListObject">SMT PdcListObject</param>
        private void WriteExperimentLevelValues(Lib.ExperimentData experiment, PDCListObject pdcListObject, UniqueExperimentKeyHandler uniqueExperimentKeyHandler)
        {
            for (int additionalRow = 0; additionalRow < experiment.MaxNumberOfMeasurementValues; additionalRow++)
            {
                myValues[myRow + additionalRow, pdcListObject.GetColumnIndex(PDCExcelConstants.COMPOUNDNO).Value] = experiment.CompoundNo;
                myValues[myRow + additionalRow, pdcListObject.GetColumnIndex(PDCExcelConstants.EXPERIMENT_NO).Value] = experiment.ExperimentNo;
                myValues[myRow + additionalRow, pdcListObject.GetColumnIndex(PDCExcelConstants.PREPARATIONNO).Value] = experiment.PreparationNo;
                int? mcNoIndex = pdcListObject.GetColumnIndex(PDCExcelConstants.MCNO);
                if (mcNoIndex != null)
                {
                    myValues[myRow + additionalRow, mcNoIndex.Value] = experiment.MCNo;
                }

                // Row must be added here, because it is possible that there are no values within experimentLevelvariable
                uniqueExperimentKeyHandler.AddMeasurementRow(experiment, myRow + additionalRow);

                foreach (Lib.TestVariableValue experimentValueForSMT in experiment.GetExperimentLevelVariableValues())
                {
                    Lib.TestVariable testVariable = pdcListObject.Testdefinition.ExperimentLevelVariables[experimentValueForSMT.VariableId];
                    int columns = pdcListObject.GetColumnIndex(testVariable.VariableId).Value;


                    if (testVariable.IsNumeric())
                    {
                        myValues[myRow + additionalRow, columns] =
                          Lib.PDCConverter.Converter.NumericString2Double(experimentValueForSMT.ValueChar, experimentValueForSMT.Prefix, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                    }
                    else
                    {
                        myValues[myRow + additionalRow, columns] = experimentValueForSMT.ValueChar;
                    }
                }
            }
        }
        #endregion
        #region MergeWithPrefix
        /// <summary>
        /// Adds the prefix to the string value.
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        private object MergeWithPrefix(string prefix, string value)
        {
            string preFix = prefix.Trim();
            if (preFix == "=")
            {
                preFix = "'=";
            }
            return preFix + (value == null ? "" : " " + value);
        }
        #endregion
        #region WriteMeasurementValues
        /// <summary>
        /// writes all MeasurementValues of one experiment into the twodimensional array (used for writing the range in one go)
        /// </summary>
        /// <param name="experiment">current experiment with the values</param>
        /// <param name="pdcListObject">SMT PdcListObject</param>
        private void WriteMeasurementValues(Lib.ExperimentData experiment, PDCListObject pdcListObject)
        {
            int additionalRow = 0;
            foreach (Lib.TestVariableValue variableValue in experiment.GetMeasurementValues())
            {
                if (!variableValue.Position.HasValue)
                {
                    continue;
                }
                Lib.TestVariable testVariable = pdcListObject.Testdefinition.MeasurementVariables[variableValue.VariableId];
                int columns = pdcListObject.GetColumnIndex(variableValue.VariableId).Value;
                additionalRow = variableValue.Position.Value - 1;

                if (testVariable.IsNumeric())
                {
                    myValues[myRow + additionalRow, columns] =
                      Lib.PDCConverter.Converter.NumericString2Double(variableValue.ValueChar, variableValue.Prefix, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                    //double tmpvalue = Lib.PDCConverter.Converter.FromDecimal(variableValue.ValueChar, ExcelUtils.TheUtils.GetExcelNumberSeparators());
                    //if (variableValue.Prefix != null && variableValue.Prefix.Trim() != "")
                    //{
                    //  myValues[myRow + additionalRow, columns] = MergeWithPrefix(variableValue.Prefix, System.Convert.ToString(tmpvalue, ExcelUtils.TheUtils.GetExcelNumberSeparators()));
                    //}
                    //else
                    //{
                    //  myValues[myRow + additionalRow, columns] = tmpvalue;
                    //}
                    ////myValues[myRow + additionalRow, columns] = variableValue.ValueNum;
                }
                else
                {
                    myValues[myRow + additionalRow, columns] = variableValue.ValueChar;
                }
            }
        }
        #endregion

        #region WriteValues
        /// <summary>
        /// writes all measurementvalues from one experiment into the twodimensional array (used for writing the range in one go)
        /// </summary>
        /// <param name="experiment">current experiment with the values to </param>
        /// <param name="pdcListObject">SMT PdcListObject</param>
        public void WriteValues(Lib.ExperimentData experiment, PDCListObject pdcListObject, UniqueExperimentKeyHandler uniqueExperimentKeyHandler, int experimentRow)
        {
            if (experiment is Lib.PlaceHolderExperiment)
            {
                //myRow++;
                return; //No data ignore
            }
            uniqueExperimentKeyHandler.SetExperiment(experiment, experimentRow);
            WriteMeasurementValues(experiment, pdcListObject);
            WriteExperimentLevelValues(experiment, pdcListObject, uniqueExperimentKeyHandler);
            myRow += experiment.MaxNumberOfMeasurementValues;
        }
        #endregion

        #endregion

        #region properties

        #region Values
        /// <summary>
        /// gets the twodimensional array
        /// </summary>
        public object[,] Values
        {
            get
            {
                return myValues;
            }
            set
            {
                myValues = value;
            }
        }
        #endregion

        #endregion
    }
}
