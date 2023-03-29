using System;
using System.Collections.Generic;
using System.Linq;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient
{
    class UniqueExperimentKey : IEquatable<UniqueExperimentKey>
    {
        internal static string NullValue
        {
            get
            {
                return "[NULL]";
            }
        }

        #region Variables

        bool myUseExperimentNo;

        SortedDictionary<string, string> myExperimentLevelVariables;
        #endregion

        #region constructor

        internal UniqueExperimentKey(bool useExperimentNo)
        {
            myExperimentLevelVariables = new SortedDictionary<string, string>();
            myUseExperimentNo = useExperimentNo;
        }

        #endregion

        #region properties

        public bool IsCompoundNoExpLevel { get; set; }
        public bool IsPrepNoExpLevel { get; set; }
        public bool IsMcNoExpLevel { get; set; }
        public string McNo { get; set; }
        internal string CompoundNo { get; set; }
        internal long? ExperimentNo { get; set; }
        internal string PreparationNo { get; set; }
        internal int RowNumber { get; set; }

        internal SortedDictionary<string, string> ExperimentLevelVariables
        {
            get { return myExperimentLevelVariables; }
            set { myExperimentLevelVariables = value; }
        }

        public string Key
        {
            get
            {
                return FormattedKey();
            }
        }

        public bool UseExperimentNo
        {
            get { return myUseExperimentNo; }
            set { myUseExperimentNo = value; }
        }

        #endregion

        public string FormattedKey(string delimiter = null)
        {
            var key = IsCompoundNoExpLevel ? CompoundNo + delimiter : string.Empty;
            key = key + (IsPrepNoExpLevel ? PreparationNo + delimiter : string.Empty);
            key = key + (IsMcNoExpLevel ? McNo + delimiter : string.Empty);
            key = key + (myUseExperimentNo ? ExperimentNo + delimiter : string.Empty);
            
            return myExperimentLevelVariables.Values.Aggregate(key, (current, value) => current + value + delimiter);
        }

        public bool Equals(UniqueExperimentKey other)
        {
            return (!IsCompoundNoExpLevel || other.CompoundNo == CompoundNo) &&
                        (!IsPrepNoExpLevel || other.PreparationNo == PreparationNo) &&
                        (!IsMcNoExpLevel || other.McNo == McNo) &&
                        !ExperimentLevelVariables.Values.Except(other.ExperimentLevelVariables.Values).Any();
        }

        public bool IsNull()
        {
            return CompoundNo == NullValue &&
                   PreparationNo == NullValue &&
                   McNo == NullValue &&
                   ExperimentLevelVariables.Values.All(value => value == NullValue);
        }
    }
}
