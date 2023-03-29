using Microsoft.Office.Interop.Excel;

namespace BBS.ST.BHC.BSP.PDC.ExcelClient.Actions
{
    internal class Ranges
    {
        public Range CompoundnoRange;
        public Range PrepnoRange;
        public Range WeightRange;
        public Range FormulaRange;
        public Range McNoRange;
        public Range StructureColumn;
        public bool Vertical;
        public bool OnPdcList;

        public Ranges GetRangesFromArea(int areaIndex)
        {
            var result = new Ranges
            {
                CompoundnoRange = CompoundnoRange != null ? CompoundnoRange.Areas[areaIndex] : null,
                PrepnoRange = PrepnoRange != null ? PrepnoRange.Areas[areaIndex] : null,
                WeightRange = WeightRange != null ? WeightRange.Areas[areaIndex] : null,
                FormulaRange = FormulaRange != null ? FormulaRange.Areas[areaIndex] : null,
                McNoRange = McNoRange != null ? McNoRange.Areas[areaIndex] : null,
                StructureColumn = StructureColumn != null ? StructureColumn.Areas[areaIndex] : null,
                Vertical = Vertical,
                OnPdcList = OnPdcList
            };

            return result;
        }
    }
}
