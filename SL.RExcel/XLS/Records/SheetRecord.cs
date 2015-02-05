using System.Collections.Generic;

namespace SL.RExcel.XLS.Records
{
    public class SheetRecord
    {
        public SheetRecord()
        {
            Cells = new List<CellRecord>();
            Rows = new List<RowRecord>();
            MergeCells = new List<MergeCellsRecord>();
        }

        public SstRecord SST { get; set; }

        public BoundSheetRecord Sheet { get; set; }

        public IndexRecord Index { get; set; }

        public List<CellRecord> Cells { get; set; }

        public List<RowRecord> Rows { get; set; }

        public List<MergeCellsRecord> MergeCells { get; set; }
    }
}