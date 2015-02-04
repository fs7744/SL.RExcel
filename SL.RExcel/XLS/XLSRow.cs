using SL.RExcel.XLS.Records;
using System.Collections.Generic;
using System.Linq;

namespace SL.RExcel.XLS
{
    public class XLSRow : IRow
    {
        public IDictionary<uint, ICell> Cells { get; private set; }

        public XLSRow(RowRecord record, SheetRecord sheet)
        {
            Cells = new Dictionary<uint, ICell>();
            foreach (var cell in sheet.Cells.Where(i => i.Row == record.RowNumber))
            {
                if (cell is LabelSstRecord)
                {
                    (cell as LabelSstRecord).SetValue(sheet.SST);
                }

                if (cell is SingleColCellRecord)
                {
                    Cells.Add((cell as SingleColCellRecord).Col, new XLSCell(cell));
                }
            }
        }

        public bool IsEmpty()
        {
            return Cells.Count == 0;
        }

        public ICell GetCell(uint index)
        {
            ICell result = null;
            Cells.TryGetValue(index, out result);
            return result;
        }
    }
}