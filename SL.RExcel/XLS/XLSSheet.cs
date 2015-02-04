using SL.RExcel.XLS.Records;
using System.Collections.Generic;

namespace SL.RExcel.XLS
{
    public class XLSSheet : IWorksheet
    {
        public string Name { get; private set; }

        public IDictionary<uint, IRow> Rows { get; private set; }

        public uint FirstRow { get; private set; }

        public uint LastRow { get; private set; }

        //public uint FirstCol { get; private set; }

        //public uint LastCol { get; private set; }

        public XLSSheet(SheetRecord record)
        {
            Name = record.Sheet.Name;
            FirstRow = record.Index.FirstRow;
            LastRow = record.Index.LastRow;
            Rows = new Dictionary<uint, IRow>();
            foreach (var item in record.Rows)
            {
                Rows.Add(item.RowNumber, new XLSRow(item, record));
            }
        }

        public IRow GetRow(uint index)
        {
            IRow result = null;
            Rows.TryGetValue(index, out result);
            return result;
        }
    }
}