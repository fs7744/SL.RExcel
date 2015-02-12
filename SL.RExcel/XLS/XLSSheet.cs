using SL.RExcel.XLS.Records;
using System.Collections.Generic;
using System.Linq;

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
            HandleMergeCells(record);
        }

        private void HandleMergeCells(SheetRecord record)
        {
            foreach (var mergeCells in record.MergeCells)
            {
                foreach (var mergeCell in mergeCells.MergeCells)
                {
                    var value = GetFirstMergeCellValue(mergeCell);
                    if (value != null)
                        SetMergeCellValue(mergeCell, value);
                }
            }
        }

        private void SetMergeCellValue(MergeCell mergeCell, object value)
        {
            for (uint i = mergeCell.FirstRow; i <= mergeCell.LastRow; i++)
            {
                IRow row = GetXLSRow(i);
                for (uint j = mergeCell.FirstCol; j <= mergeCell.LastCol; j++)
                {
                    ICell cell = null;
                    if (!row.Cells.TryGetValue(j, out cell))
                    {
                        cell = new XLSCell(value);
                        row.Cells.Add(j, cell);
                    }
                }
            }
        }

        private IRow GetXLSRow(uint i)
        {
            IRow row = null;
            if (!Rows.TryGetValue(i, out row))
            {
                row = new XLSRow();
                Rows.Add(i, row);
            }
            return row;
        }

        private object GetFirstMergeCellValue(MergeCell mergeCell)
        {
            object result = null;
            for (uint i = mergeCell.FirstRow; i <= mergeCell.LastRow; i++)
            {
                IRow row = null;
                if (!Rows.TryGetValue(i, out row))
                    continue;
                result = GetFirstMergeCellValue(mergeCell, row);
                if (result != null)
                    break;
            }
            return result;
        }

        private object GetFirstMergeCellValue(MergeCell mergeCell, IRow row)
        {
            object result = null;
            for (uint j = mergeCell.FirstCol; j <= mergeCell.LastCol; j++)
            {
                ICell cell = null;
                result = !row.Cells.TryGetValue(j, out cell) || cell.Value == null
                    ? null : result;

                if (result != null)
                    break;
            }
            return result;
        }

        public IRow GetRow(uint index)
        {
            IRow result = null;
            Rows.TryGetValue(index, out result);
            return result;
        }

        public IEnumerable<KeyValuePair<uint, IRow>> GetAllRows()
        {
            return Rows.OrderBy(i => i.Key);
        }
    }
}