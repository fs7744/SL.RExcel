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
                IRow row = null;
                if (!Rows.TryGetValue(i, out row))
                    Rows.Add(i, new XLSRow());
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

        private object GetFirstMergeCellValue(MergeCell mergeCell)
        {
            for (uint i = mergeCell.FirstRow; i <= mergeCell.LastRow; i++)
            {
                IRow row = null;
                if (!Rows.TryGetValue(i, out row))
                    continue;
                for (uint j = mergeCell.FirstCol; j <= mergeCell.LastCol; j++)
                {
                    ICell cell = null;
                    if (!row.Cells.TryGetValue(j, out cell) || cell.Value == null)
                        continue;
                    else
                        return cell.Value;
                }
            }
            return null;
        }

        public IRow GetRow(uint index)
        {
            IRow result = null;
            Rows.TryGetValue(index, out result);
            return result;
        }
    }
}