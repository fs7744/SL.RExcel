using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SL.RExcel.MHT
{
    public class MHTWorksheet : IWorksheet
    {
        public const string TableReg = "<table.+?</table>";
        public const string Tr = "tr";
        public const string TrReg = "<tr.+?</tr>";

        public string Name { get; private set; }

        public IDictionary<uint, IRow> Rows { get; private set; }

        public uint FirstRow { get; private set; }

        public uint LastRow { get; private set; }

        public MHTWorksheet(string name, Part part)
        {
            Name = name;
            Rows = new Dictionary<uint, IRow>();
            var element = GetXML(part);
            if (string.IsNullOrWhiteSpace(element)) return;
            uint index = 0;
            foreach (Match xrow in new Regex(TrReg, RegexOptions.IgnoreCase).Matches(element))
            {
                if (!xrow.Success) continue;
                var row = GetMHTRow(index);
                row.AddCells(xrow.Value);
                FillMergeDownCells(index, row.GetMergeDownCell());
                index++;
            }
        }

        private void FillMergeDownCells(uint index, IEnumerable<MHTCell> cells)
        {
            foreach (var cell in cells)
            {
                for (uint i = index + 1; i <= index + cell.RowSpan; i++)
                {
                    var row = GetMHTRow(i);
                    row.AddCell(new MHTCell(cell.Index, cell.Value));
                }
            }
        }

        private MHTRow GetMHTRow(uint index)
        {
            IRow result = null;
            if (!Rows.TryGetValue(index, out result))
            {
                result = new MHTRow();
                Rows.Add(index, result);
            }
            return result as MHTRow;
        }

        private string GetXML(Part part)
        {
            var match = new Regex(TableReg, RegexOptions.IgnoreCase).Match(part.ToString());
            return !match.Success ? null : match.Value;
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