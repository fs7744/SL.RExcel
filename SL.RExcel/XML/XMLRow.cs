using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace SL.RExcel.XML
{
    public class XMLRow : IRow
    {
        public XMLRow()
        {
            Cells = new Dictionary<uint, ICell>();
        }

        public void AddCells(XElement row)
        {
            uint index = Cells.Keys.IsEmpty() ? 0 : Cells.Keys.Max() + 1;
            foreach (var cellx in row.Descendants(XMLCommon.Cell))
            {
                var cell = new XMLCell(index, cellx);
                index = cell.Index;
                Cells.Add(index, cell);
                index++;
                index = SetMergerAcross(index, cell);
            }
        }

        public void AddCell(XMLCell cell)
        {
            if (!Cells.ContainsKey(cell.Index))
                Cells.Add(cell.Index, cell);
        }

        private uint SetMergerAcross(uint index, XMLCell cell)
        {
            if (cell.MergeAcross > 0)
            {
                for (int i = 1; i <= cell.MergeAcross; i++)
                {
                    Cells.Add(index, new XMLCell(index, cell.Value));
                    index++;
                }
            }
            return index;
        }

        public IDictionary<uint, ICell> Cells { get; private set; }

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

        public IEnumerable<XMLCell> GetMergeDownCell()
        {
            return Cells.Values.Select(i => (XMLCell)i).Where(i => i.MergeDown > 0);
        }
    }
}