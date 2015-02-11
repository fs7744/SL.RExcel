using System.Collections.Generic;
using System.Xml.Linq;

namespace SL.RExcel.XML
{
    public class XMLWorksheet : IWorksheet
    {
        public string Name { get; private set; }

        public IDictionary<uint, IRow> Rows { get; private set; }

        public uint FirstRow { get; private set; }

        public uint LastRow { get; private set; }

        public XMLWorksheet(XElement element)
        {
            Name = element.Attribute(XMLCommon.Name).Value;
            Rows = new Dictionary<uint, IRow>();
            uint index = 0;
            foreach (var xrow in element.Descendants(XMLCommon.Row))
            {
                var row = GetXMLRow(index);
                row.AddCells(xrow);
                FillMergeDownCells(index, row.GetMergeDownCell());
                index++;
            }
        }

        private XMLRow GetXMLRow(uint index)
        {
            IRow result = null;
            if (!Rows.TryGetValue(index, out result))
            {
                result = new XMLRow();
                Rows.Add(index, result);
            }
            return result as XMLRow;
        }

        private void FillMergeDownCells(uint index, IEnumerable<XMLCell> cells)
        {
            foreach (var cell in cells)
            {
                for (uint i = index + 1; i <= index + cell.MergeDown; i++)
                {
                    var row = GetXMLRow(i);
                    row.AddCell(new XMLCell(cell.Index, cell.Value));
                }
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