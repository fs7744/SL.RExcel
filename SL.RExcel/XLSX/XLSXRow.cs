using System.Collections.Generic;
using System.Xml.Linq;

namespace SL.RExcel.XLSX
{
    public class XLSXRow : IRow
    {
        public XLSXRow()
        {
            Cells = new Dictionary<uint, ICell>();
        }

        public XLSXRow(XElement row, Dictionary<string, string> sharedStrings)
        {
            Cells = new Dictionary<uint, ICell>();
            foreach (var cell in row.Descendants(XLSXCommon.ExcelNamespace + XLSXCommon.XML_C))
            {
                var index = cell.Attribute(XLSXCommon.XML_R).Value.ConvertIndex();

                if (cell.HasElements)
                {
                    var value = cell.Element(XLSXCommon.ExcelNamespace + XLSXCommon.XML_V).Value;
                    var t = cell.Attribute(XLSXCommon.XML_T);
                    if (t != null && t.Value == XLSXCommon.XML_S)
                    {
                        value = sharedStrings[value];
                    }

                    Cells.Add((uint)index[0], new XLSXCell(value));
                }
            }
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
    }
}