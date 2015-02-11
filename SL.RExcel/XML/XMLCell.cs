using System.Linq;
using System.Xml.Linq;

namespace SL.RExcel.XML
{
    public class XMLCell : ICell
    {
        public XMLCell(uint index, XElement element)
        {
            Index = ToUint(element.Attribute(XMLCommon.Index));
            if (Index == 0)
                Index = index;
            Value = element.HasElements ? element.Descendants(XMLCommon.Data).First().Value : string.Empty;
            MergeAcross = ToUint(element.Attribute(XMLCommon.MergeAcross));
            MergeDown = ToUint(element.Attribute(XMLCommon.MergeDown));
        }

        private uint ToUint(XAttribute xAttribute)
        {
            uint result = 0;
            return xAttribute == null || !uint.TryParse(xAttribute.Value, out result) ? 0 : result;
        }

        public XMLCell(uint index, object value)
        {
            Index = index;
            Value = value;
        }

        public uint Index { get; private set; }

        public uint MergeAcross { get; private set; }

        public uint MergeDown { get; private set; }

        public object Value { get; private set; }

        public string GetStringValue()
        {
            return Value.ToString();
        }
    }
}