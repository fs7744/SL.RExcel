using System.Text.RegularExpressions;

namespace SL.RExcel.MHT
{
    public class MHTCell : ICell
    {
        public const string xml_RowSpan = "rowspan";
        public const string xml_ColSpan = "colspan";
        public static readonly Regex ValueReg = new Regex(">.+?</td>", RegexOptions.IgnoreCase);
        public static readonly Regex RowSpanReg = new Regex("rowspan=.+? ", RegexOptions.IgnoreCase);
        public static readonly Regex ColSpanReg = new Regex("colspan=.+? ", RegexOptions.IgnoreCase);

        public MHTCell(uint index, string element)
        {
            Index = index;
            if (!string.IsNullOrWhiteSpace(element))
            {
                Value = GetValue(element, ValueReg, ">", "</td");
                ColSpan = ToUint(GetValue(element, ColSpanReg, "colspan=", " "));
                RowSpan = ToUint(GetValue(element, RowSpanReg, "rowspan=", " "));
            }
        }

        private string GetValue(string element, Regex reg, string begin, string end)
        {
            var match = reg.Match(element);
            return match.Success ? match.Value.Replace(begin, string.Empty).Replace(end, string.Empty) : string.Empty;
        }

        private uint ToUint(string value)
        {
            uint result = 0;
            return string.IsNullOrWhiteSpace(value) || !uint.TryParse(value, out result) ? 0 : result;
        }

        public MHTCell(uint index, object value)
        {
            Index = index;
            Value = value;
        }

        public uint Index { get; private set; }

        public uint ColSpan { get; private set; }

        public uint RowSpan { get; private set; }

        public object Value { get; private set; }

        public string GetStringValue()
        {
            return Value.ToString();
        }
    }
}