using System.Text.RegularExpressions;
using System.Windows.Browser;

namespace SL.RExcel.MHT
{
    public class MHTCell : ICell
    {
        public const string xml_RowSpan = "rowspan";
        public const string xml_ColSpan = "colspan";
        public static readonly Regex ValueReg = new Regex(">.+?</td>", RegexOptions.IgnoreCase);
        public static readonly Regex RowSpanReg = new Regex("rowspan=.+? ", RegexOptions.IgnoreCase);
        public static readonly Regex ColSpanReg = new Regex("colspan=.+? ", RegexOptions.IgnoreCase);
        public static readonly Regex SpanReg = new Regex("<span.+?>", RegexOptions.IgnoreCase);
        public static readonly Regex BeginStyleReg = new Regex("<.+?>", RegexOptions.IgnoreCase);
        public static readonly Regex EndStyleReg = new Regex("</.+?>", RegexOptions.IgnoreCase);

        public MHTCell(uint index, string element)
        {
            Index = index;
            if (!string.IsNullOrWhiteSpace(element))
            {
                var value = RemoveStyle(GetValue(element, ValueReg)).HtmlDecode();
                Value = value.Replaces(GetValue(value, SpanReg), "</span>", "</td>", ">", Multipart.SignBlank);
                ColSpan = ToUint(GetValue(element, ColSpanReg).Replaces("colspan=", " "));
                RowSpan = ToUint(GetValue(element, RowSpanReg).Replaces("rowspan=", " "));
            }
        }

        private string RemoveStyle(string element)
        {
            return string.IsNullOrWhiteSpace(element) ? element
                : EndStyleReg.Replace(BeginStyleReg.Replace(element, string.Empty), string.Empty);
        }

        private string GetValue(string element, Regex reg)
        {
            var match = reg.Match(element);
            return match.Success ? match.Value : string.Empty;
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