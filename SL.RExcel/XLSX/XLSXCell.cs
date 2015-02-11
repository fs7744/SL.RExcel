namespace SL.RExcel.XLSX
{
    public class XLSXCell : ICell
    {
        public XLSXCell(object value)
        {
            Value = value;
        }

        public object Value { get; private set; }

        public string GetStringValue()
        {
            return Value.ToString();
        }
    }
}