namespace SL.RExcel.XLS
{
    public class XLSCell : ICell
    {
        public XLSCell(object value)
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