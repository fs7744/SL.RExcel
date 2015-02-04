using SL.RExcel.XLS.Records;

namespace SL.RExcel.XLS
{
    public class XLSCell : ICell
    {
        public XLSCell(CellRecord cell)
        {
            Value = cell.Value;
        }

        public object Value { get; private set; }

        public string GetFormattedValue()
        {
            return Value.ToString();
        }
    }
}