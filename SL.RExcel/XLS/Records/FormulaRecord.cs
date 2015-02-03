using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class FormulaRecord : RowColXfCellRecord
    {
        public FormulaRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
        }
    }
}