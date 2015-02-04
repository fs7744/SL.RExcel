using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class NumberRecord : RowColXfCellRecord
    {
        //public double Value { get; private set; }

        public NumberRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
            Value = reader.ReadDouble();
        }
    }
}