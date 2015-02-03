using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class BlankRecord : RowColXfCellRecord
    {
        public BlankRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
        }
    }
}