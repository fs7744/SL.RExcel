using SL.RExcel.XLS.File;
using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class LabelRecord : RowColXfCellRecord
    {
        //public string Value { get; private set; }

        public LabelRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
            Value = reader.ReadSimpleUnicodeString();
        }
    }
}