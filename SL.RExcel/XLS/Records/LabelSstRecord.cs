using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class LabelSstRecord : RowColXfCellRecord
    {
        public string Value { get; private set; }

        public uint SSTIndex { get; private set; }

        public LabelSstRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
            SSTIndex = reader.ReadUInt32();
        }

        public void SetValue(SstRecord sst)
        {
            Value = sst.Strings[SSTIndex];
        }
    }
}