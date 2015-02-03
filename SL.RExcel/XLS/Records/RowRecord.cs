using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class RowRecord : Record
    {
        public ushort RowNumber { get; private set; }

        public ushort FirstCol { get; private set; }

        public ushort LastCol { get; private set; }

        public ushort RowHeight { get; private set; }

        public ushort Optimizer { get; private set; }

        public ushort Options { get; private set; }

        public ushort XF { get; private set; }

        public RowRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            RowNumber = reader.ReadUInt16();
            FirstCol = reader.ReadUInt16();
            LastCol = reader.ReadUInt16();
            RowHeight = reader.ReadUInt16();
            Optimizer = reader.ReadUInt16();
            Options = reader.ReadUInt16();
            XF = reader.ReadUInt16();
        }
    }
}