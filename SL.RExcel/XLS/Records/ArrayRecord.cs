using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class ArrayRecord : Record
    {
        public ushort FirstRowIdx { get; private set; }

        public ushort LastRowIdx { get; private set; }

        public byte FirstColIdx { get; private set; }

        public byte LastColIdx { get; private set; }

        public ushort Options { get; private set; }

        public uint Reserved { get; private set; }

        public byte[] Data { get; private set; }

        public ArrayRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            FirstRowIdx = reader.ReadUInt16();
            LastRowIdx = reader.ReadUInt16();
            FirstColIdx = reader.ReadByte();
            LastColIdx = reader.ReadByte();
            Options = reader.ReadUInt16();
            Reserved = reader.ReadUInt32();
            Data = reader.ReadBytes((int)(stream.Length - stream.Position));
        }
    }
}