using SL.RExcel.XLS.File;
using System.IO;

namespace SL.RExcel.XLS
{
    public class BIFFData
    {
        public const int MinSize = 4;

        public byte[] Data { get; private set; }

        public ushort ID { get; private set; }

        public ushort Length { get; private set; }

        public BIFFData(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);
            ID = reader.ReadUInt16();
            Length = reader.ReadUInt16();
            Data = reader.ReadBytes(Length);
        }

        public Stream ToStream()
        {
            return Data.ToStream();
        }
    }
}