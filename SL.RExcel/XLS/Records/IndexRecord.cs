using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class IndexRecord : Record
    {
        public uint FirstRow { get; private set; }

        public uint LastRow { get; private set; }

        public uint[] Rows { get; private set; }

        public IndexRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            uint reserved = reader.ReadUInt32();
            FirstRow = reader.ReadUInt32();
            LastRow = reader.ReadUInt32();
            reserved = reader.ReadUInt32();

            //int nrows = (int)(stream.Length - stream.Position) / 4;
            //Rows = new uint[nrows];
            //nrows = 0;
            //while (stream.Position < stream.Length)
            //    Rows[nrows++] = reader.ReadUInt32();
        }
    }
}