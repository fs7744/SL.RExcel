using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class DbCellRecord : Record
    {
        public uint RowOffset { get; private set; }

        public DbCellRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            RowOffset = reader.ReadUInt32();

            //int noffsets = (int)(stream.Length - stream.Position) / 2;
            //_streamOffsets = new ushort[noffsets];
            //noffsets = 0;
            //while (stream.Position < stream.Length)
            //    _streamOffsets[noffsets++] = reader.ReadUInt16();
        }
    }
}