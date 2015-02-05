using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class MultipleColCellRecord : CellRecord
    {
        public ushort FirstCol { get; private set; }

        public ushort LastCol { get; private set; }

        protected BinaryReader SetRowColInfo(BinaryReader reader)
        {
            SetRow(reader);
            FirstCol = reader.ReadUInt16();

            byte[] inBetween = reader.ReadBytes((int)(reader.BaseStream.Length - 6));

            LastCol = reader.ReadUInt16();

            return new BinaryReader(new MemoryStream(inBetween));
        }

        public virtual object GetValue(ushort col)
        {
            return null;
        }
    }
}