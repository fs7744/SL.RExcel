using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class SingleColCellRecord : CellRecord
    {
        public ushort Col { get; protected set; }

        protected void SetRowCol(BinaryReader reader)
        {
            SetRow(reader);
            Col = reader.ReadUInt16();
        }
    }
}