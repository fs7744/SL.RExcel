using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class CellRecord : Record
    {
        public ushort Row { get; protected set; }

        public object Value { get; protected set; }

        protected void SetRow(BinaryReader reader)
        {
            Row = reader.ReadUInt16();
        }
    }
}