using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class RowColXfCellRecord : SingleColCellRecord
    {
        public ushort XF { get; protected set; }

        protected void SetRowColXf(BinaryReader reader)
        {
            SetRowCol(reader);
            XF = reader.ReadUInt16();
        }
    }
}