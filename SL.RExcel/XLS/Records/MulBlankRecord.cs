using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class MulBlankRecord : MultipleColCellRecord
    {
        public ushort[] XFIndex;

        public MulBlankRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
        }
    }
}