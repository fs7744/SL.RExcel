using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class MulRkRecord : MultipleColCellRecord
    {
        public double[] RKs { get; private set; }

        public MulRkRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            reader = SetRowColInfo(reader);
            RKs = new double[reader.BaseStream.Length / 6];
            for (int n = 0; n < RKs.Length; ++n)
            {
                ushort xf = reader.ReadUInt16();
                int rk = reader.ReadInt32();
                RKs[n] = RkRecord.Convet(rk);
            }
        }

        public override object GetValue(ushort col)
        {
            return RKs[col - FirstCol];
        }
    }
}