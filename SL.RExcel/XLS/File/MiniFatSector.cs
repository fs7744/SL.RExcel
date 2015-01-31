using System.IO;

namespace SL.RExcel.XLS.File
{
    public class MiniFatSector : Sector
    {
        public SectorIndex[] Fats { get; private set; }

        public MiniFatSector(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);

            for (int i = 0; i < Fats.Length; ++i)
                Fats[i] = new SectorIndex(reader.ReadUInt32());
        }
    }
}