using System.IO;

namespace SL.RExcel.XLS.File
{
    public class FatSector : Sector
    {
        public SectorIndex[] Fats { get; private set; }

        public FatSector(Stream stream)
        {
            Fats = new SectorIndex[XLSFile.MaxSectorIndex];
            BinaryReader reader = new BinaryReader(stream);
            for (int i = 0; i < Fats.Length; ++i)
                Fats[i] = new SectorIndex(reader.ReadUInt32());
        }
    }
}