using System.IO;

namespace SL.RExcel.XLS.File
{
    public class DifSector : Sector
    {
        public SectorIndex[] Fats { get; private set; }

        public SectorIndex NextDif { get; private set; }

        public DifSector(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);
            Fats = new SectorIndex[127];
            for (int i = 0; i < Fats.Length; ++i)
                Fats[i] = new SectorIndex(reader.ReadUInt32());

            NextDif = new SectorIndex(reader.ReadUInt32());
        }
    }
}