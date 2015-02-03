using System.IO;

namespace SL.RExcel.XLS.File
{
    public class DirectorySector : Sector
    {
        public DirectorySectorData[] Entries { get; private set; }

        public DirectorySector(Stream stream)
        {
            Entries = new DirectorySectorData[4];
            for (int i = 0; i < Entries.Length; ++i)
                Entries[i] = new DirectorySectorData(stream);
        }
    }
}