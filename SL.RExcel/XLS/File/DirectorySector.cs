using System.IO;

namespace SL.RExcel.XLS.File
{
    public class DirectorySector : Sector
    {
        public XLSDirectory[] Entries { get; private set; }

        public DirectorySector(Stream stream)
        {
            Entries = new XLSDirectory[4];
            for (int i = 0; i < Entries.Length; ++i)
                Entries[i] = new XLSDirectory(stream);
        }
    }
}