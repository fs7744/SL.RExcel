using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SL.RExcel.XLS.File
{
    public class XLSFile
    {
        public const int MaxSectorIndex = 128;
        private XLSDirectory m_Root;

        public XLSFile(Stream stream)
        {
            var header = new XLSHeader(stream);
            var sectors = GetSectors(stream);
            List<SectorIndex> index = new List<SectorIndex>((int)(MaxSectorIndex * header.FatCount));
            Set109Fats(header, sectors, index);
            SetRemainFats(header, sectors, index);
            SetMiniFats(header, sectors, index);
            var dirs = SetDirs(header, sectors, index);
            m_Root = XLSDirectoryFactory.CreateEntry(dirs, sectors, index);
            stream.Close();
        }

        private List<DirectorySectorData> SetDirs(XLSHeader header, List<Sector> sectors, List<SectorIndex> index)
        {
            var dirs = new List<DirectorySectorData>();
            for (SectorIndex dirSect = header.DirStart;
                !dirSect.IsEndOfChain;
                dirSect = index[dirSect.ToInt()])
            {
                DirectorySector dir = new DirectorySector(sectors[dirSect.ToInt()].ToStorage().ToStream());

                foreach (DirectorySectorData entry in dir.Entries)
                    dirs.Add(entry);
                sectors[dirSect.ToInt()] = dir;
            }
            return dirs;
        }

        private static void SetMiniFats(XLSHeader header, List<Sector> sectors, List<SectorIndex> index)
        {
            SectorIndex miniFatSect;
            int miniFatCount;
            for (miniFatSect = header.MiniFatStart, miniFatCount = 0;
                !miniFatSect.IsEndOfChain && miniFatCount < header.MiniFatCount;
                miniFatSect = index[miniFatSect.ToInt()], ++miniFatCount)
            {
                MiniFatSector miniFat = new MiniFatSector(sectors[miniFatSect.ToInt()].ToStorage().ToStream());
                sectors[miniFatSect.ToInt()] = miniFat;
            }
        }

        private void SetRemainFats(XLSHeader header, List<Sector> sectors, List<SectorIndex> index)
        {
            int difCount;
            SectorIndex difIndex;
            for (difIndex = header.DifStart, difCount = 0;
                !difIndex.IsEndOfChain && difCount < header.DifCount;
                ++difCount)
            {
                DifSector dif = new DifSector(sectors[difIndex.ToInt()].ToStorage().ToStream());
                sectors[difIndex.ToInt()] = dif;

                foreach (var item in dif.Fats.Where(i => !i.IsFree))
                {
                    var i = item.ToInt();
                    FatSector fat = new FatSector(sectors[i].ToStorage().ToStream());
                    index.AddRange(fat.Fats);
                    sectors[i] = fat;
                }
            }
        }

        private void Set109Fats(XLSHeader header, List<Sector> sectors, List<SectorIndex> index)
        {
            foreach (var item in header.Fats)
            {
                var i = item.ToInt();
                if (!item.IsFree)
                {
                    FatSector fat = new FatSector(sectors[i].ToStorage().ToStream());
                    index.AddRange(fat.Fats);
                    sectors[i] = fat;
                }
            }
        }

        private List<Sector> GetSectors(Stream stream)
        {
            var result = new List<Sector>((int)(stream.Length / Storage.StorageSize));
            while (stream.Position < stream.Length)
            {
                result.Add(new Storage(stream));
            }
            return result;
        }

        public Stream OpenStream(string name)
        {
            var entry = m_Root.Find(name);
            if (entry != null && entry is XLSStreamDirectory)
                return new MemoryStream(((XLSStreamDirectory)entry).Data);

            throw new IOException("Stream [" + name + "] was not found.");
        }
    }
}