using System.Collections.Generic;
using System.Linq;

namespace SL.RExcel.XLS.File
{
    public static class XLSDirectoryFactory
    {
        public static XLSDirectory CreateEntry(List<DirectorySectorData> dirs, List<Sector> sectors, List<SectorIndex> index)
        {
            var first = dirs.FirstOrDefault();
            return CreateEntry(first, dirs, sectors, index);
        }

        public static XLSDirectory CreateEntry(DirectorySectorData root, List<DirectorySectorData> dirs, List<Sector> sectors, List<SectorIndex> index)
        {
            var result = GetEntry(root, sectors, index);
            if (!root.LeftSibling.IsEof)
                result.LeftSibling = CreateEntry(dirs[root.LeftSibling.ToInt()],
                    dirs,
                    sectors,
                    index);

            if (!root.RightSibling.IsEof)
                result.RightSibling = CreateEntry(dirs[root.RightSibling.ToInt()],
                    dirs,
                    sectors,
                    index);

            if (!root.Child.IsEof)
                result.Child = CreateEntry(dirs[root.Child.ToInt()],
                    dirs,
                    sectors,
                    index);
            return result;
        }

        private static XLSDirectory GetEntry(DirectorySectorData entry, List<Sector> sectors, List<SectorIndex> index)
        {
            XLSDirectory result = null;
            switch (entry.Type)
            {
                case Stgty.Storage:
                    result = new XLSDirectory(entry.Name);
                    break;

                case Stgty.Stream:
                    result = new XLSStreamDirectory(entry.Name, entry.Size, entry.Start, sectors, index);
                    break;

                default:
                    result = new XLSGenericDirectory(entry.Name, entry.Type);
                    break;
            }
            return result;
        }
    }
}