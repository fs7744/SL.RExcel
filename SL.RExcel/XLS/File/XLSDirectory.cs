using System;
using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.XLS.File
{
    public class XLSDirectory
    {
        public XLSDirectory LeftSibling { get; set; }

        public XLSDirectory RightSibling { get; set; }

        public XLSDirectory Child { get; set; }

        public string Name { get; set; }

        public XLSDirectory(string name)
        {
            Name = name;
        }

        public XLSDirectory Find(string name)
        {
            XLSDirectory child = Child;
            while (child != null)
            {
                int c = Compare(name, child.Name);
                if (c == 0)
                    break;
                child = GetNext(c, child);
            }

            return child;
        }

        private XLSDirectory GetNext(int c, XLSDirectory child)
        {
            XLSDirectory result = null;
            if (c < 0)
                result = child.LeftSibling;
            else if (c > 0)
                result = child.RightSibling;
            return result;
        }

        private int Compare(string l, string r)
        {
            if (l.Length < r.Length)
                return -1;
            else if (l.Length > r.Length)
                return 1;
            else
                return string.Compare(l, r, StringComparison.OrdinalIgnoreCase);
        }
    }

    public class XLSGenericDirectory : XLSDirectory
    {
        public Stgty Type { get; private set; }

        public XLSGenericDirectory(string name, Stgty type)
            : base(name)
        {
            Type = type;
        }
    }

    public class XLSStreamDirectory : XLSDirectory
    {
        public byte[] Data { get; private set; }

        public long Length { get; private set; }

        public XLSStreamDirectory(string name, long length, SectorIndex dataOffset,
            List<Sector> sectors, List<SectorIndex> index)
            : base(name)
        {
            Length = length;
            Data = new byte[length];
            SetData(length, dataOffset, sectors, index);
        }

        private void SetData(long length, SectorIndex dataOffset, List<Sector> sectors, List<SectorIndex> index)
        {
            if (!dataOffset.IsEndOfChain)
            {
                int left = (int)length;
                MemoryStream stream = new MemoryStream(Data);
                SectorIndex sect = dataOffset;

                do
                {
                    try
                    {
                        Storage sector = (Storage)sectors[sect.ToInt()];

                        int toWrite = Math.Min(sector.Data.Length, left);
                        stream.Write(sector.Data, 0, toWrite);
                        left -= toWrite;

                        sect = index[sect.ToInt()];
                    }
                    catch (Exception err)
                    {
                        return;
                    }
                } while (!sect.IsEndOfChain);
            }
        }
    }
}