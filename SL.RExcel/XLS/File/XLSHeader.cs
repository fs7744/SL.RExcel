using System;
using System.IO;

namespace SL.RExcel.XLS.File
{
    public class XLSHeader
    {
        public const ulong MagicNumber = 0xe11ab1a1e011cfd0;

        public Guid ID { get; private set; }

        public ushort MinorVer { get; private set; }

        public ushort DllVer { get; private set; }

        public ushort ByteOrder { get; private set; }

        public ushort Shift { get; private set; }

        public ushort MiniShift { get; private set; }

        public ushort Reserved { get; private set; }

        public uint Reserved1 { get; private set; }

        public uint Reserved2 { get; private set; }

        public uint FatCount { get; private set; }

        public SectorIndex DirStart { get; private set; }

        public uint Signature { get; private set; }

        public uint MiniSectorCutoff { get; private set; }

        public SectorIndex MiniFatStart { get; private set; }

        public uint MiniFatCount { get; private set; }

        public SectorIndex DifStart { get; private set; }

        public uint DifCount { get; private set; }

        public SectorIndex[] Fats { get; private set; }

        public XLSHeader(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);
            CheckFormart(reader.ReadUInt64());

            ID = new Guid(reader.ReadBytes(16));
            MinorVer = reader.ReadUInt16();
            DllVer = reader.ReadUInt16();
            ByteOrder = reader.ReadUInt16();
            Shift = reader.ReadUInt16();
            MiniShift = reader.ReadUInt16();
            Reserved = reader.ReadUInt16();
            Reserved1 = reader.ReadUInt32();
            Reserved2 = reader.ReadUInt32();
            FatCount = reader.ReadUInt32();
            DirStart = new SectorIndex(reader.ReadUInt32());
            Signature = reader.ReadUInt32();
            MiniSectorCutoff = reader.ReadUInt32();
            MiniFatStart = new SectorIndex(reader.ReadUInt32());
            MiniFatCount = reader.ReadUInt32();
            DifStart = new SectorIndex(reader.ReadUInt32());
            DifCount = reader.ReadUInt32();
            Fats = new SectorIndex[109];
            for (int i = 0; i < Fats.Length; i++)
            {
                Fats[i] = new SectorIndex(reader.ReadUInt32());
            }
        }

        private void CheckFormart(ulong sign)
        {
            if (sign != MagicNumber)
                throw new NotSupportedException("Invalid header MagicNumber.");
        }
    }
}