using System;
using System.IO;

namespace SL.RExcel.XLS.File
{
    public class DirectorySectorData
    {
        public string Name { get; private set; }

        public ushort NameLength { get; private set; }

        public Stgty Type { get; private set; }

        public DeColor Color { get; private set; }

        public SectorIndex LeftSibling { get; private set; }

        public SectorIndex RightSibling { get; private set; }

        public SectorIndex Child { get; private set; }

        public Guid ID { get; private set; }

        public uint UserFlags { get; private set; }

        public ulong CreateTimeStamp { get; private set; }

        public ulong ModifyTimeStamp { get; private set; }

        public SectorIndex Start { get; private set; }

        public uint Size { get; private set; }

        public ushort PropType { get; private set; }

        public ushort Padding { get; private set; }

        public DirectorySectorData(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);
            SetName(reader);
            Type = (Stgty)reader.ReadByte();
            Color = (DeColor)reader.ReadByte();
            LeftSibling = new SectorIndex(reader.ReadUInt32());
            RightSibling = new SectorIndex(reader.ReadUInt32());
            Child = new SectorIndex(reader.ReadUInt32());
            ID = new Guid(reader.ReadBytes(16));
            UserFlags = reader.ReadUInt32();
            CreateTimeStamp = reader.ReadUInt64();
            ModifyTimeStamp = reader.ReadUInt64();
            Start = new SectorIndex(reader.ReadUInt32());
            Size = reader.ReadUInt32();
            PropType = reader.ReadUInt16();
            Padding = reader.ReadUInt16();
        }

        private void SetName(BinaryReader reader)
        {
            Name = reader.ReadSimpleUnicodeString(64);
            NameLength = reader.ReadUInt16();
            if (NameLength > 0)
            {
                NameLength /= 2;
                --NameLength;
                Name = Name.Substring(0, NameLength);
            }
            else
                Name = string.Empty;
        }
    }
}