using SL.RExcel.XLS.File;
using System.IO;

namespace SL.RExcel.XLS.Records
{
    public enum SheetType : byte
    {
        WorkSheet = 0,
        Chart,
        VBModule
    }

    public enum VisibilityType : byte
    {
        Visible = 0,
        Hidden,
        StrongHidden
    }

    public class BoundSheetRecord : Record
    {
        public string Name { get; private set; }

        public uint BOFPos { get; private set; }

        public SheetType Type { get; private set; }

        public VisibilityType Visibility { get; private set; }

        public BoundSheetRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            BOFPos = reader.ReadUInt32();
            Visibility = (VisibilityType)reader.ReadByte();
            Type = (SheetType)reader.ReadByte();
            byte len = reader.ReadByte();
            Name = reader.ReadPossibleCompressedString(len);
        }
    }
}