using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class MergeCellsRecord : Record
    {
        public List<MergeCell> MergeCells { get; private set; }

        public MergeCellsRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            var count = reader.ReadUInt16();
            MergeCells = new List<MergeCell>();
            for (int i = 0; i < count; i++)
            {
                MergeCells.Add(new MergeCell()
                {
                    FirstRow = reader.ReadUInt16(),
                    LastRow = reader.ReadUInt16(),
                    FirstCol = reader.ReadUInt16(),
                    LastCol = reader.ReadUInt16(),
                });
            }
        }
    }
}