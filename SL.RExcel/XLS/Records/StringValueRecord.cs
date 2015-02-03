using SL.RExcel.XLS.File;
using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class StringValueRecord : Record
    {
        public string Value { get; private set; }

        public StringValueRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            Value = reader.ReadComplexString();
        }
    }
}