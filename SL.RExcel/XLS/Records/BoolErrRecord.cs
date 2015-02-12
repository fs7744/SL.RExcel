using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class BoolErrRecord : RowColXfCellRecord
    {
        public static readonly Dictionary<byte, string> BoolErrMap
            = new Dictionary<byte, string>()
        {
            {0x00,"#NULL!"},
            {0x07,"#DIV/0"},
            {0x0f,"#VALUE!"},
            {0x17,"#REF!"},
            {0x1d,"#NAME?"},
            {0x24,"#NUM!"},
            {0x2a,"#N/A"},
        };

        public bool Error { get; private set; }

        public BoolErrRecord(BIFFData data, Stream stream)
        {
            BinaryReader reader = new BinaryReader(data.ToStream());
            SetRowColXf(reader);
            byte boolErr = reader.ReadByte();
            Error = reader.ReadByte() == 1;
            if (Error)
                SetBoolErrValue(boolErr);
            else
                Value = boolErr == 1;
        }

        private void SetBoolErrValue(byte boolErr)
        {
            string result = string.Empty;
            BoolErrMap.TryGetValue(boolErr, out result);
            Value = result;
        }
    }
}