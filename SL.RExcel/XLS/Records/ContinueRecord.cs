using System.IO;

namespace SL.RExcel.XLS.Records
{
    public class ContinueRecord : Record
    {
        public byte[] Data { get; private set; }

        public ContinueRecord(BIFFData data, Stream stream)
        {
            Data = data.Data;
        }
    }
}