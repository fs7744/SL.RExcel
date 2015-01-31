using System.IO;

namespace SL.RExcel.XLS.File
{
    public class Storage : Sector
    {
        public const int StorageSize = 512;

        public byte[] Data { get; private set; }

        public Storage(Stream stream)
        {
            Data = new byte[StorageSize];
            int read = stream.Read(Data, 0, StorageSize);
            if (read < StorageSize)
                throw new IOException("The data size was shorted than expected StorageSize.");
        }

        public Stream ToStream()
        {
            return new MemoryStream(Data);
        }
    }
}