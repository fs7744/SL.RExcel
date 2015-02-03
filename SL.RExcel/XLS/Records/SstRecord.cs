using SL.RExcel.XLS.File;
using System;
using System.IO;
using System.Text;

namespace SL.RExcel.XLS.Records
{
    public class SstRecord : Record
    {
        public string[] Strings { get; private set; }

        private Stream ReadContinue(Stream stream)
        {
            BIFFData biff = new BIFFData(stream);
            return new ContinueRecord(biff, stream).Data.ToStream();
        }

        public SstRecord(BIFFData data, Stream stream)
        {
            var dataStream = data.ToStream();
            BinaryReader reader = new BinaryReader(dataStream);
            uint totalStrings = reader.ReadUInt32();
            uint totalUniqueStrings = reader.ReadUInt32();
            Strings = new string[totalUniqueStrings];
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < totalUniqueStrings; ++i)
            {
                if (dataStream.Position >= dataStream.Length)
                {
                    dataStream = ReadContinue(stream);
                    reader = new BinaryReader(dataStream);
                }

                ushort len = reader.ReadUInt16();
                byte options = reader.ReadByte();
                bool compressed = (options & 0x01) == 0;
                bool farEast = (options & 0x04) != 0;
                bool richText = (options & 0x08) != 0;
                ushort rtSize = 0;
                uint farEastSize = 0;

                if (richText)
                    rtSize = reader.ReadUInt16();
                if (farEast)
                    farEastSize = reader.ReadUInt32();

                sb.Length = 0;
                sb.EnsureCapacity(len);
                for (ushort n = 0; n < len; ++n)
                {
                    if (dataStream.Position >= dataStream.Length)
                    {
                        dataStream = ReadContinue(stream);
                        compressed = (reader.ReadByte() & 0x01) == 0;
                    }

                    if (compressed)
                        sb.Append(Convert.ToChar(reader.ReadByte()));
                    else
                        sb.Append(Convert.ToChar(reader.ReadUInt16()));
                }

                Strings[i] = sb.ToString();

                long skip = (rtSize * 4) + farEastSize;

                while (skip > 0)
                {
                    if (dataStream.Position >= dataStream.Length)
                    {
                        dataStream = ReadContinue(stream);
                        reader = new BinaryReader(dataStream);
                    }

                    long actualSkip = Math.Min(dataStream.Length - dataStream.Position, skip);

                    dataStream.Seek(actualSkip, SeekOrigin.Current);

                    skip -= actualSkip;
                }
            }
        }
    }
}