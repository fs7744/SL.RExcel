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
                GetNewReader(stream, ref dataStream, ref reader);

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

                HandleChar(stream, ref dataStream, reader, sb, len, ref compressed);
                Strings[i] = sb.ToString();
                long skip = (rtSize * 4) + farEastSize;
                ReadSkip(stream, ref dataStream, ref reader, ref skip);
            }
        }

        private void HandleChar(Stream stream, ref Stream dataStream, BinaryReader reader, StringBuilder sb, ushort len, ref bool compressed)
        {
            sb.Length = 0;
            sb.EnsureCapacity(len);
            for (ushort n = 0; n < len; ++n)
            {
                if (dataStream.Position >= dataStream.Length)
                {
                    dataStream = ReadContinue(stream);
                    compressed = (reader.ReadByte() & 0x01) == 0;
                }

                sb.Append(Convert.ToChar(compressed ? reader.ReadByte() : reader.ReadUInt16()));
            }
        }

        private void ReadSkip(Stream stream, ref Stream dataStream, ref BinaryReader reader, ref long skip)
        {
            while (skip > 0)
            {
                GetNewReader(stream, ref dataStream, ref reader);
                long actualSkip = Math.Min(dataStream.Length - dataStream.Position, skip);
                dataStream.Seek(actualSkip, SeekOrigin.Current);
                skip -= actualSkip;
            }
        }

        private void GetNewReader(Stream stream, ref Stream dataStream, ref BinaryReader reader)
        {
            if (dataStream.Position >= dataStream.Length)
            {
                dataStream = ReadContinue(stream);
                reader = new BinaryReader(dataStream);
            }
        }
    }
}