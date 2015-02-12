using System;
using System.IO;
using System.Text;

namespace SL.RExcel.XLS.File
{
    public static class ReaderHelper
    {
        public static string ReadSimpleUnicodeString(this BinaryReader reader)
        {
            ushort length = reader.ReadUInt16();
            return ReadSimpleUnicodeString(reader, length);
        }

        public static string ReadSimpleUnicodeString(this BinaryReader reader, int length)
        {
            return Encoding.Unicode.GetString(reader.ReadBytes(length), 0, length);
        }

        public static Stream ToStream(this byte[] bytes)
        {
            return new MemoryStream(bytes);
        }

        public static string ConvertFromUtf32(int utf32)
        {
            CheckUTF32(utf32);
            if (utf32 < 0x10000)
                return new string((char)utf32, 1);
            utf32 -= 0x10000;
            return new string(new[] {(char) ((utf32 >> 10) + 0xD800),
                                (char) (utf32 % 0x0400 + 0xDC00)});
        }

        private static void CheckUTF32(int utf32)
        {
            if (utf32 < 0 || utf32 > 0x10FFFF)
                throw new ArgumentOutOfRangeException("utf32", "The argument must be from 0 to 0x10FFFF.");
            CheckSurrogatePairRange(utf32);
        }

        private static void CheckSurrogatePairRange(int utf32)
        {
            if (0xD800 <= utf32 && utf32 <= 0xDFFF)
                throw new ArgumentOutOfRangeException("utf32", "The argument must not be in surrogate pair range.");
        }

        public static string ReadPossibleCompressedString(this BinaryReader reader, int length)
        {
            byte options = reader.ReadByte();
            bool compressed = (options & 0x01) == 0;

            if (compressed)
            {
                byte[] data = reader.ReadBytes(length);

                StringBuilder sb = new StringBuilder();
                foreach (byte b in data)
                {
                    string s = ConvertFromUtf32(b);

                    sb.Append(s);
                }

                return sb.ToString();
            }
            else
            {
                length *= 2;

                byte[] data = reader.ReadBytes(length);

                string retVal = Encoding.Unicode.GetString(data, 0, data.Length);

                return retVal;
            }
        }

        public static string ReadComplexString(this BinaryReader reader)
        {
            ushort len = reader.ReadUInt16();
            return ReadComplexString(reader, len);
        }

        public static string ReadComplexString(this BinaryReader reader, int length)
        {
            byte options = reader.ReadByte();
            bool compressed = (options & 0x01) == 0;

            Encoding enc;
            if (compressed)
                enc = Encoding.UTF8;
            else
            {
                length *= 2;
                enc = Encoding.Unicode;
            }

            byte[] data = reader.ReadBytes(length);

            string retVal = enc.GetString(data, 0, data.Length);

            return retVal;
        }
    }
}