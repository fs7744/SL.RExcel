using System;
using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.MHT
{
    public class Multipart
    {
        public const string Header = "MIME-VERSION";
        public const string NextPart = "NextPart";
        public const string Sign3D = "=3D";
        public const string SignEqual = "=";
        public const string SignBlank = @"&#12288;";
   
        public const string ContentLocation = "Content-Location: ";
        public const string ContentTransferEncoding = "Content-Transfer-Encoding: ";
        public const string ContentType = "Content-Type: ";

        public List<Part> Parts { get; private set; }

        public Multipart(Stream stream)
        {
            stream.Position = 0L;
            StreamReader reader = new StreamReader(stream);
            CheckHeader(reader);
            SkipToPartBegin(reader);
            SkipToPartBegin(reader);
            Parts = new List<Part>();
            while (reader.IsNotEnd())
            {
                var part = new Part(reader);
                if (part.IsNotEmpty())
                    Parts.Add(part);
            }
        }

        private void CheckHeader(StreamReader reader)
        {
            var header = reader.NextFullLine();
            if (string.IsNullOrWhiteSpace(header)
                || !header.ToUpper().Contains(Header))
                throw new NotSupportedException("The file is not mht file");
        }

        private void SkipToPartBegin(StreamReader reader)
        {
            string line = string.Empty;
            do
            {
                line = reader.NextFullLine();
            }
            while (!string.IsNullOrWhiteSpace(line)
                && !line.Contains(NextPart));
        }
    }
}