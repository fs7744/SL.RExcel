using System.IO;
using System.Text;

namespace SL.RExcel.MHT
{
    public class Part
    {
        public string Location { get; private set; }

        public string TransferEncoding { get; private set; }

        public string Type { get; private set; }

        public StringBuilder m_SB { get; private set; }

        public Part(StreamReader reader)
        {
            m_SB = new StringBuilder();
            Location = GetValue(reader, Multipart.ContentLocation);
            TransferEncoding = GetValue(reader, Multipart.ContentTransferEncoding);
            Type = GetValue(reader, Multipart.ContentType);
            while (reader.IsNotEnd())
            {
                var line = reader.NextFullLine();
                if (line.Contains(Multipart.NextPart))
                    break;
                m_SB.Append(line.Replace(Multipart.Sign3D, Multipart.SignEqual));
            }
        }

        private string GetValue(StreamReader reader, string sign)
        {
            var line = reader.NextFullLine();
            return string.IsNullOrWhiteSpace(line) ? string.Empty
                : line.Replace(sign, string.Empty);
        }

        public bool IsNotEmpty()
        {
            return m_SB.Length > 0;
        }

        public override string ToString()
        {
            return m_SB.ToString();
        }
    }
}