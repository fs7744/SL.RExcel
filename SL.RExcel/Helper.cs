using SL.RExcel.XLSX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Browser;
using System.Xml;
using System.Xml.Linq;

namespace SL.RExcel
{
    public static class Helper
    {
        private static Regex m_Reg = new Regex("^[A-Z]*");

        public static bool IsEmpty<T>(this IEnumerable<T> list)
        {
            return list == null || !list.GetEnumerator().MoveNext();
        }

        public static bool EqualsIgnoreCase(this string str, string value)
        {
            if (string.IsNullOrWhiteSpace(str))
            {
                return string.IsNullOrWhiteSpace(value);
            }

            return str.TrimString().Equals(value.TrimString(), StringComparison.OrdinalIgnoreCase);
        }

        public static string TrimString(this string str)
        {
            return str == null ? string.Empty : str.Trim();
        }

        public static XElement GetXLSXPart(this UnZipper unzip, string xmlName)
        {
            XElement partElement = null;
            using (Stream partStream = unzip.GetFileStream(xmlName))
            {
                partElement = XElement.Load(XmlReader.Create(partStream));
            }
            return partElement;
        }

        public static int[] ConvertIndex(this string index)
        {
            var col = m_Reg.Matches(index)[0].Value;
            var cols = col.ToList();
            int colInt = 0;
            for (int i = cols.Count - 1, n = 0; i >= 0; i--, n++)
            {
                colInt += ((int)cols[i] - 64) * (int)Math.Pow(26, n);
            }
            var row = int.Parse(index.Replace(col, "")) - 1;
            return new int[2] { colInt, row };
        }

        public static string NextFullLine(this StreamReader reader)
        {
            string result = string.Empty;
            while (reader.IsNotEnd() && string.IsNullOrWhiteSpace(result))
            {
                result = reader.ReadLine();
            }
            return result;
        }

        public static bool IsNotEnd(this StreamReader reader)
        {
            return !reader.EndOfStream;
        }

        public static string Replaces(this string element, params string[] replaces)
        {
            if (string.IsNullOrEmpty(element)) return element;
            var result = element;
            foreach (var item in (replaces ?? new string[1]).Where(i => !string.IsNullOrEmpty(i)))
            {
                result = result.Replace(item, string.Empty);
            }
            return result;
        }

        public static string HtmlDecode(this string value)
        {
            return string.IsNullOrWhiteSpace(value) ? value : HttpUtility.HtmlDecode(value);
        }
    }
}
