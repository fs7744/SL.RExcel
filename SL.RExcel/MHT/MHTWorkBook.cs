using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SL.RExcel.MHT
{
    public class MHTWorkBook : IWorkBook
    {
        public const string SheetsReg = "<x:ExcelWorksheet>.+?</x:ExcelWorksheet>";
        public const string SheetNameReg = "<x:Name>.+?</x:Name>";
        public const string SheetPartReg = "HRef=\".+?\"";

        public IWorksheet[] Worksheets { get; private set; }

        public int Count
        {
            get { return Worksheets.Length; }
        }

        public MHTWorkBook(Stream stream)
        {
            var multipart = new Multipart(stream);
            Worksheets = CreateSheets(multipart);
        }

        private IWorksheet[] CreateSheets(Multipart multipart)
        {
            return GetSheetMaps(multipart).Select(i => new MHTWorksheet(i.Key,
                    multipart.Parts.FirstOrDefault(j => j.Location.Contains(i.Value)))).ToArray();
        }

        private KeyValuePair<string, string>[] GetSheetMaps(Multipart multipart)
        {
            KeyValuePair<string, string>[] result = null;
            var reg = new Regex(SheetsReg, RegexOptions.IgnoreCase);
            foreach (var part in multipart.Parts)
            {
                var matches = reg.Matches(part.ToString());
                if (matches.Count > 0)
                {
                    result = matches.Cast<Match>().Select(i => GetMap(i.Value)).ToArray();
                    break;
                }
            }
            return result;
        }

        private KeyValuePair<string, string> GetMap(string xml)
        {
            var reg = new Regex(SheetNameReg, RegexOptions.IgnoreCase).Match(xml);
            var key = reg.Success ? reg.Value.Replace("<x:Name>", string.Empty)
                                            .Replace("</x:Name>", string.Empty)
                                  : string.Empty;
            reg = new Regex(SheetPartReg, RegexOptions.IgnoreCase).Match(xml);
            var value = reg.Success ? reg.Value.Replace("HRef=", string.Empty)
                                            .Replace("\"", string.Empty)
                                    : string.Empty;
            return new KeyValuePair<string, string>(key, value.HtmlDecode());
        }

        public IWorksheet GetSheetByName(string name)
        {
            return Worksheets.FirstOrDefault(i => i.Name == name);
        }

        public IWorksheet GetSheetByIndex(int index)
        {
            return Worksheets[index];
        }

        public IEnumerable<string> GetAllSheetNames()
        {
            return Worksheets.Select(i => i.Name);
        }
    }
}