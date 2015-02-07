using System.Xml.Linq;

namespace SL.RExcel.XLSX
{
    public class XLSXCommon
    {
        public static readonly XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        public const string Workbook = @"xl/workbook.xml";
        public const string Sheet = @"xl/worksheets/sheet{0}.xml";
        public const string SharedStrings = @"xl/sharedStrings.xml";
        public const string XML_T = "t";
        public const string XML_Row = "row";
        public const string XML_C = "c";
        public const string XML_V = "v";
        public const string XML_S = "s";
        public const string XML_R = "r";
        public const string XML_SI = "si";
        public const string XML_Sheet = "sheet";
        public const string XML_Name = "name";
        public const string XML_SheetId = "sheetId";
        public const string XML_MergeCell = "mergeCell";
        public const string XML_Ref = "ref";
        public static readonly char[] RefSplit = { ':' };
    }
}