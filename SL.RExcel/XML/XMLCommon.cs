using System.Xml.Linq;

namespace SL.RExcel.XML
{
    public class XMLCommon
    {
        public static readonly XNamespace SS = XNamespace.Get("urn:schemas-microsoft-com:office:spreadsheet");
        public static readonly XName Worksheet = SS + @"Worksheet";
        public static readonly XName Name = SS + @"Name";
        public static readonly XName Row = SS + @"Row";
        public static readonly XName Cell = SS + @"Cell";
        public static readonly XName MergeAcross = SS + @"MergeAcross";
        public static readonly XName Data = SS + @"Data";
        public static readonly XName MergeDown = SS + @"MergeDown";
        public static readonly XName Index = SS + @"Index";
    }
}