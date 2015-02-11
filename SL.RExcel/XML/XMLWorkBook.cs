using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace SL.RExcel.XML
{
    public class XMLWorkBook : IWorkBook
    {
        public IWorksheet[] Worksheets { get; private set; }

        public int Count
        {
            get { return Worksheets.Length; }
        }

        public XMLWorkBook(Stream stream)
        {
            stream.Position = 0L;
            var exlement = GetXML(stream);
            Worksheets = exlement.Descendants(XMLCommon.Worksheet).Select(i => new XMLWorksheet(i)).ToArray();
            stream.Close();
        }

        private XElement GetXML(Stream stream)
        {
            XElement result = null;
            try
            {
                result = XElement.Load(XmlReader.Create(stream));
            }
            catch (Exception ex)
            {
                throw new NotSupportedException("It is not xml file.", ex);
            }
            return result;
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