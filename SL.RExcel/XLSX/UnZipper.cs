using ICSharpCode.SharpZipLib.Zip;
using System.Collections.Generic;
using System.IO;

namespace SL.RExcel.XLSX
{
    public class UnZipper
    {
        private ZipFile m_ZipFile;

        public UnZipper(Stream zipFileStream)
        {
            this.m_ZipFile = new ZipFile(zipFileStream);
        }

        public Stream GetFileStream(string filename)
        {
            var entry = m_ZipFile.GetEntry(filename);
            return m_ZipFile.GetInputStream(entry);
        }

        public IEnumerable<string> GetFileNamesInZip()
        {
            List<string> names = new List<string>();
            foreach (ZipEntry item in m_ZipFile)
            {
                names.Add(item.Name);
            }
            return names;
        }
    }
}