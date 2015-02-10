using SL.RExcel.XLS;
using SL.RExcel.XLSX;
using System;
using System.IO;

namespace SL.RExcel
{
    public static class ExcelHelper
    {
        public static IWorkBook Open(Stream stream)
        {
            IWorkBook result = null;
            try
            {
                result = new XLSWorkBook(stream);
            }
            catch (NotSupportedException)
            {
                stream.Position = 0L;
                result = new XLSXWorkBook(stream);
            }
            finally
            {
                if (stream.CanRead)
                    stream.Close();
            }
            return result;
        }
    }
}