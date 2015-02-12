using SL.RExcel.MHT;
using SL.RExcel.XLS;
using SL.RExcel.XLSX;
using SL.RExcel.XML;
using System;
using System.IO;

namespace SL.RExcel
{
    public static class ExcelHelper
    {
        private static Func<Stream, IWorkBook>[] m_Funcs;

        static ExcelHelper()
        {
            m_Funcs = new Func<Stream, IWorkBook>[4]
            {
                s => new XLSXWorkBook(s),
                s => new MHTWorkBook(s),
                s => new XMLWorkBook(s),
                s => new XLSWorkBook(s)
            };
        }

        public static IWorkBook Open(Stream stream)
        {
            IWorkBook result = null;
            try
            {
                result = TryOpenByAllWay(stream);
            }
            finally
            {
                if (stream.CanRead) stream.Close();
            }
            return result;
        }

        private static IWorkBook TryOpenByAllWay(Stream stream)
        {
            IWorkBook result = null;
            int index = m_Funcs.Length - 1;
            while (index >= 0)
            {
                result = TryOpen(stream, index);
                if (result != null)
                    break;
                index--;
            }
            return result;
        }

        private static IWorkBook TryOpen(Stream stream, int index)
        {
            IWorkBook result = null;
            try
            {
                result = m_Funcs[index](stream);
            }
            catch (NotSupportedException ex)
            {
                if (index < 0)
                    throw ex;
            }
            return result;
        }
    }
}