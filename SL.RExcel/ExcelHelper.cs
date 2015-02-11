using SL.RExcel.XLS;
using SL.RExcel.XLSX;
using SL.RExcel.XML;
using System;
using System.Collections.Generic;
using System.IO;

namespace SL.RExcel
{
    public static class ExcelHelper
    {
        private static Func<Stream, IWorkBook>[] m_Funcs;

        static ExcelHelper()
        {
            var list = new List<Func<Stream, IWorkBook>>()
            {
                s => new XLSXWorkBook(s),
                s => new XMLWorkBook(s),
                s => new XLSWorkBook(s)
            };
            m_Funcs = list.ToArray();
        }

        public static IWorkBook Open(Stream stream)
        {
            IWorkBook result = null;
            try
            {
                int index = m_Funcs.Length - 1;
                while (index >= 0)
                {
                    try
                    {
                        result = m_Funcs[index](stream);
                        if (result != null)
                            break;
                    }
                    catch (NotSupportedException ex)
                    {
                        if (index < 0)
                            throw ex;
                    }
                    index--;
                }
            }
            finally
            {
                if (stream.CanRead) stream.Close();
            }
            return result;
        }
    }
}