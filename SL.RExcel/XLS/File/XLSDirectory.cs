using System;
using System.IO;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace SL.RExcel.XLS.File
{
    public class XLSDirectory
    {
        public XLSDirectory(Stream stream)
        {
            BinaryReader reader = new BinaryReader(stream);
        }

    }
}
