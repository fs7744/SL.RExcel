using System;
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
    public class DirectoryRoot
    {
        private System.Collections.Generic.List<XLSDirectory> dirs;
        private System.Collections.Generic.List<Sector> sectors;
        private System.Collections.Generic.List<SectorIndex> index;

        public DirectoryRoot(System.Collections.Generic.List<XLSDirectory> dirs, System.Collections.Generic.List<Sector> sectors, System.Collections.Generic.List<SectorIndex> index)
        {
            // TODO: Complete member initialization
            this.dirs = dirs;
            this.sectors = sectors;
            this.index = index;
        }

    }
}
