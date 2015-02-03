using SL.RExcel.XLS;
using System.Windows;
using System.Windows.Controls;

namespace Test
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            var a = new OpenFileDialog();
            a.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx";
            if (a.ShowDialog() == true)
            {
                new XLSWorkBook(a.File.OpenRead());
            }
        }
    }
}