using SL.RExcel;
using System;
using System.Text;
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

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            var a = new OpenFileDialog();
            //a.Filter = "(*.xls)|*.xls|(*.xlsx)|*.xlsx";
            if (a.ShowDialog() == true)
            {
                try
                {
                    var workBook = ExcelHelper.Open(a.File.OpenRead());
                    txtShow.Text = GetInfo(workBook);
                }
                catch (Exception ex)
                {
                    txtShow.Text = ex.Message;
                    throw ex;
                }
                
            }
        }

        private string GetInfo(IWorkBook workBook)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var sheet in workBook.Worksheets)
            {
                sb.AppendLine("sheet : " + sheet.Name);
                foreach (var rowMap in sheet.GetAllRows())
                {
                    sb.AppendLine("row " + rowMap.Key + ": " + GetRow(rowMap.Value));
                }
            }

            return sb.ToString();
        }

        private string GetRow(IRow row)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var cellMap in row.GetAllCells())
            {
                sb.Append(cellMap.Key + ": " + cellMap.Value.GetStringValue() + "\t");
            }
            return sb.ToString();
        }
    }
}