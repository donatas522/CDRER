using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using FindAndReplace;
using Microsoft.Win32;

namespace CDRER
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static readonly DependencyProperty WordFilePathProperty = DependencyProperty.Register(
            "WordFilePath", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        public string WordFilePath
        {
            get { return (string) GetValue(WordFilePathProperty); }
            set { SetValue(WordFilePathProperty, value); }
        }

        public static readonly DependencyProperty ExcelFilePathProperty = DependencyProperty.Register(
            "ExcelFilePath", typeof(string), typeof(MainWindow), new PropertyMetadata(string.Empty));

        public string ExcelFilePath
        {
            get { return (string) GetValue(ExcelFilePathProperty); }
            set { SetValue(ExcelFilePathProperty, value); }
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseWordFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = @"Word File(.docx,.doc) | *.docx; *.doc";
            if (ofd.ShowDialog() == true)
            {
                WordFilePath = ofd.FileName;
            }
        }

        private void BrowseExcelFileButton_OnClick(object sender, RoutedEventArgs e)
        {
            var ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            if (ofd.ShowDialog() == true)
            {
                ExcelFilePath = ofd.FileName;
            }
        }

        private bool CheckIfWordDocContainsString(string text)
        {
            var flatDoc = new FlatDocument(WordFilePath);
            var hasString = flatDoc.FindText(text);
            
            flatDoc.Close();
            flatDoc = null;        
            return hasString;
        }

        private async void CheckWordText_OnClick(object sender, RoutedEventArgs e)
        {
            DoExistWordString.Text = CheckIfWordDocContainsString(TextToSearchBox.Text) ? "OK." : "Fail.";
        }
    }
}
