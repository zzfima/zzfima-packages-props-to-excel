using Logic;
using System.Windows;

namespace PackagesPropsToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ConvertPropsFile _convertPropsFile;

        public MainWindow()
        {
            InitializeComponent();

            _convertPropsFile = new ConvertPropsFile();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _convertPropsFile.GenerateExcel(@"C:\Temp\pp.txt", @"C:\Temp\PackagesProps2.xls");
        }
    }
}
