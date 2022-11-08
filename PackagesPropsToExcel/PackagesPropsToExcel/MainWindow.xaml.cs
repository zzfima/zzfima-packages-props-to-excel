using Logic;
using Microsoft.Win32;
using System.Windows;

namespace PackagesPropsToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ConvertPropsFile _convertPropsFile;
        private string _packagesPropsPath, _destinationExcelPath;

        public MainWindow()
        {
            InitializeComponent();

            _convertPropsFile = new ConvertPropsFile();
            UpdateStatusLabel();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            _convertPropsFile.GenerateExcel(_packagesPropsPath, _destinationExcelPath);
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                _destinationExcelPath = openFileDialog.FileName;

            UpdateStatusLabel();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                _packagesPropsPath = openFileDialog.FileName;

            UpdateStatusLabel();
        }

        private void UpdateStatusLabel()
        {
            statusLabel1.Content = "DestinationExcelPath: " + _destinationExcelPath;
            statusLabel2.Content = "Packages.Props Path: " + _packagesPropsPath;
        }
    }
}
