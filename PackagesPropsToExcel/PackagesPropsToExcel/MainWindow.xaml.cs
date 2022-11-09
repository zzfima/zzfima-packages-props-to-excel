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
            waitClock.Spin = false;
            waitClock.Visibility = Visibility.Hidden;
            _convertPropsFile = new ConvertPropsFile();
            UpdateStatusLabel();
        }

        private void OnGenerateExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                waitClock.Spin = true;
                waitClock.Visibility = Visibility.Visible;
                _convertPropsFile.GenerateExcel(_packagesPropsPath, _destinationExcelPath);
            }
            finally
            {
                waitClock.Spin = false;
                waitClock.Visibility = Visibility.Hidden;
            }
        }

        private void OnSelectExcel(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                _destinationExcelPath = openFileDialog.FileName;

            UpdateStatusLabel();
        }

        private void OnSelectPackagesProps(object sender, RoutedEventArgs e)
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
