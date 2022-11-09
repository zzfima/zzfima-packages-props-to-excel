using Logic;
using Microsoft.Win32;
using System.Threading.Tasks;
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

        private async void OnGenerateExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                waitClock.Spin = true;
                waitClock.Visibility = Visibility.Visible;

                btnSelelectExcel.IsEnabled = false;
                btnSelelectNugets.IsEnabled = false;
                btnRun.IsEnabled = false;

                await _convertPropsFile.GenerateExcel(_packagesPropsPath, _destinationExcelPath)
                    .ContinueWith((t) =>
                    {
                        if (t.IsFaulted)
                            MessageBox.Show(t.Exception.InnerException.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    });
            }
            finally
            {
                btnSelelectExcel.IsEnabled = true;
                btnSelelectNugets.IsEnabled = true;
                btnRun.IsEnabled = true;

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
