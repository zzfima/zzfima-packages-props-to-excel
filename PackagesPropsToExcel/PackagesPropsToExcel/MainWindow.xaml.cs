using Logic;
using System;
using System.Collections.Generic;
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
            _convertPropsFile.GenerateExcel(@"C:\Sandboxes\Phoenix\efzabar_12.10.1_dev_stream\ASD_Phoenix\Packages.props", @"C:\Temp\PackagesProps1.xls");
        }
    }
}
