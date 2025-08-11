using System.Configuration;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EtmParser
{
    public partial class MainWindow : Window
    {
        private EtmParserViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = (EtmParserViewModel)DataContext;
        }

        private void SaveSettings(object sender, RoutedEventArgs e)
        {
            _viewModel.UpdateSettings();
        }

        private void StartParse(object sender, RoutedEventArgs e)
        {
            _viewModel.ParseGoods();
        }
    }
}