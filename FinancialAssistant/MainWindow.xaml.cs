using System.Windows;
using Microsoft.Win32;

namespace FinancialAssistant
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainWindowVm _vm;

        public MainWindow()
        {
            _vm = new MainWindowVm();
            DataContext = _vm;
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            _vm.DebugLoad();
        }

        private void btnLoadResearch_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog().HasValue)
            {
                _vm.ResearchPath = openFileDialog.FileName;
                _vm.LoadResearch(openFileDialog.FileName);
            }
        }

        private void btnLoadPrices_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog().HasValue)
            {
                _vm.PricesPath = openFileDialog.FileName;
                _vm.LoadPrices(openFileDialog.FileName);
            }
        }

        private void btnShowUniqueIndicators_Click(object sender, RoutedEventArgs e)
        {
            _vm.FillUniqueParameters();
        }

        private void btnCalculateCost_Click(object sender, RoutedEventArgs e)
        {
            _vm.CalculateCost();
        }

        private void btnTotalSumForOrder_Click(object sender, RoutedEventArgs e)
        {
            _vm.TotalCalculate();
        }
    }
}