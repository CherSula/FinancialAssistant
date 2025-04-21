using System.Data;
using System.IO;
using System.Reflection.Emit;
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
using Microsoft.Win32;
using System.Collections.Generic;
using NPOI.HSSF.UserModel; // Для .xls
using NPOI.XSSF.UserModel; // Для .xlsx
using NPOI.SS.UserModel;
using System.Linq;
using System.Collections.ObjectModel;
using NPOI.SS.Formula.Functions;
using NPOI.OpenXmlFormats.Wordprocessing;

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
            _vm.Load();
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
            //var analysisCost = 
            //double totalSum = 0;

            //foreach (var analysisCost in )
            //{
            //double eachCost = ;
            //totalSum += eachCost;
            //}

            //MessageBox.Show($"Итоговая стоимость заказа - '{totalSum} рублей без НДС'");
        }
    }
}