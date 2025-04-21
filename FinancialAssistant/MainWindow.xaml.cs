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

namespace FinancialAssistant
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Dictionary<string, List<string>> _analysisIndicatorsMap = new Dictionary<string, List<string>>();

        private Dictionary<string, double> _indicatorsCount = new Dictionary<string, double>();
        private Dictionary<string, double> _indicatorsPrices = new Dictionary<string, double>();

        private ObservableCollection<AnalysisData> _analysisData;
        private ObservableCollection<ParameterData> _uniqueParameters;

        public MainWindow()
        {
            InitializeComponent();
            InitializeUniqueParametersSheet();
            InitializeAnalysisDataSheet();

            // Подписываемся на событие CellValueChanged
            //dataGridView1.CurrentCellChanged += OnCurrentCellChanged;
            //dataGridView1.SelectedCellsChanged += OnSelectedCellsChanged;

            // Также необходимо подписаться на событие CurrentCellDirtyStateChanged,
            // чтобы обновить значение сразу после изменения ячейки.
            //dataGridView1.CurrentCellDirtyStateChanged += (s, e) =>
            // {
            //     if (dataGridView1.IsCurrentCellDirty)
            //     {
            //         dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //     }
            // };
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Обработка изменения выбранной ячейки
            // Например, вы можете получить текущую ячейку:
            //if (dataGridView1.SelectedCells.Count > 0)
            //{
            //    var cellInfo = dataGridView1.SelectedCells[0];
            //    // Ваш код для обработки выбранной ячейки
            //}
            MessageBox.Show("jyt5dercjyt");
        }

        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // Обработка завершения редактирования ячейки
            //if (e.EditAction == DataGridEditAction.Commit)
            //{
            //    // Здесь вы можете выполнить действия после редактирования ячейки
            //    var editedCell = e.Column.GetCellContent(e.Row);
            //    // Ваш код для обработки измененной ячейки
            //}
            MessageBox.Show("!!!!!!!!!!!kyfkyft");
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
#if DEBUG
            ReserchExcells();
#endif

            InitializeUniqueParametersSheet();
            InitializeAnalysisDataSheet();
        }

        private void InitializeUniqueParametersSheet()
        {
            _uniqueParameters = new ObservableCollection<ParameterData>()            
            {
                new ParameterData
                {
                    Name = "Параметр 1",
                    EachExpend = 100.50m,
                    Count = 5,
                    Coefficient = 1.2m,
                    EachCost = 20.00m,
                    TotalExpend = 502.50m, // EachExpend * Count
                    TotalCost = 100.00m,   // EachCost * Count
                    TotalMargin = 50.00m    // TotalCost - TotalExpend (пример)
                },
                new ParameterData
                {
                    Name = "Параметр 2",
                    EachExpend = 200.75m,
                    Count = 3,
                    Coefficient = 0.8m,
                    EachCost = 30.00m,
                    TotalExpend = 602.25m, // EachExpend * Count
                    TotalCost = 90.00m,     // EachCost * Count
                    TotalMargin = -512.25m   // TotalCost - TotalExpend (пример)
                }
            };

            dataGridView1.ItemsSource = _uniqueParameters;
            //_uniqueParameters.Rows.Clear();

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Показатель",
            //        ReadOnly = true,
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Цена за шт",
            //        ReadOnly = true
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Кол-во",
            //        ReadOnly = true
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Коэффициент",
            //        ReadOnly = false
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Цена за шт. для клиента",
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Расход за показатель",
            //        ReadOnly = true
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Цена для клиента за показатель всего"
            //    }
            // );

            //_uniqueParameters.Columns.Add(
            //    new DataColumn()
            //    {
            //        ColumnName = "Маржинальность"
            //    }
            //);

            //dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //dataGridView1.Columns["Цена за шт. для клиента"].ReadOnly = true;
            //dataGridView1.Columns["Цена для клиента за показатель всего"].ReadOnly = true;
            //dataGridView1.Columns["Маржинальность"].ReadOnly = true;
        }

        private void InitializeAnalysisDataSheet()
        {
            _analysisData = new ObservableCollection<AnalysisData>() {
                new AnalysisData
                {
                    Analysis = "Исследование 1",
                    Parameters = "Показатель 1",
                    ExpendAnalysis = 100,
                    CostAnalysis = 150,
                    MarginAnalysis = 50
                },
                new AnalysisData
                {
                    Analysis = "Исследование 2",
                    Parameters = "Показатель 2",
                    ExpendAnalysis = 200,
                    CostAnalysis = 300,
                    MarginAnalysis = 100
                },
            };

            dataGridView2.ItemsSource = _analysisData;
            //_analysisData = new DataTable();
            //_analysisData.Columns.Add("Исследование", typeof(string));
            //_analysisData.Columns.Add("Показатели", typeof(string));
            //_analysisData.Columns.Add("Расходы на исследование", typeof(float));
            //_analysisData.Columns.Add("Стоимость исследования", typeof(float));
            //_analysisData.Columns.Add("Маржинальность исследования", typeof(float));

            //dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            //dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            //dataGridView2.ItemsSource = _analysisData;
            //dataGridView2.Columns["Показатели"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //dataGridView2.Columns["Исследование"].ReadOnly = true;
            //dataGridView2.Columns["Показатели"].ReadOnly = true;
            //dataGridView2.Columns["Расходы на исследование"].ReadOnly = true;
            //dataGridView2.Columns["Стоимость исследования"].ReadOnly = true;
            //dataGridView2.Columns["Маржинальность исследования"].ReadOnly = true;
        }

        private void ReserchExcells()
        {
            var researchPath = @"C:\Users\svetl\Desktop\Маша\sheet\analysis-parameter.xlsx";
            var analisysPath = @"C:\Users\svetl\Desktop\Маша\sheet\parameters_lab.cost.xlsx";

            label1.Content = researchPath;
            label2.Content = analisysPath;

            LoadResearch(researchPath);
            LoadPrices(analisysPath);
        }

        private void btnLoadResearch_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog().HasValue)
            {
                label1.Content = openFileDialog.FileName;
                LoadResearch(openFileDialog.FileName);
            }
        }

        private void btnLoadPrices_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog().HasValue)
            {
                label2.Content = openFileDialog.FileName;
                LoadPrices(openFileDialog.FileName);
            }
        }

        private void LoadResearch(string filePath)
        {
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = System.IO.Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    string analysisName = sheet.SheetName;
                    var indicators = new List<string>();

                    if (sheet.SheetName == "Все показатели")
                    {
                        continue;
                    }

                    for (int row = 0; row <= sheet.LastRowNum; row++)
                    {
                        var cell = sheet.GetRow(row)?.GetCell(0); // Столбец A
                        if (cell == null)
                        {
                            continue;
                        }

                        string indicatorName = cell.ToString();

                        if (string.IsNullOrWhiteSpace(indicatorName))
                        {
                            continue;
                        }

                        indicators.Add(indicatorName);

                        if (_indicatorsCount.ContainsKey(indicatorName))
                        {
                            _indicatorsCount[indicatorName] += 1;
                        }
                        else
                        {
                            _indicatorsCount[indicatorName] = 1;
                        }
                    }

                    _analysisIndicatorsMap.Add(analysisName, indicators);
                }
            }
        }


        private void LoadPrices(string filePath)
        {
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = System.IO.Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);

                var sheet = workbook.GetSheetAt(0); // Предполагаем, что данные на первом листе

                for (int row = 1; row <= sheet.LastRowNum; row++) // Пропускаем заголовок
                {
                    var nameCell = sheet.GetRow(row)?.GetCell(0); // Столбец A
                    var priceCellFinal = sheet.GetRow(row)?.GetCell(1); // Столбец D !!!!!!!!!!


                    if (nameCell != null && priceCellFinal != null)
                    {
                        string indicatorName = nameCell.ToString();
                        double priceFinal;

                        if (double.TryParse(priceCellFinal.ToString(), out priceFinal))
                        {
                            _indicatorsPrices[indicatorName] = priceFinal;
                        }
                    }
                }
            }
        }

        private void btnShowUniqueIndicators_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pair in _indicatorsCount)
            {
                // Проверяем, существует ли цена для данного показателя
                if (_indicatorsPrices.TryGetValue(pair.Key, out var priceFinal))
                {
                    double eachCost = priceFinal; // Получаем цену за единицу
                    double count = pair.Value; // Количество показателя
                    double coefficient = 1; // коэффициент по умолчанию
                    double expend = priceFinal * count; // Расход за показатель
                    double eachCostClient = eachCost * coefficient;
                    double cost = count * coefficient * eachCost; // Расчет стоимости
                    double margin = cost - expend; // Маржинальность

                    // Добавляем строку с данными
                    //_uniqueParameters.Rows.Add(pair.Key, eachCost, count, coefficient, eachCostClient, expend, cost, margin);
                }
                else
                {
                    MessageBox.Show($"Цена для показателя '{pair.Key}' не найдена.");
                }
            }
        }

        private void OnCurrentCellChanged(object sender, EventArgs e)
        {
            //DataRowView dataRow = (DataRowView)dataGridView1.SelectedItem;
            //var cell = dataGridView1.CurrentCell;
            //if (e.ColumnIndex != dataGridView1.Columns["Коэффициент"].Index)
            //{
            //    return;
            //}

            //string cellValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            //float coeff = 0;

            //if (float.TryParse(cellValue, out coeff) == false)
            //{
            //    MessageBox.Show("Коэффициент должен быть числом.\nУстанавливается значение по умолчанию: 1.");
            //    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 1;
            //    return;
            //}

            //if (coeff < 0)
            //{
            //    MessageBox.Show("Коэффициент не может быть меньше нуля.\nУстанавливается значение по умолчанию: 1.");
            //    dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 1;
            //    return;
            //}
            //// Получаем количество и цену за единицу из текущей строки
            //double count = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["Кол-во"].Value);
            //double eachCost = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["Цена за шт"].Value);
            //double expend = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["Расход за показатель"].Value);

            //// Пересчитываем стоимость
            //double newCost = count * coeff * eachCost;
            //double newEachCostClient = coeff * eachCost;
            //double newMargin = newCost - expend;

            //dataGridView1.Rows[e.RowIndex].Cells["Цена за шт. для клиента"].Value = newEachCostClient;
            //dataGridView1.Rows[e.RowIndex].Cells["Цена для клиента за показатель всего"].Value = newCost;
            //dataGridView1.Rows[e.RowIndex].Cells["Маржинальность"].Value = newMargin;
        }

        private void btnCalculateCost_Click(object sender, RoutedEventArgs e)
        {
            foreach (var pair in _analysisIndicatorsMap)
            {
                string analysisName = pair.Key; // Название исследования
                var indicators = pair.Value; // Список показателей для данного исследования

                double totalExpend = 0; // общие расходы (цена лаборатории)
                double totalCost = 0; // общая стоимость

                foreach (var indicator in indicators) // расчет расходов
                {
                    if (_indicatorsPrices.TryGetValue(indicator, out var priceFinal))
                    {
                        totalExpend += priceFinal;
                    }
                    else
                    {
                        MessageBox.Show($"Цена для показателя '{indicator}' не найдена.");
                    }
                }

                //foreach (var indicator in indicators) // расчет доходов (стоимости исследования)
                //{
                //    if (_uniqueParameters.Rows[].Cells["Показатель"].Value == indicator)
                //    {
                //        double eachCost = Convert.ToDouble(_uniqueParameters.Rows[].Cells["Цена за шт. для клиента"].Value);
                //        totalCost += eachCost;
                //    }
                //    else
                //    {
                //        MessageBox.Show($"Цена для показателя '{indicator}' не найдена.");
                //    }
                //}

                //foreach (var indicator in indicators) // расчет доходов (стоимости исследования)
                //{
                //    foreach (DataGridViewRow row in dataGridView1.Rows)
                //    {
                //        if (row.Cells["Показатель"].Value != null && row.Cells["Показатель"].Value.ToString() == indicator)
                //        {
                //            double eachCost = Convert.ToDouble(row.Cells["Цена за шт. для клиента"].Value);
                //            totalCost += eachCost;

                //        }
                //    }
                //}

                double totalMarge = totalCost - totalExpend;

                //_analysisData.Rows.Add(analysisName, string.Join(Environment.NewLine, indicators), totalExpend, totalCost, totalMarge);
            }

            MessageBox.Show("Стоимость рассчитана.");
        }

        private void btnTotalSumForOrder_Click(object sender, RoutedEventArgs e)
        {
            double totalSum = 0;

            //foreach (DataGridViewRow row in dataGridView2.Rows)
            //{
            //    double eachCost = Convert.ToDouble(row.Cells["Стоимость исследования"].Value);
            //    totalSum += eachCost;
            //}

            MessageBox.Show($"Итоговая стоимость заказа - '{totalSum} рублей без НДС'");
        }
    }

    public class AnalysisData
    {
        public string Analysis { get; set; }
        public string Parameters { get; set; }
        public float ExpendAnalysis { get; set; }
        public float CostAnalysis { get; set; }
        public float MarginAnalysis { get; set; }
    }

    public class ParameterData
    {
        public string Name { get; set; }
        public decimal EachExpend { get; set; }
        public decimal Count {  get; set; }
        public decimal Coefficient { get; set; }
        public decimal EachCost { get; set; }
        public decimal TotalExpend { get; set; }
        public decimal TotalCost { get; set; }
        public decimal TotalMargin { get; set; }
    }
}