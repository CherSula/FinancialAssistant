using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Media.Media3D;
using System.Windows;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Win32; // Не забудьте добавить этот using для работы с диалогом выбора папки


namespace FinancialAssistant;

public class MainWindowVm : INotifyPropertyChanged
{
    private Dictionary<string, List<string>> _analysisIndicatorsMap = new Dictionary<string, List<string>>();

    private Dictionary<string, double> _indicatorsCount = new Dictionary<string, double>();
    private Dictionary<string, double> _indicatorsPrices = new Dictionary<string, double>();

    private string _researchPath;
    private string _analysisPath;
    private double _totalExpend;
    private double _totalCostWV;
    private double _totalCostVAT;

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainWindowVm()
    {
        AnalysisData = new ObservableCollection<AnalysisData>();        
        UniqueParameters = new ObservableCollection<ParameterData>();        
    }

    public string ResearchPath
    {
        get => _researchPath;
        set
        {
            _researchPath = value;
            NotifyPropertyChanged();
        }
    }

    public string PricesPath
    {
        get => _analysisPath;
        set
        {
            _analysisPath = value;
            NotifyPropertyChanged();
        }
    }

    public double TotalExpend
    {
        get => _totalExpend;
        set
        {
            _totalExpend = value;
            NotifyPropertyChanged();
        }
    }

    public double TotalCostWV
    {
        get => _totalCostWV;
        set
        {
            _totalCostWV = value;
            NotifyPropertyChanged();
        }
    }

    public double TotalCostVAT
    {
        get => _totalCostVAT;
        set
        {
            _totalCostVAT = value;
            NotifyPropertyChanged();
        }
    }

    public ObservableCollection<ParameterData> UniqueParameters { get; set; }

    public ObservableCollection<AnalysisData> AnalysisData { get; set; }

    public void DebugLoad()
    {
#if DEBUG
        ReserchExcells();
        FillUniqueParameters();

        void ReserchExcells()
        {
            ResearchPath = @"C:\Users\svetl\Desktop\Маша\sheet\analysis-parameter.xlsx";
            PricesPath = @"C:\Users\svetl\Desktop\Маша\sheet\parameters_lab.cost.xlsx";

            LoadResearch(ResearchPath);
            LoadPrices(PricesPath);
        }
#endif
    }

    public void LoadResearch(string filePath)
    {
        IWorkbook workbook;

        using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                string analysisName = sheet.SheetName;
                var indicators = new List<string>();

                if (sheet.SheetName == "Все показатели" || _analysisIndicatorsMap.ContainsKey(analysisName))
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

    public void LoadPrices(string filePath)
    {
        IWorkbook workbook;
        using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            workbook = Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);
            var sheet = workbook.GetSheetAt(0); // Предполагаем, что данные на первом листе

            for (int row = 1; row <= sheet.LastRowNum; row++) // Пропускаем заголовок
            {
                var nameCell = sheet.GetRow(row)?.GetCell(0); // Столбец A
                var priceWithoutVAT = sheet.GetRow(row)?.GetCell(1); // Столбец B
                var priceWithVAT = sheet.GetRow(row)?.GetCell(2); // Столбец C

                if (nameCell is null || priceWithVAT is null)
                {
                    continue;
                }

                string? parameterName = nameCell.CellType == CellType.String ? nameCell.StringCellValue : nameCell.ToString();

                if (string.IsNullOrWhiteSpace(parameterName))
                {
                    continue;
                }

                double priceFinal = 0;

                // Обработка ячейки с ценой
                if (priceWithVAT.CellType == CellType.Numeric)
                {
                    priceFinal = priceWithVAT.NumericCellValue; // Получаем числовое значение
                }
                else if (priceWithVAT.CellType == CellType.Formula)
                {
                    // Обработка формулы
                    var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                    var evalResult = evaluator.Evaluate(priceWithVAT);

                    if (evalResult != null && evalResult.CellType == CellType.Numeric)
                    {
                        priceFinal = evalResult.NumberValue;
                    }
                    else
                    {
                        continue; // Если не удалось получить числовое значение из формулы
                    }
                }
                _indicatorsPrices[parameterName] = priceFinal;
            }
        }
    }

    public void FillUniqueParameters()
    {
        UniqueParameters.Clear();

        foreach (var pair in _indicatorsCount)
        {
            // Проверяем, существует ли цена для данного показателя
            if (_indicatorsPrices.TryGetValue(pair.Key, out var priceFinal))
            {
                UniqueParameters.Add(
                    new ParameterData(pair.Key, priceFinal, pair.Value)
                );
            }
            else
            {
                // TODO заменить на Status = $"Цена для показателя '{pair.Key}' не найдена."
                Console.WriteLine($"Цена для показателя '{pair.Key}' не найдена.");
            }
        }
    }

    public void CalculateCost()
    {
        AnalysisData.Clear();

        foreach (var pair in _analysisIndicatorsMap)
        {
            string analysisName = pair.Key; // Название исследования
            var parametersNames = pair.Value; // Список показателей для данного исследования

            double totalExpend = 0; // общие расходы (цена лаборатории)
            double totalCost = 0; // общая стоимость

            foreach (var parameterName in parametersNames) // расчет расходов
            {
                if (_indicatorsPrices.TryGetValue(parameterName, out var priceFinal))
                {
                    totalExpend += priceFinal;
                }
                else
                {
                    // TODO заменить на Status = $"Цена для показателя '{parameterName}' не найдена.";
                    //MessageBox.Show($"Цена для показателя '{parameterName}' не найдена.");
                }
            }

            foreach (var parameterName in parametersNames) // расчет доходов (стоимости исследования)
            {
                ParameterData? parameter = UniqueParameters.FirstOrDefault(p => p.Name == parameterName);

                if (parameter is null)
                {
                    //MessageBox.Show($"Цена для показателя '{parameterName}' не найдена.");
                    continue;
                }

                totalCost += parameter.EachCost;
            }

            double totalCostWithVAT = totalCost * 1.2;
            double totalMargin = totalCost - totalExpend;

            AnalysisData.Add(
                new AnalysisData
                {
                    Analysis = analysisName,
                    Parameters = string.Join(Environment.NewLine, parametersNames),
                    Expend = totalExpend,
                    Cost = totalCost,
                    //CostWithVAT = totalCostWithVAT,
                    Margin = totalMargin
                }
            );
        }

        MessageBox.Show("Стоимость рассчитана.");
    }

    public void TotalCalculate()
    {
        //var analysisCost = 
        //double totalSum = 0;

        //foreach (var analysisCost in )
        //{
        //double eachCost = ;
        //totalSum += eachCost;
        //}
        TotalExpend = 0;
        TotalCostWV = 0;
        foreach(var analysis in AnalysisData)
        {
            TotalExpend += analysis.Expend;
            TotalCostWV += analysis.Cost;
        }
        TotalCostVAT = TotalCostWV * 1.2;
    }

    public void ExportDataToExcel()
    {
        //// Открываем диалог выбора папки
        //OpenFileDialog fileDialog = new OpenFileDialog();

        //fileDialog.ShowDialog();
            // Открываем диалог выбора файла (можно использовать для выбора места сохранения)
        SaveFileDialog saveFileDialog = new SaveFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx",
            Title = "Сохранить файл как"
        };

        if (saveFileDialog.ShowDialog() == true)
        {
            string filePath = saveFileDialog.FileName;

            // Создаем новый Excel файл
            IWorkbook workbook = new XSSFWorkbook();

            // Создаем первый лист и заполняем его данными из UniqueParameters
            ISheet sheet1 = workbook.CreateSheet("По параметрам");
            CreateHeaderRow(sheet1, new[] { "Показатель", "Цена лаб/шт", "Количество", "Коэффициент", "Цена клиента/шт", "Расходы за показатель всего", "Цена для клиента всего", "Маржинальность" });

            int rowIndex = 1;
            foreach (var parameter in UniqueParameters)
            {
                IRow row = sheet1.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(parameter.Name);
                row.CreateCell(1).SetCellValue(parameter.EachExpend);
                row.CreateCell(2).SetCellValue(parameter.Count);
                row.CreateCell(3).SetCellValue(parameter.Coefficient);
                row.CreateCell(3).SetCellValue(parameter.EachCost);
                row.CreateCell(3).SetCellValue(parameter.TotalExpend);
                row.CreateCell(3).SetCellValue(parameter.TotalCost);
                row.CreateCell(3).SetCellValue(parameter.TotalMargin);
            }

            // Создаем второй лист и заполняем его данными из AnalysisData
            ISheet sheet2 = workbook.CreateSheet("По исследованиям");
            CreateHeaderRow(sheet2, new[] { "Исследование", "Параметры", "Расходы", "Стоимость исследования для клиента", "Маржинальность" });

            rowIndex = 1;
            foreach (var analysis in AnalysisData)
            {
                IRow row = sheet2.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(analysis.Analysis);
                row.CreateCell(1).SetCellValue(analysis.Parameters);
                row.CreateCell(2).SetCellValue(analysis.Expend);
                row.CreateCell(3).SetCellValue(analysis.Cost);
                //row.CreateCell(3).SetCellValue(analysis.CostWithVAT);
                row.CreateCell(4).SetCellValue(analysis.Margin);
            }

            // Сохраняем файл
            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }

            MessageBox.Show("Данные успешно экспортированы в Excel!", "Экспорт завершен", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    private void CreateHeaderRow(ISheet sheet, string[] headers)
    {
        IRow headerRow = sheet.CreateRow(0);
        for (int i = 0; i < headers.Length; i++)
        {
            headerRow.CreateCell(i).SetCellValue(headers[i]);
        }
    }

    protected void NotifyPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}