using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
                //else if (priceWithVAT.CellType == CellType.String)
                //{
                //    if (!double.TryParse(priceWithVAT.StringCellValue, out priceFinal))
                //    {
                //        continue; // Если не удалось распарсить строку в число
                //    }
                //}
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

                // Если цена без НДС указана, умножаем на 1.2 для получения цены с НДС
                //if (priceWithoutVAT != null && priceWithoutVAT.CellType == CellType.Numeric)
                //{
                //    priceFinal = priceWithoutVAT.NumericCellValue * 1.2;
                //}

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

            double totalMargin = totalCost - totalExpend;

            AnalysisData.Add(
                new AnalysisData
                {
                    Analysis = analysisName,
                    Parameters = string.Join(Environment.NewLine, parametersNames),
                    Expend = totalExpend,
                    Cost = totalCost,
                    Margin = totalMargin
                }
            );
        }

        //MessageBox.Show("Стоимость рассчитана.");
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

    //static void ExportDataToExcel(string[] args)
    //{
    //    // Создаем новый рабочий файл Excel
    //    IWorkbook workbook = new XSSFWorkbook(); // Используйте HSSFWorkbook для .xls
    //    ISheet sheetParameters = workbook.CreateSheet("По параметрам");
    //    ISheet sheetAnalysis = workbook.CreateSheet("По исследованиям");
    //    ISheet sheetTotal = workbook.CreateSheet("Итого");

    //    // Пример данных для записи
    //    var data = new List<string[]>
    //    {
    //        new string[] { "Имя", "Возраст", "Город" },
    //        new string[] { "Алексей", "30", "Москва" },
    //        new string[] { "Мария", "25", "Санкт-Петербург" },
    //        new string[] { "Иван", "35", "Екатеринбург" }
    //    };

    //    // Записываем данные в ячейки
    //    for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
    //    {
    //        IRow row = sheetParameters.CreateRow(rowIndex);
    //        for (int colIndex = 0; colIndex < data[rowIndex].Length; colIndex++)
    //        {
    //            row.CreateCell(colIndex).SetCellValue(data[rowIndex][colIndex]);
    //        }
    //    }

    //    // Сохраняем файл на диск
    //    using (var fileData = new FileStream("ExportedData.xlsx", FileMode.Create))
    //    {
    //        workbook.Write(fileData);
    //    }

    //    Console.WriteLine("Данные успешно экспортированы в ExportedData.xlsx");
    //}

    protected void NotifyPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}