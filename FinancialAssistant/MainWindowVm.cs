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
    private double _totalMarginAll;
    private string _statusBarText;
    private double _vat;

    public event PropertyChangedEventHandler? PropertyChanged;

    public MainWindowVm()
    {
        AnalysisData = new ObservableCollection<AnalysisData>();        
        UniqueParameters = new ObservableCollection<ParameterData>();        
    }

    public string StatusBarText
    {
        get => _statusBarText;
        set
        {
            _statusBarText = value;
            NotifyPropertyChanged();
        }
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
            _totalExpend = Math.Round(value, 2);
            NotifyPropertyChanged();
        }
    }

    public double TotalCostWV
    {
        get => _totalCostWV;
        set
        {
            _totalCostWV = Math.Round(value, 2);
            NotifyPropertyChanged();
        }
    }

    public string VAT
    {
        get => _vat.ToString();
        set
        {
            double converted;
            if (Double.TryParse(value, out converted))
            {
                _vat = Math.Round(converted, 0);
            }
            else 
            {
                _vat = 0;
            }

            NotifyPropertyChanged();
            CalculateCost();
            TotalCalculate();
        }
    }

    public double TotalCostVAT
    {
        get => _totalCostVAT;
        set
        {
            _totalCostVAT = Math.Round(value, 2);
            NotifyPropertyChanged();
        }
    }
        public double TotalMarginAll
    {
        get => _totalMarginAll;
        set
        {
            _totalMarginAll = Math.Round(value, 2);
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
        if (File.Exists(filePath) == false)
        {
            StatusBarText = "Файл с исследованиями не прикреплён";
            return;
        }
        else
        {
            StatusBarText = "Файл с исследованиями загружен";
        }


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
        if (File.Exists(filePath) == false)
        {
            StatusBarText = "Файл с ценами не прикреплён";
            return;
        }
        else
        {
            StatusBarText = "Файл с ценами загружен";
        }

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
                if (nameCell is not null & priceWithVAT is null)
                {
                    priceFinal = Math.Round((priceWithoutVAT.NumericCellValue * 1.2), 2);
                }        
                else if (priceWithVAT.CellType == CellType.Numeric)
                {
                    priceFinal = Math.Round(priceWithVAT.NumericCellValue, 2); // Получаем числовое значение
                }
                else if (priceWithVAT.CellType == CellType.Formula)
                {
                    // Обработка формулы
                    var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
                    var evalResult = evaluator.Evaluate(priceWithVAT);

                    if (evalResult != null && evalResult.CellType == CellType.Numeric)
                    {
                        priceFinal = Math.Round(evalResult.NumberValue, 2);
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
                StatusBarText = $"Цена для показателя '{pair.Key}' не найдена.";
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
                    totalExpend += Math.Round(priceFinal, 2);
                }
                else
                {
                    StatusBarText = $"Цена для показателя '{parameterName}' не найдена.";
                }
            }

            foreach (var parameterName in parametersNames) // расчет доходов (стоимости исследования)
            {
                ParameterData? parameter = UniqueParameters.FirstOrDefault(p => p.Name == parameterName);

                if (parameter is null)
                {
                    StatusBarText = $"Цена для показателя '{parameterName}' не найдена.";
                    continue;
                }

                totalCost += Math.Round(parameter.EachCost, 2);
            }

            double totalMargin = Math.Round(((totalCost - totalExpend)/totalCost * 100), 2);
            double totalCostWithVAT;

            if (_vat == 0)
            {
                totalCostWithVAT = Math.Round((totalCost * 1), 2);
            }
            else
            {
                totalCostWithVAT = Math.Round((totalCost + (totalCost / 100 * _vat)), 2);
            }

            AnalysisData.Add(
                new AnalysisData
                {
                    Analysis = analysisName,
                    Parameters = string.Join(Environment.NewLine, parametersNames),
                    Expend = Math.Round(totalExpend, 2),
                    Cost = Math.Round(totalCost, 2),
                    CostWithVAT = Math.Round(totalCostWithVAT, 2),
                    Margin = Math.Round(totalMargin, 2)
                }
            );
        }

        StatusBarText = "Стоимость рассчитана.";
    }

    public void TotalCalculate()
    {
        TotalExpend = 0;
        TotalCostWV = 0;
        TotalMarginAll = 0;
        foreach (var analysis in AnalysisData)
        {
            TotalExpend += Math.Round(analysis.Expend, 2);
            TotalCostWV += Math.Round(analysis.Cost, 2);
        }

        if (_vat == 0)
        {
            TotalCostVAT = Math.Round((TotalCostWV * 1), 2);
        }
        else
        {
            TotalCostVAT = Math.Round((TotalCostWV + (TotalCostWV / 100 * _vat)), 2);
        }
            TotalMarginAll = Math.Round(((TotalCostWV - TotalExpend) / TotalCostWV * 100), 2);
    }

    public void ExportDataToExcel()
    {
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
                row.CreateCell(4).SetCellValue(parameter.EachCost);
                row.CreateCell(5).SetCellValue(parameter.TotalExpend);
                row.CreateCell(6).SetCellValue(parameter.TotalCost);
                row.CreateCell(7).SetCellValue(parameter.TotalMargin);
            }

            // Создаем второй лист и заполняем его данными из AnalysisData
            ISheet sheet2 = workbook.CreateSheet("По исследованиям");
            CreateHeaderRow(sheet2, new[] { "Исследование", "Параметры", "Расходы", "Стоимость исследования для клиента", "Стоимость с НДС", "Маржинальность" });

            rowIndex = 1;
            foreach (var analysis in AnalysisData)
            {
                IRow row = sheet2.CreateRow(rowIndex++);
                row.CreateCell(0).SetCellValue(analysis.Analysis);
                row.CreateCell(1).SetCellValue(analysis.Parameters);
                row.CreateCell(2).SetCellValue(analysis.Expend);
                row.CreateCell(3).SetCellValue(analysis.Cost);
                row.CreateCell(4).SetCellValue(analysis.CostWithVAT);
                row.CreateCell(5).SetCellValue(analysis.Margin);
            }

            // Сохраняем файл
            using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream);
            }
            StatusBarText = "Данные успешно экспортированы";
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