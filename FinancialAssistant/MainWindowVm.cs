using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using NPOI.HSSF.UserModel;
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

    protected void NotifyPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}