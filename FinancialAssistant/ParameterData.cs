using System.ComponentModel;

namespace FinancialAssistant
{
    public class ParameterData : INotifyPropertyChanged
    {
        private double _coefficient;

        public ParameterData(string name, double eachExpend, double count)
        {
            Name = name;
            EachExpend = Math.Round(eachExpend, 2); // Получаем цену за единицу
            Count = count; // Количество показателя
             
            Coefficient = 1; // коэффициент по умолчанию
            TotalExpend = Math.Round(eachExpend * count, 2); // Расход за показатель
            EachCost = Math.Round(eachExpend * Coefficient, 2);
            TotalCost = Math.Round(count * Coefficient * eachExpend, 2); // Расчет стоимости
            TotalMargin = Math.Round(((TotalCost - TotalExpend) / TotalCost * 100), 2); // Маржинальность
        }

        public string Name { get; set; }
        public double EachExpend { get; set; }
        public double Count {  get; set; }
        public double Coefficient
        {
            get => _coefficient;
            set
            {
                //var v = value as double;
                //if (!double.TryParse(v, out _coefficient))
                //{
                //    _coefficient = 1;
                //    MessageBox.Show("Коэффициент должен быть числом.\nУстанавливается значение по умолчанию: 1.");
                //}

                //if (_coefficient < 0)
                //{
                //    _coefficient = 1;
                //    MessageBox.Show("Коэффициент не может быть меньше нуля.\nУстанавливается значение по умолчанию: 1.");
                //}
                if (value < 0)
                {
                    _coefficient = 1;
                }

                _coefficient = value;

                EachCost = Math.Round(EachExpend * Coefficient, 2);
                TotalCost = Math.Round((Count * Coefficient * EachExpend), 2);
                TotalMargin = Math.Round(((TotalCost - TotalExpend)/TotalCost * 100), 2);

                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Coefficient)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(EachCost)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TotalCost)));
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(TotalMargin)));
            }
        }
        public double EachCost { get; set; }
        public double TotalExpend { get; set; }
        public double TotalCost { get; set; }
        public double TotalMargin { get; set; }

        public string VAT {  get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}