using System.ComponentModel;

namespace FinancialAssistant
{
    public class ParameterData : INotifyPropertyChanged
    {
        private double _coefficient;

        public ParameterData(string name, double eachExpend, double count)
        {
            Name = name;
            EachExpend = eachExpend; // Получаем цену за единицу
            Count = count; // Количество показателя
             
            Coefficient = 1; // коэффициент по умолчанию
            TotalExpend = eachExpend * count; // Расход за показатель
            EachCost = eachExpend * Coefficient;
            TotalCost = count * Coefficient * eachExpend; // Расчет стоимости
            TotalMargin = TotalCost - TotalExpend; // Маржинальность
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

                EachCost = EachExpend * Coefficient;
                TotalCost = Count * Coefficient * EachExpend;
                TotalMargin = TotalCost - TotalExpend;

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

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}