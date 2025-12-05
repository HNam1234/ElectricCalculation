using System;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public partial class MainWindowViewModel : ObservableObject
    {
        [ObservableProperty]
        private string periodLabel = string.Empty;

        public ObservableCollection<CustomerReadingViewModel> Readings { get; } = new();

        public MainWindowViewModel()
        {
            PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
        }

        [RelayCommand]
        private void AddRow()
        {
            var customer = new Customer();
            var reading = new MeterReading
            {
                Multiplier = 1
            };

            Readings.Add(new CustomerReadingViewModel(customer, reading));
        }
    }
}

