using System;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public partial class NewPeriodViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private int month = DateTime.Now.Month;

        [ObservableProperty]
        private int year = DateTime.Now.Year;

        [ObservableProperty]
        private bool copyCustomers = true;

        [ObservableProperty]
        private bool moveCurrentToPrevious = true;

        [ObservableProperty]
        private bool resetCurrentToZero = true;

        [RelayCommand]
        private void Ok()
        {
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }

        public string PeriodLabel => $"Tháng {Month:00}/{Year}";
    }
}
