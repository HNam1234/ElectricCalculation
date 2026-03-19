using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public partial class SetColumnValueViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string columnTitle = string.Empty;

        [ObservableProperty]
        private string valueText = string.Empty;

        public SetColumnValueViewModel(string? columnTitle, string? initialValue = null)
        {
            ColumnTitle = columnTitle ?? string.Empty;
            ValueText = initialValue ?? string.Empty;
        }

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
    }
}

