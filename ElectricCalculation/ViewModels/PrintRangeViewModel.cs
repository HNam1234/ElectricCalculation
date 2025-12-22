using CommunityToolkit.Mvvm.Input;
using CommunityToolkit.Mvvm.ComponentModel;

namespace ElectricCalculation.ViewModels
{
    public partial class PrintRangeViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private int fromNumber = 1;

        [ObservableProperty]
        private int toNumber = 1;

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
