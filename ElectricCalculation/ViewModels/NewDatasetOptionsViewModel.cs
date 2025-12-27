using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public enum NewDatasetCreationOption
    {
        ManualEntry,
        ImportFromExcel
    }

    public partial class NewDatasetOptionsViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        public NewDatasetCreationOption? SelectedOption { get; private set; }

        [RelayCommand]
        private void ChooseManualEntry()
        {
            SelectedOption = NewDatasetCreationOption.ManualEntry;
            DialogResult = true;
        }

        [RelayCommand]
        private void ChooseImportFromExcel()
        {
            SelectedOption = NewDatasetCreationOption.ImportFromExcel;
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }
    }
}

