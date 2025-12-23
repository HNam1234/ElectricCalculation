using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public enum SaveSnapshotPromptAction
    {
        SaveNew,
        Overwrite,
        DontSave
    }

    public partial class SaveSnapshotPromptViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string periodLabel = string.Empty;

        [ObservableProperty]
        private int customerCount;

        [ObservableProperty]
        private string snapshotName = string.Empty;

        [ObservableProperty]
        private bool canOverwrite;

        [ObservableProperty]
        private SaveSnapshotPromptAction action = SaveSnapshotPromptAction.SaveNew;

        public SaveSnapshotPromptViewModel(string periodLabel, int customerCount, string? defaultSnapshotName = null, bool canOverwrite = false)
        {
            PeriodLabel = periodLabel ?? string.Empty;
            CustomerCount = customerCount;
            SnapshotName = defaultSnapshotName ?? string.Empty;
            CanOverwrite = canOverwrite;
        }

        [RelayCommand]
        private void SaveNew()
        {
            Action = SaveSnapshotPromptAction.SaveNew;
            DialogResult = true;
        }

        [RelayCommand]
        private void Overwrite()
        {
            Action = SaveSnapshotPromptAction.Overwrite;
            DialogResult = true;
        }

        [RelayCommand]
        private void DontSave()
        {
            Action = SaveSnapshotPromptAction.DontSave;
            DialogResult = false;
        }
    }
}
