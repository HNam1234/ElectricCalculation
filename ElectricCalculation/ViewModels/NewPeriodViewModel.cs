using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace ElectricCalculation.ViewModels
{
    public partial class NewPeriodViewModel : ObservableObject
    {
        public sealed record ReferenceDatasetOption(
            string PeriodLabel,
            string DisplayName,
            string? SnapshotPath,
            bool IsCurrentDataset);

        public ObservableCollection<ReferenceDatasetOption> ReferenceDatasets { get; } = new();

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string errorMessage = string.Empty;

        [ObservableProperty]
        private int month = DateTime.Now.Month;

        [ObservableProperty]
        private int year = DateTime.Now.Year;

        [ObservableProperty]
        private bool moveCurrentToPrevious = true;

        [ObservableProperty]
        private bool resetCurrentToZero = true;

        [ObservableProperty]
        private ReferenceDatasetOption? selectedReferenceDataset;

        public NewPeriodViewModel()
        {
        }

        public NewPeriodViewModel(IEnumerable<ReferenceDatasetOption> referenceDatasets)
        {
            if (referenceDatasets != null)
            {
                foreach (var item in referenceDatasets)
                {
                    ReferenceDatasets.Add(item);
                }
            }

            SelectedReferenceDataset = ReferenceDatasets.FirstOrDefault();
        }

        [RelayCommand]
        private void Ok()
        {
            ErrorMessage = string.Empty;

            if (Month is < 1 or > 12 || Year < 2000)
            {
                ErrorMessage = "Tháng/Năm không hợp lệ.";
                return;
            }

            if (SelectedReferenceDataset == null)
            {
                ErrorMessage = "Hãy chọn bộ dữ liệu làm tháng cũ.";
                return;
            }

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
