using System;
using CommunityToolkit.Mvvm.ComponentModel;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed class ColumnMappingViewModel : ObservableObject
    {
        public const double AutoSelectThreshold = 0.55;
        public const double LockThreshold = 0.80;

        private ExcelImportService.ImportField? selectedField;
        private bool isLocked;
        private bool isConflict;
        private bool isRequiredMissing;

        public ColumnMappingViewModel(
            string columnLetter,
            string headerText,
            string samplePreview,
            ExcelImportService.ImportField? suggestedField,
            string? suggestedFieldLabel,
            double suggestedScore)
        {
            ColumnLetter = (columnLetter ?? string.Empty).Trim().ToUpperInvariant();
            HeaderText = headerText ?? string.Empty;
            SamplePreview = samplePreview ?? string.Empty;
            SuggestedField = suggestedField;
            SuggestedFieldLabel = suggestedFieldLabel ?? string.Empty;
            SuggestedScore = suggestedScore;
        }

        public event EventHandler? SelectionChanged;

        public string ColumnLetter { get; }

        public string HeaderText { get; }

        public string SamplePreview { get; }

        public ExcelImportService.ImportField? SuggestedField { get; }

        public string SuggestedFieldLabel { get; }

        public double SuggestedScore { get; }

        public string HeaderDisplay
        {
            get
            {
                var header = (HeaderText ?? string.Empty).Trim();
                return string.IsNullOrWhiteSpace(header) ? ColumnLetter : $"{ColumnLetter} - {header}";
            }
        }

        public string ConfidenceLevel =>
            SuggestedField == null ? "None" :
            SuggestedScore >= LockThreshold ? "High" :
            SuggestedScore >= AutoSelectThreshold ? "Medium" :
            SuggestedScore > 0 ? "Low" :
            "None";

        public string ConfidenceText
        {
            get
            {
                if (SuggestedField == null || string.IsNullOrWhiteSpace(SuggestedFieldLabel))
                {
                    return "Chưa có gợi ý";
                }

                return $"Gợi ý: {SuggestedFieldLabel} ({SuggestedScore:0.00})";
            }
        }

        public bool IsEditable => !IsLocked;

        public bool IsLocked
        {
            get => isLocked;
            set
            {
                if (SetProperty(ref isLocked, value))
                {
                    OnPropertyChanged(nameof(IsEditable));
                }
            }
        }

        public ExcelImportService.ImportField? SelectedField
        {
            get => selectedField;
            set
            {
                if (IsLocked)
                {
                    return;
                }

                if (SetProperty(ref selectedField, value))
                {
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public bool IsConflict
        {
            get => isConflict;
            set => SetProperty(ref isConflict, value);
        }

        public bool IsRequiredMissing
        {
            get => isRequiredMissing;
            set => SetProperty(ref isRequiredMissing, value);
        }

        public void ResetSelection()
        {
            IsLocked = false;
            SelectedField = null;
            IsConflict = false;
            IsRequiredMissing = false;
        }
    }
}
