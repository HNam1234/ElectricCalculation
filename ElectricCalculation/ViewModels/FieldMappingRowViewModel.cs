using System;
using System.Collections.Generic;
using CommunityToolkit.Mvvm.ComponentModel;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed class FieldMappingRowViewModel : ObservableObject
    {
        public const double AutoSelectThreshold = 0.55;
        public const double LockThreshold = 0.80;

        private string? selectedColumn;
        private bool isLocked;
        private string samplePreview = string.Empty;
        private bool isConflict;
        private bool isMissing;

        public FieldMappingRowViewModel(
            ExcelImportService.ImportField field,
            string label,
            string hint,
            bool isRequired,
            IReadOnlyList<ColumnOption> options,
            string? suggestedColumn,
            double suggestedScore)
        {
            Field = field;
            Label = label ?? string.Empty;
            Hint = hint ?? string.Empty;
            IsRequired = isRequired;
            Options = options ?? Array.Empty<ColumnOption>();
            SuggestedColumn = suggestedColumn;
            SuggestedScore = suggestedScore;
        }

        public event EventHandler? SelectionChanged;

        public ExcelImportService.ImportField Field { get; }

        public string Label { get; }

        public string Hint { get; }

        public bool IsRequired { get; }

        public IReadOnlyList<ColumnOption> Options { get; }

        public string? SuggestedColumn { get; }

        public double SuggestedScore { get; }

        public string ConfidenceLevel =>
            SuggestedScore >= LockThreshold ? "High" :
            SuggestedScore >= AutoSelectThreshold ? "Medium" :
            SuggestedScore > 0 ? "Low" :
            "None";

        public string ConfidenceText =>
            SuggestedScore >= LockThreshold ? $"Rất chắc ({SuggestedScore:0.00})" :
            SuggestedScore >= AutoSelectThreshold ? $"Gợi ý ({SuggestedScore:0.00})" :
            SuggestedScore > 0 ? $"Thấp ({SuggestedScore:0.00})" :
            "Chưa rõ";

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

        public string? SelectedColumn
        {
            get => selectedColumn;
            set
            {
                if (IsLocked)
                {
                    return;
                }

                if (SetProperty(ref selectedColumn, value))
                {
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public string SamplePreview
        {
            get => samplePreview;
            private set => SetProperty(ref samplePreview, value);
        }

        public bool IsConflict
        {
            get => isConflict;
            set => SetProperty(ref isConflict, value);
        }

        public bool IsMissing
        {
            get => isMissing;
            set => SetProperty(ref isMissing, value);
        }

        public void ResetSelection()
        {
            IsLocked = false;
            SelectedColumn = null;
        }

        public void UpdateSamplePreview(IReadOnlyDictionary<string, ColumnOption> optionsByLetter)
        {
            if (string.IsNullOrWhiteSpace(SelectedColumn))
            {
                SamplePreview = string.Empty;
                return;
            }

            var key = SelectedColumn.Trim().ToUpperInvariant();
            if (optionsByLetter.TryGetValue(key, out var option))
            {
                SamplePreview = option.SamplePreview ?? string.Empty;
                return;
            }

            SamplePreview = string.Empty;
        }
    }
}

