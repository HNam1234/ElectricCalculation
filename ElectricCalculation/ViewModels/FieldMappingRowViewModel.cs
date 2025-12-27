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

        private readonly IReadOnlyDictionary<string, ColumnOption> columnOptionsByLetter;

        private string? selectedColumn;
        private bool isLocked;
        private bool isConflict;
        private bool isMissingRequired;
        private string selectedSamplePreview = string.Empty;

        public FieldMappingRowViewModel(
            ExcelImportService.ImportField field,
            string label,
            string hint,
            bool isRequired,
            bool isKeyField,
            bool isAdvanced,
            IReadOnlyDictionary<string, ColumnOption> columnOptionsByLetter,
            string? suggestedColumn,
            double suggestedScore)
        {
            Field = field;
            Label = label ?? field.ToString();
            Hint = hint ?? string.Empty;
            IsRequired = isRequired;
            IsKeyField = isKeyField;
            IsAdvanced = isAdvanced;
            this.columnOptionsByLetter = columnOptionsByLetter ??
                                         new Dictionary<string, ColumnOption>(StringComparer.OrdinalIgnoreCase);

            SuggestedColumn = NormalizeColumnLetter(suggestedColumn);
            SuggestedScore = suggestedScore;
        }

        public event EventHandler? SelectionChanged;

        public ExcelImportService.ImportField Field { get; }

        public string Label { get; }

        public string Hint { get; }

        public bool IsRequired { get; }

        public bool IsKeyField { get; }

        public bool IsAdvanced { get; }

        public string? SuggestedColumn { get; }

        public double SuggestedScore { get; }

        public string SuggestedColumnDisplay
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SuggestedColumn))
                {
                    return string.Empty;
                }

                return columnOptionsByLetter.TryGetValue(SuggestedColumn, out var option)
                    ? option.DisplayName
                    : SuggestedColumn;
            }
        }

        public string ConfidenceLevel =>
            string.IsNullOrWhiteSpace(SuggestedColumn) ? "None" :
            SuggestedScore >= LockThreshold ? "High" :
            SuggestedScore >= AutoSelectThreshold ? "Medium" :
            SuggestedScore > 0 ? "Low" :
            "None";

        public string ConfidenceText
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SuggestedColumn))
                {
                    return "Chua co goi y";
                }

                return $"Goi y: {SuggestedColumnDisplay} ({SuggestedScore:0.00})";
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

        public string? SelectedColumn
        {
            get => selectedColumn;
            set
            {
                if (IsLocked)
                {
                    return;
                }

                var normalized = NormalizeColumnLetter(value);
                if (SetProperty(ref selectedColumn, normalized))
                {
                    SelectedSamplePreview = GetSamplePreview(normalized);
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        public bool IsConflict
        {
            get => isConflict;
            set => SetProperty(ref isConflict, value);
        }

        public bool IsMissingRequired
        {
            get => isMissingRequired;
            set => SetProperty(ref isMissingRequired, value);
        }

        public string SelectedSamplePreview
        {
            get => selectedSamplePreview;
            private set => SetProperty(ref selectedSamplePreview, value);
        }

        public void ResetSelection()
        {
            IsLocked = false;
            SelectedColumn = null;
            IsConflict = false;
            IsMissingRequired = false;
        }

        private string GetSamplePreview(string? columnLetter)
        {
            if (string.IsNullOrWhiteSpace(columnLetter))
            {
                return string.Empty;
            }

            return columnOptionsByLetter.TryGetValue(columnLetter, out var option)
                ? option.SamplePreview
                : string.Empty;
        }

        private static string? NormalizeColumnLetter(string? input)
        {
            var value = (input ?? string.Empty).Trim().ToUpperInvariant();
            return value.Length == 0 ? null : value;
        }
    }
}

