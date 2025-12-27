using System;
using CommunityToolkit.Mvvm.ComponentModel;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed class FieldExplanationRowViewModel : ObservableObject
    {
        private string? selectedColumnLetter;

        public FieldExplanationRowViewModel(ExcelImportService.ImportField field, string appColumn, string description)
        {
            Field = field;
            AppColumn = appColumn ?? field.ToString();
            Description = description ?? string.Empty;
        }

        public event EventHandler? SelectionChanged;

        public ExcelImportService.ImportField Field { get; }

        public string AppColumn { get; }

        public string Description { get; }

        public string? SelectedColumnLetter
        {
            get => selectedColumnLetter;
            set
            {
                var normalized = NormalizeColumnLetter(value);
                if (SetProperty(ref selectedColumnLetter, normalized))
                {
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        private static string? NormalizeColumnLetter(string? input)
        {
            var value = (input ?? string.Empty).Trim().ToUpperInvariant();
            return value.Length == 0 ? null : value;
        }
    }
}
