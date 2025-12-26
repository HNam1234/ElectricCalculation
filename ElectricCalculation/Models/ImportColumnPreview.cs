using System.Collections.Generic;
using ElectricCalculation.Services;

namespace ElectricCalculation.Models
{
    public sealed record ImportColumnPreview(
        string ColumnLetter,
        string HeaderText,
        IReadOnlyList<string> SampleValues,
        ExcelImportService.ImportField? SuggestedField,
        double SuggestedScore);
}

