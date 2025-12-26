using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed record FieldOption(ExcelImportService.ImportField? Field, string DisplayName);
}

