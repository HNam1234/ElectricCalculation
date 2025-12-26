using System.Collections.Generic;
using System.Data;

namespace ElectricCalculation.Models
{
    public sealed record ImportSampleRow(int RowIndex, IReadOnlyDictionary<string, string?> Cells);

    public sealed record ImportRunReport(
        int TotalRows,
        int ImportedRows,
        int SkippedRows,
        int WarningCount,
        int ErrorCount,
        IReadOnlyList<string> Warnings,
        IReadOnlyList<string> Errors);

    public sealed record ImportPreviewResult(
        string FilePath,
        IReadOnlyList<string> SheetNames,
        string SelectedSheetName,
        int? HeaderRowIndex,
        int? DataStartRowIndex,
        string? HeaderSignature,
        IReadOnlyList<ImportColumnPreview> Columns,
        DataTable PreviewTable,
        IReadOnlyList<ImportSampleRow> SampleRows,
        string? WarningMessage);
}

