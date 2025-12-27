namespace ElectricCalculation.ViewModels
{
    public sealed record ColumnOption(string? ColumnLetter, string HeaderText, string SamplePreview)
    {
        public string DisplayName
        {
            get
            {
                if (string.IsNullOrWhiteSpace(ColumnLetter))
                {
                    return "(Không dùng)";
                }

                var header = (HeaderText ?? string.Empty).Trim();
                return string.IsNullOrWhiteSpace(header) ? ColumnLetter! : $"{ColumnLetter} - {header}";
            }
        }
    }
}
