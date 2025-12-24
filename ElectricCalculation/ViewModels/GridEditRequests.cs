using System.Collections.Generic;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public sealed record CellEditChange(Customer Customer, string PropertyName, object? OldValue, object? NewValue);

    public sealed record ClipboardPasteRequest(
        IReadOnlyList<Customer> TargetRows,
        IReadOnlyList<string> PropertyNames,
        string ClipboardText);

    public sealed record FillDownRequest(IReadOnlyList<Customer> TargetRows, string PropertyName);
}

