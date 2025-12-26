using System;
using System.Collections.Generic;
using ElectricCalculation.Services;

namespace ElectricCalculation.Models
{
    public sealed record ImportMappingProfile(
        string ProfileName,
        string SheetName,
        string HeaderSignature,
        Dictionary<ExcelImportService.ImportField, string> ConfirmedMap,
        DateTime LastUsedAt);
}

