using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public partial class ImportMappingViewModel : ObservableObject
    {
        private const double AutoMapThreshold = 0.55;

        private readonly string filePath;
        private ImportPreviewResult preview;
        private bool suppressSheetReload;

        private Dictionary<ExcelImportService.ImportField, string> confirmedMap = new();

        public IReadOnlyList<TargetFieldOption> TargetFieldOptions { get; } = BuildTargetFieldOptions();

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string previewWarningMessage = string.Empty;

        [ObservableProperty]
        private ObservableCollection<string> sheetNames = new();

        [ObservableProperty]
        private string selectedSheetName = string.Empty;

        [ObservableProperty]
        private string profileName = string.Empty;

        [ObservableProperty]
        private bool saveProfile = true;

        [ObservableProperty]
        private DataView? previewView;

        [ObservableProperty]
        private ObservableCollection<ImportColumnMappingRowViewModel> columnMappings = new();

        [ObservableProperty]
        private ObservableCollection<string> errors = new();

        [ObservableProperty]
        private ObservableCollection<string> warnings = new();

        public ImportMappingViewModel(string filePath)
        {
            this.filePath = filePath ?? string.Empty;
            preview = ExcelImportService.BuildPreview(this.filePath);
            LoadPreview(preview);
        }

        public string FilePath => filePath;

        public int? DataStartRowIndex => preview.DataStartRowIndex;

        public IReadOnlyDictionary<ExcelImportService.ImportField, string> ConfirmedMap => confirmedMap;

        public string? HeaderSignature => preview.HeaderSignature;

        private static IReadOnlyList<TargetFieldOption> BuildTargetFieldOptions()
        {
            var options = new List<TargetFieldOption>
            {
                new("(Không dùng)", null)
            };

            foreach (var field in Enum.GetValues<ExcelImportService.ImportField>())
            {
                options.Add(new(field.ToString(), field));
            }

            return options;
        }

        partial void OnSelectedSheetNameChanged(string value)
        {
            if (suppressSheetReload)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, preview.SelectedSheetName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                LoadPreview(ExcelImportService.BuildPreview(filePath, value));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                PreviewWarningMessage = ex.Message;
            }
        }

        [RelayCommand]
        private void AutoMapAgain()
        {
            ApplySuggestedMapping();
            Errors = new ObservableCollection<string>();
            Warnings = new ObservableCollection<string>();
        }

        [RelayCommand]
        private void Validate()
        {
            _ = TryValidate(out _, out var errors, out var warnings);
            Errors = new ObservableCollection<string>(errors);
            Warnings = new ObservableCollection<string>(warnings);
        }

        [RelayCommand]
        private void Import()
        {
            if (!TryValidate(out var map, out var errors, out var warnings))
            {
                Errors = new ObservableCollection<string>(errors);
                Warnings = new ObservableCollection<string>(warnings);
                return;
            }

            confirmedMap = map;
            Errors = new ObservableCollection<string>();
            Warnings = new ObservableCollection<string>(warnings);
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }

        public void SaveProfileIfNeeded()
        {
            if (!SaveProfile ||
                confirmedMap.Count == 0 ||
                string.IsNullOrWhiteSpace(preview.HeaderSignature) ||
                string.IsNullOrWhiteSpace(preview.SelectedSheetName))
            {
                return;
            }

            var profileName = string.IsNullOrWhiteSpace(ProfileName) ? preview.SelectedSheetName : ProfileName.Trim();
            var sheetName = preview.SelectedSheetName;
            var signature = preview.HeaderSignature!;

            var profiles = LoadProfiles();
            var existingIndex = profiles.FindIndex(p =>
                string.Equals(p.HeaderSignature, signature, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(p.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(p.ProfileName, profileName, StringComparison.OrdinalIgnoreCase));

            var stored = new ImportMappingProfile(
                profileName,
                sheetName,
                signature,
                new Dictionary<ExcelImportService.ImportField, string>(confirmedMap),
                DateTime.UtcNow);

            if (existingIndex >= 0)
            {
                profiles[existingIndex] = stored;
            }
            else
            {
                profiles.Add(stored);
            }

            SaveProfiles(profiles);
        }

        private void LoadPreview(ImportPreviewResult next)
        {
            preview = next;
            PreviewView = preview.PreviewTable.DefaultView;
            PreviewWarningMessage = preview.WarningMessage ?? string.Empty;

            suppressSheetReload = true;
            try
            {
                SheetNames = new ObservableCollection<string>(preview.SheetNames);
                SelectedSheetName = preview.SelectedSheetName ?? string.Empty;
            }
            finally
            {
                suppressSheetReload = false;
            }

            var rows = preview.Columns.Select(c => new ImportColumnMappingRowViewModel(c)).ToList();
            ColumnMappings = new ObservableCollection<ImportColumnMappingRowViewModel>(rows);

            ProfileName = string.IsNullOrWhiteSpace(ProfileName) ? preview.SelectedSheetName : ProfileName;

            ApplyProfileOrSuggestions();
            Errors = new ObservableCollection<string>();
            Warnings = new ObservableCollection<string>();
        }

        private void ApplyProfileOrSuggestions()
        {
            if (!string.IsNullOrWhiteSpace(preview.HeaderSignature))
            {
                var matched = FindMatchingProfile(preview.HeaderSignature!, preview.SelectedSheetName);
                if (matched != null)
                {
                    ApplyProfile(matched);
                    return;
                }
            }

            ApplySuggestedMapping();
        }

        private void ApplyProfile(ImportMappingProfile profile)
        {
            if (profile == null)
            {
                return;
            }

            ProfileName = profile.ProfileName;

            foreach (var row in ColumnMappings)
            {
                row.TargetField = null;
            }

            foreach (var pair in profile.ConfirmedMap)
            {
                var field = pair.Key;
                var column = pair.Value;

                if (string.IsNullOrWhiteSpace(column))
                {
                    continue;
                }

                var match = ColumnMappings.FirstOrDefault(r => string.Equals(r.ColumnLetter, column, StringComparison.OrdinalIgnoreCase));
                if (match != null)
                {
                    match.TargetField = field;
                }
            }
        }

        private void ApplySuggestedMapping()
        {
            foreach (var row in ColumnMappings)
            {
                row.TargetField = null;
            }

            var bestByField = ColumnMappings
                .Where(r => r.SuggestedField != null && r.SuggestedScore >= AutoMapThreshold)
                .GroupBy(r => r.SuggestedField!.Value)
                .Select(group =>
                {
                    var best = group
                        .OrderByDescending(r => r.SuggestedScore)
                        .ThenBy(r => GetColumnIndex(r.ColumnLetter))
                        .First();

                    return (Field: group.Key, Column: best.ColumnLetter);
                })
                .ToList();

            foreach (var item in bestByField)
            {
                var row = ColumnMappings.FirstOrDefault(r => string.Equals(r.ColumnLetter, item.Column, StringComparison.OrdinalIgnoreCase));
                if (row != null)
                {
                    row.TargetField = item.Field;
                }
            }

            if (preview.HeaderRowIndex == null)
            {
                ApplyTemplateFallbackMapping();
            }
        }

        private void ApplyTemplateFallbackMapping()
        {
            var templateMap = GetTemplateFallbackMap();
            var alreadyMappedFields = new HashSet<ExcelImportService.ImportField>(
                ColumnMappings.Where(r => r.TargetField != null).Select(r => r.TargetField!.Value));

            foreach (var row in ColumnMappings)
            {
                if (row.TargetField != null)
                {
                    continue;
                }

                if (!templateMap.TryGetValue(row.ColumnLetter, out var field))
                {
                    continue;
                }

                if (alreadyMappedFields.Contains(field))
                {
                    continue;
                }

                row.TargetField = field;
                alreadyMappedFields.Add(field);
            }
        }

        private static Dictionary<string, ExcelImportService.ImportField> GetTemplateFallbackMap()
        {
            return new Dictionary<string, ExcelImportService.ImportField>(StringComparer.OrdinalIgnoreCase)
            {
                ["A"] = ExcelImportService.ImportField.SequenceNumber,
                ["B"] = ExcelImportService.ImportField.Name,
                ["C"] = ExcelImportService.ImportField.GroupName,
                ["D"] = ExcelImportService.ImportField.Address,
                ["E"] = ExcelImportService.ImportField.HouseholdPhone,
                ["F"] = ExcelImportService.ImportField.RepresentativeName,
                ["G"] = ExcelImportService.ImportField.Phone,
                ["H"] = ExcelImportService.ImportField.BuildingName,
                ["J"] = ExcelImportService.ImportField.MeterNumber,
                ["K"] = ExcelImportService.ImportField.Category,
                ["L"] = ExcelImportService.ImportField.Location,
                ["M"] = ExcelImportService.ImportField.Substation,
                ["N"] = ExcelImportService.ImportField.Page,
                ["O"] = ExcelImportService.ImportField.CurrentIndex,
                ["P"] = ExcelImportService.ImportField.PreviousIndex,
                ["Q"] = ExcelImportService.ImportField.Multiplier,
                ["S"] = ExcelImportService.ImportField.SubsidizedKwh,
                ["U"] = ExcelImportService.ImportField.UnitPrice,
                ["W"] = ExcelImportService.ImportField.PerformedBy
            };
        }

        private bool TryValidate(
            out Dictionary<ExcelImportService.ImportField, string> map,
            out List<string> errors,
            out List<string> warnings)
        {
            map = new Dictionary<ExcelImportService.ImportField, string>();
            errors = new List<string>();
            warnings = new List<string>();

            foreach (var group in ColumnMappings.Where(r => r.TargetField != null).GroupBy(r => r.TargetField!.Value))
            {
                if (group.Count() > 1)
                {
                    var cols = string.Join(", ", group.Select(r => r.ColumnLetter));
                    errors.Add($"Trung mapping: '{group.Key}' duoc chon boi nhieu cot ({cols}).");
                }
            }

            foreach (var row in ColumnMappings)
            {
                if (row.TargetField == null)
                {
                    continue;
                }

                map[row.TargetField.Value] = row.ColumnLetter;
            }

            if (!map.ContainsKey(ExcelImportService.ImportField.Name))
            {
                errors.Add("Bat buoc: phai map cot 'Name'.");
            }

            var hasAnyKeyField =
                map.ContainsKey(ExcelImportService.ImportField.MeterNumber) ||
                map.ContainsKey(ExcelImportService.ImportField.CurrentIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.PreviousIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.UnitPrice);

            if (!hasAnyKeyField)
            {
                errors.Add("Bat buoc: phai map it nhat 1 trong (MeterNumber/CurrentIndex/PreviousIndex/UnitPrice).");
            }

            if (!map.ContainsKey(ExcelImportService.ImportField.UnitPrice))
            {
                warnings.Add("Thieu mapping: UnitPrice (Don gia).");
            }

            if (!map.ContainsKey(ExcelImportService.ImportField.Multiplier))
            {
                warnings.Add("Thieu mapping: Multiplier (He so).");
            }

            warnings.AddRange(ValidateSampleData(map));

            return errors.Count == 0;
        }

        private IEnumerable<string> ValidateSampleData(Dictionary<ExcelImportService.ImportField, string> map)
        {
            if (preview.SampleRows == null || preview.SampleRows.Count == 0 || map.Count == 0)
            {
                yield break;
            }

            map.TryGetValue(ExcelImportService.ImportField.CurrentIndex, out var currentCol);
            map.TryGetValue(ExcelImportService.ImportField.PreviousIndex, out var previousCol);
            map.TryGetValue(ExcelImportService.ImportField.Multiplier, out var multiplierCol);
            map.TryGetValue(ExcelImportService.ImportField.UnitPrice, out var unitPriceCol);

            foreach (var row in preview.SampleRows.Take(5))
            {
                if (!string.IsNullOrWhiteSpace(unitPriceCol) &&
                    row.Cells.TryGetValue(unitPriceCol, out var unitPriceText) &&
                    !string.IsNullOrWhiteSpace(unitPriceText) &&
                    !decimal.TryParse(unitPriceText, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
                {
                    yield return $"Dong {row.RowIndex}: Don gia '{unitPriceText}' khong doc duoc.";
                }

                if (!string.IsNullOrWhiteSpace(multiplierCol) &&
                    row.Cells.TryGetValue(multiplierCol, out var multiplierText) &&
                    TryParseDecimalInvariant(multiplierText, out var multiplier) &&
                    multiplier <= 0)
                {
                    yield return $"Dong {row.RowIndex}: He so <= 0 (se tu dong set = 1 khi import).";
                }

                if (!string.IsNullOrWhiteSpace(currentCol) &&
                    !string.IsNullOrWhiteSpace(previousCol) &&
                    row.Cells.TryGetValue(currentCol, out var currentText) &&
                    row.Cells.TryGetValue(previousCol, out var previousText) &&
                    TryParseDecimalInvariant(currentText, out var current) &&
                    TryParseDecimalInvariant(previousText, out var previous) &&
                    current < previous)
                {
                    yield return $"Dong {row.RowIndex}: Chi so moi < chi so cu (moi={current:0.##}, cu={previous:0.##}).";
                }
            }
        }

        private static bool TryParseDecimalInvariant(string? text, out decimal value)
        {
            value = 0m;

            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private ImportMappingProfile? FindMatchingProfile(string signature, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(signature) || string.IsNullOrWhiteSpace(sheetName))
            {
                return null;
            }

            var profiles = LoadProfiles();
            return profiles
                .Where(p =>
                    string.Equals(p.HeaderSignature, signature, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(p.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(p => p.LastUsedAt)
                .FirstOrDefault();
        }

        private static string GetProfilesPath()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var folder = Path.Combine(documents, "ElectricCalculation");
            Directory.CreateDirectory(folder);
            return Path.Combine(folder, "import_mapping_profiles.json");
        }

        private static List<ImportMappingProfile> LoadProfiles()
        {
            try
            {
                var path = GetProfilesPath();
                if (!File.Exists(path))
                {
                    return new List<ImportMappingProfile>();
                }

                var json = File.ReadAllText(path);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var profiles = JsonSerializer.Deserialize<List<ImportMappingProfile>>(json, options) ?? new List<ImportMappingProfile>();
                profiles.RemoveAll(p => p == null || string.IsNullOrWhiteSpace(p.HeaderSignature) || p.ConfirmedMap == null);
                return profiles;
            }
            catch
            {
                return new List<ImportMappingProfile>();
            }
        }

        private static void SaveProfiles(List<ImportMappingProfile> profiles)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Converters = { new JsonStringEnumConverter() }
                };

                var path = GetProfilesPath();
                var json = JsonSerializer.Serialize(profiles ?? new List<ImportMappingProfile>(), options);
                File.WriteAllText(path, json);
            }
            catch
            {
                // ignore profile save errors
            }
        }

        private static int GetColumnIndex(string columnLetters)
        {
            if (string.IsNullOrWhiteSpace(columnLetters))
            {
                return int.MaxValue;
            }

            var value = 0;
            foreach (var ch in columnLetters.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z')
                {
                    break;
                }

                value = value * 26 + (ch - 'A' + 1);
            }

            return value == 0 ? int.MaxValue : value;
        }

        public sealed record TargetFieldOption(string Display, ExcelImportService.ImportField? Field);

        private sealed record ImportMappingProfile(
            string ProfileName,
            string SheetName,
            string HeaderSignature,
            Dictionary<ExcelImportService.ImportField, string> ConfirmedMap,
            DateTime LastUsedAt);

        public sealed partial class ImportColumnMappingRowViewModel : ObservableObject
        {
            public ImportColumnMappingRowViewModel(ImportColumnPreview preview)
            {
                ColumnLetter = preview.ColumnLetter ?? string.Empty;
                HeaderText = preview.HeaderText ?? string.Empty;
                SampleValuesText = preview.SampleValues == null || preview.SampleValues.Count == 0
                    ? string.Empty
                    : string.Join(" | ", preview.SampleValues);
                SuggestedField = preview.SuggestedField;
                SuggestedScore = preview.SuggestedScore;
            }

            public string ColumnLetter { get; }

            public string HeaderText { get; }

            public string SampleValuesText { get; }

            public ExcelImportService.ImportField? SuggestedField { get; }

            public double SuggestedScore { get; }

            public string SuggestedDisplay =>
                SuggestedField == null || SuggestedScore <= 0
                    ? string.Empty
                    : $"{SuggestedField} ({SuggestedScore:0.00})";

            [ObservableProperty]
            private ExcelImportService.ImportField? targetField;
        }
    }
}
