using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed partial class ImportWizardViewModel : ObservableObject
    {
        private readonly string filePath;
        private ImportPreviewResult preview;
        private bool suppressSheetReload;

        private IReadOnlyList<FieldOption> fieldOptions = Array.Empty<FieldOption>();
        private IReadOnlyDictionary<string, ColumnMappingViewModel> columnMappingsByLetter =
            new Dictionary<string, ColumnMappingViewModel>(StringComparer.OrdinalIgnoreCase);

        private List<Customer> importedCustomers = new();
        private ImportRunReport? importReport;

        public ImportWizardViewModel(string filePath)
        {
            this.filePath = filePath ?? string.Empty;
            preview = ExcelImportService.BuildPreview(this.filePath);
            LoadPreview(preview);
            CurrentStep = 0;
        }

        public string FilePath => filePath;

        public IReadOnlyList<Customer> ImportedCustomers => importedCustomers;

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsStep1))]
        [NotifyPropertyChangedFor(nameof(IsStep2))]
        [NotifyPropertyChangedFor(nameof(IsStep3))]
        [NotifyPropertyChangedFor(nameof(IsNotStep3))]
        [NotifyPropertyChangedFor(nameof(CanGoBack))]
        [NotifyPropertyChangedFor(nameof(CanGoNext))]
        private int currentStep;

        public bool IsStep1 => CurrentStep == 0;

        public bool IsStep2 => CurrentStep == 1;

        public bool IsStep3 => CurrentStep == 2;

        public bool IsNotStep3 => CurrentStep != 2;

        public bool CanGoBack => CurrentStep > 0 && !IsImporting;

        public bool CanGoNext => CurrentStep < 2 && !IsImporting;

        [ObservableProperty]
        private string previewWarningMessage = string.Empty;

        [ObservableProperty]
        private ObservableCollection<string> sheetNames = new();

        [ObservableProperty]
        private string selectedSheetName = string.Empty;

        [ObservableProperty]
        private DataView? previewView;

        [ObservableProperty]
        private ObservableCollection<PresetOption> presetOptions = new();

        [ObservableProperty]
        private PresetOption? selectedPreset;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasPresetToast))]
        private string presetToastMessage = string.Empty;

        public bool HasPresetToast => !string.IsNullOrWhiteSpace(PresetToastMessage);

        [ObservableProperty]
        private ObservableCollection<ColumnMappingViewModel> columnMappings = new();

        public IReadOnlyList<FieldOption> FieldOptions => fieldOptions;

        [ObservableProperty]
        private bool isNameMapped;

        [ObservableProperty]
        private bool isAnyKeyFieldMapped;

        [ObservableProperty]
        private ObservableCollection<string> errorMessages = new();

        [ObservableProperty]
        private ObservableCollection<string> warningMessages = new();

        public bool HasNoErrors => ErrorMessages.Count == 0;

        public bool HasNoWarnings => WarningMessages.Count == 0;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(CanImport))]
        [NotifyPropertyChangedFor(nameof(CanFinish))]
        [NotifyPropertyChangedFor(nameof(CanGoBack))]
        [NotifyPropertyChangedFor(nameof(CanGoNext))]
        private bool isImporting;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(CanFinish))]
        private bool hasImportResult;

        [ObservableProperty]
        private string importStatusText = "Bấm “Kiểm tra” để xem lỗi trước khi nhập.";

        [ObservableProperty]
        private ObservableCollection<string> importWarningsTop = new();

        public int ReportTotalRows => importReport?.TotalRows ?? 0;

        public int ReportImportedRows => importReport?.ImportedRows ?? 0;

        public int ReportSkippedRows => importReport?.SkippedRows ?? 0;

        public int ReportWarningCount => importReport?.WarningCount ?? 0;

        [ObservableProperty]
        private bool savePreset = true;

        [ObservableProperty]
        private string presetName = string.Empty;

        public bool CanSavePreset => !string.IsNullOrWhiteSpace(preview.HeaderSignature);

        public bool CanImport => !IsImporting && !HasBlockingErrors();

        public bool CanFinish => HasImportResult && !IsImporting;

        partial void OnSelectedSheetNameChanged(string value)
        {
            if (suppressSheetReload)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(value) ||
                string.Equals(value, preview.SelectedSheetName, StringComparison.OrdinalIgnoreCase))
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

        partial void OnSelectedPresetChanged(PresetOption? value)
        {
            if (value == null)
            {
                return;
            }

            PresetToastMessage = string.Empty;

            if (value.Profile == null)
            {
                ApplyAutoMapping();
                if (string.IsNullOrWhiteSpace(PresetName))
                {
                    PresetName = preview.SelectedSheetName ?? string.Empty;
                }
                return;
            }

            PresetName = value.Profile.ProfileName;
            ApplyProfile(value.Profile);
        }

        [RelayCommand]
        private void Next()
        {
            if (CurrentStep >= 2)
            {
                return;
            }

            CurrentStep += 1;
        }

        [RelayCommand]
        private void Back()
        {
            if (CurrentStep <= 0)
            {
                return;
            }

            CurrentStep -= 1;
        }

        [RelayCommand]
        private void ResetToAuto()
        {
            SelectedPreset = PresetOptions.FirstOrDefault(p => p.Profile == null);
            ApplyAutoMapping();
        }

        [RelayCommand]
        private void UnlockColumn(ColumnMappingViewModel? column)
        {
            if (column == null)
            {
                return;
            }

            column.IsLocked = false;
        }

        [RelayCommand]
        private void Validate()
        {
            var map = BuildConfirmedMap();
            ErrorMessages = new ObservableCollection<string>(BuildValidationErrors(map));
            WarningMessages = new ObservableCollection<string>(BuildValidationWarnings(map));

            ImportStatusText = ErrorMessages.Count == 0
                ? "OK. Bạn có thể bấm “Nhập dữ liệu”."
                : "Có lỗi. Hãy sửa mapping rồi bấm “Kiểm tra” lại.";

            OnPropertyChanged(nameof(HasNoErrors));
            OnPropertyChanged(nameof(HasNoWarnings));
            OnPropertyChanged(nameof(CanImport));
        }

        [RelayCommand]
        private async Task Import()
        {
            Validate();
            if (ErrorMessages.Count > 0)
            {
                return;
            }

            var map = BuildConfirmedMap();
            if (map.Count == 0)
            {
                ErrorMessages = new ObservableCollection<string>(new[] { "❌ Bạn chưa chọn cột nào để nhập." });
                return;
            }

            IsImporting = true;
            ImportStatusText = "Đang nhập dữ liệu…";
            HasImportResult = false;
            ImportWarningsTop = new ObservableCollection<string>();
            importedCustomers = new List<Customer>();
            importReport = null;

            try
            {
                var result = await Task.Run(() =>
                {
                    var list = ExcelImportService.ImportFromFile(
                        filePath,
                        preview.SelectedSheetName,
                        map,
                        preview.DataStartRowIndex,
                        out var warningMessage,
                        out var report);

                    return (Customers: list.ToList(), Report: report, WarningMessage: warningMessage);
                });

                importedCustomers = result.Customers;
                importReport = result.Report;

                if (!string.IsNullOrWhiteSpace(result.WarningMessage))
                {
                    var warnings = WarningMessages.ToList();
                    warnings.Add($"⚠️ {result.WarningMessage}");
                    WarningMessages = new ObservableCollection<string>(warnings);
                }

                if (SavePreset)
                {
                    SavePresetProfile(map);
                }

                ImportWarningsTop = new ObservableCollection<string>((importReport?.Warnings ?? Array.Empty<string>()).Take(100));
                HasImportResult = true;
                ImportStatusText = "Đã import xong. Xem kết quả bên dưới.";

                OnPropertyChanged(nameof(ReportTotalRows));
                OnPropertyChanged(nameof(ReportImportedRows));
                OnPropertyChanged(nameof(ReportSkippedRows));
                OnPropertyChanged(nameof(ReportWarningCount));
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                ErrorMessages = new ObservableCollection<string>(new[] { $"❌ Lỗi import: {ex.Message}" });
                ImportStatusText = "Không thể import. Hãy kiểm tra lại file Excel.";
            }
            finally
            {
                IsImporting = false;
                OnPropertyChanged(nameof(CanFinish));
                OnPropertyChanged(nameof(CanImport));
                OnPropertyChanged(nameof(ImportedCustomers));
            }
        }

        [RelayCommand]
        private void Finish()
        {
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }

        public Dictionary<ExcelImportService.ImportField, string> BuildConfirmedMap()
        {
            var result = new Dictionary<ExcelImportService.ImportField, string>();

            foreach (var column in ColumnMappings)
            {
                if (column.SelectedField == null)
                {
                    continue;
                }

                result[column.SelectedField.Value] = column.ColumnLetter;
            }

            return result;
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

            BuildFieldOptions();
            BuildColumnMappings();
            LoadPresetOptionsAndApplyDefault();

            ErrorMessages = new ObservableCollection<string>();
            WarningMessages = new ObservableCollection<string>();
            ImportWarningsTop = new ObservableCollection<string>();
            importedCustomers = new List<Customer>();
            importReport = null;
            HasImportResult = false;
            PresetName = string.IsNullOrWhiteSpace(PresetName) ? (preview.SelectedSheetName ?? string.Empty) : PresetName;
            ImportStatusText = "Bấm “Kiểm tra” để xem lỗi trước khi nhập.";

            UpdateConflictAndMissingFlags();
            OnPropertyChanged(nameof(CanImport));
            OnPropertyChanged(nameof(CanSavePreset));
        }

        public bool TryGetColumnMapping(string columnLetter, out ColumnMappingViewModel? mapping)
        {
            mapping = null;
            if (string.IsNullOrWhiteSpace(columnLetter))
            {
                return false;
            }

            return columnMappingsByLetter.TryGetValue(columnLetter.Trim().ToUpperInvariant(), out mapping);
        }

        private void BuildFieldOptions()
        {
            var importantFields = new[]
            {
                ExcelImportService.ImportField.Name,
                ExcelImportService.ImportField.MeterNumber,
                ExcelImportService.ImportField.PreviousIndex,
                ExcelImportService.ImportField.CurrentIndex,
                ExcelImportService.ImportField.Multiplier,
                ExcelImportService.ImportField.UnitPrice
            };

            var allFields = Enum.GetValues<ExcelImportService.ImportField>();
            var ordered = importantFields.Concat(allFields.Where(f => !importantFields.Contains(f))).ToList();

            var options = new List<FieldOption> { new(null, "(Không dùng)") };
            foreach (var field in ordered)
            {
                var label = GetFieldLabel(field);
                if (field == ExcelImportService.ImportField.Name)
                {
                    label = $"{label} (bắt buộc)";
                }

                options.Add(new FieldOption(field, label));
            }

            fieldOptions = options;
            OnPropertyChanged(nameof(FieldOptions));
        }

        private void BuildColumnMappings()
        {
            var columns = (preview.Columns ?? Array.Empty<ImportColumnPreview>())
                .Where(c => !string.IsNullOrWhiteSpace(c.ColumnLetter))
                .Select(c => c with { ColumnLetter = c.ColumnLetter.Trim().ToUpperInvariant() })
                .OrderBy(c => GetColumnIndex(c.ColumnLetter))
                .ToList();

            ColumnMappings = new ObservableCollection<ColumnMappingViewModel>(columns.Select(CreateColumnMapping));

            columnMappingsByLetter = ColumnMappings
                .Where(c => !string.IsNullOrWhiteSpace(c.ColumnLetter))
                .GroupBy(c => c.ColumnLetter, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToDictionary(c => c.ColumnLetter, c => c, StringComparer.OrdinalIgnoreCase);

            foreach (var column in ColumnMappings)
            {
                column.SelectionChanged += (_, _) =>
                {
                    UpdateConflictAndMissingFlags();
                    OnPropertyChanged(nameof(CanImport));
                };
            }
        }

        private ColumnMappingViewModel CreateColumnMapping(ImportColumnPreview column)
        {
            var sample = column.SampleValues == null || column.SampleValues.Count == 0
                ? string.Empty
                : string.Join(" | ", column.SampleValues.Where(v => !string.IsNullOrWhiteSpace(v)).Take(3));

            var suggestedFieldLabel = column.SuggestedField != null ? GetFieldLabel(column.SuggestedField.Value) : null;
            return new ColumnMappingViewModel(
                columnLetter: column.ColumnLetter.Trim().ToUpperInvariant(),
                headerText: column.HeaderText ?? string.Empty,
                samplePreview: sample,
                suggestedField: column.SuggestedField,
                suggestedFieldLabel: suggestedFieldLabel,
                suggestedScore: column.SuggestedScore);
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

        private void ApplyTemplateFallbackForUnselected()
        {
            var template = GetTemplateFallbackFieldToColumn();
            var usedFields = new HashSet<ExcelImportService.ImportField>(
                ColumnMappings.Where(c => c.SelectedField != null).Select(c => c.SelectedField!.Value));

            foreach (var pair in template)
            {
                if (usedFields.Contains(pair.Key))
                {
                    continue;
                }

                var col = (pair.Value ?? string.Empty).Trim().ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(col))
                {
                    continue;
                }

                if (!columnMappingsByLetter.TryGetValue(col, out var mapping))
                {
                    continue;
                }

                if (mapping.SelectedField != null)
                {
                    continue;
                }

                mapping.SelectedField = pair.Key;
                usedFields.Add(pair.Key);
            }
        }

        private static Dictionary<ExcelImportService.ImportField, string> GetTemplateFallbackFieldToColumn()
        {
            return new Dictionary<ExcelImportService.ImportField, string>
            {
                [ExcelImportService.ImportField.SequenceNumber] = "A",
                [ExcelImportService.ImportField.Name] = "B",
                [ExcelImportService.ImportField.GroupName] = "C",
                [ExcelImportService.ImportField.Address] = "D",
                [ExcelImportService.ImportField.HouseholdPhone] = "E",
                [ExcelImportService.ImportField.RepresentativeName] = "F",
                [ExcelImportService.ImportField.Phone] = "G",
                [ExcelImportService.ImportField.BuildingName] = "H",
                [ExcelImportService.ImportField.MeterNumber] = "J",
                [ExcelImportService.ImportField.Category] = "K",
                [ExcelImportService.ImportField.Location] = "L",
                [ExcelImportService.ImportField.Substation] = "M",
                [ExcelImportService.ImportField.Page] = "N",
                [ExcelImportService.ImportField.CurrentIndex] = "O",
                [ExcelImportService.ImportField.PreviousIndex] = "P",
                [ExcelImportService.ImportField.Multiplier] = "Q",
                [ExcelImportService.ImportField.SubsidizedKwh] = "S",
                [ExcelImportService.ImportField.UnitPrice] = "U",
                [ExcelImportService.ImportField.PerformedBy] = "W"
            };
        }

        private void LoadPresetOptionsAndApplyDefault()
        {
            PresetToastMessage = string.Empty;

            var options = new List<PresetOption> { new("Auto (gợi ý)", null) };
            var signature = preview.HeaderSignature;
            var sheetName = preview.SelectedSheetName;

            var profiles = string.IsNullOrWhiteSpace(signature) || string.IsNullOrWhiteSpace(sheetName)
                ? new List<ImportMappingProfile>()
                : LoadProfiles()
                    .Where(p =>
                        string.Equals(p.HeaderSignature, signature, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(p.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(p => p.LastUsedAt)
                    .ToList();

            foreach (var profile in profiles)
            {
                options.Add(new PresetOption(profile.ProfileName, profile));
            }

            PresetOptions = new ObservableCollection<PresetOption>(options);
            SelectedPreset = null;

            var lastUsed = profiles.FirstOrDefault();
            if (lastUsed != null)
            {
                SelectedPreset = PresetOptions.FirstOrDefault(p =>
                                   p.Profile != null &&
                                   string.Equals(p.Profile.ProfileName, lastUsed.ProfileName, StringComparison.OrdinalIgnoreCase)) ??
                               PresetOptions[0];

                PresetName = lastUsed.ProfileName;
                PresetToastMessage = "Đã áp dụng preset lần trước.";
                return;
            }

            SelectedPreset = PresetOptions[0];
        }

        private void ApplyAutoMapping()
        {
            foreach (var column in ColumnMappings)
            {
                column.ResetSelection();
            }

            var candidates = ColumnMappings
                .Where(c => c.SuggestedField != null && c.SuggestedScore >= ColumnMappingViewModel.AutoSelectThreshold)
                .OrderByDescending(c => c.SuggestedScore)
                .ThenBy(c => GetColumnIndex(c.ColumnLetter))
                .ToList();

            var usedFields = new HashSet<ExcelImportService.ImportField>();
            foreach (var column in candidates)
            {
                var field = column.SuggestedField!.Value;
                if (usedFields.Contains(field))
                {
                    continue;
                }

                column.SelectedField = field;
                column.IsLocked = column.SelectedField == column.SuggestedField &&
                                  column.SuggestedScore >= ColumnMappingViewModel.LockThreshold;
                usedFields.Add(field);
            }

            if (preview.HeaderRowIndex == null)
            {
                ApplyTemplateFallbackForUnselected();
            }

            UpdateConflictAndMissingFlags();
            OnPropertyChanged(nameof(CanImport));
        }

        private void ApplyProfile(ImportMappingProfile profile)
        {
            foreach (var column in ColumnMappings)
            {
                column.ResetSelection();
            }

            foreach (var pair in profile.ConfirmedMap ?? new Dictionary<ExcelImportService.ImportField, string>())
            {
                var col = (pair.Value ?? string.Empty).Trim().ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(col))
                {
                    continue;
                }

                if (columnMappingsByLetter.TryGetValue(col, out var mapping))
                {
                    mapping.SelectedField = pair.Key;
                }
            }

            foreach (var column in ColumnMappings)
            {
                column.IsLocked = column.SelectedField != null &&
                                  column.SelectedField == column.SuggestedField &&
                                  column.SuggestedScore >= ColumnMappingViewModel.LockThreshold;
            }

            if (preview.HeaderRowIndex == null)
            {
                ApplyTemplateFallbackForUnselected();
            }

            UpdateConflictAndMissingFlags();
            OnPropertyChanged(nameof(CanImport));
        }

        private void UpdateConflictAndMissingFlags()
        {
            foreach (var column in ColumnMappings)
            {
                column.IsConflict = false;
                column.IsRequiredMissing = false;
            }

            var conflictGroups = ColumnMappings
                .Where(c => c.SelectedField != null)
                .GroupBy(c => c.SelectedField!.Value)
                .Where(g => g.Count() > 1)
                .ToList();

            foreach (var group in conflictGroups)
            {
                foreach (var column in group)
                {
                    column.IsConflict = true;
                }
            }

            var selectedFields = ColumnMappings
                .Where(c => c.SelectedField != null)
                .Select(c => c.SelectedField!.Value)
                .ToList();

            IsNameMapped = selectedFields.Contains(ExcelImportService.ImportField.Name);

            var keyFields = new[]
            {
                ExcelImportService.ImportField.MeterNumber,
                ExcelImportService.ImportField.PreviousIndex,
                ExcelImportService.ImportField.CurrentIndex,
                ExcelImportService.ImportField.UnitPrice
            };

            IsAnyKeyFieldMapped = selectedFields.Any(f => keyFields.Contains(f));

            if (!IsNameMapped)
            {
                var candidate = ColumnMappings
                    .Where(c => c.SuggestedField == ExcelImportService.ImportField.Name && c.SuggestedScore > 0)
                    .OrderByDescending(c => c.SuggestedScore)
                    .ThenBy(c => GetColumnIndex(c.ColumnLetter))
                    .FirstOrDefault();

                if (candidate != null && candidate.SelectedField != ExcelImportService.ImportField.Name)
                {
                    candidate.IsRequiredMissing = true;
                }
            }

            if (!IsAnyKeyFieldMapped)
            {
                foreach (var field in keyFields)
                {
                    var candidate = ColumnMappings
                        .Where(c => c.SuggestedField == field && c.SuggestedScore > 0)
                        .OrderByDescending(c => c.SuggestedScore)
                        .ThenBy(c => GetColumnIndex(c.ColumnLetter))
                        .FirstOrDefault();

                    if (candidate != null && candidate.SelectedField != field)
                    {
                        candidate.IsRequiredMissing = true;
                    }
                }
            }

            OnPropertyChanged(nameof(CanImport));
        }

        private bool HasBlockingErrors()
        {
            if (ColumnMappings.Any(c => c.IsConflict))
            {
                return true;
            }

            var map = BuildConfirmedMap();
            if (!map.ContainsKey(ExcelImportService.ImportField.Name))
            {
                return true;
            }

            var hasAnyKeyField =
                map.ContainsKey(ExcelImportService.ImportField.MeterNumber) ||
                map.ContainsKey(ExcelImportService.ImportField.CurrentIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.PreviousIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.UnitPrice);

            return !hasAnyKeyField;
        }

        private List<string> BuildValidationErrors(Dictionary<ExcelImportService.ImportField, string> map)
        {
            var errors = new List<string>();

            if (!map.ContainsKey(ExcelImportService.ImportField.Name))
            {
                errors.Add("❌ Thiếu cột bắt buộc: Tên khách");
            }

            var hasAnyKeyField =
                map.ContainsKey(ExcelImportService.ImportField.MeterNumber) ||
                map.ContainsKey(ExcelImportService.ImportField.CurrentIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.PreviousIndex) ||
                map.ContainsKey(ExcelImportService.ImportField.UnitPrice);

            if (!hasAnyKeyField)
            {
                errors.Add("❌ Cần chọn ít nhất 1 trong: Số công tơ / Chỉ số cũ / Chỉ số mới / Đơn giá");
            }

            var conflictGroups = ColumnMappings
                .Where(c => c.SelectedField != null)
                .GroupBy(c => c.SelectedField!.Value)
                .Where(g => g.Count() > 1)
                .ToList();

            foreach (var group in conflictGroups)
            {
                var fieldLabel = GetFieldLabel(group.Key);
                var columns = group.Select(GetColumnDisplay).ToList();
                errors.Add($"❌ Trùng mapping: {fieldLabel} đang được chọn ở {string.Join(", ", columns)}");
            }

            return errors;
        }

        private List<string> BuildValidationWarnings(Dictionary<ExcelImportService.ImportField, string> map)
        {
            var warnings = new List<string>();

            if (!map.ContainsKey(ExcelImportService.ImportField.UnitPrice))
            {
                warnings.Add("⚠️ Thiếu cột: Đơn giá (nếu Excel có)");
            }

            if (!map.ContainsKey(ExcelImportService.ImportField.Multiplier))
            {
                warnings.Add("⚠️ Thiếu cột: Hệ số (nếu Excel có)");
            }

            warnings.AddRange(BuildSampleDataWarnings(map));
            return warnings;
        }

        private IEnumerable<string> BuildSampleDataWarnings(Dictionary<ExcelImportService.ImportField, string> map)
        {
            if (preview.SampleRows == null || preview.SampleRows.Count == 0 || map.Count == 0)
            {
                yield break;
            }

            string? unitPriceColumn = null;
            if (map.TryGetValue(ExcelImportService.ImportField.UnitPrice, out var unitPriceKey) &&
                !string.IsNullOrWhiteSpace(unitPriceKey))
            {
                unitPriceColumn = unitPriceKey;
            }

            string? multiplierColumn = null;
            if (map.TryGetValue(ExcelImportService.ImportField.Multiplier, out var multiplierKey) &&
                !string.IsNullOrWhiteSpace(multiplierKey))
            {
                multiplierColumn = multiplierKey;
            }

            string? currentColumn = null;
            if (map.TryGetValue(ExcelImportService.ImportField.CurrentIndex, out var currentKey) &&
                !string.IsNullOrWhiteSpace(currentKey))
            {
                currentColumn = currentKey;
            }

            string? previousColumn = null;
            if (map.TryGetValue(ExcelImportService.ImportField.PreviousIndex, out var previousKey) &&
                !string.IsNullOrWhiteSpace(previousKey))
            {
                previousColumn = previousKey;
            }

            var unitPriceFails = 0;
            var multiplierBad = 0;
            var indexBackwards = 0;

            foreach (var row in preview.SampleRows)
            {
                if (unitPriceColumn != null && row.Cells.TryGetValue(unitPriceColumn, out var unitPriceText))
                {
                    if (!string.IsNullOrWhiteSpace(unitPriceText) && !TryParseDecimal(unitPriceText, out _))
                    {
                        unitPriceFails++;
                    }
                }

                if (multiplierColumn != null && row.Cells.TryGetValue(multiplierColumn, out var multiplierText))
                {
                    if (TryParseDecimal(multiplierText, out var m) && m <= 0)
                    {
                        multiplierBad++;
                    }
                }

                if (currentColumn != null && previousColumn != null &&
                    row.Cells.TryGetValue(currentColumn, out var currentText) &&
                    row.Cells.TryGetValue(previousColumn, out var previousText) &&
                    TryParseDecimal(currentText, out var current) &&
                    TryParseDecimal(previousText, out var previous) &&
                    current < previous)
                {
                    indexBackwards++;
                }
            }

            if (unitPriceFails > 0)
            {
                yield return $"⚠️ Đơn giá: {unitPriceFails} dòng mẫu không đọc được số";
            }

            if (multiplierBad > 0)
            {
                yield return $"⚠️ Hệ số: {multiplierBad} dòng mẫu <= 0 (khi import sẽ tự đặt về 1)";
            }

            if (indexBackwards > 0)
            {
                yield return $"⚠️ Chỉ số: {indexBackwards} dòng mẫu (chỉ số mới < chỉ số cũ)";
            }
        }

        private static bool TryParseDecimal(string? text, out decimal value)
        {
            value = 0;
            var t = (text ?? string.Empty).Trim();
            if (t.Length == 0)
            {
                return false;
            }

            return decimal.TryParse(t, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
                   decimal.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        private void SavePresetProfile(Dictionary<ExcelImportService.ImportField, string> map)
        {
            if (string.IsNullOrWhiteSpace(preview.HeaderSignature) ||
                string.IsNullOrWhiteSpace(preview.SelectedSheetName))
            {
                return;
            }

            var signature = preview.HeaderSignature!;
            var sheetName = preview.SelectedSheetName;
            var profileName = string.IsNullOrWhiteSpace(PresetName)
                ? sheetName
                : PresetName.Trim();

            var stored = new ImportMappingProfile(
                ProfileName: profileName,
                SheetName: sheetName,
                HeaderSignature: signature,
                ConfirmedMap: new Dictionary<ExcelImportService.ImportField, string>(map),
                LastUsedAt: DateTime.UtcNow);

            var profiles = LoadProfiles();
            var index = profiles.FindIndex(p =>
                string.Equals(p.ProfileName, profileName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(p.SheetName, sheetName, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(p.HeaderSignature, signature, StringComparison.OrdinalIgnoreCase));

            if (index >= 0)
            {
                profiles[index] = stored;
            }
            else
            {
                profiles.Add(stored);
            }

            SaveProfiles(profiles);
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
                // ignore
            }
        }

        private static string GetFieldLabel(ExcelImportService.ImportField field)
        {
            return FieldLabels.TryGetValue(field, out var label) ? label : field.ToString();
        }

        private static string GetColumnDisplay(ColumnMappingViewModel column)
        {
            var header = (column.HeaderText ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(header) ? column.ColumnLetter : $"{column.ColumnLetter} - {header}";
        }

        private static readonly Dictionary<ExcelImportService.ImportField, string> FieldLabels = new()
        {
            [ExcelImportService.ImportField.SequenceNumber] = "Số thứ tự",
            [ExcelImportService.ImportField.Name] = "Tên khách",
            [ExcelImportService.ImportField.GroupName] = "Nhóm / Đơn vị",
            [ExcelImportService.ImportField.Address] = "Địa chỉ",
            [ExcelImportService.ImportField.MeterNumber] = "Số công tơ",
            [ExcelImportService.ImportField.CurrentIndex] = "Chỉ số mới",
            [ExcelImportService.ImportField.PreviousIndex] = "Chỉ số cũ",
            [ExcelImportService.ImportField.Multiplier] = "Hệ số",
            [ExcelImportService.ImportField.UnitPrice] = "Đơn giá",
            [ExcelImportService.ImportField.SubsidizedKwh] = "Bao cấp (kWh)",
            [ExcelImportService.ImportField.Phone] = "SĐT",
            [ExcelImportService.ImportField.HouseholdPhone] = "SĐT hộ",
            [ExcelImportService.ImportField.RepresentativeName] = "Đại diện",
            [ExcelImportService.ImportField.Location] = "Vị trí",
            [ExcelImportService.ImportField.Substation] = "TBA",
            [ExcelImportService.ImportField.Page] = "Trang",
            [ExcelImportService.ImportField.PerformedBy] = "Người ghi",
            [ExcelImportService.ImportField.BuildingName] = "Tòa / Mã sổ",
            [ExcelImportService.ImportField.Category] = "Loại"
        };
    }
}
