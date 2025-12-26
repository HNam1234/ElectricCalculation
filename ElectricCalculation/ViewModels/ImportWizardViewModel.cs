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

        private IReadOnlyList<ColumnOption> columnOptions = Array.Empty<ColumnOption>();
        private IReadOnlyDictionary<string, ColumnOption> columnOptionsByLetter =
            new Dictionary<string, ColumnOption>(StringComparer.OrdinalIgnoreCase);

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
        private ObservableCollection<FieldMappingRowViewModel> importantFieldMappings = new();

        [ObservableProperty]
        private ObservableCollection<FieldMappingRowViewModel> optionalFieldMappings = new();

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
        private void EditField(FieldMappingRowViewModel? row)
        {
            if (row == null)
            {
                return;
            }

            row.IsLocked = false;
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

            foreach (var row in GetAllRows())
            {
                var col = (row.SelectedColumn ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(col))
                {
                    continue;
                }

                result[row.Field] = col;
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

            BuildColumnOptions();
            BuildFieldRows();
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

        private void BuildColumnOptions()
        {
            var options = new List<ColumnOption>
            {
                new(null, string.Empty, "(Chưa chọn)"),
                new(string.Empty, string.Empty, "(Không dùng)")
            };

            foreach (var col in preview.Columns ?? Array.Empty<ImportColumnPreview>())
            {
                if (string.IsNullOrWhiteSpace(col.ColumnLetter))
                {
                    continue;
                }

                var letter = col.ColumnLetter.Trim().ToUpperInvariant();
                var header = col.HeaderText ?? string.Empty;
                var display = string.IsNullOrWhiteSpace(header) ? letter : $"{letter} - {header.Trim()}";

                var sample = col.SampleValues == null || col.SampleValues.Count == 0
                    ? string.Empty
                    : string.Join(" | ", col.SampleValues.Where(v => !string.IsNullOrWhiteSpace(v)).Take(3));

                options.Add(new ColumnOption(letter, header, display, sample));
            }

            columnOptions = options;
            columnOptionsByLetter = options
                .Where(o => !string.IsNullOrWhiteSpace(o.ColumnLetter))
                .GroupBy(o => o.ColumnLetter!, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToDictionary(o => o.ColumnLetter!, o => o, StringComparer.OrdinalIgnoreCase);
        }

        private void BuildFieldRows()
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

            var importantSet = new HashSet<ExcelImportService.ImportField>(importantFields);
            var allFields = Enum.GetValues<ExcelImportService.ImportField>().ToList();

            ImportantFieldMappings = new ObservableCollection<FieldMappingRowViewModel>(
                importantFields.Select(f => CreateRow(f, isRequired: f == ExcelImportService.ImportField.Name)));

            OptionalFieldMappings = new ObservableCollection<FieldMappingRowViewModel>(
                allFields.Where(f => !importantSet.Contains(f))
                         .Select(f => CreateRow(f, isRequired: false)));

            foreach (var row in GetAllRows())
            {
                row.SelectionChanged += (_, _) =>
                {
                    UpdateConflictAndMissingFlags();
                    OnPropertyChanged(nameof(CanImport));
                };
            }
        }

        private FieldMappingRowViewModel CreateRow(ExcelImportService.ImportField field, bool isRequired)
        {
            var label = FieldLabels.TryGetValue(field, out var l) ? l : field.ToString();
            var hint = isRequired ? "Bắt buộc" : "Tùy chọn";
            if (isRequired)
            {
                label = $"{label} *";
            }

            var suggested = GetBestSuggestedColumn(field);
            var suggestedColumn = suggested?.ColumnLetter?.Trim().ToUpperInvariant();
            var suggestedScore = suggested?.SuggestedScore ?? 0;

            return new FieldMappingRowViewModel(
                field: field,
                label: label,
                hint: hint,
                isRequired: isRequired,
                options: columnOptions,
                suggestedColumn: suggestedColumn,
                suggestedScore: suggestedScore);
        }

        private ImportColumnPreview? GetBestSuggestedColumn(ExcelImportService.ImportField field)
        {
            return (preview.Columns ?? Array.Empty<ImportColumnPreview>())
                .Where(c =>
                    c.SuggestedField != null &&
                    c.SuggestedField.Value == field &&
                    c.SuggestedScore > 0 &&
                    !string.IsNullOrWhiteSpace(c.ColumnLetter))
                .OrderByDescending(c => c.SuggestedScore)
                .ThenBy(c => GetColumnIndex(c.ColumnLetter))
                .FirstOrDefault();
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

            foreach (var row in GetAllRows())
            {
                if (row.SelectedColumn != null)
                {
                    continue;
                }

                if (!template.TryGetValue(row.Field, out var col))
                {
                    continue;
                }

                var normalized = (col ?? string.Empty).Trim().ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(normalized) || !columnOptionsByLetter.ContainsKey(normalized))
                {
                    continue;
                }

                row.SelectedColumn = normalized;
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
            foreach (var row in GetAllRows())
            {
                row.ResetSelection();
            }

            foreach (var row in GetAllRows())
            {
                if (row.SuggestedScore < FieldMappingRowViewModel.AutoSelectThreshold ||
                    string.IsNullOrWhiteSpace(row.SuggestedColumn))
                {
                    continue;
                }

                row.SelectedColumn = row.SuggestedColumn;
                row.IsLocked = row.SuggestedScore >= FieldMappingRowViewModel.LockThreshold;
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
            foreach (var row in GetAllRows())
            {
                row.ResetSelection();
            }

            foreach (var row in OptionalFieldMappings)
            {
                row.SelectedColumn = string.Empty;
            }

            foreach (var pair in profile.ConfirmedMap ?? new Dictionary<ExcelImportService.ImportField, string>())
            {
                var col = (pair.Value ?? string.Empty).Trim().ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(col) || !columnOptionsByLetter.ContainsKey(col))
                {
                    continue;
                }

                var row = GetAllRows().FirstOrDefault(r => r.Field == pair.Key);
                if (row != null)
                {
                    row.SelectedColumn = col;
                }
            }

            foreach (var row in GetAllRows())
            {
                row.IsLocked = row.SuggestedScore >= FieldMappingRowViewModel.LockThreshold &&
                               !string.IsNullOrWhiteSpace(row.SuggestedColumn) &&
                               string.Equals(row.SelectedColumn, row.SuggestedColumn, StringComparison.OrdinalIgnoreCase);
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
            foreach (var row in GetAllRows())
            {
                row.IsConflict = false;
                row.IsMissing = row.IsRequired && string.IsNullOrWhiteSpace(row.SelectedColumn);
                row.UpdateSamplePreview(columnOptionsByLetter);
            }

            var conflictGroups = GetAllRows()
                .Where(r => !string.IsNullOrWhiteSpace(r.SelectedColumn))
                .GroupBy(r => r.SelectedColumn!.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();

            foreach (var group in conflictGroups)
            {
                foreach (var row in group)
                {
                    row.IsConflict = true;
                }
            }

            OnPropertyChanged(nameof(CanImport));
        }

        private bool HasBlockingErrors()
        {
            if (GetAllRows().Any(r => r.IsConflict))
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

            var conflictGroups = GetAllRows()
                .Where(r => !string.IsNullOrWhiteSpace(r.SelectedColumn))
                .GroupBy(r => r.SelectedColumn!.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1)
                .ToList();

            foreach (var group in conflictGroups)
            {
                var col = group.Key;
                var colDisplay = columnOptionsByLetter.TryGetValue(col, out var opt) ? opt.DisplayName : col;
                var fields = group.Select(r => r.Label).ToList();

                if (fields.Count >= 2)
                {
                    errors.Add($"❌ Trùng cột: {string.Join(" và ", fields)} đều chọn cột \"{colDisplay}\"");
                }
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

        private IEnumerable<FieldMappingRowViewModel> GetAllRows()
        {
            foreach (var row in ImportantFieldMappings)
            {
                yield return row;
            }

            foreach (var row in OptionalFieldMappings)
            {
                yield return row;
            }
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
