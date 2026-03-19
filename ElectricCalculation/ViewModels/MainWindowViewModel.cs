using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Helpers;
using ElectricCalculation.Models;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public partial class MainWindowViewModel : ObservableObject
    {
        private readonly UiService _ui;
        private AppSettings _settings;
        private readonly UndoRedoManager _undoRedo = new();
        private readonly List<string> _editHistory = new();
        private bool _suppressDirty;
        private string _normalizedSearchKeyword = string.Empty;
        private int _selectedSearchFieldIndex;
        private IReadOnlyList<Customer> _viewCustomersSnapshot = Array.Empty<Customer>();
        private int _summaryCustomerCount;
        private int _summaryCompletedCount;
        private int _summaryMissingCount;
        private int _summaryWarningCount;
        private int _summaryErrorCount;
        private decimal _summaryTotalConsumption;
        private decimal _summaryTotalChargeableKwh;
        private decimal _summaryTotalAmount;
        private readonly DispatcherTimer _summaryRecalcTimer;
        private readonly HashSet<Customer> _subscribedCustomers = new();

        private const int SearchFieldNameIndex = 0;
        private const int SearchFieldGroupIndex = 1;
        private const int SearchFieldCategoryIndex = 2;
        private const int SearchFieldAddressIndex = 3;
        private const int SearchFieldPhoneIndex = 4;
        private const int SearchFieldMeterIndex = 5;

        [ObservableProperty]
        private string? currentDataFilePath;

        [ObservableProperty]
        private string periodLabel = string.Empty;

        [ObservableProperty]
        private string searchText = string.Empty;

        [ObservableProperty]
        private string selectedSearchField = string.Empty;

        [ObservableProperty]
        private Customer? selectedCustomer;

        // Invoice issuer (printed on the invoice).
        [ObservableProperty]
        private string invoiceIssuer = string.Empty;

        [ObservableProperty]
        private bool isDirty;

        [ObservableProperty]
        private string? loadedSnapshotPath;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(IsDetailMode))]
        private bool isFastEntryMode = true;

        private const string AllGroupsOption = "(Tất cả)";

        public ObservableCollection<string> GroupOptions { get; } = new();

        [ObservableProperty]
        private string selectedGroup = AllGroupsOption;

        [ObservableProperty]
        private bool filterMissing;

        [ObservableProperty]
        private bool filterWarning;

        [ObservableProperty]
        private bool filterError;

        public bool IsDetailMode => !IsFastEntryMode;

        public ObservableRangeCollection<Customer> Customers { get; } = new();

        public ICollectionView CustomersView { get; }

        public IReadOnlyList<string> SearchFields { get; } = new[]
        {
            "Tên khách",
            "Nhóm / Đơn vị",
            "Loại",
            "Địa chỉ",
            "Số ĐT",
            "Số công tơ"
        };

        private bool CanUndo() => _undoRedo.CanUndo;

        [RelayCommand(CanExecute = nameof(CanUndo))]
        private void Undo()
        {
            _undoRedo.Undo();
        }

        private bool CanRedo() => _undoRedo.CanRedo;

        [RelayCommand(CanExecute = nameof(CanRedo))]
        private void Redo()
        {
            _undoRedo.Redo();
        }

        [RelayCommand]
        private void ShowEditHistory()
        {
            if (_editHistory.Count == 0)
            {
                _ui.ShowMessage("Lịch sử chỉnh sửa", "Chưa có lịch sử.");
                return;
            }

            var lines = _editHistory
                .TakeLast(50)
                .Reverse()
                .ToList();

            _ui.ShowMessage("Lịch sử chỉnh sửa (gần nhất)", string.Join(Environment.NewLine, lines));
        }

        public MainWindowViewModel()
        {
            _suppressDirty = true;
            _ui = new UiService();
            _settings = AppSettingsService.Load();
            _normalizedSearchKeyword = string.Empty;
            _selectedSearchFieldIndex = SearchFieldNameIndex;

            _undoRedo.StateChanged += (_, _) =>
            {
                UndoCommand.NotifyCanExecuteChanged();
                RedoCommand.NotifyCanExecuteChanged();
            };

            CustomersView = CollectionViewSource.GetDefaultView(Customers);
            CustomersView.Filter = FilterCustomer;
            Customers.CollectionChanged += Customers_CollectionChanged;

            _summaryRecalcTimer = new DispatcherTimer(DispatcherPriority.Background)
            {
                Interval = TimeSpan.FromMilliseconds(150)
            };
            _summaryRecalcTimer.Tick += (_, _) =>
            {
                _summaryRecalcTimer.Stop();
                RecalculateViewSnapshotAndSummary();
            };

            if (SearchFields.Count > 0)
            {
                SelectedSearchField = SearchFields[0];
            }

            ApplySearchCache(SearchText, SelectedSearchField);
            UpdateGroupOptions();
            RecalculateViewSnapshotAndSummary();

            PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
            IsDirty = false;
            _suppressDirty = false;
        }

        private void RefreshCustomersViewAfterFilterChanged()
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
        }

        private void ApplySearchCache(string? keyword, string? searchField)
        {
            _normalizedSearchKeyword = NormalizeSearchKeyword(keyword);
            _selectedSearchFieldIndex = ResolveSearchFieldIndex(searchField);
        }

        private int ResolveSearchFieldIndex(string? searchField)
        {
            if (string.IsNullOrWhiteSpace(searchField))
            {
                return SearchFieldNameIndex;
            }

            for (var i = 0; i < SearchFields.Count; i++)
            {
                if (string.Equals(SearchFields[i], searchField, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return SearchFieldNameIndex;
        }

        private static string NormalizeSearchKeyword(string? keyword)
        {
            return string.IsNullOrWhiteSpace(keyword)
                ? string.Empty
                : keyword.Trim();
        }

        public void ImportFromExcelFile(string filePath)
        {
            try
            {
                _suppressDirty = true;

                try
                {
                    ImportFromExcel(filePath);
                }
                catch (WarningException warning)
                {
                    Debug.WriteLine(warning);
                    _ui.ShowMessage("Cảnh báo import Excel", warning.Message);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    _ui.ShowMessage("Lỗi import Excel", ex.Message);
                }
            }
            finally
            {
                _suppressDirty = false;
            }
        }

        public void LoadDataFile(string filePath, bool setCurrentDataFilePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new WarningException("Đường dẫn file dữ liệu trống.");
            }

            if (!File.Exists(filePath))
            {
                throw new WarningException("Không tìm thấy file dữ liệu.");
            }

            var (period, customers) = ProjectFileService.Load(filePath);

            try
            {
                _suppressDirty = true;
                _undoRedo.Clear();

                Customers.ReplaceRange(customers);

                if (!string.IsNullOrWhiteSpace(period))
                {
                    PeriodLabel = period;
                }

                SelectedCustomer = null;
                CurrentDataFilePath = setCurrentDataFilePath ? filePath : null;
                LoadedSnapshotPath = null;
                IsDirty = false;
                RefreshUsageAverages();
                RecalculateViewSnapshotAndSummary();
            }
            finally
            {
                _suppressDirty = false;
            }
        }

        public void LoadSnapshotFile(string filePath)
        {
            LoadDataFile(filePath, setCurrentDataFilePath: false);
            LoadedSnapshotPath = filePath;
        }

        private void RefreshUsageAverages()
        {
            var wasSuppressDirty = _suppressDirty;
            _suppressDirty = true;

            try
            {
                var averages = UsageHistoryService.BuildAverageConsumptionByMeterKey(
                    currentPeriodLabel: PeriodLabel,
                    currentCustomers: Customers,
                    periodsToAverage: 3,
                    excludeSnapshotPath: LoadedSnapshotPath);

                foreach (var customer in Customers)
                {
                    var key = UsageHistoryService.BuildMeterKey(customer);
                    customer.AverageConsumption3Periods = averages.TryGetValue(key, out var avg) ? avg : null;
                }
            }
            finally
            {
                _suppressDirty = wasSuppressDirty;
            }

            NotifySummaryChanged();
            UpdateGroupOptions();
        }

        [RelayCommand]
        private void AddRow()
        {
            var customer = new Customer();
            customer.SequenceNumber = Customers.Count == 0
                ? 1
                : Customers.Max(c => c.SequenceNumber) + 1;

            if (!string.IsNullOrWhiteSpace(InvoiceIssuer))
            {
                customer.PerformedBy = InvoiceIssuer;
            }

            ApplyDefaultsIfNeeded(customer, applyWhen: _settings.ApplyDefaultsOnNewRow, allowOverwriteExistingValues: true);
            var insertIndex = Customers.Count;

            ExecuteUndoable(new DelegateUndoableAction(
                name: "Thêm dòng",
                undo: () => Customers.Remove(customer),
                redo: () =>
                {
                    Customers.Insert(Math.Min(insertIndex, Customers.Count), customer);
                    SelectedCustomer = customer;
                }));
        }

        [RelayCommand]
        private void ClearAll()
        {
            if (Customers.Count == 0)
            {
                return;
            }

            SelectedCustomer = null;
            Customers.Clear();
        }

        // Totals for the current (filtered) view.
        private IReadOnlyList<Customer> TryGetViewCustomersSnapshot()
        {
            return _viewCustomersSnapshot.Count > 0 ? _viewCustomersSnapshot : TryGetCustomersSnapshot();
        }

        private IReadOnlyList<Customer> TryGetCustomersSnapshot()
        {
            try
            {
                return Customers.ToList();
            }
            catch
            {
                return Array.Empty<Customer>();
            }
        }

        private void EnsureViewSnapshotUpToDate()
        {
            if (_summaryRecalcTimer.IsEnabled)
            {
                RecalculateViewSnapshotAndSummary();
            }
        }

        private List<Customer> GetCurrentViewCustomers()
        {
            EnsureViewSnapshotUpToDate();
            return TryGetViewCustomersSnapshot().ToList();
        }

        public int CustomerCount => _summaryCustomerCount;

        public decimal TotalConsumption => _summaryTotalConsumption;

        public decimal TotalChargeableKwh => _summaryTotalChargeableKwh;

        public decimal TotalAmount => _summaryTotalAmount;

        public int CompletedCount => _summaryCompletedCount;

        public int MissingCount => _summaryMissingCount;

        public int WarningCount => _summaryWarningCount;

        public int ErrorCount => _summaryErrorCount;

        public double CompletionRatio
        {
            get
            {
                if (_summaryCustomerCount <= 0)
                {
                    return 0;
                }

                var ratio = (double)_summaryCompletedCount / _summaryCustomerCount;

                if (double.IsNaN(ratio) || double.IsInfinity(ratio))
                {
                    return 0;
                }

                return Math.Max(0, Math.Min(1, ratio));
            }
        }

        private void NotifySummaryChanged()
        {
            ScheduleSummaryRecalc();
        }

        private void ScheduleSummaryRecalc()
        {
            _summaryRecalcTimer.Stop();
            _summaryRecalcTimer.Start();
        }

        private void RecalculateViewSnapshotAndSummary()
        {
            _summaryRecalcTimer.Stop();

            try
            {
                var completedCount = 0;
                var missingCount = 0;
                var warningCount = 0;
                var errorCount = 0;
                decimal totalConsumption = 0;
                decimal totalChargeable = 0;
                decimal totalAmount = 0;

                var viewItems = new List<Customer>();
                foreach (var item in CustomersView)
                {
                    if (item is not Customer customer)
                    {
                        continue;
                    }

                    viewItems.Add(customer);

                    totalConsumption += customer.Consumption;
                    totalChargeable += customer.ChargeableKwh;
                    totalAmount += customer.Amount;

                    if (customer.CurrentIndex != null)
                    {
                        completedCount++;
                    }

                    if (customer.IsMissingReading)
                    {
                        missingCount++;
                    }

                    if (customer.HasUsageWarning)
                    {
                        warningCount++;
                    }

                    if (customer.HasReadingError)
                    {
                        errorCount++;
                    }
                }

                _viewCustomersSnapshot = viewItems;

                _summaryCustomerCount = viewItems.Count;
                _summaryCompletedCount = completedCount;
                _summaryMissingCount = missingCount;
                _summaryWarningCount = warningCount;
                _summaryErrorCount = errorCount;
                _summaryTotalConsumption = totalConsumption;
                _summaryTotalChargeableKwh = totalChargeable;
                _summaryTotalAmount = totalAmount;
            }
            catch
            {
                var customers = TryGetCustomersSnapshot();
                _viewCustomersSnapshot = customers;

                var completedCount = 0;
                var missingCount = 0;
                var warningCount = 0;
                var errorCount = 0;
                decimal totalConsumption = 0;
                decimal totalChargeable = 0;
                decimal totalAmount = 0;

                foreach (var customer in customers)
                {
                    totalConsumption += customer.Consumption;
                    totalChargeable += customer.ChargeableKwh;
                    totalAmount += customer.Amount;

                    if (customer.CurrentIndex != null)
                    {
                        completedCount++;
                    }

                    if (customer.IsMissingReading)
                    {
                        missingCount++;
                    }

                    if (customer.HasUsageWarning)
                    {
                        warningCount++;
                    }

                    if (customer.HasReadingError)
                    {
                        errorCount++;
                    }
                }

                _summaryCustomerCount = customers.Count;
                _summaryCompletedCount = completedCount;
                _summaryMissingCount = missingCount;
                _summaryWarningCount = warningCount;
                _summaryErrorCount = errorCount;
                _summaryTotalConsumption = totalConsumption;
                _summaryTotalChargeableKwh = totalChargeable;
                _summaryTotalAmount = totalAmount;
            }

            OnPropertyChanged(nameof(CustomerCount));
            OnPropertyChanged(nameof(CompletedCount));
            OnPropertyChanged(nameof(MissingCount));
            OnPropertyChanged(nameof(WarningCount));
            OnPropertyChanged(nameof(ErrorCount));
            OnPropertyChanged(nameof(CompletionRatio));
            OnPropertyChanged(nameof(TotalConsumption));
            OnPropertyChanged(nameof(TotalChargeableKwh));
            OnPropertyChanged(nameof(TotalAmount));

        }




        [RelayCommand]
        private void ImportFromExcel(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new WarningException("Đường dẫn file Excel cần import đang trống.");
            }

            if (!File.Exists(filePath))
            {
                throw new WarningException("Không tìm thấy file Excel để import.");
            }

            var wizard = _ui.ShowImportWizardDialog(filePath);
            if (wizard == null)
            {
                return;
            }

            var imported = wizard.ImportedCustomers?.ToList() ?? new List<Customer>();

            if (imported.Count == 0)
            {
                return;
            }

            SelectedCustomer = null;
            var prepared = new List<Customer>(imported.Count);
            foreach (var customer in imported)
            {
                if (!string.IsNullOrWhiteSpace(InvoiceIssuer) && string.IsNullOrWhiteSpace(customer.PerformedBy))
                {
                    customer.PerformedBy = InvoiceIssuer;
                }

                // Import is a full replace (Customers.ReplaceRange). Defaults should fill missing values,
                // but should not overwrite values explicitly provided by Excel.
                ApplyDefaultsIfNeeded(customer, applyWhen: _settings.ApplyDefaultsOnImport, allowOverwriteExistingValues: false);
                prepared.Add(customer);
            }

            Customers.ReplaceRange(prepared);
            _undoRedo.Clear();
            LoadedSnapshotPath = null;

            RefreshUsageAverages();
            AutoSaveImportedSnapshot();
            RecalculateViewSnapshotAndSummary();
        }

        private void AutoSaveImportedSnapshot()
        {
            try
            {
                var snapshotPath = SaveGameService.SaveSnapshot(PeriodLabel, Customers, snapshotName: "Tháng mới");
                LoadedSnapshotPath = snapshotPath;
                IsDirty = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                LoadedSnapshotPath = null;
                IsDirty = true;
                _ui.ShowMessage("Cảnh báo import Excel", $"Đã import nhưng tự động lưu thất bại:\n{ex.Message}");
            }
        }

        private sealed record CurrentIndexImportResult(
            int SourceRows,
            int UpdatedRows,
            int MissingCurrentIndexRows,
            int UnmatchedRows,
            int AmbiguousRows,
            int MissingKeyRows);

        private CurrentIndexImportResult ApplyCurrentIndexImport(
            IReadOnlyList<Customer> importedRows,
            IReadOnlyDictionary<ExcelImportService.ImportField, string> map,
            bool preferBestKeyPerRow = false)
        {
            var matchByMeter = map.ContainsKey(ExcelImportService.ImportField.MeterNumber);
            var matchBySequence = map.ContainsKey(ExcelImportService.ImportField.SequenceNumber);
            var matchByName = map.ContainsKey(ExcelImportService.ImportField.Name);

            if (!matchByMeter && !matchBySequence && !matchByName)
            {
                throw new WarningException("Cần chọn ít nhất 1 khóa ghép: Số công tơ / Số thứ tự / Tên khách.");
            }

            var byMeter = new Dictionary<string, Customer>(StringComparer.OrdinalIgnoreCase);
            var bySequence = new Dictionary<int, Customer>();
            var byName = new Dictionary<string, Customer>(StringComparer.OrdinalIgnoreCase);

            var ambiguousMeters = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var ambiguousSequences = new HashSet<int>();
            var ambiguousNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var customer in Customers)
            {
                if (matchByMeter)
                {
                    RegisterUniqueLookup(byMeter, ambiguousMeters, NormalizeLookupKey(customer.MeterNumber), customer);
                }

                if (matchBySequence && customer.SequenceNumber > 0)
                {
                    RegisterUniqueLookup(bySequence, ambiguousSequences, customer.SequenceNumber, customer);
                }

                if (matchByName)
                {
                    RegisterUniqueLookup(byName, ambiguousNames, NormalizeLookupKey(customer.Name), customer);
                }
            }

            var updatedTargets = new HashSet<Customer>();
            var updatedRows = 0;
            var missingCurrentIndexRows = 0;
            var unmatchedRows = 0;
            var ambiguousRows = 0;
            var missingKeyRows = 0;

            var wasSuppressDirty = _suppressDirty;
            _suppressDirty = true;
            try
            {
                foreach (var imported in importedRows)
                {
                    if (!imported.CurrentIndex.HasValue)
                    {
                        missingCurrentIndexRows++;
                        continue;
                    }

                    if (preferBestKeyPerRow)
                    {
                        // One-column import: avoid matching by name when meter/sequence is present.
                        // This prevents wrong updates when the same household exists in 1 pha/3 pha rows.
                        Customer? bestTarget = null;

                        if (matchByMeter)
                        {
                            var meterKey = NormalizeLookupKey(imported.MeterNumber);
                            if (!string.IsNullOrWhiteSpace(meterKey))
                            {
                                if (ambiguousMeters.Contains(meterKey))
                                {
                                    ambiguousRows++;
                                    continue;
                                }

                                if (!byMeter.TryGetValue(meterKey, out bestTarget))
                                {
                                    unmatchedRows++;
                                    continue;
                                }
                            }
                        }

                        if (bestTarget == null && matchBySequence && imported.SequenceNumber > 0)
                        {
                            if (ambiguousSequences.Contains(imported.SequenceNumber))
                            {
                                ambiguousRows++;
                                continue;
                            }

                            if (!bySequence.TryGetValue(imported.SequenceNumber, out bestTarget))
                            {
                                unmatchedRows++;
                                continue;
                            }
                        }

                        if (bestTarget == null && matchByName)
                        {
                            var nameKey = NormalizeLookupKey(imported.Name);
                            if (!string.IsNullOrWhiteSpace(nameKey))
                            {
                                if (ambiguousNames.Contains(nameKey))
                                {
                                    ambiguousRows++;
                                    continue;
                                }

                                if (!byName.TryGetValue(nameKey, out bestTarget))
                                {
                                    unmatchedRows++;
                                    continue;
                                }
                            }
                        }

                        if (bestTarget == null)
                        {
                            missingKeyRows++;
                            continue;
                        }

                        if (bestTarget.CurrentIndex != imported.CurrentIndex)
                        {
                            bestTarget.CurrentIndex = imported.CurrentIndex;
                            updatedTargets.Add(bestTarget);
                            updatedRows++;
                        }

                        continue;
                    }

                    var matches = new List<Customer>(3);
                    var hadAnyProvidedKey = false;
                    var hasAmbiguousProvidedKey = false;

                    if (matchByMeter)
                    {
                        var meterKey = NormalizeLookupKey(imported.MeterNumber);
                        if (!string.IsNullOrWhiteSpace(meterKey))
                        {
                            hadAnyProvidedKey = true;
                            if (ambiguousMeters.Contains(meterKey))
                            {
                                hasAmbiguousProvidedKey = true;
                            }
                            else if (byMeter.TryGetValue(meterKey, out var meterCustomer))
                            {
                                matches.Add(meterCustomer);
                            }
                        }
                    }

                    if (matchBySequence && imported.SequenceNumber > 0)
                    {
                        hadAnyProvidedKey = true;
                        if (ambiguousSequences.Contains(imported.SequenceNumber))
                        {
                            hasAmbiguousProvidedKey = true;
                        }
                        else if (bySequence.TryGetValue(imported.SequenceNumber, out var sequenceCustomer))
                        {
                            matches.Add(sequenceCustomer);
                        }
                    }

                    if (matchByName)
                    {
                        var nameKey = NormalizeLookupKey(imported.Name);
                        if (!string.IsNullOrWhiteSpace(nameKey))
                        {
                            hadAnyProvidedKey = true;
                            if (ambiguousNames.Contains(nameKey))
                            {
                                hasAmbiguousProvidedKey = true;
                            }
                            else if (byName.TryGetValue(nameKey, out var nameCustomer))
                            {
                                matches.Add(nameCustomer);
                            }
                        }
                    }

                    if (!hadAnyProvidedKey)
                    {
                        missingKeyRows++;
                        continue;
                    }

                    if (matches.Count == 0)
                    {
                        if (hasAmbiguousProvidedKey)
                        {
                            ambiguousRows++;
                        }
                        else
                        {
                            unmatchedRows++;
                        }

                        continue;
                    }

                    var target = matches[0];
                    if (matches.Any(c => !ReferenceEquals(c, target)))
                    {
                        ambiguousRows++;
                        continue;
                    }

                    if (target.CurrentIndex != imported.CurrentIndex)
                    {
                        target.CurrentIndex = imported.CurrentIndex;
                        updatedTargets.Add(target);
                        updatedRows++;
                    }
                }
            }
            finally
            {
                _suppressDirty = wasSuppressDirty;
            }

            if (updatedTargets.Count > 0)
            {
                IsDirty = true;
                CustomersView.Refresh();
                NotifySummaryChanged();
            }

            return new CurrentIndexImportResult(
                SourceRows: importedRows.Count,
                UpdatedRows: updatedRows,
                MissingCurrentIndexRows: missingCurrentIndexRows,
                UnmatchedRows: unmatchedRows,
                AmbiguousRows: ambiguousRows,
                MissingKeyRows: missingKeyRows);
        }

        private static string BuildCurrentIndexImportSummary(CurrentIndexImportResult result)
        {
            return
                $"Tổng dòng đọc: {result.SourceRows}\n" +
                $"Đã cập nhật: {result.UpdatedRows}\n" +
                $"Thiếu chỉ số mới: {result.MissingCurrentIndexRows}\n" +
                $"Không tìm thấy khách: {result.UnmatchedRows}\n" +
                $"Khóa ghép bị trùng/không rõ: {result.AmbiguousRows}\n" +
                $"Thiếu khóa ghép trên dòng import: {result.MissingKeyRows}";
        }

        private static string NormalizeLookupKey(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static void RegisterUniqueLookup<TKey>(
            IDictionary<TKey, Customer> map,
            ISet<TKey> ambiguousKeys,
            TKey key,
            Customer customer)
            where TKey : notnull
        {
            if (EqualityComparer<TKey>.Default.Equals(key, default!))
            {
                return;
            }

            if (key is string textKey && string.IsNullOrWhiteSpace(textKey))
            {
                return;
            }

            if (ambiguousKeys.Contains(key))
            {
                return;
            }

            if (map.TryGetValue(key, out var existing) && !ReferenceEquals(existing, customer))
            {
                map.Remove(key);
                ambiguousKeys.Add(key);
                return;
            }

            map[key] = customer;
        }

        [RelayCommand]
        private void OpenSettingsWithDialog()
        {
            try
            {
                var updated = _ui.ShowSettingsDialog(_settings ?? new AppSettings());
                if (updated == null)
                {
                    return;
                }

                AppSettingsService.Save(updated);
                _settings = updated;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Loi", ex.Message);
            }
        }

        private void ApplyDefaultsIfNeeded(Customer customer, bool applyWhen, bool allowOverwriteExistingValues)
        {
            var s = _settings ?? new AppSettings();
            ApplyDefaultsIfNeeded(customer, applyWhen, s, allowOverwriteExistingValues);
        }

        private void ApplyDefaultsIfNeeded(
            Customer customer,
            bool applyWhen,
            AppSettings settings,
            bool allowOverwriteExistingValues)
        {
            if (!applyWhen || customer == null)
            {
                return;
            }

            var s = settings ?? new AppSettings();
            var overwrite = allowOverwriteExistingValues && s.OverrideExistingValues;

            if (overwrite || customer.UnitPrice <= 0)
            {
                if (s.DefaultUnitPrice > 0)
                {
                    customer.UnitPrice = s.DefaultUnitPrice;
                }
            }

            if (overwrite || customer.Multiplier <= 0)
            {
                customer.Multiplier = s.DefaultMultiplier > 0 ? s.DefaultMultiplier : 1m;
            }

            if (overwrite || customer.SubsidizedKwh <= 0)
            {
                if (s.DefaultSubsidizedKwh > 0)
                {
                    customer.SubsidizedKwh = s.DefaultSubsidizedKwh;
                }
            }

            if (overwrite || string.IsNullOrWhiteSpace(customer.PerformedBy))
            {
                if (!string.IsNullOrWhiteSpace(s.DefaultPerformedBy))
                {
                    customer.PerformedBy = s.DefaultPerformedBy;
                }
            }
        }

        [RelayCommand]
        private void TrackCellEdit(CellEditChange? change)
        {
            if (change == null)
            {
                return;
            }

            if (!TryGetWritableCustomerProperty(change.PropertyName, out var property))
            {
                return;
            }

            var before = change.OldValue;
            var after = change.NewValue;
            if (Equals(before, after))
            {
                return;
            }

            var action = new DelegateUndoableAction(
                name: $"Sửa {change.PropertyName}",
                undo: () => property.SetValue(change.Customer, before),
                redo: () => property.SetValue(change.Customer, after));

            PushUndoable(action);
        }

        [RelayCommand]
        private void PasteFromClipboard(ClipboardPasteRequest? request)
        {
            if (request == null || request.TargetRows.Count == 0 || request.PropertyNames.Count == 0)
            {
                return;
            }

            var parsed = ParseClipboardMatrix(request.ClipboardText);
            if (parsed.Count == 0)
            {
                return;
            }

            var actions = new List<IUndoableAction>();

            for (var r = 0; r < request.TargetRows.Count && r < parsed.Count; r++)
            {
                var row = request.TargetRows[r];
                var values = parsed[r];

                for (var c = 0; c < request.PropertyNames.Count && c < values.Length; c++)
                {
                    var propertyName = request.PropertyNames[c];
                    if (!TryGetWritableCustomerProperty(propertyName, out var property))
                    {
                        continue;
                    }

                    var prop = property;

                    if (!TryConvertText(values[c], property.PropertyType, out var converted))
                    {
                        continue;
                    }

                    var before = prop.GetValue(row);
                    if (Equals(before, converted))
                    {
                        continue;
                    }

                    var capturedBefore = before;
                    var capturedAfter = converted;
                    actions.Add(new DelegateUndoableAction(
                        name: "Dán dữ liệu",
                        undo: () => prop.SetValue(row, capturedBefore),
                        redo: () => prop.SetValue(row, capturedAfter)));
                }
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Dán dữ liệu", actions));
        }

        [RelayCommand]
        private void FillDown(FillDownRequest? request)
        {
            if (request == null || request.TargetRows.Count < 2)
            {
                return;
            }

            if (!TryGetWritableCustomerProperty(request.PropertyName, out var property))
            {
                return;
            }

            var source = request.TargetRows[0];
            var sourceValue = property.GetValue(source);

            var actions = new List<IUndoableAction>();
            for (var i = 1; i < request.TargetRows.Count; i++)
            {
                var row = request.TargetRows[i];
                var before = property.GetValue(row);
                if (Equals(before, sourceValue))
                {
                    continue;
                }

                var capturedBefore = before;
                actions.Add(new DelegateUndoableAction(
                    name: "Fill-down",
                    undo: () => property.SetValue(row, capturedBefore),
                    redo: () => property.SetValue(row, sourceValue)));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Fill-down", actions));
        }

        [RelayCommand]
        private void SetColumnValue(DataGridColumn? column)
        {
            if (column == null)
            {
                return;
            }

            var propertyName = GetPropertyName(column);
            if (string.IsNullOrWhiteSpace(propertyName))
            {
                _ui.ShowMessage("Đặt toàn bộ cột", "Không xác định được cột cần chỉnh (cột không có binding).");
                return;
            }

            if (!TryGetWritableCustomerProperty(propertyName, out var property))
            {
                _ui.ShowMessage("Đặt toàn bộ cột", $"Cột '{column.Header}' không hỗ trợ chỉnh sửa hàng loạt.");
                return;
            }

            var columnTitle = (column.Header?.ToString() ?? propertyName).Trim();
            var initialValue = Customers.Count > 0
                ? Convert.ToString(property.GetValue(Customers[0]), CultureInfo.CurrentCulture)
                : null;

            var text = _ui.ShowSetColumnValueDialog(columnTitle, initialValue);
            if (text == null)
            {
                return;
            }

            if (!TryConvertText(text, property.PropertyType, out var converted))
            {
                _ui.ShowMessage("Đặt toàn bộ cột", $"Giá trị '{text}' không hợp lệ cho cột '{columnTitle}'.");
                return;
            }

            if (Customers.Count == 0)
            {
                return;
            }

            var actions = new List<IUndoableAction>();
            foreach (var customer in Customers)
            {
                var before = property.GetValue(customer);
                if (Equals(before, converted))
                {
                    continue;
                }

                var capturedCustomer = customer;
                var capturedBefore = before;
                var capturedAfter = converted;
                actions.Add(new DelegateUndoableAction(
                    name: "Đặt toàn bộ cột",
                    undo: () => property.SetValue(capturedCustomer, capturedBefore),
                    redo: () => property.SetValue(capturedCustomer, capturedAfter)));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction($"Đặt cột: {columnTitle}", actions));
        }

        [RelayCommand]
        private void DeleteSelectedRows(IList? selectedItems)
        {
            var targets = selectedItems?.OfType<Customer>()
                .Distinct()
                .ToList();

            if (targets == null || targets.Count == 0)
            {
                return;
            }

            var entries = targets
                .Select(c => (Index: Customers.IndexOf(c), Customer: c))
                .Where(x => x.Index >= 0)
                .OrderBy(x => x.Index)
                .ToList();

            if (entries.Count == 0)
            {
                return;
            }

            var action = new DelegateUndoableAction(
                name: "Xóa dòng",
                undo: () =>
                {
                    foreach (var (index, customer) in entries)
                    {
                        Customers.Insert(Math.Min(index, Customers.Count), customer);
                    }
                },
                redo: () =>
                {
                    for (var i = entries.Count - 1; i >= 0; i--)
                    {
                        Customers.Remove(entries[i].Customer);
                    }
                });

            ExecuteUndoable(action);
        }

        [RelayCommand]
        private void DuplicateRow(Customer? source)
        {
            if (source == null)
            {
                return;
            }

            var insertIndex = Customers.IndexOf(source);
            var clone = CloneCustomer(source);
            clone.SequenceNumber = Customers.Count == 0 ? 1 : Customers.Max(c => c.SequenceNumber) + 1;

            var action = new DelegateUndoableAction(
                name: "Nhân đôi dòng",
                undo: () => Customers.Remove(clone),
                redo: () =>
                {
                    var index = insertIndex >= 0 ? insertIndex + 1 : Customers.Count;
                    Customers.Insert(Math.Min(index, Customers.Count), clone);
                    SelectedCustomer = clone;
                });

            ExecuteUndoable(action);
        }

        private void ExecuteUndoable(IUndoableAction action)
        {
            _undoRedo.Execute(action);
            RecordHistory(action.Name);
        }

        private void PushUndoable(IUndoableAction action)
        {
            _undoRedo.PushDone(action);
            RecordHistory(action.Name);
        }

        private void RecordHistory(string actionName)
        {
            if (string.IsNullOrWhiteSpace(actionName))
            {
                return;
            }

            _editHistory.Add($"{DateTime.Now:HH:mm:ss} - {actionName}");

            const int max = 300;
            if (_editHistory.Count > max)
            {
                _editHistory.RemoveRange(0, _editHistory.Count - max);
            }
        }

        private static bool TryGetWritableCustomerProperty(string propertyName, out PropertyInfo property)
        {
            property = null!;

            if (string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            var p = typeof(Customer).GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
            if (p == null || !p.CanWrite)
            {
                return false;
            }

            property = p;
            return true;
        }

        private static bool TryConvertText(string text, Type targetType, out object? value)
        {
            value = null;
            var underlyingType = Nullable.GetUnderlyingType(targetType);
            var t = underlyingType ?? targetType;

            if (t == typeof(string))
            {
                value = text ?? string.Empty;
                return true;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
                if (underlyingType != null)
                {
                    value = null;
                    return true;
                }

                value = t == typeof(decimal) ? 0m : t == typeof(int) ? 0 : Activator.CreateInstance(t);
                return true;
            }

            if (t == typeof(int))
            {
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out var i) ||
                    int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out i))
                {
                    value = i;
                    return true;
                }

                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var d) ||
                    decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                {
                    value = (int)d;
                    return true;
                }

                return false;
            }

            if (t == typeof(decimal))
            {
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var d) ||
                    decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                {
                    value = d;
                    return true;
                }

                return false;
            }

            try
            {
                value = Convert.ChangeType(text, t, CultureInfo.CurrentCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string? GetPropertyName(DataGridColumn? column)
        {
            if (column is not DataGridBoundColumn bound)
            {
                return null;
            }

            if (bound.Binding is not Binding binding)
            {
                return null;
            }

            return binding.Path?.Path;
        }

        private static List<string[]> ParseClipboardMatrix(string text)
        {
            var normalized = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = normalized.Split('\n');
            var result = new List<string[]>();

            foreach (var line in lines)
            {
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                result.Add(line.Split('\t'));
            }

            return result;
        }


        [RelayCommand]
        private void ImportFromExcelWithDialog()
        {
            var filePath = _ui.ShowOpenExcelFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                ImportFromExcel(filePath);
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Cảnh báo import Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi import Excel", ex.Message);
            }
        }

        [RelayCommand]
        private void ImportCurrentIndexFromExcelWithDialog()
        {
            if (Customers.Count == 0)
            {
                _ui.ShowMessage("Import chỉ số mới", "Không có dữ liệu hiện tại để cập nhật.");
                return;
            }

            var filePath = _ui.ShowOpenExcelFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                var wizard = _ui.ShowImportWizardDialog(
                    filePath,
                    ImportWizardViewModel.ImportWizardMode.CurrentIndexOneColumn);

                if (wizard == null)
                {
                    return;
                }

                var importedRows = wizard.ImportedCustomers?.ToList() ?? new List<Customer>();
                if (importedRows.Count == 0)
                {
                    _ui.ShowMessage("Import chỉ số mới", "Không có dòng hợp lệ để cập nhật.");
                    return;
                }

                var map = wizard.BuildConfirmedMap();
                var result = ApplyCurrentIndexImport(importedRows, map, preferBestKeyPerRow: wizard.IsCurrentIndexOneColumnMode);
                var autoSaveMessage = AutoSaveAfterCurrentIndexImport();
                var message = BuildCurrentIndexImportSummary(result);
                if (!string.IsNullOrWhiteSpace(autoSaveMessage))
                {
                    message += "\n\n" + autoSaveMessage;
                }

                _ui.ShowMessage("Import chỉ số mới", message);
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Cảnh báo import chỉ số mới", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi import chỉ số mới", ex.Message);
            }
        }

        [RelayCommand]
        private async Task ExportLegacySummaryDataSheetWithDialog()
        {
            if (Customers.Count == 0)
            {
                _ui.ShowMessage("Export ngược ra Excel", "Không có dữ liệu hiện tại để export.");
                return;
            }

            string defaultFileName;
            if (TryParseMonthYear(PeriodLabel, out var month, out var year))
            {
                defaultFileName = $"Bang_tong_hop_thu_thang_{month:00}_{year}.xlsx";
            }
            else
            {
                defaultFileName = "Bang_tong_hop_thu.xlsx";
            }

            var outputPath = _ui.ShowSaveExcelFileDialog(defaultFileName, title: "Export ngược ra Excel (sheet Data)");
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            try
            {
                var templatePath = _ui.GetLegacySummaryTemplatePath();
                if (string.Equals(Path.GetFullPath(outputPath), Path.GetFullPath(templatePath), StringComparison.OrdinalIgnoreCase))
                {
                    _ui.ShowMessage("Export ngược ra Excel", "Không thể ghi đè trực tiếp lên file template. Hãy chọn tên file khác.");
                    return;
                }

                var customersSnapshot = Customers.ToList();
                using var busy = _ui.ShowBusyScope("Export ngược ra Excel", "Đang cập nhật sheet Data theo dữ liệu đã sửa...");

                var result = await Task.Run(() =>
                    ExcelRoundTripExportService.ExportLegacySummaryDataSheet(
                        templatePath,
                        outputPath,
                        customersSnapshot,
                        PeriodLabel));

                var message = $"Đã cập nhật {result.UpdatedRows} dòng và {result.UpdatedCells} ô trong sheet Data.\n\nFile: {outputPath}";
                if (result.MissingCustomers > 0)
                {
                    var sample = string.Join(", ", result.MissingSequenceNumbers.Take(15));
                    message += $"\n\nKhông tìm thấy {result.MissingCustomers} khách theo STT trong sheet Data (ví dụ: {sample}{(result.MissingCustomers > 15 ? ", ..." : string.Empty)}).";
                }

                _ui.ShowMessage("Export ngược ra Excel", message);

                var open = _ui.Confirm(
                    "Export ngược ra Excel",
                    "Bạn có muốn mở file Excel vừa xuất để kiểm tra không?");

                if (open)
                {
                    _ui.OpenWithDefaultApp(outputPath);
                }
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Export ngược ra Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi export ngược ra Excel", ex.Message);
            }
        }

        private static bool TryParseMonthYear(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            var text = (periodLabel ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            var match = Regex.Match(text, @"(\d{1,2}).*?(\d{4})", RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                return false;
            }

            if (!int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out month))
            {
                return false;
            }

            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out year))
            {
                return false;
            }

            return month is >= 1 and <= 12 && year is >= 1900 and <= 2200;
        }

        private string AutoSaveAfterCurrentIndexImport()
        {
            try
            {
                if (Customers.Count == 0)
                {
                    return string.Empty;
                }

                if (!string.IsNullOrWhiteSpace(LoadedSnapshotPath) && File.Exists(LoadedSnapshotPath))
                {
                    ProjectFileService.Save(LoadedSnapshotPath, PeriodLabel, Customers);
                    SaveGameService.SyncSnapshotFileToSharedStore(LoadedSnapshotPath, PeriodLabel);
                    IsDirty = false;
                    return $"Đã tự động lưu (ghi đè) bộ dữ liệu:\n{LoadedSnapshotPath}";
                }

                var snapshotPath = SaveGameService.SaveSnapshot(PeriodLabel, Customers, snapshotName: "Tháng mới");
                LoadedSnapshotPath = snapshotPath;
                IsDirty = false;
                return $"Đã tự động lưu bộ dữ liệu tháng mới:\n{snapshotPath}";
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                IsDirty = true;
                return $"Đã import nhưng tự động lưu thất bại:\n{ex.Message}";
            }
        }

        [RelayCommand]
        private void OpenDataFileWithDialog()
        {
            var filePath = _ui.ShowOpenDataFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                LoadDataFile(filePath, setCurrentDataFilePath: true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi mở dữ liệu", ex.Message);
            }
        }

        [RelayCommand]
        private void SaveDataFileWithDialog()
        {
            var outputPath = CurrentDataFilePath;
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                outputPath = _ui.ShowSaveDataFileDialog("Du_lieu_tien_dien.json", title: "Lưu dữ liệu");
            }

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            try
            {
                ProjectFileService.Save(outputPath, PeriodLabel, Customers);
                CurrentDataFilePath = outputPath;
                SaveGameService.SaveSnapshot(PeriodLabel, Customers);
                IsDirty = false;
                _ui.ShowMessage("Lưu dữ liệu", $"Đã lưu dữ liệu tại:\n{outputPath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi lưu dữ liệu", ex.Message);
            }
        }

        [RelayCommand]
        private void SaveSnapshot()
        {
            try
            {
                if (Customers.Count == 0)
                {
                    throw new WarningException("Không có dữ liệu để lưu snapshot.");
                }

                if (!string.IsNullOrWhiteSpace(LoadedSnapshotPath))
                {
                    var (result, action, snapshotName) = _ui.ShowSaveSnapshotPrompt(
                        PeriodLabel,
                        Customers.Count,
                        defaultSnapshotName: "Chỉnh sửa",
                        canOverwrite: true);

                    if (result == null || action == SaveSnapshotPromptAction.DontSave)
                    {
                        return;
                    }

                    if (action == SaveSnapshotPromptAction.Overwrite)
                    {
                        ProjectFileService.Save(LoadedSnapshotPath, PeriodLabel, Customers);
                        SaveGameService.SyncSnapshotFileToSharedStore(LoadedSnapshotPath, PeriodLabel);
                        IsDirty = false;
                        _ui.ShowMessage("Bộ dữ liệu", $"Đã ghi đè bộ dữ liệu:\n{LoadedSnapshotPath}");
                        return;
                    }

                    var newSnapshotPath = SaveGameService.SaveSnapshot(PeriodLabel, Customers, snapshotName);
                    IsDirty = false;
                    _ui.ShowMessage("Bộ dữ liệu", $"Đã tạo bộ dữ liệu tại:\n{newSnapshotPath}");
                    return;
                }

                var snapshotPath = SaveGameService.SaveSnapshot(PeriodLabel, Customers);
                IsDirty = false;
                _ui.ShowMessage("Bộ dữ liệu", $"Đã tạo bộ dữ liệu tại:\n{snapshotPath}");
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Bộ dữ liệu", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi snapshot", ex.Message);
            }
        }

        [RelayCommand]
        private void OpenSnapshotWithDialog()
        {
            var filePath = _ui.ShowOpenSnapshotFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            try
            {
                LoadSnapshotFile(filePath);

                // Tránh ghi đè snapshot khi bấm "Lưu dữ liệu..."
                _ui.ShowMessage("Bộ dữ liệu", $"Đã mở bộ dữ liệu:\n{filePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi mở snapshot", ex.Message);
            }
        }

        [RelayCommand]
        private void OpenSnapshotFolder()
        {
            try
            {
                var folder = _ui.GetSnapshotFolderPath();
                _ui.OpenWithDefaultApp(folder);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi", ex.Message);
            }
        }

        [RelayCommand]
        private void NewPeriodWithDialog()
        {
            var referenceOptions = new List<NewPeriodViewModel.ReferenceDatasetOption>();

            if (Customers.Count > 0)
            {
                referenceOptions.Add(new NewPeriodViewModel.ReferenceDatasetOption(
                    PeriodLabel,
                    $"Đang mở: {PeriodLabel}",
                    SnapshotPath: null,
                    IsCurrentDataset: true));
            }

            var snapshots = SaveGameService.ListSnapshots(maxCount: 200)
                .GroupBy(s => s.PeriodLabel)
                .Select(g => g.OrderByDescending(x => x.SavedAt).First())
                .OrderByDescending(s => s.SavedAt)
                .ToList();

            foreach (var snapshot in snapshots)
            {
                var namePart = string.IsNullOrWhiteSpace(snapshot.SnapshotName) ? string.Empty : $" - {snapshot.SnapshotName}";
                var display = $"{snapshot.PeriodLabel}{namePart} ({snapshot.SavedAt:dd/MM/yyyy HH:mm})";

                referenceOptions.Add(new NewPeriodViewModel.ReferenceDatasetOption(
                    snapshot.PeriodLabel,
                    display,
                    snapshot.Path,
                    IsCurrentDataset: false));
            }

            if (referenceOptions.Count == 0)
            {
                _ui.ShowMessage("Làm tháng mới", "Không có bộ dữ liệu để chọn làm tháng cũ. Hãy mở hoặc lưu snapshot trước.");
                return;
            }

            var dialogVm = new NewPeriodViewModel(referenceOptions);

            if (TryGetNextPeriod(referenceOptions[0].PeriodLabel, out var nextMonth, out var nextYear))
            {
                dialogVm.Month = nextMonth;
                dialogVm.Year = nextYear;
            }

            var vm = _ui.ShowNewPeriodDialog(dialogVm);
            if (vm == null)
            {
                return;
            }

            if (vm.Month is < 1 or > 12 || vm.Year < 2000)
            {
                _ui.ShowMessage("Làm tháng mới", "Tháng/Năm không hợp lệ.");
                return;
            }

            if (vm.SelectedReferenceDataset == null)
            {
                _ui.ShowMessage("Làm tháng mới", "Hãy chọn bộ dữ liệu làm tháng cũ.");
                return;
            }

            IReadOnlyList<Customer> sourceCustomers;

            if (vm.SelectedReferenceDataset.IsCurrentDataset)
            {
                sourceCustomers = Customers.ToList();
            }
            else
            {
                var snapshotPath = vm.SelectedReferenceDataset.SnapshotPath;
                if (string.IsNullOrWhiteSpace(snapshotPath) || !File.Exists(snapshotPath))
                {
                    _ui.ShowMessage("Làm tháng mới", "Không tìm thấy file bộ dữ liệu tháng cũ.");
                    return;
                }

                var loaded = ProjectFileService.Load(snapshotPath);
                sourceCustomers = loaded.Customers;
            }

            var newCustomers = sourceCustomers
                .OrderBy(c => c.SequenceNumber)
                .Select(CloneCustomer)
                .ToList();

            foreach (var c in newCustomers)
            {
                if (vm.MoveCurrentToPrevious)
                {
                    c.PreviousIndex = c.CurrentIndex ?? c.PreviousIndex;
                }

                if (vm.ResetCurrentToZero)
                {
                    c.CurrentIndex = null;
                }
            }

            Customers.ReplaceRange(newCustomers);

            SelectedCustomer = null;
            PeriodLabel = vm.PeriodLabel;
            CurrentDataFilePath = null;
            LoadedSnapshotPath = null;
            RefreshUsageAverages();
            RecalculateViewSnapshotAndSummary();
        }

        private static bool TryGetNextPeriod(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return false;
            }

            var parts = periodLabel.Split('/');
            if (parts.Length < 2)
            {
                return false;
            }

            var monthText = new string(parts[0].Where(char.IsDigit).ToArray());
            var yearText = new string(parts[1].Where(char.IsDigit).ToArray());

            if (!int.TryParse(monthText, out var m) || !int.TryParse(yearText, out var y))
            {
                return false;
            }

            if (m is < 1 or > 12 || y < 2000)
            {
                return false;
            }

            var next = new DateTime(y, m, 1).AddMonths(1);
            month = next.Month;
            year = next.Year;
            return true;
        }

        [RelayCommand]
        private void PrintInvoicesByRangeWithDialog()
        {
            var max = Customers.Any() ? Customers.Max(c => c.SequenceNumber) : 1;
            var range = _ui.ShowPrintRangeDialog(defaultFrom: 1, defaultTo: max);
            if (range == null)
            {
                return;
            }

            var from = Math.Min(range.FromNumber, range.ToNumber);
            var to = Math.Max(range.FromNumber, range.ToNumber);

            var customers = Customers
                .Where(c => c.SequenceNumber >= from && c.SequenceNumber <= to)
                .OrderBy(c => c.SequenceNumber)
                .ToList();

            if (customers.Count == 0)
            {
                _ui.ShowMessage("In theo số phiếu", "Không có khách nào trong khoảng số phiếu đã chọn.");
                return;
            }

            var folder = _ui.ShowFolderPickerDialog("Chọn thư mục để lưu các hóa đơn");
            if (string.IsNullOrWhiteSpace(folder))
            {
                return;
            }

            try
            {
                var templatePath = _ui.GetInvoiceTemplatePath();

                foreach (var customer in customers)
                {
                    var namePart = string.IsNullOrWhiteSpace(customer.Name) ? "Khach" : customer.Name;
                    var meterPart = string.IsNullOrWhiteSpace(customer.MeterNumber) ? string.Empty : $" - {customer.MeterNumber}";
                    var safeName = MakeSafeFileName($"{customer.SequenceNumber:0000} - {namePart}{meterPart}");
                    var filePath = Path.Combine(folder, $"Hoa don - {safeName}.xlsx");

                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        filePath,
                        customer,
                        PeriodLabel,
                        InvoiceIssuer);
                }

                _ui.ShowMessage("In theo số phiếu", $"Đã tạo {customers.Count} hóa đơn trong thư mục:\n{folder}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi in theo số phiếu", ex.Message);
            }
        }

        [RelayCommand]
        private void ExportSelectionToExcel(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new WarningException("Đường dẫn lưu file Excel không hợp lệ.");
            }

            if (SelectedCustomer == null)
            {
                throw new WarningException("Hãy chọn một khách trong bảng trước khi in.");
            }

            var list = new[] { SelectedCustomer };

            var templatePath = _ui.GetSummaryTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                list,
                PeriodLabel,
                InvoiceIssuer);
        }

        // Export the filtered list to the summary template (for grouped printing).
        [RelayCommand]
        private void ExportFilteredToExcel(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new WarningException("Đường dẫn lưu file Excel không hợp lệ.");
            }

            var filtered = GetCurrentViewCustomers();
            if (filtered.Count == 0)
            {
                throw new WarningException("Không có dòng nào trong danh sách hiện tại để in.");
            }

            var templatePath = _ui.GetSummaryTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                filtered,
                PeriodLabel,
                InvoiceIssuer);
        }

        [RelayCommand]
        private async Task PrintInvoiceWithDialog()
        {
            try
            {
                if (SelectedCustomer == null)
                {
                    _ui.ShowMessage("In Excel", "Hãy chọn một khách trong bảng trước khi in.");
                    return;
                }

                var outputPath = _ui.ShowSaveExcelFileDialog("Hoa don tien dien.xlsx");
                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    return;
                }

                var templatePath = _ui.GetInvoiceTemplatePath();
                using var busy = _ui.ShowBusyScope("In Excel", "Đang tạo hóa đơn, vui lòng chờ...");

                var customer = SelectedCustomer;
                await Task.Run(() => InvoiceExcelExportService.ExportInvoice(
                    templatePath,
                    outputPath,
                    customer,
                    PeriodLabel,
                    InvoiceIssuer));

                var openResult = _ui.Confirm(
                    "In Excel",
                    $"Đã tạo file Excel tại:\n{outputPath}\n\nBạn có muốn mở file này bằng Excel để xem / chỉnh sửa và in không?");

                if (openResult)
                {
                    _ui.OpenWithDefaultApp(outputPath);
                }
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("In Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi in Excel", ex.Message);
            }
        }

        [RelayCommand]
        private async Task PrintInvoicePdfWithDialog()
        {
            if (SelectedCustomer == null)
            {
                _ui.ShowMessage("Xuất PDF", "Hãy chọn một khách trong bảng trước khi xuất.");
                return;
            }

            var outputPdfPath = _ui.ShowSavePdfFileDialog("Hoa don tien dien.pdf");
            if (string.IsNullOrWhiteSpace(outputPdfPath))
            {
                return;
            }

            string? tempXlsxPath = null;
            try
            {
                var templatePath = _ui.GetInvoiceTemplatePath();
                tempXlsxPath = Path.Combine(Path.GetTempPath(), $"ElectricCalculation_Invoice_{Guid.NewGuid():N}.xlsx");

                using var busy = _ui.ShowBusyScope("Xuất PDF", "Đang tạo hóa đơn, vui lòng chờ...");

                var customer = SelectedCustomer;
                await Task.Run(() => InvoiceExcelExportService.ExportInvoice(
                    templatePath,
                    tempXlsxPath,
                    customer,
                    PeriodLabel,
                    InvoiceIssuer));

                await RunStaAsync(() => ExcelPdfExportService.ExportWorkbookToPdf(tempXlsxPath, outputPdfPath));

                var openResult = _ui.Confirm(
                    "Xuất PDF",
                    $"Đã tạo file PDF tại:\n{outputPdfPath}\n\nMở file này không?");

                if (openResult)
                {
                    _ui.OpenWithDefaultApp(outputPdfPath);
                }
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Xuất PDF", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);

                if (!string.IsNullOrWhiteSpace(tempXlsxPath) && File.Exists(tempXlsxPath))
                {
                    try
                    {
                        var fallbackXlsxPath = Path.ChangeExtension(outputPdfPath, ".xlsx");
                        File.Copy(tempXlsxPath, fallbackXlsxPath, overwrite: true);
                        _ui.ShowMessage(
                            "Xuất PDF",
                            $"Không thể xuất PDF (có thể máy chưa cài Microsoft Excel).\n\nĐã lưu file Excel tại:\n{fallbackXlsxPath}");
                        return;
                    }
                    catch
                    {
                        // Ignore fallback errors.
                    }
                }

                _ui.ShowMessage("Lỗi xuất PDF", ex.Message);
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempXlsxPath) && File.Exists(tempXlsxPath))
                    {
                        File.Delete(tempXlsxPath);
                    }
                }
                catch
                {
                    // Ignore temp cleanup errors.
                }
            }
        }

        [RelayCommand]
        private async Task PrintAllInvoicesWithDialog()
        {
            var outputPath = _ui.ShowSaveExcelFileDialog("Hoa don tien dien (tat ca khach).xlsx");
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            var customers = GetCurrentViewCustomers();
            if (customers.Count == 0)
            {
                _ui.ShowMessage("In Excel", "Không có dữ liệu trong danh sách hiện tại để in.");
                return;
            }

            var confirm = _ui.Confirm(
                "In Excel",
                $"Bạn có chắc muốn in hóa đơn cho tất cả {customers.Count} khách đang hiển thị trên datagrid?");

            if (!confirm)
            {
                return;
            }

            var templatePath = _ui.GetInvoiceTemplatePath();
            using var busy = _ui.ShowBusyScope("In Excel", $"Đang xử lý {customers.Count} khách, vui lòng chờ...");

            try
            {
                await Task.Run(() => InvoiceExcelExportService.ExportInvoicesToWorkbook(
                    templatePath,
                    outputPath,
                    customers,
                    PeriodLabel,
                    InvoiceIssuer));

                var openResult = _ui.Confirm(
                    "In Excel",
                    $"Đã tạo file Excel tại:\n{outputPath}\n\nBạn có muốn mở file này bằng Excel để xem / chỉnh sửa và in không?");

                if (openResult)
                {
                    _ui.OpenWithDefaultApp(outputPath);
                }
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("In Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi in Excel", ex.Message);
            }
        }

        [RelayCommand]
        private async Task PrintAllInvoicesPdfWithDialog()
        {
            var outputPdfPath = _ui.ShowSavePdfFileDialog("Hoa don tien dien (tat ca khach).pdf");
            if (string.IsNullOrWhiteSpace(outputPdfPath))
            {
                return;
            }

            string? tempXlsxPath = null;
            try
            {
                var customers = GetCurrentViewCustomers();

                if (customers.Count == 0)
                {
                    _ui.ShowMessage("Xuất PDF", "Không có dữ liệu để xuất.");
                    return;
                }

                var confirm = _ui.Confirm(
                    "Xuất PDF",
                    $"Bạn có chắc muốn in hóa đơn cho tất cả {customers.Count} khách đang hiển thị trên datagrid?");

                if (!confirm)
                {
                    return;
                }

                var templatePath = _ui.GetInvoiceTemplatePath();
                tempXlsxPath = Path.Combine(Path.GetTempPath(), $"ElectricCalculation_Invoices_{Guid.NewGuid():N}.xlsx");

                using var busy = _ui.ShowBusyScope("Xuất PDF", $"Đang xử lý {customers.Count} khách, vui lòng chờ...");

                await Task.Run(() => InvoiceExcelExportService.ExportInvoicesToWorkbook(
                    templatePath,
                    tempXlsxPath,
                    customers,
                    PeriodLabel,
                    InvoiceIssuer));

                await RunStaAsync(() => ExcelPdfExportService.ExportWorkbookToPdf(tempXlsxPath, outputPdfPath));

                var openResult = _ui.Confirm(
                    "Xuất PDF",
                    $"Đã tạo file PDF tại:\n{outputPdfPath}\n\nMở file này không?");

                if (openResult)
                {
                    _ui.OpenWithDefaultApp(outputPdfPath);
                }
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Xuất PDF", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);

                if (!string.IsNullOrWhiteSpace(tempXlsxPath) && File.Exists(tempXlsxPath))
                {
                    try
                    {
                        var fallbackXlsxPath = Path.ChangeExtension(outputPdfPath, ".xlsx");
                        File.Copy(tempXlsxPath, fallbackXlsxPath, overwrite: true);
                        _ui.ShowMessage(
                            "Xuất PDF",
                            $"Không thể xuất PDF (có thể máy chưa cài Microsoft Excel).\n\nĐã lưu file Excel tại:\n{fallbackXlsxPath}");
                        return;
                    }
                    catch
                    {
                        // Ignore fallback errors.
                    }
                }

                _ui.ShowMessage("Lỗi xuất PDF", ex.Message);
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempXlsxPath) && File.Exists(tempXlsxPath))
                    {
                        File.Delete(tempXlsxPath);
                    }
                }
                catch
                {
                    // Ignore temp cleanup errors.
                }
            }
        }

        private static Task RunStaAsync(Action action)
        {
            var tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);

            var thread = new Thread(() =>
            {
                try
                {
                    action();
                    tcs.TrySetResult(null);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            })
            {
                IsBackground = true
            };

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            return tcs.Task;
        }

        [RelayCommand]
        private void ShowReport()
        {
            var currentItems = GetCurrentViewCustomers();
            if (currentItems.Count == 0)
            {
                _ui.ShowMessage("Báo cáo", "Không có dữ liệu trong danh sách hiện tại để lập báo cáo.");
                return;
            }

            _ui.ShowReportWindow(PeriodLabel, currentItems, InvoiceIssuer);
        }

        private void Customers_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (!_suppressDirty)
            {
                IsDirty = true;
            }

            if (e.Action == NotifyCollectionChangedAction.Reset)
            {
                foreach (var customer in _subscribedCustomers)
                {
                    customer.PropertyChanged -= Customer_PropertyChanged;
                }

                _subscribedCustomers.Clear();

                foreach (var customer in Customers)
                {
                    if (_subscribedCustomers.Add(customer))
                    {
                        customer.PropertyChanged += Customer_PropertyChanged;
                    }
                }
            }
            else
            {
                if (e.OldItems != null)
                {
                    foreach (var item in e.OldItems.OfType<Customer>())
                    {
                        if (_subscribedCustomers.Remove(item))
                        {
                            item.PropertyChanged -= Customer_PropertyChanged;
                        }
                    }
                }

                if (e.NewItems != null)
                {
                    foreach (var item in e.NewItems.OfType<Customer>())
                    {
                        if (_subscribedCustomers.Add(item))
                        {
                            item.PropertyChanged += Customer_PropertyChanged;
                        }
                    }
                }
            }

            NotifySummaryChanged();
        }

        private void Customer_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_suppressDirty)
            {
                return;
            }

            IsDirty = true;

            if (e.PropertyName == nameof(Customer.Consumption) ||
                e.PropertyName == nameof(Customer.ChargeableKwh) ||
                e.PropertyName == nameof(Customer.Amount) ||
                e.PropertyName == nameof(Customer.IsMissingReading) ||
                e.PropertyName == nameof(Customer.HasReadingError) ||
                e.PropertyName == nameof(Customer.HasUsageWarning))
            {
                NotifySummaryChanged();
            }

            if (e.PropertyName == nameof(Customer.GroupName))
            {
                UpdateGroupOptions();
                CustomersView.Refresh();
                NotifySummaryChanged();
            }
        }

        private bool FilterCustomer(object? item)
        {
            if (item is not Customer customer)
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(SelectedGroup) &&
                !string.Equals(SelectedGroup.Trim(), AllGroupsOption, StringComparison.CurrentCultureIgnoreCase))
            {
                var selectedKey = NormalizeGroupKey(SelectedGroup);
                var groupKey = NormalizeGroupKey(customer.GroupName);
                if (!string.Equals(groupKey, selectedKey, StringComparison.CurrentCultureIgnoreCase))
                {
                    return false;
                }
            }

            if (FilterMissing || FilterWarning || FilterError)
            {
                var matchesStatus =
                    (FilterMissing && customer.IsMissingReading) ||
                    (FilterWarning && customer.HasUsageWarning) ||
                    (FilterError && customer.HasReadingError);

                if (!matchesStatus)
                {
                    return false;
                }
            }

            if (string.IsNullOrEmpty(_normalizedSearchKeyword))
            {
                return true;
            }

            var keyword = _normalizedSearchKeyword;

            // Fast search always includes the primary fields: name + meter + location.
            if (ContainsKeyword(customer.Name, keyword) ||
                ContainsKeyword(customer.MeterNumber, keyword) ||
                ContainsKeyword(customer.Location, keyword))
            {
                return true;
            }

            string fieldValue = GetSearchFieldValue(customer, _selectedSearchFieldIndex);

            return ContainsKeyword(fieldValue, keyword);
        }

        [RelayCommand]
        private void ApplySearch()
        {
            ApplySearchCache(SearchText, SelectedSearchField);
            RefreshCustomersViewAfterFilterChanged();
        }

        partial void OnSearchTextChanged(string value)
        {
            // Search is applied explicitly when user presses Enter.
        }

        partial void OnSelectedSearchFieldChanged(string value)
        {
            // Search field is applied explicitly when user presses Enter.
        }

        partial void OnFilterMissingChanged(bool value)
        {
            RefreshCustomersViewAfterFilterChanged();
        }

        partial void OnFilterWarningChanged(bool value)
        {
            RefreshCustomersViewAfterFilterChanged();
        }

        partial void OnFilterErrorChanged(bool value)
        {
            RefreshCustomersViewAfterFilterChanged();
        }

        partial void OnSelectedGroupChanged(string value)
        {
            RefreshCustomersViewAfterFilterChanged();
        }

        private void UpdateGroupOptions()
        {
            var groups = Customers
                .Select(c => NormalizeGroupKey(c.GroupName))
                .Where(g => !string.IsNullOrWhiteSpace(g))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(g => g, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var selectedKey = NormalizeGroupKey(SelectedGroup);

            GroupOptions.Clear();
            GroupOptions.Add(AllGroupsOption);
            foreach (var group in groups)
            {
                GroupOptions.Add(group);
            }

            if (string.IsNullOrWhiteSpace(selectedKey) ||
                string.Equals(selectedKey, NormalizeGroupKey(AllGroupsOption), StringComparison.CurrentCultureIgnoreCase))
            {
                SelectedGroup = AllGroupsOption;
                return;
            }

            var match = GroupOptions
                .Skip(1)
                .FirstOrDefault(g => string.Equals(NormalizeGroupKey(g), selectedKey, StringComparison.CurrentCultureIgnoreCase));

            SelectedGroup = match ?? AllGroupsOption;
        }

        private static string NormalizeGroupKey(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var text = value
                .Normalize(NormalizationForm.FormKC)
                .Replace('\u00A0', ' ') // no-break space
                .Replace('\u2007', ' ') // figure space
                .Replace('\u202F', ' '); // narrow no-break space

            text = ReplaceDashVariants(text);
            text = CollapseWhitespace(text).Trim();
            text = Regex.Replace(text, @"\s*-\s*", " - ");
            text = CollapseWhitespace(text).Trim();

            return text;
        }

        private static string ReplaceDashVariants(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            foreach (var ch in text)
            {
                sb.Append(ch switch
                {
                    '\u2010' => '-', // hyphen
                    '\u2011' => '-', // non-breaking hyphen
                    '\u2012' => '-', // figure dash
                    '\u2013' => '-', // en dash
                    '\u2014' => '-', // em dash
                    '\u2212' => '-', // minus sign
                    _ => ch
                });
            }

            return sb.ToString();
        }

        private static string CollapseWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(text.Length);
            var inWhitespace = false;

            foreach (var ch in text)
            {
                if (char.IsWhiteSpace(ch))
                {
                    if (!inWhitespace)
                    {
                        sb.Append(' ');
                        inWhitespace = true;
                    }

                    continue;
                }

                inWhitespace = false;
                sb.Append(ch);
            }

            return sb.ToString();
        }

        private static bool ContainsKeyword(string? value, string keyword)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   value.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private static string GetSearchFieldValue(Customer customer, int searchFieldIndex)
        {
            return searchFieldIndex switch
            {
                SearchFieldGroupIndex => customer.GroupName ?? string.Empty,
                SearchFieldCategoryIndex => customer.Category ?? string.Empty,
                SearchFieldAddressIndex => customer.Address ?? string.Empty,
                SearchFieldPhoneIndex => customer.Phone ?? string.Empty,
                SearchFieldMeterIndex => customer.MeterNumber ?? string.Empty,
                _ => customer.Name ?? string.Empty
            };
        }



        partial void OnPeriodLabelChanged(string value)
        {
            if (!_suppressDirty)
            {
                IsDirty = true;
            }
        }

        partial void OnInvoiceIssuerChanged(string value)
        {
            if (!_suppressDirty)
            {
                IsDirty = true;
            }
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return string.IsNullOrWhiteSpace(name) ? "Hoa_don" : name;
        }

        private static Customer CloneCustomer(Customer source)
        {
            return new Customer
            {
                SequenceNumber = source.SequenceNumber,
                Name = source.Name ?? string.Empty,
                GroupName = source.GroupName ?? string.Empty,
                Category = source.Category ?? string.Empty,
                Address = source.Address ?? string.Empty,
                RepresentativeName = source.RepresentativeName ?? string.Empty,
                HouseholdPhone = source.HouseholdPhone ?? string.Empty,
                Phone = source.Phone ?? string.Empty,
                BuildingName = source.BuildingName ?? string.Empty,
                MeterNumber = source.MeterNumber ?? string.Empty,
                Substation = source.Substation ?? string.Empty,
                Page = source.Page ?? string.Empty,
                PerformedBy = source.PerformedBy ?? string.Empty,
                Location = source.Location ?? string.Empty,
                PreviousIndex = source.PreviousIndex,
                CurrentIndex = source.CurrentIndex,
                Multiplier = source.Multiplier <= 0 ? 1 : source.Multiplier,
                SubsidizedKwh = source.SubsidizedKwh,
                UnitPrice = source.UnitPrice
            };
        }

    }
}
