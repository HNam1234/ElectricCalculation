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
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
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

        [ObservableProperty]
        private bool isRouteEntryMode = true;

        [ObservableProperty]
        private bool filterMissing;

        [ObservableProperty]
        private bool filterWarning;

        [ObservableProperty]
        private bool filterError;

        public bool IsDetailMode => !IsFastEntryMode;

        public string RouteEntryProgressText => BuildRouteEntryProgressText();

        public ObservableCollection<Customer> Customers { get; } = new();

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

            _undoRedo.StateChanged += (_, _) =>
            {
                UndoCommand.NotifyCanExecuteChanged();
                RedoCommand.NotifyCanExecuteChanged();
            };

            CustomersView = CollectionViewSource.GetDefaultView(Customers);
            CustomersView.Filter = FilterCustomer;
            Customers.CollectionChanged += Customers_CollectionChanged;
            ApplyRouteEntrySort();

            if (SearchFields.Count > 0)
            {
                SelectedSearchField = SearchFields[0];
            }

            PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
            IsDirty = false;
            _suppressDirty = false;
        }

        public void ImportFromExcelFile(string filePath)
        {
            try
            {
                _suppressDirty = true;
                _undoRedo.Clear();

                try
                {
                    ImportFromExcel(filePath);
                    IsDirty = true;
                    LoadedSnapshotPath = null;
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

                Customers.Clear();
                foreach (var c in customers)
                {
                    Customers.Add(c);
                }

                if (!string.IsNullOrWhiteSpace(period))
                {
                    PeriodLabel = period;
                }

                SelectedCustomer = null;
                CurrentDataFilePath = setCurrentDataFilePath ? filePath : null;
                LoadedSnapshotPath = null;
                IsDirty = false;
                RefreshUsageAverages();
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

            ApplyDefaultsIfNeeded(customer, applyWhen: _settings.ApplyDefaultsOnNewRow);
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
            try
            {
                if (CustomersView == null)
                {
                    return TryGetCustomersSnapshot();
                }

                return CustomersView
                    .Cast<object>()
                    .OfType<Customer>()
                    .ToList();
            }
            catch
            {
                return TryGetCustomersSnapshot();
            }
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

        public int CustomerCount => TryGetViewCustomersSnapshot().Count;

        public decimal TotalConsumption => TryGetViewCustomersSnapshot().Sum(c => c.Consumption);

        public decimal TotalChargeableKwh => TryGetViewCustomersSnapshot().Sum(c => c.ChargeableKwh);

        public decimal TotalAmount => TryGetViewCustomersSnapshot().Sum(c => c.Amount);

        public int CompletedCount => TryGetViewCustomersSnapshot().Count(c => c.CurrentIndex != null);

        public int MissingCount => TryGetViewCustomersSnapshot().Count(c => c.IsMissingReading);

        public int WarningCount => TryGetViewCustomersSnapshot().Count(c => c.HasUsageWarning);

        public int ErrorCount => TryGetViewCustomersSnapshot().Count(c => c.HasReadingError);

        public double CompletionRatio
        {
            get
            {
                var customers = TryGetViewCustomersSnapshot();
                if (customers.Count <= 0)
                {
                    return 0;
                }

                var completed = customers.Count(c => c.CurrentIndex != null);
                var ratio = (double)completed / customers.Count;

                if (double.IsNaN(ratio) || double.IsInfinity(ratio))
                {
                    return 0;
                }

                return Math.Max(0, Math.Min(1, ratio));
            }
        }

        private void NotifySummaryChanged()
        {
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

        private void NotifyRouteEntryProgressChanged()
        {
            OnPropertyChanged(nameof(RouteEntryProgressText));
        }

        private string BuildRouteEntryProgressText()
        {
            if (!IsRouteEntryMode)
            {
                return string.Empty;
            }

            var current = SelectedCustomer;
            if (current == null)
            {
                return string.Empty;
            }

            var locationRaw = current.Location ?? string.Empty;
            var locationDisplay = string.IsNullOrWhiteSpace(locationRaw) ? "(Chưa có vị trí)" : locationRaw.Trim();

            var customers = TryGetViewCustomersSnapshot();
            if (customers.Count == 0)
            {
                return $"Đang nhập: {locationDisplay}";
            }

            var locationKey = locationRaw.Trim();
            var group = customers
                .Where(c => string.Equals((c.Location ?? string.Empty).Trim(), locationKey, StringComparison.CurrentCultureIgnoreCase))
                .ToList();

            if (group.Count == 0)
            {
                return $"Đang nhập: {locationDisplay}";
            }

            var index = group.FindIndex(c => ReferenceEquals(c, current));
            if (index < 0)
            {
                index = 0;
            }

            return $"Đang nhập: {locationDisplay} ({index + 1}/{group.Count})";
        }

        private void ApplyRouteEntrySort()
        {
            if (CustomersView is not ListCollectionView view)
            {
                return;
            }

            view.CustomSort = IsRouteEntryMode ? CustomerRouteComparer.Instance : null;
            CustomersView.Refresh();
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

            Customers.Clear();

            var imported = Services.ExcelImportService.ImportFromFile(filePath, out var warningMessage);
            foreach (var item in imported)
            {
                ApplyDefaultsIfNeeded(item, applyWhen: _settings.ApplyDefaultsOnImport);
                Customers.Add(item);
            }

            if (!string.IsNullOrWhiteSpace(warningMessage))
            {
                Debug.WriteLine(warningMessage);
                if (!warningMessage.Contains("Khong thay dong tieu de", StringComparison.OrdinalIgnoreCase))
                {
                    _ui.ShowMessage("Thong bao import Excel", warningMessage);
                }
            }

            if (Customers.Count == 0)
            {
                throw new WarningException("Import xong nhưng không có dòng dữ liệu nào. Hãy kiểm tra lại sheet 'Data' trong file Excel nguồn.");
            }

            if (!_suppressDirty)
            {
                IsDirty = true;
            }

            RefreshUsageAverages();
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

        private void ApplyDefaultsIfNeeded(Customer customer, bool applyWhen)
        {
            var s = _settings ?? new AppSettings();
            ApplyDefaultsIfNeeded(customer, applyWhen, s);
        }

        private void ApplyDefaultsIfNeeded(Customer customer, bool applyWhen, AppSettings settings)
        {
            if (!applyWhen || customer == null)
            {
                return;
            }

            var s = settings ?? new AppSettings();
            var overwrite = true;

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
        private void ExportToExcel(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                throw new WarningException("Đường dẫn lưu file Excel không hợp lệ.");
            }

            if (Customers.Count == 0)
            {
                throw new WarningException("Không có dữ liệu trong bảng để export.");
            }

            // Always use the solution template,
            // not the last imported workbook.
            var templatePath = _ui.GetSummaryTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                Customers,
                PeriodLabel,
                InvoiceIssuer);
        }

        // Export the selected customer to the summary template (for printing in Excel).
        [RelayCommand]
        private void ExportToExcelWithDialog()
        {
            var outputPath = _ui.ShowSaveExcelFileDialog("Bang tong hop dien.xlsx");
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            try
            {
                ExportToExcel(outputPath);
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("Cảnh báo export Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi export Excel", ex.Message);
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

            Customers.Clear();
            foreach (var c in newCustomers)
            {
                Customers.Add(c);
            }

            SelectedCustomer = null;
            PeriodLabel = vm.PeriodLabel;
            CurrentDataFilePath = null;
            LoadedSnapshotPath = null;
            RefreshUsageAverages();
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

            var filtered = CustomersView.Cast<Customer>().ToList();
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
        private void PrintInvoiceWithDialog()
        {
            var outputPath = _ui.ShowSaveExcelFileDialog("Hoa don tien dien.xlsx");
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            try
            {
                if (SelectedCustomer != null)
                {
                    var templatePath = _ui.GetInvoiceTemplatePath();

                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        outputPath,
                        SelectedCustomer,
                        PeriodLabel,
                        InvoiceIssuer);
                }
                else
                {
                    ExportFilteredToExcel(outputPath);
                }

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
        private void PrintInvoicePdfWithDialog()
        {
            var outputPdfPath = _ui.ShowSavePdfFileDialog("Hoa don tien dien.pdf");
            if (string.IsNullOrWhiteSpace(outputPdfPath))
            {
                return;
            }

            string? tempXlsxPath = null;
            try
            {
                var customers = SelectedCustomer != null
                    ? new List<Customer> { SelectedCustomer }
                    : CustomersView.Cast<Customer>().ToList();

                if (customers.Count == 0)
                {
                    _ui.ShowMessage("Xuất PDF", "Không có dữ liệu để xuất.");
                    return;
                }

                var templatePath = _ui.GetInvoiceTemplatePath();
                tempXlsxPath = Path.Combine(Path.GetTempPath(), $"ElectricCalculation_Invoice_{Guid.NewGuid():N}.xlsx");

                if (customers.Count == 1)
                {
                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        tempXlsxPath,
                        customers[0],
                        PeriodLabel,
                        InvoiceIssuer);
                }
                else
                {
                    InvoiceExcelExportService.ExportInvoicesToWorkbook(
                        templatePath,
                        tempXlsxPath,
                        customers,
                        PeriodLabel,
                        InvoiceIssuer);
                }

                ExcelPdfExportService.ExportWorkbookToPdf(tempXlsxPath, outputPdfPath);

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
        private void PrintMultipleInvoicesWithDialog(IList? selectedItems)
        {
            try
            {
                var selectedCustomers = selectedItems?.OfType<Customer>().ToList();

                var customers = selectedCustomers != null && selectedCustomers.Count > 0
                    ? selectedCustomers
                    : CustomersView.Cast<Customer>().ToList();

                if (customers.Count == 0)
                {
                    _ui.ShowMessage("In nhiều phiếu", "Không có dữ liệu trong danh sách hiện tại để in.");
                    return;
                }

                var folder = _ui.ShowFolderPickerDialog("Chọn thư mục để lưu các hóa đơn");
                if (string.IsNullOrWhiteSpace(folder))
                {
                    return;
                }

                var templatePath = _ui.GetInvoiceTemplatePath();

                foreach (var customer in customers)
                {
                    var namePart = string.IsNullOrWhiteSpace(customer.Name)
                        ? "Khach"
                        : customer.Name;

                    var meterPart = string.IsNullOrWhiteSpace(customer.MeterNumber)
                        ? string.Empty
                        : $" - {customer.MeterNumber}";

                    var safeName = MakeSafeFileName($"{namePart}{meterPart}");
                    var filePath = Path.Combine(folder, $"Hoa don - {safeName}.xlsx");

                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        filePath,
                        customer,
                        PeriodLabel,
                        InvoiceIssuer);
                }

                _ui.ShowMessage("In nhiều phiếu", $"Đã tạo {customers.Count} hóa đơn trong thư mục:\n{folder}");
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                _ui.ShowMessage("In nhiều phiếu", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                _ui.ShowMessage("Lỗi in nhiều phiếu", ex.Message);
            }
        }

        [RelayCommand]
        private void ShowReport()
        {
            var currentItems = CustomersView.Cast<Customer>().ToList();
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

            if (e.OldItems != null)
            {
                foreach (var item in e.OldItems.OfType<Customer>())
                {
                    item.PropertyChanged -= Customer_PropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (var item in e.NewItems.OfType<Customer>())
                {
                    item.PropertyChanged += Customer_PropertyChanged;
                }
            }

            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
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

            if (IsRouteEntryMode &&
                (e.PropertyName == nameof(Customer.Location) ||
                 e.PropertyName == nameof(Customer.Page) ||
                 e.PropertyName == nameof(Customer.SequenceNumber)))
            {
                CustomersView.Refresh();
                NotifyRouteEntryProgressChanged();
            }
        }

        private bool FilterCustomer(object? item)
        {
            if (item is not Customer customer)
            {
                return false;
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

            if (string.IsNullOrWhiteSpace(SearchText))
            {
                return true;
            }

            var keyword = SearchText.Trim();

            static bool ContainsKeyword(string? value, string keyword) =>
                !string.IsNullOrWhiteSpace(value) &&
                value.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0;

            // Fast search always includes the primary fields: name + meter + location.
            if (ContainsKeyword(customer.Name, keyword) ||
                ContainsKeyword(customer.MeterNumber, keyword) ||
                ContainsKeyword(customer.Location, keyword))
            {
                return true;
            }

            string fieldValue = SelectedSearchField switch
            {
                "Nhóm / Đơn vị" => customer.GroupName ?? string.Empty,
                "Loại" => customer.Category ?? string.Empty,
                "Địa chỉ" => customer.Address ?? string.Empty,
                "Số ĐT" => customer.Phone ?? string.Empty,
                "Số công tơ" => customer.MeterNumber ?? string.Empty,
                _ => customer.Name ?? string.Empty // Tên khách
            };

            return ContainsKeyword(fieldValue, keyword);
        }

        partial void OnSearchTextChanged(string value)
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnSelectedSearchFieldChanged(string value)
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnFilterMissingChanged(bool value)
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnFilterWarningChanged(bool value)
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnFilterErrorChanged(bool value)
        {
            CustomersView.Refresh();
            NotifySummaryChanged();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnSelectedCustomerChanged(Customer? value)
        {
            NotifyRouteEntryProgressChanged();
        }

        partial void OnIsRouteEntryModeChanged(bool value)
        {
            ApplyRouteEntrySort();
            NotifyRouteEntryProgressChanged();
        }

        partial void OnPeriodLabelChanged(string value)
        {
            if (!_suppressDirty)
            {
                IsDirty = true;
            }
        }

        [RelayCommand]
        private void PrintMultipleInvoicesPdfWithDialog(IList? selectedItems)
        {
            var outputPdfPath = _ui.ShowSavePdfFileDialog("Hoa don tien dien - nhieu khach.pdf");
            if (string.IsNullOrWhiteSpace(outputPdfPath))
            {
                return;
            }

            string? tempXlsxPath = null;
            try
            {
                var selectedCustomers = selectedItems?.OfType<Customer>().ToList();
                var customers = selectedCustomers != null && selectedCustomers.Count > 0
                    ? selectedCustomers
                    : CustomersView.Cast<Customer>().ToList();

                if (customers.Count == 0)
                {
                    _ui.ShowMessage("Xuất PDF", "Không có dữ liệu để xuất.");
                    return;
                }

                var templatePath = _ui.GetInvoiceTemplatePath();
                tempXlsxPath = Path.Combine(Path.GetTempPath(), $"ElectricCalculation_Invoices_{Guid.NewGuid():N}.xlsx");

                if (customers.Count == 1)
                {
                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        tempXlsxPath,
                        customers[0],
                        PeriodLabel,
                        InvoiceIssuer);
                }
                else
                {
                    InvoiceExcelExportService.ExportInvoicesToWorkbook(
                        templatePath,
                        tempXlsxPath,
                        customers,
                        PeriodLabel,
                        InvoiceIssuer);
                }

                ExcelPdfExportService.ExportWorkbookToPdf(tempXlsxPath, outputPdfPath);

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
