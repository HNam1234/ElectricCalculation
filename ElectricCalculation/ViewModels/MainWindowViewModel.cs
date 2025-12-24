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

        // Quick defaults (UI shortcut for AppSettings).
        [ObservableProperty]
        private bool isQuickDefaultsExpanded;

        [ObservableProperty]
        private string quickDefaultUnitPrice = "0";

        [ObservableProperty]
        private string quickDefaultMultiplier = "1";

        [ObservableProperty]
        private string quickDefaultSubsidizedKwh = "0";

        [ObservableProperty]
        private string quickDefaultPerformedBy = string.Empty;

        [ObservableProperty]
        private bool quickApplyDefaultsOnNewRow = true;

        [ObservableProperty]
        private bool quickApplyDefaultsOnImport = true;

        [ObservableProperty]
        private string quickDefaultsErrorMessage = string.Empty;

        [ObservableProperty]
        private bool isDirty;

        [ObservableProperty]
        private string? loadedSnapshotPath;

        [ObservableProperty]
        private bool showAdvancedColumns;

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
            LoadQuickDefaultsFromSettings();

            _undoRedo.StateChanged += (_, _) =>
            {
                UndoCommand.NotifyCanExecuteChanged();
                RedoCommand.NotifyCanExecuteChanged();
            };

            CustomersView = CollectionViewSource.GetDefaultView(Customers);
            CustomersView.Filter = FilterCustomer;
            Customers.CollectionChanged += Customers_CollectionChanged;

            if (SearchFields.Count > 0)
            {
                SelectedSearchField = SearchFields[0];
            }

            PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
            IsDirty = false;
            _suppressDirty = false;
        }

        private void LoadQuickDefaultsFromSettings()
        {
            var s = _settings ?? new AppSettings();
            QuickDefaultUnitPrice = s.DefaultUnitPrice.ToString("0.##", CultureInfo.CurrentCulture);
            QuickDefaultMultiplier = s.DefaultMultiplier.ToString("0.##", CultureInfo.CurrentCulture);
            QuickDefaultSubsidizedKwh = s.DefaultSubsidizedKwh.ToString("0.##", CultureInfo.CurrentCulture);
            QuickDefaultPerformedBy = s.DefaultPerformedBy ?? string.Empty;
            QuickApplyDefaultsOnNewRow = s.ApplyDefaultsOnNewRow;
            QuickApplyDefaultsOnImport = s.ApplyDefaultsOnImport;
            QuickDefaultsErrorMessage = string.Empty;
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
        public int CustomerCount => CustomersView.Cast<Customer>().Count();

        public decimal TotalConsumption => CustomersView.Cast<Customer>().Sum(c => c.Consumption);

        public decimal TotalChargeableKwh => CustomersView.Cast<Customer>().Sum(c => c.ChargeableKwh);

        public decimal TotalAmount => CustomersView.Cast<Customer>().Sum(c => c.Amount);

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
        private void ToggleQuickDefaults()
        {
            IsQuickDefaultsExpanded = !IsQuickDefaultsExpanded;
        }

        [RelayCommand]
        private void ApplyQuickDefaultsToSelected(IList? selectedItems)
        {
            if (!TryBuildSettingsFromQuickDefaults(out var updated))
            {
                return;
            }

            var targets = selectedItems?.OfType<Customer>().ToList();
            if (targets == null || targets.Count == 0)
            {
                _ui.ShowMessage("Nhập nhanh", "Hãy chọn 1 hoặc nhiều dòng để áp dụng.");
                return;
            }

            var actions = new List<IUndoableAction>();
            foreach (var customer in targets)
            {
                actions.AddRange(BuildDefaultsActions(customer, updated));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Nhập nhanh (dòng chọn)", actions));
        }

        [RelayCommand]
        private void ApplyQuickDefaultsToFiltered()
        {
            if (!TryBuildSettingsFromQuickDefaults(out var updated))
            {
                return;
            }

            var targets = CustomersView.Cast<Customer>().ToList();
            if (targets.Count == 0)
            {
                _ui.ShowMessage("Nhập nhanh", "Không có dữ liệu trong danh sách đang lọc để áp dụng.");
                return;
            }

            var actions = new List<IUndoableAction>();
            foreach (var customer in targets)
            {
                actions.AddRange(BuildDefaultsActions(customer, updated));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Nhập nhanh (đang lọc)", actions));
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

        [RelayCommand]
        private void ApplyQuickDefaultsToSameGroup(Customer? reference)
        {
            if (reference == null)
            {
                return;
            }

            var group = reference.GroupName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(group))
            {
                return;
            }

            if (!TryBuildSettingsFromQuickDefaults(out var updated))
            {
                return;
            }

            var targets = Customers.Where(c => string.Equals(c.GroupName, group, StringComparison.OrdinalIgnoreCase)).ToList();
            if (targets.Count == 0)
            {
                return;
            }

            var actions = new List<IUndoableAction>();
            foreach (var c in targets)
            {
                actions.AddRange(BuildDefaultsActions(c, updated));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Nhập nhanh theo nhóm", actions));
        }

        [RelayCommand]
        private void ApplyQuickDefaultsToSameCategory(Customer? reference)
        {
            if (reference == null)
            {
                return;
            }

            var category = reference.Category?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(category))
            {
                return;
            }

            if (!TryBuildSettingsFromQuickDefaults(out var updated))
            {
                return;
            }

            var targets = Customers.Where(c => string.Equals(c.Category, category, StringComparison.OrdinalIgnoreCase)).ToList();
            if (targets.Count == 0)
            {
                return;
            }

            var actions = new List<IUndoableAction>();
            foreach (var c in targets)
            {
                actions.AddRange(BuildDefaultsActions(c, updated));
            }

            if (actions.Count == 0)
            {
                return;
            }

            ExecuteUndoable(new CompositeUndoableAction("Nhập nhanh theo loại", actions));
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

        private static IReadOnlyList<IUndoableAction> BuildDefaultsActions(Customer customer, AppSettings settings)
        {
            var actions = new List<IUndoableAction>();

            if (settings.DefaultUnitPrice > 0 && customer.UnitPrice != settings.DefaultUnitPrice)
            {
                var before = customer.UnitPrice;
                var after = settings.DefaultUnitPrice;
                actions.Add(new DelegateUndoableAction("Đơn giá", () => customer.UnitPrice = before, () => customer.UnitPrice = after));
            }

            var multiplier = settings.DefaultMultiplier > 0 ? settings.DefaultMultiplier : 1m;
            if (customer.Multiplier != multiplier)
            {
                var before = customer.Multiplier;
                var after = multiplier;
                actions.Add(new DelegateUndoableAction("Hệ số", () => customer.Multiplier = before, () => customer.Multiplier = after));
            }

            if (settings.DefaultSubsidizedKwh > 0)
            {
                if (customer.SubsidizedKwh != settings.DefaultSubsidizedKwh)
                {
                    var beforeKwh = customer.SubsidizedKwh;
                    var afterKwh = settings.DefaultSubsidizedKwh;
                    actions.Add(new DelegateUndoableAction(
                        "Bao cấp (kWh)",
                        () =>
                        {
                            customer.SubsidizedKwh = beforeKwh;
                        },
                        () =>
                        {
                            customer.SubsidizedKwh = afterKwh;
                        }));
                }
            }

            if (!string.IsNullOrWhiteSpace(settings.DefaultPerformedBy) &&
                !string.Equals(customer.PerformedBy, settings.DefaultPerformedBy, StringComparison.Ordinal))
            {
                var before = customer.PerformedBy ?? string.Empty;
                var after = settings.DefaultPerformedBy;
                actions.Add(new DelegateUndoableAction("Người ghi", () => customer.PerformedBy = before, () => customer.PerformedBy = after));
            }

            return actions;
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
            var t = Nullable.GetUnderlyingType(targetType) ?? targetType;

            if (t == typeof(string))
            {
                value = text ?? string.Empty;
                return true;
            }

            if (string.IsNullOrWhiteSpace(text))
            {
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

        private bool TryBuildSettingsFromQuickDefaults(out AppSettings settings)
        {
            settings = _settings ?? new AppSettings();
            QuickDefaultsErrorMessage = string.Empty;

            if (!TryParseDecimal(QuickDefaultUnitPrice, out var unitPrice) ||
                !TryParseDecimal(QuickDefaultMultiplier, out var multiplier) ||
                !TryParseDecimal(QuickDefaultSubsidizedKwh, out var subsidizedKwh))
            {
                return false;
            }

            if (multiplier <= 0)
            {
                multiplier = 1;
            }

            settings = new AppSettings
            {
                DefaultUnitPrice = unitPrice,
                DefaultMultiplier = multiplier,
                DefaultSubsidizedKwh = subsidizedKwh,
                DefaultPerformedBy = QuickDefaultPerformedBy ?? string.Empty,
                ApplyDefaultsOnNewRow = QuickApplyDefaultsOnNewRow,
                ApplyDefaultsOnImport = QuickApplyDefaultsOnImport,
                OverrideExistingValues = true
            };

            return true;
        }

        private bool TryParseDecimal(string? text, out decimal value)
        {
            value = 0m;

            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            if (!decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value) &&
                !decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            {
                QuickDefaultsErrorMessage = $"Không đọc được giá trị: '{text}'.";
                return false;
            }

            if (value < 0)
            {
                QuickDefaultsErrorMessage = "Giá trị phải >= 0.";
                return false;
            }

            return true;
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
            var vm = _ui.ShowNewPeriodDialog();
            if (vm == null)
            {
                return;
            }

            if (vm.Month is < 1 or > 12 || vm.Year < 2000)
            {
                _ui.ShowMessage("Làm tháng mới", "Tháng/Năm không hợp lệ.");
                return;
            }

            var newCustomers = vm.CopyCustomers
                ? Customers.OrderBy(c => c.SequenceNumber).Select(CloneCustomer).ToList()
                : new List<Customer>();

            if (vm.CopyCustomers)
            {
                foreach (var c in newCustomers)
                {
                    if (vm.MoveCurrentToPrevious)
                    {
                        c.PreviousIndex = c.CurrentIndex;
                    }

                    if (vm.ResetCurrentToZero)
                    {
                        c.CurrentIndex = 0;
                    }
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

            OnPropertyChanged(nameof(CustomerCount));
            OnPropertyChanged(nameof(TotalConsumption));
            OnPropertyChanged(nameof(TotalChargeableKwh));
            OnPropertyChanged(nameof(TotalAmount));
        }

        private void Customer_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (!_suppressDirty)
            {
                IsDirty = true;
            }

            if (e.PropertyName == nameof(Customer.Consumption) ||
                e.PropertyName == nameof(Customer.ChargeableKwh) ||
                e.PropertyName == nameof(Customer.Amount))
            {
                OnPropertyChanged(nameof(TotalConsumption));
                OnPropertyChanged(nameof(TotalChargeableKwh));
                OnPropertyChanged(nameof(TotalAmount));
            }
        }

        private bool FilterCustomer(object? item)
        {
            if (item is not Customer customer)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SearchText))
            {
                return true;
            }

            var keyword = SearchText.Trim();

            string fieldValue = SelectedSearchField switch
            {
                "Nhóm / Đơn vị" => customer.GroupName ?? string.Empty,
                "Loại" => customer.Category ?? string.Empty,
                "Địa chỉ" => customer.Address ?? string.Empty,
                "Số ĐT" => customer.Phone ?? string.Empty,
                "Số công tơ" => customer.MeterNumber ?? string.Empty,
                _ => customer.Name ?? string.Empty // Tên khách
            };

            return fieldValue.IndexOf(keyword, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        partial void OnSearchTextChanged(string value)
        {
            CustomersView.Refresh();
            OnPropertyChanged(nameof(CustomerCount));
            OnPropertyChanged(nameof(TotalConsumption));
            OnPropertyChanged(nameof(TotalChargeableKwh));
            OnPropertyChanged(nameof(TotalAmount));
        }

        partial void OnSelectedSearchFieldChanged(string value)
        {
            CustomersView.Refresh();
            OnPropertyChanged(nameof(CustomerCount));
            OnPropertyChanged(nameof(TotalConsumption));
            OnPropertyChanged(nameof(TotalChargeableKwh));
            OnPropertyChanged(nameof(TotalAmount));
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
