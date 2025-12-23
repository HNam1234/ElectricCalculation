using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

        public MainWindowViewModel()
        {
            _suppressDirty = true;
            _ui = new UiService();
            _settings = AppSettingsService.Load();

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

        public void ImportFromExcelFile(string filePath)
        {
            try
            {
                _suppressDirty = true;
                ImportFromExcel(filePath);
                IsDirty = true;
                LoadedSnapshotPath = null;
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
            Customers.Add(customer);
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
            if (!applyWhen || customer == null)
            {
                return;
            }

            var s = _settings ?? new AppSettings();
            var overwrite = s.OverrideExistingValues;

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

            if (overwrite || (customer.SubsidizedKwh <= 0 && customer.SubsidizedPercent <= 0))
            {
                if (s.DefaultSubsidizedPercent > 0)
                {
                    customer.SubsidizedPercent = s.DefaultSubsidizedPercent;
                    if (overwrite)
                    {
                        customer.SubsidizedKwh = 0;
                    }
                }
                else if (s.DefaultSubsidizedKwh > 0)
                {
                    customer.SubsidizedKwh = s.DefaultSubsidizedKwh;
                    if (overwrite)
                    {
                        customer.SubsidizedPercent = 0;
                    }
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
                SubsidizedPercent = source.SubsidizedPercent,
                UnitPrice = source.UnitPrice
            };
        }

    }
}
