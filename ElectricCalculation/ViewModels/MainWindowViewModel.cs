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
            _ui = new UiService();

            CustomersView = CollectionViewSource.GetDefaultView(Customers);
            CustomersView.Filter = FilterCustomer;
            Customers.CollectionChanged += Customers_CollectionChanged;

            if (SearchFields.Count > 0)
            {
                SelectedSearchField = SearchFields[0];
            }

            PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
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
                Customers.Add(item);
            }

            if (!string.IsNullOrWhiteSpace(warningMessage))
            {
                throw new WarningException(warningMessage);
            }

            if (Customers.Count == 0)
            {
                throw new WarningException("Import xong nhưng không có dòng dữ liệu nào. Hãy kiểm tra lại sheet 'Data' trong file Excel nguồn.");
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
                PeriodLabel);
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
                PeriodLabel);
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
                PeriodLabel);
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

            _ui.ShowReportWindow(PeriodLabel, currentItems);
        }

        private void Customers_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
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

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return string.IsNullOrWhiteSpace(name) ? "Hoa_don" : name;
        }

    }
}
