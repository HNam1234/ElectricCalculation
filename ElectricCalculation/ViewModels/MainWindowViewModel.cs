using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public partial class MainWindowViewModel : ObservableObject
    {
        private string? _lastImportedExcelPath;

        [ObservableProperty]
        private string periodLabel = string.Empty;

        [ObservableProperty]
        private string searchText = string.Empty;

        [ObservableProperty]
        private string selectedSearchField = string.Empty;

        [ObservableProperty]
        private Customer? selectedCustomer;

        // Người lập đơn (in lên hóa đơn Excel)
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
            Customers.Add(new Customer());
        }

        [RelayCommand]
        private void ClearAll()
        {
            if (Customers.Count == 0)
            {
                return;
            }

            Customers.Clear();
        }

        // Thống kê theo danh sách đang hiển thị (đã lọc)
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

            _lastImportedExcelPath = filePath;

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

            // Luôn dùng file template mặc định trong solution,
            // không phụ thuộc vào file Excel vừa import.
            var templatePath = GetDefaultTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                Customers);
        }

        // Export 1 khách (dòng đang chọn) ra Excel theo template tổng hợp mặc định để in bằng Excel.
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

            var templatePath = GetDefaultTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                list);
        }

        // Export toàn bộ danh sách đang lọc ra Excel theo template tổng hợp mặc định (in theo nhóm/cụm).
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

            var templatePath = GetDefaultTemplatePath();

            Services.ExcelExportService.ExportToFile(
                templatePath,
                outputPath,
                filtered);
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

        private static string GetDefaultTemplatePath()
        {
            // Mặc định dùng file mẫu tổng hợp cạnh solution:
            // "Bảng tổng hợp thu tháng 6 năm 2025.xlsx"
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // bin/Debug/net8.0-windows -> quay lên thư mục solution
            var rootDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
            var templatePath = Path.Combine(rootDir, "Bảng tổng hợp thu tháng 6 năm 2025.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new WarningException("Không tìm thấy file Excel template mặc định 'Bảng tổng hợp thu tháng 6 năm 2025.xlsx' cạnh solution.");
            }

            return templatePath;
        }
    }
}
