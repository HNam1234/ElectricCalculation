using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public partial class ReportItem : ObservableObject
    {
        [ObservableProperty]
        private string groupName = string.Empty;

        [ObservableProperty]
        private int customerCount;

        [ObservableProperty]
        private decimal totalConsumption;

        [ObservableProperty]
        private decimal totalChargeableKwh;

        [ObservableProperty]
        private decimal totalAmount;

        [ObservableProperty]
        private double kwhBarRatio;

        [ObservableProperty]
        private double amountBarRatio;

        [ObservableProperty]
        private bool isSelected;
    }

    public partial class ReportViewModel : ObservableObject
    {
        private readonly UiService _ui;

        private readonly decimal _maxKwh;
        private readonly decimal _maxAmount;
        private readonly string _periodLabel;
        private readonly string _issuerName;
        private readonly List<Customer> _sourceCustomers;

        public string Title { get; }

        public ObservableCollection<ReportItem> Items { get; } = new();

        public IReadOnlyList<string> Metrics { get; } = new[] { "kWh", "Tiền điện" };

        [ObservableProperty]
        private string selectedMetric = "kWh";

        public string YAxisLabel => SelectedMetric == "Tiền điện" ? "Tiền (VNĐ)" : "kWh";

        public bool ShowKwh => !string.Equals(SelectedMetric, "Tiền điện", StringComparison.OrdinalIgnoreCase);

        public bool ShowAmount => !ShowKwh;

        public decimal GrandTotalConsumption => Items.Sum(i => i.TotalConsumption);

        public decimal GrandTotalChargeableKwh => Items.Sum(i => i.TotalChargeableKwh);

        public decimal GrandTotalAmount => Items.Sum(i => i.TotalAmount);

        [ObservableProperty]
        private ReportItem? selectedItem;

        [ObservableProperty]
        private bool? dialogResult;

        public string KwhTick100 => FormatTick(MaxKwh);
        public string KwhTick75 => FormatTick(MaxKwh * 0.75m);
        public string KwhTick50 => FormatTick(MaxKwh * 0.5m);
        public string KwhTick25 => FormatTick(MaxKwh * 0.25m);
        public string KwhTick0 => "0";

        public string AmountTick100 => FormatTick(MaxAmount);
        public string AmountTick75 => FormatTick(MaxAmount * 0.75m);
        public string AmountTick50 => FormatTick(MaxAmount * 0.5m);
        public string AmountTick25 => FormatTick(MaxAmount * 0.25m);
        public string AmountTick0 => "0";

        public decimal MaxKwh => _maxKwh;
        public decimal MaxAmount => _maxAmount;

        public IReadOnlyList<Customer> SourceCustomers => _sourceCustomers;

        partial void OnSelectedMetricChanged(string value)
        {
            OnPropertyChanged(nameof(ShowKwh));
            OnPropertyChanged(nameof(ShowAmount));
            OnPropertyChanged(nameof(YAxisLabel));
        }

        partial void OnSelectedItemChanged(ReportItem? value)
        {
            foreach (var item in Items)
            {
                item.IsSelected = item == value;
            }
        }

        [RelayCommand]
        private void Close()
        {
            DialogResult = false;
        }

        public ReportViewModel(string periodLabel, IEnumerable<Customer> customers)
            : this(periodLabel, customers, new UiService())
        {
        }

        public ReportViewModel(
            string periodLabel,
            IEnumerable<Customer> customers,
            UiService ui)
            : this(periodLabel, customers, issuerName: null, ui)
        {
        }

        public ReportViewModel(
            string periodLabel,
            IEnumerable<Customer> customers,
            string? issuerName,
            UiService ui)
        {
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));

            _periodLabel = periodLabel ?? string.Empty;
            _issuerName = issuerName?.Trim() ?? string.Empty;
            Title = $"Thống kê theo nhóm - {periodLabel}";

            _sourceCustomers = customers?.ToList() ?? new List<Customer>();

            var groups = _sourceCustomers
                .GroupBy(c => string.IsNullOrWhiteSpace(c.GroupName) ? "(Không có nhóm)" : c.GroupName)
                .OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase);

            foreach (var g in groups)
            {
                var item = new ReportItem
                {
                    GroupName = g.Key,
                    CustomerCount = g.Count(),
                    TotalConsumption = g.Sum(c => c.Consumption),
                    TotalChargeableKwh = g.Sum(c => c.ChargeableKwh),
                    TotalAmount = g.Sum(c => c.Amount)
                };

                Items.Add(item);
            }

            var maxKwh = Items.Any() ? Items.Max(i => i.TotalConsumption) : 0m;
            var maxAmount = Items.Any() ? Items.Max(i => i.TotalAmount) : 0m;

            if (maxKwh <= 0)
            {
                maxKwh = 1;
            }

            if (maxAmount <= 0)
            {
                maxAmount = 1;
            }

            _maxKwh = maxKwh;
            _maxAmount = maxAmount;

            foreach (var item in Items)
            {
                item.KwhBarRatio = maxKwh <= 0 ? 0 : (double)(item.TotalConsumption / maxKwh);
                item.AmountBarRatio = maxAmount <= 0 ? 0 : (double)(item.TotalAmount / maxAmount);
            }
        }

        [RelayCommand]
        private void PrintGroup()
        {
            var item = SelectedItem;
            if (item == null)
            {
                _ui.ShowMessage("In hóa đơn nhóm", "Hãy chọn một nhóm / đơn vị ở bảng bên phải trước.");
                return;
            }

            var customers = GetCustomersForGroup(item)
                .OrderBy(c => c.SequenceNumber)
                .ToList();

            if (customers.Count == 0)
            {
                _ui.ShowMessage("In hóa đơn nhóm", "Nhóm được chọn hiện không có dữ liệu khách hàng.");
                return;
            }

            try
            {
                var groupName = string.IsNullOrWhiteSpace(item.GroupName) ? "(Không có nhóm)" : item.GroupName.Trim();
                var confirm = _ui.Confirm(
                    "In hóa đơn nhóm",
                    $"Nhóm: {groupName}\nSố khách: {customers.Count}\n\nXuất 1 hóa đơn gộp cho nhóm này?");

                if (!confirm)
                {
                    return;
                }

                var safeGroupName = MakeSafeFileName(groupName);
                var outputPath = _ui.ShowSaveExcelFileDialog(
                    $"Hoa don - {safeGroupName}.xlsx",
                    title: "In hóa đơn nhóm");

                if (string.IsNullOrWhiteSpace(outputPath))
                {
                    return;
                }

                var templatePath = _ui.GetLegacySummaryTemplatePath();
                var sheetCount = LegacyGroupInvoiceExportService.ExportGroupInvoice(
                    templatePath,
                    outputPath,
                    groupName,
                    customers,
                    _periodLabel,
                    _issuerName);

                _ui.ShowMessage(
                    "In hóa đơn nhóm",
                    $"Đã tạo {sheetCount} sheet hóa đơn nhóm ({customers.Count} khách) cho '{groupName}' tại:\n{outputPath}");
            }
            catch (WarningException warning)
            {
                _ui.ShowMessage("In hóa đơn nhóm", warning.Message);
            }
            catch (Exception ex)
            {
                _ui.ShowMessage("Lỗi in hóa đơn nhóm", ex.Message);
            }
        }

        [RelayCommand]
        private void PrintAllGroups()
        {
            var groupedItems = Items
                .OrderBy(i => i.GroupName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (groupedItems.Count == 0)
            {
                _ui.ShowMessage("In hóa đơn theo nhóm", "Không có nhóm nào để xuất.");
                return;
            }

            var confirm = _ui.Confirm(
                "In hóa đơn theo nhóm",
                $"Sẽ xuất {groupedItems.Count} hóa đơn (mỗi nhóm 1 hóa đơn gộp). Tiếp tục?");
            if (!confirm)
            {
                return;
            }

            var folderPath = _ui.ShowFolderPickerDialog("Chọn thư mục để lưu hóa đơn theo nhóm");
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return;
            }

            try
            {
                Directory.CreateDirectory(folderPath);
                var templatePath = _ui.GetLegacySummaryTemplatePath();
                var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                var successCount = 0;
                var failedGroups = new List<string>();

                foreach (var group in groupedItems)
                {
                    var groupName = string.IsNullOrWhiteSpace(group.GroupName) ? "(Không có nhóm)" : group.GroupName.Trim();
                    var customers = GetCustomersForGroup(group)
                        .OrderBy(c => c.SequenceNumber)
                        .ToList();

                    if (customers.Count == 0)
                    {
                        continue;
                    }

                    try
                    {
                        var safeGroupName = MakeSafeFileName(groupName);
                        var outputPath = BuildUniqueFilePath(
                            folderPath,
                            $"Hoa don - {safeGroupName}.xlsx",
                            usedPaths);

                        LegacyGroupInvoiceExportService.ExportGroupInvoice(
                            templatePath,
                            outputPath,
                            groupName,
                            customers,
                            _periodLabel,
                            _issuerName);

                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failedGroups.Add($"{groupName}: {ex.Message}");
                    }
                }

                var message =
                    $"Đã xuất {successCount}/{groupedItems.Count} hóa đơn theo nhóm tại:\n{folderPath}";

                if (failedGroups.Count > 0)
                {
                    var preview = string.Join("\n", failedGroups.Take(8));
                    if (failedGroups.Count > 8)
                    {
                        preview += $"\n... ({failedGroups.Count - 8} nhóm nữa)";
                    }

                    message += $"\n\nKhông xuất được {failedGroups.Count} nhóm:\n{preview}";
                }

                _ui.ShowMessage("In hóa đơn theo nhóm", message);

                if (successCount > 0)
                {
                    var openFolder = _ui.Confirm("In hóa đơn theo nhóm", "Mở thư mục kết quả?");
                    if (openFolder)
                    {
                        _ui.OpenWithDefaultApp(folderPath);
                    }
                }
            }
            catch (WarningException warning)
            {
                _ui.ShowMessage("In hóa đơn theo nhóm", warning.Message);
            }
            catch (Exception ex)
            {
                _ui.ShowMessage("Lỗi in hóa đơn theo nhóm", ex.Message);
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

        private static string BuildUniqueFilePath(string folderPath, string fileName, ISet<string> usedPaths)
        {
            var baseName = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);

            for (var index = 1; index <= 10000; index++)
            {
                var candidateName = index == 1
                    ? $"{baseName}{extension}"
                    : $"{baseName} ({index}){extension}";
                var candidatePath = Path.Combine(folderPath, candidateName);

                if (File.Exists(candidatePath))
                {
                    continue;
                }

                if (usedPaths.Add(candidatePath))
                {
                    return candidatePath;
                }
            }

            return Path.Combine(folderPath, $"{baseName}-{Guid.NewGuid():N}{extension}");
        }

        public IEnumerable<Customer> GetCustomersForGroup(ReportItem? item)
        {
            if (item == null)
            {
                return Enumerable.Empty<Customer>();
            }

            var target = item.GroupName ?? string.Empty;

            return _sourceCustomers.Where(c =>
            {
                var key = string.IsNullOrWhiteSpace(c.GroupName) ? "(Không có nhóm)" : c.GroupName;
                return string.Equals(key, target, StringComparison.CurrentCultureIgnoreCase);
            });
        }

        private static string FormatTick(decimal value)
        {
            if (value <= 0)
            {
                return "0";
            }

            var rounded = decimal.Round(value, 0, MidpointRounding.AwayFromZero);
            return string.Format(CultureInfo.CurrentCulture, "{0:N0}", rounded);
        }
    }
}
