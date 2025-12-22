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
        private double kwhBarHeight;

        [ObservableProperty]
        private double amountBarHeight;

        [ObservableProperty]
        private bool isSelected;
    }

    public partial class ReportViewModel : ObservableObject
    {
        private const double MaxBarHeight = 180.0;

        private readonly UiService _ui;

        private readonly decimal _maxKwh;
        private readonly decimal _maxAmount;
        private readonly string _periodLabel;
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

        public ReportViewModel(string periodLabel, IEnumerable<Customer> customers)
            : this(periodLabel, customers, new UiService())
        {
        }

        public ReportViewModel(
            string periodLabel,
            IEnumerable<Customer> customers,
            UiService ui)
        {
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));

            _periodLabel = periodLabel ?? string.Empty;
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
                item.KwhBarHeight = (double)(item.TotalConsumption / maxKwh) * MaxBarHeight;
                item.AmountBarHeight = (double)(item.TotalAmount / maxAmount) * MaxBarHeight;
            }
        }

        [RelayCommand]
        private void PrintGroup()
        {
            var item = SelectedItem;
            if (item == null)
            {
                _ui.ShowMessage("In Excel nhóm", "Hãy chọn một nhóm / đơn vị ở bảng bên phải trước.");
                return;
            }

            var customers = GetCustomersForGroup(item).ToList();
            if (customers.Count == 0)
            {
                _ui.ShowMessage("In Excel nhóm", "Nhóm được chọn hiện không có dữ liệu khách hàng.");
                return;
            }

            var safeGroupName = MakeSafeFileName(item.GroupName);
            var outputPath = _ui.ShowSaveExcelFileDialog(
                $"Tien dien - {safeGroupName}.xlsx",
                title: "In Excel nhóm");

            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return;
            }

            try
            {
                var templatePath = _ui.GetSummaryTemplatePath();
                ExcelExportService.ExportToFile(templatePath, outputPath, customers, _periodLabel);
            }
            catch (WarningException warning)
            {
                _ui.ShowMessage("In Excel nhóm", warning.Message);
            }
            catch (Exception ex)
            {
                _ui.ShowMessage("Lỗi in Excel nhóm", ex.Message);
            }
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return string.IsNullOrWhiteSpace(name) ? "Tien_dien" : name;
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
