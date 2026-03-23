using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public sealed record GroupInvoiceSelectionResult(
        IReadOnlyList<Customer> SelectedCustomers,
        string PeriodLabel,
        string IssuerName,
        string IssuePlace,
        DateTime IssueDate,
        string RecipientName,
        string ConsumptionAddress,
        string RepresentativeName,
        string HouseholdPhone,
        string RepresentativePhone);

    public sealed class GroupInvoicePreviewRow
    {
        public int Stt { get; }
        public decimal? CurrentIndex { get; }
        public decimal PreviousIndex { get; }
        public decimal Multiplier { get; }
        public decimal Consumption { get; }
        public decimal SubsidizedKwh { get; }
        public decimal UnitPrice { get; }
        public decimal Amount { get; }
        public string Name { get; }
        public string Address { get; }

        public GroupInvoicePreviewRow(int stt, Customer customer, string displayName)
        {
            if (customer == null)
            {
                throw new ArgumentNullException(nameof(customer));
            }

            Stt = stt;
            CurrentIndex = customer.CurrentIndex;
            PreviousIndex = customer.PreviousIndex;
            Multiplier = customer.Multiplier <= 0 ? 1 : customer.Multiplier;
            Consumption = customer.Consumption;
            SubsidizedKwh = customer.SubsidizedKwh;
            UnitPrice = customer.UnitPrice;
            Amount = customer.Amount;
            Name = displayName ?? string.Empty;
            Address = customer.Address ?? string.Empty;
        }
    }

    public partial class GroupInvoiceSelectionItem : ObservableObject
    {
        public Customer Customer { get; }

        [ObservableProperty]
        private bool isSelected = true;

        public int SequenceNumber => Customer.SequenceNumber;
        public string Name => Customer.Name;
        public string MeterNumber => Customer.MeterNumber;
        public string Location => Customer.Location;
        public string Address => Customer.Address;
        public string HouseholdPhone => Customer.HouseholdPhone;
        public string Phone => Customer.Phone;

        public GroupInvoiceSelectionItem(Customer customer)
        {
            Customer = customer ?? throw new ArgumentNullException(nameof(customer));
        }
    }

    public partial class GroupInvoiceSelectionViewModel : ObservableObject
    {
        public string GroupName { get; }

        public ObservableCollection<GroupInvoiceSelectionItem> Items { get; }

        public ObservableCollection<GroupInvoicePreviewRow> PreviewRows { get; } = new();

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string validationMessage = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        [NotifyPropertyChangedFor(nameof(PreviewFooterText))]
        private string periodLabel = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewFooterText))]
        private string issuerName = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewFooterText))]
        private string issuePlace = "Hà Nội";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewFooterText))]
        private DateTime? issueDate = DateTime.Today;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private bool useAutoHeaderFields = true;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private string recipientName = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private string consumptionAddress = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private string representativeName = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private string householdPhone = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PreviewHeaderText))]
        private string representativePhone = string.Empty;

        public int TotalCount => Items.Count;

        public int SelectedCount => Items.Count(item => item.IsSelected);

        private IReadOnlyList<Customer> selectedCustomersSnapshot = Array.Empty<Customer>();
        private string resolvedAddress = string.Empty;
        private string resolvedRepresentativeDisplay = string.Empty;
        private string resolvedHouseholdPhone = string.Empty;
        private string resolvedRepresentativePhone = string.Empty;
        private bool suppressSelectionRefresh;

        public string PreviewHeaderText
        {
            get
            {
                var lines = new List<string>();

                var period = (PeriodLabel ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(period))
                {
                    lines.Add($"Kỳ: {period}");
                }

                var effectiveRecipient = UseAutoHeaderFields ? GroupName : (RecipientName ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(effectiveRecipient))
                {
                    effectiveRecipient = GroupName;
                }

                lines.Add($"Kính gửi: {effectiveRecipient}");

                var effectiveAddress = UseAutoHeaderFields ? resolvedAddress : (ConsumptionAddress ?? string.Empty).Trim();
                var effectiveRepresentative = UseAutoHeaderFields ? resolvedRepresentativeDisplay : (RepresentativeName ?? string.Empty).Trim();
                var effectiveHouseholdPhone = UseAutoHeaderFields ? resolvedHouseholdPhone : (HouseholdPhone ?? string.Empty).Trim();
                var effectiveRepresentativePhone = UseAutoHeaderFields ? resolvedRepresentativePhone : (RepresentativePhone ?? string.Empty).Trim();

                if (!string.IsNullOrWhiteSpace(effectiveAddress))
                {
                    lines.Add($"Địa chỉ hộ tiêu thụ: {EnsureTrailingPeriod(effectiveAddress)}");
                }

                if (!string.IsNullOrWhiteSpace(effectiveRepresentative))
                {
                    lines.Add($"Đại diện: {EnsureTrailingPeriod(effectiveRepresentative)}");
                }

                if (!string.IsNullOrWhiteSpace(effectiveHouseholdPhone))
                {
                    lines.Add($"Điện thoại: {effectiveHouseholdPhone}");
                }

                if (!string.IsNullOrWhiteSpace(effectiveRepresentativePhone))
                {
                    lines.Add($"Điện thoại: {effectiveRepresentativePhone}");
                }

                return string.Join(Environment.NewLine, lines);
            }
        }

        public decimal PreviewTotalAmount { get; private set; }

        public string PreviewAmountText { get; private set; } = string.Empty;

        public string PreviewFooterText
        {
            get
            {
                var lines = new List<string>();

                var place = (IssuePlace ?? string.Empty).Trim();
                var date = IssueDate ?? DateTime.Today;
                var dateText = string.IsNullOrWhiteSpace(place)
                    ? $"Ngày {date.Day} tháng {date.Month} năm {date.Year}"
                    : $"{place}, ngày {date.Day} tháng {date.Month} năm {date.Year}";

                lines.Add(dateText);

                var issuer = (IssuerName ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(issuer))
                {
                    lines.Add($"Người lập: {issuer}");
                }

                if (PreviewTotalAmount > 0)
                {
                    lines.Add($"Tổng tiền: {PreviewTotalAmount:N0} VNĐ");
                }

                if (!string.IsNullOrWhiteSpace(PreviewAmountText))
                {
                    lines.Add($"Bằng chữ: {PreviewAmountText}");
                }

                return string.Join(Environment.NewLine, lines);
            }
        }

        public GroupInvoiceSelectionViewModel(
            string groupName,
            IEnumerable<Customer> customers,
            string? periodLabel = null,
            string? issuerName = null,
            string? issuePlace = null,
            DateTime? issueDate = null)
        {
            GroupName = (groupName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(GroupName))
            {
                GroupName = "(Không có nhóm)";
            }

            PeriodLabel = periodLabel?.Trim() ?? string.Empty;
            IssuerName = issuerName?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(issuePlace))
            {
                IssuePlace = issuePlace.Trim();
            }

            IssueDate = issueDate ?? DateTime.Today;

            var list = (customers ?? Array.Empty<Customer>())
                .Where(c => c != null)
                .OrderBy(c => c.SequenceNumber > 0 ? c.SequenceNumber : int.MaxValue)
                .ThenBy(c => c.Name)
                .Select(c => new GroupInvoiceSelectionItem(c))
                .ToList();

            Items = new ObservableCollection<GroupInvoiceSelectionItem>(list);

            foreach (var item in Items)
            {
                item.PropertyChanged += OnItemPropertyChanged;
            }

            RefreshPreview();
        }

        public IReadOnlyList<Customer> GetSelectedCustomers()
        {
            return Items
                .Where(item => item.IsSelected)
                .Select(item => item.Customer)
                .ToList();
        }

        public GroupInvoiceSelectionResult GetResult()
        {
            var customers = GetSelectedCustomers()
                .Where(c => c != null)
                .OrderBy(c => c.SequenceNumber > 0 ? c.SequenceNumber : int.MaxValue)
                .ThenBy(c => c.Name)
                .ToList();

            var effectiveRecipient = UseAutoHeaderFields ? GroupName : (RecipientName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(effectiveRecipient))
            {
                effectiveRecipient = GroupName;
            }

            var effectiveAddress = UseAutoHeaderFields ? resolvedAddress : (ConsumptionAddress ?? string.Empty).Trim();
            var effectiveRepresentative = UseAutoHeaderFields ? resolvedRepresentativeDisplay : (RepresentativeName ?? string.Empty).Trim();
            var effectiveHouseholdPhone = UseAutoHeaderFields ? resolvedHouseholdPhone : (HouseholdPhone ?? string.Empty).Trim();
            var effectiveRepresentativePhone = UseAutoHeaderFields ? resolvedRepresentativePhone : (RepresentativePhone ?? string.Empty).Trim();

            return new GroupInvoiceSelectionResult(
                customers,
                (PeriodLabel ?? string.Empty).Trim(),
                (IssuerName ?? string.Empty).Trim(),
                (IssuePlace ?? string.Empty).Trim(),
                IssueDate ?? DateTime.Today,
                effectiveRecipient,
                effectiveAddress,
                effectiveRepresentative,
                effectiveHouseholdPhone,
                effectiveRepresentativePhone);
        }

        [RelayCommand]
        private void SelectAll()
        {
            ApplySelectionToItems(Items, isSelected: true);
        }

        [RelayCommand]
        private void ClearAll()
        {
            ApplySelectionToItems(Items, isSelected: false);
        }

        [RelayCommand]
        private void SelectHighlighted(IList? selectedItems)
        {
            var targets = selectedItems?.OfType<GroupInvoiceSelectionItem>().ToList();
            if (targets == null || targets.Count == 0)
            {
                return;
            }

            ApplySelectionToItems(targets, isSelected: true);
        }

        [RelayCommand]
        private void ClearHighlighted(IList? selectedItems)
        {
            var targets = selectedItems?.OfType<GroupInvoiceSelectionItem>().ToList();
            if (targets == null || targets.Count == 0)
            {
                return;
            }

            ApplySelectionToItems(targets, isSelected: false);
        }

        [RelayCommand]
        private void ToggleHighlighted(IList? selectedItems)
        {
            var targets = selectedItems?.OfType<GroupInvoiceSelectionItem>().ToList();
            if (targets == null || targets.Count == 0)
            {
                return;
            }

            ApplySelectionToItems(targets, isSelected: null);
        }

        [RelayCommand]
        private void Ok()
        {
            if (SelectedCount <= 0)
            {
                ValidationMessage = "Bạn chưa chọn hộ nào để in.";
                return;
            }

            ValidationMessage = string.Empty;
            DialogResult = true;
        }

        [RelayCommand]
        private void Cancel()
        {
            DialogResult = false;
        }

        private void OnItemPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (string.Equals(e.PropertyName, nameof(GroupInvoiceSelectionItem.IsSelected), StringComparison.Ordinal))
            {
                OnPropertyChanged(nameof(SelectedCount));
                if (SelectedCount > 0 && !string.IsNullOrWhiteSpace(ValidationMessage))
                {
                    ValidationMessage = string.Empty;
                }

                if (suppressSelectionRefresh)
                {
                    return;
                }

                RefreshPreview();
            }
        }

        private void ApplySelectionToItems(IEnumerable<GroupInvoiceSelectionItem> targets, bool? isSelected)
        {
            if (targets == null)
            {
                return;
            }

            suppressSelectionRefresh = true;
            try
            {
                foreach (var item in targets)
                {
                    if (isSelected == null)
                    {
                        item.IsSelected = !item.IsSelected;
                        continue;
                    }

                    item.IsSelected = isSelected.Value;
                }
            }
            finally
            {
                suppressSelectionRefresh = false;
            }

            OnPropertyChanged(nameof(SelectedCount));
            if (SelectedCount > 0 && !string.IsNullOrWhiteSpace(ValidationMessage))
            {
                ValidationMessage = string.Empty;
            }

            RefreshPreview();
        }

        private void RefreshPreview()
        {
            selectedCustomersSnapshot = Items
                .Where(item => item.IsSelected)
                .Select(item => item.Customer)
                .Where(c => c != null)
                .OrderBy(c => c.SequenceNumber > 0 ? c.SequenceNumber : int.MaxValue)
                .ThenBy(c => c.Name)
                .ToList();

            resolvedAddress = ResolveGroupHeaderAddress(GroupName, selectedCustomersSnapshot);

            var sharedRepresentative = GetSharedNonEmptyValue(selectedCustomersSnapshot, c => c.RepresentativeName) ?? string.Empty;
            var sharedHouseholdPhone = ResolveBestPhone(selectedCustomersSnapshot, c => c.HouseholdPhone);
            var sharedRepresentativePhone = ResolveBestPhone(selectedCustomersSnapshot, c => c.Phone);

            var normalizedHouseholdPhone = NormalizePhoneForComparison(sharedHouseholdPhone);
            var normalizedRepresentativePhone = NormalizePhoneForComparison(sharedRepresentativePhone);

            if (!string.IsNullOrWhiteSpace(normalizedHouseholdPhone) &&
                !string.IsNullOrWhiteSpace(normalizedRepresentativePhone) &&
                string.Equals(normalizedHouseholdPhone, normalizedRepresentativePhone, StringComparison.OrdinalIgnoreCase))
            {
                sharedRepresentativePhone = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(sharedHouseholdPhone) && !string.IsNullOrWhiteSpace(sharedRepresentativePhone))
            {
                sharedHouseholdPhone = sharedRepresentativePhone;
                sharedRepresentativePhone = string.Empty;
            }

            resolvedHouseholdPhone = sharedHouseholdPhone;
            resolvedRepresentativePhone = sharedRepresentativePhone;
            resolvedRepresentativeDisplay = !string.IsNullOrWhiteSpace(sharedRepresentative) ? sharedRepresentative : GroupName;

            if (UseAutoHeaderFields)
            {
                RecipientName = GroupName;
                ConsumptionAddress = resolvedAddress;
                RepresentativeName = resolvedRepresentativeDisplay;
                HouseholdPhone = resolvedHouseholdPhone;
                RepresentativePhone = resolvedRepresentativePhone;
            }

            PreviewRows.Clear();
            for (var i = 0; i < selectedCustomersSnapshot.Count; i++)
            {
                var customer = selectedCustomersSnapshot[i];
                var displayName = string.IsNullOrWhiteSpace(customer.Name) ? $"Hộ {i + 1}" : customer.Name.Trim();
                PreviewRows.Add(new GroupInvoicePreviewRow(i + 1, customer, displayName));
            }

            PreviewTotalAmount = selectedCustomersSnapshot.Sum(c => c.Amount);
            PreviewAmountText = VietnameseNumberTextService.ConvertAmountToText(PreviewTotalAmount);

            OnPropertyChanged(nameof(PreviewTotalAmount));
            OnPropertyChanged(nameof(PreviewAmountText));
            OnPropertyChanged(nameof(PreviewHeaderText));
            OnPropertyChanged(nameof(PreviewFooterText));
        }

        partial void OnUseAutoHeaderFieldsChanged(bool value)
        {
            RefreshPreview();
        }

        private static string ResolveGroupHeaderAddress(string groupName, IReadOnlyList<Customer> customers)
        {
            var shared = GetSharedNonEmptyValue(customers, c => c.Address);
            if (!string.IsNullOrWhiteSpace(shared))
            {
                return shared;
            }

            return (groupName ?? string.Empty).Trim();
        }

        private static string? GetSharedNonEmptyValue(IReadOnlyList<Customer> customers, Func<Customer, string?> selector)
        {
            string? shared = null;

            foreach (var customer in customers)
            {
                var value = selector(customer)?.Trim();
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (shared == null)
                {
                    shared = value;
                    continue;
                }

                if (!string.Equals(shared, value, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
            }

            return shared;
        }

        private static string ResolveBestPhone(IReadOnlyList<Customer> customers, Func<Customer, string?> selector)
        {
            if (customers.Count == 0)
            {
                return string.Empty;
            }

            var groups = new Dictionary<string, (int Count, string Display, int DisplayScore)>(StringComparer.OrdinalIgnoreCase);

            foreach (var customer in customers)
            {
                var raw = selector(customer)?.Trim();
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                var normalized = NormalizePhoneForComparison(raw);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                if (!groups.TryGetValue(normalized, out var existing))
                {
                    groups[normalized] = (1, raw, ScorePhoneDisplay(raw));
                    continue;
                }

                var bestDisplay = existing.Display;
                var bestScore = existing.DisplayScore;
                var score = ScorePhoneDisplay(raw);
                if (score > bestScore)
                {
                    bestDisplay = raw;
                    bestScore = score;
                }

                groups[normalized] = (existing.Count + 1, bestDisplay, bestScore);
            }

            if (groups.Count == 0)
            {
                return string.Empty;
            }

            if (groups.Count == 1)
            {
                return groups.Values.First().Display;
            }

            var best = groups
                .OrderByDescending(kvp => kvp.Value.Count)
                .ThenByDescending(kvp => kvp.Value.DisplayScore)
                .First();

            return best.Value.Display;
        }

        private static int ScorePhoneDisplay(string phone)
        {
            if (string.IsNullOrWhiteSpace(phone))
            {
                return -1;
            }

            var text = phone.Trim();
            var digitCount = text.Count(char.IsDigit);
            var score = digitCount;

            if (text.StartsWith("0", StringComparison.Ordinal))
            {
                score += 1000;
            }

            if (text.StartsWith("+", StringComparison.Ordinal))
            {
                score += 900;
            }

            if (text.StartsWith("84", StringComparison.Ordinal))
            {
                score += 800;
            }

            return score;
        }

        private static string NormalizePhoneForComparison(string phone)
        {
            var raw = phone?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(raw.Length);
            foreach (var ch in raw)
            {
                if (char.IsDigit(ch))
                {
                    builder.Append(ch);
                    continue;
                }
            }

            var digits = builder.ToString();
            if (string.IsNullOrWhiteSpace(digits))
            {
                return string.Empty;
            }

            if (digits.StartsWith("84", StringComparison.Ordinal) && digits.Length >= 9)
            {
                digits = "0" + digits.Substring(2);
            }

            digits = digits.TrimStart('0');
            return digits;
        }

        private static string EnsureTrailingPeriod(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            var trimmed = text.Trim();
            return trimmed.EndsWith(".", StringComparison.Ordinal) ? trimmed : $"{trimmed}.";
        }
    }
}
