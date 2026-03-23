using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
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

        [ObservableProperty]
        private bool? dialogResult;

        [ObservableProperty]
        private string validationMessage = string.Empty;

        public int TotalCount => Items.Count;

        public int SelectedCount => Items.Count(item => item.IsSelected);

        public GroupInvoiceSelectionViewModel(string groupName, IEnumerable<Customer> customers)
        {
            GroupName = (groupName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(GroupName))
            {
                GroupName = "(Không có nhóm)";
            }

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
        }

        public IReadOnlyList<Customer> GetSelectedCustomers()
        {
            return Items
                .Where(item => item.IsSelected)
                .Select(item => item.Customer)
                .ToList();
        }

        [RelayCommand]
        private void SelectAll()
        {
            foreach (var item in Items)
            {
                item.IsSelected = true;
            }

            OnPropertyChanged(nameof(SelectedCount));
        }

        [RelayCommand]
        private void ClearAll()
        {
            foreach (var item in Items)
            {
                item.IsSelected = false;
            }

            OnPropertyChanged(nameof(SelectedCount));
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
            }
        }
    }
}
