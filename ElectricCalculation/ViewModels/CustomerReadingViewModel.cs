using CommunityToolkit.Mvvm.ComponentModel;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public partial class CustomerReadingViewModel : ObservableObject
    {
        public Customer Customer { get; }

        public MeterReading MeterReading { get; }

        public CustomerReadingViewModel(Customer customer, MeterReading meterReading)
        {
            Customer = customer;
            MeterReading = meterReading;

            name = customer.Name;
            groupName = customer.GroupName;
            category = customer.Category;
            address = customer.Address;
            phone = customer.Phone;
            meterNumber = customer.MeterNumber;

            previousIndex = meterReading.PreviousIndex;
            currentIndex = meterReading.CurrentIndex;
            multiplier = meterReading.Multiplier;
            subsidizedKwh = meterReading.SubsidizedKwh;
            unitPrice = meterReading.UnitPrice;
        }

        [ObservableProperty]
        private string name = string.Empty;

        partial void OnNameChanged(string value)
        {
            Customer.Name = value;
        }

        [ObservableProperty]
        private string groupName = string.Empty;

        partial void OnGroupNameChanged(string value)
        {
            Customer.GroupName = value;
        }

        [ObservableProperty]
        private string category = string.Empty;

        partial void OnCategoryChanged(string value)
        {
            Customer.Category = value;
        }

        [ObservableProperty]
        private string address = string.Empty;

        partial void OnAddressChanged(string value)
        {
            Customer.Address = value;
        }

        [ObservableProperty]
        private string phone = string.Empty;

        partial void OnPhoneChanged(string value)
        {
            Customer.Phone = value;
        }

        [ObservableProperty]
        private string meterNumber = string.Empty;

        partial void OnMeterNumberChanged(string value)
        {
            Customer.MeterNumber = value;
        }

        [ObservableProperty]
        private decimal previousIndex;

        partial void OnPreviousIndexChanged(decimal value)
        {
            MeterReading.PreviousIndex = value;
            OnPropertyChanged(nameof(Consumption));
            OnPropertyChanged(nameof(ChargeableKwh));
            OnPropertyChanged(nameof(Amount));
        }

        [ObservableProperty]
        private decimal currentIndex;

        partial void OnCurrentIndexChanged(decimal value)
        {
            MeterReading.CurrentIndex = value;
            OnPropertyChanged(nameof(Consumption));
            OnPropertyChanged(nameof(ChargeableKwh));
            OnPropertyChanged(nameof(Amount));
        }

        [ObservableProperty]
        private decimal multiplier = 1;

        partial void OnMultiplierChanged(decimal value)
        {
            MeterReading.Multiplier = value;
            OnPropertyChanged(nameof(Consumption));
            OnPropertyChanged(nameof(ChargeableKwh));
            OnPropertyChanged(nameof(Amount));
        }

        [ObservableProperty]
        private decimal subsidizedKwh;

        partial void OnSubsidizedKwhChanged(decimal value)
        {
            MeterReading.SubsidizedKwh = value;
            OnPropertyChanged(nameof(ChargeableKwh));
            OnPropertyChanged(nameof(Amount));
        }

        [ObservableProperty]
        private decimal unitPrice;

        partial void OnUnitPriceChanged(decimal value)
        {
            MeterReading.UnitPrice = value;
            OnPropertyChanged(nameof(Amount));
        }

        public decimal Consumption
        {
            get
            {
                var delta = MeterReading.CurrentIndex - MeterReading.PreviousIndex;
                if (delta < 0)
                {
                    delta = 0;
                }

                if (MeterReading.Multiplier <= 0)
                {
                    return 0;
                }

                return delta * MeterReading.Multiplier;
            }
        }

        public decimal ChargeableKwh
        {
            get
            {
                var result = Consumption - MeterReading.SubsidizedKwh;
                return result > 0 ? result : 0;
            }
        }

        public decimal Amount => ChargeableKwh * MeterReading.UnitPrice;
    }
}
