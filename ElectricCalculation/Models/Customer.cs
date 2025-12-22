using System;
using CommunityToolkit.Mvvm.ComponentModel;

namespace ElectricCalculation.Models
{
    public partial class Customer : ObservableObject
    {
        [ObservableProperty]
        private int sequenceNumber;

        [ObservableProperty]
        private string name = string.Empty;

        [ObservableProperty]
        private string groupName = string.Empty;

        [ObservableProperty]
        private string category = string.Empty;

        [ObservableProperty]
        private string address = string.Empty;

        [ObservableProperty]
        private string representativeName = string.Empty;

        [ObservableProperty]
        private string householdPhone = string.Empty;

        [ObservableProperty]
        private string phone = string.Empty;

        [ObservableProperty]
        private string buildingName = string.Empty;

        [ObservableProperty]
        private string meterNumber = string.Empty;

        [ObservableProperty]
        private string substation = string.Empty;

        [ObservableProperty]
        private string page = string.Empty;

        [ObservableProperty]
        private string performedBy = string.Empty;

        // Meter location (optional).
        [ObservableProperty]
        private string location = string.Empty;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Consumption))]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal previousIndex;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Consumption))]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal currentIndex;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Consumption))]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal multiplier = 1;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal subsidizedKwh;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal unitPrice;

        public decimal Consumption
        {
            get
            {
                var delta = CurrentIndex - PreviousIndex;
                if (delta <= 0 || Multiplier <= 0)
                {
                    return 0;
                }

                return delta * Multiplier;
            }
        }

        public decimal ChargeableKwh => Math.Max(0, Consumption - SubsidizedKwh);

        public decimal Amount => ChargeableKwh * UnitPrice;
    }
}
