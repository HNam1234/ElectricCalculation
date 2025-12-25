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
        [NotifyPropertyChangedFor(nameof(IsMissingReading))]
        [NotifyPropertyChangedFor(nameof(HasReadingError))]
        [NotifyPropertyChangedFor(nameof(HasUsageWarning))]
        [NotifyPropertyChangedFor(nameof(IsZeroUsage))]
        [NotifyPropertyChangedFor(nameof(StatusText))]
        [NotifyPropertyChangedFor(nameof(StatusTooltip))]
        private decimal previousIndex;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Consumption))]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        [NotifyPropertyChangedFor(nameof(IsMissingReading))]
        [NotifyPropertyChangedFor(nameof(HasReadingError))]
        [NotifyPropertyChangedFor(nameof(HasUsageWarning))]
        [NotifyPropertyChangedFor(nameof(IsZeroUsage))]
        [NotifyPropertyChangedFor(nameof(StatusText))]
        [NotifyPropertyChangedFor(nameof(StatusTooltip))]
        private decimal? currentIndex;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Consumption))]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        [NotifyPropertyChangedFor(nameof(HasUsageWarning))]
        [NotifyPropertyChangedFor(nameof(StatusText))]
        [NotifyPropertyChangedFor(nameof(StatusTooltip))]
        private decimal multiplier = 1;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ChargeableKwh))]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal subsidizedKwh;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(Amount))]
        private decimal unitPrice;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(HasUsageWarning))]
        [NotifyPropertyChangedFor(nameof(StatusText))]
        [NotifyPropertyChangedFor(nameof(StatusTooltip))]
        private decimal? averageConsumption3Periods;

        public bool IsMissingReading => CurrentIndex == null;

        public bool HasReadingError => CurrentIndex != null && CurrentIndex.Value < PreviousIndex;

        public bool IsZeroUsage => CurrentIndex != null && CurrentIndex.Value == PreviousIndex;

        public bool HasUsageWarning
        {
            get
            {
                if (CurrentIndex == null || HasReadingError)
                {
                    return false;
                }

                if (AverageConsumption3Periods == null || AverageConsumption3Periods <= 0)
                {
                    return false;
                }

                return Consumption > AverageConsumption3Periods.Value * 2;
            }
        }

        public string StatusText
        {
            get
            {
                if (IsMissingReading)
                {
                    return "Thiếu chỉ số";
                }

                if (HasReadingError)
                {
                    return "Lỗi: CS mới < CS cũ";
                }

                if (HasUsageWarning)
                {
                    return "Cảnh báo: Tăng cao";
                }

                if (IsZeroUsage)
                {
                    return "0 kWh";
                }

                return "OK";
            }
        }

        public string StatusTooltip
        {
            get
            {
                if (IsMissingReading)
                {
                    return "Thiếu: chưa nhập chỉ số mới";
                }

                if (HasReadingError)
                {
                    return "Lỗi: chỉ số mới < chỉ số cũ";
                }

                if (HasUsageWarning)
                {
                    if (AverageConsumption3Periods is > 0)
                    {
                        var ratio = Consumption / AverageConsumption3Periods.Value;
                        var increasePercent = (ratio - 1) * 100;
                        return $"Cảnh báo: kWh tăng {increasePercent:0}% so với TB 3 tháng";
                    }

                    return "Cảnh báo: kWh tăng cao";
                }

                if (IsZeroUsage)
                {
                    return "0 kWh: không tiêu thụ (CS mới = CS cũ)";
                }

                return "OK";
            }
        }

        public decimal Consumption
        {
            get
            {
                if (CurrentIndex == null)
                {
                    return 0;
                }

                var delta = CurrentIndex.Value - PreviousIndex;
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
