using CommunityToolkit.Mvvm.ComponentModel;
using ElectricCalculation.Models;

namespace ElectricCalculation.ViewModels
{
    public partial class InvoiceViewModel : ObservableObject
    {
        public Customer Customer { get; }

        public string PeriodLabel { get; }

        public InvoiceViewModel(Customer customer, string periodLabel)
        {
            Customer = customer;
            PeriodLabel = periodLabel;
        }
    }
}

