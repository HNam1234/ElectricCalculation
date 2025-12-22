using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Models;
using System.Windows;
using System.Windows.Controls;

namespace ElectricCalculation.ViewModels
{
    public partial class InvoiceViewModel : ObservableObject
    {
        [ObservableProperty]
        private bool? dialogResult;

        public Customer Customer { get; }

        public string PeriodLabel { get; }

        public IRelayCommand<FrameworkElement> PrintCommand { get; }

        public IRelayCommand CloseCommand { get; }

        public InvoiceViewModel(Customer customer, string periodLabel)
        {
            Customer = customer;
            PeriodLabel = periodLabel;

            PrintCommand = new RelayCommand<FrameworkElement>(PrintInvoice, canExecute: element => element != null);
            CloseCommand = new RelayCommand(() => DialogResult = false);
        }

        private void PrintInvoice(FrameworkElement? invoiceRoot)
        {
            if (invoiceRoot == null)
            {
                return;
            }

            var dialog = new PrintDialog();
            if (dialog.ShowDialog() == true)
            {
                dialog.PrintVisual(invoiceRoot, "Phiếu thu tiền điện");
            }
        }
    }
}

