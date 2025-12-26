using System;
using System.Windows;
using System.Windows.Controls;
using ElectricCalculation.ViewModels;

namespace ElectricCalculation.Views
{
    public partial class ImportWizardWindow : Window
    {
        public ImportWizardWindow()
        {
            InitializeComponent();
        }

        private void MappingGrid_AutoGeneratingColumn(object? sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (DataContext is not ImportWizardViewModel viewModel)
            {
                return;
            }

            if (string.Equals(e.PropertyName, "Row", StringComparison.OrdinalIgnoreCase))
            {
                e.Column.Header = "Dòng";
                return;
            }

            if (!viewModel.TryGetColumnMapping(e.PropertyName, out var mapping) || mapping == null)
            {
                return;
            }

            e.Column.Header = mapping;
            e.Column.HeaderTemplate = (DataTemplate)Resources["ColumnMappingHeaderTemplate"];
        }
    }
}
