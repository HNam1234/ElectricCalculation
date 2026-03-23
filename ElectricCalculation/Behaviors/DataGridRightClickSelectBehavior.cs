using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridRightClickSelectBehavior
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(DataGridRightClickSelectBehavior),
                new PropertyMetadata(false, OnIsEnabledChanged));

        public static bool GetIsEnabled(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsEnabledProperty);
        }

        public static void SetIsEnabled(DependencyObject obj, bool value)
        {
            obj.SetValue(IsEnabledProperty, value);
        }

        private static void OnIsEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not DataGrid grid)
            {
                return;
            }

            if (e.NewValue is true)
            {
                grid.PreviewMouseRightButtonDown += OnPreviewMouseRightButtonDown;
                return;
            }

            grid.PreviewMouseRightButtonDown -= OnPreviewMouseRightButtonDown;
        }

        private static void OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            var source = e.OriginalSource as DependencyObject;
            if (source == null)
            {
                return;
            }

            var row = ItemsControl.ContainerFromElement(grid, source) as DataGridRow;
            if (row == null)
            {
                return;
            }

            if (!row.IsSelected)
            {
                grid.SelectedItem = row.Item;
            }

            row.Focus();
        }
    }
}

