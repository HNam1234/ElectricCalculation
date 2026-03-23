using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridSelectRowOnClickBehavior
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(DataGridSelectRowOnClickBehavior),
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
                grid.PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
                return;
            }

            grid.PreviewMouseLeftButtonDown -= OnPreviewMouseLeftButtonDown;
        }

        private static void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            if (e.OriginalSource is not DependencyObject source)
            {
                return;
            }

            // Only handle clicks on checkbox/toggle buttons inside the grid.
            if (FindAncestor<ToggleButton>(source) == null)
            {
                return;
            }

            var row = FindAncestor<DataGridRow>(source);
            if (row == null)
            {
                return;
            }

            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                row.IsSelected = !row.IsSelected;
                row.Focus();
                return;
            }

            // If the row is already selected (especially in multi-selection), keep selection.
            if (row.IsSelected)
            {
                row.Focus();
                return;
            }

            grid.SelectedItems.Clear();
            row.IsSelected = true;
            grid.SelectedItem = row.Item;
            row.Focus();
        }

        private static T? FindAncestor<T>(DependencyObject node) where T : DependencyObject
        {
            while (node != null)
            {
                if (node is T typed)
                {
                    return typed;
                }

                node = VisualTreeHelper.GetParent(node);
            }

            return null;
        }
    }
}

