using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ElectricCalculation.Behaviors
{
    public static class ListBoxRightClickSelectBehavior
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(ListBoxRightClickSelectBehavior),
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
            if (d is not ListBox listBox)
            {
                return;
            }

            if (e.NewValue is true)
            {
                listBox.PreviewMouseRightButtonDown += OnPreviewMouseRightButtonDown;
                return;
            }

            listBox.PreviewMouseRightButtonDown -= OnPreviewMouseRightButtonDown;
        }

        private static void OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not ListBox listBox)
            {
                return;
            }

            var source = e.OriginalSource as DependencyObject;
            var container = ItemsControl.ContainerFromElement(listBox, source) as ListBoxItem;
            if (container == null)
            {
                return;
            }

            listBox.SelectedItem = container.DataContext;
            container.Focus();
        }
    }
}
