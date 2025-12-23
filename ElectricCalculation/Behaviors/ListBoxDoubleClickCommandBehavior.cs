using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ElectricCalculation.Behaviors
{
    public static class ListBoxDoubleClickCommandBehavior
    {
        public static readonly DependencyProperty CommandProperty =
            DependencyProperty.RegisterAttached(
                "Command",
                typeof(ICommand),
                typeof(ListBoxDoubleClickCommandBehavior),
                new PropertyMetadata(null, OnCommandChanged));

        public static ICommand? GetCommand(DependencyObject obj)
        {
            return (ICommand?)obj.GetValue(CommandProperty);
        }

        public static void SetCommand(DependencyObject obj, ICommand? value)
        {
            obj.SetValue(CommandProperty, value);
        }

        private static void OnCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not ListBox listBox)
            {
                return;
            }

            listBox.PreviewMouseDoubleClick -= OnPreviewMouseDoubleClick;

            if (e.NewValue is ICommand)
            {
                listBox.PreviewMouseDoubleClick += OnPreviewMouseDoubleClick;
            }
        }

        private static void OnPreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not ListBox listBox)
            {
                return;
            }

            var command = GetCommand(listBox);
            if (command == null)
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

            if (command.CanExecute(container.DataContext))
            {
                command.Execute(container.DataContext);
                e.Handled = true;
            }
        }
    }
}
