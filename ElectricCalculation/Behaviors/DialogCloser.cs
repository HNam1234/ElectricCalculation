using System;
using System.Windows;

namespace ElectricCalculation.Behaviors
{
    public static class DialogCloser
    {
        public static readonly DependencyProperty DialogResultProperty =
            DependencyProperty.RegisterAttached(
                "DialogResult",
                typeof(bool?),
                typeof(DialogCloser),
                new PropertyMetadata(null, OnDialogResultChanged));

        public static void SetDialogResult(Window target, bool? value)
        {
            target.SetValue(DialogResultProperty, value);
        }

        public static bool? GetDialogResult(Window target)
        {
            return (bool?)target.GetValue(DialogResultProperty);
        }

        private static void OnDialogResultChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not Window window)
            {
                return;
            }

            if (e.NewValue is not bool dialogResult)
            {
                return;
            }

            try
            {
                window.DialogResult = dialogResult;
            }
            catch (InvalidOperationException)
            {
                // Not shown as a dialog; fall back to closing.
            }

            window.Close();
        }
    }
}

