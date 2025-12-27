using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace ElectricCalculation.Behaviors
{
    public static class FileDropCommandBehavior
    {
        public static readonly DependencyProperty CommandProperty =
            DependencyProperty.RegisterAttached(
                "Command",
                typeof(ICommand),
                typeof(FileDropCommandBehavior),
                new PropertyMetadata(null, OnCommandChanged));

        public static readonly DependencyProperty FileExtensionsProperty =
            DependencyProperty.RegisterAttached(
                "FileExtensions",
                typeof(string),
                typeof(FileDropCommandBehavior),
                new PropertyMetadata(null));

        public static ICommand? GetCommand(DependencyObject obj)
        {
            return (ICommand?)obj.GetValue(CommandProperty);
        }

        public static void SetCommand(DependencyObject obj, ICommand? value)
        {
            obj.SetValue(CommandProperty, value);
        }

        public static string? GetFileExtensions(DependencyObject obj)
        {
            return (string?)obj.GetValue(FileExtensionsProperty);
        }

        public static void SetFileExtensions(DependencyObject obj, string? value)
        {
            obj.SetValue(FileExtensionsProperty, value);
        }

        private static void OnCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not UIElement element)
            {
                return;
            }

            element.PreviewDragOver -= OnPreviewDragOver;
            element.Drop -= OnDrop;

            if (e.NewValue is ICommand)
            {
                element.PreviewDragOver += OnPreviewDragOver;
                element.Drop += OnDrop;
            }
        }

        private static void OnPreviewDragOver(object sender, DragEventArgs e)
        {
            if (!TryGetFirstFilePath(e.Data, out var filePath))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            if (sender is DependencyObject obj && !IsAllowedByExtension(obj, filePath))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
                return;
            }

            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private static void OnDrop(object sender, DragEventArgs e)
        {
            if (sender is not DependencyObject obj)
            {
                return;
            }

            var command = GetCommand(obj);
            if (command == null)
            {
                return;
            }

            if (!TryGetFirstFilePath(e.Data, out var filePath))
            {
                return;
            }

            if (!IsAllowedByExtension(obj, filePath))
            {
                return;
            }

            if (command.CanExecute(filePath))
            {
                command.Execute(filePath);
                e.Handled = true;
            }
        }

        private static bool TryGetFirstFilePath(IDataObject data, out string filePath)
        {
            filePath = string.Empty;
            if (!data.GetDataPresent(DataFormats.FileDrop))
            {
                return false;
            }

            var files = data.GetData(DataFormats.FileDrop) as string[];
            var picked = files?.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
            if (string.IsNullOrWhiteSpace(picked) || !File.Exists(picked))
            {
                return false;
            }

            filePath = picked;
            return true;
        }

        private static bool IsAllowedByExtension(DependencyObject obj, string filePath)
        {
            var filter = (GetFileExtensions(obj) ?? string.Empty).Trim();
            if (filter.Length == 0)
            {
                return true;
            }

            var extensions = filter
                .Split(new[] { ';', ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.StartsWith(".", StringComparison.Ordinal) ? x : "." + x)
                .Select(x => x.ToLowerInvariant())
                .ToArray();

            if (extensions.Length == 0)
            {
                return true;
            }

            var actual = Path.GetExtension(filePath)?.ToLowerInvariant() ?? string.Empty;
            return extensions.Contains(actual);
        }
    }
}
