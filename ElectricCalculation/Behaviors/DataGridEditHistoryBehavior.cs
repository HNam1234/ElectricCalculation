using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using ElectricCalculation.Models;
using ElectricCalculation.ViewModels;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridEditHistoryBehavior
    {
        public static readonly DependencyProperty TrackCellEditCommandProperty =
            DependencyProperty.RegisterAttached(
                "TrackCellEditCommand",
                typeof(ICommand),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        public static readonly DependencyProperty PasteFromClipboardCommandProperty =
            DependencyProperty.RegisterAttached(
                "PasteFromClipboardCommand",
                typeof(ICommand),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        public static readonly DependencyProperty FillDownCommandProperty =
            DependencyProperty.RegisterAttached(
                "FillDownCommand",
                typeof(ICommand),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        public static readonly DependencyProperty DeleteSelectedRowsCommandProperty =
            DependencyProperty.RegisterAttached(
                "DeleteSelectedRowsCommand",
                typeof(ICommand),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        public static readonly DependencyProperty DuplicateRowCommandProperty =
            DependencyProperty.RegisterAttached(
                "DuplicateRowCommand",
                typeof(ICommand),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        private static readonly DependencyProperty StateProperty =
            DependencyProperty.RegisterAttached(
                "State",
                typeof(State),
                typeof(DataGridEditHistoryBehavior),
                new PropertyMetadata(null));

        public static ICommand? GetTrackCellEditCommand(DependencyObject obj) =>
            (ICommand?)obj.GetValue(TrackCellEditCommandProperty);

        public static void SetTrackCellEditCommand(DependencyObject obj, ICommand? value) =>
            obj.SetValue(TrackCellEditCommandProperty, value);

        public static ICommand? GetPasteFromClipboardCommand(DependencyObject obj) =>
            (ICommand?)obj.GetValue(PasteFromClipboardCommandProperty);

        public static void SetPasteFromClipboardCommand(DependencyObject obj, ICommand? value) =>
            obj.SetValue(PasteFromClipboardCommandProperty, value);

        public static ICommand? GetFillDownCommand(DependencyObject obj) =>
            (ICommand?)obj.GetValue(FillDownCommandProperty);

        public static void SetFillDownCommand(DependencyObject obj, ICommand? value) =>
            obj.SetValue(FillDownCommandProperty, value);

        public static ICommand? GetDeleteSelectedRowsCommand(DependencyObject obj) =>
            (ICommand?)obj.GetValue(DeleteSelectedRowsCommandProperty);

        public static void SetDeleteSelectedRowsCommand(DependencyObject obj, ICommand? value) =>
            obj.SetValue(DeleteSelectedRowsCommandProperty, value);

        public static ICommand? GetDuplicateRowCommand(DependencyObject obj) =>
            (ICommand?)obj.GetValue(DuplicateRowCommandProperty);

        public static void SetDuplicateRowCommand(DependencyObject obj, ICommand? value) =>
            obj.SetValue(DuplicateRowCommandProperty, value);

        private static State GetOrCreateState(DataGrid grid)
        {
            var state = (State?)grid.GetValue(StateProperty);
            if (state != null)
            {
                return state;
            }

            state = new State(grid);
            grid.SetValue(StateProperty, state);
            return state;
        }

        private static void OnAnyPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not DataGrid grid)
            {
                return;
            }

            var hasAny =
                GetTrackCellEditCommand(grid) != null ||
                GetPasteFromClipboardCommand(grid) != null ||
                GetFillDownCommand(grid) != null ||
                GetDeleteSelectedRowsCommand(grid) != null ||
                GetDuplicateRowCommand(grid) != null;

            var state = (State?)grid.GetValue(StateProperty);
            if (!hasAny)
            {
                state?.Detach();
                grid.ClearValue(StateProperty);
                return;
            }

            GetOrCreateState(grid).Attach();
        }

        private sealed class State
        {
            private readonly DataGrid _grid;
            private bool _attached;
            private readonly Dictionary<(Customer, string), object?> _oldValues = new();

            public State(DataGrid grid)
            {
                _grid = grid;
            }

            public void Attach()
            {
                if (_attached)
                {
                    return;
                }

                _grid.BeginningEdit += Grid_BeginningEdit;
                _grid.CellEditEnding += Grid_CellEditEnding;
                _grid.PreviewKeyDown += Grid_PreviewKeyDown;
                _attached = true;
            }

            public void Detach()
            {
                if (!_attached)
                {
                    return;
                }

                _grid.BeginningEdit -= Grid_BeginningEdit;
                _grid.CellEditEnding -= Grid_CellEditEnding;
                _grid.PreviewKeyDown -= Grid_PreviewKeyDown;
                _oldValues.Clear();
                _attached = false;
            }

            private void Grid_BeginningEdit(object? sender, DataGridBeginningEditEventArgs e)
            {
                if (e.Row?.Item is not Customer customer)
                {
                    return;
                }

                var propertyName = GetPropertyName(e.Column);
                if (string.IsNullOrWhiteSpace(propertyName))
                {
                    return;
                }

                var oldValue = GetPropertyValue(customer, propertyName);
                _oldValues[(customer, propertyName)] = oldValue;
            }

            private void Grid_CellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
            {
                if (e.EditAction != DataGridEditAction.Commit)
                {
                    return;
                }

                if (e.Row?.Item is not Customer customer)
                {
                    return;
                }

                var command = GetTrackCellEditCommand(_grid);
                if (command == null)
                {
                    return;
                }

                var propertyName = GetPropertyName(e.Column);
                if (string.IsNullOrWhiteSpace(propertyName))
                {
                    return;
                }

                if (!_oldValues.TryGetValue((customer, propertyName), out var oldValue))
                {
                    oldValue = GetPropertyValue(customer, propertyName);
                }

                _grid.Dispatcher.BeginInvoke(new Action(() =>
                {
                    var newValue = GetPropertyValue(customer, propertyName);
                    if (Equals(oldValue, newValue))
                    {
                        return;
                    }

                    var change = new CellEditChange(customer, propertyName, oldValue, newValue);
                    if (command.CanExecute(change))
                    {
                        command.Execute(change);
                    }
                }));
            }

            private void Grid_PreviewKeyDown(object? sender, KeyEventArgs e)
            {
                if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.V)
                {
                    if (TryPasteFromClipboard())
                    {
                        e.Handled = true;
                    }

                    return;
                }

                if (Keyboard.Modifiers == ModifierKeys.Control && e.Key == Key.D)
                {
                    if (TryFillDown())
                    {
                        e.Handled = true;
                    }

                    return;
                }

                if (Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift) && e.Key == Key.D)
                {
                    if (TryDuplicateRow())
                    {
                        e.Handled = true;
                    }

                    return;
                }

                if (Keyboard.Modifiers == ModifierKeys.None && e.Key == Key.Delete)
                {
                    if (TryDeleteSelectedRows())
                    {
                        e.Handled = true;
                    }
                }
            }

            private bool TryDeleteSelectedRows()
            {
                var command = GetDeleteSelectedRowsCommand(_grid);
                if (command == null)
                {
                    return false;
                }

                var selected = _grid.SelectedItems?.OfType<Customer>().ToList() ?? new List<Customer>();
                if (selected.Count == 0)
                {
                    return false;
                }

                if (command.CanExecute(selected))
                {
                    command.Execute(selected);
                    return true;
                }

                return false;
            }

            private bool TryDuplicateRow()
            {
                var command = GetDuplicateRowCommand(_grid);
                if (command == null)
                {
                    return false;
                }

                if (_grid.SelectedItem is not Customer customer)
                {
                    return false;
                }

                if (command.CanExecute(customer))
                {
                    command.Execute(customer);
                    return true;
                }

                return false;
            }

            private bool TryFillDown()
            {
                var command = GetFillDownCommand(_grid);
                if (command == null)
                {
                    return false;
                }

                var propertyName = GetPropertyName(_grid.CurrentColumn);
                if (string.IsNullOrWhiteSpace(propertyName))
                {
                    return false;
                }

                var selected = _grid.SelectedItems?.OfType<Customer>().ToList() ?? new List<Customer>();
                if (selected.Count < 2)
                {
                    return false;
                }

                selected = selected
                    .OrderBy(c => _grid.Items.IndexOf(c))
                    .ToList();

                var request = new FillDownRequest(selected, propertyName);
                if (command.CanExecute(request))
                {
                    command.Execute(request);
                    return true;
                }

                return false;
            }

            private bool TryPasteFromClipboard()
            {
                var command = GetPasteFromClipboardCommand(_grid);
                if (command == null)
                {
                    return false;
                }

                var text = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                var rows = SplitClipboardRows(text);
                if (rows.Count == 0)
                {
                    return false;
                }

                var colCount = rows.Max(r => r.Length);
                if (colCount <= 0)
                {
                    return false;
                }

                var currentItem = _grid.CurrentItem;
                var startRowIndex = currentItem != null ? _grid.Items.IndexOf(currentItem) : 0;
                if (startRowIndex < 0)
                {
                    startRowIndex = 0;
                }

                var columns = _grid.Columns
                    .OrderBy(c => c.DisplayIndex)
                    .ToList();

                var startColumnIndex = _grid.CurrentColumn != null
                    ? columns.FindIndex(c => ReferenceEquals(c, _grid.CurrentColumn))
                    : 0;

                if (startColumnIndex < 0)
                {
                    startColumnIndex = 0;
                }

                var propertyNames = new List<string>();
                for (var c = 0; c < colCount; c++)
                {
                    var col = startColumnIndex + c;
                    if (col >= columns.Count)
                    {
                        break;
                    }

                    var prop = GetPropertyName(columns[col]);
                    if (string.IsNullOrWhiteSpace(prop))
                    {
                        break;
                    }

                    propertyNames.Add(prop);
                }

                if (propertyNames.Count == 0)
                {
                    return false;
                }

                var targetRows = new List<Customer>();
                for (var r = 0; r < rows.Count; r++)
                {
                    var rowIndex = startRowIndex + r;
                    if (rowIndex >= _grid.Items.Count)
                    {
                        break;
                    }

                    if (_grid.Items[rowIndex] is Customer customer)
                    {
                        targetRows.Add(customer);
                    }
                }

                if (targetRows.Count == 0)
                {
                    return false;
                }

                var request = new ClipboardPasteRequest(targetRows, propertyNames, text);
                if (command.CanExecute(request))
                {
                    command.Execute(request);
                    return true;
                }

                return false;
            }

            private static List<string[]> SplitClipboardRows(string text)
            {
                var normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
                var lines = normalized.Split('\n');
                var result = new List<string[]>();
                foreach (var line in lines)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }

                    result.Add(line.Split('\t'));
                }

                return result;
            }

            private static string? GetPropertyName(DataGridColumn? column)
            {
                if (column is not DataGridBoundColumn bound)
                {
                    return null;
                }

                if (bound.Binding is not Binding binding)
                {
                    return null;
                }

                return binding.Path?.Path;
            }

            private static object? GetPropertyValue(Customer customer, string propertyName)
            {
                try
                {
                    var prop = customer.GetType().GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance);
                    return prop?.GetValue(customer);
                }
                catch
                {
                    return null;
                }
            }
        }
    }
}
