using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridFastEntryLockBehavior
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(DataGridFastEntryLockBehavior),
                new PropertyMetadata(false, OnAnyPropertyChanged));

        public static readonly DependencyProperty TargetPropertyNameProperty =
            DependencyProperty.RegisterAttached(
                "TargetPropertyName",
                typeof(string),
                typeof(DataGridFastEntryLockBehavior),
                new PropertyMetadata(string.Empty, OnAnyPropertyChanged));

        private static readonly DependencyProperty StateProperty =
            DependencyProperty.RegisterAttached(
                "State",
                typeof(State),
                typeof(DataGridFastEntryLockBehavior),
                new PropertyMetadata(null));

        public static bool GetIsEnabled(DependencyObject obj) =>
            (bool)obj.GetValue(IsEnabledProperty);

        public static void SetIsEnabled(DependencyObject obj, bool value) =>
            obj.SetValue(IsEnabledProperty, value);

        public static string GetTargetPropertyName(DependencyObject obj) =>
            (string)obj.GetValue(TargetPropertyNameProperty);

        public static void SetTargetPropertyName(DependencyObject obj, string value) =>
            obj.SetValue(TargetPropertyNameProperty, value);

        private static void OnAnyPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not DataGrid grid)
            {
                return;
            }

            var enabled = GetIsEnabled(grid);
            var state = (State?)grid.GetValue(StateProperty);

            if (!enabled)
            {
                state?.Detach();
                grid.ClearValue(StateProperty);
                return;
            }

            state ??= new State(grid);
            grid.SetValue(StateProperty, state);
            state.Attach();
        }

        private sealed class State
        {
            private readonly DataGrid _grid;
            private bool _attached;
            private bool _adjusting;

            public State(DataGrid grid)
            {
                _grid = grid;
            }

            public void Attach()
            {
                if (_attached)
                {
                    EnforceTargetCell(beginEdit: false);
                    return;
                }

                _grid.Loaded += Grid_Loaded;
                _grid.CurrentCellChanged += Grid_CurrentCellChanged;
                _grid.BeginningEdit += Grid_BeginningEdit;
                _grid.PreviewKeyDown += Grid_PreviewKeyDown;
                _grid.PreviewMouseLeftButtonDown += Grid_PreviewMouseLeftButtonDown;
                _attached = true;

                EnforceTargetCell(beginEdit: false);
            }

            public void Detach()
            {
                if (!_attached)
                {
                    return;
                }

                _grid.Loaded -= Grid_Loaded;
                _grid.CurrentCellChanged -= Grid_CurrentCellChanged;
                _grid.BeginningEdit -= Grid_BeginningEdit;
                _grid.PreviewKeyDown -= Grid_PreviewKeyDown;
                _grid.PreviewMouseLeftButtonDown -= Grid_PreviewMouseLeftButtonDown;
                _attached = false;
            }

            private void Grid_Loaded(object sender, RoutedEventArgs e)
            {
                EnforceTargetCell(beginEdit: true);
            }

            private void Grid_CurrentCellChanged(object? sender, EventArgs e)
            {
                if (_adjusting)
                {
                    return;
                }

                EnforceTargetCell(beginEdit: false);
            }

            private void Grid_BeginningEdit(object? sender, DataGridBeginningEditEventArgs e)
            {
                if (_adjusting)
                {
                    return;
                }

                var targetColumn = FindTargetColumn();
                if (targetColumn == null || ReferenceEquals(e.Column, targetColumn))
                {
                    return;
                }

                e.Cancel = true;
                EnforceTargetCell(beginEdit: true);
            }

            private void Grid_PreviewKeyDown(object sender, KeyEventArgs e)
            {
                if (e.Key != Key.Tab && e.Key != Key.Left && e.Key != Key.Right)
                {
                    return;
                }

                if (e.Key is Key.Left or Key.Right && Keyboard.FocusedElement is TextBox)
                {
                    return;
                }

                e.Handled = true;
                EnforceTargetCell(beginEdit: true);
            }

            private void Grid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
            {
                if (_adjusting)
                {
                    return;
                }

                var cell = FindAncestor<DataGridCell>(e.OriginalSource as DependencyObject);
                if (cell == null)
                {
                    return;
                }

                var targetColumn = FindTargetColumn();
                if (targetColumn == null || ReferenceEquals(cell.Column, targetColumn))
                {
                    return;
                }

                _grid.Dispatcher.BeginInvoke(new Action(() => EnforceTargetCell(beginEdit: true)), DispatcherPriority.Input);
            }

            private void EnforceTargetCell(bool beginEdit)
            {
                var targetColumn = FindTargetColumn();
                if (targetColumn == null || _grid.Items.Count == 0)
                {
                    return;
                }

                var item = _grid.CurrentItem ?? _grid.SelectedItem ?? _grid.Items[0];
                if (item == null)
                {
                    return;
                }

                if (!_grid.CurrentCell.IsValid || !ReferenceEquals(_grid.CurrentCell.Column, targetColumn))
                {
                    _adjusting = true;

                    try
                    {
                        _grid.SelectedItem = item;
                        _grid.CurrentCell = new DataGridCellInfo(item, targetColumn);
                        _grid.ScrollIntoView(item, targetColumn);
                    }
                    finally
                    {
                        _adjusting = false;
                    }
                }

                if (beginEdit)
                {
                    _grid.Dispatcher.BeginInvoke(new Action(() => _grid.BeginEdit()), DispatcherPriority.Background);
                }
            }

            private DataGridColumn? FindTargetColumn()
            {
                var targetPropertyName = GetTargetPropertyName(_grid);
                if (string.IsNullOrWhiteSpace(targetPropertyName))
                {
                    return null;
                }

                return _grid.Columns
                    .OrderBy(c => c.DisplayIndex)
                    .FirstOrDefault(c => string.Equals(GetPropertyName(c), targetPropertyName, StringComparison.Ordinal));
            }

            private static string? GetPropertyName(DataGridColumn column)
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

            private static T? FindAncestor<T>(DependencyObject? start) where T : DependencyObject
            {
                var current = start;
                while (current != null)
                {
                    if (current is T typed)
                    {
                        return typed;
                    }

                    current = VisualTreeHelper.GetParent(current);
                }

                return null;
            }
        }
    }
}

