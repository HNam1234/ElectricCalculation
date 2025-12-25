using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridEnterToNextRowBehavior
    {
        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(DataGridEnterToNextRowBehavior),
                new PropertyMetadata(false, OnAnyPropertyChanged));

        public static readonly DependencyProperty TargetPropertyNameProperty =
            DependencyProperty.RegisterAttached(
                "TargetPropertyName",
                typeof(string),
                typeof(DataGridEnterToNextRowBehavior),
                new PropertyMetadata(string.Empty, OnAnyPropertyChanged));

        public static readonly DependencyProperty SelectAllOnEditProperty =
            DependencyProperty.RegisterAttached(
                "SelectAllOnEdit",
                typeof(bool),
                typeof(DataGridEnterToNextRowBehavior),
                new PropertyMetadata(true, OnAnyPropertyChanged));

        private static readonly DependencyProperty StateProperty =
            DependencyProperty.RegisterAttached(
                "State",
                typeof(State),
                typeof(DataGridEnterToNextRowBehavior),
                new PropertyMetadata(null));

        public static bool GetIsEnabled(DependencyObject obj) =>
            (bool)obj.GetValue(IsEnabledProperty);

        public static void SetIsEnabled(DependencyObject obj, bool value) =>
            obj.SetValue(IsEnabledProperty, value);

        public static string GetTargetPropertyName(DependencyObject obj) =>
            (string)obj.GetValue(TargetPropertyNameProperty);

        public static void SetTargetPropertyName(DependencyObject obj, string value) =>
            obj.SetValue(TargetPropertyNameProperty, value);

        public static bool GetSelectAllOnEdit(DependencyObject obj) =>
            (bool)obj.GetValue(SelectAllOnEditProperty);

        public static void SetSelectAllOnEdit(DependencyObject obj, bool value) =>
            obj.SetValue(SelectAllOnEditProperty, value);

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

                _grid.PreviewKeyDown += Grid_PreviewKeyDown;
                _grid.PreparingCellForEdit += Grid_PreparingCellForEdit;
                _attached = true;
            }

            public void Detach()
            {
                if (!_attached)
                {
                    return;
                }

                _grid.PreviewKeyDown -= Grid_PreviewKeyDown;
                _grid.PreparingCellForEdit -= Grid_PreparingCellForEdit;
                _attached = false;
            }

            private void Grid_PreparingCellForEdit(object? sender, DataGridPreparingCellForEditEventArgs e)
            {
                if (!GetSelectAllOnEdit(_grid))
                {
                    return;
                }

                if (e.EditingElement is not TextBox textBox)
                {
                    return;
                }

                var targetColumn = FindTargetColumn();
                if (targetColumn == null || !ReferenceEquals(targetColumn, e.Column))
                {
                    return;
                }

                _grid.Dispatcher.BeginInvoke(new Action(textBox.SelectAll), DispatcherPriority.Input);
            }

            private void Grid_PreviewKeyDown(object? sender, KeyEventArgs e)
            {
                if (Keyboard.Modifiers != ModifierKeys.None || e.Key != Key.Enter)
                {
                    return;
                }

                if (_grid.IsReadOnly || _grid.CurrentItem == null)
                {
                    return;
                }

                var committedCell = _grid.CommitEdit(DataGridEditingUnit.Cell, true);
                if (!committedCell)
                {
                    return;
                }

                _grid.CommitEdit(DataGridEditingUnit.Row, true);

                var currentRowIndex = _grid.Items.IndexOf(_grid.CurrentItem);
                if (currentRowIndex < 0)
                {
                    return;
                }

                var nextRowIndex = currentRowIndex + 1;
                if (nextRowIndex >= _grid.Items.Count)
                {
                    e.Handled = true;
                    return;
                }

                var nextItem = _grid.Items[nextRowIndex];
                var targetColumn = FindTargetColumn() ?? _grid.CurrentColumn;
                if (targetColumn == null)
                {
                    return;
                }

                _grid.SelectedItem = nextItem;
                _grid.ScrollIntoView(nextItem);
                _grid.CurrentCell = new DataGridCellInfo(nextItem, targetColumn);

                _grid.Dispatcher.BeginInvoke(new Action(() =>
                {
                    _grid.BeginEdit();
                }), DispatcherPriority.Background);

                e.Handled = true;
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
        }
    }
}
