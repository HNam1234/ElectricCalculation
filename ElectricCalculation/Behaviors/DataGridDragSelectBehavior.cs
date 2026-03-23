using System;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace ElectricCalculation.Behaviors
{
    public static class DataGridDragSelectBehavior
    {
        private sealed class DragState
        {
            public bool IsTracking { get; set; }
            public bool IsDragging { get; set; }
            public int StartIndex { get; set; }
            public int LastIndex { get; set; }
            public Point StartPoint { get; set; }
        }

        public static readonly DependencyProperty IsEnabledProperty =
            DependencyProperty.RegisterAttached(
                "IsEnabled",
                typeof(bool),
                typeof(DataGridDragSelectBehavior),
                new PropertyMetadata(false, OnIsEnabledChanged));

        private static readonly DependencyProperty DragStateProperty =
            DependencyProperty.RegisterAttached(
                "DragState",
                typeof(DragState),
                typeof(DataGridDragSelectBehavior),
                new PropertyMetadata(null));

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
                grid.PreviewMouseMove += OnPreviewMouseMove;
                grid.PreviewMouseLeftButtonUp += OnPreviewMouseLeftButtonUp;
                return;
            }

            grid.PreviewMouseLeftButtonDown -= OnPreviewMouseLeftButtonDown;
            grid.PreviewMouseMove -= OnPreviewMouseMove;
            grid.PreviewMouseLeftButtonUp -= OnPreviewMouseLeftButtonUp;
        }

        private static void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            if (Keyboard.Modifiers != ModifierKeys.None)
            {
                return;
            }

            var source = e.OriginalSource as DependencyObject;
            if (source == null)
            {
                return;
            }

            var row = FindAncestor<DataGridRow>(source);
            if (row == null)
            {
                return;
            }

            var index = grid.ItemContainerGenerator.IndexFromContainer(row);
            if (index < 0)
            {
                return;
            }

            var state = new DragState
            {
                IsTracking = true,
                IsDragging = false,
                StartIndex = index,
                LastIndex = index,
                StartPoint = e.GetPosition(grid)
            };

            grid.SetValue(DragStateProperty, state);
        }

        private static void OnPreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            if (grid.GetValue(DragStateProperty) is not DragState state || !state.IsTracking)
            {
                return;
            }

            if (!state.IsDragging)
            {
                var current = e.GetPosition(grid);
                var deltaX = current.X - state.StartPoint.X;
                var deltaY = current.Y - state.StartPoint.Y;

                // Wait for a small movement threshold so clicking a checkbox still toggles normally.
                const double threshold = 6;
                if (Math.Abs(deltaX) < threshold && Math.Abs(deltaY) < threshold)
                {
                    return;
                }

                state.IsDragging = true;
                grid.CaptureMouse();
                SelectRange(grid, state.StartIndex, state.StartIndex);
            }

            var position = e.GetPosition(grid);
            if (grid.InputHitTest(position) is not DependencyObject hit)
            {
                return;
            }

            var row = FindAncestor<DataGridRow>(hit);
            if (row == null)
            {
                return;
            }

            var index = grid.ItemContainerGenerator.IndexFromContainer(row);
            if (index < 0 || index == state.LastIndex)
            {
                return;
            }

            state.LastIndex = index;
            SelectRange(grid, state.StartIndex, index);
        }

        private static void OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGrid grid)
            {
                return;
            }

            if (grid.GetValue(DragStateProperty) is DragState state)
            {
                state.IsTracking = false;
                state.IsDragging = false;
            }

            grid.ReleaseMouseCapture();
            grid.ClearValue(DragStateProperty);
        }

        private static void SelectRange(DataGrid grid, int startIndex, int endIndex)
        {
            if (grid == null)
            {
                return;
            }

            var count = grid.Items.Count;
            if (count <= 0)
            {
                return;
            }

            startIndex = Math.Clamp(startIndex, 0, count - 1);
            endIndex = Math.Clamp(endIndex, 0, count - 1);

            var min = Math.Min(startIndex, endIndex);
            var max = Math.Max(startIndex, endIndex);

            grid.SelectedItems.Clear();
            for (var i = min; i <= max; i++)
            {
                grid.SelectedItems.Add(grid.Items[i]);
            }

            grid.SelectedItem = grid.Items[endIndex];
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
