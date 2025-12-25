using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace ElectricCalculation.Behaviors
{
    public static class ShortcutFocusBehavior
    {
        public static readonly DependencyProperty KeyProperty =
            DependencyProperty.RegisterAttached(
                "Key",
                typeof(Key),
                typeof(ShortcutFocusBehavior),
                new PropertyMetadata(Key.None, OnAnyPropertyChanged));

        public static readonly DependencyProperty ModifiersProperty =
            DependencyProperty.RegisterAttached(
                "Modifiers",
                typeof(ModifierKeys),
                typeof(ShortcutFocusBehavior),
                new PropertyMetadata(ModifierKeys.None, OnAnyPropertyChanged));

        public static readonly DependencyProperty TargetProperty =
            DependencyProperty.RegisterAttached(
                "Target",
                typeof(IInputElement),
                typeof(ShortcutFocusBehavior),
                new PropertyMetadata(null, OnAnyPropertyChanged));

        public static readonly DependencyProperty SelectAllTextProperty =
            DependencyProperty.RegisterAttached(
                "SelectAllText",
                typeof(bool),
                typeof(ShortcutFocusBehavior),
                new PropertyMetadata(true, OnAnyPropertyChanged));

        private static readonly DependencyProperty StateProperty =
            DependencyProperty.RegisterAttached(
                "State",
                typeof(State),
                typeof(ShortcutFocusBehavior),
                new PropertyMetadata(null));

        public static Key GetKey(DependencyObject obj) =>
            (Key)obj.GetValue(KeyProperty);

        public static void SetKey(DependencyObject obj, Key value) =>
            obj.SetValue(KeyProperty, value);

        public static ModifierKeys GetModifiers(DependencyObject obj) =>
            (ModifierKeys)obj.GetValue(ModifiersProperty);

        public static void SetModifiers(DependencyObject obj, ModifierKeys value) =>
            obj.SetValue(ModifiersProperty, value);

        public static IInputElement? GetTarget(DependencyObject obj) =>
            (IInputElement?)obj.GetValue(TargetProperty);

        public static void SetTarget(DependencyObject obj, IInputElement? value) =>
            obj.SetValue(TargetProperty, value);

        public static bool GetSelectAllText(DependencyObject obj) =>
            (bool)obj.GetValue(SelectAllTextProperty);

        public static void SetSelectAllText(DependencyObject obj, bool value) =>
            obj.SetValue(SelectAllTextProperty, value);

        private static void OnAnyPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not UIElement root)
            {
                return;
            }

            var hasShortcut = GetKey(root) != Key.None && GetTarget(root) != null;
            var state = (State?)root.GetValue(StateProperty);

            if (!hasShortcut)
            {
                state?.Detach();
                root.ClearValue(StateProperty);
                return;
            }

            state ??= new State(root);
            root.SetValue(StateProperty, state);
            state.Attach();
        }

        private sealed class State
        {
            private readonly UIElement _root;
            private bool _attached;

            public State(UIElement root)
            {
                _root = root;
            }

            public void Attach()
            {
                if (_attached)
                {
                    return;
                }

                _root.PreviewKeyDown += Root_PreviewKeyDown;
                _attached = true;
            }

            public void Detach()
            {
                if (!_attached)
                {
                    return;
                }

                _root.PreviewKeyDown -= Root_PreviewKeyDown;
                _attached = false;
            }

            private void Root_PreviewKeyDown(object sender, KeyEventArgs e)
            {
                var key = e.Key == Key.System ? e.SystemKey : e.Key;
                if (key != GetKey(_root) || Keyboard.Modifiers != GetModifiers(_root))
                {
                    return;
                }

                var target = GetTarget(_root);
                if (target == null)
                {
                    return;
                }

                if (target is FrameworkElement element)
                {
                    element.Focus();

                    if (GetSelectAllText(_root) && element is TextBox textBox)
                    {
                        _root.Dispatcher.BeginInvoke(new Action(textBox.SelectAll), DispatcherPriority.Input);
                    }

                    e.Handled = true;
                }
            }
        }
    }
}

