using System;
using System.Collections.Generic;

namespace ElectricCalculation.Services
{
    public interface IUndoableAction
    {
        string Name { get; }
        void Undo();
        void Redo();
    }

    public sealed class DelegateUndoableAction : IUndoableAction
    {
        private readonly Action _undo;
        private readonly Action _redo;

        public DelegateUndoableAction(string name, Action undo, Action redo)
        {
            Name = name ?? string.Empty;
            _undo = undo ?? throw new ArgumentNullException(nameof(undo));
            _redo = redo ?? throw new ArgumentNullException(nameof(redo));
        }

        public string Name { get; }

        public void Undo() => _undo();
        public void Redo() => _redo();
    }

    public sealed class CompositeUndoableAction : IUndoableAction
    {
        private readonly IReadOnlyList<IUndoableAction> _actions;

        public CompositeUndoableAction(string name, IReadOnlyList<IUndoableAction> actions)
        {
            Name = name ?? string.Empty;
            _actions = actions ?? Array.Empty<IUndoableAction>();
        }

        public string Name { get; }

        public void Undo()
        {
            for (var i = _actions.Count - 1; i >= 0; i--)
            {
                _actions[i].Undo();
            }
        }

        public void Redo()
        {
            foreach (var a in _actions)
            {
                a.Redo();
            }
        }
    }

    public sealed class UndoRedoManager
    {
        private readonly Stack<IUndoableAction> _undo = new();
        private readonly Stack<IUndoableAction> _redo = new();

        public event EventHandler? StateChanged;

        public bool CanUndo => _undo.Count > 0;
        public bool CanRedo => _redo.Count > 0;

        public string? UndoTitle => _undo.Count > 0 ? _undo.Peek().Name : null;
        public string? RedoTitle => _redo.Count > 0 ? _redo.Peek().Name : null;

        public void Clear()
        {
            _undo.Clear();
            _redo.Clear();
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        public void PushDone(IUndoableAction action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            _undo.Push(action);
            _redo.Clear();
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Execute(IUndoableAction action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            action.Redo();
            PushDone(action);
        }

        public void Undo()
        {
            if (_undo.Count == 0)
            {
                return;
            }

            var action = _undo.Pop();
            action.Undo();
            _redo.Push(action);
            StateChanged?.Invoke(this, EventArgs.Empty);
        }

        public void Redo()
        {
            if (_redo.Count == 0)
            {
                return;
            }

            var action = _redo.Pop();
            action.Redo();
            _undo.Push(action);
            StateChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}

