using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Services;
using ElectricCalculation.Views;

namespace ElectricCalculation.ViewModels
{
    public sealed record SnapshotItem(
        string Path,
        string DisplayTitle,
        string DisplaySubtitle,
        bool IsPinned,
        DateTime? PinnedAtUtc,
        DateTime SavedAt)
    {
        public bool CanDelete => true;
        public bool CanTogglePin => true;
        public string PinMenuText => IsPinned ? "Unpin" : "Pin";
    }

    public partial class StartupViewModel : ObservableObject
    {
        private readonly UiService _ui;

        public ObservableCollection<SnapshotItem> RecentSnapshots { get; } = new();

        [ObservableProperty]
        private SnapshotItem? selectedSnapshot;

        public StartupViewModel()
        {
            _ui = new UiService();
            _ = SampleDataService.TrySeedJune2025SampleSnapshotOnce();
            RefreshSnapshots();
        }

        [RelayCommand]
        private void RefreshSnapshots()
        {
            try
            {
                var keepSelectedPath = SelectedSnapshot?.Path;

                var snapshots = SaveGameService.ListSnapshots(maxCount: 50);

                var pins = PinnedDatasetService.LoadPins();
                if (PinnedDatasetService.TryCleanupMissingPins(pins, snapshots.Select(s => s.Path)))
                {
                    PinnedDatasetService.SavePins(pins);
                }

                RecentSnapshots.Clear();

                var items = new List<SnapshotItem>();

                foreach (var snapshot in snapshots)
                {
                    var period = string.IsNullOrWhiteSpace(snapshot.PeriodLabel) ? "(Unknown period)" : snapshot.PeriodLabel;
                    var title = string.IsNullOrWhiteSpace(snapshot.SnapshotName) ? period : $"{period} - {snapshot.SnapshotName}";
                    var subtitle = $"{snapshot.SavedAt:dd/MM/yyyy HH:mm:ss}";

                    var isPinned = pins.TryGetValue(snapshot.Path, out var pinnedAtUtc);
                    items.Add(new SnapshotItem(
                        snapshot.Path,
                        title,
                        subtitle,
                        IsPinned: isPinned,
                        PinnedAtUtc: isPinned ? pinnedAtUtc : null,
                        SavedAt: snapshot.SavedAt));
                }

                foreach (var item in items
                    .OrderByDescending(i => i.IsPinned)
                    .ThenByDescending(i => i.PinnedAtUtc ?? DateTime.MinValue)
                    .ThenByDescending(i => i.SavedAt)
                    .ThenBy(i => i.DisplayTitle, StringComparer.OrdinalIgnoreCase))
                {
                    RecentSnapshots.Add(item);
                }

                if (!string.IsNullOrWhiteSpace(keepSelectedPath))
                {
                    SelectedSnapshot = RecentSnapshots.FirstOrDefault(i =>
                        string.Equals(i.Path, keepSelectedPath, StringComparison.OrdinalIgnoreCase));
                }
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void NewSession()
        {
            OpenEditorWindow(init: null);
        }

        [RelayCommand]
        private void ImportExcel()
        {
            var filePath = _ui.ShowOpenExcelFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            OpenEditorWindow(editor => editor.ImportFromExcelFile(filePath));
        }

        [RelayCommand]
        private void OpenSnapshot()
        {
            var filePath = _ui.ShowOpenSnapshotFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            OpenEditorWindow(editor =>
            {
                if (IsSnapshotPath(filePath))
                {
                    editor.LoadSnapshotFile(filePath);
                    return;
                }

                editor.LoadDataFile(filePath, setCurrentDataFilePath: true);
            });
        }

        [RelayCommand]
        private void ContinueLatestSnapshot()
        {
            try
            {
                // Prefer the newest item in the current list (top snapshot section),
                // so the "continue" action matches what users see on Startup.
                RefreshSnapshots();

                var latest = RecentSnapshots.FirstOrDefault(i => !i.IsPinned)?.Path
                    ?? RecentSnapshots.FirstOrDefault()?.Path;

                if (string.IsNullOrWhiteSpace(latest))
                {
                    throw new WarningException("No snapshot to continue.");
                }

                OpenEditorWindow(editor => editor.LoadSnapshotFile(latest));
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void OpenSelectedSnapshot()
        {
            try
            {
                if (SelectedSnapshot == null)
                {
                    throw new WarningException("Select a dataset first.");
                }

                var item = SelectedSnapshot;
                OpenEditorWindow(editor => editor.LoadSnapshotFile(item.Path));
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void TogglePinSelectedSnapshot()
        {
            try
            {
                if (SelectedSnapshot == null)
                {
                    throw new WarningException("Select a dataset first.");
                }

                var pins = PinnedDatasetService.LoadPins();
                if (pins.ContainsKey(SelectedSnapshot.Path))
                {
                    pins.Remove(SelectedSnapshot.Path);
                }
                else
                {
                    pins[SelectedSnapshot.Path] = DateTime.UtcNow;
                }

                PinnedDatasetService.SavePins(pins);
                RefreshSnapshots();
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void DeleteSelectedSnapshot()
        {
            try
            {
                if (SelectedSnapshot == null)
                {
                    throw new WarningException("Select a dataset first.");
                }

                var snapshot = SelectedSnapshot;
                var ok = _ui.Confirm(
                    "Delete dataset",
                    $"Delete this dataset?\n\n{snapshot.DisplayTitle}\n{snapshot.Path}");

                if (!ok)
                {
                    return;
                }

                var pins = PinnedDatasetService.LoadPins();
                if (pins.Remove(snapshot.Path))
                {
                    PinnedDatasetService.SavePins(pins);
                }

                if (!SaveGameService.TryDeleteSnapshot(snapshot.Path, out var error))
                {
                    throw new WarningException(error ?? "Failed to delete dataset.");
                }

                RefreshSnapshots();
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void OpenSnapshotFolder()
        {
            try
            {
                var folder = _ui.GetSnapshotFolderPath();
                _ui.OpenWithDefaultApp(folder);
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void OpenLogoUrl()
        {
            try
            {
                _ui.OpenWithDefaultApp("https://www.youtube.com/watch?v=xvFZjo5PgG0");
            }
            catch
            {
                // Ignore: opening a browser may fail in locked-down environments.
            }
        }

        private static bool IsSnapshotPath(string filePath)
        {
            try
            {
                var root = Path.GetFullPath(SaveGameService.GetSaveRootDirectory())
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;

                var fullPath = Path.GetFullPath(filePath);
                return fullPath.StartsWith(root, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private void OpenEditorWindow(Action<MainWindowViewModel>? init)
        {
            var editorWindow = new MainWindow();
            if (editorWindow.DataContext is not MainWindowViewModel editorVm)
            {
                throw new InvalidOperationException("MainWindow DataContext is not MainWindowViewModel.");
            }

            try
            {
                init?.Invoke(editorVm);
            }
            catch (Exception ex)
            {
                editorWindow.Close();
                HandleRequestError(ex);
                return;
            }

            Application.Current.MainWindow = editorWindow;
            editorWindow.Show();

            var hostWindow = Application.Current?.Windows
                .OfType<Window>()
                .FirstOrDefault(w => ReferenceEquals(w.DataContext, this));

            hostWindow?.Close();
        }

        public void HandleRequestError(Exception ex)
        {
            Debug.WriteLine(ex);

            if (ex is WarningException warning)
            {
                _ui.ShowMessage("Notice", warning.Message);
                return;
            }

            _ui.ShowMessage("Error", ex.Message);
        }
    }
}
