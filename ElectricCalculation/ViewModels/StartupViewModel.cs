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

                CreateNewPeriodFromDatasetCommand.NotifyCanExecuteChanged();
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

        private bool CanCreateNewPeriodFromDataset() => RecentSnapshots.Count > 0;

        [RelayCommand(CanExecute = nameof(CanCreateNewPeriodFromDataset))]
        private void CreateNewPeriodFromDataset()
        {
            try
            {
                RefreshSnapshots();

                var snapshots = SaveGameService.ListSnapshots(maxCount: 200);
                if (snapshots.Count == 0)
                {
                    throw new WarningException("Không có bộ dữ liệu để tạo tháng mới.");
                }

                var pins = PinnedDatasetService.LoadPins();

                var ordered = snapshots
                    .Select(s =>
                    {
                        var isPinned = pins.TryGetValue(s.Path, out var pinnedAtUtc);
                        return new
                        {
                            Snapshot = s,
                            IsPinned = isPinned,
                            PinnedAtUtc = isPinned ? pinnedAtUtc : (DateTime?)null
                        };
                    })
                    .OrderByDescending(x => x.IsPinned)
                    .ThenByDescending(x => x.PinnedAtUtc ?? DateTime.MinValue)
                    .ThenByDescending(x => x.Snapshot.SavedAt)
                    .ToList();

                var referenceOptions = ordered
                    .Select(x =>
                    {
                        var period = string.IsNullOrWhiteSpace(x.Snapshot.PeriodLabel)
                            ? "(Unknown period)"
                            : x.Snapshot.PeriodLabel;

                        var title = string.IsNullOrWhiteSpace(x.Snapshot.SnapshotName)
                            ? period
                            : $"{period} - {x.Snapshot.SnapshotName}";

                        var pinPrefix = x.IsPinned ? "[PIN] " : string.Empty;
                        var displayName = $"{pinPrefix}{title} ({x.Snapshot.SavedAt:dd/MM/yyyy HH:mm})";

                        return new NewPeriodViewModel.ReferenceDatasetOption(
                            PeriodLabel: period,
                            DisplayName: displayName,
                            SnapshotPath: x.Snapshot.Path,
                            IsCurrentDataset: false);
                    })
                    .ToList();

                var dialogVm = new NewPeriodViewModel(referenceOptions);

                var defaultPath = SelectedSnapshot?.Path
                    ?? RecentSnapshots.FirstOrDefault(i => !i.IsPinned)?.Path
                    ?? RecentSnapshots.FirstOrDefault()?.Path;

                if (!string.IsNullOrWhiteSpace(defaultPath))
                {
                    dialogVm.SelectedReferenceDataset = dialogVm.ReferenceDatasets.FirstOrDefault(o =>
                        string.Equals(o.SnapshotPath, defaultPath, StringComparison.OrdinalIgnoreCase))
                        ?? dialogVm.ReferenceDatasets.FirstOrDefault();
                }

                if (TryGetNextPeriod(dialogVm.SelectedReferenceDataset?.PeriodLabel, out var nextMonth, out var nextYear))
                {
                    dialogVm.Month = nextMonth;
                    dialogVm.Year = nextYear;
                }

                var vm = _ui.ShowNewPeriodDialog(dialogVm);
                if (vm == null)
                {
                    return;
                }

                var reference = vm.SelectedReferenceDataset;
                var snapshotPath = reference?.SnapshotPath;
                if (string.IsNullOrWhiteSpace(snapshotPath))
                {
                    throw new WarningException("Bộ dữ liệu tháng cũ không hợp lệ.");
                }

                var (_, customers) = ProjectFileService.Load(snapshotPath);
                customers = customers
                    .OrderBy(c => c.SequenceNumber)
                    .ToList();

                foreach (var customer in customers)
                {
                    if (vm.MoveCurrentToPrevious)
                    {
                        customer.PreviousIndex = customer.CurrentIndex ?? customer.PreviousIndex;
                    }

                    if (vm.ResetCurrentToZero)
                    {
                        customer.CurrentIndex = null;
                    }
                }

                var periodLabel = vm.PeriodLabel;

                OpenEditorWindow(editor =>
                {
                    var tempPath = Path.Combine(
                        Path.GetTempPath(),
                        $"ElectricCalculation_NewPeriod_{Guid.NewGuid():N}.json");

                    try
                    {
                        ProjectFileService.Save(tempPath, periodLabel, customers);
                        editor.LoadDataFile(tempPath, setCurrentDataFilePath: false);
                        editor.IsDirty = true;
                    }
                    finally
                    {
                        try
                        {
                            if (File.Exists(tempPath))
                            {
                                File.Delete(tempPath);
                            }
                        }
                        catch
                        {
                            // Best-effort cleanup.
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        private static bool TryGetNextPeriod(string? periodLabel, out int month, out int year)
        {
            month = 0;
            year = 0;

            if (string.IsNullOrWhiteSpace(periodLabel))
            {
                return false;
            }

            var parts = periodLabel.Split('/');
            if (parts.Length < 2)
            {
                return false;
            }

            var monthText = new string(parts[0].Where(char.IsDigit).ToArray());
            var yearText = new string(parts[1].Where(char.IsDigit).ToArray());

            if (!int.TryParse(monthText, out var m) || !int.TryParse(yearText, out var y))
            {
                return false;
            }

            if (m is < 1 or > 12 || y < 2000)
            {
                return false;
            }

            var next = new DateTime(y, m, 1).AddMonths(1);
            month = next.Month;
            year = next.Year;
            return true;
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
