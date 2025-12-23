using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElectricCalculation.Services;

namespace ElectricCalculation.ViewModels
{
    public enum StartupRequestType
    {
        NewSession,
        ImportExcel,
        OpenDataFile,
        OpenSnapshot,
        OpenLatestSnapshot,
        OpenSnapshotFolder,
        OpenSnapshotPath
    }

    public sealed record StartupRequest(StartupRequestType Type, string? Path = null);

    public sealed record SnapshotItem(string Path, string PeriodLabel, string? SnapshotName, DateTime SavedAt)
    {
        public string DisplayTitle
        {
            get
            {
                var period = string.IsNullOrWhiteSpace(PeriodLabel) ? "(Không rõ kỳ tính)" : PeriodLabel;
                return string.IsNullOrWhiteSpace(SnapshotName) ? period : $"{period} - {SnapshotName}";
            }
        }
        public string DisplaySubtitle => $"{SavedAt:dd/MM/yyyy HH:mm:ss}";
    }

    public partial class StartupViewModel : ObservableObject
    {
        private readonly UiService _ui;

        public ObservableCollection<SnapshotItem> RecentSnapshots { get; } = new();

        [ObservableProperty]
        private SnapshotItem? selectedSnapshot;

        public event Action<StartupRequest>? Request;

        public StartupViewModel()
        {
            _ui = new UiService();
            RefreshSnapshots();
        }

        [RelayCommand]
        private void RefreshSnapshots()
        {
            try
            {
                RecentSnapshots.Clear();
                foreach (var item in SaveGameService.ListSnapshots(maxCount: 50))
                {
                    RecentSnapshots.Add(new SnapshotItem(item.Path, item.PeriodLabel, item.SnapshotName, item.SavedAt));
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
            Request?.Invoke(new StartupRequest(StartupRequestType.NewSession));
        }

        [RelayCommand]
        private void OpenDataFile()
        {
            var filePath = _ui.ShowOpenDataFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            Request?.Invoke(new StartupRequest(StartupRequestType.OpenDataFile, filePath));
        }

        [RelayCommand]
        private void ImportExcel()
        {
            var filePath = _ui.ShowOpenExcelFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            Request?.Invoke(new StartupRequest(StartupRequestType.ImportExcel, filePath));
        }

        [RelayCommand]
        private void OpenSnapshot()
        {
            var filePath = _ui.ShowOpenSnapshotFileDialog();
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            Request?.Invoke(new StartupRequest(StartupRequestType.OpenSnapshot, filePath));
        }

        [RelayCommand]
        private void ContinueLatestSnapshot()
        {
            try
            {
                var latest = SaveGameService.GetLatestSnapshotPath();
                if (string.IsNullOrWhiteSpace(latest))
                {
                    throw new WarningException("Chưa có snapshot nào để tiếp tục.");
                }

                Request?.Invoke(new StartupRequest(StartupRequestType.OpenLatestSnapshot, latest));
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
                    throw new WarningException("Chọn 1 snapshot trong danh sách trước.");
                }

                Request?.Invoke(new StartupRequest(StartupRequestType.OpenSnapshotPath, SelectedSnapshot.Path));
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
                    throw new WarningException("Ch ¯?n 1 snapshot trong danh sA­ch tr’ø ¯>c.");
                }

                var snapshot = SelectedSnapshot;
                var ok = _ui.Confirm(
                    "Xóa bộ dữ liệu",
                    $"Bạn có chắc muốn xóa bộ dữ liệu này không?\n\n{snapshot.DisplayTitle}\n{snapshot.Path}");

                if (!ok)
                {
                    return;
                }

                if (!SaveGameService.TryDeleteSnapshot(snapshot.Path, out var error))
                {
                    throw new WarningException(error ?? "Không thể xóa bộ dữ liệu.");
                }

                RecentSnapshots.Remove(snapshot);
                if (ReferenceEquals(SelectedSnapshot, snapshot))
                {
                    SelectedSnapshot = null;
                }
            }
            catch (Exception ex)
            {
                HandleRequestError(ex);
            }
        }

        [RelayCommand]
        private void OpenSnapshotFolder()
        {
            Request?.Invoke(new StartupRequest(StartupRequestType.OpenSnapshotFolder));
        }

        public void HandleRequestError(Exception ex)
        {
            Debug.WriteLine(ex);

            if (ex is WarningException warning)
            {
                _ui.ShowMessage("Thông báo", warning.Message);
                return;
            }

            _ui.ShowMessage("Lỗi", ex.Message);
        }
    }
}
