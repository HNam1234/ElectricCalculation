using System;
using System.Diagnostics;
using System.Windows;
using ElectricCalculation.Services;
using ElectricCalculation.ViewModels;

namespace ElectricCalculation.Views
{
    public partial class MainWindow : Window
    {
        private bool _skipClosePrompt;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BackToHome_Click(object sender, RoutedEventArgs e)
        {
            if (!PromptSnapshotIfNeeded())
            {
                return;
            }

            _skipClosePrompt = true;
            var home = new StartupWindow();
            Application.Current.MainWindow = home;
            home.Show();
            Close();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_skipClosePrompt)
            {
                return;
            }

            if (!PromptSnapshotIfNeeded())
            {
                e.Cancel = true;
            }
        }

        private bool PromptSnapshotIfNeeded()
        {
            try
            {
                if (DataContext is not MainWindowViewModel vm)
                {
                    return true;
                }

                if (vm.Customers.Count == 0 || !vm.IsDirty)
                {
                    return true;
                }

                var ui = new UiService();
                var canOverwrite = !string.IsNullOrWhiteSpace(vm.LoadedSnapshotPath);
                var (result, action, snapshotName) = ui.ShowSaveSnapshotPrompt(
                    vm.PeriodLabel,
                    vm.Customers.Count,
                    defaultSnapshotName: "Chỉnh sửa",
                    canOverwrite: canOverwrite);

                if (result == null)
                {
                    return false; // cancel
                }

                if (action == SaveSnapshotPromptAction.DontSave)
                {
                    return true;
                }

                if (action == SaveSnapshotPromptAction.Overwrite && !string.IsNullOrWhiteSpace(vm.LoadedSnapshotPath))
                {
                    ProjectFileService.Save(vm.LoadedSnapshotPath, vm.PeriodLabel, vm.Customers);
                    vm.IsDirty = false;
                    return true;
                }

                SaveGameService.SaveSnapshot(vm.PeriodLabel, vm.Customers, snapshotName);
                vm.IsDirty = false;

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return true;
            }
        }
    }
}

