using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ElectricCalculation.Services;
using ElectricCalculation.ViewModels;

namespace ElectricCalculation.Views
{
    public partial class StartupWindow : Window
    {
        public StartupWindow()
        {
            InitializeComponent();

            if (DataContext is StartupViewModel vm)
            {
                vm.Request += HandleRequest;
            }
        }

        private void HandleRequest(StartupRequest request)
        {
            if (DataContext is not StartupViewModel vm)
            {
                return;
            }

            try
            {
                switch (request.Type)
                {
                    case StartupRequestType.OpenSnapshotFolder:
                        OpenSnapshotFolderOnly();
                        return;
                    case StartupRequestType.NewSession:
                        OpenEditorWindow(init: null);
                        return;
                    case StartupRequestType.ImportExcel:
                        OpenEditorWindow(editor =>
                        {
                            if (!string.IsNullOrWhiteSpace(request.Path))
                            {
                                editor.ImportFromExcelFile(request.Path);
                            }
                        });
                        return;
                    case StartupRequestType.OpenDataFile:
                        OpenEditorWindow(editor =>
                        {
                            if (!string.IsNullOrWhiteSpace(request.Path))
                            {
                                editor.LoadDataFile(request.Path, setCurrentDataFilePath: true);
                            }
                        });
                        return;
                    case StartupRequestType.OpenSnapshot:
                    case StartupRequestType.OpenLatestSnapshot:
                    case StartupRequestType.OpenSnapshotPath:
                        OpenEditorWindow(editor =>
                        {
                            if (!string.IsNullOrWhiteSpace(request.Path))
                            {
                                editor.LoadSnapshotFile(request.Path);
                            }
                        });
                        return;
                    default:
                        throw new WarningException("Tác vụ không được hỗ trợ.");
                }
            }
            catch (Exception ex)
            {
                vm.HandleRequestError(ex);
            }
        }

        private void OpenSnapshotFolderOnly()
        {
            try
            {
                var folder = SaveGameService.GetSaveRootDirectory();
                Directory.CreateDirectory(folder);
                Process.Start(new ProcessStartInfo(folder) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                if (DataContext is StartupViewModel vm)
                {
                    vm.HandleRequestError(ex);
                }
            }
        }

        private void OpenEditorWindow(Action<MainWindowViewModel>? init)
        {
            var editorWindow = new MainWindow();
            if (editorWindow.DataContext is not MainWindowViewModel editorVm)
            {
                throw new InvalidOperationException("MainWindow DataContext is not MainWindowViewModel.");
            }

            init?.Invoke(editorVm);

            Application.Current.MainWindow = editorWindow;
            editorWindow.Show();
            Close();
        }

        private void RecentSnapshots_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is not StartupViewModel vm)
            {
                return;
            }

            var command = vm.OpenSelectedSnapshotCommand;
            if (command.CanExecute(null))
            {
                command.Execute(null);
            }
        }

        private void RecentSnapshots_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is not ListBox listBox)
            {
                return;
            }

            var element = e.OriginalSource as DependencyObject;
            while (element != null && element is not ListBoxItem)
            {
                element = VisualTreeHelper.GetParent(element);
            }

            if (element is ListBoxItem item)
            {
                listBox.SelectedItem = item.DataContext;
                item.Focus();
            }
        }
    }
}
