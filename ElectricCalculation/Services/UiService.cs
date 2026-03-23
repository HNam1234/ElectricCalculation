using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using ElectricCalculation.Models;
using ElectricCalculation.ViewModels;
using ElectricCalculation.Views;
using Microsoft.Win32;

namespace ElectricCalculation.Services
{
    public sealed class UiService
    {
        private const string InvoiceTemplateFileName = "DefaultTemplate.xlsx";
        private const string PackagedInvoiceTemplateRelativePath = @"Templates\DefaultTemplate.xlsx";
        private const string PackagedSummaryTemplateRelativePath = @"SampleData\Bang_tong_hop_thu_thang_06_2025.xlsx";
        private const string LegacySummaryTemplateFileName = "Bảng tổng hợp thu tháng 6 năm 2025.xlsx";

        private static Window? GetOwner()
        {
            var app = Application.Current;
            if (app == null)
            {
                return null;
            }

            var active = app.Windows
                .OfType<Window>()
                .FirstOrDefault(w => w.IsActive);

            return active ?? app.MainWindow;
        }

        public string? ShowOpenExcelFileDialog()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string? ShowOpenDataFileDialog()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Electric Calculation data (*.json)|*.json|All files (*.*)|*.*"
            };

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string? ShowOpenSnapshotFileDialog()
        {
            var saveRoot = SaveGameService.GetSaveRootDirectory();
            Directory.CreateDirectory(saveRoot);

            var dialog = new OpenFileDialog
            {
                Filter = "Electric Calculation bộ dữ liệu (*.json)|*.json|All files (*.*)|*.*",
                InitialDirectory = saveRoot
            };

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string GetSnapshotFolderPath()
        {
            var saveRoot = SaveGameService.GetSaveRootDirectory();
            Directory.CreateDirectory(saveRoot);
            return saveRoot;
        }

        public string? ShowSaveExcelFileDialog(string defaultFileName, string? title = null)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = defaultFileName
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                dialog.Title = title;
            }

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string? ShowSavePdfFileDialog(string defaultFileName, string? title = null)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*",
                FileName = defaultFileName
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                dialog.Title = title;
            }

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string? ShowSaveDataFileDialog(string defaultFileName, string? title = null)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Electric Calculation data (*.json)|*.json|All files (*.*)|*.*",
                FileName = defaultFileName
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                dialog.Title = title;
            }

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            return result == true ? dialog.FileName : null;
        }

        public string? ShowFolderPickerDialog(string title)
        {
            var dialog = new SaveFileDialog
            {
                Title = title,
                Filter = "Thư mục|*.folder",
                FileName = "Chon_thu_muc_o_day"
            };

            var owner = GetOwner();
            var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
            if (result != true)
            {
                return null;
            }

            return Path.GetDirectoryName(dialog.FileName);
        }

        public bool Confirm(string title, string message)
        {
            var owner = GetOwner();
            var result = owner != null
                ? MessageBox.Show(owner, message, title, MessageBoxButton.YesNo, MessageBoxImage.Question)
                : MessageBox.Show(message, title, MessageBoxButton.YesNo, MessageBoxImage.Question);

            return result == MessageBoxResult.Yes;
        }

        public void ShowMessage(string title, string message)
        {
            var vm = new MessageDialogViewModel(title, message);
            var dialog = new MessageDialogWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            dialog.ShowDialog();
        }

        public IDisposable ShowBusyScope(string title, string message)
        {
            var owner = GetOwner();
            var vm = new BusyDialogViewModel(title, message);
            var dialog = new BusyDialogWindow
            {
                Owner = owner,
                DataContext = vm
            };

            if (owner != null)
            {
                owner.IsEnabled = false;
            }

            dialog.Show();

            return new BusyScope(owner, dialog);
        }

        private sealed class BusyScope : IDisposable
        {
            private readonly Window? owner;
            private readonly Window dialog;

            public BusyScope(Window? owner, Window dialog)
            {
                this.owner = owner;
                this.dialog = dialog;
            }

            public void Dispose()
            {
                void CloseOnUi()
                {
                    try
                    {
                        dialog.Close();
                    }
                    catch
                    {
                        // Ignore close errors.
                    }

                    if (owner != null)
                    {
                        owner.IsEnabled = true;
                        owner.Activate();
                    }
                }

                if (dialog.Dispatcher.CheckAccess())
                {
                    CloseOnUi();
                }
                else
                {
                    dialog.Dispatcher.Invoke(CloseOnUi);
                }
            }
        }

        public void ShowUserGuideDialog()
        {
            var owner = GetOwner();
            var sections = UserGuideSnapshotService.BuildGuideSections(owner);
            if (sections.Count == 0)
            {
                throw new WarningException("Không thể tạo ảnh hướng dẫn. Hãy thử mở lại hướng dẫn sau.");
            }

            var vm = new UserGuideViewModel(sections);
            var dialog = new UserGuideWindow
            {
                Owner = owner,
                DataContext = vm
            };

            dialog.ShowDialog();
        }

        public (bool? Result, SaveSnapshotPromptAction Action, string SnapshotName) ShowSaveSnapshotPrompt(
            string periodLabel,
            int customerCount,
            string? defaultSnapshotName = null,
            bool canOverwrite = false)
        {
            var vm = new SaveSnapshotPromptViewModel(periodLabel ?? string.Empty, customerCount, defaultSnapshotName, canOverwrite);
            var dialog = new SaveSnapshotWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            var result = dialog.ShowDialog();
            return (result, vm.Action, vm.SnapshotName ?? string.Empty);
        }

        public AppSettings? ShowSettingsDialog(AppSettings settings)
        {
            var vm = new SettingsViewModel(settings ?? new AppSettings());
            var dialog = new SettingsWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return dialog.ShowDialog() == true ? vm.BuildSettings() : null;
        }

        public ImportWizardViewModel? ShowImportWizardDialog(
            string filePath,
            ImportWizardViewModel.ImportWizardMode mode = ImportWizardViewModel.ImportWizardMode.FullDataset)
        {
            var vm = new ImportWizardViewModel(this, filePath, mode);
            Window dialog = mode == ImportWizardViewModel.ImportWizardMode.FullDataset
                ? new ImportWizardWindow()
                : new CurrentIndexImportWindow();

            dialog.Owner = GetOwner();
            dialog.DataContext = vm;

            return dialog.ShowDialog() == true ? vm : null;
        }

        public string? ShowSetColumnValueDialog(string columnTitle, string? initialValue = null)
        {
            var vm = new SetColumnValueViewModel(columnTitle, initialValue);
            var dialog = new SetColumnValueWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return dialog.ShowDialog() == true ? vm.ValueText : null;
        }

        public NewDatasetCreationOption? ShowNewDatasetOptionsDialog()
        {
            var vm = new NewDatasetOptionsViewModel();
            var dialog = new NewDatasetOptionsWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return dialog.ShowDialog() == true ? vm.SelectedOption : null;
        }

        public void OpenWithDefaultApp(string path)
        {
            var info = new ProcessStartInfo(path)
            {
                UseShellExecute = true
            };

            Process.Start(info);
        }

        public string GetSummaryTemplatePath()
        {
            var packaged = Path.Combine(AppContext.BaseDirectory, PackagedSummaryTemplateRelativePath);
            if (File.Exists(packaged))
            {
                return packaged;
            }

            var rootDir = GetSolutionRootDirectory();

            var legacyPath = Path.Combine(rootDir, LegacySummaryTemplateFileName);
            if (File.Exists(legacyPath))
            {
                return legacyPath;
            }

            var candidates = Directory
                .EnumerateFiles(rootDir, "Bảng tổng hợp thu*.xlsx", SearchOption.TopDirectoryOnly)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .ToList();

            var picked = candidates.FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(picked))
            {
                return picked;
            }

            throw new WarningException("Không tìm thấy file Excel template tổng hợp (Bảng tổng hợp thu*.xlsx) cạnh solution.");
        }

        public string GetLegacySummaryTemplatePath()
        {
            var rootDir = GetSolutionRootDirectory();
            var legacyPath = Path.Combine(rootDir, LegacySummaryTemplateFileName);
            if (File.Exists(legacyPath))
            {
                return legacyPath;
            }

            var packaged = Path.Combine(AppContext.BaseDirectory, PackagedSummaryTemplateRelativePath);
            if (File.Exists(packaged))
            {
                return packaged;
            }

            return GetSummaryTemplatePath();
        }

        public string GetInvoiceTemplatePath()
        {
            var packaged = Path.Combine(AppContext.BaseDirectory, PackagedInvoiceTemplateRelativePath);
            if (File.Exists(packaged))
            {
                return packaged;
            }

            var rootDir = GetSolutionRootDirectory();
            var path = Path.Combine(rootDir, InvoiceTemplateFileName);

            if (!File.Exists(path))
            {
                throw new WarningException($"Không tìm thấy file Excel template in mặc định '{InvoiceTemplateFileName}' cạnh solution.");
            }

            return path;
        }

        public void ShowReportWindow(string periodLabel, IEnumerable<Customer> customers, string? issuerName = null)
        {
            var list = customers?.ToList() ?? new List<Customer>();

            var vm = new ReportViewModel(periodLabel, list, issuerName, this);
            var window = new ReportWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            window.ShowDialog();
        }

        public NewPeriodViewModel? ShowNewPeriodDialog(NewPeriodViewModel vm)
        {
            var window = new NewPeriodWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return window.ShowDialog() == true ? vm : null;
        }

        public NewPeriodViewModel? ShowNewPeriodDialog()
        {
            return ShowNewPeriodDialog(new NewPeriodViewModel());
        }

        public PrintRangeViewModel? ShowPrintRangeDialog(int defaultFrom, int defaultTo)
        {
            var vm = new PrintRangeViewModel
            {
                FromNumber = defaultFrom,
                ToNumber = defaultTo
            };

            var window = new PrintRangeWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return window.ShowDialog() == true ? vm : null;
        }

        public IReadOnlyList<Customer>? ShowGroupInvoiceSelectionDialog(string groupName, IReadOnlyList<Customer> customers)
        {
            var vm = new GroupInvoiceSelectionViewModel(groupName ?? string.Empty, customers ?? Array.Empty<Customer>());
            var window = new GroupInvoiceSelectionWindow
            {
                Owner = GetOwner(),
                DataContext = vm
            };

            return window.ShowDialog() == true ? vm.GetSelectedCustomers() : null;
        }

        private static string GetSolutionRootDirectory()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
        }
    }
}
