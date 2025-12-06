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
using Microsoft.Win32;

namespace ElectricCalculation.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            if (DataContext is not MainWindowViewModel viewModel)
            {
                return;
            }

            try
            {
                viewModel.ImportFromExcelCommand.Execute(dialog.FileName);
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                ShowMessageDialog("Cảnh báo import Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                ShowMessageDialog("Lỗi import Excel", ex.Message);
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = "Bang tong hop dien.xlsx"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            if (DataContext is not MainWindowViewModel viewModel)
            {
                return;
            }

            try
            {
                viewModel.ExportToExcelCommand.Execute(dialog.FileName);
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                ShowMessageDialog("Cảnh báo export Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                ShowMessageDialog("Lỗi export Excel", ex.Message);
            }
        }

        // Nút In: ưu tiên dòng đang chọn, nếu không có thì in toàn bộ danh sách đang lọc ra một file Excel theo mẫu gốc
        private void PrintInvoice_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainWindowViewModel viewModel)
            {
                return;
            }

            try
            {
                var dialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    FileName = "Hoa don tien dien.xlsx"
                };

                if (dialog.ShowDialog(this) != true)
                {
                    return;
                }

                var outputPath = dialog.FileName;

                if (viewModel.SelectedCustomer != null)
                {
                    var templatePath = GetInvoiceTemplatePath();

                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        outputPath,
                        viewModel.SelectedCustomer,
                        viewModel.PeriodLabel,
                        viewModel.InvoiceIssuer);
                }
                else
                {
                    viewModel.ExportFilteredToExcelCommand.Execute(outputPath);
                }

                ShowMessageDialog("In Excel",
                    $"Đã tạo file hóa đơn Excel tại:\n{outputPath}\n\nMở file này bằng Excel để xem Print Preview / in ra giấy.");
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                ShowMessageDialog("In Excel", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                ShowMessageDialog("Lỗi in Excel", ex.Message);
            }
        }

        private static void OpenWithDefaultApp(string path)
        {
            var info = new ProcessStartInfo(path)
            {
                UseShellExecute = true
            };
            Process.Start(info);
        }

        private static string GetInvoiceTemplatePath()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var rootDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
            var templatePath = Path.Combine(rootDir, "DefaultTemplate.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new WarningException("KhA'ng tAªm th §y file Excel template in m §úc Ž` ¯<nh 'DefaultTemplate.xlsx' c §­nh solution.");
            }

            return templatePath;
        }

        private void ShowMessageDialog(string title, string message)
        {
            var vm = new MessageDialogViewModel(title, message);
            var dialog = new MessageDialogWindow
            {
                Owner = this,
                DataContext = vm
            };

            dialog.ShowDialog();
        }

        private void DataGrid_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (sender is not DependencyObject dep)
            {
                return;
            }

            var scrollViewer = FindVisualChild<ScrollViewer>(dep);
            if (scrollViewer == null)
            {
                return;
            }

            var offset = scrollViewer.VerticalOffset - e.Delta / 3.0;
            if (offset < 0)
            {
                offset = 0;
            }

            scrollViewer.ScrollToVerticalOffset(offset);
            e.Handled = true;
        }

        private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            var childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (var i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T match)
                {
                    return match;
                }

                var descendant = FindVisualChild<T>(child);
                if (descendant != null)
                {
                    return descendant;
                }
            }

            return null;
        }
    }
}
