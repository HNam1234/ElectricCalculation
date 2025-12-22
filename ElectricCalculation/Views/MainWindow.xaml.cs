using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ElectricCalculation.Models;
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

        // Nút In: ưu tiên dùng khách đang chọn; nếu không có thì in toàn bộ danh sách đang lọc ra Excel
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

                var openResult = MessageBox.Show(
                    this,
                    $"Đã tạo file Excel tại:\n{outputPath}\n\nBạn có muốn mở file này bằng Excel để xem / chỉnh sửa và in không?",
                    "In Excel",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (openResult == MessageBoxResult.Yes)
                {
                    OpenWithDefaultApp(outputPath);
                }
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

        // In nhiều phiếu theo template DefaultTemplate.xlsx cho danh sách khách hàng (ưu tiên các dòng đang chọn)
        private void PrintMultipleInvoices_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainWindowViewModel viewModel)
            {
                return;
            }

            try
            {
                // Ưu tiên in các khách hàng đang được chọn; nếu không chọn gì thì in toàn bộ danh sách đang lọc
                var selectedCustomers = CustomerGrid.SelectedItems
                    .OfType<Customer>()
                    .ToList();

                var customers = selectedCustomers.Count > 0
                    ? selectedCustomers
                    : viewModel.CustomersView.Cast<Customer>().ToList();

                if (customers.Count == 0)
                {
                    ShowMessageDialog("In nhiều phiếu", "Không có dữ liệu trong danh sách hiện tại để in.");
                    return;
                }

                var dialog = new SaveFileDialog
                {
                    Title = "Chọn thư mục để lưu các hóa đơn",
                    Filter = "Thư mục|*.folder",
                    FileName = "Chon_thu_muc_o_day"
                };

                if (dialog.ShowDialog(this) != true)
                {
                    return;
                }

                var folder = Path.GetDirectoryName(dialog.FileName);
                if (string.IsNullOrWhiteSpace(folder))
                {
                    ShowMessageDialog("In nhiều phiếu", "Đường dẫn thư mục không hợp lệ.");
                    return;
                }

                var templatePath = GetInvoiceTemplatePath();

                foreach (var customer in customers)
                {
                    var namePart = string.IsNullOrWhiteSpace(customer.Name)
                        ? "Khach"
                        : customer.Name;

                    var meterPart = string.IsNullOrWhiteSpace(customer.MeterNumber)
                        ? string.Empty
                        : $" - {customer.MeterNumber}";

                    var safeName = MakeSafeFileName($"{namePart}{meterPart}");
                    var filePath = Path.Combine(folder, $"Hoa don - {safeName}.xlsx");

                    InvoiceExcelExportService.ExportInvoice(
                        templatePath,
                        filePath,
                        customer,
                        viewModel.PeriodLabel,
                        viewModel.InvoiceIssuer);
                }

                ShowMessageDialog("In nhiều phiếu", $"Đã tạo {customers.Count} hóa đơn trong thư mục:\n{folder}");
            }
            catch (WarningException warning)
            {
                Debug.WriteLine(warning);
                ShowMessageDialog("In nhiều phiếu", warning.Message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                ShowMessageDialog("Lỗi in nhiều phiếu", ex.Message);
            }
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return "Hoa_don";
            }

            return name;
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
                throw new WarningException("Không tìm thấy file Excel template in mặc định 'DefaultTemplate.xlsx' cạnh solution.");
            }

            return templatePath;
        }

        private void ShowReport_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainWindowViewModel viewModel)
            {
                return;
            }

            var currentItems = viewModel.CustomersView.Cast<Customer>().ToList();
            if (currentItems.Count == 0)
            {
                ShowMessageDialog("Báo cáo", "Không có dữ liệu trong danh sách hiện tại để lập báo cáo.");
                return;
            }

            var vm = new ReportViewModel(viewModel.PeriodLabel, currentItems);
            var window = new ReportWindow
            {
                Owner = this,
                DataContext = vm
            };

            window.ShowDialog();
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

