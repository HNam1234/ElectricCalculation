using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using ElectricCalculation.Models;
using ElectricCalculation.Services;
using ElectricCalculation.ViewModels;
using Microsoft.Win32;

namespace ElectricCalculation.Views
{
    public partial class ReportWindow : Window
    {
        public ReportWindow()
        {
            InitializeComponent();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void PrintGroupButton_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not ReportViewModel viewModel)
            {
                return;
            }

            var item = viewModel.SelectedItem;
            if (item == null)
            {
                MessageBox.Show(
                    this,
                    "Hãy chọn một nhóm / đơn vị ở bảng bên phải trước.",
                    "In Excel nhóm",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var customers = viewModel.GetCustomersForGroup(item).ToList();
            if (customers.Count == 0)
            {
                MessageBox.Show(
                    this,
                    "Nhóm được chọn hiện không có dữ liệu khách hàng.",
                    "In Excel nhóm",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                FileName = $"Tien dien - {item.GroupName}.xlsx"
            };

            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            string templatePath;
            try
            {
                templatePath = GetGroupTemplatePath();
            }
            catch (WarningException ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "In Excel nhóm",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "In Excel nhóm",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            try
            {
                ExcelExportService.ExportToFile(templatePath, dialog.FileName, customers);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "Lỗi in Excel nhóm",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private static string GetGroupTemplatePath()
        {
            // Sử dụng file tổng hợp mặc định cạnh solution
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var rootDir = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", ".."));
            var templatePath = Path.Combine(rootDir, "Bảng tổng hợp thu tháng 6 năm 2025.xlsx");

            if (!File.Exists(templatePath))
            {
                throw new WarningException(
                    "Không tìm thấy file Excel template tổng hợp mặc định 'Bảng tổng hợp thu tháng 6 năm 2025.xlsx' cạnh solution.");
            }

            return templatePath;
        }
    }
}

