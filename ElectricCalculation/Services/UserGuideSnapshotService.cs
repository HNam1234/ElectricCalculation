using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ElectricCalculation.Models;
using ElectricCalculation.ViewModels;
using ElectricCalculation.Views;

namespace ElectricCalculation.Services
{
    internal static class UserGuideSnapshotService
    {
        public static IReadOnlyList<UserGuideStepItem> BuildGuideSteps(Window? startupWindow)
        {
            var steps = new List<UserGuideStepItem>();

            if (TryCaptureStartupWindow(startupWindow, out var startupScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "Bước 1: Chọn thao tác ở màn hình Trang chủ",
                    Description: "Trang chủ: bạn chọn Tạo bộ dữ liệu mới, Tạo tháng mới hoặc mở bộ dữ liệu gần đây.",
                    Screenshot: startupScreenshot));
            }

            AddDetailedImportExcelSteps(steps);

            if (TryCaptureMainWindowWithSampleData(searchDemo: false, out var editorScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "Bước cuối: Nhập/chỉnh chỉ số trên màn hình chính",
                    Description: "Màn hình nhập liệu với dữ liệu mẫu. Bạn nhập chỉ số mới trực tiếp trong bảng.",
                    Screenshot: editorScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(searchDemo: true, out var searchScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "Mẹo: Tìm kiếm nhanh và lọc dữ liệu",
                    Description: "Ví dụ sau khi nhập từ khóa và nhấn Enter để áp dụng tìm kiếm.",
                    Screenshot: searchScreenshot));
            }

            return steps;
        }

        private static void AddDetailedImportExcelSteps(ICollection<UserGuideStepItem> steps)
        {
            var sampleExcelPath = GetGuideSampleExcelPath();
            if (string.IsNullOrWhiteSpace(sampleExcelPath) || !File.Exists(sampleExcelPath))
            {
                return;
            }

            var ui = new UiService();

            if (TryCaptureImportWizardFullDatasetSteps(ui, sampleExcelPath, out var fullDatasetSteps))
            {
                foreach (var step in fullDatasetSteps)
                {
                    steps.Add(step);
                }
            }

            if (TryCaptureImportWizardCurrentIndexSteps(ui, sampleExcelPath, out var currentIndexSteps))
            {
                foreach (var step in currentIndexSteps)
                {
                    steps.Add(step);
                }
            }
        }

        private static bool TryCaptureImportWizardFullDatasetSteps(
            UiService ui,
            string sampleExcelPath,
            out IReadOnlyList<UserGuideStepItem> steps)
        {
            var captured = new List<UserGuideStepItem>();
            ImportWizardWindow? window = null;

            try
            {
                var vm = new ImportWizardViewModel(ui, sampleExcelPath, ImportWizardViewModel.ImportWizardMode.FullDataset);
                window = new ImportWizardWindow
                {
                    DataContext = vm
                };

                vm.CurrentStep = 0;
                if (TryCaptureWindowVisual(window, 1180, 780, out var step1Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "Import Excel (chi tiết) - Bước 1: Chọn file/sheet",
                        Description:
                            "1. Vào Tệp -> Import từ Excel.\n" +
                            "2. Chọn đúng file .xlsx và đúng Sheet.\n" +
                            "3. Kiểm tra 'Dòng tiêu đề là dòng số' để cột nhận diện đúng.",
                        Screenshot: step1Screenshot));
                }

                vm.CurrentStep = 1;
                if (TryCaptureWindowVisual(window, 1180, 780, out var step2Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "Import Excel (chi tiết) - Bước 2: Ghép cột (mapping)",
                        Description:
                            "1. Chọn cột trong file cho từng trường trên phần mềm.\n" +
                            "2. Bắt buộc phải có 'Tên khách'.\n" +
                            "3. Nên map thêm Số công tơ, Chỉ số cũ, Chỉ số mới, Đơn giá để dữ liệu đầy đủ.",
                        Screenshot: step2Screenshot));
                }

                vm.CurrentStep = 2;
                vm.ValidateCommand.Execute(null);
                if (TryCaptureWindowVisual(window, 1180, 780, out var step3Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "Import Excel (chi tiết) - Bước 3: Kiểm tra và nhập",
                        Description:
                            "1. Bấm 'Kiểm tra' để xem lỗi/cảnh báo.\n" +
                            "2. Sửa mapping nếu còn lỗi đỏ.\n" +
                            "3. Bấm 'Nhập dữ liệu' rồi 'Xong' khi kết quả import đạt yêu cầu.",
                        Screenshot: step3Screenshot));
                }
            }
            catch
            {
                // Keep other guide sections available if import capture fails.
            }
            finally
            {
                try
                {
                    window?.Close();
                }
                catch
                {
                    // Ignore cleanup errors.
                }
            }

            steps = captured;
            return captured.Count > 0;
        }

        private static bool TryCaptureImportWizardCurrentIndexSteps(
            UiService ui,
            string sampleExcelPath,
            out IReadOnlyList<UserGuideStepItem> steps)
        {
            var captured = new List<UserGuideStepItem>();
            ImportWizardWindow? window = null;

            try
            {
                var vm = new ImportWizardViewModel(ui, sampleExcelPath, ImportWizardViewModel.ImportWizardMode.CurrentIndexOnly);
                window = new ImportWizardWindow
                {
                    DataContext = vm
                };

                vm.CurrentStep = 1;
                if (TryCaptureWindowVisual(window, 1180, 780, out var stepScreenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "Import chỉ số mới theo cột (mapping)",
                        Description:
                            "1. Vào Tệp -> Import chỉ số mới từ Excel.\n" +
                            "2. Bắt buộc map 'Chỉ số mới'.\n" +
                            "3. Bắt buộc chọn ít nhất 1 khóa ghép: Số công tơ / Số thứ tự / Tên khách.\n" +
                            "4. Kiểm tra lại kết quả cập nhật sau khi import.",
                        Screenshot: stepScreenshot));
                }
            }
            catch
            {
                // Keep other guide sections available if import capture fails.
            }
            finally
            {
                try
                {
                    window?.Close();
                }
                catch
                {
                    // Ignore cleanup errors.
                }
            }

            steps = captured;
            return captured.Count > 0;
        }

        private static bool TryCaptureStartupWindow(Window? startupWindow, out BitmapSource screenshot)
        {
            screenshot = null!;

            if (startupWindow == null)
            {
                return false;
            }

            try
            {
                screenshot = CaptureWindowVisual(startupWindow, fallbackWidth: 1100, fallbackHeight: 720);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryCaptureMainWindowWithSampleData(bool searchDemo, out BitmapSource screenshot)
        {
            screenshot = null!;
            var window = new MainWindow();

            try
            {
                if (window.DataContext is MainWindowViewModel vm)
                {
                    PopulateDemoCustomers(vm);

                    if (searchDemo)
                    {
                        vm.SelectedSearchField = vm.SearchFields.Count > 0 ? vm.SearchFields[0] : string.Empty;
                        vm.SearchText = "A";
                        vm.ApplySearchCommand.Execute(null);
                    }
                }

                screenshot = CaptureWindowVisual(window, fallbackWidth: 1280, fallbackHeight: 720);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    window.Close();
                }
                catch
                {
                    // Best-effort cleanup.
                }
            }
        }

        private static bool TryCaptureWindowVisual(
            Window window,
            double fallbackWidth,
            double fallbackHeight,
            out BitmapSource screenshot)
        {
            screenshot = null!;

            try
            {
                screenshot = CaptureWindowVisual(window, fallbackWidth, fallbackHeight);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static BitmapSource CaptureWindowVisual(Window window, double fallbackWidth, double fallbackHeight)
        {
            if (window.Content is not FrameworkElement root)
            {
                throw new InvalidOperationException("Window does not have a renderable root element.");
            }

            var width = ResolveSize(window.ActualWidth, window.Width, root.ActualWidth, root.Width, fallbackWidth);
            var height = ResolveSize(window.ActualHeight, window.Height, root.ActualHeight, root.Height, fallbackHeight);

            root.Measure(new Size(width, height));
            root.Arrange(new Rect(0, 0, width, height));
            root.UpdateLayout();

            var pixelWidth = Math.Max(1, (int)Math.Ceiling(width));
            var pixelHeight = Math.Max(1, (int)Math.Ceiling(height));

            var bitmap = new RenderTargetBitmap(pixelWidth, pixelHeight, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(root);
            bitmap.Freeze();
            return bitmap;
        }

        private static double ResolveSize(
            double windowActual,
            double windowDeclared,
            double rootActual,
            double rootDeclared,
            double fallback)
        {
            if (IsUsableSize(windowActual))
            {
                return windowActual;
            }

            if (IsUsableSize(windowDeclared))
            {
                return windowDeclared;
            }

            if (IsUsableSize(rootActual))
            {
                return rootActual;
            }

            if (IsUsableSize(rootDeclared))
            {
                return rootDeclared;
            }

            return fallback;
        }

        private static bool IsUsableSize(double value)
        {
            return !double.IsNaN(value) && !double.IsInfinity(value) && value > 1;
        }

        private static void PopulateDemoCustomers(MainWindowViewModel vm)
        {
            vm.PeriodLabel = $"Tháng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
            vm.InvoiceIssuer = "Nguyễn Văn A";

            vm.Customers.Clear();

            vm.Customers.Add(new Customer
            {
                SequenceNumber = 1,
                Name = "Phòng A101",
                GroupName = "Khu A",
                Category = "Sinh hoạt",
                Address = "Tòa A - Tầng 1",
                MeterNumber = "CT-A101",
                Location = "A1",
                PreviousIndex = 1250m,
                CurrentIndex = 1288m,
                Multiplier = 1m,
                UnitPrice = 3505m
            });

            vm.Customers.Add(new Customer
            {
                SequenceNumber = 2,
                Name = "Phòng A102",
                GroupName = "Khu A",
                Category = "Sinh hoạt",
                Address = "Tòa A - Tầng 1",
                MeterNumber = "CT-A102",
                Location = "A1",
                PreviousIndex = 980m,
                CurrentIndex = 1032m,
                Multiplier = 1m,
                UnitPrice = 3505m
            });

            vm.Customers.Add(new Customer
            {
                SequenceNumber = 3,
                Name = "Phòng B201",
                GroupName = "Khu B",
                Category = "Dịch vụ",
                Address = "Tòa B - Tầng 2",
                MeterNumber = "CT-B201",
                Location = "B2",
                PreviousIndex = 2010m,
                CurrentIndex = null,
                Multiplier = 1m,
                UnitPrice = 4169m
            });

            vm.Customers.Add(new Customer
            {
                SequenceNumber = 4,
                Name = "Phòng B202",
                GroupName = "Khu B",
                Category = "Dịch vụ",
                Address = "Tòa B - Tầng 2",
                MeterNumber = "CT-B202",
                Location = "B2",
                PreviousIndex = 1540m,
                CurrentIndex = 1530m,
                Multiplier = 1m,
                UnitPrice = 4169m
            });

            vm.CustomersView.Refresh();
            vm.IsDirty = false;
        }

        private static string GetGuideSampleExcelPath()
        {
            return Path.Combine(
                AppContext.BaseDirectory,
                "SampleData",
                "Bang_tong_hop_thu_thang_06_2025.xlsx");
        }
    }
}
