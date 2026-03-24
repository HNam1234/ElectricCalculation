using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Interop;
using System.Windows.Threading;
using ElectricCalculation.Models;
using ElectricCalculation.ViewModels;
using ElectricCalculation.Views;

namespace ElectricCalculation.Services
{
    internal static class UserGuideSnapshotService
    {
        private const string ReportMenuGuideImageFileName = "report_menu.png";
        private const string ReportGroupGuideImageFileName = "report_group.png";
        private const string GroupInvoiceSelectionGuideImageFileName = "group_invoice_selection.png";

        public static IReadOnlyList<UserGuideSectionItem> BuildGuideSections(Window? startupWindow)
        {
            var sections = new List<UserGuideSectionItem>();
            if (TryBuildSingleBasicFlowSection(startupWindow, out var basicFlowSection))
            {
                sections.Add(basicFlowSection);
            }

            if (TryBuildGroupInvoiceFlowSection(out var groupInvoiceSection))
            {
                sections.Add(groupInvoiceSection);
            }

            return sections;
        }

        private static bool TryBuildGroupInvoiceFlowSection(out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            BitmapSource? reportMenuScreenshot = null;
            if (TryLoadUserGuideImage(ReportMenuGuideImageFileName, out var loadedMenuScreenshot))
            {
                reportMenuScreenshot = loadedMenuScreenshot;
            }
            else if (TryCaptureReportMenuOnMainWindow(out var capturedMenuScreenshot))
            {
                reportMenuScreenshot = capturedMenuScreenshot;
            }

            BitmapSource? mainDetailScreenshot = null;
            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var capturedDetailScreenshot))
            {
                mainDetailScreenshot = capturedDetailScreenshot;
            }

            BitmapSource? reportScreenshot = null;
            if (TryCaptureOpenReportWindow(out var openReportScreenshot))
            {
                reportScreenshot = openReportScreenshot;
            }
            else if (TryLoadUserGuideImage(ReportGroupGuideImageFileName, out var loadedReport))
            {
                reportScreenshot = loadedReport;
            }
            else if (TryCaptureReportWindowWithSampleData(out var capturedReportScreenshot))
            {
                reportScreenshot = capturedReportScreenshot;
            }

            BitmapSource? selectionScreenshot = null;
            if (TryCaptureOpenGroupInvoiceSelectionWindow(out var openSelectionScreenshot))
            {
                selectionScreenshot = openSelectionScreenshot;
            }
            else if (TryLoadUserGuideImage(GroupInvoiceSelectionGuideImageFileName, out var loadedSelection))
            {
                selectionScreenshot = loadedSelection;
            }
            else if (TryCaptureGroupInvoiceSelectionWindowWithSampleData(out var capturedSelectionScreenshot))
            {
                selectionScreenshot = capturedSelectionScreenshot;
            }

            var fallbackScreenshot =
                reportScreenshot ??
                selectionScreenshot ??
                mainDetailScreenshot;

            if (fallbackScreenshot == null)
            {
                section = new UserGuideSectionItem(
                    TabTitle: "In nhóm",
                    Heading: "In hóa đơn theo nhóm",
                    Description: "Không thể chụp ảnh hướng dẫn trong phiên hiện tại.",
                    Steps: Array.Empty<UserGuideStepItem>());
                return false;
            }

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 1: Mở báo cáo theo nhóm",
                Description: "Vào menu Báo cáo → Thống kê theo nhóm… để mở màn hình thống kê.",
                Screenshot: reportMenuScreenshot ?? reportScreenshot ?? selectionScreenshot ?? mainDetailScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 2: Chọn nhóm cần in",
                Description: "Trong màn hình thống kê, chọn nhóm/đơn vị rồi bấm In hóa đơn nhóm.",
                Screenshot: reportScreenshot ?? selectionScreenshot ?? mainDetailScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 3: Chọn hộ và xem trước",
                Description: "Tick/bỏ tick các hộ cần in. Giữ Ctrl + chuột trái để bôi đen nhiều dòng, sau đó chuột phải để Chọn in/Bỏ in nhanh. Nếu muốn custom header, bỏ chọn 'Tự động lấy thông tin (Kính gửi/Địa chỉ/Đại diện/Điện thoại)' rồi nhập tay và xem trước trực tiếp bên phải.",
                Screenshot: selectionScreenshot ?? reportScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 4: Xuất file Excel",
                Description: "Bấm OK → xác nhận → chọn nơi lưu. Hệ thống sẽ tạo file .xlsx gồm 1 sheet hóa đơn gộp cho nhóm.",
                Screenshot: selectionScreenshot ?? reportScreenshot ?? fallbackScreenshot));

            section = new UserGuideSectionItem(
                TabTitle: "In nhóm",
                Heading: "In hóa đơn theo nhóm",
                Description: "4 bước: mở thống kê theo nhóm, chọn nhóm, chọn hộ + tuỳ biến + xem trước, và xuất file Excel.",
                Steps: steps);

            return true;
        }

        private static bool TryLoadUserGuideImage(string fileName, out BitmapSource screenshot)
        {
            screenshot = null!;

            if (string.IsNullOrWhiteSpace(fileName))
            {
                return false;
            }

            try
            {
                var cachePath = Path.Combine(GetUserGuideImageCacheDirectory(), fileName);
                var baseDirPath = Path.Combine(AppContext.BaseDirectory, "Assets", "UserGuide", fileName);

                var path = File.Exists(cachePath) ? cachePath : baseDirPath;
                if (!File.Exists(path))
                {
                    return false;
                }

                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.UriSource = new Uri(path, UriKind.Absolute);
                image.EndInit();
                image.Freeze();

                screenshot = image;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetUserGuideImageCacheDirectory()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "ElectricCalculation", "UserGuide");
        }

        private static void TrySaveUserGuideImage(string fileName, BitmapSource screenshot)
        {
            if (string.IsNullOrWhiteSpace(fileName) || screenshot == null)
            {
                return;
            }

            try
            {
                var folder = GetUserGuideImageCacheDirectory();
                Directory.CreateDirectory(folder);
                var path = Path.Combine(folder, fileName);

                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(screenshot));

                using var stream = File.Create(path);
                encoder.Save(stream);
            }
            catch
            {
                // Best-effort cache.
            }
        }

        private static bool TryCaptureReportMenuOnMainWindow(out BitmapSource screenshot)
        {
            screenshot = null!;

            var app = Application.Current;
            if (app == null)
            {
                return false;
            }

            var window = app.Windows
                .OfType<MainWindow>()
                .FirstOrDefault(w => w.IsVisible && w.WindowState != WindowState.Minimized)
                ?? app.Windows.OfType<MainWindow>().FirstOrDefault()
                ?? app.MainWindow as MainWindow;

            if (window == null)
            {
                return false;
            }

            MenuItem? reportMenu = null;

            try
            {
                window.Dispatcher.Invoke(() =>
                {
                    window.Activate();
                    reportMenu = window.FindName("ReportMenuItem") as MenuItem;
                    if (reportMenu != null)
                    {
                        reportMenu.IsSubmenuOpen = true;
                        reportMenu.UpdateLayout();
                    }
                }, DispatcherPriority.Normal);

                window.Dispatcher.Invoke(() => { }, DispatcherPriority.ApplicationIdle);
                Thread.Sleep(80);

                if (!TryCaptureWindowFromScreen(window, out screenshot))
                {
                    screenshot = CaptureWindowVisual(window, fallbackWidth: 1200, fallbackHeight: 700);
                }

                TrySaveUserGuideImage(ReportMenuGuideImageFileName, screenshot);
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
                    if (reportMenu != null)
                    {
                        window.Dispatcher.Invoke(() => reportMenu.IsSubmenuOpen = false, DispatcherPriority.Normal);
                    }
                }
                catch
                {
                    // Best-effort cleanup.
                }
            }
        }

        private static bool TryCaptureOpenReportWindow(out BitmapSource screenshot)
        {
            screenshot = null!;

            var app = Application.Current;
            if (app == null)
            {
                return false;
            }

            var window = app.Windows
                .OfType<ReportWindow>()
                .FirstOrDefault(w => w.IsVisible && w.WindowState != WindowState.Minimized)
                ?? app.Windows.OfType<ReportWindow>().FirstOrDefault();

            if (window == null)
            {
                return false;
            }

            try
            {
                window.Activate();

                if (!TryCaptureWindowFromScreen(window, out screenshot))
                {
                    screenshot = CaptureWindowVisual(window, fallbackWidth: 1000, fallbackHeight: 540);
                }

                TrySaveUserGuideImage(ReportGroupGuideImageFileName, screenshot);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryCaptureOpenGroupInvoiceSelectionWindow(out BitmapSource screenshot)
        {
            screenshot = null!;

            var app = Application.Current;
            if (app == null)
            {
                return false;
            }

            var window = app.Windows
                .OfType<GroupInvoiceSelectionWindow>()
                .FirstOrDefault(w => w.IsVisible && w.WindowState != WindowState.Minimized)
                ?? app.Windows.OfType<GroupInvoiceSelectionWindow>().FirstOrDefault();

            if (window == null)
            {
                return false;
            }

            try
            {
                window.Activate();

                if (!TryCaptureWindowFromScreen(window, out screenshot))
                {
                    screenshot = CaptureWindowVisual(window, fallbackWidth: 1100, fallbackHeight: 650);
                }

                TrySaveUserGuideImage(GroupInvoiceSelectionGuideImageFileName, screenshot);
                return true;
            }
            catch
            {
                return false;
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out NativeRect rect);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hDc);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hDc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hDc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hDc, IntPtr hObject);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool BitBlt(
            IntPtr hdcDest,
            int nXDest,
            int nYDest,
            int nWidth,
            int nHeight,
            IntPtr hdcSrc,
            int nXSrc,
            int nYSrc,
            int dwRop);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        private static bool TryCaptureWindowFromScreen(Window window, out BitmapSource screenshot)
        {
            screenshot = null!;

            var handle = new WindowInteropHelper(window).Handle;
            if (handle == IntPtr.Zero)
            {
                return false;
            }

            if (!GetWindowRect(handle, out var rect))
            {
                return false;
            }

            var width = rect.Right - rect.Left;
            var height = rect.Bottom - rect.Top;
            if (width <= 0 || height <= 0)
            {
                return false;
            }

            IntPtr screenDc = IntPtr.Zero;
            IntPtr memoryDc = IntPtr.Zero;
            IntPtr hBitmap = IntPtr.Zero;
            IntPtr oldObject = IntPtr.Zero;

            try
            {
                const int SRCCOPY = 0x00CC0020;
                const int CAPTUREBLT = 0x40000000;

                screenDc = GetDC(IntPtr.Zero);
                if (screenDc == IntPtr.Zero)
                {
                    return false;
                }

                memoryDc = CreateCompatibleDC(screenDc);
                if (memoryDc == IntPtr.Zero)
                {
                    return false;
                }

                hBitmap = CreateCompatibleBitmap(screenDc, width, height);
                if (hBitmap == IntPtr.Zero)
                {
                    return false;
                }

                oldObject = SelectObject(memoryDc, hBitmap);
                if (oldObject == IntPtr.Zero)
                {
                    return false;
                }

                if (!BitBlt(memoryDc, 0, 0, width, height, screenDc, rect.Left, rect.Top, SRCCOPY | CAPTUREBLT))
                {
                    return false;
                }

                var source = Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());

                source.Freeze();
                screenshot = source;
                return true;
            }
            finally
            {
                if (oldObject != IntPtr.Zero && memoryDc != IntPtr.Zero)
                {
                    SelectObject(memoryDc, oldObject);
                }

                if (hBitmap != IntPtr.Zero)
                {
                    DeleteObject(hBitmap);
                }

                if (memoryDc != IntPtr.Zero)
                {
                    DeleteDC(memoryDc);
                }

                if (screenDc != IntPtr.Zero)
                {
                    ReleaseDC(IntPtr.Zero, screenDc);
                }
            }
        }

        private static bool TryBuildSingleBasicFlowSection(Window? startupWindow, out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            BitmapSource? startupScreenshot = null;
            if (TryCaptureStartupWindow(startupWindow, out var capturedStartupScreenshot))
            {
                startupScreenshot = capturedStartupScreenshot;
            }

            BitmapSource? importFromExcelOptionScreenshot = null;
            if (TryCaptureNewDatasetOptionsWindow(out var capturedImportFromExcelOptionScreenshot))
            {
                importFromExcelOptionScreenshot = capturedImportFromExcelOptionScreenshot;
            }

            BitmapSource? mainDetailScreenshot = null;
            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var capturedDetailScreenshot))
            {
                mainDetailScreenshot = capturedDetailScreenshot;
            }

            BitmapSource? fastEntryScreenshot = null;
            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.FastEntry, out var capturedFastEntryScreenshot))
            {
                fastEntryScreenshot = capturedFastEntryScreenshot;
            }

            BitmapSource? singleInvoiceScreenshot = null;
            if (TryCaptureSingleInvoiceWindow(out var capturedSingleInvoiceScreenshot))
            {
                singleInvoiceScreenshot = capturedSingleInvoiceScreenshot;
            }

            var fallbackScreenshot =
                startupScreenshot ??
                importFromExcelOptionScreenshot ??
                mainDetailScreenshot ??
                fastEntryScreenshot ??
                singleInvoiceScreenshot;

            if (fallbackScreenshot == null)
            {
                section = new UserGuideSectionItem(
                    TabTitle: "Flow cơ bản",
                    Heading: "Flow cơ bản từ dữ liệu trống đến xuất hóa đơn 1 hộ",
                    Description: "Không thể chụp ảnh hướng dẫn trong phiên hiện tại.",
                    Steps: Array.Empty<UserGuideStepItem>());
                return false;
            }

            BitmapSource? wizardFileSelectionScreenshot = null;
            BitmapSource? wizardMappingScreenshot = null;
            BitmapSource? wizardReviewAndImportScreenshot = null;

            var sampleExcelPath = GetGuideSampleExcelPath();

            if (!string.IsNullOrWhiteSpace(sampleExcelPath) && File.Exists(sampleExcelPath))
            {
                var ui = new UiService();
                var wizardScreenshots = CaptureBasicImportWizardScreenshots(ui, sampleExcelPath);
                wizardFileSelectionScreenshot = wizardScreenshots.FileSelection;
                wizardMappingScreenshot = wizardScreenshots.Mapping;
                wizardReviewAndImportScreenshot = wizardScreenshots.ReviewAndImport;
            }

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 1: Màn hình chính",
                Description: "Bắt đầu từ màn hình chính khi chưa có dữ liệu nhập, chuẩn bị tạo bộ dữ liệu mới.",
                Screenshot: startupScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 2: Chọn Import từ Excel",
                Description: "Trong cửa sổ tạo bộ dữ liệu mới, chọn phương án Import từ Excel.",
                Screenshot: importFromExcelOptionScreenshot ?? startupScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 3: Chọn file Excel chứa data",
                Description: "Trong ImportWizard, chọn file dữ liệu Excel cần nhập.",
                Screenshot: wizardFileSelectionScreenshot ?? importFromExcelOptionScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 4: Mapping dữ liệu",
                Description: "Ghép cột trong file Excel với các trường cần nhập trên hệ thống tại bước Ghép cột.",
                Screenshot: wizardMappingScreenshot ?? wizardFileSelectionScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 5: Kiểm tra, nhập dữ liệu, hoàn tất",
                Description: "Ở bước Kiểm tra & nhập, bấm Kiểm tra, sau đó Nhập dữ liệu và Hoàn tất để kết thúc import.",
                Screenshot: wizardReviewAndImportScreenshot ?? wizardMappingScreenshot ?? wizardFileSelectionScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 6: Giao diện Nhập nhanh",
                Description: "Sau khi import xong, chuyển sang chế độ Nhập nhanh để cập nhật chỉ số nhanh theo dòng.",
                Screenshot: fastEntryScreenshot ?? mainDetailScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 7: Chọn Chi tiết để hiển thị tất cả",
                Description: "Chuyển lại tab Chi tiết để xem đầy đủ tất cả cột thông tin của khách hàng.",
                Screenshot: mainDetailScreenshot ?? fastEntryScreenshot ?? fallbackScreenshot));

            steps.Add(new UserGuideStepItem(
                StepTitle: "Bước 8: Xuất hóa đơn cho 1 hộ",
                Description: "Chọn một khách hàng rồi thực hiện xuất hóa đơn/in hóa đơn cho đúng 1 hộ.",
                Screenshot: singleInvoiceScreenshot ?? mainDetailScreenshot ?? fallbackScreenshot));

            section = new UserGuideSectionItem(
                TabTitle: "Flow cơ bản",
                Heading: "Flow chuẩn: import dữ liệu và xuất hóa đơn 1 hộ",
                Description: "Flow duy nhất gồm 8 bước: màn hình chính, import Excel, chọn file, mapping, kiểm tra/nhập/hoàn tất, nhập nhanh, quay lại chi tiết và xuất hóa đơn 1 hộ.",
                Steps: steps);

            return true;
        }

        private static (BitmapSource? FileSelection, BitmapSource? Mapping, BitmapSource? ReviewAndImport) CaptureBasicImportWizardScreenshots(
            UiService ui,
            string sampleExcelPath)
        {
            BitmapSource? fileSelectionScreenshot = null;
            BitmapSource? mappingScreenshot = null;
            BitmapSource? reviewAndImportScreenshot = null;
            ImportWizardWindow? window = null;

            try
            {
                var vm = new ImportWizardViewModel(ui, sampleExcelPath, ImportWizardViewModel.ImportWizardMode.FullDataset);
                window = new ImportWizardWindow
                {
                    DataContext = vm
                };

                vm.CurrentStep = 0;
                if (TryCaptureWindowVisual(window, 1180, 780, out var sourceStepScreenshot))
                {
                    fileSelectionScreenshot = sourceStepScreenshot;
                }

                vm.CurrentStep = 1;
                if (TryCaptureWindowVisual(window, 1180, 780, out var mappingStepScreenshot))
                {
                    mappingScreenshot = mappingStepScreenshot;
                }

                vm.CurrentStep = 2;
                vm.ValidateCommand.Execute(null);
                if (TryCaptureWindowVisual(window, 1180, 780, out var reviewStepScreenshot))
                {
                    reviewAndImportScreenshot = reviewStepScreenshot;
                }
            }
            catch
            {
                // Keep basic flow available even if wizard capture fails.
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

            return (fileSelectionScreenshot, mappingScreenshot, reviewAndImportScreenshot);
        }

        private static bool TryCaptureNewDatasetOptionsWindow(out BitmapSource screenshot)
        {
            screenshot = null!;
            var window = new NewDatasetOptionsWindow
            {
                DataContext = new NewDatasetOptionsViewModel()
            };

            try
            {
                screenshot = CaptureWindowVisual(window, fallbackWidth: 900, fallbackHeight: 620);
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
                    // Ignore cleanup errors.
                }
            }
        }

        private static bool TryCaptureSingleInvoiceWindow(out BitmapSource screenshot)
        {
            screenshot = null!;
            var window = new InvoiceWindow
            {
                DataContext = new InvoiceViewModel(BuildGuideInvoiceCustomer(), $"Tháng {DateTime.Now.Month} năm {DateTime.Now.Year}")
            };

            try
            {
                screenshot = CaptureWindowVisual(window, fallbackWidth: 860, fallbackHeight: 620);
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
                    // Ignore cleanup errors.
                }
            }
        }

        private static bool TryCaptureReportWindowWithSampleData(out BitmapSource screenshot)
        {
            screenshot = null!;

            var customers = BuildGuideReportCustomers();
            if (customers.Count == 0)
            {
                return false;
            }

            var window = new ReportWindow
            {
                DataContext = new ReportViewModel(
                    periodLabel: $"Tháng {DateTime.Now:MM/yyyy}",
                    customers: customers,
                    issuerName: "Nguyễn Văn A",
                    ui: new UiService())
            };

            try
            {
                screenshot = CaptureWindowVisual(window, fallbackWidth: 1000, fallbackHeight: 540);
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
                    // Ignore cleanup errors.
                }
            }
        }

        private static bool TryCaptureGroupInvoiceSelectionWindowWithSampleData(out BitmapSource screenshot)
        {
            screenshot = null!;

            var customers = BuildGuideReportCustomers()
                .Where(c => string.Equals(c.GroupName, "Khu A", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (customers.Count == 0)
            {
                return false;
            }

            var vm = new GroupInvoiceSelectionViewModel(
                groupName: "Khu A",
                customers: customers,
                periodLabel: $"Tháng {DateTime.Now:MM/yyyy}",
                issuerName: "Nguyễn Văn A",
                issuePlace: "Hà Nội",
                issueDate: DateTime.Today)
            {
                UseAutoHeaderFields = false,
                RecipientName = "Khu A (Custom)",
                ConsumptionAddress = "Tòa A - Tầng 1",
                RepresentativeName = "Ban quản lý Khu A",
                HouseholdPhone = "0243 123 4567",
                RepresentativePhone = "0906 123 357"
            };

            var window = new GroupInvoiceSelectionWindow
            {
                DataContext = vm
            };

            try
            {
                screenshot = CaptureWindowVisual(window, fallbackWidth: 1100, fallbackHeight: 650);
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
                    // Ignore cleanup errors.
                }
            }
        }

        private static List<Customer> BuildGuideReportCustomers()
        {
            return new List<Customer>
            {
                new Customer
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
                },
                new Customer
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
                },
                new Customer
                {
                    SequenceNumber = 3,
                    Name = "Phòng B201",
                    GroupName = "Khu B",
                    Category = "Dịch vụ",
                    Address = "Tòa B - Tầng 2",
                    MeterNumber = "CT-B201",
                    Location = "B2",
                    PreviousIndex = 2010m,
                    CurrentIndex = 2080m,
                    Multiplier = 1m,
                    UnitPrice = 4169m
                },
                new Customer
                {
                    SequenceNumber = 4,
                    Name = "Phòng B202",
                    GroupName = "Khu B",
                    Category = "Dịch vụ",
                    Address = "Tòa B - Tầng 2",
                    MeterNumber = "CT-B202",
                    Location = "B2",
                    PreviousIndex = 1540m,
                    CurrentIndex = 1615m,
                    Multiplier = 1m,
                    UnitPrice = 4169m
                }
            };
        }

        private static Customer BuildGuideInvoiceCustomer()
        {
            return new Customer
            {
                SequenceNumber = 7,
                Name = "Quán ăn uống giải khát c.Ly",
                Address = "Tầng 1 Số 10 TQB",
                GroupName = "HDKT",
                RepresentativeName = "Nguyễn Hương Ly",
                Phone = "0945656446",
                HouseholdPhone = "0945656446",
                MeterNumber = "5089",
                Substation = "B1",
                BuildingName = "Số 10 TQB",
                Page = "4",
                Location = "Tủ tổng T1",
                PreviousIndex = 436148m,
                CurrentIndex = 439092m,
                Multiplier = 1m,
                UnitPrice = 4169m,
                SubsidizedKwh = 0m
            };
        }

        private static bool TryCaptureStartupWindow(Window? startupWindow, out BitmapSource screenshot)
        {
            screenshot = null!;

            if (startupWindow == null)
            {
                return false;
            }

            var overlayElement = startupWindow.FindName("LoadingOverlay") as UIElement;
            var overlayOpacity = overlayElement?.Opacity;

            try
            {
                if (overlayElement != null && overlayOpacity is > 0)
                {
                    overlayElement.Opacity = 0;
                }

                screenshot = CaptureWindowVisual(startupWindow, fallbackWidth: 1100, fallbackHeight: 720);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (overlayElement != null && overlayOpacity.HasValue)
                {
                    overlayElement.Opacity = overlayOpacity.Value;
                }
            }
        }

        private static bool TryCaptureMainWindowWithSampleData(
            MainWindowGuideCaptureMode mode,
            out BitmapSource screenshot)
        {
            screenshot = null!;
            var window = new MainWindow();

            try
            {
                if (window.DataContext is MainWindowViewModel vm)
                {
                    PopulateGuideCustomers(vm);

                    switch (mode)
                    {
                        case MainWindowGuideCaptureMode.Search:
                            vm.SelectedSearchField = vm.SearchFields.Count > 0 ? vm.SearchFields[0] : string.Empty;
                            vm.SearchText = "A";
                            vm.ApplySearchCommand.Execute(null);
                            break;
                        case MainWindowGuideCaptureMode.FastEntry:
                            vm.IsFastEntryMode = true;
                            break;
                        default:
                            vm.IsFastEntryMode = false;
                            break;
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

        private static void PopulateGuideCustomers(MainWindowViewModel vm)
        {
            if (TryPopulateCustomersFromSampleExcel(vm))
            {
                return;
            }

            PopulateDemoCustomers(vm);
        }

        private static bool TryPopulateCustomersFromSampleExcel(MainWindowViewModel vm)
        {
            var sampleExcelPath = GetGuideSampleExcelPath();
            if (string.IsNullOrWhiteSpace(sampleExcelPath) || !File.Exists(sampleExcelPath))
            {
                return false;
            }

            try
            {
                var preview = ExcelImportService.BuildPreview(sampleExcelPath);
                var map = BuildGuideImportMap(preview.Columns);
                if (!map.ContainsKey(ExcelImportService.ImportField.Name))
                {
                    return false;
                }

                var imported = ExcelImportService.ImportFromFile(
                    sampleExcelPath,
                    preview.SelectedSheetName,
                    map,
                    preview.DataStartRowIndex,
                    out _,
                    out _);

                if (imported.Count == 0)
                {
                    return false;
                }

                vm.PeriodLabel = $"Tháng {DateTime.Now:MM/yyyy}";
                vm.InvoiceIssuer = "Nguyễn Văn A";
                vm.Customers.Clear();
                foreach (var customer in imported)
                {
                    vm.Customers.Add(customer);
                }

                vm.CustomersView.Refresh();
                vm.IsDirty = false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Dictionary<ExcelImportService.ImportField, string> BuildGuideImportMap(
            IReadOnlyList<ImportColumnPreview> columns)
        {
            var bestColumns = new Dictionary<ExcelImportService.ImportField, (string ColumnLetter, double Score)>();

            foreach (var column in columns)
            {
                if (!column.SuggestedField.HasValue || string.IsNullOrWhiteSpace(column.ColumnLetter))
                {
                    continue;
                }

                var field = column.SuggestedField.Value;
                var candidate = (ColumnLetter: column.ColumnLetter, Score: column.SuggestedScore);
                if (!bestColumns.TryGetValue(field, out var existing) || candidate.Score > existing.Score)
                {
                    bestColumns[field] = candidate;
                }
            }

            var result = new Dictionary<ExcelImportService.ImportField, string>();
            foreach (var pair in bestColumns)
            {
                result[pair.Key] = pair.Value.ColumnLetter;
            }

            return result;
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

        private enum MainWindowGuideCaptureMode
        {
            Detail,
            Search,
            FastEntry
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

