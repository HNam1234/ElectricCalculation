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
        public static IReadOnlyList<UserGuideSectionItem> BuildGuideSections(Window? startupWindow)
        {
            var sections = new List<UserGuideSectionItem>();
            if (TryBuildSingleBasicFlowSection(startupWindow, out var basicFlowSection))
            {
                sections.Add(basicFlowSection);
            }

            return sections;
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
                Description: "Trong ImportWizard, chọn file dữ liệu Excel và chọn đúng sheet dữ liệu.",
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

        private static bool TryBuildOverviewSection(BitmapSource? startupScreenshot, out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            if (startupScreenshot != null)
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 1: Chá»n thao tÃ¡c á»Ÿ mÃ n hÃ¬nh Trang chá»§",
                    Description: "Tá»« Trang chá»§, báº¡n chá»n táº¡o bá»™ dá»¯ liá»‡u má»›i, lÃ m thÃ¡ng má»›i hoáº·c má»Ÿ bá»™ dá»¯ liá»‡u Ä‘Ã£ lÆ°u Ä‘á»ƒ báº¯t Ä‘áº§u lÃ m viá»‡c.",
                    Screenshot: startupScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var editorScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 2: LÃ m quen mÃ n hÃ¬nh nháº­p chi tiáº¿t",
                    Description: "ÄÃ¢y lÃ  mÃ n hÃ¬nh chÃ­nh Ä‘á»ƒ nháº­p, sá»­a vÃ  kiá»ƒm tra dá»¯ liá»‡u khÃ¡ch hÃ ng. Báº¡n cÃ³ thá»ƒ chá»‰nh trá»±c tiáº¿p tá»«ng cá»™t trong báº£ng.",
                    Screenshot: editorScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Search, out var searchScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 3: TÃ¬m kiáº¿m vÃ  lá»c nhanh dá»¯ liá»‡u",
                    Description: "Nháº­p tá»« khÃ³a vÃ o Ã´ tÃ¬m kiáº¿m, nháº¥n Enter vÃ  káº¿t há»£p cÃ¡c bá»™ lá»c Ä‘á»ƒ táº­p trung vÃ o cÃ¡c dÃ²ng Ä‘ang cáº§n xá»­ lÃ½.",
                    Screenshot: searchScreenshot));
            }

            var normalizedSteps = RenumberSteps(steps);

            section = new UserGuideSectionItem(
                TabTitle: "LÃ m sao Ä‘á»ƒ báº¯t Ä‘áº§u?",
                Heading: "LÃ m sao Ä‘á»ƒ báº¯t Ä‘áº§u vÃ  lÃ m quen pháº§n má»m?",
                Description: "Flow cÆ¡ báº£n tá»« Trang chá»§ Ä‘áº¿n mÃ n hÃ¬nh nháº­p liá»‡u chÃ­nh.",
                Steps: normalizedSteps);

            return normalizedSteps.Count > 0;
        }

        private static void AddImportGuideSections(
            ICollection<UserGuideSectionItem> sections,
            BitmapSource? startupScreenshot)
        {
            var sampleExcelPath = GetGuideSampleExcelPath();
            if (string.IsNullOrWhiteSpace(sampleExcelPath) || !File.Exists(sampleExcelPath))
            {
                return;
            }

            var ui = new UiService();

            if (TryBuildDetailedImportSection(ui, sampleExcelPath, startupScreenshot, out var detailedImportSection))
            {
                sections.Add(detailedImportSection);
            }

            if (TryBuildCurrentIndexImportSection(ui, sampleExcelPath, startupScreenshot, out var currentIndexImportSection))
            {
                sections.Add(currentIndexImportSection);
            }
        }

        private static bool TryBuildDetailedImportSection(
            UiService ui,
            string sampleExcelPath,
            BitmapSource? startupScreenshot,
            out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            if (startupScreenshot != null)
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 1: Má»Ÿ hoáº·c táº¡o bá»™ dá»¯ liá»‡u cáº§n nháº­p",
                    Description: "Tá»« Trang chá»§, vÃ o bá»™ dá»¯ liá»‡u báº¡n muá»‘n lÃ m viá»‡c trÆ°á»›c khi thá»±c hiá»‡n import Excel chi tiáº¿t.",
                    Screenshot: startupScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var openImportScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 2: VÃ o Tá»‡p -> Import tá»« Excel",
                    Description: "Táº¡i mÃ n hÃ¬nh chÃ­nh, chá»n menu Tá»‡p rá»“i báº¥m 'Import tá»« Excel...' Ä‘á»ƒ má»Ÿ form import dá»¯ liá»‡u Ä‘áº§y Ä‘á»§.",
                    Screenshot: openImportScreenshot));
            }

            if (TryCaptureImportWizardFullDatasetSteps(ui, sampleExcelPath, out var wizardSteps))
            {
                foreach (var step in wizardSteps)
                {
                    steps.Add(step);
                }
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var resultScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 6: Kiá»ƒm tra láº¡i dá»¯ liá»‡u sau khi import",
                    Description: "Sau khi nháº­p xong, rÃ  láº¡i cÃ¡c cá»™t chá»‰ sá»‘, Ä‘Æ¡n giÃ¡, nhÃ³m vÃ  Ä‘á»‹a chá»‰ ngay trÃªn báº£ng chÃ­nh Ä‘á»ƒ trÃ¡nh thiáº¿u dá»¯ liá»‡u.",
                    Screenshot: resultScreenshot));
            }

            var normalizedSteps = RenumberSteps(steps);

            section = new UserGuideSectionItem(
                TabTitle: "LÃ m sao Ä‘á»ƒ nháº­p tá»« Excel?",
                Heading: "LÃ m sao Ä‘á»ƒ nháº­p dá»¯ liá»‡u chi tiáº¿t tá»« Excel?",
                Description: "Flow nháº­p Ä‘áº§y Ä‘á»§ dá»¯ liá»‡u khÃ¡ch hÃ ng vÃ  chá»‰ sá»‘ tá»« file Excel.",
                Steps: normalizedSteps);

            return normalizedSteps.Count > 0;
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
                        StepTitle: "BÆ°á»›c 3: Chá»n file, sheet vÃ  dÃ²ng tiÃªu Ä‘á»",
                        Description:
                            "Chá»n Ä‘Ãºng file .xlsx, Ä‘Ãºng sheet vÃ  Ä‘Ãºng dÃ²ng tiÃªu Ä‘á» Ä‘á»ƒ há»‡ thá»‘ng Ä‘á»c tÃªn cá»™t chÃ­nh xÃ¡c ngay tá»« Ä‘áº§u.",
                        Screenshot: step1Screenshot));
                }

                vm.CurrentStep = 1;
                if (TryCaptureWindowVisual(window, 1180, 780, out var step2Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "BÆ°á»›c 4: GhÃ©p cá»™t dá»¯ liá»‡u",
                        Description:
                            "GhÃ©p tá»«ng cá»™t trong file Excel vá»›i trÆ°á»ng tÆ°Æ¡ng á»©ng trÃªn pháº§n má»m. Báº¯t buá»™c cÃ³ 'TÃªn khÃ¡ch', vÃ  nÃªn map thÃªm Sá»‘ cÃ´ng tÆ¡, Chá»‰ sá»‘ cÅ©, Chá»‰ sá»‘ má»›i, ÄÆ¡n giÃ¡.",
                        Screenshot: step2Screenshot));
                }

                vm.CurrentStep = 2;
                vm.ValidateCommand.Execute(null);
                if (TryCaptureWindowVisual(window, 1180, 780, out var step3Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "BÆ°á»›c 5: Kiá»ƒm tra lá»—i rá»“i nháº­p dá»¯ liá»‡u",
                        Description:
                            "Báº¥m 'Kiá»ƒm tra' Ä‘á»ƒ xem lá»—i vÃ  cáº£nh bÃ¡o. Náº¿u khÃ´ng cÃ²n lá»—i Ä‘á», báº¥m 'Nháº­p dá»¯ liá»‡u' rá»“i hoÃ n táº¥t.",
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

        private static bool TryBuildCurrentIndexImportSection(
            UiService ui,
            string sampleExcelPath,
            BitmapSource? startupScreenshot,
            out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            if (startupScreenshot != null)
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 1: Má»Ÿ bá»™ dá»¯ liá»‡u thÃ¡ng hiá»‡n táº¡i",
                    Description: "Tá»« Trang chá»§, má»Ÿ bá»™ dá»¯ liá»‡u cá»§a thÃ¡ng Ä‘ang lÃ m Ä‘á»ƒ chuáº©n bá»‹ cáº­p nháº­t chá»‰ sá»‘ má»›i.",
                    Screenshot: startupScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var openImportScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 2: Báº¥m Import chá»‰ sá»‘ má»›i",
                    Description: "Táº¡i mÃ n hÃ¬nh chÃ­nh, dÃ¹ng nÃºt 'Import chá»‰ sá»‘ má»›i' hoáº·c menu Tá»‡p -> Import chá»‰ sá»‘ má»›i tá»« Excel...",
                    Screenshot: openImportScreenshot));
            }

            if (TryCaptureImportWizardCurrentIndexSteps(ui, sampleExcelPath, out var wizardSteps))
            {
                foreach (var step in wizardSteps)
                {
                    steps.Add(step);
                }
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var resultScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 6: Xem láº¡i cá»™t Chá»‰ sá»‘ má»›i trÃªn báº£ng",
                    Description: "Sau khi import xong, kiá»ƒm tra láº¡i cÃ¡c dÃ²ng vá»«a cáº­p nháº­t vÃ  lÆ°u Ã½ há»‡ thá»‘ng sáº½ tá»± lÆ°u snapshot thÃ¡ng má»›i.",
                    Screenshot: resultScreenshot));
            }

            var normalizedSteps = RenumberSteps(steps);

            section = new UserGuideSectionItem(
                TabTitle: "LÃ m sao Ä‘á»ƒ nháº­p chá»‰ sá»‘ má»›i?",
                Heading: "LÃ m sao Ä‘á»ƒ nháº­p chá»‰ sá»‘ má»›i tá»« Excel?",
                Description: "Flow cáº­p nháº­t má»™t cá»™t chá»‰ sá»‘ má»›i vÃ  ghÃ©p vá»›i dá»¯ liá»‡u hiá»‡n táº¡i.",
                Steps: normalizedSteps);

            return normalizedSteps.Count > 0;
        }

        private static bool TryCaptureImportWizardCurrentIndexSteps(
            UiService ui,
            string sampleExcelPath,
            out IReadOnlyList<UserGuideStepItem> steps)
        {
            var captured = new List<UserGuideStepItem>();
            CurrentIndexImportWindow? window = null;

            try
            {
                var vm = new ImportWizardViewModel(ui, sampleExcelPath, ImportWizardViewModel.ImportWizardMode.CurrentIndexOneColumn);
                window = new CurrentIndexImportWindow
                {
                    DataContext = vm
                };

                vm.CurrentStep = 0;
                if (TryCaptureWindowVisual(window, 1180, 780, out var step1Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "BÆ°á»›c 3: Chá»n file, sheet vÃ  dÃ²ng tiÃªu Ä‘á»",
                        Description:
                            "Chá»n file Excel chá»©a chá»‰ sá»‘ má»›i, Ä‘Ãºng sheet vÃ  Ä‘Ãºng dÃ²ng tiÃªu Ä‘á» Ä‘á»ƒ há»‡ thá»‘ng nháº­n biáº¿t cá»™t chÃ­nh xÃ¡c.",
                        Screenshot: step1Screenshot));
                }

                vm.CurrentStep = 1;
                if (TryCaptureWindowVisual(window, 1180, 780, out var step2Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "BÆ°á»›c 4: Chá»n cá»™t Chá»‰ sá»‘ má»›i vÃ  khÃ³a ghÃ©p",
                        Description:
                            "Báº¯t buá»™c map 'Chá»‰ sá»‘ má»›i'. Kiá»ƒm tra láº¡i cÃ¡c khÃ³a ghÃ©p nhÆ° Sá»‘ cÃ´ng tÆ¡, Sá»‘ thá»© tá»±, TÃªn khÃ¡ch; há»‡ thá»‘ng cÃ³ tá»± Ä‘oÃ¡n nhÆ°ng báº¡n váº«n cÃ³ thá»ƒ sá»­a tay.",
                        Screenshot: step2Screenshot));
                }

                vm.CurrentStep = 2;
                vm.ValidateCommand.Execute(null);
                if (TryCaptureWindowVisual(window, 1180, 780, out var step3Screenshot))
                {
                    captured.Add(new UserGuideStepItem(
                        StepTitle: "BÆ°á»›c 5: Kiá»ƒm tra rá»“i nháº­p chá»‰ sá»‘",
                        Description:
                            "Báº¥m 'Kiá»ƒm tra' Ä‘á»ƒ rÃ  lá»—i/cáº£nh bÃ¡o, sau Ä‘Ã³ import. Káº¿t quáº£ cáº­p nháº­t sáº½ Ä‘Æ°á»£c ghi láº¡i vÃ  tá»± lÆ°u bá»™ dá»¯ liá»‡u.",
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

        private static bool TryBuildFastEntrySection(BitmapSource? startupScreenshot, out UserGuideSectionItem section)
        {
            var steps = new List<UserGuideStepItem>();

            if (startupScreenshot != null)
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 1: Má»Ÿ bá»™ dá»¯ liá»‡u cáº§n nháº­p nhanh",
                    Description: "Báº¯t Ä‘áº§u tá»« Trang chá»§ vÃ  má»Ÿ bá»™ dá»¯ liá»‡u cá»§a thÃ¡ng Ä‘ang lÃ m Ä‘á»ƒ chuyá»ƒn sang cháº¿ Ä‘á»™ nháº­p nhanh.",
                    Screenshot: startupScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var detailScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 2: VÃ o mÃ n hÃ¬nh nháº­p chi tiáº¿t",
                    Description: "Táº¡i mÃ n hÃ¬nh chÃ­nh, kiá»ƒm tra Ä‘Ãºng ká»³ tÃ­nh vÃ  dá»¯ liá»‡u khÃ¡ch trÆ°á»›c khi chuyá»ƒn sang cháº¿ Ä‘á»™ nháº­p nhanh.",
                    Screenshot: detailScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.FastEntry, out var fastEntryScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 3: Báº­t cháº¿ Ä‘á»™ Nháº­p nhanh",
                    Description: "Gáº¡t nÃºt cháº¿ Ä‘á»™ sang 'Nháº­p nhanh'. Khi Ä‘Ã³ báº£ng sáº½ táº­p trung thao tÃ¡c vÃ o cá»™t 'Chá»‰ sá»‘ má»›i'.",
                    Screenshot: fastEntryScreenshot));

                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 4: Nháº­p chá»‰ sá»‘ má»›i liÃªn tá»¥c",
                    Description: "Chá»‰ nháº­p á»Ÿ cá»™t 'Chá»‰ sá»‘ má»›i', nháº¥n Enter Ä‘á»ƒ nháº£y xuá»‘ng dÃ²ng tiáº¿p theo. DÃ¹ng bá»™ lá»c 'Thiáº¿u', 'Cáº£nh bÃ¡o', 'Lá»—i' Ä‘á»ƒ rÃ  nhanh.",
                    Screenshot: fastEntryScreenshot));
            }

            if (TryCaptureMainWindowWithSampleData(MainWindowGuideCaptureMode.Detail, out var backToDetailScreenshot))
            {
                steps.Add(new UserGuideStepItem(
                    StepTitle: "BÆ°á»›c 5: Quay láº¡i Chi tiáº¿t khi cáº§n sá»­a sÃ¢u",
                    Description: "Sau khi nháº­p nhanh xong, chuyá»ƒn láº¡i 'Chi tiáº¿t' Ä‘á»ƒ sá»­a cÃ¡c cá»™t khÃ¡c nhÆ° nhÃ³m, Ä‘á»‹a chá»‰, TBA hoáº·c sá»‘ cÃ´ng tÆ¡.",
                    Screenshot: backToDetailScreenshot));
            }

            var normalizedSteps = RenumberSteps(steps);

            section = new UserGuideSectionItem(
                TabTitle: "LÃ m sao Ä‘á»ƒ nháº­p nhanh?",
                Heading: "LÃ m sao Ä‘á»ƒ nháº­p nhanh chá»‰ sá»‘ trÃªn mÃ n hÃ¬nh chÃ­nh?",
                Description: "Flow thao tÃ¡c trá»±c tiáº¿p trÃªn báº£ng Ä‘á»ƒ cáº­p nháº­t chá»‰ sá»‘ nhanh theo tá»«ng dÃ²ng.",
                Steps: normalizedSteps);

            return normalizedSteps.Count > 0;
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

        private static IReadOnlyList<UserGuideStepItem> RenumberSteps(IEnumerable<UserGuideStepItem> steps)
        {
            var result = new List<UserGuideStepItem>();
            var stepNumber = 1;

            foreach (var step in steps)
            {
                var title = step.StepTitle ?? string.Empty;
                if (title.StartsWith("BÆ°á»›c ", StringComparison.OrdinalIgnoreCase))
                {
                    var colonIndex = title.IndexOf(':');
                    if (colonIndex >= 0 && colonIndex < title.Length - 1)
                    {
                        title = $"BÆ°á»›c {stepNumber}:{title[(colonIndex + 1)..]}";
                    }
                    else
                    {
                        title = $"BÆ°á»›c {stepNumber}: {title}";
                    }
                }
                else
                {
                    title = $"BÆ°á»›c {stepNumber}: {title}";
                }

                result.Add(step with { StepTitle = title });
                stepNumber++;
            }

            return result;
        }

        private enum MainWindowGuideCaptureMode
        {
            Detail,
            Search,
            FastEntry
        }

        private static void PopulateDemoCustomers(MainWindowViewModel vm)
        {
            vm.PeriodLabel = $"ThÃ¡ng {DateTime.Now.Month:00}/{DateTime.Now.Year}";
            vm.InvoiceIssuer = "Nguyá»…n VÄƒn A";

            vm.Customers.Clear();

            vm.Customers.Add(new Customer
            {
                SequenceNumber = 1,
                Name = "PhÃ²ng A101",
                GroupName = "Khu A",
                Category = "Sinh hoáº¡t",
                Address = "TÃ²a A - Táº§ng 1",
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
                Name = "PhÃ²ng A102",
                GroupName = "Khu A",
                Category = "Sinh hoáº¡t",
                Address = "TÃ²a A - Táº§ng 1",
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
                Name = "PhÃ²ng B201",
                GroupName = "Khu B",
                Category = "Dá»‹ch vá»¥",
                Address = "TÃ²a B - Táº§ng 2",
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
                Name = "PhÃ²ng B202",
                GroupName = "Khu B",
                Category = "Dá»‹ch vá»¥",
                Address = "TÃ²a B - Táº§ng 2",
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

