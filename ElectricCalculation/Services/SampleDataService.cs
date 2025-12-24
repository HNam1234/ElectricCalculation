using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace ElectricCalculation.Services
{
    public static class SampleDataService
    {
        private const string SeedMarkerFileName = "sample_seed.json";

        private sealed class SeedMarker
        {
            public DateTime SeededAtUtc { get; set; }
            public string SourceExcelFile { get; set; } = string.Empty;
            public string SnapshotPath { get; set; } = string.Empty;
        }

        public static bool TrySeedJune2025SampleSnapshotOnce()
        {
            try
            {
                var markerPath = GetSeedMarkerPath();
                if (File.Exists(markerPath))
                {
                    return false;
                }

                var excelPath = GetJune2025SampleExcelPath();
                if (!File.Exists(excelPath))
                {
                    return false;
                }

                var customers = ExcelImportService.ImportFromFile(excelPath, out _);

                var periodLabel = "06/2025";
                var snapshotName = "Bang tong hop thu thang 06 2025 - Sample";
                var snapshotPath = SaveGameService.SaveSnapshot(periodLabel, customers, snapshotName);

                var pins = PinnedDatasetService.LoadPins();
                pins[snapshotPath] = DateTime.UtcNow;
                PinnedDatasetService.SavePins(pins);

                Directory.CreateDirectory(AppSettingsService.GetSettingsDirectory());
                var marker = new SeedMarker
                {
                    SeededAtUtc = DateTime.UtcNow,
                    SourceExcelFile = Path.GetFileName(excelPath),
                    SnapshotPath = snapshotPath
                };
                var json = JsonSerializer.Serialize(marker, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(markerPath, json);

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string GetSeedMarkerPath()
        {
            return Path.Combine(AppSettingsService.GetSettingsDirectory(), SeedMarkerFileName);
        }

        private static string GetJune2025SampleExcelPath()
        {
            return Path.Combine(AppContext.BaseDirectory, "SampleData", "Bang_tong_hop_thu_thang_06_2025.xlsx");
        }
    }
}
