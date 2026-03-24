using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class SaveGameService
    {
        private const int MaxSnapshotsPerPeriod = 30;

        public sealed record SnapshotInfo(string Path, string PeriodLabel, string? SnapshotName, DateTime SavedAt);

        public static bool IsSharedSavesFolderEnabled()
        {
            try
            {
                var settings = AppSettingsService.Load();
                return !string.IsNullOrWhiteSpace(settings.SharedSavesDirectory);
            }
            catch
            {
                return false;
            }
        }

        public static bool IsSharedSyncEnabled()
        {
            // SQL shared-sync has been reverted/disabled.
            return false;
        }

        public static string GetSaveRootDirectory()
        {
            try
            {
                var settings = AppSettingsService.Load();
                if (!string.IsNullOrWhiteSpace(settings.SharedSavesDirectory))
                {
                    return settings.SharedSavesDirectory.Trim();
                }
            }
            catch
            {
                // Ignore settings read failures.
            }

            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "ElectricCalculation", "Saves");
        }

        public static string SaveSnapshot(string periodLabel, IEnumerable<Customer> customers, string? snapshotName = null)
        {
            var safePeriod = MakeSafeFileName(periodLabel ?? string.Empty);
            if (string.IsNullOrWhiteSpace(safePeriod))
            {
                safePeriod = "UnknownPeriod";
            }

            var root = GetSaveRootDirectory();
            var folder = IsSharedSavesFolderEnabled()
                ? Path.Combine(root, safePeriod, MakeSafeFileName(GetClientId()))
                : Path.Combine(root, safePeriod);
            Directory.CreateDirectory(folder);

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            stamp = $"{stamp}_{Random.Shared.Next(1000, 10000)}";
            var safeSnapshotName = MakeSafeFileName(snapshotName ?? string.Empty);
            var fileName = string.IsNullOrWhiteSpace(safeSnapshotName)
                ? $"{stamp} - {safePeriod}.json"
                : $"{stamp} - {safePeriod} - {safeSnapshotName}.json";
            var path = Path.Combine(folder, fileName);

            ProjectFileService.Save(path, periodLabel ?? string.Empty, customers);

            // In shared-folder mode, snapshots are per-client subfolder, so trimming stays safe.
            TrimOldSnapshots(folder);

            return path;
        }

        public static IReadOnlyList<SnapshotInfo> ListSnapshots(int maxCount = 50)
        {
            if (maxCount <= 0)
            {
                return Array.Empty<SnapshotInfo>();
            }

            var root = GetSaveRootDirectory();
            Directory.CreateDirectory(root);

            if (!Directory.Exists(root))
            {
                return Array.Empty<SnapshotInfo>();
            }

            var files = Directory
                .EnumerateFiles(root, "*.json", SearchOption.AllDirectories)
                .Select(path => new FileInfo(path))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .Take(maxCount)
                .ToList();

            var result = new List<SnapshotInfo>(files.Count);
            foreach (var file in files)
            {
                var periodLabel = TryReadPeriodLabel(file.FullName) ?? Path.GetFileNameWithoutExtension(file.Name);
                var snapshotName = TryReadSnapshotNameFromFileName(file.Name);
                result.Add(new SnapshotInfo(file.FullName, periodLabel, snapshotName, file.LastWriteTime));
            }

            return result;
        }

        public static void SyncSnapshotFileToSharedStore(string snapshotPath, string periodLabel)
        {
            // SQL shared-sync has been reverted/disabled.
        }

        public static bool TryDeleteSnapshot(string snapshotPath, out string? error)
        {
            error = null;

            if (string.IsNullOrWhiteSpace(snapshotPath))
            {
                error = "Duong dan snapshot trong.";
                return false;
            }

            try
            {
                if (!string.Equals(Path.GetExtension(snapshotPath), ".json", StringComparison.OrdinalIgnoreCase))
                {
                    error = "Snapshot khong hop le (khong phai file .json).";
                    return false;
                }

                var root = Path.GetFullPath(GetSaveRootDirectory())
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                var rootWithSeparator = root + Path.DirectorySeparatorChar;
                var fullPath = Path.GetFullPath(snapshotPath);

                if (!fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(fullPath, root, StringComparison.OrdinalIgnoreCase))
                {
                    error = "Chi cho phep xoa snapshot trong thu muc snapshot.";
                    return false;
                }

                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                    return true;
                }

                error = "Snapshot khong con ton tai.";
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void TrimOldSnapshots(string folder)
        {
            try
            {
                var files = Directory
                    .EnumerateFiles(folder, "*.json", SearchOption.TopDirectoryOnly)
                    .Select(path => new FileInfo(path))
                    .OrderByDescending(f => f.LastWriteTimeUtc)
                    .ToList();

                foreach (var file in files.Skip(MaxSnapshotsPerPeriod))
                {
                    file.Delete();
                }
            }
            catch
            {
                // Best-effort cleanup; snapshot save should still succeed.
            }
        }

        private static string GetClientId()
        {
            try
            {
                var machine = Environment.MachineName?.Trim() ?? "UnknownMachine";
                var user = Environment.UserName?.Trim() ?? "UnknownUser";
                return $"{machine}_{user}";
            }
            catch
            {
                return "UnknownClient";
            }
        }

        private static string? TryReadPeriodLabel(string filePath)
        {
            try
            {
                using var stream = File.OpenRead(filePath);
                using var doc = JsonDocument.Parse(stream);

                if (doc.RootElement.ValueKind != JsonValueKind.Object)
                {
                    return null;
                }

                if (doc.RootElement.TryGetProperty("PeriodLabel", out var element) &&
                    element.ValueKind == JsonValueKind.String)
                {
                    return element.GetString();
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string? TryReadSnapshotNameFromFileName(string fileName)
        {
            try
            {
                var baseName = Path.GetFileNameWithoutExtension(fileName);
                if (string.IsNullOrWhiteSpace(baseName))
                {
                    return null;
                }

                var parts = baseName.Split(new[] { " - " }, StringSplitOptions.None);
                if (parts.Length < 3)
                {
                    return null;
                }

                return string.Join(" - ", parts.Skip(2)).Trim();
            }
            catch
            {
                return null;
            }
        }

        private static string MakeSafeFileName(string name)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name.Trim();
        }
    }
}
