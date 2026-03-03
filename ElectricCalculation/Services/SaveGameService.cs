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

        public static bool IsSharedSyncEnabled()
        {
            return SharedSnapshotDatabaseService.IsEnabled();
        }

        public static string GetSaveRootDirectory()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "ElectricCalculation", "Saves");
        }

        public static void SyncFromSharedStoreIfEnabled()
        {
            var root = GetSaveRootDirectory();
            Directory.CreateDirectory(root);

            if (!IsSharedSyncEnabled())
            {
                return;
            }

            var records = SharedSnapshotDatabaseService.ListSnapshots(maxCount: 5000);
            foreach (var record in records)
            {
                SharedSnapshotDatabaseService.MaterializeToLocal(root, record);
            }
        }

        public static string SaveSnapshot(string periodLabel, IEnumerable<Customer> customers, string? snapshotName = null)
        {
            var safePeriod = MakeSafeFileName(periodLabel ?? string.Empty);
            if (string.IsNullOrWhiteSpace(safePeriod))
            {
                safePeriod = "UnknownPeriod";
            }

            var folder = Path.Combine(GetSaveRootDirectory(), safePeriod);
            Directory.CreateDirectory(folder);

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var safeSnapshotName = MakeSafeFileName(snapshotName ?? string.Empty);
            var fileName = string.IsNullOrWhiteSpace(safeSnapshotName)
                ? $"{stamp} - {safePeriod}.json"
                : $"{stamp} - {safePeriod} - {safeSnapshotName}.json";
            var path = Path.Combine(folder, fileName);

            ProjectFileService.Save(path, periodLabel ?? string.Empty, customers);
            TrimOldSnapshots(folder);

            if (IsSharedSyncEnabled())
            {
                var savedAtUtc = File.GetLastWriteTimeUtc(path);
                SharedSnapshotDatabaseService.UpsertSnapshot(
                    saveRootDirectory: GetSaveRootDirectory(),
                    snapshotPath: path,
                    periodLabel: periodLabel ?? string.Empty,
                    snapshotName: snapshotName,
                    savedAtUtc: savedAtUtc);
            }

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

            if (IsSharedSyncEnabled())
            {
                var records = SharedSnapshotDatabaseService.ListSnapshots(maxCount);
                var synced = new List<SnapshotInfo>(records.Count);

                foreach (var record in records)
                {
                    var localPath = SharedSnapshotDatabaseService.MaterializeToLocal(root, record);
                    synced.Add(new SnapshotInfo(
                        localPath,
                        record.PeriodLabel,
                        record.SnapshotName,
                        record.SavedAtUtc.ToLocalTime()));
                }

                return synced;
            }

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
            if (!IsSharedSyncEnabled() || string.IsNullOrWhiteSpace(snapshotPath))
            {
                return;
            }

            if (!File.Exists(snapshotPath))
            {
                throw new WarningException("Khong tim thay file snapshot de dong bo.");
            }

            var snapshotName = TryReadSnapshotNameFromFileName(Path.GetFileName(snapshotPath));
            SharedSnapshotDatabaseService.UpsertSnapshot(
                saveRootDirectory: GetSaveRootDirectory(),
                snapshotPath: snapshotPath,
                periodLabel: periodLabel ?? string.Empty,
                snapshotName: snapshotName,
                savedAtUtc: File.GetLastWriteTimeUtc(snapshotPath));
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

                if (IsSharedSyncEnabled())
                {
                    SharedSnapshotDatabaseService.DeleteSnapshot(GetSaveRootDirectory(), fullPath);
                }

                if (File.Exists(fullPath))
                {
                    File.Delete(fullPath);
                    return true;
                }

                if (IsSharedSyncEnabled())
                {
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
