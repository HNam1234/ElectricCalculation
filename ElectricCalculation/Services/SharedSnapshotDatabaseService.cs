using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using ElectricCalculation.Models;
using Microsoft.Data.Sqlite;

namespace ElectricCalculation.Services
{
    internal static class SharedSnapshotDatabaseService
    {
        private const string SharedDbPathEnvVar = "ELECTRIC_CALC_SYNC_DB_PATH";
        private const string TableName = "shared_snapshots";

        internal sealed record SharedSnapshotRecord(
            string RelativePath,
            string PeriodLabel,
            string? SnapshotName,
            DateTime SavedAtUtc,
            string ContentJson);

        private static readonly object InitLock = new();
        private static string? _initializedForPath;

        public static bool IsEnabled()
        {
            return !string.IsNullOrWhiteSpace(GetDatabasePath());
        }

        public static string? GetDatabasePath()
        {
            var envPath = Environment.GetEnvironmentVariable(SharedDbPathEnvVar);
            if (!string.IsNullOrWhiteSpace(envPath))
            {
                return envPath.Trim();
            }

            var settings = AppSettingsService.Load();
            if (!string.IsNullOrWhiteSpace(settings.SharedSyncDatabasePath))
            {
                return settings.SharedSyncDatabasePath.Trim();
            }

            return null;
        }

        public static void EnsureInitialized()
        {
            var dbPath = GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                return;
            }

            lock (InitLock)
            {
                if (string.Equals(_initializedForPath, dbPath, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                var dir = Path.GetDirectoryName(dbPath);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                using var connection = new SqliteConnection(BuildConnectionString(dbPath));
                connection.Open();

                using var command = connection.CreateCommand();
                command.CommandText =
                    $"CREATE TABLE IF NOT EXISTS {TableName} (" +
                    "relative_path TEXT PRIMARY KEY," +
                    "period_label TEXT NOT NULL," +
                    "snapshot_name TEXT NULL," +
                    "saved_at_utc TEXT NOT NULL," +
                    "content_json TEXT NOT NULL," +
                    "updated_at_utc TEXT NOT NULL" +
                    ");" +
                    $"CREATE INDEX IF NOT EXISTS IX_{TableName}_saved_at ON {TableName}(saved_at_utc DESC);";
                command.ExecuteNonQuery();

                _initializedForPath = dbPath;
            }
        }

        public static IReadOnlyList<SharedSnapshotRecord> ListSnapshots(int maxCount)
        {
            if (maxCount <= 0)
            {
                return Array.Empty<SharedSnapshotRecord>();
            }

            var dbPath = GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                return Array.Empty<SharedSnapshotRecord>();
            }

            EnsureInitialized();

            using var connection = new SqliteConnection(BuildConnectionString(dbPath));
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText =
                $"SELECT relative_path, period_label, snapshot_name, saved_at_utc, content_json " +
                $"FROM {TableName} " +
                "ORDER BY saved_at_utc DESC " +
                "LIMIT $limit;";
            command.Parameters.AddWithValue("$limit", maxCount);

            using var reader = command.ExecuteReader();
            var result = new List<SharedSnapshotRecord>();
            while (reader.Read())
            {
                var relativePath = reader.GetString(0);
                var periodLabel = reader.GetString(1);
                var snapshotName = reader.IsDBNull(2) ? null : reader.GetString(2);
                var savedAtUtcText = reader.GetString(3);
                var contentJson = reader.GetString(4);
                var savedAtUtc = ParseUtc(sanitized: savedAtUtcText);

                result.Add(new SharedSnapshotRecord(
                    RelativePath: relativePath,
                    PeriodLabel: periodLabel,
                    SnapshotName: snapshotName,
                    SavedAtUtc: savedAtUtc,
                    ContentJson: contentJson));
            }

            return result;
        }

        public static void UpsertSnapshot(
            string saveRootDirectory,
            string snapshotPath,
            string periodLabel,
            string? snapshotName,
            DateTime savedAtUtc)
        {
            var dbPath = GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                return;
            }

            if (!File.Exists(snapshotPath))
            {
                throw new WarningException("Không tìm thấy snapshot để đồng bộ lên database dùng chung.");
            }

            var relativePath = TryBuildRelativePath(saveRootDirectory, snapshotPath);
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                throw new WarningException("Snapshot nằm ngoài thư mục Saves, không thể đồng bộ dùng chung.");
            }

            EnsureInitialized();

            var contentJson = File.ReadAllText(snapshotPath);
            var savedUtc = savedAtUtc.Kind == DateTimeKind.Utc
                ? savedAtUtc
                : savedAtUtc.ToUniversalTime();
            var nowUtc = DateTime.UtcNow;

            using var connection = new SqliteConnection(BuildConnectionString(dbPath));
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText =
                $"INSERT INTO {TableName}(relative_path, period_label, snapshot_name, saved_at_utc, content_json, updated_at_utc) " +
                "VALUES($relative_path, $period_label, $snapshot_name, $saved_at_utc, $content_json, $updated_at_utc) " +
                "ON CONFLICT(relative_path) DO UPDATE SET " +
                "period_label = excluded.period_label, " +
                "snapshot_name = excluded.snapshot_name, " +
                "saved_at_utc = excluded.saved_at_utc, " +
                "content_json = excluded.content_json, " +
                "updated_at_utc = excluded.updated_at_utc;";
            command.Parameters.AddWithValue("$relative_path", relativePath);
            command.Parameters.AddWithValue("$period_label", periodLabel ?? string.Empty);
            command.Parameters.AddWithValue("$snapshot_name", (object?)snapshotName ?? DBNull.Value);
            command.Parameters.AddWithValue("$saved_at_utc", savedUtc.ToString("O", CultureInfo.InvariantCulture));
            command.Parameters.AddWithValue("$content_json", contentJson);
            command.Parameters.AddWithValue("$updated_at_utc", nowUtc.ToString("O", CultureInfo.InvariantCulture));
            command.ExecuteNonQuery();
        }

        public static void DeleteSnapshot(string saveRootDirectory, string snapshotPath)
        {
            var dbPath = GetDatabasePath();
            if (string.IsNullOrWhiteSpace(dbPath))
            {
                return;
            }

            var relativePath = TryBuildRelativePath(saveRootDirectory, snapshotPath);
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return;
            }

            EnsureInitialized();

            using var connection = new SqliteConnection(BuildConnectionString(dbPath));
            connection.Open();

            using var command = connection.CreateCommand();
            command.CommandText = $"DELETE FROM {TableName} WHERE relative_path = $relative_path;";
            command.Parameters.AddWithValue("$relative_path", relativePath);
            command.ExecuteNonQuery();
        }

        public static string MaterializeToLocal(string saveRootDirectory, SharedSnapshotRecord record)
        {
            var fullPath = BuildSafeLocalPath(saveRootDirectory, record.RelativePath);
            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var shouldWrite = true;
            if (File.Exists(fullPath))
            {
                try
                {
                    var existingContent = File.ReadAllText(fullPath);
                    shouldWrite = !string.Equals(existingContent, record.ContentJson, StringComparison.Ordinal);
                }
                catch
                {
                    shouldWrite = true;
                }
            }

            if (shouldWrite)
            {
                File.WriteAllText(fullPath, record.ContentJson);
            }

            File.SetLastWriteTimeUtc(fullPath, record.SavedAtUtc);
            return fullPath;
        }

        private static string BuildConnectionString(string dbPath)
        {
            var builder = new SqliteConnectionStringBuilder
            {
                DataSource = dbPath,
                Mode = SqliteOpenMode.ReadWriteCreate,
                Cache = SqliteCacheMode.Shared
            };

            return builder.ToString();
        }

        private static string BuildSafeLocalPath(string saveRootDirectory, string relativePath)
        {
            var root = Path.GetFullPath(saveRootDirectory);
            var normalized = NormalizeRelativePath(relativePath);
            var combined = Path.Combine(root, normalized.Replace('/', Path.DirectorySeparatorChar));
            var full = Path.GetFullPath(combined);

            if (!IsUnderRoot(full, root))
            {
                throw new WarningException("Phát hiện dữ liệu đường dẫn không hợp lệ từ database dùng chung.");
            }

            return full;
        }

        private static string? TryBuildRelativePath(string saveRootDirectory, string snapshotPath)
        {
            var root = Path.GetFullPath(saveRootDirectory);
            var fullPath = Path.GetFullPath(snapshotPath);

            if (!IsUnderRoot(fullPath, root))
            {
                return null;
            }

            var relative = Path.GetRelativePath(root, fullPath);
            if (string.IsNullOrWhiteSpace(relative) ||
                relative.StartsWith("..", StringComparison.Ordinal))
            {
                return null;
            }

            return NormalizeRelativePath(relative);
        }

        private static string NormalizeRelativePath(string relativePath)
        {
            var normalized = relativePath
                .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
                .Replace(Path.DirectorySeparatorChar, '/')
                .Trim();

            while (normalized.StartsWith("/", StringComparison.Ordinal))
            {
                normalized = normalized[1..];
            }

            return normalized;
        }

        private static bool IsUnderRoot(string fullPath, string rootPath)
        {
            var root = rootPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var rootWithSeparator = root + Path.DirectorySeparatorChar;

            return fullPath.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(fullPath, root, StringComparison.OrdinalIgnoreCase);
        }

        private static DateTime ParseUtc(string sanitized)
        {
            if (DateTime.TryParse(
                    sanitized,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var value))
            {
                return value;
            }

            return DateTime.UtcNow;
        }
    }
}
