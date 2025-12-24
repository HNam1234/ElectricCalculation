using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace ElectricCalculation.Services
{
    public static class PinnedDatasetService
    {
        private const string FileName = "pinned_datasets.json";

        private sealed class Payload
        {
            public Dictionary<string, DateTime> Pins { get; set; } = new();
        }

        private static string GetPinsPath()
        {
            var dir = AppSettingsService.GetSettingsDirectory();
            return Path.Combine(dir, FileName);
        }

        public static Dictionary<string, DateTime> LoadPins()
        {
            try
            {
                var path = GetPinsPath();
                if (!File.Exists(path))
                {
                    return new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
                }

                var json = File.ReadAllText(path);
                var payload = JsonSerializer.Deserialize<Payload>(
                    json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                return NormalizePins(payload?.Pins);
            }
            catch
            {
                return new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public static void SavePins(Dictionary<string, DateTime> pins)
        {
            var normalized = NormalizePins(pins);

            var dir = AppSettingsService.GetSettingsDirectory();
            Directory.CreateDirectory(dir);

            var path = GetPinsPath();
            var payload = new Payload { Pins = normalized };
            var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }

        private static Dictionary<string, DateTime> NormalizePins(Dictionary<string, DateTime>? pins)
        {
            var normalized = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            if (pins == null || pins.Count == 0)
            {
                return normalized;
            }

            foreach (var (path, pinnedAt) in pins)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }

                var fullPath = path.Trim();
                normalized[fullPath] = pinnedAt.Kind == DateTimeKind.Unspecified
                    ? DateTime.SpecifyKind(pinnedAt, DateTimeKind.Utc)
                    : pinnedAt.ToUniversalTime();
            }

            return normalized;
        }

        public static bool TryCleanupMissingPins(Dictionary<string, DateTime> pins, IEnumerable<string> existingSnapshotPaths)
        {
            if (pins == null || pins.Count == 0)
            {
                return false;
            }

            var existing = new HashSet<string>(
                existingSnapshotPaths ?? Array.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var removedAny = false;
            foreach (var key in pins.Keys.ToList())
            {
                if (!existing.Contains(key))
                {
                    pins.Remove(key);
                    removedAny = true;
                }
            }

            return removedAny;
        }
    }
}

