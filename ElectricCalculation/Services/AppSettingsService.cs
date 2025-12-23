using System;
using System.IO;
using System.Text.Json;
using ElectricCalculation.Models;

namespace ElectricCalculation.Services
{
    public static class AppSettingsService
    {
        private const string SettingsFileName = "settings.json";

        public static string GetSettingsDirectory()
        {
            var documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(documents, "ElectricCalculation");
        }

        public static string GetSettingsPath()
        {
            return Path.Combine(GetSettingsDirectory(), SettingsFileName);
        }

        public static AppSettings Load()
        {
            try
            {
                var path = GetSettingsPath();
                if (!File.Exists(path))
                {
                    return new AppSettings();
                }

                var json = File.ReadAllText(path);
                var settings = JsonSerializer.Deserialize<AppSettings>(
                    json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                return settings ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public static void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            var dir = GetSettingsDirectory();
            Directory.CreateDirectory(dir);

            var path = GetSettingsPath();
            var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
    }
}

