#nullable enable
using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace AutoJMS
{
    public class AppSettings
    {
        public double ZoomFactor { get; set; } = 1.0;
        public string DefaultUrl { get; set; } = AppConfig.Current.JmsBaseUrl.TrimEnd('/');
        public string LastAuthToken { get; set; } = "";
        public string DownloadFolder { get; set; } = "";
        public string DefaultSheet { get; set; } = "DKCH";
        public bool UseSheetByDefault { get; set; } = false;
        public bool AutoRefreshToken { get; set; } = true;
        public string LastMode { get; set; } = "DKCH1";
        public int DefaultRowCount { get; set; } = 1;
        public string PrinterName { get; set; } = "";
        public int PaperWidth { get; set; } = 762;  
        public int PaperHeight { get; set; } = 762; 
        
        public string LicenseKey { get; set; } = ""; 
    }

    public static class SettingsManager
    {
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoJMS.config.enc");
        private static readonly string LegacyConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoJMS.json");
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };

        public static AppSettings Load()
        {
            var settings = new AppSettings();

            if (File.Exists(ConfigPath))
            {
                try
                {
                    string protectedJson = File.ReadAllText(ConfigPath);
                    string json = SecureConfigCrypto.UnprotectString(protectedJson, BuildSettingsSecret());
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        var loaded = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
                        return Normalize(loaded ?? settings);
                    }
                }
                catch
                {
                    return settings;
                }
            }

            AppSettings? legacySettings = TryLoadLegacyPlainJson();
            if (legacySettings != null)
            {
                Save(legacySettings);
                return Normalize(legacySettings);
            }

            Save(settings);
            return Normalize(settings);
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                string json = JsonSerializer.Serialize(Normalize(settings), JsonOptions);
                string protectedJson = SecureConfigCrypto.ProtectString(json, BuildSettingsSecret());
                File.WriteAllText(ConfigPath, protectedJson);
            }
            catch { }
        }

        public static async Task SaveAsync(AppSettings settings)
        {
            try
            {
                string json = JsonSerializer.Serialize(Normalize(settings), JsonOptions);
                string protectedJson = SecureConfigCrypto.ProtectString(json, BuildSettingsSecret());
                await File.WriteAllTextAsync(ConfigPath, protectedJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lưu cấu hình: {ex.Message}");
            }
        }

        private static AppSettings Normalize(AppSettings settings)
        {
            if (string.IsNullOrWhiteSpace(settings.DefaultUrl))
                settings.DefaultUrl = AppConfig.Current.JmsBaseUrl.TrimEnd('/');
            return settings;
        }

        private static AppSettings? TryLoadLegacyPlainJson()
        {
            if (!File.Exists(LegacyConfigPath)) return null;

            try
            {
                string json = File.ReadAllText(LegacyConfigPath);
                return JsonSerializer.Deserialize<AppSettings>(json, JsonOptions);
            }
            catch
            {
                return null;
            }
        }

        private static string BuildSettingsSecret()
            => $"{Environment.MachineName}|{Environment.UserName}|AutoJMS|settings";
    }
}