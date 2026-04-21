using System;
using System.IO;
using System.Text.Json;

namespace AutoJMS
{
    public class AppSettings
    {
        public double ZoomFactor { get; set; } = 1.0;
        public string DefaultUrl { get; set; } = "https://jms.jtexpress.vn";
        public string LastAuthToken { get; set; } = "";
        public string DownloadFolder { get; set; } = "";
        public string DefaultSheet { get; set; } = "DKCH";
        public bool UseSheetByDefault { get; set; } = false;
        public bool AutoRefreshToken { get; set; } = true;
        public string LastMode { get; set; } = "DKCH1";
        public int DefaultRowCount { get; set; } = 1;
        //...........................
    }

    public static class SettingsManager
    {
        private static readonly string ConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoJMS.json");

        public static AppSettings Load()
        {
            var settings = new AppSettings();

            if (!File.Exists(ConfigPath))
            {
                Save(settings); // tạo file mặc định
                return settings;
            }

            try
            {
                string json = File.ReadAllText(ConfigPath);
                var loaded = JsonSerializer.Deserialize<AppSettings>(json, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });
                return loaded ?? settings;
            }
            catch
            {
                return settings;
            }
        }

        public static void Save(AppSettings settings)
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                string json = JsonSerializer.Serialize(settings, options);
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }
        public static async Task SaveAsync(AppSettings settings)
        {
            try
            {
                var options = new JsonSerializerOptions { WriteIndented = true };
                // Sử dụng FileStream với FileMode.Create để ghi file bất đồng bộ
                using (FileStream fs = new FileStream(ConfigPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true))
                {
                    await JsonSerializer.SerializeAsync(fs, settings, options);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi lưu cấu hình: {ex.Message}");
            }



        }
    }
}