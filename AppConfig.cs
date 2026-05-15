#nullable enable
using System;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace AutoJMS
{
    public sealed class AppRuntimeConfig
    {
        public string FirebaseUrl { get; set; } = "";
        public string FirebaseDatabaseSecret { get; set; } = "";
        public string DataSpreadsheetId { get; set; } = "";
        public string GoogleSheetUrl { get; set; } = "";
        public string LicenseSpreadsheetId { get; set; } = "";
        public string AppsScriptUrl { get; set; } = "";
        public string GoogleServiceAccountJson { get; set; } = "";
        public string GoogleCredentialPath { get; set; } = "service_account.json";
        public string JmsBaseUrl { get; set; } = "https://jms.jtexpress.vn";
        public string JmsApiBaseUrl { get; set; } = "https://jmsgw.jtexpress.vn";
        public string InternetCheckUrl { get; set; } = "http://clients3.google.com/generate_204";
        public string UpdateXmlUrl { get; set; } = "";
        public string ActionSiteCode { get; set; } = "";
        public string LicenseKey { get; set; } = "";

        public void Normalize()
        {
            FirebaseUrl = NormalizeBaseUrl(FirebaseUrl);
            JmsBaseUrl = NormalizeBaseUrl(string.IsNullOrWhiteSpace(JmsBaseUrl) ? "https://jms.jtexpress.vn" : JmsBaseUrl);
            JmsApiBaseUrl = NormalizeBaseUrl(string.IsNullOrWhiteSpace(JmsApiBaseUrl) ? "https://jmsgw.jtexpress.vn" : JmsApiBaseUrl);

            if (string.IsNullOrWhiteSpace(InternetCheckUrl))
                InternetCheckUrl = "http://clients3.google.com/generate_204";

            if (!string.IsNullOrWhiteSpace(GoogleSheetUrl) && string.IsNullOrWhiteSpace(DataSpreadsheetId))
                DataSpreadsheetId = ExtractSpreadsheetId(GoogleSheetUrl);

            if (!string.IsNullOrWhiteSpace(DataSpreadsheetId) && string.IsNullOrWhiteSpace(GoogleSheetUrl))
                GoogleSheetUrl = $"https://docs.google.com/spreadsheets/d/{DataSpreadsheetId}";
        }

        public void MergeFrom(AppRuntimeConfig? source, bool includeFirebase)
        {
            if (source == null) return;

            if (includeFirebase)
            {
                FirebaseUrl = Pick(source.FirebaseUrl, FirebaseUrl);
                FirebaseDatabaseSecret = Pick(source.FirebaseDatabaseSecret, FirebaseDatabaseSecret);
            }

            DataSpreadsheetId = Pick(source.DataSpreadsheetId, DataSpreadsheetId);
            GoogleSheetUrl = Pick(source.GoogleSheetUrl, GoogleSheetUrl);
            LicenseSpreadsheetId = Pick(source.LicenseSpreadsheetId, LicenseSpreadsheetId);
            AppsScriptUrl = Pick(source.AppsScriptUrl, AppsScriptUrl);
            GoogleServiceAccountJson = Pick(source.GoogleServiceAccountJson, GoogleServiceAccountJson);
            GoogleCredentialPath = Pick(source.GoogleCredentialPath, GoogleCredentialPath);
            JmsBaseUrl = Pick(source.JmsBaseUrl, JmsBaseUrl);
            JmsApiBaseUrl = Pick(source.JmsApiBaseUrl, JmsApiBaseUrl);
            InternetCheckUrl = Pick(source.InternetCheckUrl, InternetCheckUrl);
            UpdateXmlUrl = Pick(source.UpdateXmlUrl, UpdateXmlUrl);

            ActionSiteCode = Pick(source.ActionSiteCode, ActionSiteCode);
            LicenseKey = Pick(source.LicenseKey, LicenseKey);

            Normalize();
        }

        public string BuildJmsUrl(string relativePath) => CombineUrl(JmsBaseUrl, relativePath);
        public string BuildJmsApiUrl(string relativePath) => CombineUrl(JmsApiBaseUrl, relativePath);

        public static string ExtractSpreadsheetId(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "";
            string trimmed = value.Trim();
            Match match = Regex.Match(trimmed, @"/spreadsheets/d/([^/?#]+)", RegexOptions.IgnoreCase);
            if (match.Success) return match.Groups[1].Value;
            return trimmed;
        }

        private static string Pick(string? candidate, string current) => string.IsNullOrWhiteSpace(candidate) ? current : candidate.Trim();
        private static string NormalizeBaseUrl(string value) => string.IsNullOrWhiteSpace(value) ? "" : value.Trim().TrimEnd('/') + "/";
        private static string CombineUrl(string baseUrl, string relativePath)
        {
            if (string.IsNullOrWhiteSpace(baseUrl)) return relativePath;
            return baseUrl.TrimEnd('/') + "/" + relativePath.TrimStart('/');
        }
    }

    public static class AppConfig
    {
        private static readonly string SecureConfigPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AutoJMS.secure");
        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions { WriteIndented = true, PropertyNameCaseInsensitive = true };
        private static AppRuntimeConfig _current = new AppRuntimeConfig();
        private static string _lastSecret = "";

        public static AppRuntimeConfig Current => _current;

        public static void LoadBootstrap(string machineKey)
        {
            string secret = ResolveSecret(machineKey);
            var config = new AppRuntimeConfig();

            if (File.Exists(SecureConfigPath))
            {
                try
                {
                    string protectedJson = File.ReadAllText(SecureConfigPath);
                    string json = SecureConfigCrypto.UnprotectString(protectedJson, secret);
                    if (!string.IsNullOrWhiteSpace(json))
                    {
                        config = JsonSerializer.Deserialize<AppRuntimeConfig>(json, JsonOptions) ?? new AppRuntimeConfig();
                    }
                }
                catch
                {
                    // Nếu giải mã thất bại (VD: đem sang máy khác chạy), tạo mới cấu hình
                    config = new AppRuntimeConfig();
                }
            }

            ApplyEnvironment(config);
            config.Normalize();
            _current = config;
            _lastSecret = secret;

            if (!File.Exists(SecureConfigPath) && HasBootstrapEnvironment())
                SaveCurrent();
        }

        public static bool HasFirebaseConfig() => !string.IsNullOrWhiteSpace(Current.FirebaseUrl);

        public static void ApplyLicenseConfig(AppRuntimeConfig? licenseConfig)
        {
            _current.MergeFrom(licenseConfig, includeFirebase: false);
            SaveCurrent();
        }

        public static void SaveCurrent()
        {
            if (string.IsNullOrWhiteSpace(_lastSecret))
                _lastSecret = ResolveSecret(Environment.MachineName);

            try
            {
                string json = JsonSerializer.Serialize(_current, JsonOptions);
                string protectedJson = SecureConfigCrypto.ProtectString(json, _lastSecret);
                File.WriteAllText(SecureConfigPath, protectedJson);
            }
            catch { }
        }

        private static string ResolveSecret(string machineKey)
        {
            string? configuredKey = Environment.GetEnvironmentVariable("AUTOJMS_CONFIG_KEY");
            if (!string.IsNullOrWhiteSpace(configuredKey)) return configuredKey;
            return $"{Environment.MachineName}|{Environment.UserName}|{machineKey}|AutoJMS";
        }

        private static bool HasBootstrapEnvironment() => !string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("AUTOJMS_FIREBASE_URL"));

        private static void ApplyEnvironment(AppRuntimeConfig config)
        {
            Apply("AUTOJMS_FIREBASE_URL", value => config.FirebaseUrl = value);
            Apply("AUTOJMS_FIREBASE_SECRET", value => config.FirebaseDatabaseSecret = value);
            Apply("AUTOJMS_DATA_SPREADSHEET_ID", value => config.DataSpreadsheetId = value);
            Apply("AUTOJMS_GOOGLE_SHEET_URL", value => config.GoogleSheetUrl = value);
            Apply("AUTOJMS_LICENSE_SPREADSHEET_ID", value => config.LicenseSpreadsheetId = value);
            Apply("AUTOJMS_APPS_SCRIPT_URL", value => config.AppsScriptUrl = value);
            Apply("AUTOJMS_GOOGLE_SERVICE_ACCOUNT_JSON", value => config.GoogleServiceAccountJson = value);
            Apply("AUTOJMS_GOOGLE_CREDENTIAL_PATH", value => config.GoogleCredentialPath = value);
            Apply("AUTOJMS_JMS_BASE_URL", value => config.JmsBaseUrl = value);
            Apply("AUTOJMS_JMS_API_BASE_URL", value => config.JmsApiBaseUrl = value);
            Apply("AUTOJMS_INTERNET_CHECK_URL", value => config.InternetCheckUrl = value);
            Apply("AUTOJMS_UPDATE_XML_URL", value => config.UpdateXmlUrl = value);
        }

        private static void Apply(string name, Action<string> setter)
        {
            string? value = Environment.GetEnvironmentVariable(name);
            if (!string.IsNullOrWhiteSpace(value)) setter(value.Trim());
        }
    }
}