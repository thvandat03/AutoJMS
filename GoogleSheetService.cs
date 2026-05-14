using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    public static class GoogleSheetService
    {
        public static string DATA_SPREADSHEET_ID => AppConfig.Current.DataSpreadsheetId;
        public static string LICENSE_SPREADSHEET_ID => AppConfig.Current.LicenseSpreadsheetId;

        private static SheetsService Service;

        // ==========================================
        // BỨC TƯỜNG THÀNH 1: BỘ NHỚ TẠM (CACHE 30 GIÂY)
        // ==========================================
        private static readonly ConcurrentDictionary<string, (List<IList<object>> data, DateTime timestamp)> _rangeCache
            = new ConcurrentDictionary<string, (List<IList<object>>, DateTime)>();
        private static readonly TimeSpan CacheDuration = TimeSpan.FromSeconds(30);

        // ==========================================
        // BỨC TƯỜNG THÀNH 2: TRẠM THU PHÍ (RATE LIMITER CHỐNG QUOTA 429)
        // ==========================================
        private static readonly SemaphoreSlim _apiSemaphore = new SemaphoreSlim(1, 1);
        private static DateTime _lastApiCallTime = DateTime.MinValue;
        private const int API_DELAY_MS = 1500;

        // ==========================================
        // 1. KHỞI TẠO & CẤU HÌNH DỊCH VỤ
        // ==========================================
        public static void InitService()
        {
            if (Service != null) return;

            try
            {
                GoogleCredential credential;
                string serviceAccountJson = AppConfig.Current.GoogleServiceAccountJson;

                if (!string.IsNullOrWhiteSpace(serviceAccountJson))
                {
                    credential = GoogleCredential.FromJson(serviceAccountJson)
                        .CreateScoped(SheetsService.Scope.Spreadsheets);
                }
                else
                {
                    string credentialPath = ResolveCredentialPath();
                    if (!File.Exists(credentialPath))
                        throw new FileNotFoundException($"Không tìm thấy tệp xác thực Google Sheet");

                    using var stream = new FileStream(credentialPath, FileMode.Open, FileAccess.Read);
                    credential = GoogleCredential.FromStream(stream)
                        .CreateScoped(SheetsService.Scope.Spreadsheets);
                }

                Service = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "AutoJMS"
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khởi tạo xác thực Google:\n{ex.Message}", "Lỗi InitService", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        public static void ResetService()
        {
            if (Service is IDisposable disposable) disposable.Dispose();
            Service = null;
        }

        private static string ResolveCredentialPath()
        {
            string configuredPath = AppConfig.Current.GoogleCredentialPath;
            if (string.IsNullOrWhiteSpace(configuredPath)) configuredPath = "service_account.json";

            return Path.IsPathRooted(configuredPath)
                ? configuredPath
                : Path.Combine(AppDomain.CurrentDomain.BaseDirectory, configuredPath);
        }

        public static bool IsConfigValid()
        {
            bool hasCredential = !string.IsNullOrWhiteSpace(AppConfig.Current.GoogleServiceAccountJson)
                || File.Exists(ResolveCredentialPath());
            return hasCredential && !string.IsNullOrWhiteSpace(DATA_SPREADSHEET_ID);
        }

        // ==========================================
        // 2. LÕI CHỐNG LỖI QUOTA VÀ LỖI MẠNG
        // ==========================================

        // HÀM MỚI: Nhận diện lỗi do rớt mạng / đứt kết nối
        private static bool IsNetworkError(Exception ex)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return true;

            string msg = ex.Message.ToLower();
            return msg.Contains("no such host is known") ||
                   msg.Contains("a connection attempt failed") ||
                   msg.Contains("network is unreachable") ||
                   msg.Contains("the operation was canceled") ||
                   msg.Contains("task was canceled");
        }

        private static bool TryGetFromCache(string spreadsheetId, string range, out List<IList<object>> data)
        {
            string key = $"{spreadsheetId}|{range}";
            if (_rangeCache.TryGetValue(key, out var entry) && DateTime.UtcNow - entry.timestamp < CacheDuration)
            {
                data = entry.data;
                return true;
            }
            data = null;
            return false;
        }

        private static void SetCache(string spreadsheetId, string range, IList<IList<object>> data)
        {
            string key = $"{spreadsheetId}|{range}";
            _rangeCache[key] = (data?.ToList() ?? new List<IList<object>>(), DateTime.UtcNow);
        }

        private static void InvalidateCache(string spreadsheetId, string range)
        {
            string prefix = $"{spreadsheetId}|";
            var keysToRemove = _rangeCache.Keys.Where(k => k.StartsWith(prefix)).ToList();
            foreach (var key in keysToRemove)
            {
                _rangeCache.TryRemove(key, out _);
            }
        }

        private static async Task ThrottleApiCallAsync()
        {
            var now = DateTime.UtcNow;
            var timeSinceLastCall = now - _lastApiCallTime;

            if (timeSinceLastCall.TotalMilliseconds < API_DELAY_MS)
            {
                int delay = API_DELAY_MS - (int)timeSinceLastCall.TotalMilliseconds;
                await Task.Delay(delay).ConfigureAwait(false);
            }

            _lastApiCallTime = DateTime.UtcNow;
        }

        private static async Task<T> CallWithRetryAsync<T>(Func<Task<T>> func, int maxRetries = 5)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return default;

            await _apiSemaphore.WaitAsync().ConfigureAwait(false);
            try
            {
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    if (NetworkState.Current == NetworkStatus.Offline) return default;

                    await ThrottleApiCallAsync().ConfigureAwait(false);
                    try
                    {
                        return await func().ConfigureAwait(false);
                    }
                    catch (GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
                    {
                        if (attempt == maxRetries) throw;
                        int backoff = (int)Math.Pow(2, attempt) * 2000;
                        await Task.Delay(backoff).ConfigureAwait(false);
                    }
                }
                return default;
            }
            finally
            {
                _apiSemaphore.Release();
            }
        }

        private static async Task CallWithRetryVoidAsync(Func<Task> func, int maxRetries = 5)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return;

            await _apiSemaphore.WaitAsync().ConfigureAwait(false);
            try
            {
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    if (NetworkState.Current == NetworkStatus.Offline) return;

                    await ThrottleApiCallAsync().ConfigureAwait(false);
                    try
                    {
                        await func().ConfigureAwait(false);
                        return;
                    }
                    catch (GoogleApiException ex) when (ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests)
                    {
                        if (attempt == maxRetries) throw;
                        int backoff = (int)Math.Pow(2, attempt) * 2000;
                        await Task.Delay(backoff).ConfigureAwait(false);
                    }
                }
            }
            finally
            {
                _apiSemaphore.Release();
            }
        }

        // ==========================================
        // 3. CÁC HÀM TƯƠNG TÁC GOOGLE SHEET 
        // ==========================================

        public static async Task<List<IList<IList<object>>>> BatchReadAsync(string spreadsheetId, List<string> ranges)
        {
            // NGẮT NGAY TỪ CỬA NẾU MẤT MẠNG
            if (NetworkState.Current == NetworkStatus.Offline) return new List<IList<IList<object>>>();

            try
            {
                if (Service == null) InitService();

                var result = new List<IList<IList<object>>>();
                var rangesToFetch = new List<string>();
                var cacheResults = new Dictionary<string, IList<IList<object>>>();

                foreach (var range in ranges)
                {
                    if (TryGetFromCache(spreadsheetId, range, out var cachedData))
                        cacheResults[range] = cachedData;
                    else
                        rangesToFetch.Add(range);
                }

                if (rangesToFetch.Count > 0)
                {
                    var response = await CallWithRetryAsync(async () =>
                    {
                        var request = Service.Spreadsheets.Values.BatchGet(spreadsheetId);
                        request.Ranges = rangesToFetch;
                        return await request.ExecuteAsync().ConfigureAwait(false);
                    }).ConfigureAwait(false);

                    if (response?.ValueRanges != null)
                    {
                        for (int i = 0; i < response.ValueRanges.Count; i++)
                        {
                            var fetchedData = response.ValueRanges[i].Values ?? new List<IList<object>>();
                            SetCache(spreadsheetId, rangesToFetch[i], fetchedData);
                            cacheResults[rangesToFetch[i]] = fetchedData;
                        }
                    }
                }

                foreach (var range in ranges)
                {
                    result.Add(cacheResults.ContainsKey(range) ? cacheResults[range] : new List<IList<object>>());
                }
                return result;
            }
            catch (Exception ex)
            {
                // NẾU LỖI DO MẠNG YẾU/MẤT KẾT NỐI -> LẶNG LẼ TRẢ VỀ RỖNG, KHÔNG HIỆN THÔNG BÁO
                if (IsNetworkError(ex)) return new List<IList<IList<object>>>();

                // NẾU LÀ LỖI CODE THẬT (VD: SAI ID SHEET) -> MỚI HIỆN BẢNG
                AppLogger.Error($"Lỗi hệ thống (BatchRead):\n{ex.Message}");
                return new List<IList<IList<object>>>();
            }
        }

        public static async Task<IList<IList<object>>> ReadRangeAsync(string spreadsheetId, string range)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return new List<IList<object>>();

            var batchResult = await BatchReadAsync(spreadsheetId, new List<string> { range }).ConfigureAwait(false);
            return batchResult?.FirstOrDefault() ?? new List<IList<object>>();
        }

        public static async Task ClearSheetAsync(string spreadsheetId, string range)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return;

            try
            {
                if (Service == null) InitService();
                await CallWithRetryVoidAsync(async () =>
                {
                    var clearRequest = Service.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadsheetId, range);
                    await clearRequest.ExecuteAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);

                InvalidateCache(spreadsheetId, range);
            }
            catch (Exception ex)
            {
                if (IsNetworkError(ex)) return;
                MessageBox.Show(ex.Message, "Lỗi ClearSheet");
            }
        }

        public static async Task UpdateBumpSheetAsync(IList<IList<object>> values, string spreadsheetId, string range)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return;

            try
            {
                if (Service == null) InitService();
                await CallWithRetryVoidAsync(async () =>
                {
                    var valueRange = new ValueRange { Values = values };
                    var updateRequest = Service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    await updateRequest.ExecuteAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);

                InvalidateCache(spreadsheetId, range);
            }
            catch (Exception ex)
            {
                if (IsNetworkError(ex)) return;
                MessageBox.Show(ex.Message, "Lỗi UpdateSheet");
            }
        }

        public static async Task UpdateCellAsync(string spreadsheetId, string range, string value)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return;

            try
            {
                if (Service == null) InitService();
                await CallWithRetryVoidAsync(async () =>
                {
                    var oblist = new List<object>() { value };
                    var valueRange = new ValueRange { Values = new List<IList<object>> { oblist } };
                    var updateRequest = Service.Spreadsheets.Values.Update(valueRange, spreadsheetId, range);
                    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    await updateRequest.ExecuteAsync().ConfigureAwait(false);
                }).ConfigureAwait(false);

                InvalidateCache(spreadsheetId, range);
            }
            catch (Exception ex)
            {
                if (IsNetworkError(ex)) return;
                MessageBox.Show(ex.Message, "Lỗi UpdateCell");
            }
        }

        // ==========================================
        // 4. HÀM TƯƠNG THÍCH NGƯỢC
        // ==========================================

        public static List<string> ReadColumnBySpreadsheetId(string spreadsheetId, string sheetName, int columnIndex)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return new List<string>();

            try
            {
                string colLetter = GetColumnLetter(columnIndex);
                string range = $"{sheetName}!{colLetter}2:{colLetter}";

                var values = ReadRangeAsync(spreadsheetId, range).GetAwaiter().GetResult();

                List<string> result = new List<string>();
                if (values != null)
                {
                    foreach (var row in values)
                    {
                        if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0]?.ToString()))
                            result.Add(row[0].ToString().Trim());
                    }
                }
                return result;
            }
            catch { return new List<string>(); }
        }

        public static List<string> ReadColumn(string sheetName, int colIndex)
        {
            return ReadColumnBySpreadsheetId(DATA_SPREADSHEET_ID, sheetName, colIndex);
        }

        public static IList<IList<object>> ReadRange(string spreadsheetId, string range)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return new List<IList<object>>();
            return ReadRangeAsync(spreadsheetId, range).GetAwaiter().GetResult();
        }

        public static List<IList<IList<object>>> BatchRead(string spreadsheetId, List<string> ranges)
        {
            if (NetworkState.Current == NetworkStatus.Offline) return new List<IList<IList<object>>>();
            return BatchReadAsync(spreadsheetId, ranges).GetAwaiter().GetResult();
        }

        private static string GetColumnLetter(int col)
        {
            int dividend = col;
            string columnName = String.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
    }
}