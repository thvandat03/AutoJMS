using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class InventorySyncService
    {
        private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(60) };
        private const int LeaseSeconds = 1800; // 30 phút
        private const int PageSize = 100;
        private const int MaxRetriesPerPage = 3;

        private static string GetActionSiteCode() => "214A02";

        public static async Task RunInventorySyncAsync(CancellationToken ct = default)
        {
            var now = DateTime.Now;
            if (now.Hour < 7 || now.Hour > 12) return;

            AppLogger.Info("[InventorySync] Kiểm tra lease lock...");
            bool acquired = await SupabaseDbService.TryAcquireInventoryLeaseAsync(LeaseSeconds);
            if (!acquired)
            {
                AppLogger.Info("[InventorySync] Máy khác đang giữ quyền sync tồn kho.");
                return;
            }

            AppLogger.Info("[InventorySync] Đã chiếm quyền. Bắt đầu kéo tồn kho.");

            using var heartbeatCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            var heartbeatTask = Task.Run(() => LeaseHeartbeatLoopAsync(heartbeatCts.Token), heartbeatCts.Token);
            bool success = false;

            try
            {
                var inventoryWaybills = await FetchAllInventoryWaybillsWithRetryAsync(ct);
                if (inventoryWaybills.Count > 0 && !ct.IsCancellationRequested)
                {
                    int inserted = await SupabaseDbService.UpsertNewWaybillsOnlyAsync(inventoryWaybills);
                    AppLogger.Info($"[InventorySync] Hoàn tất. Lấy {inventoryWaybills.Count} mã, thêm mới {inserted} mã.");
                }
                success = true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("[InventorySync] Lỗi kéo tồn kho", ex);
            }
            finally
            {
                heartbeatCts.Cancel();
                try { await heartbeatTask; } catch { }

                if (success) await SupabaseDbService.CompleteInventorySyncAsync();
                else await SupabaseDbService.ReleaseInventoryLeaseAsync();
            }
        }

        private static async Task LeaseHeartbeatLoopAsync(CancellationToken ct)
        {
            try
            {
                while (!ct.IsCancellationRequested)
                {
                    await Task.Delay(TimeSpan.FromMinutes(1), ct);
                    if (ct.IsCancellationRequested) break;
                    await SupabaseDbService.RefreshInventoryLeaseAsync(LeaseSeconds);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { AppLogger.Warning($"[InventorySync] Heartbeat lỗi: {ex.Message}"); }
        }

        private static async Task<List<string>> FetchAllInventoryWaybillsWithRetryAsync(CancellationToken ct)
        {
            var waybills = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(Main.CapturedAuthToken)) return waybills.ToList();

            string url = AppConfig.Current.BuildJmsApiUrl("businessindicator/bigdataReport/detail/take_ret_mon_detail_doris2");
            string actionSiteCode = GetActionSiteCode();
            string startDate = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd 00:00:00");
            string endDate = DateTime.Now.ToString("yyyy-MM-dd 23:59:59");
            int currentPage = 1;
            int totalPages = 1;

            do
            {
                bool pageSuccess = false;
                for (int retry = 1; retry <= MaxRetriesPerPage && !pageSuccess; retry++)
                {
                    ct.ThrowIfCancellationRequested();
                    try
                    {
                        var payload = new Dictionary<string, object>
                        {
                            { "current", currentPage }, { "size", PageSize }, { "dimension", "2" },
                            { "isFlag", "1" }, { "actionSiteCode", actionSiteCode },
                            { "startDate", startDate }, { "endDate", endDate }, { "countryId", "1" }
                        };

                        var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json") };
                        req.Headers.Add("authToken", Main.CapturedAuthToken);
                        req.Headers.Add("lang", "VN");

                        using var res = await _httpClient.SendAsync(req, ct);
                        var json = await res.Content.ReadAsStringAsync(ct);

                        if (!res.IsSuccessStatusCode) throw new Exception($"HTTP {(int)res.StatusCode}");

                        using var doc = JsonDocument.Parse(json);
                        var root = doc.RootElement;

                        if (!root.TryGetProperty("succ", out var succ) || !succ.GetBoolean()) throw new Exception("API trả về succ=false");

                        var data = root.GetProperty("data");
                        totalPages = data.GetProperty("pages").GetInt32();

                        foreach (var record in data.GetProperty("records").EnumerateArray())
                        {
                            if (record.TryGetProperty("billcode", out var billcodeProp))
                            {
                                var code = billcodeProp.GetString();
                                if (!string.IsNullOrWhiteSpace(code)) waybills.Add(code.Trim());
                            }
                        }
                        pageSuccess = true;
                    }
                    catch (Exception ex)
                    {
                        if (retry >= MaxRetriesPerPage) { AppLogger.Warning($"[InventorySync] Bỏ qua trang {currentPage} sau {retry} lần lỗi: {ex.Message}"); break; }
                        var delay = TimeSpan.FromSeconds(Math.Min(10, retry * 2));
                        AppLogger.Warning($"[InventorySync] Trang {currentPage} lỗi, thử lại {retry}/{MaxRetriesPerPage} sau {delay.TotalSeconds}s: {ex.Message}");
                        await Task.Delay(delay, ct);
                    }
                }
                currentPage++;
                await Task.Delay(250, ct);
            }
            while (currentPage <= totalPages && !ct.IsCancellationRequested);

            return waybills.ToList();
        }
    }
}