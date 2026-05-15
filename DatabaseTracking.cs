using AutoJMS.Data;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillHistoryResponse
    {
        [JsonPropertyName("succ")] public bool succ { get; set; }
        [JsonPropertyName("msg")] public string? msg { get; set; }
        [JsonPropertyName("data")] public List<WaybillHistoryItem> data { get; set; } = new();
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillHistoryItem
    {
        [JsonPropertyName("keyword")] public string? keyword { get; set; }
        [JsonPropertyName("billCode")] public string? billCode { get; set; }
        [JsonPropertyName("billcode")] public string? billcode { get; set; }
        [JsonPropertyName("details")] public List<WaybillDetail> details { get; set; } = new();
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillDetail
    {
        [JsonPropertyName("time")] public string? time { get; set; }
        [JsonPropertyName("action")] public string? action { get; set; }
        [JsonPropertyName("siteCode")] public string? siteCode { get; set; }
        [JsonPropertyName("siteName")] public string? siteName { get; set; }
        [JsonPropertyName("operatorName")] public string? operatorName { get; set; }
        [JsonPropertyName("remark")] public string? remark { get; set; }
        [JsonPropertyName("desc")] public string? desc { get; set; }
    }

    public static class DatabaseTracking
    {
        private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(60) };

        private const int BatchSize = 40;
        private const int MaxDegreeOfParallelism = 3;

        public static async Task RunBackgroundTrackingAsync(IEnumerable<string> waybills, CancellationToken ct = default)
        {
            if (waybills == null) return;

            var list = waybills
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim().ToUpperInvariant())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (list.Count == 0 || string.IsNullOrWhiteSpace(Main.CapturedAuthToken))
                return;

            AppLogger.Info($"[DatabaseTracking] Bắt đầu tracking {list.Count} đơn.");

            var dict = new ConcurrentDictionary<string, WaybillDbModel>(StringComparer.OrdinalIgnoreCase);
            var queryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var wb in list)
            {
                ct.ThrowIfCancellationRequested();
                queryCodes.Add(wb);
                dict[wb] = new WaybillDbModel { WaybillNo = wb };

                if (wb.Contains("-"))
                {
                    var original = wb.Split('-')[0].Trim();
                    if (!string.IsNullOrWhiteSpace(original))
                    {
                        queryCodes.Add(original);
                        dict.TryAdd(original, new WaybillDbModel { WaybillNo = original });
                    }
                }
            }

            var batches = queryCodes.ToList().Chunk(BatchSize).ToList();

            await Parallel.ForEachAsync(
                batches,
                new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism, CancellationToken = ct },
                async (batch, token) =>
                {
                    await ProcessTrackingBatchAsync(batch.ToArray(), dict, token);
                });

            var finalUploadList = new List<WaybillDbModel>();

            foreach (var wb in list)
            {
                if (!dict.TryGetValue(wb, out var row)) continue;

                lock (row)
                {
                    if (!string.IsNullOrEmpty(row.ThaoTacCuoi) &&
                        row.ThaoTacCuoi.Contains("Ký nhận", StringComparison.OrdinalIgnoreCase))
                    {
                        row.IsActive = false;
                    }

                    row.WaybillNo = wb;
                    row.UpdatedAt = DateTime.UtcNow;
                    finalUploadList.Add(row);
                }
            }

            if (finalUploadList.Count > 0 && !ct.IsCancellationRequested)
            {
                await SupabaseDbService.UpsertManyWaybillsAsync(finalUploadList);
                AppLogger.Info($"[DatabaseTracking] Đã update lên database {finalUploadList.Count} đơn.");
            }
        }

        private static async Task ProcessTrackingBatchAsync(string[] batch, ConcurrentDictionary<string, WaybillDbModel> dict, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();

            string url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/podTracking/inner/query/keywordList");

            var payload = new Dictionary<string, object>
            {
                { "keywordList", batch },
                { "trackingTypeEnum", "WAYBILL" },
                { "countryId", "1" }
            };

            var req = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), System.Text.Encoding.UTF8, "application/json")
            };

            req.Headers.Add("authToken", Main.CapturedAuthToken);
            req.Headers.Add("lang", "VN");
            req.Headers.Add("routeName", "trackingExpress");

            using var res = await _httpClient.SendAsync(req, ct);
            if (!res.IsSuccessStatusCode) return;

            var json = await res.Content.ReadAsStringAsync(ct);

            var result = JsonSerializer.Deserialize<WaybillHistoryResponse>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (result?.succ != true || result.data == null) return;

            foreach (var item in result.data)
            {
                var wb = item.keyword ?? item.billCode ?? item.billcode ?? string.Empty;
                if (string.IsNullOrWhiteSpace(wb)) continue;

                if (dict.TryGetValue(wb, out var row))
                {
                    lock (row)
                    {
                        ExtractTrackingData(row, item.details);
                    }
                }
            }
        }

        private static void ExtractTrackingData(WaybillDbModel row, List<WaybillDetail>? details)
        {
            row.LastTrackedAt = DateTime.UtcNow;

            if (details == null || details.Count == 0)
            {
                row.NextTrackAt = DateTime.UtcNow.AddMinutes(row.TrackingIntervalMins ?? 30);
                row.UpdatedAt = DateTime.UtcNow;
                return;
            }

            var last = details.Last();

            row.ThaoTacCuoi = last.action ?? last.desc ?? last.remark ?? row.ThaoTacCuoi;
            row.ThoiGianThaoTac = last.time ?? row.ThoiGianThaoTac;
            row.BuuCucThaoTac = last.siteName ?? row.BuuCucThaoTac;
            row.NguoiThaoTac = last.operatorName ?? row.NguoiThaoTac;

            if (!string.IsNullOrWhiteSpace(row.ThaoTacCuoi))
            {
                row.TrangThaiHienTai = row.ThaoTacCuoi;
            }

            row.NextTrackAt = DateTime.UtcNow.AddMinutes(row.TrackingIntervalMins ?? 30);
            row.UpdatedAt = DateTime.UtcNow;
        }
    }
}