using AutoJMS.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    // =====================================================================
    // CÁC CLASS MODEL HỨNG DỮ LIỆU JSON TỪ API JMS
    // Vẫn giữ lại thẻ Obfuscation để .NET Reactor biết đường bỏ qua, 
    // tránh lỗi mã hóa làm hỏng cấu trúc JSON.
    // =====================================================================

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillHistoryResponse
    {
        [JsonPropertyName("succ")]
        public bool succ { get; set; }

        [JsonPropertyName("msg")]
        public string? msg { get; set; }

        [JsonPropertyName("data")]
        public List<WaybillHistoryItem> data { get; set; } = new();
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillHistoryItem
    {
        [JsonPropertyName("keyword")]
        public string? keyword { get; set; }

        [JsonPropertyName("billCode")]
        public string? billCode { get; set; }

        [JsonPropertyName("billcode")]
        public string? billcode { get; set; }

        [JsonPropertyName("details")]
        public List<WaybillDetail> details { get; set; } = new();
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillDetail
    {
        [JsonPropertyName("time")]
        public string? time { get; set; }

        [JsonPropertyName("action")]
        public string? action { get; set; }

        [JsonPropertyName("siteCode")]
        public string? siteCode { get; set; }

        [JsonPropertyName("siteName")]
        public string? siteName { get; set; }

        [JsonPropertyName("operatorName")]
        public string? operatorName { get; set; }

        [JsonPropertyName("remark")]
        public string? remark { get; set; }

        [JsonPropertyName("desc")]
        public string? desc { get; set; }
    }

    // =====================================================================
    // CÔNG NHÂN CHẠY NGẦM: TRACKING HÀNH TRÌNH TỪ JMS
    // =====================================================================

    public static class DatabaseTracking
    {
        private static readonly HttpClient _httpClient = new()
        {
            Timeout = TimeSpan.FromSeconds(60)
        };

        private const int BatchSize = 40;
        private const int MaxDegreeOfParallelism = 3;

        public static async Task RunBackgroundTrackingAsync(IEnumerable<string> waybills, CancellationToken ct = default)
        {
            if (waybills == null) return;

            // 1. Chuẩn hóa danh sách mã vận đơn đầu vào
            var list = waybills
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim().ToUpperInvariant())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (list.Count == 0 || string.IsNullOrWhiteSpace(Main.CapturedAuthToken))
                return;

            AppLogger.Info($"[DatabaseTracking] Bắt đầu tracking {list.Count} đơn.");

            var dict = new Dictionary<string, WaybillDbModel>(StringComparer.OrdinalIgnoreCase);
            var queryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var aliasMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // 2. Xử lý mã có chứa hậu tố (VD: -001) để lấy đúng mã gốc đi hỏi API
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
                        aliasMap[wb] = original;

                        if (!dict.ContainsKey(original))
                        {
                            dict[original] = new WaybillDbModel { WaybillNo = original };
                        }
                    }
                }
            }

            // 3. Chia nhỏ danh sách thành các Batch (mỗi Batch 40 mã)
            var batches = queryCodes.ToList().Chunk(BatchSize).ToList();

            // 4. Gọi API JMS song song (Giới hạn tối đa 3 luồng để chống khóa IP)
            await Parallel.ForEachAsync(
                batches,
                new ParallelOptions
                {
                    MaxDegreeOfParallelism = MaxDegreeOfParallelism,
                    CancellationToken = ct
                },
                async (batch, token) =>
                {
                    await ProcessTrackingBatchAsync(batch.ToArray(), dict, token);
                    await ProcessOrderDetailBatchAsync(batch.ToArray(), dict, token);
                });

            // 5. Đồng bộ dữ liệu từ mã gốc sang mã có hậu tố
            foreach (var kv in aliasMap)
            {
                if (dict.TryGetValue(kv.Value, out var src) && dict.TryGetValue(kv.Key, out var dst))
                {
                    dst.NhanVienNhanHang = src.NhanVienNhanHang ?? "";
                    dst.TenNguoiGui = src.TenNguoiGui ?? "";
                    dst.DiaChiLayHang = src.DiaChiLayHang ?? "";
                }
            }

            // 6. Gói ghém kết quả cuối cùng để đẩy lên Database
            var finalUploadList = new List<WaybillDbModel>();

            foreach (var wb in list)
            {
                if (!dict.TryGetValue(wb, out var row))
                    continue;

                // Nếu đơn đã ký nhận, đánh dấu tắt tracking vĩnh viễn
                if (!string.IsNullOrEmpty(row.ThaoTacCuoi) &&
                    row.ThaoTacCuoi.Contains("Ký nhận", StringComparison.OrdinalIgnoreCase))
                {
                    row.IsActive = false;
                }

                row.WaybillNo = wb;
                finalUploadList.Add(row);
            }

            // 7. Lưu tất cả kết quả lên Supabase
            if (finalUploadList.Count > 0 && !ct.IsCancellationRequested)
            {
                await SupabaseDbService.UpsertManyWaybillsAsync(finalUploadList);
                AppLogger.Info($"[DatabaseTracking] Đã update lên Database {finalUploadList.Count} đơn.");
            }
        }

        private static async Task ProcessTrackingBatchAsync(
            string[] batch,
            Dictionary<string, WaybillDbModel> dict,
            CancellationToken ct)
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
                    ExtractTrackingData(row, item.details);
                }
            }
        }

        private static async Task ProcessOrderDetailBatchAsync(
            string[] batch,
            Dictionary<string, WaybillDbModel> dict,
            CancellationToken ct)
        {
            // Placeholder: Bạn có thể implement logic lấy chi tiết đơn (người gửi, người nhận) ở đây
            // Hoặc bỏ qua nếu đã lấy đủ thông tin từ hành trình.
            await Task.CompletedTask;
        }

        private static void ExtractTrackingData(WaybillDbModel row, List<WaybillDetail>? details)
        {
            if (details == null || details.Count == 0)
            {
                // Nếu chưa có hành trình, cập nhật lại chu kỳ quét tiếp theo
                row.LastTrackedAt = DateTime.UtcNow;
                row.NextTrackAt = DateTime.UtcNow.AddMinutes(row.TrackingIntervalMins <= 0 ? 30 : row.TrackingIntervalMins);
                return;
            }

            // Lấy dòng hành trình mới nhất (nằm ở cuối list)
            var last = details.Last();

            row.ThaoTacCuoi = last.action ?? last.desc ?? last.remark ?? row.ThaoTacCuoi;
            
            // Tìm trạng thái "Kiện vấn đề"
            var vanDe = details.Where(d => d.action?.Contains("vấn đề", StringComparison.OrdinalIgnoreCase) == true)
                               .LastOrDefault();
            if (vanDe != null)
            {
                row.NhanVienKienVanDe = vanDe.operatorName ?? "";
            }

            // Cập nhật nhãn thời gian tracking
            row.LastTrackedAt = DateTime.UtcNow;
            var interval = row.TrackingIntervalMins <= 0 ? 30 : row.TrackingIntervalMins;
            row.NextTrackAt = DateTime.UtcNow.AddMinutes(interval);
        }
    }
}
