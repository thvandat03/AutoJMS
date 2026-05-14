using AutoJMS.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class DatabaseTracking
    {
        private static readonly HttpClient _httpClient = new() { Timeout = TimeSpan.FromSeconds(60) };
        private const int BatchSize = 40;
        private const int MaxDegreeOfParallelism = 3;

        public static async Task RunBackgroundTrackingAsync(IEnumerable<string> waybills, CancellationToken ct = default)
        {
            if (waybills == null) return;
            var list = waybills.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim().ToUpperInvariant()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (list.Count == 0 || string.IsNullOrWhiteSpace(Main.CapturedAuthToken)) return;

            AppLogger.Info($"[DatabaseTracking] Tracking {list.Count} đơn.");

            var dict = new Dictionary<string, WaybillDbModel>(StringComparer.OrdinalIgnoreCase);
            var queryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var aliasMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
                        if (!dict.ContainsKey(original)) dict[original] = new WaybillDbModel { WaybillNo = original };
                    }
                }
            }

            var codes = queryCodes.ToList();
            var batches = codes.Chunk(BatchSize).ToList();

            await Parallel.ForEachAsync(batches, new ParallelOptions { MaxDegreeOfParallelism = MaxDegreeOfParallelism, CancellationToken = ct }, async (batch, token) =>
            {
                await ProcessTrackingBatchAsync(batch.ToArray(), dict, token);
                await ProcessOrderDetailBatchAsync(batch.ToArray(), dict, token);
            });

            foreach (var kv in aliasMap)
            {
                if (dict.TryGetValue(kv.Value, out var src) && dict.TryGetValue(kv.Key, out var dst))
                {
                    dst.NhanVienNhanHang = src.NhanVienNhanHang ?? "";
                    dst.TenNguoiGui = src.TenNguoiGui ?? "";
                    dst.DiaChiLayHang = src.DiaChiLayHang ?? "";
                }
            }

            var finalUploadList = new List<WaybillDbModel>();
            foreach (var wb in list)
            {
                if (dict.TryGetValue(wb, out var row))
                {
                    if (!string.IsNullOrEmpty(row.ThaoTacCuoi) && row.ThaoTacCuoi.Contains("Ký nhận", StringComparison.OrdinalIgnoreCase))
                    {
                        row.IsActive = false;
                    }
                    row.WaybillNo = wb;
                    row.TrackingIntervalMins = 30; // Reset lại 30 phút cho lần sau
                    row.NextTrackAt = DateTime.UtcNow.AddMinutes(30);
                    row.LastTrackedAt = DateTime.UtcNow;
                    finalUploadList.Add(row);
                }
            }

            if (finalUploadList.Count > 0 && !ct.IsCancellationRequested)
            {
                await SupabaseDbService.UpsertManyWaybillsAsync(finalUploadList);
                AppLogger.Info($"[DatabaseTracking] Update xong {finalUploadList.Count} đơn.");
            }
        }

        private static async Task ProcessTrackingBatchAsync(string[] batch, Dictionary<string, WaybillDbModel> dict, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            var url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/podTracking/inner/query/keywordList");
            var payload = new Dictionary<string, object> { { "keywordList", batch }, { "trackingTypeEnum", "WAYBILL" }, { "countryId", "1" } };

            var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json") };
            req.Headers.Add("authToken", Main.CapturedAuthToken);
            req.Headers.Add("lang", "VN");
            req.Headers.Add("routeName", "trackingExpress");

            using var res = await _httpClient.SendAsync(req, ct);
            if (!res.IsSuccessStatusCode) return;

            var json = await res.Content.ReadAsStringAsync(ct);
            var result = JsonSerializer.Deserialize<WaybillHistoryResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

            if (result?.succ != true || result.data == null) return;

            foreach (var item in result.data)
            {
                var wb = item.keyword ?? item.billCode ?? "";
                if (string.IsNullOrWhiteSpace(wb)) continue;

                if (dict.TryGetValue(wb, out var row)) ExtractTrackingData(row, item.details);
            }
        }

        private static async Task ProcessOrderDetailBatchAsync(string[] batch, Dictionary<string, WaybillDbModel> dict, CancellationToken ct)
        {
            foreach (var wb in batch)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    var url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/order/getOrderDetail");
                    var payload = new Dictionary<string, object> { { "waybillNo", wb } };
                    var req = new HttpRequestMessage(HttpMethod.Post, url) { Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json") };
                    req.Headers.Add("authToken", Main.CapturedAuthToken);
                    req.Headers.Add("lang", "VN");

                    using var res = await _httpClient.SendAsync(req, ct);
                    if (!res.IsSuccessStatusCode) continue;

                    var json = await res.Content.ReadAsStringAsync(ct);
                    var result = JsonSerializer.Deserialize<OrderDetailResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    if (result?.succ == true && result.data?.details != null && dict.TryGetValue(wb, out var row))
                    {
                        var info = result.data.details;
                        row.NhanVienNhanHang = info.staffName ?? "";
                        row.DiaChiLayHang = info.senderDetailedAddress ?? "";
                        row.ThoiGianNhanHang = info.pickTime ?? "";
                        row.NoiDungHangHoa = info.goodsName ?? "";
                        row.CODThucTe = info.codMoney ?? "";
                        row.TenNguoiGui = info.customerName ?? "";
                        row.TrongLuong = info.packageChargeWeight?.ToString() ?? "";
                        row.PTTT = info.paymentModeName ?? "";
                        row.DiaChiNhanHang = info.receiverDetailedAddress ?? "";
                        row.Phuong = info.destinationName ?? "";
                        if (!string.IsNullOrEmpty(info.terminalDispatchCode))
                        {
                            var parts = info.terminalDispatchCode.Split('-');
                            row.MaDoan1 = parts.Length > 0 ? parts[0] : "";
                            row.MaDoan2 = parts.Length > 1 ? parts[1] : "";
                            row.MaDoan3 = parts.Length > 2 ? parts[2] : "";
                            row.MaDoanFull = info.terminalDispatchCode;
                        }
                    }
                }
                catch { }
            }
        }

        private static void ExtractTrackingData(WaybillDbModel row, List<WaybillDetail> details)
        {
            if (details == null || details.Count == 0) return;
            row.DauChuyenHoan = details.Any(d => d.status == "已审核") ? "Có" : "Không";
            
            var staffNameFisrt = details.Where(d => (d.scanTypeName?.Contains("Nhận hàng") == true || d.scanTypeName?.Contains("Lấy hàng") == true || d.status == "已揽件") && (!string.IsNullOrEmpty(d.staffName) || !string.IsNullOrEmpty(d.scanByName))).OrderBy(d => DateTime.TryParse(!string.IsNullOrEmpty(d.uploadTime) ? d.uploadTime : (d.scanTime ?? "9999-12-31"), out DateTime dt) ? dt : DateTime.MaxValue).FirstOrDefault();
            if (staffNameFisrt != null) row.NhanVienNhanHang = staffNameFisrt.staffName ?? staffNameFisrt.scanByName ?? "";

            var giaoLaiGanNhat = details.Where(d => d.scanTypeName != null && d.scanTypeName.Contains("Giao lại hàng")).OrderByDescending(d => DateTime.TryParse(d.uploadTime ?? d.scanTime ?? "", out DateTime dt) ? dt : DateTime.MinValue).FirstOrDefault();
            row.ThoiGianYeuCauPhatLai = giaoLaiGanNhat?.remark2 ?? "";

            var vanDe = details.Where(d => d.scanTypeName?.Contains("vấn đề") == true || d.scanTypeName?.Contains("Kiện vấn đề") == true).OrderByDescending(d => DateTime.TryParse(d.uploadTime ?? d.scanTime ?? "", out DateTime dt) ? dt : DateTime.MinValue).FirstOrDefault();
            if (vanDe != null) { row.NhanVienKienVanDe = vanDe.scanByName ?? ""; row.NguyenNhanKienVanDe = vanDe.remark1 ?? ""; }

            var latest = details.LastOrDefault();
            if (latest != null)
            {
                row.TrangThaiHienTai = latest.waybillTrackingContent ?? "";
                row.ThaoTacCuoi = latest.scanTypeName ?? "";
                row.ThoiGianThaoTac = latest.uploadTime ?? latest.scanTime ?? "";
                row.BuuCucThaoTac = latest.scanNetworkName ?? "";
                row.NguoiThaoTac = latest.scanByName ?? "";
            }
        }
    }
}