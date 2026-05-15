using Supabase.Postgrest.Attributes;
using Supabase.Postgrest.Models;
using Supabase.Postgrest.Responses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS.Data
{
    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    [Table("waybills")]
    public class WaybillDbModel : BaseModel
    {
        [PrimaryKey("waybill_no", false)]
        [Column("waybill_no")]
        public string WaybillNo { get; set; } = string.Empty;

        [Column("is_active")]
        public bool? IsActive { get; set; } = true;

        [Column("tracking_interval_mins")]
        public int? TrackingIntervalMins { get; set; } = 30;

        [Column("next_track_at")]
        public DateTime? NextTrackAt { get; set; }

        [Column("last_tracked_at")]
        public DateTime? LastTrackedAt { get; set; }

        [Column("trang_thai_hien_tai")]
        public string? TrangThaiHienTai { get; set; }

        [Column("thao_tac_cuoi")]
        public string? ThaoTacCuoi { get; set; }

        [Column("thoi_gian_thao_tac")]
        public string? ThoiGianThaoTac { get; set; }

        [Column("nhan_vien_kien_van_de")]
        public string? NhanVienKienVanDe { get; set; }

        [Column("nguyen_nhan_kien_van_de")]
        public string? NguyenNhanKienVanDe { get; set; }

        [Column("buu_cuc_thao_tac")]
        public string? BuuCucThaoTac { get; set; }

        [Column("nguoi_thao_tac")]
        public string? NguoiThaoTac { get; set; }

        [Column("dau_chuyen_hoan")]
        public string? DauChuyenHoan { get; set; }

        [Column("dia_chi_nhan_hang")]
        public string? DiaChiNhanHang { get; set; }

        [Column("phuong")]
        public string? Phuong { get; set; }

        [Column("noi_dung_hang_hoa")]
        public string? NoiDungHangHoa { get; set; }

        [Column("cod_thuc_te")]
        public string? CODThucTe { get; set; }

        [Column("pttt")]
        public string? PTTT { get; set; }

        [Column("nhan_vien_nhan_hang")]
        public string? NhanVienNhanHang { get; set; }

        [Column("dia_chi_lay_hang")]
        public string? DiaChiLayHang { get; set; }

        [Column("thoi_gian_nhan_hang")]
        public string? ThoiGianNhanHang { get; set; }

        [Column("ten_nguoi_gui")]
        public string? TenNguoiGui { get; set; }

        [Column("trong_luong")]
        public string? TrongLuong { get; set; }

        [Column("ma_doan_full")]
        public string? MaDoanFull { get; set; }

        [Column("ma_doan_1")]
        public string? MaDoan1 { get; set; }

        [Column("ma_doan_2")]
        public string? MaDoan2 { get; set; }

        [Column("ma_doan_3")]
        public string? MaDoan3 { get; set; }

        [Column("print_count")]
        public int? PrintCount { get; set; }

        [Column("created_at")]
        public DateTime? CreatedAt { get; set; }

        [Column("updated_at")]
        public DateTime? UpdatedAt { get; set; }
    }

    public static class SupabaseDbService
    {
        private static Supabase.Client? _client;
        private static readonly SemaphoreSlim _initGate = new(1, 1);

        // THAY BẰNG URL/KEY CỦA BẠN (KHÔNG CÓ DẤU NGOẶC VUÔNG)
        private const string SUPABASE_URL = "https://valmbajjpkjccqslsuou.supabase.co";
        private const string SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZhbG1iYWpqcGtqY2Nxc2xzdW91Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzg2MDM5OTMsImV4cCI6MjA5NDE3OTk5M30.dwuPB1nlzNpFdWYR4fuvTOP7w6wB8U4fWE0cW_rOJ-o";

        public static string MachineId { get; } = LoadOrCreateMachineId();

        private static string LoadOrCreateMachineId()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "machine-id.txt");
            try
            {
                if (File.Exists(path)) return File.ReadAllText(path).Trim();
                string id = $"{Environment.MachineName}_{Guid.NewGuid():N}";
                File.WriteAllText(path, id);
                return id;
            }
            catch
            {
                return $"{Environment.MachineName}_{Guid.NewGuid():N}";
            }
        }

        public static async Task InitializeAsync()
        {
            if (_client != null) return;
            await _initGate.WaitAsync();
            try
            {
                if (_client != null) return;
                var options = new Supabase.SupabaseOptions { AutoConnectRealtime = true };
                _client = new Supabase.Client(SUPABASE_URL, SUPABASE_KEY, options);
                await _client.InitializeAsync();
            }
            finally { _initGate.Release(); }
        }

        private static async Task<Supabase.Client> GetClientAsync()
        {
            if (_client == null) await InitializeAsync();
            return _client!;
        }

        public static async Task<bool> TryAcquireInventoryLeaseAsync(int leaseSeconds = 1800)
        {
            var client = await GetClientAsync();
            return await client.RpcBoolAsync("try_acquire_inventory_lease", new
            {
                p_owner_id = MachineId,
                p_lease_seconds = leaseSeconds
            });
        }

        public static async Task<bool> RefreshInventoryLeaseAsync(int leaseSeconds = 1800)
        {
            var client = await GetClientAsync();
            return await client.RpcBoolAsync("refresh_inventory_lease", new
            {
                p_owner_id = MachineId,
                p_lease_seconds = leaseSeconds
            });
        }

        public static async Task<bool> ReleaseInventoryLeaseAsync()
        {
            var client = await GetClientAsync();
            return await client.RpcBoolAsync("release_inventory_lease", new { p_owner_id = MachineId });
        }

        public static async Task<bool> CompleteInventorySyncAsync()
        {
            var client = await GetClientAsync();
            return await client.RpcBoolAsync("complete_inventory_sync", new { p_owner_id = MachineId });
        }

        public static async Task<int> InsertNewWaybillsOnlyAsync(IEnumerable<string> fetchedWaybills)
        {
            var client = await GetClientAsync();
            var arr = fetchedWaybills?.Where(x => !string.IsNullOrWhiteSpace(x))
                                      .Select(x => x.Trim())
                                      .Distinct(StringComparer.OrdinalIgnoreCase)
                                      .ToArray() ?? Array.Empty<string>();

            if (arr.Length == 0) return 0;
            return await client.RpcIntAsync("upsert_new_waybills", new { p_waybills = arr });
        }

        public static async Task<int> UpsertManyWaybillsAsync(List<WaybillDbModel> rows)
        {
            var client = await GetClientAsync();
            if (rows == null || rows.Count == 0) return 0;

            var payload = rows.Select(x => new
            {
                waybill_no = x.WaybillNo,
                is_active = x.IsActive,
                tracking_interval_mins = x.TrackingIntervalMins,
                next_track_at = x.NextTrackAt,
                last_tracked_at = x.LastTrackedAt,
                trang_thai_hien_tai = x.TrangThaiHienTai,
                thao_tac_cuoi = x.ThaoTacCuoi,
                thoi_gian_thao_tac = x.ThoiGianThaoTac,
                nhan_vien_kien_van_de = x.NhanVienKienVanDe,
                nguyen_nhan_kien_van_de = x.NguyenNhanKienVanDe,
                buu_cuc_thao_tac = x.BuuCucThaoTac,
                nguoi_thao_tac = x.NguoiThaoTac,
                dau_chuyen_hoan = x.DauChuyenHoan,
                dia_chi_nhan_hang = x.DiaChiNhanHang,
                phuong = x.Phuong,
                noi_dung_hang_hoa = x.NoiDungHangHoa,
                cod_thuc_te = x.CODThucTe,
                pttt = x.PTTT,
                nhan_vien_nhan_hang = x.NhanVienNhanHang,
                dia_chi_lay_hang = x.DiaChiLayHang,
                thoi_gian_nhan_hang = x.ThoiGianNhanHang,
                ten_nguoi_gui = x.TenNguoiGui,
                trong_luong = x.TrongLuong,
                ma_doan_full = x.MaDoanFull,
                ma_doan_1 = x.MaDoan1,
                ma_doan_2 = x.MaDoan2,
                ma_doan_3 = x.MaDoan3,
                print_count = x.PrintCount,
                updated_at = x.UpdatedAt ?? DateTime.UtcNow
            }).ToArray();

            return await client.RpcIntAsync("merge_waybill_tracking_rows", new { p_rows = payload });
        }

        public static async Task<List<WaybillDbModel>> GetActiveWaybillsAsync(int pageSize = 500, int maxPages = 200)
        {
            var client = await GetClientAsync();
            var results = new List<WaybillDbModel>();
            int offset = 0;

            for (int page = 0; page < maxPages; page++)
            {
                var response = await client.From<WaybillDbModel>()
                    .Where(x => x.IsActive == true)
                    .Range(offset, offset + pageSize - 1)
                    .Get();

                var pageRows = response.Models?.ToList() ?? new List<WaybillDbModel>();
                if (pageRows.Count == 0) break;
                
                results.AddRange(pageRows);
                if (pageRows.Count < pageSize) break;
                offset += pageSize;
            }
            return results;
        }

        public static async Task<List<string>> GetWaybillsDueForTrackingAsync(int pageSize = 500, int maxPages = 200)
        {
            var client = await GetClientAsync();
            var results = new List<string>();
            int offset = 0;
            var now = DateTime.UtcNow;

            for (int page = 0; page < maxPages; page++)
            {
                var response = await client.From<WaybillDbModel>()
                    .Where(x => x.IsActive == true)
                    .Where(x => x.NextTrackAt <= now)
                    .Select("waybill_no")
                    .Range(offset, offset + pageSize - 1)
                    .Get();

                var pageRows = response.Models?.Select(x => x.WaybillNo).Where(x => !string.IsNullOrWhiteSpace(x)).ToList() ?? new List<string>();
                if (pageRows.Count == 0) break;
                
                results.AddRange(pageRows);
                if (pageRows.Count < pageSize) break;
                offset += pageSize;
            }

            return results.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
    }

    public static class SupabaseRpcHelper
    {
        public static async Task<bool> RpcBoolAsync(this Supabase.Client client, string functionName, object? args = null)
        {
            BaseResponse response = await client.Rpc(functionName, args);
            return ParseBool(ReadContent(response));
        }

        public static async Task<int> RpcIntAsync(this Supabase.Client client, string functionName, object? args = null)
        {
            BaseResponse response = await client.Rpc(functionName, args);
            return ParseInt(ReadContent(response));
        }

        public static string? ReadContent(object? response)
        {
            if (response == null) return null;
            try
            {
                var t = response.GetType();
                var contentProp = t.GetProperty("Content", BindingFlags.Public | BindingFlags.Instance);
                if (contentProp != null)
                {
                    var value = contentProp.GetValue(response);
                    if (value is string s) return s;
                    return value?.ToString();
                }
                var jsonProp = t.GetProperty("Response", BindingFlags.Public | BindingFlags.Instance);
                if (jsonProp != null) return jsonProp.GetValue(response)?.ToString();
            }
            catch { }
            return response.ToString();
        }

        public static bool ParseBool(string? content)
        {
            if (string.IsNullOrWhiteSpace(content)) return false;
            content = content.Trim();
            if (bool.TryParse(content, out bool b)) return b;
            try
            {
                using var doc = JsonDocument.Parse(content);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.True) return true;
                if (root.ValueKind == JsonValueKind.False) return false;
                if (root.ValueKind == JsonValueKind.String && bool.TryParse(root.GetString(), out b)) return b;
            }
            catch { }
            return content.Contains("true", StringComparison.OrdinalIgnoreCase);
        }

        public static int ParseInt(string? content)
        {
            if (string.IsNullOrWhiteSpace(content)) return 0;
            content = content.Trim();
            if (int.TryParse(content, out int n)) return n;
            try
            {
                using var doc = JsonDocument.Parse(content);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Number && root.TryGetInt32(out n)) return n;
                if (root.ValueKind == JsonValueKind.String && int.TryParse(root.GetString(), out n)) return n;
            }
            catch { }
            return 0;
        }
    }
}