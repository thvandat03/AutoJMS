using AutoJMS.Data;
using Postgrest.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class SupabaseDbService
    {
        private static Supabase.Client _client;
        private static readonly SemaphoreSlim _initGate = new(1, 1);

        private const string SUPABASE_URL = "https://[PROJECT_ID].supabase.co";
        private const string SUPABASE_KEY = "eyJh[YOUR_ANON_KEY]...";

        public static string MachineId { get; private set; } = LoadOrCreateMachineId();

        private static string LoadOrCreateMachineId()
        {
            return Environment.MachineName + "_" + Guid.NewGuid().ToString("N");
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
            finally
            {
                _initGate.Release();
            }
        }

        private static async Task<T> RpcAsync<T>(string fn, object args)
        {
            if (_client == null) await InitializeAsync();
            var res = await _client.Rpc(fn, args);
            // Parse chuỗi trả về thành kiểu T
            if (typeof(T) == typeof(bool)) return (T)(object)(res.Content != null && res.Content.Trim() == "true");
            if (typeof(T) == typeof(int)) return (T)(object)(int.Parse(res.Content ?? "0"));
            return default;
        }

        public static Task<bool> TryAcquireInventoryLeaseAsync(int leaseSeconds = 1800) => 
            RpcAsync<bool>("try_acquire_inventory_lease", new { p_owner_id = MachineId, p_lease_seconds = leaseSeconds });

        public static Task<bool> RefreshInventoryLeaseAsync(int leaseSeconds = 1800) => 
            RpcAsync<bool>("refresh_inventory_lease", new { p_owner_id = MachineId, p_lease_seconds = leaseSeconds });

        public static Task<bool> ReleaseInventoryLeaseAsync() => 
            RpcAsync<bool>("release_inventory_lease", new { p_owner_id = MachineId });

        public static Task<bool> CompleteInventorySyncAsync() => 
            RpcAsync<bool>("complete_inventory_sync", new { p_owner_id = MachineId });

        public static async Task<int> UpsertNewWaybillsOnlyAsync(IEnumerable<string> fetchedWaybills)
        {
            if (_client == null) await InitializeAsync();
            var arr = fetchedWaybills?.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToArray() ?? Array.Empty<string>();
            if (arr.Length == 0) return 0;
            return await RpcAsync<int>("upsert_new_waybills", new { p_waybills = arr });
        }

        public static async Task<List<WaybillDbModel>> GetActiveWaybillsAsync()
        {
            if (_client == null) await InitializeAsync();
            var response = await _client.From<WaybillDbModel>().Where(x => x.IsActive == true).Get();
            return response.Models?.ToList() ?? new List<WaybillDbModel>();
        }

        public static async Task<List<string>> GetWaybillsDueForTrackingAsync()
        {
            if (_client == null) await InitializeAsync();
            var response = await _client.From<WaybillDbModel>()
                .Where(x => x.IsActive == true)
                .Where(x => x.NextTrackAt <= DateTime.UtcNow)
                .Select("waybill_no").Get();
            return response.Models.Select(x => x.WaybillNo).Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
        }

        public static async Task UpsertManyWaybillsAsync(List<WaybillDbModel> rows)
        {
            if (_client == null) await InitializeAsync();
            if (rows == null || rows.Count == 0) return;
            await _client.From<WaybillDbModel>().Upsert(rows);
        }
    }
}