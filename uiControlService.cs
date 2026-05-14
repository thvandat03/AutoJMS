using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;

namespace AutoJMS
{
    public sealed class uiControlService
    {
        private readonly TimeSpan _interval = TimeSpan.FromSeconds(1);

        private static readonly HttpClient _httpClient;
        private const string NETWORK_TEST_URL = "http://clients3.google.com/generate_204";

        static uiControlService()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromMilliseconds(2000);
        }

        public async Task StartAsync(CancellationToken ct = default)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    if (!NetworkInterface.GetIsNetworkAvailable())
                    {
                        NetworkState.Set(NetworkStatus.Offline);
                    }
                    else
                    {
                        var sw = Stopwatch.StartNew();
                        using var request = new HttpRequestMessage(HttpMethod.Head, NETWORK_TEST_URL);

                        var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, ct);
                        sw.Stop();

                        if (response.IsSuccessStatusCode)
                        {
                            if (sw.ElapsedMilliseconds <= 400)
                            {
                                NetworkState.Set(NetworkStatus.Online);
                            }
                            else
                            {
                                NetworkState.Set(NetworkStatus.Unstable);
                            }
                        }
                        else
                        {
                            NetworkState.Set(NetworkStatus.Offline);
                        }
                    }
                }
                catch
                {
                    NetworkState.Set(NetworkStatus.Offline);
                }

                // Tổng thời gian chờ mỗi vòng = 1s + (0.1s -> 0.5s)
                int jitterMs = Random.Shared.Next(100, 500);
                await Task.Delay(_interval + TimeSpan.FromMilliseconds(jitterMs), ct);
            }
        }
    }

    // ==========================================
    // CÁC ENUM VÀ STATE QUẢN LÝ TRẠNG THÁI
    // ==========================================
    public enum NetworkStatus
    {
        Online,
        Unstable,
        Offline
    }

    public static class NetworkState
    {
        public static NetworkStatus Current { get; private set; } = NetworkStatus.Online;
        public static event Action<NetworkStatus> OnChanged;

        public static void Set(NetworkStatus status)
        {
            if (Current != status)
            {
                Current = status;
                OnChanged?.Invoke(status);
            }
        }
    }
}