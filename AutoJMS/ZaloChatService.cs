using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;

namespace AutoJMS
{
    public class ZaloChatService : IDisposable
    {
        private readonly WebView2 _webView;
        private readonly string _appsScriptUrl;
        private readonly HttpClient _httpClient;
        private readonly System.Windows.Forms.Timer _timer;
        private readonly JsonSerializerOptions _jsonOptions;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        private bool _isProcessing = false;
        private bool _isDisposed = false;

        public ZaloChatService(WebView2 webView, string appsScriptUrl)
        {
            _webView = webView;
            _appsScriptUrl = appsScriptUrl;
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };

            _jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };

            _timer = new System.Windows.Forms.Timer();
            _timer.Tick += async (s, e) => await Timer_TickAsync();
        }

        public void StartAutoReminder(int intervalSeconds = 300)
        {
            if (intervalSeconds < 5) intervalSeconds = 300;
            _timer.Interval = intervalSeconds * 1000;
            _timer.Start();
            Console.WriteLine($"[ZaloBot] ✅ Bot đã chạy - Nhắc nhở mỗi {intervalSeconds} giây");
        }

        public void StopAutoReminder()
        {
            _timer.Stop();
            Console.WriteLine("[ZaloBot] ⛔ Bot đã dừng");
        }

        private async Task Timer_TickAsync()
        {
            if (_isProcessing || _isDisposed) return;

            try
            {
                _isProcessing = true;
                await ProcessRemindersAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloBot] ❌ Lỗi Timer: {ex.Message}");
            }
            finally
            {
                _isProcessing = false;
            }
        }

        private async Task ProcessRemindersAsync()
        {
            try
            {
                var reminders = await GetDataFromSheetAsync();

                if (reminders.Count == 0)
                {
                    Console.WriteLine("[ZaloBot] Không có đơn nào cần nhắc.");
                    return;
                }

                Console.WriteLine($"[ZaloBot] Đang xử lý {reminders.Count} đơn...");

                foreach (var item in reminders)
                {
                    if (_cts.Token.IsCancellationRequested) break;

                    string message = $"{item.maDon}\n@{item.nhanVien}";

                    bool sent = await SendZaloMessageJSAsync(message);

                    if (sent)
                    {
                        await UpdateReminderStatusAsync(item.row);
                        Console.WriteLine($"[ZaloBot] ✅ Đã gửi: {item.maDon}");
                        await Task.Delay(2000, _cts.Token);
                    }
                    else
                    {
                        Console.WriteLine($"[ZaloBot] ⚠️ Không gửi được: {item.maDon}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloBot] ❌ ProcessReminders lỗi: {ex.Message}");
            }
        }

        public async Task<List<Reminder>> GetDataFromSheetAsync()
        {
            try
            {
                Console.WriteLine("[ZaloBot] Đang đọc danh sách từ cột B (PHATLAI!B2:B)...");

                // Đọc trực tiếp từ cột B của sheet PHATLAI
                var data = GoogleSheetService.ReadRange(GoogleSheetService.DATA_SPREADSHEET_ID, "PHATLAI!B2:B");

                List<Reminder> reminders = new List<Reminder>();

                if (data != null)
                {
                    foreach (var row in data)
                    {
                        if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0]?.ToString()))
                        {
                            string maDon = row[0].ToString().Trim();

                            reminders.Add(new Reminder
                            {
                                     // row giả để update sau (nếu cần)
                                maDon = maDon,
                                nhanVien = "",                  // tạm để trống (bạn có thể bổ sung sau)
                                trangThai = "Đang theo dõi",    // trạng thái mặc định
                                soLanNhac = 0,
                            });
                        }
                    }
                }

                Console.WriteLine($"[ZaloBot] ✅ Đã lấy được {reminders.Count} vận đơn từ cột B");

                return reminders;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloBot] ❌ Lỗi đọc cột B: {ex.Message}");
                return new List<Reminder>();
            }
        }

        private async Task<bool> SendZaloMessageJSAsync(string message)
        {
            if (_webView?.CoreWebView2 == null) return false;

            string safeMessage = message
                .Replace("\\", "\\\\")
                .Replace("'", "\\'")
                .Replace("\n", "\\n")
                .Replace("\r", "");

            string script = $@"
                try {{
                    var input = document.getElementById('richInput');
                    if (!input) return 'false';
                    input.focus();
                    document.execCommand('insertText', false, '{safeMessage}');
                    
                    setTimeout(() => {{
                        var enterEvent = new KeyboardEvent('keydown', {{
                            bubbles: true, cancelable: true, keyCode: 13, which: 13,
                            key: 'Enter', code: 'Enter'
                        }});
                        input.dispatchEvent(enterEvent);
                    }}, 120);
                    return 'true';
                }} catch (e) {{ return 'false'; }}
            ";

            try
            {
                var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);
                return result?.Contains("true") == true;
            }
            catch
            {
                return false;
            }
        }

        private async Task UpdateReminderStatusAsync(int row)
        {
            try
            {
                var payload = new { row };
                var json = JsonSerializer.Serialize(payload);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                await _httpClient.PostAsync(_appsScriptUrl, content);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloBot] Không update row {row}: {ex.Message}");
            }
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _cts.Cancel();
            _timer.Stop();
            _timer.Dispose();
            _httpClient.Dispose();
            _cts.Dispose();
            _isDisposed = true;
        }
    }

    /// <summary>
    /// Model hứng dữ liệu JSON từ Google Apps Script
    /// </summary>
    public class Reminder
    {
        // LƯU Ý: Phải có chữ 'public', kiểu dữ liệu, tên biến, và { get; set; } viết liền nhau
        public int row { get; set; }
        public string maDon { get; set; }
        public string nhanVien { get; set; }
        public string trangThai { get; set; }
        public int soLanNhac { get; set; }
        public string thoiGianNhac { get; set; }
    }
}