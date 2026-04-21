using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using System.Linq;

namespace AutoJMS
{
    public class ZaloChatService
    {
        private readonly WebView2 _webView;
        private readonly string _appsScriptUrl;
        private readonly HttpClient _httpClient;
        private System.Windows.Forms.Timer _timer;
        private bool _isProcessing = false;

        public ZaloChatService(WebView2 webView, string appsScriptUrl)
        {
            _webView = webView;
            _appsScriptUrl = appsScriptUrl;
            _httpClient = new HttpClient();

            _timer = new System.Windows.Forms.Timer();
            _timer.Tick += async (s, e) => await Timer_Tick();
        }

        public void StartAutoReminder(int intervalMinutes = 5)
        {
            _timer.Interval = intervalMinutes * 60 * 1000;
            _timer.Start();
            Console.WriteLine($"[ZaloService] Đã bắt đầu tiến trình nhắc nhở tự động ({intervalMinutes} phút/lần).");
        }

        public void StopAutoReminder()
        {
            _timer.Stop();
            Console.WriteLine("[ZaloService] Đã dừng tiến trình nhắc nhở tự động.");
        }

        private async Task Timer_Tick()
        {
            if (_isProcessing) return;

            try
            {
                _isProcessing = true;
                await ProcessReminders();
            }
            catch (Exception ex)
            {
                // Bạn có thể ghi log ra file hoặc UI tại đây
                Console.WriteLine($"[ZaloService] Lỗi Timer: {ex.Message}");
            }
            finally
            {
                _isProcessing = false;
            }
        }

        /// <summary>
        /// Lấy dữ liệu Reminder từ Google Sheet theo yêu cầu mới:
        /// - BUMP!A:F     → maDon(A), nhanVien(F), trangThai(C)
        /// - PHATLAI!E    → soLanNhac (khớp theo maDon)
        /// </summary>
        /// <summary>
        /// DEBUG SIÊU CHI TIẾT - Lấy dữ liệu từ BUMP + PHATLAI
        /// </summary>
        /// <summary>
        /// DEBUG SIÊU CHI TIẾT - Bắt Google.GoogleApiException
        /// </summary>
        public async Task<List<Reminder>> GetDataFromSheetAsync()
        {
            try
            {
                string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;
                Console.WriteLine($"[DEBUG] Spreadsheet ID = {spreadsheetId}");

                Console.WriteLine("[DEBUG] Đang đọc BUMP!A2:F ...");
                var bumpData = GoogleSheetService.ReadRange(spreadsheetId, "BUMP!A2:F");
                Console.WriteLine($"[DEBUG] BUMP đọc được: {(bumpData?.Count ?? 0)} dòng");

                if (bumpData == null || bumpData.Count == 0)
                {
                    Console.WriteLine("❌ Sheet BUMP trống hoặc không đọc được!");
                    return new List<Reminder>();
                }

                // In 3 dòng đầu để kiểm tra dữ liệu
                Console.WriteLine("[DEBUG] 3 dòng đầu BUMP:");
                for (int i = 0; i < Math.Min(3, bumpData.Count); i++)
                {
                    var row = bumpData[i];
                    string line = string.Join(" | ", row.Select(c => c?.ToString() ?? "NULL"));
                    Console.WriteLine($"   Dòng {i + 2}: {line}");
                }

                // Đọc PHATLAI
                var phatLaiData = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!A2:E");
                Console.WriteLine($"[DEBUG] PHATLAI đọc được: {(phatLaiData?.Count ?? 0)} dòng");

                // Xử lý dữ liệu...
                var reminders = new List<Reminder>();
                int rowNumber = 2;
                var phatLaiDict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                if (phatLaiData != null)
                {
                    foreach (var row in phatLaiData)
                    {
                        if (row.Count > 4)
                        {
                            string maDon = row[0]?.ToString()?.Trim();
                            if (!string.IsNullOrEmpty(maDon) && int.TryParse(row[4]?.ToString(), out int soLan))
                                phatLaiDict[maDon] = soLan;
                        }
                    }
                }

                foreach (var row in bumpData)
                {
                    if (row.Count < 6) continue;
                    string maDon = row[0]?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(maDon)) continue;

                    string nhanVien = row[5]?.ToString()?.Trim() ?? "";
                    string trangThai = row[2]?.ToString()?.Trim() ?? "";
                    int soLanNhac = phatLaiDict.TryGetValue(maDon, out int sl) ? sl : 0;

                    reminders.Add(new Reminder
                    {
                        row = rowNumber,
                        maDon = maDon,
                        nhanVien = nhanVien,
                        trangThai = trangThai,
                        soLanNhac = soLanNhac,
                        thoiGianNhac = ""
                    });
                    rowNumber++;
                }

                Console.WriteLine($"✅ [SUCCESS] Load được {reminders.Count} đơn!");
                return reminders;
            }
            catch (Google.GoogleApiException gex)
            {
                Console.WriteLine("❌ === GOOGLE API EXCEPTION ===");
                Console.WriteLine($"Error Code     : {gex.Error?.Code}");
                Console.WriteLine($"Error Message  : {gex.Error?.Message}");
                Console.WriteLine($"Http Status    : {gex.HttpStatusCode}");
                Console.WriteLine($"Service Name   : {gex.ServiceName}");
                Console.WriteLine($"Full Message   : {gex.Message}");
                Console.WriteLine("StackTrace:");
                Console.WriteLine(gex.StackTrace);
                return new List<Reminder>();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ [OTHER EXCEPTION] {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return new List<Reminder>();
            }
        }

        //=======================================================================================================
        /// <summary>
        /// Kéo dữ liệu từ Google Sheet và đổ trực tiếp lên DataGridView truyền vào
        /// </summary>
        /// <param name="grid">Control DataGridView trên form</param>
        public async Task LoadDataToGridAsync(DataGridView grid)
        {
            if (grid == null) return;

            try
            {
                // 1. Kéo data gốc từ Google Sheets (Kiểu List<Reminder>)
                var data = await GetDataFromSheetAsync();

                // 2. Nếu data null hoặc rỗng thì cảnh báo để ta biết
                if (data == null || data.Count == 0)
                {
                    MessageBox.Show("Google Sheets trống!", "Thông báo");
                    return;
                }

                // 3. Đổ trực tiếp List<Reminder> vào bảng
                Action updateUI = new Action(() => {
                    grid.AutoGenerateColumns = false;     // Tắt đẻ cột bậy bạ
                    grid.AllowUserToAddRows = false;      // Tắt dòng rỗng cuối cùng
                    grid.DataSource = null;               // Reset data cũ (nếu có)
                    grid.DataSource = data;               // Gán data mới
                });

                if (grid.InvokeRequired)
                {
                    grid.Invoke(updateUI);
                }
                else
                {
                    updateUI();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi hiển thị dữ liệu Zalo: {ex.Message}");
            }
        }
        public async Task ProcessReminders()
        {
            try
            {
                // 1. Gọi API Get để lấy danh sách đơn cần nhắc
                var response = await _httpClient.GetStringAsync(_appsScriptUrl);
                var reminders = JsonSerializer.Deserialize<List<Reminder>>(response);

                if (reminders == null || reminders.Count == 0)
                {
                    return; // Không có đơn nào cần nhắc
                }

                foreach (var item in reminders)
                {
                    // Tạo nội dung tin nhắn
                    string message = $"{item.maDon} \n @{item.nhanVien}";

                    // 2. Bắn JS vào Zalo Web để gửi tin nhắn
                    bool sent = await SendZaloMessageJS(message);

                    if (sent)
                    {
                        // 3. Gọi API Post để cập nhật thời gian "Lần Nhắc Cuối" nếu gửi thành công
                        var payload = new { row = item.row };
                        var jsonPayload = JsonSerializer.Serialize(payload);
                        var content = new StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json");

                        await _httpClient.PostAsync(_appsScriptUrl, content);

                        // Nghỉ 2 giây giữa các tin nhắn để tránh Zalo chặn vì spam tốc độ cao
                        await Task.Delay(2000);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloService] Lỗi trong quá trình lấy data/gửi tin: {ex.Message}");
            }
        }

        /// <summary>
        /// Bắn Javascript vào WebView2 để giả lập thao tác nhập và gửi
        /// </summary>
        private async Task<bool> SendZaloMessageJS(string message)
        {
            if (_webView == null || _webView.CoreWebView2 == null) return false;

            // Xử lý chuỗi để tránh làm gãy cú pháp Javascript (chống lỗi nháy đơn, xuống dòng)
            string safeMessage = message.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\n", "\\n").Replace("\r", "");

            // Script tối ưu dùng execCommand để bypass ReactJS
            string script = $@"
                try {{
                    var input = document.getElementById('richInput'); 
                    if (!input) return 'false';
                    
                    // 1. Focus vào ô nhập liệu
                    input.focus();
                    
                    // 2. Chèn text bằng execCommand (giả lập thao tác Paste/Ctrl+V)
                    document.execCommand('insertText', false, '{safeMessage}');

                    // 3. Chờ 100ms để ReactJS của Zalo cập nhật trạng thái nút Gửi
                    setTimeout(function() {{
                        var enterEvent = new KeyboardEvent('keydown', {{
                            bubbles: true,
                            cancelable: true,
                            keyCode: 13,
                            which: 13,
                            key: 'Enter',
                            code: 'Enter'
                        }});
                        input.dispatchEvent(enterEvent);
                    }}, 100);

                    return 'true';
                }} catch (e) {{
                    return 'false';
                }}
            ";

            var result = await _webView.CoreWebView2.ExecuteScriptAsync(script);

            // WebView2 trả về chuỗi JSON (ví dụ: "\"true\""), nên cần kiểm tra Contains
            return result != null && result.Contains("true");
        }
        /// <summary>
        /// LẤY DANH SÁCH MÃ ĐƠN TỪ PHATLAI!A2:A (chỉ cột A)
        /// </summary>
        public List<string> GetWaybillsFromPhatLai()
        {
            try
            {
                var data = GoogleSheetService.ReadRange(
                    GoogleSheetService.DATA_SPREADSHEET_ID,
                    "PHATLAI!A2:A");

                return data
                    .SelectMany(row => row)
                    .Select(cell => cell?.ToString()?.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)   // tránh trùng lặp
                    .ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ZaloService] Lỗi đọc danh sách mã đơn từ PHATLAI: {ex.Message}");
                return new List<string>();
            }
        }
        // Thêm vào cuối class ZaloChatService
        public static List<string> GetWaybillsFromPhatLaiStatic()
        {
            try
            {
                var data = GoogleSheetService.ReadRange(
                    GoogleSheetService.DATA_SPREADSHEET_ID, "PHATLAI!A2:A");

                return data
                    .SelectMany(row => row)
                    .Select(cell => cell?.ToString()?.Trim())
                    .Where(s => !string.IsNullOrEmpty(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }
    }



    /// <summary>
    /// Model hứng dữ liệu JSON trả về từ Google Apps Script
    /// </summary>
    public class Reminder
    {
        public int row { get; set; }
        public string maDon { get; set; }
        public string nhanVien { get; set; }
        public string trangThai { get; set; }
        public int soLanNhac { get; set; }
        public string thoiGianNhac { get; set; }
    }
}