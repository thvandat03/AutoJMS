using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

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
        /// Hàm cốt lõi: Lấy dữ liệu từ Sheet và gửi tin nhắn
        /// </summary>
        /// 
        /// <summary>
        /// Hàm này chỉ dùng để kéo dữ liệu mới nhất từ Sheet về (Dùng cho nút Làm mới Grid)
        /// </summary>
        public async Task<List<Reminder>> GetDataFromSheetAsync()
        {
            try
            {
                var response = await _httpClient.GetStringAsync(_appsScriptUrl);

                // Sử dụng options để bỏ qua phân biệt hoa thường khi map JSON
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                var reminders = JsonSerializer.Deserialize<List<Reminder>>(response, options);

                return reminders ?? new List<Reminder>();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi tải dữ liệu từ Sheet: {ex.Message}");
                return new List<Reminder>();
            }
        }
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