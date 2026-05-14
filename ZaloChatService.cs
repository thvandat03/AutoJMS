using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization; 
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    public class ZaloChatService
    {
        private readonly WebView2 _webView;
        private readonly string _appsScriptUrl;
        private readonly HttpClient _httpClient;
        private System.Windows.Forms.Timer _timer;

        public ZaloChatService(WebView2 webView, string appsScriptUrl)
        {
            _webView = webView;
            _appsScriptUrl = appsScriptUrl;
            _httpClient = new HttpClient();

            _timer = new System.Windows.Forms.Timer();
            // Timer này có thể dùng để gọi API Bot tương lai
            // Hiện tại Main.cs đang chịu trách nhiệm lấy Data
        }

        public void StartAutoReminder(int intervalMinutes = 5)
        {
            _timer.Interval = intervalMinutes * 60 * 1000;
            _timer.Start();
            Console.WriteLine($"[ZaloService] Đã bắt đầu ({intervalMinutes} phút/lần).");
        }

        public void StopAutoReminder()
        {
            _timer.Stop();
        }

        public async Task<bool> SendZaloMessage(string message)
        {
            if (_webView == null || _webView.CoreWebView2 == null) return false;

            try
            {
                string[] parts = message.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                string tenNV = "";
                var dsMa = new List<string>();
                foreach (var p in parts)
                {
                    string trimmed = p.Trim();
                    if (trimmed.StartsWith("@")) tenNV = trimmed.Substring(1).Trim();
                    else if (!string.IsNullOrWhiteSpace(trimmed)) dsMa.Add(trimmed);
                }
                if (dsMa.Count == 0 || string.IsNullOrEmpty(tenNV)) return false;

                string dsMaStr = string.Join("\n", dsMa);
                string tenNVJson = JsonSerializer.Serialize(tenNV);

                string script = $@"
                (async function() {{
                    try {{
                        var ten = {tenNVJson};
                        var dsMaText = ' \n' + `{dsMaStr.Replace("'", "\\'").Replace("\n", "\\n")}`;
                        var input = document.querySelector('#richInput');
                        if (!input) return JSON.stringify({{ok:false, error:'Không tìm thấy ô nhập liệu'}});

                        input.focus();
                        if (input.contentEditable !== 'true') input.contentEditable = 'true';

                        document.execCommand('selectAll', false, null);
                        document.execCommand('delete', false, null);
                        document.execCommand('insertText', false, '@' + ten);
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        
                        await new Promise(r => setTimeout(r, 800));

                        var popup = document.querySelector('#mentionPopover');
                        if (popup) {{
                            var items = popup.querySelectorAll('.mention-popover__item');
                            if (items.length > 0) {{
                                var target = ten.toLowerCase().trim();
                                var isClicked = false;
                                for (var i = 0; i < items.length; i++) {{
                                    var nameEl = items[i].querySelector('.tg-name');
                                    if (nameEl && nameEl.innerText.toLowerCase().includes(target)) {{
                                        items[i].dispatchEvent(new MouseEvent('mousedown', {{ bubbles: true, cancelable: true }}));
                                        items[i].click();
                                        isClicked = true;
                                        break;
                                    }}
                                }}
                                if (!isClicked) {{
                                    items[0].dispatchEvent(new MouseEvent('mousedown', {{ bubbles: true, cancelable: true }}));
                                    items[0].click();
                                }}
                            }}
                        }}

                        await new Promise(r => setTimeout(r, 500));
                        input.focus();
                        var sel = window.getSelection();
                        var range = document.createRange();
                        range.selectNodeContents(input);
                        range.collapse(false);
                        sel.removeAllRanges();
                        sel.addRange(range);

                        document.execCommand('insertText', false, ' \n');

                        var pasteEvent = new ClipboardEvent('paste', {{
                            bubbles: true, cancelable: true, clipboardData: new DataTransfer()
                        }});
                        pasteEvent.clipboardData.setData('text/plain', dsMaText);
                        input.dispatchEvent(pasteEvent);
                        input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        
                        await new Promise(r => setTimeout(r, 500));
                        var btn = document.querySelector('div.send-msg-btn, div[icon=""Sent-msg_24_Line""]');
                        if (btn) btn.click();

                        return JSON.stringify({{ok:true}});
                    }} catch(e) {{
                        return JSON.stringify({{ok:false, error: e.message}});
                    }}
                }})();";

                string result = await _webView.CoreWebView2.ExecuteScriptAsync(script);
                return result != null && result.Contains("true");
            }
            catch { return false; }
        }
    }

    public class Reminder
    {
        [JsonPropertyName("row")] public int row { get; set; }
        [JsonPropertyName("maDon")] public string maDon { get; set; }
        [JsonPropertyName("nhanVien")] public string nhanVien { get; set; }
        [JsonPropertyName("trangThai")] public string trangThai { get; set; }
        [JsonPropertyName("soLanNhac")] public int soLanNhac { get; set; }
        [JsonPropertyName("thoiGianNhac")] public string thoiGianNhac { get; set; }
    }
}