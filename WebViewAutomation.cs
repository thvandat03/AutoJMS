using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    public class NoDataWaybillException : Exception { }
    public class NeedSwitchToDkch1Exception : Exception { }
    public class NeedSwitchToDkch2Exception : Exception { }


    public static class WebViewAutomation
    {
        private static async Task<string> ExecuteScriptSafeAsync(WebView2 webView, string script)
        {
            if (webView.InvokeRequired)
                return await (Task<string>)webView.Invoke(new Func<Task<string>>(async () => await webView.ExecuteScriptAsync(script)));
            return await webView.ExecuteScriptAsync(script);
        }

        private static async Task<string> WaitForApiResponseAsync(WebView2 webView, Func<string, bool> predicate, int timeoutMs = 10000)
        {
            var tcs = new TaskCompletionSource<string>();
            EventHandler<CoreWebView2WebResourceResponseReceivedEventArgs> handler = async (sender, args) =>
            {
                if (args.Response.StatusCode == 200)
                {
                    try
                    {
                        var stream = await args.Response.GetContentAsync();
                        if (stream != null)
                        {
                            using (var reader = new StreamReader(stream))
                            {
                                string json = await reader.ReadToEndAsync();
                                if (predicate(json))
                                    tcs.TrySetResult(json);
                            }
                        }
                    }
                    catch { }
                }
            };

            if (webView.InvokeRequired)
                webView.Invoke(new Action(() => webView.CoreWebView2.WebResourceResponseReceived += handler));
            else
                webView.CoreWebView2.WebResourceResponseReceived += handler;

            var completedTask = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs));

            if (webView.InvokeRequired)
                webView.Invoke(new Action(() => webView.CoreWebView2.WebResourceResponseReceived -= handler));
            else
                webView.CoreWebView2.WebResourceResponseReceived -= handler;

            if (completedTask != tcs.Task)
                throw new TimeoutException("Timeout: Không có phản hồi");

            return await tcs.Task;
        }

        private static void CheckAndThrowIfError(string json)
        {
            if (json.Contains("\"succ\":false") || json.Contains("\"fail\":true"))
            {
                Exception exToThrow = null;
                try
                {
                    // Chỉ dùng try-catch để an toàn khi phân tích JSON
                    using (JsonDocument doc = JsonDocument.Parse(json))
                    {
                        var root = doc.RootElement;
                        string msg = root.TryGetProperty("msg", out var m) ? m.GetString() : "";
                        string code = root.TryGetProperty("code", out var c) ? c.ToString() : "";

                        if (msg.Contains("Chưa có") || msg.Contains("hoàn lần 2") || code == "999006328")
                            exToThrow = new NeedSwitchToDkch1Exception();
                        else if (code == "137043004" || code == "999006082" || code.Contains(":"))
                            exToThrow = new NeedSwitchToDkch2Exception();
                        else if (msg.Contains("không có dữ liệu") || msg.Contains("Vận đơn không tồn tại"))
                            exToThrow = new NoDataWaybillException();
                        else if (!string.IsNullOrEmpty(msg))
                            exToThrow = new Exception($"Lỗi: {msg}");
                    }
                }
                catch { } // Bỏ qua nếu lỗi format JSON

                // Ném lỗi một cách gọn gàng ở ngoài khối try-catch
                if (exToThrow != null)
                {
                    throw exToThrow;
                }
            }
        }

        public static async Task FillWaybillAsync(WebView2 webView, string waybill, CancellationToken token)
        {
            string collapseJs = @"(function() {
                var headers = document.querySelectorAll('.el-collapse-item__header.is-active');
                for (var i = 0; i < headers.length; i++) {
                    var text = headers[i].innerText.trim();
                    if (text.includes('Thông tin người gửi') || text.includes('hóa đơn gốc')) {
                        headers[i].click();
                    }
                }
            })();"; 
            await ExecuteScriptSafeAsync(webView, collapseJs);
            string js = $@"
            (function() {{
                var container = document.querySelector('div[id^=""el-collapse-content-""]');
                var input = container ? container.querySelector('input') : null;
                if (!input) {{
                    var inputs = document.querySelectorAll('input[type=text], input:not([type])');
                    for(var i=0; i<inputs.length; i++) {{
                        if(inputs[i].offsetParent !== null && !inputs[i].disabled) {{
                            input = inputs[i]; break;
                        }}
                    }}
                }}
                if (input) {{
                    if (input.value.trim().toUpperCase() === '{waybill}'.toUpperCase()) return 'already_filled';
                    
                    // Ép VueJS nhận diện sự thay đổi
                    var nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    nativeInputValueSetter.call(input, '{waybill}');
                    
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    input.blur();
                    return 'filled';
                }}
                return 'not_found';
            }})();";

            string res = await ExecuteScriptSafeAsync(webView, js);
            if (res.Contains("not_found")) throw new Exception("Không tìm thấy ô nhập mã vận đơn.");
        }

        public static async Task CheckAndSelectDropdownAsync(WebView2 webView, string targetOption, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            string checkJs = @"(function() {
                var input = document.querySelector('.el-select .el-input__inner');
                return input ? input.value : '';
            })();";
            string currentJson = await ExecuteScriptSafeAsync(webView, checkJs);
            string currentValue = currentJson.Trim('"');

            if (currentValue.Equals(targetOption, StringComparison.OrdinalIgnoreCase))
                return;

            //string selectJs = $@"
            //(async function() {{
            //    let ddInput = document.querySelector('.el-select .el-input__inner[readonly]');
            //    if (!ddInput) return 'no_input';
            //    ddInput.click();
            //    let maxRetries = 20;
            //    while(maxRetries > 0) {{
            //        await new Promise(r => setTimeout(r, 100));
            //        let items = document.querySelectorAll('li.el-select-dropdown__item span');
            //        for (let span of items) {{
            //            if (span.innerText.trim() === '{targetOption}') {{
            //                span.parentElement.click();
            //                return 'ok';
            //            }}
            //        }}
            //        maxRetries--;
            //    }}
            //    return 'item_not_found';
            //}})();";
            string selectJs = $@"
            (async function() {{
                try {{
                    // 1. Tìm và mở Dropdown
                    let ddInput = document.querySelector('.el-select .el-input__inner[readonly]');
                    if (!ddInput) return 'no_input';
                    
                    ddInput.dispatchEvent(new MouseEvent('mousedown', {{ bubbles: true }}));
                    ddInput.click();

                    // 2. Chờ Dropdown render
                    let maxRetries = 20;
                    while(maxRetries > 0) {{
                        await new Promise(r => setTimeout(r, 100));
                        
                        let visibleDropdowns = document.querySelectorAll('.el-select-dropdown:not([style*=""display: none""])');
                        
                        for (let dd of visibleDropdowns) {{
                            let items = dd.querySelectorAll('li.el-select-dropdown__item');
                            
                            for (let item of items) {{
                                
                                // CHỈNH SỬA TẠI ĐÂY: Dùng so sánh tuyệt đối (===) thay vì .includes()
                                let currentText = item.innerText.toLowerCase().trim();
                                let targetText = '{targetOption}'.toLowerCase().trim();

                                if (currentText === targetText && !item.classList.contains('is-disabled')) {{
                                    
                                    item.scrollIntoView({{ block: 'center' }});
                                    
                                    item.dispatchEvent(new MouseEvent('mouseenter', {{ bubbles: true }}));
                                    item.dispatchEvent(new MouseEvent('mousedown', {{ bubbles: true }}));
                                    item.click();
                                    item.dispatchEvent(new MouseEvent('mouseup', {{ bubbles: true }}));
                                    
                                    return 'ok';
                                }}
                            }}
                        }}
                        maxRetries--;
                    }}
                    return 'item_not_found';
                }} catch(e) {{
                    return 'error: ' + e.message;
                }}
            }})();";

            string res = await ExecuteScriptSafeAsync(webView, selectJs);
            if (res.Contains("no_input")) throw new Exception("Không tìm thấy ô Dropdown.");
            if (res.Contains("item_not_found")) throw new Exception($"Không tìm thấy mục '{targetOption}'");
            await Task.Delay(200, token);
        }

        public static async Task ClickSearchAsync(WebView2 webView, string expectedWaybill, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            string clickJs = @"(function() {
                var container = document.querySelector('div[id^=""el-collapse-content-""]');
                if (container) {
                    var btn = container.querySelector('button.el-button--primary');
                    if (btn) {
                        // Kiểm tra xem nút có bị VueJS khóa không
                        if (btn.disabled || btn.classList.contains('is-disabled')) return 'disabled';
                        btn.click(); 
                        return 'clicked';
                    }
                }
                return 'not_found';
            })();";

            var responseTask = WaitForApiResponseAsync(webView,
                json => json.Contains($"\"waybillNo\":\"{expectedWaybill}\"") || json.Contains("\"succ\":false"),
                timeoutMs: 3000); // Cho web 3s để phản hồi

            string clickRes = await ExecuteScriptSafeAsync(webView, clickJs);
            clickRes = clickRes.Trim('"');

            // NẾU BỊ KHÓA, QUĂNG LỖI NGAY ĐỂ ĐƯỢC CHẠY LẠI
            if (clickRes == "disabled") throw new Exception("Nút Tìm kiếm đang bị mờ. Vue chưa nhận dữ liệu.");
            if (clickRes == "not_found") throw new Exception("Không tìm thấy nút Tìm kiếm.");

            try
            {
                string jsonResult = await responseTask;
                CheckAndThrowIfError(jsonResult);
            }
            catch (TimeoutException)
            {
                string errorJs = "document.querySelector('.el-message--error') ? document.querySelector('.el-message--error').innerText : ''";
                string error = await ExecuteScriptSafeAsync(webView, errorJs);
                error = error.Trim('"');
                if (!string.IsNullOrEmpty(error)) throw new Exception($"Lỗi UI: {error}");

                throw new Exception("Timeout: Đã bấm nhưng Web không trả về API.");
            }
        }

        public static async Task ClickSaveAndVerifyAsync(WebView2 webView, CancellationToken token)
        {
            var responseTask = WaitForApiResponseAsync(webView,
                json => (json.Contains("Thao tác thành công") && json.Contains("\"code\":1")) || json.Contains("\"succ\":false"),
                timeoutMs: 1500);

            string saveJs = @"(function() {
                var btn = document.querySelector('button[title=""Lưu và thêm mới""]');
                if(btn) { btn.click(); return 'clicked'; }
                return 'not_found';
            })();";

            string clickRes = await ExecuteScriptSafeAsync(webView, saveJs);
            if (clickRes.Contains("not_found")) throw new Exception("Không tìm thấy nút Lưu.");

            try
            {
                string jsonResult = await responseTask;
                CheckAndThrowIfError(jsonResult);
            }
            catch (TimeoutException)
            {
                // BẮT BUỘC NÉM LỖI ĐỂ KÍCH HOẠT SWITCH DKCH2
                string errorJs = "document.querySelector('.el-message--error') ? document.querySelector('.el-message--error').innerText : ''";
                string error = await ExecuteScriptSafeAsync(webView, errorJs);
                error = error.Trim('"');
                if (!string.IsNullOrEmpty(error)) throw new Exception($"Lỗi UI: {error}");

                throw new Exception("Timeout: Web không phản hồi khi Lưu đơn."); // Ném thẳng ra ngoài
            }
        }
    }

        public class DkchManager
    {

        public event Action<string> OnLog;
        public event Action<string> OnStatusUpdate;
        public event Action<string> OnCurrentWaybillChanged;
        public event Action<int> OnSaveCountChanged;
        public event Action<string> OnTrackingHistoryChanged;
        public event Action<string> OnWaybillCompleted;

        private PeriodicTimer _mainLoadTimer;
        private CancellationTokenSource _dkchCts;
        private bool _isRunning = false;
        private bool _isProcessing = false;
        private string _currentMode;
        private int _lastProcessedIndex = 0;
        private int _saveCount = 0;
        private List<string> _priorityQueue = new List<string>();
        private object _queueLock = new object();
        private static readonly Regex WaybillRegex = new Regex("^[A-Za-z0-9]{1,20}$", RegexOptions.Compiled);
        private DateTime _lastSheetFetchTime = DateTime.MinValue;
        private List<string> _cachedSheetData = new List<string>();
        // Dependencies
        private WebView2 _webView;
        private WaybillTrackingService _trackingService;
        private Func<(bool useSheet, string sheetName, int rowCount)> _settingsGetter;

        public bool IsRunning => _isRunning;

        public void SetWebView(WebView2 webView) => _webView = webView;
        public void SetTrackingService(WaybillTrackingService service) => _trackingService = service;
        public void SetSettingsGetter(Func<(bool useSheet, string sheetName, int rowCount)> getter) => _settingsGetter = getter;


        public void AddPriorityWaybill(string waybill)
        {
            if (!IsValidWaybill(waybill)) return;
            lock (_queueLock)
                _priorityQueue.Add(waybill.Trim());
        }

        private static bool IsValidWaybill(string waybill)
        {
            if (string.IsNullOrWhiteSpace(waybill)) return false;
            return WaybillRegex.IsMatch(waybill.Trim());
        }

        public void StartDaemon()
        {
            _mainLoadTimer = new PeriodicTimer(TimeSpan.FromMilliseconds(500));
            _ = Task.Run(async () =>
            {
                while (await _mainLoadTimer.WaitForNextTickAsync())
                {
                    try
                    {
                        if (!_isRunning || _isProcessing) continue;

                        bool hasPriorityJob = false;
                        lock (_queueLock) hasPriorityJob = _priorityQueue.Count > 0;

                        List<string> sheetData = new List<string>();
                        bool useSheet = false;
                        string sheetName = "";
                        int rowCount = 0;

                        if (!hasPriorityJob && _settingsGetter != null)
                        {
                            var sets = _settingsGetter();
                            useSheet = sets.useSheet;
                            sheetName = sets.sheetName;
                            rowCount = sets.rowCount;

                            if (useSheet)
                            {
                                //Giới hạn 15 giây mới Request 1 lần
                                if ((DateTime.Now - _lastSheetFetchTime).TotalSeconds >= 15)
                                {
                                    _cachedSheetData = GoogleSheetService.ReadColumn(sheetName, rowCount);
                                    _lastSheetFetchTime = DateTime.Now;
                                }

                                sheetData = _cachedSheetData;
                            }
                            else
                            {

                            }
                        }

                        bool hasNewJob = hasPriorityJob || (useSheet && _lastProcessedIndex < sheetData.Count);

                        if (hasNewJob && _dkchCts != null && !_dkchCts.IsCancellationRequested)
                        {
                            _isProcessing = true;
                            try
                            {
                                await ProcessAutomationQueueAsync(sheetData);
                            }
                            finally { _isProcessing = false; }
                        }
                    }
                    catch { _isProcessing = false; }
                }
            });
        }

        public async Task StartAsync(string mode)
        {
            string targetUrl = AppConfig.Current.BuildJmsUrl("app/operatingPlatformIndex/returnAndForwardMaintainAddSite");
            string current = _webView.Source?.ToString() ?? "";
            if (!current.Contains("returnAndForwardMaintainAddSite"))
            {
                _webView.Source = new Uri(targetUrl);
                await Task.Delay(1000);
            }

            Stop();

            _currentMode = mode;
            _dkchCts = new CancellationTokenSource();
            _isRunning = true;
            _isProcessing = false;

            lock (_queueLock) _priorityQueue.Clear();

            _lastProcessedIndex = 0;
            _saveCount = 0;
            OnSaveCountChanged?.Invoke(0);
            _lastSheetFetchTime = DateTime.MinValue;
            OnLog?.Invoke($"=== START {_currentMode} ===");
        }

        public void Stop()
        {
            _isRunning = false;
            _dkchCts?.Cancel();
            _dkchCts = null;
            _isProcessing = false;
        }

        private async Task ProcessAutomationQueueAsync(List<string> sheetData)
        {
            var tokenSource = _dkchCts;
            if (tokenSource == null || tokenSource.IsCancellationRequested) return;

            // Xử lý hàng đợi ưu tiên trước
            while (true)
            {
                if (tokenSource.IsCancellationRequested) return;

                string waybill = null;
                lock (_queueLock)
                {
                    if (_priorityQueue.Count > 0) waybill = _priorityQueue[0];
                }

                if (waybill == null) break;
                if (!IsValidWaybill(waybill))
                {
                    lock (_queueLock)
                    {
                        if (_priorityQueue.Count > 0) _priorityQueue.RemoveAt(0);
                    }
                    continue;
                }

                OnCurrentWaybillChanged?.Invoke(waybill);
                await ExecuteOneWaybill(waybill, tokenSource.Token);
                OnWaybillCompleted?.Invoke(waybill);

                lock (_queueLock)
                {
                    if (_priorityQueue.Count > 0) _priorityQueue.RemoveAt(0);
                }
            }

            // Xử lý dữ liệu từ sheet
            if (_lastProcessedIndex < sheetData.Count)
            {
                while (_lastProcessedIndex < sheetData.Count)
                {
                    if (tokenSource.IsCancellationRequested) return;

                    bool hasNewInput = false;
                    lock (_queueLock) hasNewInput = _priorityQueue.Count > 0;
                    if (hasNewInput) return;

                    string waybill = sheetData[_lastProcessedIndex];
                    if (!IsValidWaybill(waybill))
                    {
                        _lastProcessedIndex++;
                        continue;
                    }
                    OnLog?.Invoke($"▶ [{_currentMode}] Row {_lastProcessedIndex + 1}: {waybill}");
                    OnCurrentWaybillChanged?.Invoke(waybill);

                    await ExecuteOneWaybill(waybill, tokenSource.Token);
                    OnWaybillCompleted?.Invoke(waybill);
                    _lastProcessedIndex++;
                }
            }
        }


        private async Task<bool> ExecuteOneWaybill(string waybill, CancellationToken token)
        {
            int maxRetries = 2;
            int attempt = 0;

            while (attempt < maxRetries)
            {
                attempt++;
                try
                {
                    if (attempt > 1)
                        OnStatusUpdate?.Invoke($"1. Đang điền đơn (Thử lại lần {attempt})...");
                    

                    string currentTarget = (_currentMode == "DKCH2") ? "Chuyển hoàn lần 2" : "Chuyển hoàn";

                    for (int switchAttempt = 0; switchAttempt < 2; switchAttempt++)
                    {
                        try
                        {
                            // Chống rỗng ô: Điền mã trước mỗi lần thực thi các bước
                            await WebViewAutomation.FillWaybillAsync(_webView, waybill, token);
                            await RunSteps(waybill, currentTarget, token);

                            // Thành công → lưu lại và thoát
                            _saveCount++;
                            OnSaveCountChanged?.Invoke(_saveCount);
                            return true;
                        }
                        catch (NoDataWaybillException)
                        {
                            OnLog?.Invoke($"{waybill}: Skip (Không có dữ liệu)");
                            return false;
                        }
                        catch (NeedSwitchToDkch2Exception)
                        {
                            if (switchAttempt == 1) throw;
                            currentTarget = "Chuyển hoàn lần 2";
                            OnLog?.Invoke($"{waybill}: → Chuyển sang DKCH2");

                        }
                        catch (NeedSwitchToDkch1Exception)
                        {
                            if (switchAttempt == 1) throw;
                            currentTarget = "Chuyển hoàn";
                            OnLog?.Invoke($"{waybill}: → Chuyển sang DKCH1");
                        }
                        catch (Exception ex)
                        {
                            // Mọi lỗi khác khi chưa ở DKCH2 → ép thử DKCH2 ít nhất 1 lần
                            if (currentTarget != "Chuyển hoàn lần 2" && switchAttempt == 0)
                            {
                                OnLog?.Invoke($"{waybill}: Lỗi ({ex.Message}) → Ép chuyển DKCH2");
                                currentTarget = "Chuyển hoàn lần 2";
                            }
                            else
                            {
                                // Nếu đang ở DKCH2 rồi, hoặc đã switch rồi mà vẫn lỗi -> Ném ra ngoài để retry
                                throw;
                            }
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                    return false;
                }
                catch (Exception ex)
                {
                    OnLog?.Invoke($"{waybill}: Lỗi vòng ngoài lần {attempt}: {ex.Message}");
                    await Task.Delay(200, token); // Đợi mạng/trình duyệt ổn định lại
                }
            }

            OnLog?.Invoke($"{waybill}: Thất bại sau {maxRetries} lần thử.");
            return false;
        }

        private async Task RunSteps(string waybill, string dropdownOption, CancellationToken token)
        {
            OnStatusUpdate?.Invoke($"1. Check Menu: {dropdownOption}...");
            await WebViewAutomation.CheckAndSelectDropdownAsync(_webView, dropdownOption, token);

            await Task.Delay(100, token);

            OnStatusUpdate?.Invoke("2. Đang điền/kiểm tra mã...");
            await WebViewAutomation.FillWaybillAsync(_webView, waybill, token);

            await Task.Delay(100, token);

            OnStatusUpdate?.Invoke("3. Đang tìm kiếm...");
            await WebViewAutomation.ClickSearchAsync(_webView, waybill, token);

            if (_trackingService != null)
            {
                string history = await _trackingService.GetDKCHHistoryAsync(waybill);
                OnTrackingHistoryChanged?.Invoke(history ?? "Không có dữ liệu");
            }
            else
            {
                OnTrackingHistoryChanged?.Invoke("Chưa khởi tạo Tracking Service.");
            }

            OnStatusUpdate?.Invoke("4. Đang Lưu...");
            await WebViewAutomation.ClickSaveAndVerifyAsync(_webView, token);
        }



    }
}
