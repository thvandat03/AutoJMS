using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
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
                try
                {
                    using (JsonDocument doc = JsonDocument.Parse(json))
                    {
                        var root = doc.RootElement;
                        string msg = root.TryGetProperty("msg", out var m) ? m.GetString() : "";
                        string code = root.TryGetProperty("code", out var c) ? c.ToString() : "";

                        if (msg.Contains("Chưa có") || msg.Contains("hoàn lần 2") || code == "999006328")
                            throw new NeedSwitchToDkch1Exception();
                        if (code == "137043004" || code == "999006082" || code.Contains(":"))
                            throw new NeedSwitchToDkch2Exception();
                        if (msg.Contains("không có dữ liệu") || msg.Contains("Vận đơn không tồn tại"))
                            throw new NoDataWaybillException();
                        if (!string.IsNullOrEmpty(msg))
                            throw new Exception($"Lỗi: {msg}");
                    }
                }
                catch (Exception ex) when (ex is NoDataWaybillException || ex is NeedSwitchToDkch1Exception || ex is NeedSwitchToDkch2Exception)
                {
                    throw;
                }
                catch { }
            }
        }

        public static async Task FillWaybillAsync(WebView2 webView, string waybill, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

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
                    input.value = '';
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.focus();
                    input.value = '{waybill}';
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    input.blur();
                    return input.value;
                }}
                return 'notfound';
            }})();";

            int retries = 2;
            while (retries > 0)
            {
                string result = await ExecuteScriptSafeAsync(webView, js);
                string currentVal = result.Trim('"');
                if (currentVal == "notfound")
                    throw new Exception("Không tìm thấy ô nhập liệu.");
                if (currentVal == waybill)
                    return;
                await Task.Delay(200, token);
                retries--;
            }
            throw new Exception($"Không thể điền mã '{waybill}' sau nhiều lần thử.");
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

            string selectJs = $@"
            (async function() {{
                let ddInput = document.querySelector('.el-select .el-input__inner[readonly]');
                if (!ddInput) return 'no_input';
                ddInput.click();
                let maxRetries = 20;
                while(maxRetries > 0) {{
                    await new Promise(r => setTimeout(r, 100));
                    let items = document.querySelectorAll('li.el-select-dropdown__item span');
                    for (let span of items) {{
                        if (span.innerText.trim() === '{targetOption}') {{
                            span.parentElement.click();
                            return 'ok';
                        }}
                    }}
                    maxRetries--;
                }}
                return 'item_not_found';
            }})();";

            string res = await ExecuteScriptSafeAsync(webView, selectJs);
            if (res.Contains("no_input")) throw new Exception("Không tìm thấy ô Dropdown.");
            if (res.Contains("item_not_found")) throw new Exception($"Không tìm thấy mục '{targetOption}'");
            await Task.Delay(200, token);
        }

        public static async Task ClickSearchAsync(WebView2 webView, string expectedWaybill, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            var responseTask = WaitForApiResponseAsync(webView,
                json => json.Contains($"\"waybillNo\":\"{expectedWaybill}\"") || json.Contains("\"succ\":false"),
                timeoutMs: 2000);

            string clickJs = @"(function() {
                var container = document.querySelector('div[id^=""el-collapse-content-""]');
                if (container) {
                    var btn = container.querySelector('button.el-button--primary');
                    if(btn) { btn.click(); return 'clicked'; }
                }
                return 'not_found';
            })();";
            await ExecuteScriptSafeAsync(webView, clickJs);

            try
            {
                string jsonResult = await responseTask;
                CheckAndThrowIfError(jsonResult);
            }
            catch (TimeoutException) { }
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
                string errorJs = "document.querySelector('.el-message--error') ? document.querySelector('.el-message--error').innerText : ''";
                string error = await ExecuteScriptSafeAsync(webView, errorJs);
                error = error.Trim('"');
                if (!string.IsNullOrEmpty(error)) throw new Exception($"Lỗi: {error}");
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

        private PeriodicTimer _mainLoadTimer;
        private CancellationTokenSource _dkchCts;
        private bool _isRunning = false;
        private bool _isProcessing = false;
        private string _currentMode;
        private int _lastProcessedIndex = 0;
        private int _saveCount = 0;
        private List<string> _priorityQueue = new List<string>();
        private object _queueLock = new object();

        // Dependencies
        private WebView2 _webView;
        private WaybillTrackingService _trackingService;
        private Func<(bool useSheet, string sheetName, int rowCount)> _settingsGetter;
        private TextBox _previewBox;
        private Label _countLabel;

        public bool IsRunning => _isRunning;

        public void SetWebView(WebView2 webView) => _webView = webView;
        public void SetTrackingService(WaybillTrackingService service) => _trackingService = service;
        public void SetSettingsGetter(Func<(bool, string, int)> getter) => _settingsGetter = getter;
        public void SetUILogger(TextBox preview, Label count)
        {
            _previewBox = preview;
            _countLabel = count;
        }

        public void AddPriorityWaybill(string waybill)
        {
            lock (_queueLock)
                _priorityQueue.Add(waybill);
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
                                sheetData = GoogleSheetService.ReadColumn(sheetName, rowCount);
                                UpdatePreview(string.Join(Environment.NewLine, sheetData), sheetData.Count);
                            }
                            else
                            {
                                UpdatePreview("=Hiện tại không lấy data từ Sheet=\r\n=====================\r\n||=======Đạt đzvl=======||\r\n======================", 0);
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
            string targetUrl = "https://jms.jtexpress.vn/app/operatingPlatformIndex/returnAndForwardMaintainAddSite";
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

            OnLog?.Invoke($"=== START {_currentMode} ===");
        }

        public void Stop()
        {
            _isRunning = false;
            _dkchCts?.Cancel();
            _dkchCts = null;
            _isProcessing = false;
            OnStatusUpdate?.Invoke("Đã dừng.");
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

                OnCurrentWaybillChanged?.Invoke(waybill);
                await ExecuteOneWaybill(waybill, tokenSource.Token);

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
                    OnLog?.Invoke($"▶ [{_currentMode}] Row {_lastProcessedIndex + 1}: {waybill}");
                    OnCurrentWaybillChanged?.Invoke(waybill);

                    await ExecuteOneWaybill(waybill, tokenSource.Token);
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
                    if (attempt > 1) OnLog?.Invoke($"[{waybill}] Thử lại lần {attempt}...");
                    OnStatusUpdate?.Invoke("1. Đang điền đơn...");
                    await WebViewAutomation.FillWaybillAsync(_webView, waybill, token);

                    bool retryMode = false;
                    string currentTarget = (_currentMode == "DKCH2") ? "Chuyển hoàn lần 2" : "Chuyển hoàn";

                    try
                    {
                        await RunSteps(waybill, currentTarget, token);
                    }
                    catch (NoDataWaybillException)
                    {
                        OnLog?.Invoke($"{waybill}: Skip");
                        return false;
                    }
                    catch (NeedSwitchToDkch2Exception)
                    {
                        retryMode = true;
                        currentTarget = "Chuyển hoàn lần 2";
                    }

                    if (retryMode)
                        await RunSteps(waybill, currentTarget, token);

                    _saveCount++;
                    OnSaveCountChanged?.Invoke(_saveCount);
                    OnLog?.Invoke($"✅: {waybill}");
                    return true;
                }
                catch (OperationCanceledException) { return false; }
                catch (Exception ex)
                {
                    OnLog?.Invoke($"❌ ({attempt}/{maxRetries}): {ex.Message}");
                    await Task.Delay(500, token);
                }
            }
            OnLog?.Invoke($"{waybill} Thất bại");
            return false;
        }

        private async Task RunSteps(string waybill, string dropdownOption, CancellationToken token)
        {
            OnStatusUpdate?.Invoke($"2. Check: {dropdownOption}...");
            await WebViewAutomation.CheckAndSelectDropdownAsync(_webView, dropdownOption, token);

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

        private void UpdatePreview(string text, int count)
        {
            if (_previewBox != null && !_previewBox.IsDisposed)
            {
                if (_previewBox.InvokeRequired)
                    _previewBox.Invoke(() => { _previewBox.Text = text; });
                else
                    _previewBox.Text = text;
            }
            if (_countLabel != null && ! _countLabel.IsDisposed)
            {
                if (_countLabel.InvokeRequired)
                    _countLabel.Invoke(() => { _countLabel.Text = count.ToString(); });
                else
                    _countLabel.Text = count.ToString();
            }
        }

    }
}