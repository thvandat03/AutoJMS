using AutoJMS.Data;
using Microsoft.Web.WebView2.Core;
using Sunny.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    public partial class Main : UIForm
    {
        private const string LICENSE_SPREADSHEET_ID = "1nx2VoXnAU3h8GRPxXkZ4c9Ev8jwHg3Iyor5o6wvsLNY";
        private const string LICENSE_SHEET_NAME = "LICENSE";
        string appsScriptUrl = "https://script.google.com/macros/s/AKfycbzVoJuFl-jf5wWXKEkRCLUXP8_cQ8yYWdDojFo-V19FH1G6Lm0AxhEyjZ9XI85pXIKGTQ/exec";

        private AppSettings _settings;
        public static string CapturedAuthToken = "";
        public static WaybillTrackingService _trackingService;
        private DkchManager _dkchManager;
        private PrintService _printService;
        private ZaloChatService _zaloChatService;
        private System.Windows.Forms.Timer _queueTimer;
        private List<Reminder> _trackedReminders = new List<Reminder>();   // Nguồn dữ liệu chính cho tabChat
        private bool _isUpdatingUI = false;
        private BindingSource _chatBindingSource;

        private bool _isZaloLoaded = false;
        private bool _isDkchNeedReload = true;

        public Main()
        {
            InitializeComponent();
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.MinimumSize = new Size(1366, 768);
            _settings = SettingsManager.Load();

            // UI styling
            tabDKCH_inputNewBill.Font = new System.Drawing.Font("Segoe UI", 14f, System.Drawing.FontStyle.Bold);
            int singleLineHeight = TextRenderer.MeasureText("A", tabDKCH_inputNewBill.Font).Height;
            int textHeightFor6Lines = singleLineHeight * 6;
            tabDKCH_inputNewBill.Height = textHeightFor6Lines + tabDKCH_inputNewBill.Padding.Top + tabDKCH_inputNewBill.Padding.Bottom + 4;

            string sharedFolder = Path.Combine(Application.StartupPath, "SharedBrowserData");
            var sharedProps = new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties()
            {
                UserDataFolder = sharedFolder
            };

            // TabChat: BindingSource + DataTable (giữ nguyên)
            _chatBindingSource = LocalTrackingManager.Instance.ChatBindingSource;
            tabChat_dataGrid.DataSource = _chatBindingSource;
            tabChat_dataGrid.AutoGenerateColumns = false;
            tabChat_dataGrid.AllowUserToAddRows = false;
            tabChat_dataGrid.ReadOnly = true;
            tabChat_dataGrid.RowHeadersVisible = false;


            // TabDash: Virtual Mode
            SetupDashGridVirtualMode();

            // Sự kiện cập nhật từ LocalTrackingManager (chỉ dành cho tabDash)
            LocalTrackingManager.Instance.OnDataUpdated += () =>
            {
                if (this.InvokeRequired)
                    this.Invoke((Action)(() =>
                    {
                        RefreshDashboardUI(); // cho tabDash
                        if (tabControl.SelectedTab == tabChat)
                        {
                            PopulateStatusSelect();
                            ApplyFilterAndSort();
                        }
                    }));
                else
                {
                    RefreshDashboardUI();
                    if (tabControl.SelectedTab == tabChat)
                    {
                        PopulateStatusSelect();
                        ApplyFilterAndSort();
                    }
                }
            };

            // Sự kiện lọc cho tabChat
            tabChat_statusSelect.SelectedIndexChanged += tabChat_statusSelect_SelectedIndexChanged;
            tabChat_timeSelect.SelectedIndexChanged += (s, e) => UpdateTrackingInterval();

            // Double buffering cho tabTracking
            var doubleBufferPropertyInfo = tabTracking_dataView.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            doubleBufferPropertyInfo?.SetValue(tabTracking_dataView, true, null);

            // WebView
            tabHome_webView.CreationProperties = sharedProps;
            tabDKCH_webView.CreationProperties = sharedProps;
            tabDKCH_sheetName.SelectedItem = _settings.DefaultSheet;
            tabDKCH_useSheet.Active = _settings.UseSheetByDefault;
            tabDKCH_numRow.Value = _settings.DefaultRowCount;

            // Sự kiện bàn phím
            tabTracking_inputWaybill.KeyDown += tabTracking_inputWaybill_KeyDown;
            tabDKCH_nowTracking.KeyDown += tabDKCH_txtNewBill_KeyDown;
            tabPrint_btnSelectAll.CheckedChanged += tabPrint_btnSelectAll_CheckedChanged;
            tabHome_webView.NavigationCompleted += tabHome_WebView_NavigationCompleted;

            // Timer xử lý lệnh RUN
            _queueTimer = new System.Windows.Forms.Timer();
            _queueTimer.Interval = 30000;
            _queueTimer.Tick += async (s, e) => await ProcessDirectTrackingAsync();
            _queueTimer.Start();

            tabChat_timeSelect.DropDownStyle = UIDropDownStyle.DropDown;

            // DKCH buttons
            tabDKCH_btnDKCH1.Visible = true;
            tabDKCH_btnDKCH2.Visible = true;
            tabDKCH_btnStop.Visible = false;
            tabDKCH_btnDKCH1.Enabled = false;
            tabDKCH_btnDKCH2.Enabled = false;

            CheckForIllegalCrossThreadCalls = false;
            this.KeyPreview = true;

            _dkchManager = new DkchManager();
            _dkchManager.OnSaveCountChanged += (count) => tabDKCH_countSave.Text = count.ToString();
            _dkchManager.OnTrackingHistoryChanged += (history) => tabDKCH_inputNewBill.Text = history;
        }

        // ========== CẤU HÌNH VIRTUAL MODE CHO TAB DASH ==========
        private void SetupDashGridVirtualMode()
        {
            var dgv = tabDash_dataGrid;
            dgv.VirtualMode = true;
            dgv.AutoGenerateColumns = false;
            dgv.AllowUserToAddRows = false;
            dgv.ReadOnly = true;
            dgv.RowHeadersVisible = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dgv.Columns.Clear();
            dgv.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Mã vận đơn", HeaderText = "Mã vận đơn" },
                new DataGridViewTextBoxColumn { Name = "Trạng thái hiện tại", HeaderText = "Trạng thái hiện tại" },
                new DataGridViewTextBoxColumn { Name = "Loại quét kiện cuối", HeaderText = "Loại quét kiện cuối" },
                new DataGridViewTextBoxColumn { Name = "Tên nhân viên", HeaderText = "Tên nhân viên" },
                new DataGridViewTextBoxColumn { Name = "Thời gian thao tác", HeaderText = "Thời gian thao tác" },
                new DataGridViewTextBoxColumn { Name = "Thời gian yêu cầu phát lại", HeaderText = "Thời gian yêu cầu phát lại" },
                new DataGridViewTextBoxColumn { Name = "Số lần nhắc", HeaderText = "Số lần nhắc" },
                new DataGridViewTextBoxColumn { Name = "Thời gian nhắc gần nhất", HeaderText = "Thời gian nhắc gần nhất" },
                new DataGridViewTextBoxColumn { Name = "Số lần đăng ký chuyển hoàn", HeaderText = "Số lần đăng ký chuyển hoàn" },
                new DataGridViewTextBoxColumn { Name = "Số lần giao lại hàng", HeaderText = "Số lần giao lại hàng" },
                new DataGridViewTextBoxColumn { Name = "Nguồn đơn đặt", HeaderText = "Nguồn đơn đặt" }
            });

            dgv.CellValueNeeded += DashGrid_CellValueNeeded;
            dgv.RowCount = 0;
        }

        private void DashGrid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            var list = LocalTrackingManager.Instance.GetVirtualList();
            if (e.RowIndex < 0 || e.RowIndex >= list.Count) return;
            var row = list[e.RowIndex];
            var colName = tabDash_dataGrid.Columns[e.ColumnIndex].Name;

            switch (colName)
            {
                case "Mã vận đơn": e.Value = row.WaybillNo; break;
                case "Trạng thái hiện tại": e.Value = row.TrangThaiHienTai; break;
                case "Loại quét kiện cuối": e.Value = row.ThaoTacCuoi; break;
                case "Tên nhân viên": e.Value = row.NguoiThaoTac; break;
                case "Thời gian thao tác": e.Value = row.ThoiGianThaoTac; break;
                case "Thời gian yêu cầu phát lại": e.Value = row.ThoiGianYeuCauPhatLai; break;
                case "Số lần nhắc": e.Value = row.SoLanNhac; break;
                case "Thời gian nhắc gần nhất": e.Value = row.ThoiGianNhacGanNhat; break;
                case "Số lần đăng ký chuyển hoàn": e.Value = row.SoLanDangKyChuyenHoan; break;
                case "Số lần giao lại hàng": e.Value = row.SoLanGiaoLaiHang; break;
                case "Nguồn đơn đặt": e.Value = row.NguonDonDat; break;
            }
        }

        private void RefreshDashboardUI()
        {
            if (tabDash_dataGrid == null) return;
            var list = LocalTrackingManager.Instance.GetVirtualList();
            tabDash_dataGrid.RowCount = list.Count;
            tabDash_dataGrid.Invalidate();
            tabDash_dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            if (tabDash_lblDebug != null)
                tabDash_lblDebug.Text = $"✅ Đã load {list.Count} đơn | Cập nhật: {LocalTrackingManager.Instance.LastTrackedTime:HH:mm:ss}";
        }

        // ========== TAB CHAT: LỌC VÀ HIỂN THỊ ==========
        private void tabChat_statusSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFilterAndSort();
        }

        private async Task RefreshTrackingDataAsync()
        {
            try
            {
                var waybills = _zaloChatService?.GetWaybillsFromPhatLai();
                if (waybills == null || waybills.Count == 0) return;

                string waybillText = string.Join("\n", waybills);
                await Main._trackingService.SearchTrackingAsync(waybillText, updateMainGrid: false);

                var trackedRows = _trackingService.GetAllRows();
                _trackedReminders = trackedRows.Select(r => new Reminder
                {
                    maDon = r.WaybillNo,
                    nhanVien = r.NhanVienKienVanDe ?? r.NguoiThaoTac ?? "",
                    trangThai = r.ThaoTacCuoi ?? "",
                    soLanNhac = 0,
                    thoiGianNhac = "",
                    row = 0
                }).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[RefreshTrackingDataAsync] Lỗi: {ex.Message}");
            }
        }

        private async Task RefreshZaloDataAsync(bool forceReload = false)
        {
            // Nếu chưa có dữ liệu hoặc cần tải lại, gọi RefreshTrackingDataAsync
            if (_trackedReminders == null || _trackedReminders.Count == 0 || forceReload)
            {
                await RefreshTrackingDataAsync();
            }

            // Cập nhật ComboBox trạng thái dựa trên _trackedReminders
            PopulateStatusSelect();

            // Áp dụng lọc và hiển thị
            ApplyFilterAndSort();
        }

        private void ApplyFilterAndSort()
        {
            if (_isUpdatingUI) return;
            _isUpdatingUI = true;

            try
            {
                string selected = tabChat_statusSelect.SelectedItem?.ToString() ?? "Tất cả";

                if (selected == "Tất cả")
                    _chatBindingSource.RemoveFilter();
                else
                    _chatBindingSource.Filter = $"[Trạng thái hiện tại] = '{selected.Replace("'", "''")}'";

                // Sắp xếp theo tên nhân viên
                _chatBindingSource.Sort = "Tên nhân viên ASC";
            }
            finally
            {
                _isUpdatingUI = false;
            }
        }

        private void UpdateZaloGridUI(List<Reminder> data)
        {
            if (tabChat_dataGrid.InvokeRequired)
            {
                tabChat_dataGrid.Invoke(() => UpdateZaloGridUI(data));
                return;
            }

            tabChat_dataGrid.SuspendLayout();
            try
            {
                tabChat_dataGrid.AutoGenerateColumns = false;
                tabChat_dataGrid.AllowUserToAddRows = false;
                tabChat_dataGrid.RowHeadersVisible = false;

                if (tabChat_dataGrid.Columns.Count >= 4)
                {
                    tabChat_dataGrid.Columns["maDon"].DataPropertyName = "maDon";
                    tabChat_dataGrid.Columns["nhanVien"].DataPropertyName = "nhanVien";
                    tabChat_dataGrid.Columns["trangThai"].DataPropertyName = "trangThai";
                    tabChat_dataGrid.Columns["soLanNhac"].DataPropertyName = "soLanNhac";
                }

                tabChat_dataGrid.DataSource = null;
                tabChat_dataGrid.DataSource = data;
                tabChat_dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

                // Buộc làm mới ngay lập tức
                tabChat_dataGrid.Refresh();
            }
            finally
            {
                tabChat_dataGrid.ResumeLayout();
            }
        }

        private void PopulateStatusSelect()
        {
            if (tabChat_statusSelect == null) return;
            tabChat_statusSelect.Items.Clear();
            tabChat_statusSelect.Items.Add("Tất cả");

            DataTable dt = (DataTable)_chatBindingSource.DataSource;
            if (dt == null || dt.Rows.Count == 0) return;

            var uniqueStatuses = dt.AsEnumerable()
                .Select(r => r.Field<string>("Trạng thái hiện tại")?.Trim())
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s)
                .ToList();

            foreach (var s in uniqueStatuses)
                tabChat_statusSelect.Items.Add(s);

            tabChat_statusSelect.SelectedIndex = 0;
        }

        private void UpdateTrackingInterval()
        {
            string text = tabChat_timeSelect.Text?.Replace(" phút", "").Trim() ?? "5";
            if (int.TryParse(text, out int minutes) && minutes > 0)
            {
                LocalTrackingManager.Instance.StopAutoTracking();
                LocalTrackingManager.Instance.StartAutoTracking(minutes);
            }
        }

        // ========== ONLOAD ==========
        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (this.IsDisposed) return;

            _ = tabHome_webView.Handle;
            _ = tabDKCH_webView.Handle;

            try
            {
                await tabHome_webView.EnsureCoreWebView2Async(null);
                await tabDKCH_webView.EnsureCoreWebView2Async(null);
                GoogleSheetService.InitService();

                tabHome_webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
                tabHome_webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
                tabDKCH_webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
                tabDKCH_webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;

                tabHome_webView.CoreWebView2.Navigate("https://jms.jtexpress.vn");
                tabDKCH_webView.CoreWebView2.Navigate("https://jms.jtexpress.vn");

                tabDKCH_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Fetch);
                tabDKCH_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.XmlHttpRequest);
                tabDKCH_webView.CoreWebView2.WebResourceRequested += CoreWebView2_WebResourceRequested;
                tabDKCH_webView.CoreWebView2.NavigationCompleted += (s, args) => { if (args.IsSuccess) ApplyZoomFactor(); };
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khởi tạo trình duyệt: " + ex.Message);
            }

            // Khởi tạo tracking service
            _trackingService = new WaybillTrackingService(tabTracking_dataView, tabTracking_process);
            LocalTrackingManager.Instance.TrackingService = _trackingService;

            // Chạy tracking lần đầu
            await LocalTrackingManager.Instance.PerformIncrementalTrackingAsync();
            LocalTrackingManager.Instance.StartAutoTracking(5);

            WaybillTrackingService.EnableDoubleBuffering(tabTracking_dataView);
            tabTracking_dataView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;

            // DKCH
            _dkchManager.SetWebView(tabDKCH_webView);
            _dkchManager.SetTrackingService(_trackingService);
            _dkchManager.SetSettingsGetter(() => (
                useSheet: tabDKCH_useSheet.Active,
                sheetName: tabDKCH_sheetName.Text,
                rowCount: (int)tabDKCH_numRow.Value
            ));
            _dkchManager.StartDaemon();

            tabDKCH_btnDKCH1.Enabled = true;
            tabDKCH_btnDKCH2.Enabled = true;

            await WebViewHost.InitAsync(tabDKCH_webView);
            ApplyZoomFactor();

            tabDKCH_webView.ZoomFactorChanged += (s, args) =>
            {
                _settings.ZoomFactor = tabDKCH_webView.ZoomFactor;
                SettingsManager.Save(_settings);
            };

            await WebViewHost.NavigateAsync(_settings.DefaultUrl);
            await Task.Delay(2000);
            await RefreshAuthTokenAsync();
        }

        // ========== CÁC SỰ KIỆN KHÁC ==========
        private void ApplyZoomFactor()
        {
            if (tabDKCH_webView?.CoreWebView2 != null)
                tabDKCH_webView.ZoomFactor = _settings.ZoomFactor;
        }

        private void tabTracking_inputWaybill_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                e.SuppressKeyPress = true;
                if (Clipboard.ContainsText())
                {
                    string plainText = Clipboard.GetText(TextDataFormat.UnicodeText);
                    string cleaned = string.Join(Environment.NewLine,
                        plainText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                 .Select(line => line.Trim())
                                 .Where(line => !string.IsNullOrWhiteSpace(line)));
                    tabTracking_inputWaybill.SelectedText = cleaned;
                }
            }
        }

        private void tabDKCH_txtNewBill_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                var allLines = tabDKCH_nowTracking.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (allLines.Count == 0 && tabDKCH_nowTracking.Text.Trim().Length > 0)
                    allLines.Add(tabDKCH_nowTracking.Text.Trim());
                else if (allLines.Count == 0) return;

                string newWaybill = allLines.Last().Trim();
                if (!string.IsNullOrEmpty(newWaybill))
                {
                    _dkchManager.AddPriorityWaybill(newWaybill);
                    if (allLines.Count >= 5)
                    {
                        var keepLines = allLines.Skip(allLines.Count - 5).ToList();
                        tabDKCH_nowTracking.Text = string.Join(Environment.NewLine, keepLines);
                    }
                    else
                    {
                        tabDKCH_nowTracking.Text = string.Join(Environment.NewLine, allLines);
                    }
                    tabDKCH_nowTracking.AppendText(Environment.NewLine);
                    tabDKCH_nowTracking.SelectionStart = tabDKCH_nowTracking.Text.Length;
                    tabDKCH_nowTracking.ScrollToCaret();
                }
            }
        }

        private void UpdateWaybillCount()
        {
            if (tabDKCH_countSum == null) return;
            var uniqueCodes = tabTracking_inputWaybill.Text
                .Split(new[] { '\r', '\n', ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim().ToUpper())
                .Where(x => x.Length > 5)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            tabDKCH_countSum.Text = uniqueCodes.Count().ToString("N0");
        }

        private async void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
            try
            {
                if (tabHome_webView != null && !tabHome_webView.IsDisposed && tabHome_webView.CoreWebView2 != null)
                {
                    if (tabControl.SelectedTab != tabHome)
                        await tabHome_webView.CoreWebView2.TrySuspendAsync();
                    else
                        tabHome_webView.CoreWebView2.Resume();
                }
            }
            catch { }

            try
            {
                if (tabControl.SelectedTab == tabDKCH)
                {
                    if (_isDkchNeedReload && tabDKCH_webView != null && tabDKCH_webView.CoreWebView2 != null)
                    {
                        tabDKCH_webView.CoreWebView2.Reload();
                        _isDkchNeedReload = false;
                    }
                }
                else if (tabControl.SelectedTab == tabTracking)
                {
                    if (string.IsNullOrEmpty(Main.CapturedAuthToken))
                        await RefreshAuthTokenAsync();
                }
                else if (tabControl.SelectedTab == tabChat)
                {
                    if (!_isZaloLoaded && tabChat_webViewZalo.CoreWebView2 == null)
                    {
                        try
                        {
                            string userDataFolder = Path.Combine(Application.StartupPath, "ZaloProfile");
                            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                            await tabChat_webViewZalo.EnsureCoreWebView2Async(env);
                            tabChat_webViewZalo.CoreWebView2.Navigate("https://chat.zalo.me/index.html");
                            tabChat_webViewZalo.NavigationCompleted += (s, args) =>
                            {
                                if (_zaloChatService == null)
                                {
                                    _zaloChatService = new ZaloChatService(tabChat_webViewZalo, appsScriptUrl);
                                    _zaloChatService.StartAutoReminder(5);
                                }
                            };
                            await LocalTrackingManager.Instance.PerformIncrementalTrackingAsync();
                            // Cập nhật ComboBox trạng thái
                            PopulateStatusSelect();
                            // Áp dụng lọc hiện tại
                            ApplyFilterAndSort();
                            _isZaloLoaded = true;
                        }
                        catch (Exception ex) { MessageBox.Show("Lỗi khởi tạo Zalo Web: " + ex.Message); }
                    }
                    await RefreshZaloDataAsync(forceReload: false);
                }
                else if (tabControl.SelectedTab == tabDash)
                {
                    RefreshDashboardUI();
                }
            }
            catch (Exception ex) { MessageBox.Show("Lỗi xử lý Tab: " + ex.Message); }
        }

        private void CoreWebView2_WebResourceRequested(object sender, CoreWebView2WebResourceRequestedEventArgs e)
        {
            try
            {
                if (e.Request.Headers.Contains("authToken"))
                {
                    string token = e.Request.Headers.GetHeader("authToken");
                    if (!string.IsNullOrEmpty(token) && token.Length > 20)
                    {
                        CapturedAuthToken = token;
                        _settings.LastAuthToken = token;
                        SettingsManager.Save(_settings);
                    }
                }
            }
            catch { }
        }

        public async Task RefreshAuthTokenAsync()
        {
            if (tabDKCH_webView.CoreWebView2 == null) return;
            string js = @"(function() {
                let token = localStorage.getItem('authToken') || localStorage.getItem('token');
                if (!token || token.length < 30) {
                    const userData = localStorage.getItem('userData');
                    if (userData) {
                        try {
                            const obj = JSON.parse(userData);
                            if (obj.uuid && obj.uuid.length > 20) token = obj.uuid;
                        } catch(e){}
                    }
                }
                return { found: !!token && token.length > 20, value: token || '' };
            })();";
            try
            {
                string result = await WebViewHost.ExecJsAsync(js);
                var obj = JsonSerializer.Deserialize<JsonElement>(result);
                if (obj.GetProperty("found").GetBoolean())
                    CapturedAuthToken = obj.GetProperty("value").GetString() ?? "";
            }
            catch { }
        }

        private async Task ProcessDirectTrackingAsync()
        {
            _queueTimer.Stop();
            try
            {
                string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;
                string c5Value = "";
                var c5Data = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!C5");
                if (c5Data != null && c5Data.Count > 0 && c5Data[0].Count > 0)
                    c5Value = c5Data[0][0].ToString().Trim().ToUpper();

                if (c5Value == "RUN")
                {
                    GoogleSheetService.UpdateCell(spreadsheetId, "PHATLAI!C5", "PROCESSING");
                    var columnAData = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!A2:A");
                    List<string> waybills = new List<string>();
                    if (columnAData != null)
                    {
                        foreach (var row in columnAData)
                            if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0].ToString()))
                                waybills.Add(row[0].ToString().Trim());
                    }
                    if (waybills.Count > 0)
                    {
                        string waybillsText = string.Join("\n", waybills);
                        await _trackingService.SearchTrackingAsync(waybillsText, false);
                        var results = _trackingService.GetAllRows();
                        var sheetData = new List<IList<object>>();
                        foreach (var item in results)
                        {
                            sheetData.Add(new List<object>()
                            {
                                item.WaybillNo, item.TrangThaiHienTai, item.ThaoTacCuoi,
                                item.ThoiGianThaoTac, item.ThoiGianYeuCauPhatLai, item.NhanVienKienVanDe
                            });
                        }
                        GoogleSheetService.UpdateBumpSheet(sheetData, spreadsheetId, "BUMP!A2");
                    }
                    GoogleSheetService.UpdateCell(spreadsheetId, "PHATLAI!C5", "DONE");
                }
            }
            catch (Exception ex) { Console.WriteLine($"[Lỗi Tracking Direct] {ex.Message}"); }
            finally { _queueTimer.Start(); }
        }

        private async void tabHome_WebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (tabHome_webView.CoreWebView2 != null)
            {
                tabHome_btnBack.Enabled = tabHome_webView.CoreWebView2.CanGoBack;
                tabHome_btnForward.Enabled = tabHome_webView.CoreWebView2.CanGoForward;
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (tabHome_webView?.CoreWebView2 != null && tabHome_webView.CoreWebView2.CanGoBack)
                tabHome_webView.CoreWebView2.GoBack();
        }

        private void btnForward_Click(object sender, EventArgs e)
        {
            if (tabHome_webView?.CoreWebView2 != null && tabHome_webView.CoreWebView2.CanGoForward)
                tabHome_webView.CoreWebView2.GoForward();
        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            if (tabHome_webView?.CoreWebView2 != null)
            {
                tabHome_webView.CoreWebView2.Reload();
                tabHome_webView.CoreWebView2.Navigate("https://jms.jtexpress.vn");
            }
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            if (tabHome_webView?.CoreWebView2 != null)
                tabHome_webView.CoreWebView2.Navigate("https://jms.jtexpress.vn");
        }

        private async void tabDKCH_btnDKCH1_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning) return;
            tabDKCH_btnDKCH1.Enabled = false;
            tabDKCH_btnDKCH2.Enabled = false;
            await _dkchManager.StartAsync("DKCH1");
            tabDKCH_btnDKCH1.Visible = false;
            tabDKCH_btnDKCH2.Visible = false;
            tabDKCH_btnStop.Visible = true;
            await RefreshAuthTokenAsync();
        }

        private async void tabDKCH_btnDKCH2_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning) return;
            tabDKCH_btnDKCH1.Enabled = false;
            tabDKCH_btnDKCH2.Enabled = false;
            await _dkchManager.StartAsync("DKCH2");
            tabDKCH_btnDKCH1.Visible = false;
            tabDKCH_btnDKCH2.Visible = false;
            tabDKCH_btnStop.Visible = true;
        }

        private void tabDKCH_btnStop_Click(object sender, EventArgs e)
        {
            _dkchManager.Stop();
            tabDKCH_btnDKCH1.Visible = true;
            tabDKCH_btnDKCH2.Visible = true;
            tabDKCH_btnDKCH1.Enabled = true;
            tabDKCH_btnDKCH2.Enabled = true;
            tabDKCH_btnStop.Visible = false;
        }

        private async void btn_Refresh_Click(object sender, EventArgs e)
        {
            _dkchManager.Stop();
            await WebViewHost.NavigateAsync("https://jms.jtexpress.vn");
        }

        private async void btnSearch_Click(object sender, EventArgs e)
        {
            string input = tabTracking_inputWaybill.Text.Trim();
            if (string.IsNullOrWhiteSpace(input))
            {
                UIMessageTip.ShowWarning("Nhập mã vận đơn vào ô bên trên!");
                return;
            }
            try
            {
                tabTracking_btnSearch.Enabled = false;
                await _trackingService.SearchTrackingAsync(input);
                int tongSoDon = _trackingService.GetAllRows().Count;
                tabTracking_countSum.Text = tongSoDon.ToString("N0");
                if (tongSoDon == 0) UIMessageTip.ShowWarning("Không tìm thấy vận đơn nào!");
            }
            catch (Exception ex) { UIMessageTip.ShowError("Lỗi khi tra cứu: " + ex.Message); }
            finally { tabTracking_btnSearch.Enabled = true; UpdateWaybillCount(); }
        }

        private void btn_Export_Click(object sender, EventArgs e) => _trackingService?.ExportToExcel();
        private void btn_Clear_Click(object sender, EventArgs e) { _trackingService?.ClearData(); tabTracking_inputWaybill.Clear(); }
        private void btn_Export_Spe_Click(object sender, EventArgs e) => _trackingService.ExportSpecial();

        private async void tabTracking_btnUpload_Click(object sender, EventArgs e)
        {
            tabTracking_btnUpload.Enabled = false;
            string oldText = tabTracking_btnUpload.Text;
            tabTracking_btnUpload.Text = "Đang đồng bộ...";
            try
            {
                DataGridView grid = tabTracking_dataView;
                if (grid.Rows.Count == 0 && grid.Columns.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu trên bảng để tải lên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                var sheetData = new List<IList<object>>();
                var headerRow = new List<object>();
                foreach (DataGridViewColumn col in grid.Columns) headerRow.Add(col.HeaderText);
                sheetData.Add(headerRow);
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;
                    var rowData = new List<object>();
                    foreach (DataGridViewCell cell in row.Cells) rowData.Add(cell.Value?.ToString() ?? "");
                    sheetData.Add(rowData);
                }
                await Task.Run(() =>
                {
                    string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;
                    string targetSheetName = "BUMP";
                    GoogleSheetService.ClearSheet(spreadsheetId, targetSheetName);
                    GoogleSheetService.UpdateBumpSheet(sheetData, spreadsheetId, $"{targetSheetName}!A1");
                });
                MessageBox.Show("Đã tải lên thành công!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { tabTracking_btnUpload.Enabled = true; tabTracking_btnUpload.Text = oldText; }
        }

        private string _downloadFolderPath = Path.Combine(Application.StartupPath, "Downloads");
        private void btn_Download_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(_downloadFolderPath)) Directory.CreateDirectory(_downloadFolderPath);
                Process.Start(new ProcessStartInfo { FileName = _downloadFolderPath, UseShellExecute = true, Verb = "open" });
            }
            catch (Exception ex) { MessageBox.Show("Không thể mở thư mục: " + ex.Message); }
        }

        private async void print_TimKiem_Click(object sender, EventArgs e)
        {
            string input = tabPrint_inputWaybill.Text.Trim();
            await _printService.SearchAndLoadAsync(input, _printService.CurrentMode);
        }

        private void print_LamMoi_Click(object sender, EventArgs e)
        {
            _trackingService.ClearData();
            _printService.Reset();
            tabPrint_btnSelectAll.Checked = false;
            tabPrint_countSelect.Text = "000";
            tabPrint_countSum.Text = "000";
        }

        private void print_InChuyenHoan_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InHoan);
        private void print_InChuyenTiep_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InChuyenTiep);
        private void print_InLaiDon_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InLaiDon);
        private void print_InReverse_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InReverse);
        private void tabPrint_btnSelectAll_CheckedChanged(object sender, EventArgs e) => _printService.SelectAll(tabPrint_btnSelectAll.Checked);

        private async void tabPrint_btnPrint_Click(object sender, EventArgs e)
        {
            if (_printService == null)
            {
                MessageBox.Show("Chưa khởi tạo PrintService.", "Lỗi");
                return;
            }
            var selected = _printService.GetSelectedWaybills();
            if (selected == null || selected.Count == 0)
            {
                MessageBox.Show("Chưa chọn vận đơn nào!", "Thông báo");
                return;
            }
            var originalOrder = tabPrint_inputWaybill.Text
                .Split(new[] { '\r', '\n', ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim().ToUpper())
                .Where(x => x.Length > 5)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            selected = selected.OrderBy(wb => originalOrder.IndexOf(wb)).ToList();

            tabPrint_btnPrint.Enabled = false;
            tabPrint_btnPrint.Text = "Đang in...";
            try
            {
                int printType = 1;
                int applyTypeCode = (_printService.CurrentMode == PrintMode.InChuyenTiep) ? 2 : 4;
                string pdfUrl = await GetPdfUrlViaCSharpAsync(selected, printType, applyTypeCode);
                if (!string.IsNullOrEmpty(pdfUrl))
                {
                    string localPath = await DownloadPdfToTempFileAsync(pdfUrl);
                    if (!string.IsNullOrEmpty(localPath))
                    {
                        for (int i = 0; i < selected.Count; i++)
                        {
                            var psi = new ProcessStartInfo(localPath) { Verb = "print", UseShellExecute = true };
                            Process.Start(psi);
                            await Task.Delay(800);
                        }
                    }
                }
                else MessageBox.Show("Không lấy được PDF.", "Lỗi");
            }
            catch (Exception ex) { }
            finally { tabPrint_btnPrint.Enabled = true; tabPrint_btnPrint.Text = "IN"; }
        }

        private async Task<string> GetPdfUrlViaCSharpAsync(List<string> waybills, int printType, int applyTypeCode)
        {
            try
            {
                await RefreshAuthTokenAsync();
                string token = Main.CapturedAuthToken ?? "";
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Add("authToken", token);
                    client.DefaultRequestHeaders.Add("lang", "VN");
                    client.DefaultRequestHeaders.Add("langType", "VN");
                    client.DefaultRequestHeaders.Add("routeName", "returnAndForwardMaintain");
                    client.DefaultRequestHeaders.Add("routerNameList", "%E6%93%8D%E4%BD%9C%E5%B9%B3%E5%8F%B0%3E%E6%8B%A6%E6%88%AA%E9%80%80%E8%BD%AC%3E%E9%80%80%E8%BD%AC%E4%BB%B6%E7%AE%A1%E7%90%86");
                    client.DefaultRequestHeaders.Add("sec-ch-ua", "\"Not-A.Brand\";v=\"24\", \"Microsoft Edge\";v=\"146\", \"Chromium\";v=\"146\", \"Microsoft Edge WebView2\";v=\"146\"");
                    client.DefaultRequestHeaders.Add("sec-ch-ua-mobile", "?0");
                    client.DefaultRequestHeaders.Add("sec-ch-ua-platform", "\"Windows\"");
                    client.DefaultRequestHeaders.Add("timezone", "GMT+0700");
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36 Edg/146.0.0.0");
                    client.DefaultRequestHeaders.Add("Origin", "https://jms.jtexpress.vn");
                    client.DefaultRequestHeaders.Add("Referer", "https://jms.jtexpress.vn/");
                    client.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");
                    client.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.9");

                    var payload1 = new { current = 1, size = 20, pringFlag = 0, applyNetworkId = 5165, waybillIds = waybills, applyTimeFrom = "2020-01-01 00:00:00", applyTimeTo = "2030-01-01 23:59:59", pringType = printType, countryId = "1" };
                    var content1 = new StringContent(JsonSerializer.Serialize(payload1), Encoding.UTF8, "application/json");
                    await client.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/pringListPage", content1);

                    var payload2 = new { waybillIds = waybills, applyTypeCode = applyTypeCode, countryId = "1" };
                    var content2 = new StringContent(JsonSerializer.Serialize(payload2), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/printWaybill", content2);
                    string rawJson = await response.Content.ReadAsStringAsync();
                    using (JsonDocument doc = JsonDocument.Parse(rawJson))
                    {
                        if (doc.RootElement.TryGetProperty("data", out JsonElement data) && data.ValueKind == JsonValueKind.String)
                        {
                            string url = data.GetString();
                            if (url.StartsWith("http")) return url;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private async Task<string> DownloadPdfToTempFileAsync(string pdfUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(pdfUrl) || !Uri.TryCreate(pdfUrl.Trim(), UriKind.Absolute, out _))
                    throw new Exception("Link PDF không hợp lệ.");
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    byte[] bytes = await client.GetByteArrayAsync(pdfUrl.Trim());
                    string path = Path.Combine(Path.GetTempPath(), $"AutoJMS_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
                    File.WriteAllBytes(path, bytes);
                    return path;
                }
            }
            catch { return null; }
        }

        private async void tabChat_btnReload_Click(object sender, EventArgs e)
        {
            if (_zaloChatService == null)
            {
                MessageBox.Show("Vui lòng đợi Zalo khởi tạo xong!");
                return;
            }
            tabChat_btnReload.Enabled = false;
            tabChat_btnReload.Text = "Đang tải...";

            await LocalTrackingManager.Instance.PerformIncrementalTrackingAsync();
            PopulateStatusSelect();
            ApplyFilterAndSort();

            tabChat_btnReload.Enabled = true;
            tabChat_btnReload.Text = "Làm mới Data";
        }

        private async void mainNav_btnExit_Click(object sender, EventArgs e) => Application.Exit();
    }
}