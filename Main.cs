using AutoJMS.Data;
using AutoUpdaterDotNET;
using Microsoft.Web.WebView2.Core;
using PdfiumViewer;
using Sunny.UI;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Size = System.Drawing.Size;

namespace AutoJMS
{
    public partial class Main : UIForm
    {
        private static string JmsHomeUrl => AppConfig.Current.JmsBaseUrl.TrimEnd('/');
        private static string AppsScriptUrl => AppConfig.Current.AppsScriptUrl;

        private AppSettings _settings;
        public static string CapturedAuthToken = "";

        // ================= CÁC DỊCH VỤ CỐT LÕI =================
        public static WaybillTrackingService _trackingService;
        private DkchManager _dkchManager;
        private PrintService _printService;
        public ZaloChatService _zaloChatService;

        // ================= BACKGROUND SYNC =================
        private readonly CancellationTokenSource _appCts = new();
        private readonly SemaphoreSlim _syncGate = new(1, 1);
        private Task _syncLoopTask;
        private bool _isUpdatingUI = false;
        private List<WaybillDbModel> _cloudData = new();

        // ================= CỜ TRẠNG THÁI =================
        public bool _isZaloLoaded = false;
        private bool _isDkchNeedReload = true;
        private bool _isHomeNeedReload = false;
        private readonly object _authTokenLock = new object();
        private CancellationTokenSource _authTokenSaveCts;
        private System.Windows.Forms.Timer _dkchUiStateTimer;
        private bool _isDkchStarting = false;

        // ================= BỘ NHỚ IN ẤN & GIAO DIỆN =================
        private readonly Dictionary<string, (DateTime LastTime, int Count)> _printedHistory = new(StringComparer.OrdinalIgnoreCase);
        private readonly SemaphoreSlim _printLock = new(1, 1);
        private System.Windows.Forms.Timer hideTimer = null;
        private string _downloadFolderPath = Path.Combine(Application.StartupPath, "Downloads");

        private static readonly Regex DkchWaybillRegex = new Regex("^[A-Za-z0-9]{1,20}$", RegexOptions.Compiled);
        private const string CHROME_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
        private UILabel lblNetworkStatus;

        public Main()
        {
            InitializeComponent();
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.AutoScaleDimensions = new SizeF(96F, 96F);
            this.MinimumSize = new Size(1366, 768);
            _settings = SettingsManager.Load();

            File.WriteAllText("debug.log", "App started\n");
            AppDomain.CurrentDomain.UnhandledException += (s, e) => 
            { 
                File.WriteAllText("crash.log", e.ExceptionObject.ToString()); 
            };

            // UI styling
            tabDKCH_inputNewBill.Font = new System.Drawing.Font("Segoe UI Semibold", 12f, System.Drawing.FontStyle.Bold);
            tabDKCH_inputNewBill.WordWrap = false;
            tabDKCH_newBillDone.WordWrap = false;

            string sharedFolder = Path.Combine(Application.StartupPath, "SharedBrowserData");
            var sharedProps = new Microsoft.Web.WebView2.WinForms.CoreWebView2CreationProperties() 
            { 
                UserDataFolder = sharedFolder 
            };

            // ================= SETUP GRID BINDING =================
            SetupDashGridBinding();

            tabChat_dataGrid.AutoGenerateColumns = false;
            tabChat_dataGrid.Columns.Clear();
            tabChat_dataGrid.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "maDon", HeaderText = "Mã vận đơn", DataPropertyName = "WaybillNo", AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells, SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "nhanVien", HeaderText = "Tên nhân viên", DataPropertyName = "NguoiThaoTac", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "trangThai", HeaderText = "Trạng thái hiện tại", DataPropertyName = "ThaoTacCuoi", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "soLanNhac", HeaderText = "Số lần nhắc", DataPropertyName = "PrintCount", SortMode = DataGridViewColumnSortMode.Programmatic }
            });
            tabChat_dataGrid.ColumnHeaderMouseClick += tabChat_dataGrid_ColumnHeaderMouseClick;

            tabDash_updateData.Click += async (s, e) => await HandleAutoSyncCycleAsync(_appCts.Token);

            var doubleBufferPropertyInfo = tabTracking_dataView.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            doubleBufferPropertyInfo?.SetValue(tabTracking_dataView, true, null);

            // ================= CẤU HÌNH WEBVIEW =================
            tabHome_webView.CreationProperties = sharedProps;
            tabDKCH_webView.CreationProperties = sharedProps;
            tabPrint_printPreview.CreationProperties = sharedProps;

            tabDKCH_sheetName.SelectedItem = _settings.DefaultSheet;
            tabDKCH_useSheet.Active = _settings.UseSheetByDefault;
            tabDKCH_numRow.Value = _settings.DefaultRowCount;

            // ================= GẮN EVENT UI =================
            tabTracking_inputWaybill.KeyDown += tabTracking_inputWaybill_KeyDown;
            tabPrint_inputWaybill.KeyDown += tabPrint_inputWaybill_KeyDown;
            tabPrint_btnSelectAll.CheckedChanged += tabPrint_btnSelectAll_CheckedChanged;
            tabPrint_printFunc.SelectedIndexChanged += TabPrint_printFunc_SelectedIndexChanged;
            tabHome_webView.NavigationCompleted += tabHome_WebView_NavigationCompleted;

            // DKCH buttons
            tabDKCH_btnDKCH1.Visible = true;
            tabDKCH_btnDKCH2.Visible = true;
            tabDKCH_btnStop.Visible = false;
            tabDKCH_btnDKCH1.Enabled = false;
            tabDKCH_btnDKCH2.Enabled = false;
            UpdateDkchButtonsByState(false);

            _dkchUiStateTimer = new System.Windows.Forms.Timer();
            _dkchUiStateTimer.Interval = 300;
            _dkchUiStateTimer.Tick += (s, e) => UpdateDkchButtonsByState((_dkchManager?.IsRunning == true) || _isDkchStarting);
            _dkchUiStateTimer.Start();

            CheckForIllegalCrossThreadCalls = false;
            this.KeyPreview = true;

            _dkchManager = new DkchManager();
            _dkchManager.OnSaveCountChanged += (count) => tabDKCH_countSave.Text = $" OK: " + count.ToString();
            _dkchManager.OnTrackingHistoryChanged += (history) =>
            {
                if (this.InvokeRequired) 
                {
                    this.Invoke(new Action(() => FormatNowTracking(history ?? "Không có dữ liệu")));
                }
                else 
                {
                    FormatNowTracking(history ?? "Không có dữ liệu");
                }
            };
            _dkchManager.OnWaybillCompleted += AppendDoneWaybill;
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            InitNetworkUI();

            Version appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string versionString = appVersion.ToString(); // Trả về dạng chuẩn: 1.0.0.1

            if (tabAbout_lblVersion != null)
            {
                tabAbout_lblVersion.Text = $"Phiên bản hiện tại: v{versionString}";
            }

            if (this.IsDisposed) return;

            _ = tabHome_webView.Handle;
            _ = tabDKCH_webView.Handle;
            _ = tabPrint_printPreview.Handle;

            try
            {
                await tabHome_webView.EnsureCoreWebView2Async(null);
                await tabDKCH_webView.EnsureCoreWebView2Async(null);
                await tabPrint_printPreview.EnsureCoreWebView2Async(null);

                tabHome_webView.CoreWebView2.Settings.UserAgent = CHROME_USER_AGENT;
                tabDKCH_webView.CoreWebView2.Settings.UserAgent = CHROME_USER_AGENT;
                tabPrint_printPreview.CoreWebView2.Settings.UserAgent = CHROME_USER_AGENT;

                tabHome_webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
                tabHome_webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;
                tabDKCH_webView.CoreWebView2.Settings.IsPasswordAutosaveEnabled = true;
                tabDKCH_webView.CoreWebView2.Settings.IsGeneralAutofillEnabled = true;

                tabHome_webView.CoreWebView2.Navigate(JmsHomeUrl);
                tabDKCH_webView.CoreWebView2.Navigate(JmsHomeUrl);

                tabHome_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Fetch);
                tabHome_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.XmlHttpRequest);
                tabHome_webView.CoreWebView2.WebResourceRequested += CoreWebView2_WebResourceRequested;

                tabDKCH_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Fetch);
                tabDKCH_webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.XmlHttpRequest);
                tabDKCH_webView.CoreWebView2.WebResourceRequested += CoreWebView2_WebResourceRequested;
                tabDKCH_webView.CoreWebView2.NavigationCompleted += (s, args) => { if (args.IsSuccess) ApplyZoomFactor(); };
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khởi tạo trình duyệt: " + ex.Message);
            }

            // ================= KHỞI TẠO SERVICE =================
            _trackingService = new WaybillTrackingService(tabTracking_dataView, tabTracking_process);
            _printService = new PrintService(tabPrint_dataView, _trackingService);
            tabPrint_dataView.Visible = false;

            _printService.OnPrintStatsChanged += (selectedCount, totalCount) =>
            {
                this.Invoke((MethodInvoker)delegate {
                    tabPrint_countSelect.Text = "Đang chọn:" + selectedCount.ToString();
                    tabPrint_countSum.Text = "Tổng:" + totalCount.ToString();
                });
            };

            ApplyStandardGridSettings(tabChat_dataGrid);
            ApplyStandardGridSettings(tabDash_dataGridView);
            ApplyStandardGridSettings(tabTracking_dataView);
            ApplyStandardGridSettings(tabPrint_dataView);

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

            // ================= CHẠY BACKGROUND WORKER =================
            try
            {
                await SupabaseDbService.InitializeAsync();
                _syncLoopTask = Task.Run(() => AutoSyncLoopAsync(_appCts.Token));
            }
            catch (Exception ex)
            {
                AppLogger.Error("Lỗi khởi tạo Supabase", ex);
            }
        }

        // ======================================================================================
        // HỆ THỐNG ĐỒNG BỘ BACKGROUND (SUPABASE CLOUD)
        // ======================================================================================

        private async Task AutoSyncLoopAsync(CancellationToken ct)
        {
            // Chạy ngay lần đầu tiên
            try { await HandleAutoSyncCycleAsync(ct); } catch { }

            using var timer = new PeriodicTimer(TimeSpan.FromMinutes(5));
            while (await timer.WaitForNextTickAsync(ct))
            {
                try
                {
                    await HandleAutoSyncCycleAsync(ct);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex) { AppLogger.Error("Lỗi vòng sync tự động", ex); }
            }
        }

        private async Task HandleAutoSyncCycleAsync(CancellationToken ct)
        {
            if (!await _syncGate.WaitAsync(0, ct)) return;

            try
            {
                // 1. Quét Tồn Kho (Có Lease Lock bảo vệ bên trong)
                await InventorySyncService.RunInventorySyncAsync(ct);

                // 2. Lấy đơn đến hạn & Tracking
                var dueWaybills = await SupabaseDbService.GetWaybillsDueForTrackingAsync();
                if (dueWaybills.Count > 0)
                {
                    await DatabaseTracking.RunBackgroundTrackingAsync(dueWaybills, ct);
                }

                if (ct.IsCancellationRequested) return;

                // 3. Tải Data cập nhật UI
                _cloudData = await SupabaseDbService.GetActiveWaybillsAsync();
                await UpdateDashAndChatViewsAsync(ct);
            }
            finally
            {
                _syncGate.Release();
            }
        }
        private void tabchat_datagrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
        }
        private async Task UpdateDashAndChatViewsAsync(CancellationToken ct)
        {
            if (_isUpdatingUI) return;
            _isUpdatingUI = true;

            try
            {
                string dataSourceOption = GetControlTextSafe(tabDash_dataSource);
                string selectedStatus = GetControlTextSafe(tabDash_statusSelect);

                List<string> phatLaiWaybills = new();

                if (dataSourceOption == "PHATLAI")
                {
                    phatLaiWaybills = await ZaloChatService.GetWaybillsFromPhatLaiAsync();
                }

                var baseData = _cloudData;

                if (dataSourceOption == "PHATLAI" && phatLaiWaybills.Count > 0)
                {
                    var set = new HashSet<string>(phatLaiWaybills, StringComparer.OrdinalIgnoreCase);
                    baseData = baseData.Where(x => set.Contains(x.WaybillNo)).ToList();
                }

                List<WaybillDbModel> dashData;

                if (string.IsNullOrWhiteSpace(selectedStatus) || selectedStatus == "Tất cả")
                {
                    dashData = baseData.OrderBy(x => x.WaybillNo).ToList();
                }
                else
                {
                    var multiStatuses = selectedStatus.Split(',')
                                                      .Select(s => s.Trim())
                                                      .Where(s => s.Length > 0)
                                                      .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    dashData = baseData.Where(x => multiStatuses.Contains(x.ThaoTacCuoi ?? ""))
                                       .OrderBy(x => x.WaybillNo).ToList();
                }

                if (tabDash_dataGridView.InvokeRequired)
                {
                    tabDash_dataGridView.Invoke(new Action(() => { tabDash_dataGridView.DataSource = dashData; }));
                }
                else
                {
                    tabDash_dataGridView.DataSource = dashData;
                }

                ApplyChatFilter(baseData);

                if (tabDash_lblLastUpdate.InvokeRequired)
                {
                    tabDash_lblLastUpdate.Invoke(new Action(() =>
                    {
                        PopulateStatusSelects();
                        tabDash_lblLastUpdate.Text = $"Cập nhật lần cuối: {DateTime.Now:HH:mm:ss}";
                        tabDash_lblDebug.Text = $"Đã load {baseData.Count} đơn từ Cloud";
                        if (tabChat_sumFollow != null)
                        {
                            tabChat_sumFollow.Text = $"Tổng đang theo dõi: {tabChat_dataGrid.RowCount}";
                            int kienVanDeCount = baseData.Count(r => r.TrangThaiHienTai?.Contains("vấn đề") == true);
                            if (tabChat_hasKVD != null) tabChat_hasKVD.Text = $"Kiện vấn đề: {kienVanDeCount}";
                            int xnchCount = baseData.Count(r => r.TrangThaiHienTai?.Contains("Xác nhận chuyển hoàn") == true);
                            if (tabChat_hasXNCH != null) tabChat_hasXNCH.Text = $"Xác nhận CH: {xnchCount}";
                        }
                    }));
                }
                else
                {
                    PopulateStatusSelects();
                    tabDash_lblLastUpdate.Text = $"Cập nhật lần cuối: {DateTime.Now:HH:mm:ss}";
                    tabDash_lblDebug.Text = $"Đã load {baseData.Count} đơn từ Cloud";
                }
            }
            finally
            {
                _isUpdatingUI = false;
            }
        }

        private void ApplyChatFilter(List<WaybillDbModel> baseData)
        {
            string selected = GetControlTextSafe(tabChat_statusSelect);
            List<WaybillDbModel> filtered;

            if (selected == "Tất cả" || string.IsNullOrEmpty(selected))
            {
                filtered = baseData.OrderBy(x => x.NguoiThaoTac).ToList();
            }
            else
            {
                var multiStatuses = selected.Split(',').Select(s => s.Trim()).Where(s => s.Length > 0).ToList();
                filtered = baseData.Where(x => multiStatuses.Contains(x.ThaoTacCuoi ?? "")).OrderBy(x => x.NguoiThaoTac).ToList();
            }

            if (tabChat_dataGrid.InvokeRequired)
            {
                tabChat_dataGrid.Invoke(new Action(() => { tabChat_dataGrid.DataSource = filtered; }));
            }
            else
            {
                tabChat_dataGrid.DataSource = filtered;
            }
        }

        // ======================================================================================
        // CÁC HÀM UI HỖ TRỢ VÀ TAB DASH / CHAT
        // ======================================================================================

        private string GetControlTextSafe(Control ctrl)
        {
            if (ctrl.InvokeRequired)
            {
                return (string)ctrl.Invoke(new Func<string>(() => ctrl.Text?.Trim() ?? ""));
            }
            return ctrl.Text?.Trim() ?? "";
        }

        private void PopulateStatusSelects()
        {
            var statuses = _cloudData.Select(x => x.ThaoTacCuoi).Where(s => !string.IsNullOrEmpty(s)).Distinct().OrderBy(s => s).ToList();

            if (tabDash_statusSelect != null)
            {
                string currentDash = tabDash_statusSelect.Text;
                tabDash_statusSelect.Items.Clear();
                tabDash_statusSelect.Items.Add("Tất cả");
                foreach (var s in statuses) tabDash_statusSelect.Items.Add(s);
                if (tabDash_statusSelect.Items.Contains(currentDash)) tabDash_statusSelect.Text = currentDash;
                else tabDash_statusSelect.SelectedIndex = 0;
            }

            if (tabChat_statusSelect != null)
            {
                string currentChat = tabChat_statusSelect.Text;
                tabChat_statusSelect.Items.Clear();
                tabChat_statusSelect.Items.Add("Tất cả");
                foreach (var s in statuses) tabChat_statusSelect.Items.Add(s);
                if (tabChat_statusSelect.Items.Contains(currentChat)) tabChat_statusSelect.Text = currentChat;
                else tabChat_statusSelect.SelectedIndex = 0;
            }
        }

        private void tabDash_statusSelect_SelectedIndexChanged(object sender, EventArgs e) => _ = UpdateDashAndChatViewsAsync(_appCts.Token);
        private void tabChat_statusSelect_SelectedIndexChanged(object sender, EventArgs e) => _ = UpdateDashAndChatViewsAsync(_appCts.Token);

        private void tabDash_timeUpdateData_SelectedIndexChanged(object sender, EventArgs e)
        {
            string text = tabDash_timeUpdateData.Text?.Replace(" PHÚT", "")?.Replace(" GIỜ", "")?.Trim() ?? "30";
            int minutes = 30;
            if (tabDash_timeUpdateData.Text.Contains("GIỜ"))
            {
                if (int.TryParse(text, out int h)) minutes = h * 60;
            }
            else
            {
                int.TryParse(text, out minutes);
            }

            // Ghi chú: Chỉnh sửa interval ở đây cho PeriodicTimer không trực tiếp được như System.Timer
            // Bạn có thể lưu vào setting để lần sau load lại hoặc bỏ qua nếu dùng default 5 min
        }

        private void SetupDashGridBinding()
        {
            var dgv = tabDash_dataGridView;
            dgv.VirtualMode = false;
            dgv.AutoGenerateColumns = false;
            dgv.MultiSelect = true;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dgv.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dgv.ColumnHeaderMouseClick += tabDash_dataGrid_ColumnHeaderMouseClick;

            dgv.Columns.Clear();
            dgv.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Mã vận đơn", HeaderText = "Mã vận đơn", DataPropertyName = "WaybillNo", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Trạng thái hiện tại", HeaderText = "Trạng thái hiện tại", DataPropertyName = "TrangThaiHienTai", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Loại quét kiện cuối", HeaderText = "Loại quét kiện cuối", DataPropertyName = "ThaoTacCuoi", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Tên nhân viên", HeaderText = "Tên nhân viên", DataPropertyName = "NguoiThaoTac", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thời gian thao tác", HeaderText = "Thời gian thao tác", DataPropertyName = "ThoiGianThaoTac", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thời gian yêu cầu phát lại", HeaderText = "Thời gian yêu cầu phát lại", DataPropertyName = "ThoiGianYeuCauPhatLai", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Số lần nhắc", HeaderText = "Số lần nhắc", DataPropertyName = "PrintCount", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Cập nhật lần cuối", HeaderText = "Cập nhật lần cuối", DataPropertyName = "LastTrackedAt", SortMode = DataGridViewColumnSortMode.Programmatic }
            });
        }

        private void ApplyStandardGridSettings(DataGridView grid)
        {
            if (grid == null) return;
            grid.ReadOnly = true;
            grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.AllowUserToResizeColumns = true;
            grid.AllowUserToResizeRows = false;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.MultiSelect = true;
            grid.RowHeadersVisible = false;
            if (grid is Sunny.UI.UIDataGridView uiGrid)
            {
                uiGrid.StripeOddColor = System.Drawing.Color.White;
                uiGrid.StripeEvenColor = System.Drawing.Color.White;
            }
        }

        private void tabDash_dataGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var dgv = tabDash_dataGridView;
            if (dgv == null || e.RowIndex != -1 || e.ColumnIndex < 0) return;
            var clicked = dgv.Columns[e.ColumnIndex];
            if (clicked == null || string.IsNullOrEmpty(clicked.DataPropertyName)) return;

            var direction = clicked.HeaderCell.SortGlyphDirection == SortOrder.Ascending ? ListSortDirection.Descending : ListSortDirection.Ascending;

            var dataSource = dgv.DataSource as List<WaybillDbModel>;
            if (dataSource != null)
            {
                var propInfo = typeof(WaybillDbModel).GetProperty(clicked.DataPropertyName);
                if (propInfo != null)
                {
                    if (direction == ListSortDirection.Ascending) dgv.DataSource = dataSource.OrderBy(x => propInfo.GetValue(x, null)).ToList();
                    else dgv.DataSource = dataSource.OrderByDescending(x => propInfo.GetValue(x, null)).ToList();
                }
            }

            clicked.HeaderCell.SortGlyphDirection = direction == ListSortDirection.Ascending ? SortOrder.Ascending : SortOrder.Descending;
            foreach (DataGridViewColumn col in dgv.Columns) if (col != clicked) col.HeaderCell.SortGlyphDirection = SortOrder.None;
        }

        private void tabChat_dataGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return;
            var column = tabChat_dataGrid.Columns[e.ColumnIndex];
            if (column == null || string.IsNullOrEmpty(column.DataPropertyName)) return;

            var direction = column.HeaderCell.SortGlyphDirection == SortOrder.Ascending ? ListSortDirection.Descending : ListSortDirection.Ascending;

            var dataSource = tabChat_dataGrid.DataSource as List<WaybillDbModel>;
            if (dataSource != null)
            {
                var propInfo = typeof(WaybillDbModel).GetProperty(column.DataPropertyName);
                if (propInfo != null)
                {
                    if (direction == ListSortDirection.Ascending) tabChat_dataGrid.DataSource = dataSource.OrderBy(x => propInfo.GetValue(x, null)).ToList();
                    else tabChat_dataGrid.DataSource = dataSource.OrderByDescending(x => propInfo.GetValue(x, null)).ToList();
                }
            }

            column.HeaderCell.SortGlyphDirection = direction == ListSortDirection.Ascending ? SortOrder.Ascending : SortOrder.Descending;
            foreach (DataGridViewColumn col in tabChat_dataGrid.Columns) if (col != column) col.HeaderCell.SortGlyphDirection = SortOrder.None;
        }

        // ======================================================================================
        // QUẢN LÝ APP & MẠNG 
        // ======================================================================================

        private bool _isExiting = false;

        private void Main_FormClosing(object? sender, FormClosingEventArgs e)
        {
            if (_isExiting) return;

            if (e.CloseReason == CloseReason.UserClosing)
            {
                bool confirm = ShowCustomExitDialog();
                if (!confirm)
                {
                    e.Cancel = true;
                    return;
                }
            }

            _isExiting = true;

            this.Hide(); 
            _appCts.Cancel(); 

            Task.Run(async () => {
                try
                {
                    await SupabaseDbService.ReleaseInventoryLeaseAsync();
                    // Nếu bạn có dùng khóa tracking thì mở dòng dưới
                    // await SupabaseDbService.ReleaseLeaseAsync("tracking_worker"); 
                }
                catch { }
            }).Wait(1000);
        }

        private bool ShowCustomExitDialog()
        {
            using (UIForm form = new UIForm())
            {
                form.Text = "Đóng ứng dụng";
                form.ClientSize = new System.Drawing.Size(450, 220);
                form.StartPosition = FormStartPosition.CenterScreen;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular);

                UILabel lblMsg = new UILabel();
                lblMsg.Text = "Cứ ngỡ cống hiến trăm năm...\nAi ngờ 5h00.pm";
                lblMsg.Font = new System.Drawing.Font("Tahoma", 13F, System.Drawing.FontStyle.Regular);
                lblMsg.Location = new System.Drawing.Point(20, 60);
                lblMsg.Size = new System.Drawing.Size(410, 70);
                lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                form.Controls.Add(lblMsg);

                UIButton btnYes = new UIButton();
                btnYes.Text = "Thoát ngay";
                btnYes.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
                btnYes.Size = new System.Drawing.Size(140, 40);
                btnYes.Location = new System.Drawing.Point(60, 150);
                btnYes.DialogResult = DialogResult.Yes;
                btnYes.FillColor = System.Drawing.Color.IndianRed;
                btnYes.RectColor = System.Drawing.Color.IndianRed;
                btnYes.FillHoverColor = System.Drawing.Color.Red;
                btnYes.RectHoverColor = System.Drawing.Color.DarkRed;
                btnYes.FillPressColor = System.Drawing.Color.Maroon;
                btnYes.RectPressColor = System.Drawing.Color.Maroon;
                form.Controls.Add(btnYes);

                UIButton btnNo = new UIButton();
                btnNo.Text = "Hủy bỏ";
                btnNo.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
                btnNo.Size = new System.Drawing.Size(140, 40);
                btnNo.Location = new System.Drawing.Point(250, 150);
                btnNo.DialogResult = DialogResult.No;
                form.Controls.Add(btnNo);

                return form.ShowDialog() == DialogResult.Yes;
            }
        }

        private void InitNetworkUI()
        {
            lblNetworkStatus = new UILabel();
            lblNetworkStatus.AutoSize = true;
            lblNetworkStatus.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
            lblNetworkStatus.BackColor = Color.Transparent;
            lblNetworkStatus.TextAlign = ContentAlignment.MiddleRight;
            lblNetworkStatus.Parent = this;
            lblNetworkStatus.BringToFront();
            this.Controls.Add(lblNetworkStatus);
            
            UpdateNetworkUI(NetworkStatus.Online);
            NetworkState.OnChanged += UpdateNetworkUI;
            this.SizeChanged += (s, e) => RepositionNetworkLabel();
        }

        private void RepositionNetworkLabel()
        {
            if (lblNetworkStatus != null)
                lblNetworkStatus.Location = new Point(this.Width - lblNetworkStatus.Width - 100, 7);
        }

        private void UpdateNetworkUI(NetworkStatus status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateNetworkUI(status)));
                return;
            }

            switch (status)
            {
                case NetworkStatus.Online:
                    lblNetworkStatus.Text = "● Online";
                    lblNetworkStatus.ForeColor = Color.FromArgb(0, 240, 100);
                    break;
                case NetworkStatus.Unstable:
                    lblNetworkStatus.Text = "● Mạng chậm";
                    lblNetworkStatus.ForeColor = Color.FromArgb(253, 224, 71);
                    break;
                case NetworkStatus.Offline:
                    lblNetworkStatus.Text = "● Mất kết nối";
                    lblNetworkStatus.ForeColor = Color.FromArgb(252, 115, 115);
                    break;
            }
            RepositionNetworkLabel();
        }

        private void tabAbout_btnCheckUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                AutoUpdater.ShowSkipButton = false;
                AutoUpdater.ReportErrors = true;

                // 1. Lấy chuẩn Version đang chạy
                Version appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;

                AutoUpdater.InstalledVersion = appVersion;

                AutoUpdater.HttpUserAgent = $"AutoJMS/{appVersion}";
                AutoUpdater.ExecutablePath = Path.Combine(Application.StartupPath, "AutoJMS.exe");

                string xmlUrl = AppConfig.Current.UpdateXmlUrl;
                if (string.IsNullOrWhiteSpace(xmlUrl))
                {
                    MessageBox.Show("Chưa cấu hình file update.", "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                AutoUpdater.Start(xmlUrl);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể kiểm tra cập nhật.\n\nChi tiết: " + ex.Message, "Update Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btn_Download_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(_downloadFolderPath)) Directory.CreateDirectory(_downloadFolderPath);
                Process.Start(new ProcessStartInfo { FileName = _downloadFolderPath, UseShellExecute = true, Verb = "open" });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể mở thư mục: " + ex.Message);
            }
        }

        // ======================================================================================
        // CÁC HÀM WEBVIEW, TOKEN & TAB ĐIỀU HƯỚNG
        // ======================================================================================

        private async void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.IsDisposed) return;
            try
            {
                if (tabControl.SelectedTab == tabHome)
                {
                    if (_isHomeNeedReload && tabHome_webView != null && tabHome_webView.CoreWebView2 != null)
                    {
                        tabHome_webView.CoreWebView2.Reload();
                        _isHomeNeedReload = false;
                    }
                }
                else if (tabControl.SelectedTab == tabDKCH)
                {
                    if (_isDkchNeedReload && tabDKCH_webView != null && tabDKCH_webView.CoreWebView2 != null)
                    {
                        tabDKCH_webView.CoreWebView2.Reload();
                        _isDkchNeedReload = false;
                    }
                }
                else if (tabControl.SelectedTab == tabTracking)
                {
                    if (string.IsNullOrEmpty(Main.CapturedAuthToken)) await RefreshAuthTokenAsync();
                    if (tabTracking_btnSearch.Enabled)
                    {
                        if (tabTracking_process != null && !tabTracking_process.IsDisposed)
                        {
                            tabTracking_process.Value = 0;
                            tabTracking_process.Visible = false;
                        }
                    }
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
                            tabChat_webViewZalo.CoreWebView2.Settings.UserAgent = CHROME_USER_AGENT;
                            tabChat_webViewZalo.CoreWebView2.Navigate("https://chat.zalo.me/index.html");
                            tabChat_webViewZalo.NavigationCompleted += (s, args) =>
                            {
                                if (_zaloChatService == null)
                                {
                                    _zaloChatService = new ZaloChatService(tabChat_webViewZalo, AppsScriptUrl);
                                    _zaloChatService.StartAutoReminder(5);
                                }
                            };
                            _isZaloLoaded = true;
                        }
                        catch (Exception ex) { MessageBox.Show("Lỗi khởi tạo Zalo Web: " + ex.Message); }
                    }
                    await UpdateDashAndChatViewsAsync(_appCts.Token);
                }
                else if (tabControl.SelectedTab == tabDash)
                {
                    await UpdateDashAndChatViewsAsync(_appCts.Token);
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
                        bool changed = false;
                        lock (_authTokenLock)
                        {
                            if (!string.Equals(_settings.LastAuthToken, token, StringComparison.Ordinal) || CapturedAuthToken != token)
                            {
                                _settings.LastAuthToken = token;
                                CapturedAuthToken = token;
                                changed = true;
                            }
                        }
                        if (changed)
                        {
                            _ = SaveAuthTokenDebouncedAsync();
                            if (this.InvokeRequired)
                            {
                                this.Invoke(new Action(() =>
                                {
                                    if (tabControl.SelectedTab == tabHome) _isDkchNeedReload = true;
                                    else if (tabControl.SelectedTab == tabDKCH) _isHomeNeedReload = true;
                                }));
                            }
                            else
                            {
                                if (tabControl.SelectedTab == tabHome) _isDkchNeedReload = true;
                                else if (tabControl.SelectedTab == tabDKCH) _isHomeNeedReload = true;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        private async Task SaveAuthTokenDebouncedAsync()
        {
            CancellationTokenSource currentCts;
            lock (_authTokenLock)
            {
                _authTokenSaveCts?.Cancel();
                _authTokenSaveCts?.Dispose();
                _authTokenSaveCts = new CancellationTokenSource();
                currentCts = _authTokenSaveCts;
            }
            try
            {
                await Task.Delay(400, currentCts.Token);
                AppSettings snapshot;
                lock (_authTokenLock)
                {
                    snapshot = new AppSettings
                    {
                        ZoomFactor = _settings.ZoomFactor,
                        DefaultUrl = _settings.DefaultUrl,
                        LastAuthToken = _settings.LastAuthToken,
                        DownloadFolder = _settings.DownloadFolder,
                        DefaultSheet = _settings.DefaultSheet,
                        UseSheetByDefault = _settings.UseSheetByDefault,
                        AutoRefreshToken = _settings.AutoRefreshToken,
                        LastMode = _settings.LastMode,
                        DefaultRowCount = _settings.DefaultRowCount,
                        PrinterName = _settings.PrinterName,
                        PaperWidth = _settings.PaperWidth,
                        PaperHeight = _settings.PaperHeight
                    };
                }
                await SettingsManager.SaveAsync(snapshot);
            }
            catch (OperationCanceledException) { }
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
                return JSON.stringify({ found: !!token && token.length > 20, value: token || '' }); 
            })();";
            try
            {
                string result = await WebViewHost.ExecJsAsync(js);
                string unescaped = JsonSerializer.Deserialize<string>(result);
                using (JsonDocument doc = JsonDocument.Parse(unescaped))
                {
                    var root = doc.RootElement;
                    if (root.GetProperty("found").GetBoolean())
                        CapturedAuthToken = root.GetProperty("value").GetString() ?? "";
                }
            }
            catch { }
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
                tabHome_webView.CoreWebView2.Navigate(JmsHomeUrl);
            }
        }

        private void btnHome_Click(object sender, EventArgs e)
        {
            if (tabHome_webView?.CoreWebView2 != null)
                tabHome_webView.CoreWebView2.Navigate(JmsHomeUrl);
        }

        // ======================================================================================
        // TAB ZALO CHAT
        // ======================================================================================

        private async void tabChat_btnStart_Click(object sender, EventArgs e)
        {
            if (!_isZaloLoaded) { MessageBox.Show("Zalo chưa sẵn sàng!"); return; }
            tabChat_btnStart.Enabled = false;

            try
            {
                var reminders = new List<Reminder>();
                var modelsToUpdate = new List<WaybillDbModel>();

                foreach (var item in _cloudData)
                {
                    if (item.TrangThaiHienTai?.Trim().Equals("Quét phát hàng", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        reminders.Add(new Reminder { maDon = item.WaybillNo ?? "", nhanVien = item.NguoiThaoTac ?? "", trangThai = item.TrangThaiHienTai });
                    }
                }

                if (reminders.Count == 0) { MessageBox.Show("Không có đơn nào ở trạng thái 'Quét phát hàng'."); return; }

                var groups = reminders.GroupBy(r => r.nhanVien.Trim(), StringComparer.OrdinalIgnoreCase).ToList();
                int totalGroups = groups.Count;
                int successCount = 0;

                for (int i = 0; i < totalGroups; i++)
                {
                    var group = groups[i];
                    string tenNV = group.Key;
                    var danhSachMa = group.Select(r => r.maDon).Distinct().ToList();
                    string danhSachMaDon = string.Join("\n", danhSachMa);
                    string noiDung = $"@{tenNV}\n{danhSachMaDon}";

                    bool result = await _zaloChatService.SendZaloMessage(noiDung);

                    if (result)
                    {
                        successCount++;
                        foreach (var wb in danhSachMa)
                        {
                            var target = _cloudData.FirstOrDefault(x => x.WaybillNo == wb);
                            if (target != null)
                            {
                                target.PrintCount++;
                                modelsToUpdate.Add(target);
                            }
                        }
                        await UpdateDashAndChatViewsAsync(_appCts.Token);
                        await Task.Delay(2500);
                    }
                    else
                    {
                        if (MessageBox.Show($"Gửi thất bại cho nhân viên: {tenNV}\n\nTiếp tục gửi?", "Lỗi gửi tin", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) break;
                    }
                }

                if (modelsToUpdate.Count > 0) await SupabaseDbService.UpsertManyWaybillsAsync(modelsToUpdate);
                MessageBox.Show($"Hoàn tất! Đã gửi thành công {successCount}/{totalGroups} nhân viên.", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { tabChat_btnStart.Enabled = true; }
        }

        private async void tabChat_btnReload_Click(object sender, EventArgs e)
        {
            if (_zaloChatService == null) { MessageBox.Show("Vui lòng đợi Zalo khởi tạo xong!"); return; }
            tabChat_btnReload.Enabled = false; tabChat_btnReload.Text = "Đang tải...";
            await HandleAutoSyncCycleAsync(_appCts.Token);
            tabChat_btnReload.Enabled = true; tabChat_btnReload.Text = "-Làm mới-";
        }

        // ======================================================================================
        // TAB DKCH
        // ======================================================================================

        private async void tabDKCH_btnDKCH1_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning || _isDkchStarting) return;
            _isDkchStarting = true;
            UpdateDkchButtonsByState(true);
            try
            {
                await _dkchManager.StartAsync("DKCH1");
                await RefreshAuthTokenAsync();
                UpdateDkchButtonsByState(_dkchManager.IsRunning);
            }
            catch (Exception ex)
            {
                _dkchManager.Stop();
                UpdateDkchButtonsByState(false);
                UIMessageTip.ShowError("Không thể khởi động DKCH1: " + ex.Message);
            }
            finally
            {
                _isDkchStarting = false;
                UpdateDkchButtonsByState(_dkchManager.IsRunning);
            }
        }

        private async void tabDKCH_btnDKCH2_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning || _isDkchStarting) return;
            _isDkchStarting = true;
            UpdateDkchButtonsByState(true);
            try
            {
                await _dkchManager.StartAsync("DKCH2");
                UpdateDkchButtonsByState(_dkchManager.IsRunning);
            }
            catch (Exception ex)
            {
                _dkchManager.Stop();
                UpdateDkchButtonsByState(false);
                UIMessageTip.ShowError("Không thể khởi động DKCH2: " + ex.Message);
            }
            finally
            {
                _isDkchStarting = false;
                UpdateDkchButtonsByState(_dkchManager.IsRunning);
            }
        }

        private void tabDKCH_btnStop_Click(object sender, EventArgs e)
        {
            _isDkchStarting = false;
            _dkchManager.Stop();
            UpdateDkchButtonsByState(false);
        }

        private async void btn_Refresh_Click(object sender, EventArgs e)
        {
            _dkchManager.Stop();
            await WebViewHost.NavigateAsync(JmsHomeUrl);
        }

        private void UpdateDkchButtonsByState(bool isRunning)
        {
            if (tabDKCH_btnDKCH1 == null || tabDKCH_btnDKCH2 == null || tabDKCH_btnStop == null) return;
            tabDKCH_btnDKCH1.Visible = !isRunning;
            tabDKCH_btnDKCH2.Visible = !isRunning;
            tabDKCH_btnDKCH1.Enabled = !isRunning;
            tabDKCH_btnDKCH2.Enabled = !isRunning;
            tabDKCH_btnStop.Visible = isRunning;
            tabDKCH_btnStop.Enabled = isRunning;
            if (isRunning) tabDKCH_btnStop.BringToFront();
        }

        private void tabDKCH_inputNewBill_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                e.Handled = true;
                var allLines = tabDKCH_inputNewBill.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => !string.IsNullOrWhiteSpace(x) && IsValidDkchWaybill(x)).ToList();
                if (allLines.Count == 0) return;
                tabDKCH_inputNewBill.Text = "";
                foreach (var waybill in allLines)
                {
                    AppendToNewBillDone(waybill);
                    _dkchManager.AddPriorityWaybill(waybill);
                }
            }
        }

        private void AppendToNewBillDone(string waybill)
        {
            if (tabDKCH_newBillDone == null || tabDKCH_newBillDone.IsDisposed) return;
            if (tabDKCH_newBillDone.InvokeRequired)
            {
                tabDKCH_newBillDone.Invoke(new Action(() => AppendToNewBillDone(waybill)));
                return;
            }
            var currentLines = tabDKCH_newBillDone.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).ToList();
            if (!currentLines.Any(x => string.Equals(x, waybill, StringComparison.OrdinalIgnoreCase)))
            {
                currentLines.Add(waybill);
            }
            tabDKCH_newBillDone.Text = string.Join(Environment.NewLine, currentLines);
            tabDKCH_newBillDone.SelectionStart = tabDKCH_newBillDone.TextLength;
            tabDKCH_newBillDone.ScrollToCaret();
        }

        private void AppendDoneWaybill(string waybill)
        {
            if (tabDKCH_newBillDone == null || tabDKCH_newBillDone.IsDisposed) return;
            if (tabDKCH_newBillDone.InvokeRequired)
            {
                tabDKCH_newBillDone.Invoke(new Action(() => AppendDoneWaybill(waybill)));
                return;
            }
            var currentLines = tabDKCH_newBillDone.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim()).Where(x => !string.Equals(x, waybill, StringComparison.OrdinalIgnoreCase)).ToList();
            tabDKCH_newBillDone.Text = string.Join(Environment.NewLine, currentLines);
            tabDKCH_newBillDone.SelectionStart = tabDKCH_newBillDone.TextLength;
            tabDKCH_newBillDone.ScrollToCaret();
        }

        private static bool IsValidDkchWaybill(string waybill)
        {
            if (string.IsNullOrWhiteSpace(waybill)) return false;
            return DkchWaybillRegex.IsMatch(waybill.Trim());
        }

        private void FormatNowTracking(string historyText)
        {
            tabDKCH_nowTracking.Clear();
            tabDKCH_nowTracking.Text = historyText;
            var formatRules = new Dictionary<string, System.Drawing.Color>
            {
                { "Đăng ký chuyển hoàn", System.Drawing.Color.Red },
                { "Đăng ký chuyển hoàn lần 2", System.Drawing.Color.Red },
                { "Quét kiện vấn đề", System.Drawing.Color.Red },
                { "Giao lại hàng", System.Drawing.Color.DodgerBlue },
                { "Ký nhận CPN", System.Drawing.Color.ForestGreen },
                { "Đang chuyển hoàn", System.Drawing.Color.DarkOrange },
                { "Xác nhận chuyển hoàn", System.Drawing.Color.DarkOrange }
            };

            foreach (var rule in formatRules)
            {
                int startIndex = 0;
                while (startIndex < tabDKCH_nowTracking.TextLength)
                {
                    int wordStartIndex = tabDKCH_nowTracking.Find(rule.Key, startIndex, RichTextBoxFinds.None);
                    if (wordStartIndex != -1)
                    {
                        tabDKCH_nowTracking.SelectionStart = wordStartIndex;
                        tabDKCH_nowTracking.SelectionLength = rule.Key.Length;
                        tabDKCH_nowTracking.SelectionColor = rule.Value;
                        tabDKCH_nowTracking.SelectionFont = new System.Drawing.Font(tabDKCH_nowTracking.Font, System.Drawing.FontStyle.Bold);
                        startIndex = wordStartIndex + rule.Key.Length;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            tabDKCH_nowTracking.SelectionStart = tabDKCH_nowTracking.TextLength;
            tabDKCH_nowTracking.SelectionLength = 0;
            tabDKCH_nowTracking.SelectionColor = tabDKCH_nowTracking.ForeColor;
            tabDKCH_nowTracking.SelectionFont = tabDKCH_nowTracking.Font;
        }

        private void ApplyZoomFactor()
        {
            if (tabDKCH_webView?.CoreWebView2 != null) tabDKCH_webView.ZoomFactor = _settings.ZoomFactor;
        }


        // ======================================================================================
        // TAB TRACKING & UPLOAD BUMP
        // ======================================================================================

        private async void btnSearch_Click(object sender, EventArgs e)
        {
            string input = NormalizeWaybillInput(tabTracking_inputWaybill.Text);
            tabTracking_inputWaybill.Text = input;

            if (string.IsNullOrWhiteSpace(input))
            {
                UIMessageTip.ShowWarning("Chưa nhập mã vận đơn!");
                return;
            }

            try
            {
                tabTracking_btnSearch.Enabled = false;
                if (tabTracking_process != null)
                {
                    tabTracking_process.Value = 0;
                    tabTracking_process.Visible = true;
                }

                await _trackingService.SearchTrackingAsync(input);
                int tongSoDon = _trackingService.GetAllRows().Count;
                tabTracking_countSum.Text = tongSoDon.ToString("N0");

                if (tongSoDon == 0)
                {
                    UIMessageTip.ShowWarning("Không tìm thấy vận đơn nào!");
                }
            }
            catch (Exception ex)
            {
                UIMessageTip.ShowError("Lỗi khi tra cứu: " + ex.Message);
            }
            finally
            {
                tabTracking_btnSearch.Enabled = true;
                UpdateWaybillCount();
                if (tabTracking_process != null)
                {
                    tabTracking_process.Value = tabTracking_process.Maximum;
                    tabTracking_process.Visible = false;
                }
            }
        }

        private void btn_Export_Click(object sender, EventArgs e) => _trackingService?.ExportToExcel();

        private void btn_Clear_Click(object sender, EventArgs e)
        {
            _trackingService?.ClearData();
            tabTracking_inputWaybill.Clear();
        }

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
                    GoogleSheetService.ClearSheetAsync(spreadsheetId, targetSheetName);
                    GoogleSheetService.UpdateBumpSheetAsync(sheetData, spreadsheetId, $"{targetSheetName}!A1");
                });
                MessageBox.Show("Đã tải lên thành công!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Lỗi: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { tabTracking_btnUpload.Enabled = true; tabTracking_btnUpload.Text = oldText; }
        }

        private void tabTracking_inputWaybill_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V)
            {
                e.SuppressKeyPress = true;
                if (Clipboard.ContainsText())
                {
                    string plainText = Clipboard.GetText(TextDataFormat.UnicodeText);
                    string cleaned = NormalizeWaybillInput(plainText);
                    if (!string.IsNullOrEmpty(cleaned))
                    {
                        tabTracking_inputWaybill.SelectedText = cleaned + Environment.NewLine;
                    }
                }
            }
        }

        private string NormalizeWaybillInput(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var parts = text.Split(new[] { '\r', '\n', ',', ';', '|', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(t => t.Trim().ToUpper())
                            .Where(t => t.Length >= 6)
                            .ToList();
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var waybillRegex = new System.Text.RegularExpressions.Regex(@"((8\d{11}|[A-Za-z][A-Za-z0-9]{4,17})(-\d{3})?)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            foreach (var part in parts)
            {
                var matches = waybillRegex.Matches(part);
                if (matches.Count > 0)
                {
                    foreach (System.Text.RegularExpressions.Match m in matches)
                    {
                        string code = m.Value.ToUpper();
                        if (seen.Add(code)) result.Add(code);
                    }
                }
                else if (part.Length <= 25)
                {
                    if (seen.Add(part)) result.Add(part);
                }
            }
            return string.Join(Environment.NewLine, result);
        }

        private void UpdateWaybillCount()
        {
            if (tabDKCH_countSum == null) return;
            var uniqueCodes = tabTracking_inputWaybill.Text.Split(new[] { '\r', '\n', ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().ToUpper()).Where(x => x.Length > 5).Distinct(StringComparer.OrdinalIgnoreCase);
            tabDKCH_countSum.Text = $":Tổng: " + uniqueCodes.Count().ToString("N0");
        }


        // ======================================================================================
        // TAB PRINT
        // ======================================================================================

        private void print_LamMoi_Click(object sender, EventArgs e)
        {
            _trackingService.ClearData();
            _printService.Reset();
            tabPrint_btnSelectAll.Checked = false;
            tabPrint_countSelect.Text = "Đang chọn: 0";
            tabPrint_countSum.Text = "Tổng: 0";
            if (tabPrint_inputWaybill != null)
            {
                tabPrint_inputWaybill.Text = "";
            }
            if (tabPrint_printPreview?.CoreWebView2 != null)
            {
                tabPrint_printPreview.CoreWebView2.Navigate("about:blank");
            }
        }

        private void TabPrint_printFunc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_printService == null) return;

            if (tabPrint_printFunc.SelectedTab == tabPrint_inCH) _printService.SetMode(PrintMode.InHoan);
            else if (tabPrint_printFunc.SelectedTab == tabPrint_inCT) _printService.SetMode(PrintMode.InChuyenTiep);
            else if (tabPrint_printFunc.SelectedTab == tabPrint_inLaiDon) _printService.SetMode(PrintMode.InLaiDon);
            else if (tabPrint_printFunc.SelectedTab == tabPrint_inRV) _printService.SetMode(PrintMode.InReverse);

            tabPrint_btnSelectAll.Checked = false;
            tabPrint_inputWaybill.Text = "";
        }

        private void print_InChuyenHoan_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InHoan);
        private void print_InChuyenTiep_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InChuyenTiep);
        private void print_InLaiDon_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InLaiDon);
        private void print_InReverse_Click(object sender, EventArgs e) => _printService.SetMode(PrintMode.InReverse);

        private string GetBaseWaybill(string waybill)
        {
            if (string.IsNullOrEmpty(waybill)) return "";
            int hyphenIndex = waybill.IndexOf('-');
            return hyphenIndex > 0 ? waybill.Substring(0, hyphenIndex).Trim().ToUpper() : waybill.Trim().ToUpper();
        }

        private async void tabPrint_inputWaybill_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                if (tabPrint_AutoMode.Active)
                {
                    string input = tabPrint_inputWaybill.Text.Trim();
                    if (string.IsNullOrWhiteSpace(input)) return;

                    await _printService.SearchAndLoadAsync(input, _printService.CurrentMode);
                    _printService.SelectAll(true);
                    tabPrint_btnSelectAll.Checked = true;
                    await ExecutePrintAsync(true);
                    tabPrint_inputWaybill.Text = "";
                }
                else
                {
                    tabPrint_inputWaybill.AppendText(Environment.NewLine);
                    print_TimKiem_Click(null, null);
                }
            }
        }

        private async void print_TimKiem_Click(object sender, EventArgs e)
        {
            string input = tabPrint_inputWaybill.Text.Trim();
            if (string.IsNullOrWhiteSpace(input)) return;

            tabPrint_btnTimKiem.Enabled = false;
            try
            {
                await _printService.SearchAndLoadAsync(input, _printService.CurrentMode);
                _printService.SelectAll(true);
                tabPrint_btnSelectAll.Checked = true;
            }
            finally
            {
                tabPrint_btnTimKiem.Enabled = true;
            }
        }

        private async void tabPrint_btnPrint_Click(object sender, EventArgs e)
        {
            await ExecutePrintAsync(isAutoMode: false);
        }

        private async Task ExecutePrintAsync(bool isAutoMode)
        {
            if (_printService == null) return;
            if (!await _printLock.WaitAsync(0))
            {
                if (!isAutoMode) ShowPrintMessage("Đang in...", true);
                return;
            }
            bool printButtonChanged = false;
            try
            {
                var selected = _printService.GetSelectedWaybills();
                if (selected == null || selected.Count == 0)
                {
                    if (!isAutoMode) ShowPrintMessage("Chưa chọn vận đơn nào!", true);
                    return;
                }
                List<string> originalOrder = ParseWaybillOrder(tabPrint_inputWaybill.Text);
                if (originalOrder.Count > 0)
                {
                    var orderMap = originalOrder.Select((wb, index) => new { wb, index }).GroupBy(x => x.wb, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First().index, StringComparer.OrdinalIgnoreCase);
                    selected = selected.OrderBy(wb => orderMap.TryGetValue(wb, out int idx) ? idx : int.MaxValue).ToList();
                }
                SetPrintButtonState(false);
                printButtonChanged = true;
                int printType = 1;
                int applyTypeCode = (_printService.CurrentMode == PrintMode.InChuyenTiep) ? 2 : 4;
                string pdfUrl = await GetPdfUrlViaCSharpAsync(selected, printType, applyTypeCode);
                if (string.IsNullOrWhiteSpace(pdfUrl))
                {
                    ShowPrintMessage("Không thể in.", true);
                    return;
                }
                int keepPdfs = 500;
                int keepLogsDays = 3;
                TryReadPrintConfig(out keepPdfs, out keepLogsDays);
                string firstWaybill = selected[0];
                string localPath = await DownloadPdfWithRetryAsync(pdfUrl, keepPdfs, firstWaybill);
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    ShowPrintMessage("In thất bại.", true);
                    return;
                }
                string baseInput = GetBaseWaybill(firstWaybill);
                int currentPrintCount = 0;
                if (_printedHistory.TryGetValue(baseInput, out var history))
                {
                    currentPrintCount = history.Count;
                    if (currentPrintCount >= 3)
                    {
                        ShowPrintMessage($"{firstWaybill} đã in {currentPrintCount} lần!", true, 5000);
                        return;
                    }
                    ShowPrintMessage($"Đã in {currentPrintCount} lần. Gần nhất {history.LastTime:HH:mm:ss}", true, 4000);
                }
                if (tabPrint_printPreview?.CoreWebView2 != null)
                {
                    string fileUri = new Uri(localPath).AbsoluteUri;
                    tabPrint_printPreview.CoreWebView2.Navigate($"{fileUri}#view=FitH&toolbar=0&navpanes=0&scrollbar=0");
                }
                SavePrintLog(selected, keepLogsDays);
                try
                {
                    byte[] pdfBytes = await File.ReadAllBytesAsync(localPath);
                    using var stream = new MemoryStream(pdfBytes);
                    using var document = PdfDocument.Load(stream);
                    using var printDocument = document.CreatePrintDocument();
                    string printerName = string.IsNullOrWhiteSpace(_settings.PrinterName) ? "" : _settings.PrinterName;
                    if (!string.IsNullOrEmpty(printerName) && printerName != "-1")
                    {
                        printDocument.PrinterSettings.PrinterName = printerName;
                    }
                    printDocument.PrintController = new StandardPrintController();
                    printDocument.Print();
                    _printedHistory[baseInput] = (DateTime.Now, currentPrintCount + 1);
                }
                catch (Exception ex)
                {
                    ShowPrintMessage($"Lỗi máy in: {ex.Message}", true, 5000);
                }
            }
            catch (Exception ex)
            {
                ShowPrintMessage($"Lỗi hệ thống: {ex.Message}", true);
            }
            finally
            {
                tabPrint_btnPrint.Enabled = true;
                tabPrint_btnPrint.Text = "IN";
                _printLock.Release();

                _printService.ClearGridSelection();
                tabPrint_inputWaybill.Text = "";
            }
        }

        private async Task<string> GetPdfUrlViaCSharpAsync(List<string> waybills, int printType, int applyTypeCode)
        {
            try
            {
                await RefreshAuthTokenAsync();
                string token = Main.CapturedAuthToken ?? "";
                if (string.IsNullOrEmpty(token) || token.Length < 20)
                {
                    ShowPrintMessage("Không tìm thấy Token xác thực.", true);
                    return null;
                }
                using (System.Net.Http.HttpClient client = new System.Net.Http.HttpClient())
                {
                    client.DefaultRequestHeaders.Clear();
                    client.DefaultRequestHeaders.Add("authToken", token);
                    client.DefaultRequestHeaders.Add("lang", "VN");
                    client.DefaultRequestHeaders.Add("langType", "VN");
                    client.DefaultRequestHeaders.Add("User-Agent", CHROME_USER_AGENT);
                    client.DefaultRequestHeaders.Add("Accept", "application/json, text/plain, */*");
                    var payload = new Dictionary<string, object> { { "waybillIds", waybills }, { "applyTypeCode", applyTypeCode }, { "printType", printType }, { "pringType", printType }, { "countryId", "1" } };
                    string jsonPayload = System.Text.Json.JsonSerializer.Serialize(payload);
                    var content = new System.Net.Http.StringContent(jsonPayload, System.Text.Encoding.UTF8, "application/json");
                    string apiUrl = AppConfig.Current.BuildJmsApiUrl("operatingplatform/rebackTransferExpress/printWaybill");
                    var response = await client.PostAsync(apiUrl, content);
                    string rawJson = await response.Content.ReadAsStringAsync();
                    if (!response.IsSuccessStatusCode) return null;
                    using (System.Text.Json.JsonDocument doc = System.Text.Json.JsonDocument.Parse(rawJson))
                    {
                        var root = doc.RootElement;
                        if (root.TryGetProperty("code", out System.Text.Json.JsonElement codeElement))
                        {
                            string codeVal = codeElement.ToString();
                            if (codeVal == "200" || codeVal == "0" || codeVal == "1")
                            {
                                if (root.TryGetProperty("data", out System.Text.Json.JsonElement data))
                                {
                                    if (data.ValueKind == System.Text.Json.JsonValueKind.String)
                                    {
                                        string url = data.GetString();
                                        if (!string.IsNullOrEmpty(url) && url.StartsWith("http")) return url;
                                    }
                                    else if (data.ValueKind == System.Text.Json.JsonValueKind.Array && data.GetArrayLength() > 0)
                                    {
                                        string url = data[0].GetString();
                                        if (!string.IsNullOrEmpty(url) && url.StartsWith("http")) return url;
                                    }
                                }
                                ShowPrintMessage("Không có bản in.", true);
                            }
                            else
                            {
                                string msg = root.TryGetProperty("msg", out System.Text.Json.JsonElement msgElement) ? msgElement.GetString() : "Lỗi từ máy chủ JMS";
                                ShowPrintMessage(msg, true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi mạng: {ex.Message}");
            }
            return null;
        }

        private async Task<string> DownloadPdfWithRetryAsync(string pdfUrl, int keepPdfs, string waybillTag = "")
        {
            int maxRetries = 3;
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(pdfUrl)) return null;
                    string printFolder = Path.Combine(Application.StartupPath, "Downloads", "Vận đơn đã in");
                    if (!Directory.Exists(printFolder)) Directory.CreateDirectory(printFolder);
                    using (var client = new System.Net.Http.HttpClient())
                    {
                        client.Timeout = TimeSpan.FromMinutes(3);
                        client.DefaultRequestHeaders.Add("User-Agent", CHROME_USER_AGENT);
                        byte[] bytes = await client.GetByteArrayAsync(pdfUrl.Trim());
                        string fileName = string.IsNullOrEmpty(waybillTag) ? $"AutoJMS_{DateTime.Now:yyyyMMdd_HHmmssfff}.pdf" : $"{waybillTag.Replace("/", "_")}-{DateTime.Now:yyyyMMdd_HHmmssfff}.pdf";
                        string path = Path.Combine(printFolder, fileName);
                        File.WriteAllBytes(path, bytes);
                        var files = new DirectoryInfo(printFolder).GetFiles("*.pdf").OrderByDescending(f => f.CreationTime).ToList();
                        for (int j = keepPdfs; j < files.Count; j++)
                        {
                            try { files[j].Delete(); } catch { }
                        }
                        return path;
                    }
                }
                catch (Exception ex)
                {
                    if (i == maxRetries - 1)
                    {
                        ShowPrintMessage($"Không thể tải PDF: {ex.Message}", true);
                        return null;
                    }
                    await Task.Delay(2000);
                }
            }
            return null;
        }

        private void SavePrintLog(List<string> waybills, int keepLogsDays)
        {
            try
            {
                string logFolder = Path.Combine(Application.StartupPath, "Downloads", "Vận đơn đã in", "Logs");
                if (!Directory.Exists(logFolder)) Directory.CreateDirectory(logFolder);
                string logFile = Path.Combine(logFolder, $"Log_{DateTime.Now:yyyyMMdd}.txt");
                string timeStr = DateTime.Now.ToString("HH:mm:ss");
                string line = $"[{timeStr}] Đã in {waybills.Count} đơn: {string.Join(", ", waybills)}{Environment.NewLine}";
                File.AppendAllText(logFile, line);
                DateTime cutoff = DateTime.Now.Date.AddDays(-keepLogsDays);
                var oldLogs = new DirectoryInfo(logFolder).GetFiles("Log_*.txt").Where(f => f.CreationTime.Date < cutoff).ToList();
                foreach (var f in oldLogs)
                {
                    try { f.Delete(); } catch { }
                }
            }
            catch { }
        }

        private static List<string> ParseWaybillOrder(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return new List<string>();
            return input.Split(new[] { '\r', '\n', ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Select(x => x.Trim().ToUpperInvariant()).Where(x => x.Length > 5).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private void TryReadPrintConfig(out int keepPdfs, out int keepLogsDays)
        {
            keepPdfs = 500;
            keepLogsDays = 3;
            try
            {
                string jsonPath = Path.Combine(Application.StartupPath, "AutoJMS.json");
                if (!File.Exists(jsonPath)) return;
                using var doc = JsonDocument.Parse(File.ReadAllText(jsonPath));
                var root = doc.RootElement;
                if (root.TryGetProperty("KeepRecentPdfCount", out var pPdf) && pPdf.TryGetInt32(out int vPdf) && vPdf > 0)
                {
                    keepPdfs = vPdf;
                }
                if (root.TryGetProperty("PrintLogRetentionDays", out var pLog) && pLog.TryGetInt32(out int vLog) && vLog > 0)
                {
                    keepLogsDays = vLog;
                }
            }
            catch { }
        }

        private void SetPrintButtonState(bool enabled)
        {
            if (tabPrint_btnPrint.InvokeRequired)
            {
                tabPrint_btnPrint.Invoke(new Action(() => SetPrintButtonState(enabled)));
                return;
            }
            tabPrint_btnPrint.Enabled = enabled;
            tabPrint_btnPrint.Text = enabled ? "IN" : "Đang in...";
        }

        private void tabPrint_btnSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (_printService != null)
            {
                _printService.SelectAll(tabPrint_btnSelectAll.Checked);
            }
        }

        private void ShowPrintMessage(string message, bool isError = false, int timeout = 2000)
        {
            if (tabPrint_messLable == null || tabPrint_messLable.IsDisposed) return;
            if (hideTimer != null)
            {
                hideTimer.Stop();
                hideTimer.Dispose();
                hideTimer = null;
            }
            tabPrint_messLable.Text = message;
            tabPrint_messLable.ForeColor = isError ? Color.Red : Color.Black;
            hideTimer = new System.Windows.Forms.Timer();
            hideTimer.Interval = timeout;
            hideTimer.Tick += (sender, e) =>
            {
                if (tabPrint_messLable != null && !tabPrint_messLable.IsDisposed) tabPrint_messLable.Text = "";
                hideTimer.Stop();
                hideTimer.Dispose();
                hideTimer = null;
            };
            hideTimer.Start();
        }
    }
}
