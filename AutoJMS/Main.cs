using AutoJMS.Data;
using AutoUpdaterDotNET;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace AutoJMS
{
    public partial class Main : Form
    {
        private const string LICENSE_SPREADSHEET_ID = "1nx2VoXnAU3h8GRPxXkZ4c9Ev8jwHg3Iyor5o6wvsLNY";
        private const string LICENSE_SHEET_NAME = "LICENSE";
        string appsScriptUrl = "https://script.google.com/macros/s/AKfycbzVoJuFl-jf5wWXKEkRCLUXP8_cQ8yYWdDojFo-V19FH1G6Lm0AxhEyjZ9XI85pXIKGTQ/exec";

        private AppSettings _settings;
        public static string CapturedAuthToken = "";
        private WaybillTrackingService _trackingService;
        private FocusModeHelper _focusMode;
        private DkchManager _dkchManager;
        private PrintService _printService;
        private ZaloChatService _zaloChatService;
        private System.Windows.Forms.Timer _queueTimer;
        private BindingSource _zaloBindingSource = new BindingSource();
        private int _originalPnlLeftWidth = 260;
        private bool isZaloBotRunning = false;
        private System.Windows.Forms.Timer timer_AutoUpdateStatus;


        public Main()
        {
            InitializeComponent();
            _settings = SettingsManager.Load();
            Version appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            string versionText = $"Phiên bản: v{appVersion}";
            lbl_version.Text = versionText;
            lbl_version.ForeColor = System.Drawing.Color.Red;
            lbl_version.Visible = true;
            lbl_version.BringToFront();

            cb_SheetName.SelectedItem = _settings.DefaultSheet;
            ckb_UseSheet.Checked = _settings.UseSheetByDefault;
            _originalPnlLeftWidth = pnl_Left.Width;
            num_Row.Value = _settings.DefaultRowCount;
            Data_Waybill.KeyDown += Data_Waybill_KeyDown;
            print_SelectAll.CheckedChanged += print_SelectAll_CheckedChanged;
            _queueTimer = new System.Windows.Forms.Timer();
            _queueTimer.Interval = 30000; // 30 giây check 1 lần
            _queueTimer.Tick += async (s, e) => await ProcessDirectTrackingAsync();
            _queueTimer.Start();
            InitStatusTimer();


            btn_DKCH1.Visible = true;
            btn_DKCH2.Visible = true;
            btn_Stop.Visible = false;
            btn_DKCH1.Enabled = false;
            btn_DKCH2.Enabled = false;


            pnl_InChuyenHoan.Visible = false;
            pnl_ZaloChatz.Visible = false;


            if (pnl_TraHanhTrinh != null)
            {
                pnl_TraHanhTrinh.Parent = this;
                pnl_TraHanhTrinh.Dock = DockStyle.Fill;
                pnl_TraHanhTrinh.Visible = false;
                pnl_TraHanhTrinh.BringToFront();
            }

            txt_InputNew.KeyDown += txt_InputNew_KeyDown;
            CheckForIllegalCrossThreadCalls = false;

            _focusMode = new FocusModeHelper(pnl_Left, btn_FocusMode);

            this.KeyPreview = true;
            _dkchManager = new DkchManager();
            _dkchManager.OnSaveCountChanged += (count) => lbl_CountSave.Text = count.ToString();
            _dkchManager.OnTrackingHistoryChanged += (history) => now_Tracking.Text = history;
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F11 && Tab_Control.SelectedTab != Tab_NangCao)
            {
                _focusMode.Toggle();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void btn_FocusMode_Click(object sender, EventArgs e)
        {
            if (_focusMode != null)
            {
                _focusMode.Toggle();
            }
        }
        private async Task ProcessDirectTrackingAsync()
        {
            _queueTimer.Stop();
            try
            {
                string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;

                // 1. Kiểm tra lệnh RUN ở C5
                var c5Data = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!C5");
                string c5Value = (c5Data != null && c5Data.Count > 0 && c5Data[0].Count > 0)
                    ? c5Data[0][0].ToString().Trim().ToUpper()
                    : "";

                if (c5Value == "RUN")
                {
                    GoogleSheetService.UpdateCell(spreadsheetId, "PHATLAI!C5", "PROCESSING");

                    // Đọc danh sách từ cột A
                    var columnAData = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!A2:A");
                    List<string> waybills = new List<string>();

                    if (columnAData != null)
                    {
                        foreach (var row in columnAData)
                        {
                            if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0]?.ToString()))
                                waybills.Add(row[0].ToString().Trim());
                        }
                    }

                    if (waybills.Count > 0)
                    {
                        string waybillsText = string.Join("\n", waybills);

                        // Tracking
                        await _trackingService.SearchTrackingAsync(waybillsText, false);
                        var results = _trackingService.GetAllRows();

                        // Ghi kết quả vào BUMP
                        var sheetData = new List<IList<object>>();
                        foreach (var item in results)
                        {
                            sheetData.Add(new List<object>()
                    {
                        item.WaybillNo,
                        item.TrangThaiHienTai,
                        item.ThaoTacCuoi,
                        item.ThoiGianThaoTac,
                        item.ThoiGianYeuCauPhatLai,
                        item.NhanVienKienVanDe
                    });
                        }
                        GoogleSheetService.UpdateBumpSheet(sheetData, spreadsheetId, "BUMP!A2");

                        // === PHẦN MỚI: COPY danh sách vừa tracking vào cột B2:B ===
                        var bValues = new List<IList<object>>();
                        foreach (var wb in waybills)
                        {
                            bValues.Add(new List<object> { wb });
                        }
                        GoogleSheetService.UpdateBumpSheet(bValues, spreadsheetId, "PHATLAI!B2");

                    }

                    GoogleSheetService.UpdateCell(spreadsheetId, "PHATLAI!C5", "DONE");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Lỗi Tracking Direct] {ex.Message}");
            }
            finally
            {
                _queueTimer.Start();
            }
        }
        private void ApplyZoomFactor()
        {
            if (Main_Webview?.CoreWebView2 != null)
                Main_Webview.ZoomFactor = _settings.ZoomFactor;
        }

        private void Data_Waybill_KeyDown(object sender, KeyEventArgs e)
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
                    Data_Waybill.SelectedText = cleaned;
                }
            }
        }

        private void UpdateWaybillCount()
        {
            if (inputCount == null) return;
            var uniqueCodes = Data_Waybill.Text
                .Split(new[] { '\r', '\n', ' ', ',', ';', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim().ToUpper())
                .Where(x => x.Length > 5)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            inputCount.Text = uniqueCodes.Count().ToString("N0");
        }

        private bool _isZaloLoaded = false;
        private bool _zaloServiceInitialized = false;

        private async void Tab_Control_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 1. Tạm dừng vẽ giao diện (chỉ áp dụng cho phần đổi UI đồng bộ)
            this.SuspendLayout();

            try
            {
                // Ẩn tất cả trước khi mở tab mới để tránh chồng hình
                pnl_Webview.Visible = false;
                if (pnl_TraHanhTrinh != null) pnl_TraHanhTrinh.Visible = false;
                if (pnl_ZaloChatz != null) pnl_ZaloChatz.Visible = false;

                // Trả pnl_Left về kích thước mặc định (Sẽ được kéo giãn lại nếu là tab Zalo)
                if (Tab_Control.SelectedTab != Tab_ZaloChat)
                {
                    pnl_Left.Width = _originalPnlLeftWidth;
                }

                // --- XỬ LÝ GIAO DIỆN ---
                if (Tab_Control.SelectedTab == Tab_DKCH)
                {
                    pnl_Webview.Visible = true;
                    pnl_Webview.BringToFront();
                    if (btn_FocusMode != null)
                    {
                        if (btn_FocusMode.Parent != panel_Main_act) btn_FocusMode.Parent = panel_Main_act;
                        btn_FocusMode.Visible = true;
                        btn_FocusMode.BringToFront();
                    }
                }
                else if (Tab_Control.SelectedTab == Tab_NangCao)
                {
                    if (btn_FocusMode != null) btn_FocusMode.Visible = false;

                    if (pnl_TraHanhTrinh != null)
                    {
                        if (pnl_TraHanhTrinh.Parent != panel_Main_act) pnl_TraHanhTrinh.Parent = panel_Main_act;
                        pnl_TraHanhTrinh.Visible = true;
                        pnl_TraHanhTrinh.BringToFront();
                    }
                }
                else if (Tab_Control.SelectedTab == Tab_ZaloChat)
                {
                    if (pnl_ZaloChatz != null)
                    {
                        pnl_Left.Width = 400;
                        if (pnl_ZaloChatz.Parent != panel_Main_act) pnl_ZaloChatz.Parent = panel_Main_act;
                        pnl_ZaloChatz.Visible = true;
                        pnl_ZaloChatz.BringToFront();

                        if (btn_FocusMode != null)
                        {
                            if (btn_FocusMode.Parent != panel_Main_act) btn_FocusMode.Parent = panel_Main_act;
                            btn_FocusMode.Visible = true;
                            btn_FocusMode.BringToFront();
                        }
                    }

                    // === LUÔN SETUP GRID NGAY KHI MỞ TAB (để header hiện liền) ===
                    if (!_isZaloLoaded)
                    {
                        zaloChat_DataView.AutoGenerateColumns = false;
                        UniformHeaderColor();
                        EnableDoubleBufferedForGunaGrid();
                        try
                        {
                            await Main_ZaloChat.EnsureCoreWebView2Async();
                            Main_ZaloChat.CoreWebView2.Navigate("https://chat.zalo.me/index.html");
                            Main_ZaloChat.NavigationCompleted += ZaloNavigationCompletedHandler;
                            _isZaloLoaded = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Lỗi khởi tạo Zalo Web: " + ex.Message);
                        }
                    }
                    else if (_zaloServiceInitialized && _zaloChatService != null)
                    {
                        await LoadZaloData();
                    }
                }
            }
            finally
            {
                // 2. BẮT BUỘC mở lại vẽ UI TRƯỚC KHI gọi bất kỳ lệnh await nào
                this.ResumeLayout(true);
            }

            // --- XỬ LÝ DỮ LIỆU & LOAD NGẦM BẤT ĐỒNG BỘ ---
            if (Tab_Control.SelectedTab == Tab_NangCao)
            {
                // Chỉ thực hiện RefreshAuthTokenAsync() khi chưa có authtoken
                if (string.IsNullOrEmpty(Main.CapturedAuthToken))
                {
                    await RefreshAuthTokenAsync();
                }
            }
            else if (Tab_Control.SelectedTab == Tab_ZaloChat)
            {
                if (!_isZaloLoaded && Main_ZaloChat.CoreWebView2 == null)
                {
                    try
                    {
                        await Main_ZaloChat.EnsureCoreWebView2Async();
                        Main_ZaloChat.CoreWebView2.Navigate("https://chat.zalo.me/index.html");

                        _isZaloLoaded = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Lỗi khởi tạo Zalo Web: " + ex.Message);
                    }
                }
            }
        }
        private async void ZaloNavigationCompletedHandler(object sender, CoreWebView2NavigationCompletedEventArgs args)
        {
            if (_zaloChatService == null)
            {
                _zaloChatService = new ZaloChatService(Main_ZaloChat, appsScriptUrl);
                // KHÔNG auto start reminder ở đây (để user bấm nút Start)
            }

            if (!_zaloServiceInitialized)
            {
                await LoadZaloData();           // Load dữ liệu lần đầu sau khi service sẵn sàng
                _zaloServiceInitialized = true;
            }
        }
        private async void Main_Load(object sender, EventArgs e)
        {
            Version appVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            lbl_version.Text = $"Phiên bản: v{appVersion}";
            await Main_Webview.EnsureCoreWebView2Async(null);
            await print_Preview.EnsureCoreWebView2Async(); string
            userDataFolder = Path.Combine(Application.StartupPath, "ZaloProfile");
            var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
            await Main_ZaloChat.EnsureCoreWebView2Async(env);
            Main_Webview.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Fetch);
            Main_Webview.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.XmlHttpRequest);
            Main_Webview.CoreWebView2.WebResourceRequested += CoreWebView2_WebResourceRequested;
            Main_Webview.CoreWebView2.NavigationCompleted += (s, args) =>
            {
                if (args.IsSuccess) ApplyZoomFactor();
            };

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
            if (Main_Webview.CoreWebView2 == null) return;
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

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            await WebViewHost.InitAsync(Main_Webview);
            ApplyZoomFactor();
            Main_Webview.ZoomFactorChanged += (s, args) =>
            {
                _settings.ZoomFactor = Main_Webview.ZoomFactor;
                SettingsManager.Save(_settings);
            };
            await WebViewHost.NavigateAsync(_settings.DefaultUrl);
            await Task.Delay(2000);
            await RefreshAuthTokenAsync();

            _trackingService = new WaybillTrackingService(Data_Result, progressBarTracking, lblPercent);
            WaybillTrackingService.EnableDoubleBuffering(Data_Result);
            Data_Result.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _printService = new PrintService(print_OptionWaybill, _trackingService);

            _dkchManager.SetWebView(Main_Webview);
            _dkchManager.SetTrackingService(_trackingService);
            _dkchManager.SetSettingsGetter(() => (
                useSheet: ckb_UseSheet.Checked,
                sheetName: cb_SheetName.Text,
                rowCount: (int)num_Row.Value
            ));
            _dkchManager.SetUILogger(txt_Preview, lbl_Count);
            _dkchManager.StartDaemon();
            btn_DKCH1.Enabled = true;
            btn_DKCH2.Enabled = true;
        }

        private void txt_InputNew_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                var allLines = txt_InputNew.Text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                if (allLines.Count == 0 && txt_InputNew.Text.Trim().Length > 0)
                    allLines.Add(txt_InputNew.Text.Trim());
                else if (allLines.Count == 0) return;

                string newWaybill = allLines.Last().Trim();
                if (!string.IsNullOrEmpty(newWaybill))
                {
                    _dkchManager.AddPriorityWaybill(newWaybill);
                    if (allLines.Count >= 5)
                    {
                        var keepLines = allLines.Skip(allLines.Count - 5).ToList();
                        txt_InputNew.Text = string.Join(Environment.NewLine, keepLines);
                    }
                    else
                    {
                        txt_InputNew.Text = string.Join(Environment.NewLine, allLines);
                    }
                    txt_InputNew.AppendText(Environment.NewLine);
                    txt_InputNew.SelectionStart = txt_InputNew.Text.Length;
                    txt_InputNew.ScrollToCaret();
                }
            }
        }

        private async void btn_DKCH1_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning) return;
            btn_DKCH1.Enabled = false;
            btn_DKCH2.Enabled = false;
            await _dkchManager.StartAsync("DKCH1");
            btn_DKCH1.Visible = false;
            btn_DKCH2.Visible = false;
            btn_Stop.Visible = true;
            await RefreshAuthTokenAsync();
        }

        private async void btn_DKCH2_Click(object sender, EventArgs e)
        {
            if (_dkchManager.IsRunning) return;
            btn_DKCH1.Enabled = false;
            btn_DKCH2.Enabled = false;
            await _dkchManager.StartAsync("DKCH2");
            btn_DKCH1.Visible = false;
            btn_DKCH2.Visible = false;
            btn_Stop.Visible = true;
        }

        private void btn_Stop_Click(object sender, EventArgs e)
        {
            _dkchManager.Stop();
            btn_DKCH1.Visible = true;
            btn_DKCH2.Visible = true;
            btn_DKCH1.Enabled = true;
            btn_DKCH2.Enabled = true;
            btn_Stop.Visible = false;
        }

        private async void btn_Refresh_Click(object sender, EventArgs e)
        {
            _dkchManager.Stop();
            await WebViewHost.NavigateAsync("https://jms.jtexpress.vn");
        }


        private async void btn_TimKiem_Click(object sender, EventArgs e)
        {
            string input = Data_Waybill.Text.Trim();
            UpdateWaybillCount();
            if (string.IsNullOrWhiteSpace(input))
            {
                MessageBox.Show("Nhập mã vận đơn vào ô bên trên!", "Thông báo");
                return;
            }
            await _trackingService.SearchTrackingAsync(input);
        }

        private void btn_Export_Click(object sender, EventArgs e) => _trackingService?.ExportToExcel();
        private void btn_Clear_Click(object sender, EventArgs e) { _trackingService?.ClearData(); Data_Waybill.Clear(); }
        private async void btn_Upload_Click(object sender, EventArgs e)
        {
            btn_Upload.Enabled = false;
            string oldText = btn_Upload.Text;
            btn_Upload.Text = "Đang đồng bộ...";

            try
            {
                DataGridView grid = Data_Result;

                if (grid.Rows.Count == 0 && grid.Columns.Count == 0)
                {
                    MessageBox.Show("Không có dữ liệu trên bảng để tải lên!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sheetData = new List<IList<object>>();

                // 1. LẤY TIÊU ĐỀ CỘT (HEADER)
                var headerRow = new List<object>();
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    headerRow.Add(col.HeaderText);
                }
                sheetData.Add(headerRow);

                // 2. LẤY DỮ LIỆU TỪNG DÒNG
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue; // Bỏ qua dòng trống cuối cùng để nhập liệu (nếu có)

                    var rowData = new List<object>();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        rowData.Add(cell.Value?.ToString() ?? "");
                    }
                    sheetData.Add(rowData);
                }

                // 3. ĐẨY LÊN GOOGLE SHEETS (Chạy ngầm)
                await Task.Run(() =>
                {
                    string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;

                    // Tên Sheet bạn muốn thao tác. 
                    // Nếu bạn muốn lấy theo Dropdown trên giao diện thì dùng: cb_SheetName.Text
                    string targetSheetName = "BUMP";

                    // Bước 3.1: XÓA TRẮNG DỮ LIỆU CŨ (Chỉ truyền tên Sheet thì nó sẽ xóa toàn bộ sheet đó)
                    GoogleSheetService.ClearSheet(spreadsheetId, targetSheetName);

                    // Bước 3.2: GHI ĐÈ DỮ LIỆU MỚI (Bắt đầu từ ô A1 vì đã có chứa Header)
                    GoogleSheetService.UpdateBumpSheet(sheetData, spreadsheetId, $"{targetSheetName}!A1");
                });


                MessageBox.Show("Đã tải lên thành công!", "Hoàn tất", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Đã xảy ra lỗi khi tải dữ liệu lên:\n{ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Console.WriteLine($"[LỖI UPLOAD] {ex.Message}");
            }
            finally
            {
                btn_Upload.Enabled = true;
                btn_Upload.Text = oldText;
            }
        }
        private void btn_Export_Spe_Click(object sender, EventArgs e) => _trackingService.ExportSpecial();

        private string _downloadFolderPath = Path.Combine(Application.StartupPath, "Downloads");
        private void btn_Download_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Directory.Exists(_downloadFolderPath)) Directory.CreateDirectory(_downloadFolderPath);
                Process.Start(new ProcessStartInfo
                {
                    FileName = _downloadFolderPath,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể mở thư mục: " + ex.Message);
            }
        }

        private void btn_TraHanhTrinh_Click(object sender, EventArgs e)
        {
            pnl_TraHanhTrinh.Visible = true;
            pnl_InChuyenHoan.Visible = false;
            pnl_ZaloChatz.Visible = false;
        }
        private void btn_NangCao_In_Click(object sender, EventArgs e)
        {
            pnl_TraHanhTrinh.Visible = false;
            pnl_InChuyenHoan.Visible = true;
            pnl_ZaloChatz.Visible = false;
        }
        private async void print_TimKiem_Click(object sender, EventArgs e)
        {
            string input = print_Input.Text.Trim();
            await _printService.SearchAndLoadAsync(input, _printService.CurrentMode);
        }
        private void print_LamMoi_Click(object sender, EventArgs e)
        {
            _trackingService.ClearData();
            _printService.Reset();
            print_Input.Clear();
            print_SelectAll.Checked = false;
            label10.Text = "000";
            label9.Text = "000";

            if (print_Preview != null && print_Preview.CoreWebView2 != null)
            {
                print_Preview.CoreWebView2.Navigate("about:blank");
            }

        }

        private void print_InChuyenHoan_Click(object sender, EventArgs e)
        {
            _printService.SetMode(PrintMode.InHoan);
        }
        private void print_InChuyenTiep_Click(object sender, EventArgs e)
        {
            _printService.SetMode(PrintMode.InChuyenTiep);
        }
        private void print_InLaiDon_Click(object sender, EventArgs e)
        {
            _printService.SetMode(PrintMode.InLaiDon);
        }
        private void print_InReverse_Click(object sender, EventArgs e)
        {
            _printService.SetMode(PrintMode.InReverse);
        }

        private void print_SelectAll_CheckedChanged(object sender, EventArgs e)
        {
            _printService.SelectAll(print_SelectAll.Checked);
        }





        private async Task<string> GetPdfUrlViaCSharpAsync(List<string> waybills, int printType, int applyTypeCode)
        {
            try
            {
                await RefreshAuthTokenAsync();

                string token = Main.CapturedAuthToken ?? "";

                using (HttpClient client = new HttpClient())
                {
                    // === REPLICATE ĐÚNG 100% HEADERS TỪ BROWSER ===
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


                    // 1. Gọi pringListPage trước (đánh dấu hệ thống)
                    var payload1 = new
                    {
                        current = 1,
                        size = 20,
                        pringFlag = 0,
                        applyNetworkId = 5165,
                        waybillIds = waybills,
                        applyTimeFrom = "2020-01-01 00:00:00",
                        applyTimeTo = "2030-01-01 23:59:59",
                        pringType = printType,
                        countryId = "1"
                    };

                    var content1 = new StringContent(JsonSerializer.Serialize(payload1), Encoding.UTF8, "application/json");
                    await client.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/pringListPage", content1);

                    // 2. Gọi printWaybill để lấy link PDF
                    var payload2 = new
                    {
                        waybillIds = waybills,
                        applyTypeCode = applyTypeCode,
                        countryId = "1"
                    };

                    var content2 = new StringContent(JsonSerializer.Serialize(payload2), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/printWaybill", content2);

                    string rawJson = await response.Content.ReadAsStringAsync();

                    using (JsonDocument doc = JsonDocument.Parse(rawJson))
                    {
                        var root = doc.RootElement;
                        if (root.TryGetProperty("data", out JsonElement data) && data.ValueKind == JsonValueKind.String)
                        {
                            string url = data.GetString();
                            if (url.StartsWith("http"))
                            {
                                Clipboard.SetText(url);
                                return url;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            return null;
        }


        private async Task<string> DownloadPdfToTempFileAsync(string pdfUrl)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(pdfUrl) || !Uri.TryCreate(pdfUrl.Trim(), UriKind.Absolute, out _))
                {
                    throw new Exception("Link PDF không hợp lệ (không phải URL tuyệt đối).");
                }

                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    byte[] bytes = await client.GetByteArrayAsync(pdfUrl.Trim());

                    string path = Path.Combine(Path.GetTempPath(), $"AutoJMS_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
                    File.WriteAllBytes(path, bytes);
                    return path;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private async void print_btn_Print_Click(object sender, EventArgs e)
        {
            // KIỂM TRA NULL TRƯỚC KHI CHẠY
            if (_printService == null)
            {
                MessageBox.Show("Chưa khởi tạo PrintService. Vui lòng load lại form.", "Lỗi");
                return;
            }

            var selected = _printService.GetSelectedWaybills();
            if (selected == null || selected.Count == 0)
            {
                MessageBox.Show("Chưa chọn vận đơn nào!", "Thông báo");
                return;
            }

            if (print_Preview?.CoreWebView2 == null)
            {
                await print_Preview.EnsureCoreWebView2Async(null);
            }

            var originalOrder = print_Input.Text
        .Split(new[] { '\r', '\n', ',', ';', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
        .Select(x => x.Trim().ToUpper())
        .Where(x => x.Length > 5)
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .ToList();

            // Sắp xếp lại selected theo thứ tự gốc
            selected = selected
                .OrderBy(wb => originalOrder.IndexOf(wb))
                .ToList();

            print_btn_Print.Enabled = false;
            print_btn_Print.Text = "Đang in...";


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
                        print_Preview.CoreWebView2.Navigate($"file:///{localPath.Replace("\\", "/")}");
                        for (int i = 0; i < selected.Count; i++)
                        {
                            var psi = new ProcessStartInfo(localPath) { Verb = "print", UseShellExecute = true };
                            Process.Start(psi);
                            await Task.Delay(800); // chờ máy in nhận lệnh
                        }
                    }
                }
                else
                {
                    MessageBox.Show("PDF", "Lỗi");
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                print_btn_Print.Enabled = true;
                print_btn_Print.Text = "IN";
            }
        }

        private void btn_CheckUpdate_Click(object sender, EventArgs e)
        {
            AutoUpdater.ShowSkipButton = false;
            AutoUpdater.ReportErrors = true;
            string xmlUrl = "https://raw.githubusercontent.com/thvandat03/AutoJMS/refs/heads/master/AutoJMS/update.xml";
            AutoUpdater.ExecutablePath = "AutoJMS.exe";
            AutoUpdater.Start(xmlUrl);
        }


        private void InitStatusTimer()
        {
            timer_AutoUpdateStatus = new System.Windows.Forms.Timer();
            timer_AutoUpdateStatus.Interval = 30 * 60 * 1000; // 30 phút
            timer_AutoUpdateStatus.Tick += async (s, e) => await ProcessPhatLaiReTrackAsync();
        }

        private void ZaloChat_btn_start_Click(object sender, EventArgs e)
        {
            if (_zaloChatService == null || !_zaloServiceInitialized)
            {
                MessageBox.Show("Zalo Chat chưa sẵn sàng.\nVui lòng chờ Zalo Web load xong.",
                                "Chưa sẵn sàng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!isZaloBotRunning)
            {
                // Kiểm tra thời gian nhắc nhở
                if (zaloChat_selectTimeRemind.SelectedItem == null ||
                    string.IsNullOrWhiteSpace(zaloChat_selectTimeRemind.Text))
                {
                    MessageBox.Show("Chưa chọn thời gian nhắc lại", "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string timeText = zaloChat_selectTimeRemind.Text.Trim();
                string numberStr = Regex.Match(timeText, @"\d+").Value;

                if (!int.TryParse(numberStr, out int timeValue) || timeValue <= 0)
                {
                    MessageBox.Show("Định dạng thời gian không hợp lệ!", "Lỗi",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int intervalSeconds = timeText.ToLower().Contains("giây") ? timeValue : timeValue * 60;

                // === BẮT ĐẦU CẢ HAI TIMER ===
                _zaloChatService.StartAutoReminder(intervalSeconds);
                timer_AutoUpdateStatus.Start();                    // ← Timer tracking cột B

                isZaloBotRunning = true;
                ZaloChat_btn_start.Text = "Dừng";
                Console.WriteLine("[ZaloBot] ✅ Đã khởi động cả 2 timer (Remind + Tracking cột B)");
            }
            else
            {
                // DỪNG CẢ HAI TIMER
                _zaloChatService.StopAutoReminder();
                timer_AutoUpdateStatus.Stop();

                isZaloBotRunning = false;
                ZaloChat_btn_start.Text = "Bắt đầu - Bot";
                Console.WriteLine("[ZaloBot] ⛔ Đã dừng cả 2 timer");
            }
        }
        // Khai báo biến toàn cục ở đầu file Main.cs
        private List<Reminder> _allZaloReminders = new List<Reminder>();

        // --- Hãy sửa lại hàm tải dữ liệu của bạn để gán biến này ---
        private void PopulateZaloStatusCombo()
        {
            if (zaloChat_selectStatus == null || _allZaloReminders == null) return;

            var unique = _allZaloReminders
                .Where(r => !string.IsNullOrWhiteSpace(r.trangThai))
                .Select(r => r.trangThai.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s)
                .ToList();

            zaloChat_selectStatus.Items.Clear();
            zaloChat_selectStatus.Items.Add("Tất cả");

            foreach (var s in unique)
                zaloChat_selectStatus.Items.Add(s);

            zaloChat_selectStatus.SelectedIndex = 0;
        }
        /// <summary>
        /// Buộc tất cả header cột cùng màu
        /// </summary>
        private void UniformHeaderColor()
        {
            var style = zaloChat_DataView.ColumnHeadersDefaultCellStyle;

            style.BackColor = Color.FromArgb(0, 122, 204); 
            style.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }
        private async Task LoadZaloData()
        {
            if (_zaloChatService == null)
            {
                Console.WriteLine("[Zalo] Service chưa khởi tạo.");
                return;
            }

            try
            {
                Console.WriteLine("[Zalo] Đang tải dữ liệu từ Sheet...");

                _allZaloReminders = await _zaloChatService.GetDataFromSheetAsync();

                Console.WriteLine($"[Zalo] ✅ Lấy được {_allZaloReminders.Count} vận đơn");

                if (_allZaloReminders.Count == 0)
                {
                    zaloChat_DataView.DataSource = null;
                    MessageBox.Show("Sheet hiện tại không có vận đơn nào!", "Thông báo");
                    return;
                }

                // Tự động tạo danh sách trạng thái cho ComboBox
                PopulateZaloStatusCombo();

                // Gán dữ liệu
                _zaloBindingSource.DataSource = _allZaloReminders;

                // Force refresh
                if (zaloChat_DataView.InvokeRequired)
                {
                    zaloChat_DataView.Invoke(new Action(() =>
                    {
                        zaloChat_DataView.Refresh();
                        zaloChat_DataView.Update();
                    }));
                }
                else
                {
                    zaloChat_DataView.Refresh();
                    zaloChat_DataView.Update();
                }

                MessageBox.Show($"ĐÃ LOAD THÀNH CÔNG {_allZaloReminders.Count} vận đơn vào grid!",
                               "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[LoadZaloData] Lỗi: {ex.Message}");
                MessageBox.Show($"Lỗi load: {ex.Message}");
            }
        }
        /// <summary>
        /// Bật DoubleBuffered cho Guna2DataGridView (tăng tốc vẽ, giảm flicker)
        /// </summary>
        private void EnableDoubleBufferedForGunaGrid()
        {
            try
            {
                // Reflection để bật DoubleBuffered (Guna2 không expose public)
                var dgvType = zaloChat_DataView.GetType();
                var prop = dgvType.GetProperty("DoubleBuffered",
                            System.Reflection.BindingFlags.NonPublic |
                            System.Reflection.BindingFlags.Instance);

                prop?.SetValue(zaloChat_DataView, true);

                Console.WriteLine("[Zalo Performance] ✅ DoubleBuffered đã bật cho zaloChat_DataView");
            }
            catch
            {
                // Fallback nếu reflection lỗi
                zaloChat_DataView.GetType().GetProperty("DoubleBuffered")?
                    .SetValue(zaloChat_DataView, true);
            }
        }
        private void FilterZaloGrid()
        {
            if (_allZaloReminders == null || zaloChat_DataView == null) return;

            string selected = zaloChat_selectStatus?.Text?.Trim() ?? "Tất cả";

            List<Reminder> listToShow = (string.IsNullOrWhiteSpace(selected) || selected == "Tất cả")
                ? _allZaloReminders
                : _allZaloReminders.Where(r =>
                    !string.IsNullOrEmpty(r.trangThai) &&
                    r.trangThai.IndexOf(selected, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            zaloChat_DataView.DataSource = null;        // Reset hoàn toàn
            zaloChat_DataView.DataSource = listToShow;

            zaloChat_DataView.Refresh();
            zaloChat_DataView.Update();

            Console.WriteLine($"[Zalo] Grid hiển thị {listToShow.Count} dòng (filter: {selected})");
        }
        private void zaloChat_selectStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Chỉ filter khi đã có dữ liệu
            if (_allZaloReminders != null && _allZaloReminders.Count > 0)
            {
                FilterZaloGrid();
            }
        }
        private async void zaloChat_btn_Refesh_Click(object sender, EventArgs e)
        {
            zaloChat_btn_Refesh.Enabled = false;
            zaloChat_btn_Refesh.Text = "Đang Tracking...";

            try
            {
                // Tracking NGAY từ cột B
                await ProcessPhatLaiReTrackAsync();

                // Load lại dữ liệu mới nhất vào Grid Zalo
                await LoadZaloData();

                Console.WriteLine("[Zalo] ✅ Làm mới hoàn tất: Đã tracking cột B và load grid");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi làm mới: {ex.Message}", "Lỗi",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                zaloChat_btn_Refesh.Enabled = true;
                zaloChat_btn_Refesh.Text = "Làm mới";
            }
        }
        private async Task ProcessPhatLaiReTrackAsync()
        {
            try
            {
                string spreadsheetId = GoogleSheetService.DATA_SPREADSHEET_ID;

                var bData = GoogleSheetService.ReadRange(spreadsheetId, "PHATLAI!B2:B");
                List<string> waybills = new List<string>();

                if (bData != null)
                {
                    foreach (var row in bData)
                    {
                        if (row.Count > 0 && !string.IsNullOrWhiteSpace(row[0]?.ToString()))
                            waybills.Add(row[0].ToString().Trim());
                    }
                }

                if (waybills.Count == 0)
                {
                    Console.WriteLine("[PHATLAI] Cột B trống → bỏ qua.");
                    return;
                }

                Console.WriteLine($"[PHATLAI] Tracking ngay {waybills.Count} đơn từ cột B...");

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

                Console.WriteLine($"[PHATLAI] Hoàn tất tracking cột B và update BUMP");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PHATLAI ReTrack] Lỗi: {ex.Message}");
            }
        }

















    }
}