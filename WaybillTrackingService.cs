using AutoJMS.Data;
using ClosedXML.Excel;
using Sunny.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    public class WaybillTrackingService
    {
        private readonly DataGridView _dataGrid;
        private readonly DataTable _displayTable;
        private readonly BindingSource _bindingSource = new BindingSource();
        private readonly string _exportFolder;
        private readonly HttpClient _httpClient;

        private readonly UIProcessBar _progressBar;
        private int _totalItems;
        private int _completedItems;
        private int _lastPercent = -1;

        private readonly List<TrackingRow> _allRows = new List<TrackingRow>();
        private readonly Dictionary<string, TrackingRow> _rowDict = new Dictionary<string, TrackingRow>(StringComparer.OrdinalIgnoreCase);
        
        // 1. CHỐNG RACE CONDITION (BẢO VỆ GHI DỮ LIỆU)
        private readonly object _rowLock = new object();
        
        // 2. CHỐNG LAG UI (CHỈ RESIZE GRID 1 LẦN)
        private bool _columnsResized = false;

        private string _currentSortColumn = string.Empty;
        private ListSortDirection _currentSortDirection = ListSortDirection.Ascending;

        // 3. CHUẨN HÓA BATCHING & CONCURRENCY
        private const int BatchSize = 25;
        private const int MaxConcurrency = 4;
        private const int MaxRetry = 3;

        public WaybillTrackingService(DataGridView dataGrid, UIProcessBar progressBar = null)
        {
            _dataGrid = dataGrid ?? throw new ArgumentNullException(nameof(dataGrid));
            _progressBar = progressBar;

            _exportFolder = Path.Combine(Application.StartupPath, "Downloads", "Trạng thái hiện tại");
            Directory.CreateDirectory(_exportFolder);

            _displayTable = CreateDataTable();

            // 4. TĂNG TỐC HTTP MẠNH MẼ VỚI SOCKETS HTTP HANDLER
            var handler = new SocketsHttpHandler
            {
                PooledConnectionLifetime = TimeSpan.FromMinutes(10),
                PooledConnectionIdleTimeout = TimeSpan.FromMinutes(5),
                MaxConnectionsPerServer = 50,
                EnableMultipleHttp2Connections = true
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(30)
            };

            SetupDataGridViewBinding();
            InitProgressControls();

            _bindingSource.DataSource = _displayTable;
            _dataGrid.DataSource = _bindingSource;
        }

        private void SetupDataGridViewBinding()
        {
            EnableDoubleBuffering(_dataGrid);
            _dataGrid.VirtualMode = false;
            _dataGrid.AutoGenerateColumns = false;
            _dataGrid.AllowUserToAddRows = false;
            _dataGrid.AllowUserToDeleteRows = false;
            _dataGrid.ReadOnly = true;
            _dataGrid.AllowUserToResizeRows = false;
            _dataGrid.RowHeadersVisible = false;
            _dataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dataGrid.MultiSelect = false;
            
            // Xóa tự động Resize theo từng Cell (Rất nặng)
            _dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None; 
            
            _dataGrid.ColumnHeaderMouseClick += DataGrid_ColumnHeaderMouseClick;
            _dataGrid.Columns.Clear();
            _dataGrid.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Mã vận đơn", HeaderText = "Mã vận đơn", DataPropertyName = "Mã vận đơn", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Trạng thái hiện tại", HeaderText = "Trạng thái hiện tại", DataPropertyName = "Trạng thái hiện tại", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thao tác cuối cùng", HeaderText = "Thao tác cuối cùng", DataPropertyName = "Thao tác cuối cùng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thời gian thao tác", HeaderText = "Thời gian thao tác", DataPropertyName = "Thời gian thao tác", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thời gian yêu cầu phát lại", HeaderText = "Thời gian yêu cầu phát lại", DataPropertyName = "Thời gian yêu cầu phát lại", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Nhân viên kiện vấn đề", HeaderText = "Nhân viên kiện vấn đề", DataPropertyName = "Nhân viên kiện vấn đề", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Nguyên nhân kiện vấn đề", HeaderText = "Nguyên nhân kiện vấn đề", DataPropertyName = "Nguyên nhân kiện vấn đề", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Bưu cục thao tác cuối", HeaderText = "Bưu cục thao tác cuối", DataPropertyName = "Bưu cục thao tác cuối", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Người thao tác", HeaderText = "Người thao tác", DataPropertyName = "Người thao tác", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Dấu chuyển hoàn", HeaderText = "Dấu chuyển hoàn", DataPropertyName = "Dấu chuyển hoàn", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Địa chỉ nhận hàng", HeaderText = "Địa chỉ nhận hàng", DataPropertyName = "Địa chỉ nhận hàng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Phường", HeaderText = "Phường/Xã", DataPropertyName = "Phường", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Nội dung hàng hóa", HeaderText = "Nội dung hàng hóa", DataPropertyName = "Nội dung hàng hóa", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "COD thực tế", HeaderText = "COD thực tế", DataPropertyName = "COD thực tế", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "PTTT", HeaderText = "PTTT", DataPropertyName = "PTTT", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Nhân viên nhận hàng", HeaderText = "Nhân viên nhận hàng", DataPropertyName = "Nhân viên nhận hàng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Địa chỉ lấy hàng", HeaderText = "Địa chỉ lấy hàng", DataPropertyName = "Địa chỉ lấy hàng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Thời gian nhận hàng", HeaderText = "Thời gian nhận hàng", DataPropertyName = "Thời gian nhận hàng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Tên người gửi", HeaderText = "Tên người gửi", DataPropertyName = "Tên người gửi", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Trọng lượng", HeaderText = "Trọng lượng", DataPropertyName = "Trọng lượng", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 1", HeaderText = "Mã đoạn 1", DataPropertyName = "Mã đoạn 1", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 2", HeaderText = "Mã đoạn 2", DataPropertyName = "Mã đoạn 2", SortMode = DataGridViewColumnSortMode.Programmatic },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 3", HeaderText = "Mã đoạn 3", DataPropertyName = "Mã đoạn 3", SortMode = DataGridViewColumnSortMode.Programmatic }
            });
        }

        public void LoadAllDataToGrid()
        {
            if (_dataGrid.InvokeRequired)
            {
                _dataGrid.Invoke(new Action(LoadAllDataToGrid));
                return;
            }

            SyncDisplayTableFromAllRows();

            if (!string.IsNullOrEmpty(_currentSortColumn))
                ApplyCurrentSort();

            if (!_columnsResized)
            {
                _columnsResized = true;
                _dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }

            UpdateColumnSortGlyphs();
        }

        private void SyncDisplayTableFromAllRows()
        {
            _displayTable.Rows.Clear();
            foreach (var r in _allRows)
            {
                _displayTable.Rows.Add(
                    r.WaybillNo,
                    r.TrangThaiHienTai,
                    r.ThaoTacCuoi,
                    r.ThoiGianThaoTac,
                    r.ThoiGianYeuCauPhatLai,
                    r.NhanVienKienVanDe,
                    r.NguyenNhanKienVanDe,
                    r.BuuCucThaoTac,
                    r.NguoiThaoTac,
                    r.DauChuyenHoan,
                    r.DiaChiNhanHang,
                    r.Phuong,
                    r.NoiDungHangHoa,
                    r.CODThucTe,
                    r.PTTT,
                    r.NhanVienNhanHang,
                    r.DiaChiLayHang,
                    r.ThoiGianNhanHang,
                    r.TenNguoiGui,
                    r.TrongLuong,
                    r.MaDoan1,
                    r.MaDoan2,
                    r.MaDoan3
                );
            }
        }

        private void InitProgressControls()
        {
            if (_progressBar != null)
            {
                _progressBar.Maximum = 100;
                _progressBar.Value = 0;
                _progressBar.Visible = false;
                _progressBar.ShowValue = true;
            }
        }

        private void InitializeProgress(int total)
        {
            _totalItems = total;
            _completedItems = 0;
            _lastPercent = -1;
            UpdateProgress(0);
        }

        private void ReportProgress(int increment)
        {
            int current = Interlocked.Add(ref _completedItems, increment);
            int percent = _totalItems > 0 ? (current * 100 / _totalItems) : 0;
            
            if (percent != _lastPercent)
            {
                _lastPercent = percent;
                UpdateProgress(percent);
            }
        }

        private void UpdateProgress(int percent)
        {
            if (_progressBar == null || _progressBar.IsDisposed) return;
            if (_progressBar.InvokeRequired)
            {
                _progressBar.BeginInvoke(new Action(() => UpdateProgress(percent)));
                return;
            }
            _progressBar.Value = Math.Min(percent, 100);
            if (!_progressBar.Visible) _progressBar.Visible = true;
        }

        private void CompleteProgress()
        {
            UpdateProgress(100);
            Task.Delay(800).ContinueWith(_ =>
            {
                if (_progressBar != null && !_progressBar.IsDisposed)
                {
                    _progressBar.BeginInvoke(new Action(() =>
                    {
                        _progressBar.Visible = false;
                        _progressBar.Value = 0;
                    }));
                }
            });
        }

        private DataTable CreateDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Mã vận đơn", typeof(string));
            dt.Columns.Add("Trạng thái hiện tại", typeof(string));
            dt.Columns.Add("Thao tác cuối cùng", typeof(string));
            dt.Columns.Add("Thời gian thao tác", typeof(string));
            dt.Columns.Add("Thời gian yêu cầu phát lại", typeof(string));
            dt.Columns.Add("Nhân viên kiện vấn đề", typeof(string));
            dt.Columns.Add("Nguyên nhân kiện vấn đề", typeof(string));
            dt.Columns.Add("Bưu cục thao tác cuối", typeof(string));
            dt.Columns.Add("Người thao tác", typeof(string));
            dt.Columns.Add("Dấu chuyển hoàn", typeof(string));
            dt.Columns.Add("Địa chỉ nhận hàng", typeof(string));
            dt.Columns.Add("Phường", typeof(string));
            dt.Columns.Add("Nội dung hàng hóa", typeof(string));
            dt.Columns.Add("COD thực tế", typeof(string));
            dt.Columns.Add("PTTT", typeof(string));
            dt.Columns.Add("Nhân viên nhận hàng", typeof(string));
            dt.Columns.Add("Địa chỉ lấy hàng", typeof(string));
            dt.Columns.Add("Thời gian nhận hàng", typeof(string));
            dt.Columns.Add("Tên người gửi", typeof(string));
            dt.Columns.Add("Trọng lượng", typeof(string));
            dt.Columns.Add("Mã đoạn 1", typeof(string));
            dt.Columns.Add("Mã đoạn 2", typeof(string));
            dt.Columns.Add("Mã đoạn 3", typeof(string));
            return dt;
        }

        public static void EnableDoubleBuffering(DataGridView dgv)
        {
            if (SystemInformation.TerminalServerSession) return;
            var pi = typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            pi?.SetValue(dgv, true, null);
        }

        public async Task SearchTrackingAsync(string waybillsText, bool updateMainGrid = true)
        {
            if (string.IsNullOrEmpty(Main.CapturedAuthToken)) return;

            // 5. YIELD NGAY TỪ ĐẦU ĐỂ KHÔNG BLOCK UI KHI BỊ MINIMIZE
            await Task.Yield();

            var waybillList = ExtractWaybills(waybillsText);
            if (waybillList.Count == 0)
            {
                if (updateMainGrid) MessageBox.Show("Nhập mã vận đơn!", "Thông báo");
                return;
            }

            if (updateMainGrid) InitializeProgress(waybillList.Count * 2);

            _allRows.Clear();
            _rowDict.Clear();
            _columnsResized = false; // Đặt lại cờ resize khi bắt đầu tra cứu mới

            var codesToQuery = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var mapping001 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var wb in waybillList)
            {
                string code = wb.Trim().ToUpper();
                codesToQuery.Add(code);

                var row = new TrackingRow { WaybillNo = code };
                _allRows.Add(row);
                _rowDict[code] = row;

                if (code.Contains("-"))
                {
                    string original = code.Split('-')[0];
                    codesToQuery.Add(original);
                    mapping001[code] = original;

                    if (!_rowDict.ContainsKey(original))
                    {
                        _rowDict[original] = new TrackingRow { WaybillNo = original };
                    }
                }
            }

            if (updateMainGrid) _displayTable.Rows.Clear();

            // Gọi API song song
            await ProcessTrackingBatchesAsync(codesToQuery.ToList());
            await ProcessOrderDetailBatchesAsync(codesToQuery.ToList());

            // 6. VERIFY: QUÉT LẠI CÁC ĐƠN BỊ TRẮNG DATA (RẤT QUAN TRỌNG)
            var missing = _allRows.Where(x => string.IsNullOrWhiteSpace(x.TrangThaiHienTai))
                                  .Select(x => x.WaybillNo)
                                  .ToList();
            if (missing.Count > 0)
            {
                await ProcessTrackingBatchesAsync(missing);
                await ProcessOrderDetailBatchesAsync(missing);
            }

            // Đồng bộ dữ liệu mã gốc sang mã hậu tố
            foreach (var kvp in mapping001)
            {
                string codeWithDash = kvp.Key;
                string originalCode = kvp.Value;

                if (_rowDict.TryGetValue(originalCode, out var srcRow) && _rowDict.TryGetValue(codeWithDash, out var destRow))
                {
                    lock (_rowLock)
                    {
                        destRow.NhanVienNhanHang = srcRow.NhanVienNhanHang ?? "";
                        if (string.IsNullOrEmpty(destRow.TenNguoiGui)) destRow.TenNguoiGui = srcRow.TenNguoiGui ?? "";
                        if (string.IsNullOrEmpty(destRow.DiaChiLayHang)) destRow.DiaChiLayHang = srcRow.DiaChiLayHang ?? "";
                    }
                }
            }

            if (updateMainGrid)
            {
                LoadAllDataToGrid();
                CompleteProgress();
            }
        }

        private List<string> ExtractWaybills(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return new List<string>();
            return text.Split(new[] { '\r', '\n', ',', ';', '|', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries)
                       .Select(t => t.Trim().ToUpper())
                       .Where(clean => clean.Length > 5)
                       .Distinct(StringComparer.OrdinalIgnoreCase)
                       .ToList();
        }

        private async Task ProcessTrackingBatchesAsync(List<string> waybills)
        {
            var batches = waybills.Distinct(StringComparer.OrdinalIgnoreCase).Chunk(BatchSize).ToList();

            // ĐÃ XÓA TOÀN BỘ SEMAPHORE
            await Parallel.ForEachAsync(batches, new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrency }, async (batch, ct) =>
            {
                await CallTrackingBatchWithRetryAsync(batch.ToList(), ct);
                ReportProgress(batch.Length);
            });
        }

        private async Task CallTrackingBatchWithRetryAsync(List<string> batch, CancellationToken ct)
        {
            for (int retry = 1; retry <= MaxRetry; retry++)
            {
                try
                {
                    await CallTrackingBatchAsync(batch);
                    return;
                }
                catch
                {
                    if (retry >= MaxRetry) return;
                    await Task.Delay(1000 * retry, ct);
                }
            }
        }

        private async Task CallTrackingBatchAsync(IEnumerable<string> batch)
        {
            var url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/podTracking/inner/query/keywordList");
            var payload = new { keywordList = batch, trackingTypeEnum = "WAYBILL", countryId = "1" };
            
            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("authToken", Main.CapturedAuthToken);
            request.Headers.Add("lang", "VN");
            request.Headers.Add("langType", "VN");
            request.Headers.Add("routeName", "trackingExpress");

            // Cấu hình False để chạy thoát khỏi SyncContext của UI
            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);
            
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                var result = JsonSerializer.Deserialize<WaybillHistoryResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                
                if (result?.succ == true && result.data != null)
                {
                    foreach (var item in result.data)
                    {
                        string wb = item.keyword ?? item.billCode ?? "";
                        if (!string.IsNullOrEmpty(wb)) ProcessTrackingData(wb, item);
                    }
                }
            }
            else
            {
                // Quăng lỗi để kích hoạt vòng lặp Retry ở hàm mẹ
                throw new Exception("HTTP Lỗi");
            }
        }

        private void ProcessTrackingData(string waybill, WaybillData wd)
        {
            if (string.IsNullOrEmpty(waybill)) return;

            lock (_rowLock)
            {
                if (!_rowDict.TryGetValue(waybill, out var row)) return;

                var details = wd.details ?? new List<WaybillDetail>();
                if (details.Count == 0) return;

                row.DauChuyenHoan = details.Any(d => d.status == "已审核") ? "Có" : "Không";

                var staffNameFisrt = details
                    .Where(d => (d.scanTypeName?.Contains("Nhận hàng") == true || d.scanTypeName?.Contains("Lấy hàng") == true || d.status == "已揽件") &&
                                (!string.IsNullOrEmpty(d.staffName) || !string.IsNullOrEmpty(d.scanByName)))
                    .OrderBy(d => DateTime.TryParse(!string.IsNullOrEmpty(d.uploadTime) ? d.uploadTime : (d.scanTime ?? "9999-12-31"), out DateTime dt) ? dt : DateTime.MaxValue)
                    .FirstOrDefault();
                if (staffNameFisrt != null) row.NhanVienNhanHang = staffNameFisrt.staffName ?? staffNameFisrt.scanByName ?? "";

                var giaoLaiGanNhat = details.Where(d => d.scanTypeName != null && d.scanTypeName.Contains("Giao lại hàng"))
                    .OrderByDescending(d => DateTime.TryParse(d.uploadTime ?? d.scanTime ?? "", out DateTime dt) ? dt : DateTime.MinValue).FirstOrDefault();
                row.ThoiGianYeuCauPhatLai = giaoLaiGanNhat?.remark2 ?? "";

                var vanDe = details.Where(d => d.scanTypeName?.Contains("vấn đề") == true || d.scanTypeName?.Contains("Kiện vấn đề") == true)
                    .OrderByDescending(d => DateTime.TryParse(d.uploadTime ?? d.scanTime ?? "", out DateTime dt) ? dt : DateTime.MinValue).FirstOrDefault();
                if (vanDe != null)
                {
                    row.NhanVienKienVanDe = vanDe.scanByName ?? "";
                    row.NguyenNhanKienVanDe = vanDe.remark1 ?? "";
                }

                var latest = details.OrderByDescending(d =>
                {
                    DateTime.TryParse(d.uploadTime ?? d.scanTime, out var dt);
                    return dt;
                }).FirstOrDefault();

                if (latest != null)
                {
                    row.TrangThaiHienTai = latest.waybillTrackingContent ?? "";
                    row.ThaoTacCuoi = latest.scanTypeName ?? "";
                    row.ThoiGianThaoTac = latest.uploadTime ?? latest.scanTime ?? "";
                    row.BuuCucThaoTac = latest.scanNetworkName ?? "";
                    row.NguoiThaoTac = latest.scanByName ?? "";
                }
            }
        }

        private async Task ProcessOrderDetailBatchesAsync(List<string> waybills)
        {
            await Parallel.ForEachAsync(waybills, new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrency }, async (waybill, ct) =>
            {
                await CallOrderDetailWithRetryAsync(waybill, ct);
                ReportProgress(1);
            });
        }

        private async Task CallOrderDetailWithRetryAsync(string waybill, CancellationToken ct)
        {
            for (int retry = 1; retry <= MaxRetry; retry++)
            {
                try
                {
                    await CallOrderDetailAsync(waybill);
                    return;
                }
                catch
                {
                    if (retry >= MaxRetry) return;
                    await Task.Delay(1000 * retry, ct);
                }
            }
        }

        private async Task CallOrderDetailAsync(string waybill)
        {
            var url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/order/getOrderDetail");
            var payload = new { waybillNo = waybill };
            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("authToken", Main.CapturedAuthToken);
            request.Headers.Add("lang", "VN");
            request.Headers.Add("langType", "VN");
            request.Headers.Add("routeName", "trackingExpress");

            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);
            
            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                var result = JsonSerializer.Deserialize<OrderDetailResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (result?.succ == true && result.data?.details != null)
                {
                    UpdateOrderDetail(waybill, result.data.details);
                }
            }
            else
            {
                throw new Exception("HTTP Lỗi");
            }
        }

        private void UpdateOrderDetail(string waybill, OrderDetailInfo info)
        {
            if (string.IsNullOrEmpty(waybill)) return;

            lock (_rowLock)
            {
                if (!_rowDict.TryGetValue(waybill, out var row)) return;

                row.NhanVienNhanHang = string.IsNullOrEmpty(row.NhanVienNhanHang) ? (info.staffName ?? "") : row.NhanVienNhanHang;
                row.DiaChiLayHang = info.senderDetailedAddress ?? "";
                row.ThoiGianNhanHang = info.pickTime ?? "";
                row.NoiDungHangHoa = info.goodsName ?? "";
                row.CODThucTe = info.codMoney ?? "";
                row.TenNguoiGui = info.customerName ?? "";
                row.TrongLuong = info.packageChargeWeight?.ToString() ?? "";
                row.PTTT = info.paymentModeName ?? "";
                row.DiaChiNhanHang = info.receiverDetailedAddress ?? "";
                row.Phuong = info.destinationName ?? "";

                if (!string.IsNullOrEmpty(info.terminalDispatchCode))
                {
                    var parts = info.terminalDispatchCode.Split('-');
                    row.MaDoan1 = parts.Length > 0 ? parts[0] : "";
                    row.MaDoan2 = parts.Length > 1 ? parts[1] : "";
                    row.MaDoan3 = parts.Length > 2 ? parts[2] : "";
                    row.MaDoanFull = info.terminalDispatchCode;
                }
            }
        }

        private void DataGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex != -1 || e.ColumnIndex < 0) return;

            var column = _dataGrid.Columns[e.ColumnIndex];
            if (column == null) return;

            var nextDirection = ListSortDirection.Ascending;
            if (_currentSortColumn == column.Name && _currentSortDirection == ListSortDirection.Ascending)
                nextDirection = ListSortDirection.Descending;

            _currentSortColumn = column.Name;
            _currentSortDirection = nextDirection;
            ApplyCurrentSort();
        }

        private void ApplyCurrentSort()
        {
            if (_allRows.Count == 0 || string.IsNullOrWhiteSpace(_currentSortColumn))
            {
                _bindingSource.Sort = string.Empty;
                return;
            }

            string columnName = _currentSortColumn switch
            {
                "Mã vận đơn" => "Mã vận đơn",
                "Trạng thái hiện tại" => "Trạng thái hiện tại",
                "Thao tác cuối cùng" => "Thao tác cuối cùng",
                "Thời gian thao tác" => "Thời gian thao tác",
                "Thời gian yêu cầu phát lại" => "Thời gian yêu cầu phát lại",
                "Nhân viên kiện vấn đề" => "Nhân viên kiện vấn đề",
                "Nguyên nhân kiện vấn đề" => "Nguyên nhân kiện vấn đề",
                "Bưu cục thao tác cuối" => "Bưu cục thao tác cuối",
                "Người thao tác" => "Người thao tác",
                "Dấu chuyển hoàn" => "Dấu chuyển hoàn",
                "Địa chỉ nhận hàng" => "Địa chỉ nhận hàng",
                "Phường" => "Phường",
                "Nội dung hàng hóa" => "Nội dung hàng hóa",
                "COD thực tế" => "COD thực tế",
                "PTTT" => "PTTT",
                "Nhân viên nhận hàng" => "Nhân viên nhận hàng",
                "Địa chỉ lấy hàng" => "Địa chỉ lấy hàng",
                "Thời gian nhận hàng" => "Thời gian nhận hàng",
                "Tên người gửi" => "Tên người gửi",
                "Trọng lượng" => "Trọng lượng",
                "Mã đoạn 1" => "Mã đoạn 1",
                "Mã đoạn 2" => "Mã đoạn 2",
                "Mã đoạn 3" => "Mã đoạn 3",
                _ => null
            };

            if (columnName == null || !_displayTable.Columns.Contains(columnName))
                return;

            string direction = _currentSortDirection == ListSortDirection.Ascending ? "ASC" : "DESC";
            _bindingSource.Sort = $"[{columnName}] {direction}";
            UpdateColumnSortGlyphs();
        }

        private void UpdateColumnSortGlyphs()
        {
            foreach (DataGridViewColumn col in _dataGrid.Columns)
            {
                if (col.Name == _currentSortColumn)
                    col.HeaderCell.SortGlyphDirection = _currentSortDirection == ListSortDirection.Ascending
                        ? SortOrder.Ascending
                        : SortOrder.Descending;
                else
                    col.HeaderCell.SortGlyphDirection = SortOrder.None;
            }
        }

        public void ClearData()
        {
            _allRows.Clear();
            _rowDict.Clear();
            _displayTable?.Rows.Clear();
            _currentSortColumn = string.Empty;
            _currentSortDirection = ListSortDirection.Ascending;
            _bindingSource.Sort = string.Empty;
            _bindingSource.ResetBindings(false);
            _columnsResized = false;
            foreach (DataGridViewColumn col in _dataGrid.Columns)
                col.HeaderCell.SortGlyphDirection = SortOrder.None;
            if (_progressBar != null) _progressBar.Visible = false;
        }

        // ====================== EXPORT EXCEL ======================
        public void ExportToExcel()
        {
            if (_allRows.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu!", "Thông báo");
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd-MM-yyyy}.xlsx",
                InitialDirectory = _exportFolder
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Trạng thái hiện tại");
                var exportTable = _displayTable.Copy();

                var table = ws.Cell(1, 1).InsertTable(exportTable, "TrackingTable", true);

                ws.Range(1, 1, table.RowCount() + 1, table.ColumnCount()).Clear(XLClearOptions.AllFormats);
                var headerRow = table.HeadersRow();
                headerRow.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.2);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Font.FontColor = XLColor.Black;
                headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                if (table.RowCount() > 1)
                {
                    var dataRange = ws.Range(2, 1, table.RowCount() + 1, table.ColumnCount());
                    dataRange.Style.Fill.BackgroundColor = XLColor.White;
                }
                ws.Columns().AdjustToContents();

                wb.SaveAs(sfd.FileName);

                DialogResult result = MessageBox.Show($"Xuất {sfd.FileName} thành công, Mở file ngay?",
                                              "Xuất file thành công",
                                              MessageBoxButtons.YesNo,
                                              MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName)
                    {
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi export: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ====================== EXPORT CHIA CHỨC NĂNG (btn_Export_Spe) ======================
        public void ExportSpecial()
        {
            if (_allRows.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd-MM-yyyy}.xlsx",
                InitialDirectory = _exportFolder
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();

                var phatHangRows = _allRows.Where(r => r.DauChuyenHoan != "Có").ToList();
                if (phatHangRows.Any())
                {
                    var ws = wb.Worksheets.Add("PHÁT HÀNG");
                    var dtPhat = CreatePhatHangDataTable(phatHangRows);
                    var table = ws.Cell(1, 1).InsertTable(dtPhat, "PhatHangTable", true);
                    ApplyHeaderStyle(ws, table);
                }

                var hoanPhatRows = _allRows.Where(r => r.DauChuyenHoan == "Có").ToList();
                if (hoanPhatRows.Any())
                {
                    var ws = wb.Worksheets.Add("HOÀN PHÁT");
                    var dtHoan = CreateHoanPhatDataTable(hoanPhatRows);
                    var table = ws.Cell(1, 1).InsertTable(dtHoan, "HoanPhatTable", true);
                    ApplyHeaderStyle(ws, table);
                }

                wb.SaveAs(sfd.FileName);
                if (MessageBox.Show($"Xuất file thành công!\n\nFile: {sfd.FileName}\n\nMở file ngay?", "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show("Lỗi export: " + ex.Message); }
        }

        private DataTable CreatePhatHangDataTable(List<TrackingRow> rows)
        {
            var dt = new DataTable("PHÁT HÀNG");
            dt.Columns.Add("Mã vận đơn");
            dt.Columns.Add("Mã đoạn 3");
            dt.Columns.Add("Địa chỉ nhận hàng");
            dt.Columns.Add("Phường/Xã");
            dt.Columns.Add("Trạng thái hiện tại");
            dt.Columns.Add("Thao tác cuối cùng");
            dt.Columns.Add("Thời gian thao tác");
            dt.Columns.Add("Nội dung hàng hóa");
            dt.Columns.Add("COD thực tế");
            dt.Columns.Add("Trọng lượng");
            dt.Columns.Add("Mã đoạn 1");
            dt.Columns.Add("Mã đoạn 2");
            foreach (var r in rows)
                dt.Rows.Add(r.WaybillNo, r.MaDoan3, r.DiaChiNhanHang, r.Phuong, r.TrangThaiHienTai, r.ThaoTacCuoi, r.ThoiGianThaoTac, r.NoiDungHangHoa, r.CODThucTe, r.TrongLuong, r.MaDoan1, r.MaDoan2);
            return dt;
        }

        private DataTable CreateHoanPhatDataTable(List<TrackingRow> rows)
        {
            var dt = new DataTable("HOÀN PHÁT");
            dt.Columns.Add("Mã vận đơn");
            dt.Columns.Add("Nhân viên nhận hàng");
            dt.Columns.Add("Tên người gửi");
            dt.Columns.Add("Địa chỉ lấy hàng");
            dt.Columns.Add("Nội dung hàng hóa");
            dt.Columns.Add("COD thực tế");
            dt.Columns.Add("PTTT");
            dt.Columns.Add("Trọng lượng");
            foreach (var r in rows)
                dt.Rows.Add(r.WaybillNo, r.NhanVienNhanHang, r.TenNguoiGui, r.DiaChiLayHang, r.NoiDungHangHoa, r.CODThucTe, r.PTTT, r.TrongLuong);
            return dt;
        }

        private void ApplyHeaderStyle(IXLWorksheet ws, IXLTable table)
        {
            ws.Range(1, 1, table.RowCount() + 1, table.ColumnCount()).Clear(XLClearOptions.AllFormats);

            var headerRow = table.HeadersRow();
            headerRow.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.2);
            headerRow.Style.Font.Bold = true;
            headerRow.Style.Font.FontColor = XLColor.Black;
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            headerRow.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            if (table.RowCount() > 1)
            {
                var dataRange = ws.Range(2, 1, table.RowCount() + 1, table.ColumnCount());
                dataRange.Style.Fill.BackgroundColor = XLColor.White;
            }

            ws.Columns().AdjustToContents();
        }

        public List<TrackingRow> GetAllRows() => _allRows ?? new List<TrackingRow>();

        public async Task<string> GetDKCHHistoryAsync(string waybill) => await GetFullHistoryFromArrival(waybill);

        private async Task<string> GetFullHistoryFromArrival(string waybill)
        {
            try
            {
                var url = AppConfig.Current.BuildJmsApiUrl("operatingplatform/podTracking/inner/query/keywordList");
                var payload = new { keywordList = new[] { waybill }, trackingTypeEnum = "WAYBILL", countryId = "1" };
                var request = new HttpRequestMessage(HttpMethod.Post, url)
                {
                    Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
                };
                request.Headers.Add("authToken", Main.CapturedAuthToken);
                request.Headers.Add("lang", "VN");
                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();
                var result = JsonSerializer.Deserialize<WaybillHistoryResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (result?.succ != true || result?.data?[0]?.details == null)
                    return "Không có dữ liệu hành trình.";
                var details = result.data[0].details;
                int arrivalIndex = -1;
                for (int i = 0; i < details.Count; i++)
                {
                    var d = details[i];
                    string type = d.scanTypeName ?? "";
                    string network = d.scanNetworkName ?? "";
                    bool isArrival = type.Contains("Xuống hàng kiện đến") || type.Contains("Xuống kiện") ||
                                     type.Contains("卸车到件") || type.Contains("到件");
                    bool isKimTan = network.Contains("Kim Tân") || network.Contains("(LCI)");
                    if (isArrival && isKimTan)
                    {
                        arrivalIndex = i;
                        break;
                    }
                }
                if (arrivalIndex == -1) return "";
                var sb = new StringBuilder();
                for (int i = 0; i <= arrivalIndex; i++)
                {
                    var d = details[i];
                    string type = d.scanTypeName ?? "";
                    if (type.Contains("Kiểm tra hàng tồn kho") || type.Contains("派件电联") ||
                        type.Contains("Lịch sử cuộc gọi") || type.Contains("cuộc gọi-phát"))
                        continue;
                    sb.AppendLine($"Mã vận đơn: {waybill}");
                    sb.AppendLine($"👤 {d.scanByName}");
                    sb.AppendLine($"Thao tác: {CleanType(type)}");
                    if (!string.IsNullOrEmpty(d.remark1))
                        sb.AppendLine($" └─ Ghi chú: {d.remark1}");
                    sb.AppendLine("──────────────────────────────────────────────");
                }
                return sb.ToString();
            }
            catch (Exception ex) { return $"Lỗi hiển thị lịch sử: {ex.Message}"; }
        }

        private string CleanType(string type)
        {
            if (string.IsNullOrEmpty(type)) return type;
            return type
                .Replace("卸车到件", "Xuống hàng kiện đến")
                .Replace("库存盘点", "Kiểm tra hàng tồn kho")
                .Replace("退件登记", "Đăng ký chuyển hoàn")
                .Replace("再次登记", "Đăng ký chuyển hoàn lần 2")
                .Replace("问题件扫描", "Quét kiện vấn đề")
                .Replace("出仓扫描", "Quét phát hàng")
                .Replace("快件签收", "Ký nhận CPN")
                .Replace("退件签收", "Ký nhận chuyển hoàn")
                .Replace("拆包扫描", "Gỡ bao")
                .Replace("退件扫描", "In đơn chuyển hoàn")
                .Replace("重派", "Giao lại hàng")
                .Replace("退件确认", "Xác nhận chuyển hoàn")
                .Replace("【", "[")
                .Replace("】", "]")
                .Replace("已", "Đã ")
                .Replace("进行", "")
                .Replace("扫描员", "Nhân viên quét mã ")
                .Replace("登记人员是", "người đăng ký là ")
                .Replace("退回原因", "nguyên nhân chuyển hoàn ")
                .Replace("问题件原因", "nguyên nhân kiện vấn đề ")
                .Replace("上一站是", "trạm trước là ")
                .Replace("任务编号", "mã nhiệm vụ ")
                .Replace("的派件员", "nhân viên phát hàng ")
                .Replace("正在派件", "đang phát hàng");
        }
    }

    // Các class Model giữ nguyên...
    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class TrackingRow
    {
        public string WaybillNo { get; set; }
        public string TrangThaiHienTai { get; set; }
        public string ThaoTacCuoi { get; set; }
        public string ThoiGianThaoTac { get; set; }
        public string ThoiGianYeuCauPhatLai { get; set; } = "";
        public string NhanVienKienVanDe { get; set; } = "";
        public string NguyenNhanKienVanDe { get; set; } = "";
        public string BuuCucThaoTac { get; set; }
        public string NguoiThaoTac { get; set; }
        public string DauChuyenHoan { get; set; } = "Không";
        public string DiaChiNhanHang { get; set; }
        public string Phuong { get; set; }
        public string NoiDungHangHoa { get; set; }
        public string CODThucTe { get; set; }
        public string PTTT { get; set; }
        public string NhanVienNhanHang { get; set; } = "";
        public string DiaChiLayHang { get; set; }
        public string ThoiGianNhanHang { get; set; }
        public string TenNguoiGui { get; set; }
        public string TrongLuong { get; set; }
        public string MaDoanFull { get; set; } = "";
        public string MaDoan1 { get; set; }
        public string MaDoan2 { get; set; }
        public string MaDoan3 { get; set; }
        public string RebackStatus { get; set; }
        public int PrintCount { get; set; }
        public string NewTerminalDispatchCode { get; set; }
        public string InHoanScanTime { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillHistoryResponse
    {
        public bool succ { get; set; }
        public List<WaybillData> data { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillData
    {
        public string keyword { get; set; }
        public string billCode { get; set; }
        public List<WaybillDetail> details { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class WaybillDetail
    {
        public string status { get; set; }
        public string scanTypeName { get; set; }
        public string uploadTime { get; set; }
        public string scanTime { get; set; }
        public string scanNetworkName { get; set; }
        public string scanByName { get; set; }
        public string staffName { get; set; }
        public string remark1 { get; set; }
        public string remark2 { get; set; }
        public string waybillTrackingContent { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class OrderDetailResponse
    {
        public bool succ { get; set; }
        public OrderDetailData data { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class OrderDetailData
    {
        public OrderDetailInfo details { get; set; }
    }

    [System.Reflection.Obfuscation(Exclude = true, ApplyToMembers = true)]
    public class OrderDetailInfo
    {
        public string staffName { get; set; }
        public string senderDetailedAddress { get; set; }
        public string pickTime { get; set; }
        public string goodsName { get; set; }
        public string codMoney { get; set; }
        public string customerName { get; set; }
        public double? packageChargeWeight { get; set; }
        public string paymentModeName { get; set; }
        public string receiverDetailedAddress { get; set; }
        public string destinationName { get; set; }
        public string terminalDispatchCode { get; set; }
    }
}