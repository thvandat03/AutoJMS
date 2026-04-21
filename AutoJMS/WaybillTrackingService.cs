using AutoJMS.Data;
using ClosedXML.Excel;
using Sunny.UI;
using System;
using System.Collections.Generic;
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
        private readonly string _exportFolder;
        private readonly HttpClient _httpClient;
        private readonly SemaphoreSlim _semaphore;

        private readonly UIProcessBar _progressBar;
        private int _totalItems;
        private int _completedItems;

        private readonly List<TrackingRow> _allRows = new List<TrackingRow>();
        private readonly Dictionary<string, TrackingRow> _rowDict = new Dictionary<string, TrackingRow>(StringComparer.OrdinalIgnoreCase);

        private const int BatchSize = 40;
        private const int MaxConcurrency = 8;

        public WaybillTrackingService(DataGridView dataGrid, UIProcessBar progressBar = null)
        {
            _dataGrid = dataGrid ?? throw new ArgumentNullException(nameof(dataGrid));
            _progressBar = progressBar;

            _exportFolder = Path.Combine(Application.StartupPath, "Downloads", "Trạng thái hiện tại");
            Directory.CreateDirectory(_exportFolder);

            _displayTable = CreateDataTable();

            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            _semaphore = new SemaphoreSlim(MaxConcurrency, MaxConcurrency);

            SetupDataGridViewVirtualMode();
            InitProgressControls();
        }

        private void SetupDataGridViewVirtualMode()
        {
            EnableDoubleBuffering(_dataGrid);
            _dataGrid.VirtualMode = true;
            _dataGrid.AutoGenerateColumns = false;
            _dataGrid.AllowUserToAddRows = false;
            _dataGrid.ReadOnly = true;
            _dataGrid.RowHeadersVisible = false;
            _dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            _dataGrid.ScrollBars = ScrollBars.Both;
            _dataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            _dataGrid.Columns.Clear();
            _dataGrid.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { Name = "Mã vận đơn", HeaderText = "Mã vận đơn" },
                new DataGridViewTextBoxColumn { Name = "Trạng thái hiện tại", HeaderText = "Trạng thái hiện tại" },
                new DataGridViewTextBoxColumn { Name = "Thao tác cuối cùng", HeaderText = "Thao tác cuối cùng" },
                new DataGridViewTextBoxColumn { Name = "Thời gian thao tác", HeaderText = "Thời gian thao tác" },
                new DataGridViewTextBoxColumn { Name = "Thời gian yêu cầu phát lại", HeaderText = "Thời gian yêu cầu phát lại" },
                new DataGridViewTextBoxColumn { Name = "Nhân viên kiện vấn đề", HeaderText = "Nhân viên kiện vấn đề" },
                new DataGridViewTextBoxColumn { Name = "Nguyên nhân kiện vấn đề", HeaderText = "Nguyên nhân kiện vấn đề" },
                new DataGridViewTextBoxColumn { Name = "Bưu cục thao tác cuối", HeaderText = "Bưu cục thao tác cuối" },
                new DataGridViewTextBoxColumn { Name = "Người thao tác", HeaderText = "Người thao tác" },
                new DataGridViewTextBoxColumn { Name = "Dấu chuyển hoàn", HeaderText = "Dấu chuyển hoàn" },
                new DataGridViewTextBoxColumn { Name = "Địa chỉ nhận hàng", HeaderText = "Địa chỉ nhận hàng" },
                new DataGridViewTextBoxColumn { Name = "Phường", HeaderText = "Phường/Xã" },
                new DataGridViewTextBoxColumn { Name = "Nội dung hàng hóa", HeaderText = "Nội dung hàng hóa" },
                new DataGridViewTextBoxColumn { Name = "COD thực tế", HeaderText = "COD thực tế" },
                new DataGridViewTextBoxColumn { Name = "PTTT", HeaderText = "PTTT" },
                new DataGridViewTextBoxColumn { Name = "Nhân viên nhận hàng", HeaderText = "Nhân viên nhận hàng" },
                new DataGridViewTextBoxColumn { Name = "Địa chỉ lấy hàng", HeaderText = "Địa chỉ lấy hàng" },
                new DataGridViewTextBoxColumn { Name = "Thời gian nhận hàng", HeaderText = "Thời gian nhận hàng" },
                new DataGridViewTextBoxColumn { Name = "Tên người gửi", HeaderText = "Tên người gửi" },
                new DataGridViewTextBoxColumn { Name = "Trọng lượng", HeaderText = "Trọng lượng" },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 1", HeaderText = "Mã đoạn 1" },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 2", HeaderText = "Mã đoạn 2" },
                new DataGridViewTextBoxColumn { Name = "Mã đoạn 3", HeaderText = "Mã đoạn 3" }
            });

            _dataGrid.CellValueNeeded += OnCellValueNeeded;
            _dataGrid.RowCount = 0;
        }

        private void OnCellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _allRows.Count) return;
            var row = _allRows[e.RowIndex];
            string colName = _dataGrid.Columns[e.ColumnIndex].Name;

            switch (colName)
            {
                case "Mã vận đơn": e.Value = row.WaybillNo; break;
                case "Trạng thái hiện tại": e.Value = row.TrangThaiHienTai; break;
                case "Thao tác cuối cùng": e.Value = row.ThaoTacCuoi; break;
                case "Thời gian thao tác": e.Value = row.ThoiGianThaoTac; break;
                case "Thời gian yêu cầu phát lại": e.Value = row.ThoiGianYeuCauPhatLai; break;
                case "Nhân viên kiện vấn đề": e.Value = row.NhanVienKienVanDe; break;
                case "Nguyên nhân kiện vấn đề": e.Value = row.NguyenNhanKienVanDe; break;
                case "Bưu cục thao tác cuối": e.Value = row.BuuCucThaoTac; break;
                case "Người thao tác": e.Value = row.NguoiThaoTac; break;
                case "Dấu chuyển hoàn": e.Value = row.DauChuyenHoan; break;
                case "Địa chỉ nhận hàng": e.Value = row.DiaChiNhanHang; break;
                case "Phường": e.Value = row.Phuong; break;
                case "Nội dung hàng hóa": e.Value = row.NoiDungHangHoa; break;
                case "COD thực tế": e.Value = row.CODThucTe; break;
                case "PTTT": e.Value = row.PTTT; break;
                case "Nhân viên nhận hàng": e.Value = row.NhanVienNhanHang; break;
                case "Địa chỉ lấy hàng": e.Value = row.DiaChiLayHang; break;
                case "Thời gian nhận hàng": e.Value = row.ThoiGianNhanHang; break;
                case "Tên người gửi": e.Value = row.TenNguoiGui; break;
                case "Trọng lượng": e.Value = row.TrongLuong; break;
                case "Mã đoạn 1": e.Value = row.MaDoan1; break;
                case "Mã đoạn 2": e.Value = row.MaDoan2; break;
                case "Mã đoạn 3": e.Value = row.MaDoan3; break;
            }
        }

        public void LoadAllDataToGrid()
        {
            if (_dataGrid.InvokeRequired)
            {
                _dataGrid.Invoke(new Action(LoadAllDataToGrid));
                return;
            }
            _dataGrid.RowCount = _allRows.Count;
            _dataGrid.Invalidate();
            _dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
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
            UpdateProgress(0);
        }

        private void ReportProgress(int increment)
        {
            int current = Interlocked.Add(ref _completedItems, increment);
            int percent = _totalItems > 0 ? (current * 100 / _totalItems) : 0;
            UpdateProgress(percent);
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

            var waybillList = ExtractWaybills(waybillsText);
            if (waybillList.Count == 0)
            {
                if (updateMainGrid) MessageBox.Show("Nhập mã vận đơn!", "Thông báo");
                return;
            }

            if (updateMainGrid) InitializeProgress(waybillList.Count * 2);

            _allRows.Clear();
            _rowDict.Clear();

            var codesToQuery = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var mapping001 = new Dictionary<string, string>();

            foreach (var wb in waybillList)
            {
                string code = wb.Trim().ToUpper();
                codesToQuery.Add(code);
                if (code.EndsWith("-001"))
                {
                    string original = code.Substring(0, code.Length - 4);
                    codesToQuery.Add(original);
                    mapping001[code] = original;
                }
                var row = new TrackingRow { WaybillNo = code };
                _allRows.Add(row);
                _rowDict[code] = row;
            }

            if (updateMainGrid)
            {
                _displayTable.Rows.Clear();
                _dataGrid.RowCount = 0;
            }

            await ProcessTrackingBatchesAsync(codesToQuery.ToList());
            await ProcessOrderDetailBatchesAsync(codesToQuery.ToList());

            foreach (var kvp in mapping001)
            {
                string code001 = kvp.Key;
                string originalCode = kvp.Value;
                if (_rowDict.TryGetValue(originalCode, out var srcRow) && _rowDict.TryGetValue(code001, out var destRow))
                {
                    destRow.NhanVienNhanHang = srcRow.NhanVienNhanHang ?? "";
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
            var batches = waybills.Chunk(BatchSize).ToList();
            await Parallel.ForEachAsync(batches, new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrency }, async (batch, ct) =>
            {
                await _semaphore.WaitAsync(ct);
                try { await CallTrackingBatchAsync(batch); ReportProgress(batch.Length); }
                finally { _semaphore.Release(); }
            });
        }

        private async Task CallTrackingBatchAsync(IEnumerable<string> batch)
        {
            var url = "https://jmsgw.jtexpress.vn/operatingplatform/podTracking/inner/query/keywordList";
            var payload = new { keywordList = batch, trackingTypeEnum = "WAYBILL", countryId = "1" };
            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("authToken", Main.CapturedAuthToken);
            request.Headers.Add("lang", "VN");
            request.Headers.Add("langType", "VN");
            request.Headers.Add("routeName", "trackingExpress");

            try
            {
                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();
                if (response.IsSuccessStatusCode)
                {
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
            }
            catch { }
        }

        private void ProcessTrackingData(string waybill, WaybillData wd)
        {
            if (string.IsNullOrEmpty(waybill) || !_rowDict.TryGetValue(waybill, out var row)) return;
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

            WaybillDetail latest = null;
            DateTime maxTime = DateTime.MinValue;
            foreach (var d in details)
            {
                string type = d.scanTypeName ?? "";
                if (type == "Kiểm tra hàng tồn kho" || type.Contains("Lịch sử cuộc gọi") || type.Contains("cuộc gọi-phát")) continue;
                string timeStr = d.uploadTime ?? d.scanTime;
                if (string.IsNullOrEmpty(timeStr)) continue;
                if (DateTime.TryParse(timeStr, out DateTime dt) && dt > maxTime)
                {
                    maxTime = dt;
                    latest = d;
                }
            }
            latest ??= details.LastOrDefault();
            if (latest != null)
            {
                row.TrangThaiHienTai = latest.waybillTrackingContent ?? "";
                row.ThaoTacCuoi = latest.scanTypeName ?? "";
                row.ThoiGianThaoTac = latest.uploadTime ?? latest.scanTime ?? "";
                row.BuuCucThaoTac = latest.scanNetworkName ?? "";
                row.NguoiThaoTac = latest.scanByName ?? "";
            }
        }

        private async Task ProcessOrderDetailBatchesAsync(List<string> waybills)
        {
            await Parallel.ForEachAsync(waybills, new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrency }, async (waybill, ct) =>
            {
                await _semaphore.WaitAsync(ct);
                try { await CallOrderDetailAsync(waybill); ReportProgress(1); }
                finally { _semaphore.Release(); }
            });
        }

        private async Task CallOrderDetailAsync(string waybill)
        {
            var url = "https://jmsgw.jtexpress.vn/operatingplatform/order/getOrderDetail";
            var payload = new { waybillNo = waybill };
            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            request.Headers.Add("authToken", Main.CapturedAuthToken);
            request.Headers.Add("lang", "VN");
            request.Headers.Add("langType", "VN");
            request.Headers.Add("routeName", "trackingExpress");

            try
            {
                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();
                if (response.IsSuccessStatusCode)
                {
                    var result = JsonSerializer.Deserialize<OrderDetailResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                    if (result?.succ == true && result.data?.details != null)
                        UpdateOrderDetail(waybill, result.data.details);
                }
            }
            catch { }
        }

        private void UpdateOrderDetail(string waybill, OrderDetailInfo info)
        {
            if (string.IsNullOrEmpty(waybill) || !_rowDict.TryGetValue(waybill, out var row)) return;
            if (string.IsNullOrEmpty(row.NhanVienNhanHang)) row.NhanVienNhanHang = info.staffName ?? "";
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

        public void ClearData()
        {
            _allRows.Clear();
            _rowDict.Clear();
            _displayTable?.Rows.Clear();
            _dataGrid.RowCount = 0;
            if (_progressBar != null) _progressBar.Visible = false;
        }

        public void ExportToExcel()
        {
            if (_allRows.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu!", "Thông báo");
                return;
            }
            _displayTable.Rows.Clear();
            foreach (var r in _allRows)
            {
                _displayTable.Rows.Add(
                    r.WaybillNo, r.TrangThaiHienTai, r.ThaoTacCuoi, r.ThoiGianThaoTac,
                    r.ThoiGianYeuCauPhatLai, r.NhanVienKienVanDe, r.NguyenNhanKienVanDe,
                    r.BuuCucThaoTac, r.NguoiThaoTac, r.DauChuyenHoan,
                    r.DiaChiNhanHang, r.Phuong, r.NoiDungHangHoa, r.CODThucTe,
                    r.PTTT, r.NhanVienNhanHang, r.DiaChiLayHang, r.ThoiGianNhanHang,
                    r.TenNguoiGui, r.TrongLuong, r.MaDoan1, r.MaDoan2, r.MaDoan3);
            }
            using var sfd = new SaveFileDialog { Filter = "Excel Files (*.xlsx)|*.xlsx", FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd/MM/yyyy}.xlsx", InitialDirectory = _exportFolder };
            if (sfd.ShowDialog() != DialogResult.OK) return;
            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Trạng thái hiện tại");
                var table = ws.Cell(1, 1).InsertTable(_displayTable.Copy(), "TrackingTable", true);
                ws.Range(1, 1, table.RowCount() + 1, table.ColumnCount()).Clear(XLClearOptions.AllFormats);
                var headerRow = table.HeadersRow();
                headerRow.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.2);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Columns().AdjustToContents();
                wb.SaveAs(sfd.FileName);
                if (MessageBox.Show($"Xuất {sfd.FileName} thành công, Mở file ngay?", "Xuất file thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show("Lỗi export: " + ex.Message); }
        }

        public void ExportSpecial()
        {
            if (_allRows.Count == 0)
            {
                MessageBox.Show("Chưa có dữ liệu để xuất!", "Thông báo");
                return;
            }
            using var sfd = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx",
                FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd/MM/yyyy}.xlsx",
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
            headerRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            table.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.Columns().AdjustToContents();
        }

        public List<TrackingRow> GetAllRows() => _allRows ?? new List<TrackingRow>();

        public async Task<string> GetDKCHHistoryAsync(string waybill) => await GetFullHistoryFromArrival(waybill);

        private async Task<string> GetFullHistoryFromArrival(string waybill)
        {
            try
            {
                var url = "https://jmsgw.jtexpress.vn/operatingplatform/podTracking/inner/query/keywordList";
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
}