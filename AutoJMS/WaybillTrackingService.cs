using AutoJMS.Data;
using ClosedXML.Excel;
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

        // === PROGRESS ===
        private readonly ProgressBar _progressBar;
        private readonly Label _lblPercent;
        private int _totalItems;
        private int _completedItems;

        private readonly List<TrackingRow> _allRows = new List<TrackingRow>();
        private readonly Dictionary<string, TrackingRow> _rowDict = new Dictionary<string, TrackingRow>(StringComparer.OrdinalIgnoreCase);

        private const int BatchSize = 40;
        private const int MaxConcurrency = 8;


        public WaybillTrackingService(DataGridView dataGrid, ProgressBar progressBar = null, Label lblPercent = null)
        {
            _dataGrid = dataGrid ?? throw new ArgumentNullException(nameof(dataGrid));
            _progressBar = progressBar;
            _lblPercent = lblPercent;

            _exportFolder = Path.Combine(Application.StartupPath, "Downloads", "Trạng thái hiện tại");
            Directory.CreateDirectory(_exportFolder);

            _displayTable = CreateDataTable();
            _dataGrid.DataSource = _displayTable;

            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
            _semaphore = new SemaphoreSlim(MaxConcurrency, MaxConcurrency);

            SetupDataGridView();
            InitProgressControls();
        }

        private void InitProgressControls()
        {
            if (_progressBar != null)
            {
                _progressBar.Minimum = 0;
                _progressBar.Maximum = 100;
                _progressBar.Style = ProgressBarStyle.Continuous;
                _progressBar.Visible = false;
            }
            if (_lblPercent != null)
                _lblPercent.Visible = false;
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
            Action action = () =>
            {
                if (_progressBar != null)
                {
                    _progressBar.Value = Math.Min(percent, 100);
                    _progressBar.Visible = true;
                }
                if (_lblPercent != null)
                {
                    _lblPercent.Text = $"{percent}%";
                    _lblPercent.Visible = true;
                }
            };
            if ((_progressBar?.InvokeRequired == true) || (_lblPercent?.InvokeRequired == true))
                _dataGrid.BeginInvoke(action);
            else
                action();
        }

        private void CompleteProgress()
        {
            UpdateProgress(100);
            Task.Delay(800).ContinueWith(_ =>
            {
                Action hide = () =>
                {
                    if (_progressBar != null) _progressBar.Visible = false;
                    if (_lblPercent != null) _lblPercent.Visible = false;
                };
                _dataGrid.BeginInvoke(hide);
            });
        }

        // ====================== DATA TABLE ======================
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

            dt.Columns["Mã vận đơn"].Caption = "Mã vận đơn";
            dt.Columns["Trạng thái hiện tại"].Caption = "Trạng thái hiện tại";
            dt.Columns["Thao tác cuối cùng"].Caption = "Thao tác cuối cùng";
            dt.Columns["Thời gian thao tác"].Caption = "Thời gian thao tác";
            dt.Columns["Thời gian yêu cầu phát lại"].Caption = "Thời gian yêu cầu phát lại";
            dt.Columns["Nhân viên kiện vấn đề"].Caption = "Nhân viên kiện vấn đề";
            dt.Columns["Nguyên nhân kiện vấn đề"].Caption = "Nguyên nhân kiện vấn đề";
            dt.Columns["Bưu cục thao tác cuối"].Caption = "Bưu cục thao tác cuối";
            dt.Columns["Người thao tác"].Caption = "Người thao tác";
            dt.Columns["Dấu chuyển hoàn"].Caption = "Dấu chuyển hoàn";
            dt.Columns["Địa chỉ nhận hàng"].Caption = "Địa chỉ nhận hàng";
            dt.Columns["Phường"].Caption = "Phường/Xã";
            dt.Columns["Nội dung hàng hóa"].Caption = "Nội dung hàng hóa";
            dt.Columns["COD thực tế"].Caption = "COD thực tế";
            dt.Columns["PTTT"].Caption = "PTTT";
            dt.Columns["Nhân viên nhận hàng"].Caption = "Nhân viên nhận hàng";
            dt.Columns["Địa chỉ lấy hàng"].Caption = "Địa chỉ lấy hàng";
            dt.Columns["Thời gian nhận hàng"].Caption = "Thời gian nhận hàng";
            dt.Columns["Tên người gửi"].Caption = "Tên người gửi";
            dt.Columns["Trọng lượng"].Caption = "Trọng lượng";
            dt.Columns["Mã đoạn 1"].Caption = "Mã đoạn 1";
            dt.Columns["Mã đoạn 2"].Caption = "Mã đoạn 2";
            dt.Columns["Mã đoạn 3"].Caption = "Mã đoạn 3";
            return dt;
        }

        private void SetupDataGridView()
        {
            EnableDoubleBuffering(_dataGrid);
            _dataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _dataGrid.ScrollBars = ScrollBars.Both;
            _dataGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _dataGrid.RowHeadersVisible = false;
            _dataGrid.AllowUserToResizeColumns = true;
            _dataGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _dataGrid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

        }

        public static void EnableDoubleBuffering(DataGridView dgv)
        {
            if (SystemInformation.TerminalServerSession) return;
            var pi = typeof(DataGridView).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            pi?.SetValue(dgv, true, null);
        }

        // ====================== TÌM KIẾM CHÍNH ======================
        public async Task SearchTrackingAsync(string waybillsText, bool updateMainGrid = true)
        {
            if (string.IsNullOrEmpty(Main.CapturedAuthToken))
            {
                
                return;
            }

            var waybillList = ExtractWaybills(waybillsText);
            if (waybillList.Count == 0)
            {
                if (updateMainGrid)
                    MessageBox.Show("Nhập mã vận đơn!", "Thông báo");
                return;
            }

            if (updateMainGrid)
                InitializeProgress(waybillList.Count * 2);

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
                _displayTable.Rows.Clear();

            await ProcessTrackingBatchesAsync(codesToQuery.ToList());
            await ProcessOrderDetailBatchesAsync(codesToQuery.ToList());

            foreach (var kvp in mapping001)
            {
                string code001 = kvp.Key;
                string originalCode = kvp.Value;

                if (_rowDict.TryGetValue(originalCode, out var srcRow) &&
                    _rowDict.TryGetValue(code001, out var destRow))
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

        // ====================== XỬ LÝ TRACKING =====================
        private async Task ProcessTrackingBatchesAsync(List<string> waybills)
        {
            var batches = waybills.Chunk(BatchSize).ToList();
            await Parallel.ForEachAsync(batches, new ParallelOptions { MaxDegreeOfParallelism = MaxConcurrency }, async (batch, ct) =>
            {
                await _semaphore.WaitAsync(ct);
                try
                {
                    await CallTrackingBatchAsync(batch);
                    ReportProgress(batch.Length);
                }
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
                            if (!string.IsNullOrEmpty(wb))
                                ProcessTrackingData(wb, item);
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
    .Where(d =>
        (d.scanTypeName?.Contains("Nhận hàng") == true ||
         d.scanTypeName?.Contains("Lấy hàng") == true ||
         d.status == "已揽件") &&
        (!string.IsNullOrEmpty(d.staffName) || !string.IsNullOrEmpty(d.scanByName))
    )
    .OrderBy(d =>
    {
        string timeStr = !string.IsNullOrEmpty(d.uploadTime) ? d.uploadTime :
                        (d.scanTime ?? "9999-12-31");
        return DateTime.TryParse(timeStr, out DateTime dt) ? dt : DateTime.MaxValue;
    })
    .FirstOrDefault();

            if (staffNameFisrt != null)
            {
                row.NhanVienNhanHang = staffNameFisrt.staffName ??
                                       staffNameFisrt.scanByName ?? "";
            }

            var giaoLaiGanNhat = details.Where(d => d.scanTypeName != null && d.scanTypeName.Contains("Giao lại hàng"))
        .OrderByDescending(d =>
        {
            string timeStr = d.uploadTime ?? d.scanTime ?? "";
            return DateTime.TryParse(timeStr, out DateTime dt) ? dt : DateTime.MinValue;
        }).FirstOrDefault();

            row.ThoiGianYeuCauPhatLai = giaoLaiGanNhat?.remark2 ?? "";


            var vanDe = details.Where(d => d.scanTypeName?.Contains("vấn đề") == true || d.scanTypeName?.Contains("Kiện vấn đề") == true)
        .OrderByDescending(d =>
        {
            string timeStr = d.uploadTime ?? d.scanTime ?? "";
            return DateTime.TryParse(timeStr, out DateTime dt) ? dt : DateTime.MinValue;
        }).FirstOrDefault();
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
                if (type == "Kiểm tra hàng tồn kho" || type.Contains("Lịch sử cuộc gọi") || type.Contains("cuộc gọi-phát"))
                    continue;

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
                try
                {
                    await CallOrderDetailAsync(waybill);
                    ReportProgress(1);
                }
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

            if (string.IsNullOrEmpty(row.NhanVienNhanHang))
                row.NhanVienNhanHang = info.staffName ?? "";

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

        // ====================== API CHUYỂN HOÀN - pringListPage ======================
        public async Task LoadRebackPrintInfoAsync(List<string> waybills)
        {
            if (waybills.Count == 0) return;

            var batches = waybills.Chunk(40).ToList();
            await Parallel.ForEachAsync(batches, new ParallelOptions { MaxDegreeOfParallelism = 8 }, async (batch, ct) =>
            {
                await _semaphore.WaitAsync(ct);
                try
                {
                    await CallRebackPrintBatchAsync(batch);
                }
                finally { _semaphore.Release(); }
            });
        }

        private async Task CallRebackPrintBatchAsync(IEnumerable<string> batch)
        {
            var url = "https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/pringListPage";
            var payload = new { waybillNoList = batch.ToList(), page = 1, size = batch.Count() };

            var request = new HttpRequestMessage(HttpMethod.Post, url)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            request.Headers.Add("authToken", Main.CapturedAuthToken);
            request.Headers.Add("lang", "VN");

            try
            {
                var response = await _httpClient.SendAsync(request);
                var json = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    var result = JsonSerializer.Deserialize<RebackPrintResponse>(json,
                        new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                    if (result?.succ == true && result.data?.list != null)
                    {
                        foreach (var item in result.data.list)
                        {
                            if (_rowDict.TryGetValue(item.waybillNo ?? "", out var row))
                            {
                                row.PrintCount = item.printCount;
                                row.NewTerminalDispatchCode = item.newTerminalDispatchCode ?? "";

                                row.RebackStatus = item.statusName
                                    ?? item.status
                                    ?? item.auditStatus
                                    ?? "Chưa duyệt";

                                var inScan = item.details?.FirstOrDefault(d =>
                                    d.scanTypeName?.Contains("In đơn chuyển hoàn") == true
                                    || d.scanTypeName?.Contains("退件扫描") == true);

                                row.InHoanScanTime = inScan?.scanTime ?? "";
                            }
                        }
                    }
                }
            }
            catch { }
        }
        private void LoadAllDataToGrid()
        {
            if (_allRows.Count == 0) return;

            _dataGrid.SuspendLayout();
            _displayTable.BeginLoadData();
            try
            {
                _displayTable.Rows.Clear();

                var newRows = new object[_allRows.Count][];
                for (int i = 0; i < _allRows.Count; i++)
                {
                    var r = _allRows[i];
                    newRows[i] = new object[]
                    {
                        r.WaybillNo, r.TrangThaiHienTai, r.ThaoTacCuoi, r.ThoiGianThaoTac,
                        r.ThoiGianYeuCauPhatLai, r.NhanVienKienVanDe, r.NguyenNhanKienVanDe,
                        r.BuuCucThaoTac, r.NguoiThaoTac, r.DauChuyenHoan,
                        r.DiaChiNhanHang, r.Phuong, r.NoiDungHangHoa, r.CODThucTe,
                        r.PTTT, r.NhanVienNhanHang, r.DiaChiLayHang, r.ThoiGianNhanHang,
                        r.TenNguoiGui, r.TrongLuong, r.MaDoan1, r.MaDoan2, r.MaDoan3
                    };
                }
                for (int i = 0; i < newRows.Length; i++)
                {
                    _displayTable.Rows.Add(newRows[i]);
                }
            }
            finally
            {
                _displayTable.EndLoadData();
                _dataGrid.ResumeLayout();
                _dataGrid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
        }

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
            catch (Exception ex)
            {
                return $"Lỗi hiển thị lịch sử: {ex.Message}";
            }
        }
        private string CleanType(string type)
        {
            if (string.IsNullOrEmpty(type)) return type;
            return type

                // VIỆT SUB

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
                FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd/MM/yyyy}.xlsx",
                InitialDirectory = _exportFolder
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Trạng thái hiện tại");
                var exportTable = _displayTable.Copy();

                var table = ws.Cell(1, 1).InsertTable(exportTable, "TrackingTable", true);


                // === FORMAT HEADER ===
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


        /// <summary>
        /// 
        ///
        //

        ////
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
                FileName = $"Trạng thái vận đơn__{DateTime.Now:HH-mm-ss  dd/MM/yyyy}.xlsx",
                InitialDirectory = _exportFolder
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var wb = new XLWorkbook();

                // ==================== SHEET 1: PHÁT HÀNG ====================
                var phatHangRows = _allRows.Where(r => r.DauChuyenHoan != "Có").ToList();
                if (phatHangRows.Any())
                {
                    var ws = wb.Worksheets.Add("PHÁT HÀNG");
                    var dtPhat = CreatePhatHangDataTable(phatHangRows);
                    var table = ws.Cell(1, 1).InsertTable(dtPhat, "PhatHangTable", true);
                    ApplyHeaderStyle(ws, table);
                }

                // ==================== SHEET 2: HOÀN PHÁT ====================
                var hoanPhatRows = _allRows.Where(r => r.DauChuyenHoan == "Có").ToList();
                if (hoanPhatRows.Any())
                {
                    var ws = wb.Worksheets.Add("HOÀN PHÁT");
                    var dtHoan = CreateHoanPhatDataTable(hoanPhatRows);
                    var table = ws.Cell(1, 1).InsertTable(dtHoan, "HoanPhatTable", true);
                    ApplyHeaderStyle(ws, table);
                }

                wb.SaveAs(sfd.FileName);

                DialogResult result = MessageBox.Show($"Xuất file thành công!\n\nFile: {sfd.FileName}\n\nMở file ngay?",
                                                      "Thành công", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi export: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                dt.Rows.Add(
                    r.WaybillNo,
                    r.MaDoan3,
                    r.DiaChiNhanHang,
                    r.Phuong,
                    r.TrangThaiHienTai,
                    r.ThaoTacCuoi,
                    r.ThoiGianThaoTac,
                    r.NoiDungHangHoa,
                    r.CODThucTe,
                    r.TrongLuong,
                    r.MaDoan1,
                    r.MaDoan2
                );
            }
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

            for (int i = 0; i < rows.Count; i++)
            {
                var r = rows[i];
                dt.Rows.Add(
                    r.WaybillNo,
                    r.NhanVienNhanHang,
                    r.TenNguoiGui,
                    r.DiaChiLayHang,
                    r.NoiDungHangHoa,
                    r.CODThucTe,
                    r.PTTT,
                    r.TrongLuong
                );
            }
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

        public List<TrackingRow> GetAllRows()
        {
            _allRows.ToList();
            return _allRows;
        }

        public void ClearData()
        {
            _allRows.Clear();
            _rowDict.Clear();
            _displayTable.Rows.Clear();
            if (_progressBar != null) _progressBar.Visible = false;
            if (_lblPercent != null) _lblPercent.Visible = false;
        }

     

        /////////// PRINT PREVIEW////////////
        public async Task<string> GetPrintPdfUrlAsync(List<string> selectedWaybills, int printType, int applyTypeCode)
        {
            try
            {
                string token = Main.CapturedAuthToken; // Lấy token hiện tại của hệ thống
                if (string.IsNullOrEmpty(token)) throw new Exception("Chưa có Auth Token!");

                _httpClient.DefaultRequestHeaders.Clear();
                _httpClient.DefaultRequestHeaders.Add("Authorization", token);
                // Thêm các header cần thiết khác của J&T nếu có (vd: x-app-id,...)

                // --- BƯỚC 1: GỌI API pringListPage ---
                var payload1 = new
                {
                    current = 1,
                    size = 20,
                    pringFlag = 0,
                    applyNetworkId = 5165,
                    waybillIds = selectedWaybills, // Danh sách truyền vào giữ nguyên thứ tự
                    applyTimeFrom = DateTime.Now.ToString("yyyy-MM-dd 00:00:00"),
                    applyTimeTo = DateTime.Now.ToString("yyyy-MM-dd 23:59:59"),
                    pringType = printType,
                    countryId = "1"
                };

                var content1 = new StringContent(JsonSerializer.Serialize(payload1), Encoding.UTF8, "application/json");
                var response1 = await _httpClient.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/pringListPage", content1);
                response1.EnsureSuccessStatusCode();
                // (Tùy chọn) Đọc kết quả response1 nếu hệ thống yêu cầu lấy dữ liệu từ đây để truyền sang bước 2

                // --- BƯỚC 2: GỌI API printWaybill ĐỂ LẤY LINK PDF ---
                var payload2 = new
                {
                    waybillIds = selectedWaybills, // Tiếp tục truyền đúng thứ tự
                    applyTypeCode = applyTypeCode,
                    countryId = "1"
                };

                var content2 = new StringContent(JsonSerializer.Serialize(payload2), Encoding.UTF8, "application/json");
                var response2 = await _httpClient.PostAsync("https://jmsgw.jtexpress.vn/operatingplatform/rebackTransferExpress/printWaybill", content2);
                response2.EnsureSuccessStatusCode();

                string jsonResponse2 = await response2.Content.ReadAsStringAsync();

                // Parse JSON để lấy link PDF trả về. 
                // LƯU Ý: Cấu trúc parse bên dưới giả định link PDF nằm trong thuộc tính "data". 
                // Bạn cần điều chỉnh lại JsonDocument.Parse nếu API trả về cấu trúc khác.
                using (JsonDocument doc = JsonDocument.Parse(jsonResponse2))
                {
                    var root = doc.RootElement;
                    if (root.TryGetProperty("data", out JsonElement dataElement))
                    {
                        // Giả sử API trả về chuỗi URL trực tiếp trong "data"
                        return dataElement.GetString();
                    }
                }

                throw new Exception("Không tìm thấy link PDF trong phản hồi của API.");
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void UpdatePringListData(List<PringListRecord> printRecords)
        {
            foreach (var record in printRecords)
            {
                // Tìm xem mã vận đơn này có trong danh sách gốc không
                if (_rowDict.TryGetValue(record.WaybillNo, out TrackingRow r))
                {
                    r.RebackStatus = record.StatusName;
                    r.PrintCount = record.PrintCount;
                    r.NewTerminalDispatchCode = record.NewTerminalDispatchCode;
                    // r.InHoanScanTime = DateTime.Now.ToString("HH:mm:ss"); // Có thể tự gắn giờ nếu cần
                }
            }
        }

        private List<PringListRecord> ExtractPringListData(string jsonString)
{
    List<PringListRecord> result = new List<PringListRecord>();
    try
    {
        using (JsonDocument doc = JsonDocument.Parse(jsonString))
        {
            var root = doc.RootElement;
            
            // Vào property "data" -> "records"
            if (root.TryGetProperty("data", out JsonElement dataElement) && dataElement.ValueKind != JsonValueKind.Null &&
                dataElement.TryGetProperty("records", out JsonElement recordsElement) && recordsElement.ValueKind == JsonValueKind.Array)
            {
                // Vòng lặp lấy từng vận đơn trong danh sách
                foreach (var record in recordsElement.EnumerateArray())
                {
                    var item = new PringListRecord();
                    
                    if (record.TryGetProperty("waybillNo", out JsonElement wb) && wb.ValueKind != JsonValueKind.Null) 
                        item.WaybillNo = wb.GetString();
                        
                    if (record.TryGetProperty("status", out JsonElement st) && st.ValueKind != JsonValueKind.Null) 
                        item.Status = st.GetInt32();
                        
                    if (record.TryGetProperty("statusName", out JsonElement stn) && stn.ValueKind != JsonValueKind.Null) 
                        item.StatusName = stn.GetString();
                        
                    if (record.TryGetProperty("newTerminalDispatchCode", out JsonElement ntdc) && ntdc.ValueKind != JsonValueKind.Null) 
                        item.NewTerminalDispatchCode = ntdc.GetString();
                        
                    if (record.TryGetProperty("printCount", out JsonElement pc) && pc.ValueKind != JsonValueKind.Null) 
                        item.PrintCount = pc.GetInt32();

                    result.Add(item);
                }
            }
        }
    }
    catch (Exception ex)
    {}
    return result;
}

    }



    public class PringListRecord
    {
        public string WaybillNo { get; set; }
        public int Status { get; set; }
        public string StatusName { get; set; }
        public string NewTerminalDispatchCode { get; set; }
        public int PrintCount { get; set; }
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