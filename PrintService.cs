using AutoJMS.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace AutoJMS
{
    public enum PrintMode
    {
        InHoan,
        InChuyenTiep,
        InLaiDon,
        InReverse
    }

    public class PrintService
    {
        private readonly DataGridView _grid;
        private readonly WaybillTrackingService _trackingService;
        private readonly DataTable _displayTable = new DataTable();
        private PrintMode _currentMode = PrintMode.InHoan;

        public PrintService(DataGridView grid, WaybillTrackingService trackingService)
        {
            _grid = grid ?? throw new ArgumentNullException(nameof(grid));
            _trackingService = trackingService ?? throw new ArgumentNullException(nameof(trackingService));

            // THỦ THUẬT 1 CLICK: Click vào bất kỳ điểm nào trong Cell cũng sẽ tự động đổi trạng thái
            _grid.CellClick += (s, e) =>
            {
                if (e.RowIndex >= 0 && _grid.Columns[e.ColumnIndex].Name == "Select")
                {
                    var cell = _grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                    bool isChecked = cell.Value != DBNull.Value && (bool)cell.Value;
                    cell.Value = !isChecked;

                    _grid.EndEdit();
                    UpdateStatsAndVisibility();
                }
            };

            WaybillTrackingService.EnableDoubleBuffering(_grid);
            _grid.DataSource = _displayTable;

            SetupGrid();
            RebuildTable();
        }

        // 1. HÀM MỚI: Xóa sạch trạng thái chọn ngầm của WinForms
        public void ClearGridSelection()
        {
            if (_grid.InvokeRequired)
            {
                _grid.Invoke(new Action(ClearGridSelection));
                return;
            }

            _grid.ClearSelection();
            _grid.CurrentCell = null;

            foreach (DataGridViewRow row in _grid.Rows)
            {
                row.Selected = false;
            }

            // Ép bỏ tick toàn bộ checkbox trong DataTable
            foreach (DataRow row in _displayTable.Rows)
            {
                row["Select"] = false;
            }

            UpdateStatsAndVisibility();
            _grid.Refresh();
        }

        public void Reset()
        {
            _trackingService.ClearData();
            _currentMode = PrintMode.InHoan;
            _displayTable.Rows.Clear();
            _displayTable.Columns.Clear();
            ClearGridSelection(); // Gọi thêm Clear
            RebuildTable();
        }

        public event Action<int, int> OnPrintStatsChanged;

        private void UpdateStatsAndVisibility()
        {
            int total = _displayTable.Rows.Count;
            int selected = 0;

            foreach (DataRow row in _displayTable.Rows)
            {
                if (row["Select"] != DBNull.Value && (bool)row["Select"])
                {
                    selected++;
                }
            }

            if (_grid.InvokeRequired)
                _grid.Invoke(new Action(() => _grid.Visible = total > 0));
            else
                _grid.Visible = total > 0;

            OnPrintStatsChanged?.Invoke(selected, total);
        }

        private void SetupGrid()
        {
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter;
            _grid.RowHeadersVisible = false;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.AutoGenerateColumns = true;

            DisableSorting();
        }

        private void SetColumnAlignments()
        {
            if (_grid.Columns.Count == 0) return;

            foreach (DataGridViewColumn col in _grid.Columns)
            {
                col.ReadOnly = true;

                if (col.Name == "Select")
                {
                    col.HeaderText = "Chọn";
                    col.Width = 50;
                    col.ReadOnly = false; // Đảm bảo cột Select có thể sửa được bằng click
                }

                switch (col.Name)
                {
                    case "Select":
                    case "Mã vận đơn":
                    case "Mã đoạn":
                    case "Số lượng bản in":
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        break;
                    default:
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        break;
                }
            }
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        public void SetMode(PrintMode mode)
        {
            if (_currentMode == mode) return;
            _currentMode = mode;
            RebuildTable();
        }

        private void RebuildTable()
        {
            _grid.SuspendLayout();
            try
            {
                _displayTable.Columns.Clear();
                _displayTable.Rows.Clear();

                _displayTable.Columns.Add("Select", typeof(bool));

                switch (_currentMode)
                {
                    case PrintMode.InHoan:
                        _displayTable.Columns.Add("Mã vận đơn", typeof(string));
                        _displayTable.Columns.Add("Trạng thái", typeof(string));
                        _displayTable.Columns.Add("Trạng thái duyệt", typeof(string));
                        _displayTable.Columns.Add("Số lượng bản in", typeof(int));
                        _displayTable.Columns.Add("Mã đoạn", typeof(string));
                        _displayTable.Columns.Add("Thời gian in", typeof(string));
                        _displayTable.Columns.Add("Người in", typeof(string));
                        break;
                    case PrintMode.InChuyenTiep:
                        _displayTable.Columns.Add("Mã vận đơn", typeof(string));
                        _displayTable.Columns.Add("Trạng thái", typeof(string));
                        _displayTable.Columns.Add("Số lượng bản in", typeof(int));
                        _displayTable.Columns.Add("SĐT người nhận", typeof(string));
                        _displayTable.Columns.Add("Địa chỉ người nhận", typeof(string));
                        _displayTable.Columns.Add("Mã đoạn", typeof(string));
                        _displayTable.Columns.Add("Thời gian in", typeof(string));
                        _displayTable.Columns.Add("Người in", typeof(string));
                        break;
                    case PrintMode.InLaiDon:
                        _displayTable.Columns.Add("Mã vận đơn", typeof(string));
                        _displayTable.Columns.Add("Số lượng bản in", typeof(int));
                        _displayTable.Columns.Add("Nội dung hàng hóa", typeof(string));
                        _displayTable.Columns.Add("Mã đoạn", typeof(string));
                        break;
                    case PrintMode.InReverse:
                        _displayTable.Columns.Add("Mã vận đơn", typeof(string));
                        _displayTable.Columns.Add("Nhân viên lấy hàng", typeof(string));
                        _displayTable.Columns.Add("Địa chỉ lấy hàng", typeof(string));
                        _displayTable.Columns.Add("Tên người gửi", typeof(string));
                        _displayTable.Columns.Add("Thời gian gửi", typeof(string));
                        _displayTable.Columns.Add("Số lượng bản in", typeof(int));
                        _displayTable.Columns.Add("Nội dung hàng hóa", typeof(string));
                        _displayTable.Columns.Add("Mã đoạn", typeof(string));
                        break;
                }
            }
            finally
            {
                _grid.ResumeLayout();
            }

            LoadDataToGrid();
            SetColumnAlignments();
        }

        private void LoadDataToGrid()
        {
            var rows = _trackingService.GetAllRows();
            if (rows.Count == 0) return;

            _grid.SuspendLayout();
            _displayTable.BeginLoadData();
            try
            {
                _displayTable.Rows.Clear();
                string now = DateTime.Now.ToString("dd/MM/yyyy HH:mm");
                string printerName = Environment.UserName;

                foreach (var r in rows)
                {
                    string finalWaybill = r.WaybillNo;
                    bool isKyNhanCPN = (r.TrangThaiHienTai?.Contains("Ký nhận CPN") == true) ||
                                       (r.ThaoTacCuoi?.Contains("Ký nhận CPN") == true);

                    if (isKyNhanCPN && !finalWaybill.EndsWith("-001"))
                    {
                        finalWaybill += "-001";
                    }

                    // 2. CHUYỂN TRUE THÀNH FALSE: Ngừng tự động Tick để chống in nhầm
                    object[] rowData = _currentMode switch
                    {
                        PrintMode.InHoan => new object[] { false, finalWaybill, r.ThaoTacCuoi, r.RebackStatus, r.PrintCount, r.NewTerminalDispatchCode ?? "", string.IsNullOrEmpty(r.InHoanScanTime) ? "" : r.InHoanScanTime, printerName },
                        PrintMode.InChuyenTiep => new object[] { false, finalWaybill, r.ThaoTacCuoi, 1, "", r.DiaChiNhanHang ?? "", r.MaDoanFull ?? "", now, printerName },
                        PrintMode.InLaiDon => new object[] { false, finalWaybill, 1, r.NoiDungHangHoa ?? "", r.MaDoanFull ?? "" },
                        PrintMode.InReverse => new object[] { false, finalWaybill, r.NhanVienNhanHang ?? "", r.DiaChiLayHang ?? "", r.TenNguoiGui ?? "", r.ThoiGianNhanHang ?? "", 1, r.NoiDungHangHoa ?? "", r.MaDoanFull ?? "" },
                        _ => null
                    };

                    if (rowData != null) _displayTable.Rows.Add(rowData);
                }
            }
            finally
            {
                _displayTable.EndLoadData();
                _grid.ResumeLayout();
                _grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                UpdateStatsAndVisibility();
            }
        }

        private void DisableSorting()
        {
            foreach (DataGridViewColumn col in _grid.Columns)
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        // 3. XÓA SẠCH DỮ LIỆU CŨ TRƯỚC KHI TÌM KIẾM
        public async System.Threading.Tasks.Task SearchAndLoadAsync(string waybillsText, PrintMode mode)
        {
            _currentMode = mode;

            // Xóa rỗng Tracking cũ để không bị dính mã cũ
            _trackingService.ClearData();

            await _trackingService.SearchTrackingAsync(waybillsText, updateMainGrid: false);
            RebuildTable();

            // Nếu tìm xong không có kết quả hợp lệ => Clear Selection
            if (_displayTable.Rows.Count == 0)
            {
                ClearGridSelection();
            }
        }

        public void SelectAll(bool isChecked)
        {
            _grid.EndEdit();
            foreach (DataRow row in _displayTable.Rows)
            {
                row["Select"] = isChecked;
            }
            _displayTable.AcceptChanges();
            UpdateStatsAndVisibility();
        }

        // 4. LỌC NGHIÊM NGẶT CÁC MÃ ĐƠN HỢP LỆ VÀ XÓA RÁC
        public List<string> GetSelectedWaybills()
        {
            _grid.EndEdit();
            var list = new List<string>();

            foreach (DataRow row in _displayTable.Rows)
            {
                bool isChecked = row["Select"] != DBNull.Value && Convert.ToBoolean(row["Select"]);
                if (!isChecked) continue;

                string wb = row["Mã vận đơn"]?.ToString()?.Trim();

                // Chặn null, rỗng, hoặc quá ngắn
                if (string.IsNullOrWhiteSpace(wb) || wb.Length < 5) continue;

                // Chỉ nhận A-Z, 0-9 và dấu gạch ngang '-', chặn ký tự rác/đặc biệt
                bool isValid = wb.All(c => char.IsLetterOrDigit(c) || c == '-');
                if (!isValid) continue;

                list.Add(wb);
            }

            if (list.Count == 0)
            {
                ClearGridSelection();
            }

            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        public PrintMode CurrentMode => _currentMode;
    }
}