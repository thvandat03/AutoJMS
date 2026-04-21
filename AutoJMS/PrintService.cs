using AutoJMS.Data;
using DocumentFormat.OpenXml.VariantTypes;
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

            WaybillTrackingService.EnableDoubleBuffering(_grid);
            _grid.DataSource = _displayTable;

            SetupGrid();
            RebuildTable();
        }

        public void Reset()
        {
            _trackingService.ClearData();
            _currentMode = PrintMode.InHoan;
            _displayTable.Rows.Clear();
            _displayTable.Columns.Clear();
            RebuildTable();
        }

        private void SetupGrid()
        {
            _grid.AllowUserToAddRows = false;
            _grid.AllowUserToDeleteRows = false;
            _grid.EditMode = DataGridViewEditMode.EditOnEnter; 

            _grid.RowHeadersVisible = false;
            _grid.AllowUserToResizeColumns = true;
            _grid.ScrollBars = ScrollBars.Both;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False;

            // Checkbox cột đầu tiên
            if (!_grid.Columns.Contains("Select"))
            {
                var checkCol = new DataGridViewCheckBoxColumn
                {
                    Name = "Select",
                    HeaderText = "",
                    Width = 60,
                    TrueValue = true,
                    FalseValue = false
                };
                _grid.Columns.Insert(0, checkCol);
            }
            DisableSorting();
        }

        //public void UpdatePrintStatus(List<PringListRecord> printRecords)
        //{
        //    // Kiểm tra và thêm cột nếu chưa có (giao diện sẽ tự sinh ra cột tương ứng)
        //    if (!_displayTable.Columns.Contains("Trạng thái duyệt"))
        //        _displayTable.Columns.Add("Trạng thái duyệt", typeof(string));

        //    if (!_displayTable.Columns.Contains("Mã luân chuyển"))
        //        _displayTable.Columns.Add("Mã luân chuyển", typeof(string));

        //    if (!_displayTable.Columns.Contains("Số lần in"))
        //        _displayTable.Columns.Add("Số lần in", typeof(int));

        //    // Quét và cập nhật dữ liệu cho từng dòng
        //    foreach (var record in printRecords)
        //    {
        //        foreach (DataRow row in _displayTable.Rows)
        //        {
        //            // Kiểm tra khớp Mã vận đơn
        //            if (row["Mã vận đơn"] != DBNull.Value && row["Mã vận đơn"].ToString() == record.WaybillNo)
        //            {
        //                row["Trạng thái duyệt"] = record.StatusName ?? "";
        //                row["Mã luân chuyển"] = record.NewTerminalDispatchCode ?? "";
        //                row["Số lần in"] = record.PrintCount;
        //                break; // Tìm thấy thì dừng vòng lặp trong, chuyển sang đơn tiếp theo
        //            }
        //        }
        //    }

        //}

        /// CĂN CHỈNH RIÊNG TỪNG CỘT VÀ KHOÁ DỮ LIỆU CHỈ ĐỌC
        private void SetColumnAlignments()
        {
            foreach (DataGridViewColumn col in _grid.Columns)
            {
                // FIX: Đảm bảo cột Checkbox bấm được, các cột khác thì khoá lại
                if (col.Name == "Select")
                {
                    col.ReadOnly = false;
                    continue;
                }
                else
                {
                    col.ReadOnly = true; // Chỉ đọc dữ liệu, không cho gõ bậy vào grid
                }

                switch (col.Name)
                {
                    case "Mã vận đơn":
                    case "Mã đoạn":
                    case "Thời gian in":
                    case "Người in":
                    case "Thời gian gửi":
                    case "SĐT người nhận":
                    case "Số lượng bản in":
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        break;

                    case "Địa chỉ người nhận":
                    case "Địa chỉ lấy hàng":
                    case "Nội dung hàng hóa":
                    case "Tên người gửi":
                    case "Nhân viên lấy hàng":
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        break;

                    case "Trạng thái":
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                        break;

                    default:
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
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

                var newRows = new object[rows.Count][];
                for (int i = 0; i < rows.Count; i++)
                {
                    var r = rows[i];
                    newRows[i] = _currentMode switch
                    {
                        PrintMode.InHoan => new object[]
                        {
                            r.WaybillNo,
                            r.ThaoTacCuoi,
                            r.RebackStatus,
                            r.PrintCount,
                            r.NewTerminalDispatchCode ?? "",
                            string.IsNullOrEmpty(r.InHoanScanTime) ? "" : r.InHoanScanTime,
                            printerName
                        },
                        PrintMode.InChuyenTiep => new object[]
                        {
                            r.WaybillNo,
                            r.ThaoTacCuoi, 1, "",
                            r.DiaChiNhanHang ?? "",
                            r.MaDoanFull ?? "", now, printerName
                        },
                        PrintMode.InLaiDon => new object[]
                        {
                            r.WaybillNo, 1,
                            r.NoiDungHangHoa ?? "",
                            r.MaDoanFull ?? ""
                        },
                        PrintMode.InReverse => new object[]
                        {
                            r.WaybillNo,
                            r.NhanVienNhanHang ?? "",
                            r.DiaChiLayHang ?? "",
                            r.TenNguoiGui ?? "",
                            r.ThoiGianNhanHang ?? "", 1,
                            r.NoiDungHangHoa ?? "",
                            r.MaDoanFull ?? ""
                        },
                        _ => Array.Empty<object>()
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
                _grid.ResumeLayout();
                _grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
        }

        private void DisableSorting()
        {
            foreach (DataGridViewColumn col in _grid.Columns)
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        public async Task SearchAndLoadAsync(string waybillsText, PrintMode mode = PrintMode.InHoan)
        {
            _currentMode = mode;
            await _trackingService.SearchTrackingAsync(waybillsText, updateMainGrid: false);
            RebuildTable();
        }

        // HÀM ĐƯỢC THÊM LẠI ĐỂ SỬ DỤNG VỚI NÚT "CHỌN TẤT CẢ" BÊN MAIN.CS
        public void SelectAll(bool isChecked)
        {
            if (_grid.Columns.Contains("Select"))
            {
                foreach (DataGridViewRow row in _grid.Rows)
                {
                    row.Cells["Select"].Value = isChecked;
                }
                _grid.EndEdit(); // Ép giao diện lưu trạng thái check ngay lập tức
            }
        }

        public List<string> GetSelectedWaybills()
        {
            var list = new List<string>();
            for (int i = 0; i < _grid.Rows.Count; i++)
            {
                if (_grid.Rows[i].Cells["Select"].Value is true)
                {
                    var wb = _grid.Rows[i].Cells["Mã vận đơn"].Value?.ToString();
                    if (!string.IsNullOrEmpty(wb))
                        list.Add(wb);
                }
            }
            return list;   
        }

        public PrintMode CurrentMode => _currentMode;
    }
}