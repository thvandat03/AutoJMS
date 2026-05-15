using Postgrest.Attributes;
using Postgrest.Models;
using System;
using System.Reflection;

namespace AutoJMS.Data
{
    [Obfuscation(Exclude = true, ApplyToMembers = true)]
    [Table("waybills")]
    public class WaybillDbModel : BaseModel
    {
        [PrimaryKey("waybill_no", false)] public string WaybillNo { get; set; }
        [Column("trang_thai_hien_tai")] public string TrangThaiHienTai { get; set; }
        [Column("thao_tac_cuoi")] public string ThaoTacCuoi { get; set; }
        [Column("thoi_gian_thao_tac")] public string ThoiGianThaoTac { get; set; }
        [Column("thoi_gian_yeu_cau_phat_lai")] public string ThoiGianYeuCauPhatLai { get; set; }
        [Column("nhan_vien_kien_van_de")] public string NhanVienKienVanDe { get; set; }
        [Column("nguyen_nhan_kien_van_de")] public string NguyenNhanKienVanDe { get; set; }
        [Column("buu_cuc_thao_tac")] public string BuuCucThaoTac { get; set; }
        [Column("nguoi_thao_tac")] public string NguoiThaoTac { get; set; }
        [Column("dau_chuyen_hoan")] public string DauChuyenHoan { get; set; }
        [Column("dia_chi_nhan_hang")] public string DiaChiNhanHang { get; set; }
        [Column("phuong")] public string Phuong { get; set; }
        [Column("noi_dung_hang_hoa")] public string NoiDungHangHoa { get; set; }
        [Column("cod_thuc_te")] public string CODThucTe { get; set; }
        [Column("pttt")] public string PTTT { get; set; }
        [Column("nhan_vien_nhan_hang")] public string NhanVienNhanHang { get; set; }
        [Column("dia_chi_lay_hang")] public string DiaChiLayHang { get; set; }
        [Column("thoi_gian_nhan_hang")] public string ThoiGianNhanHang { get; set; }
        [Column("ten_nguoi_gui")] public string TenNguoiGui { get; set; }
        [Column("trong_luong")] public string TrongLuong { get; set; }
        [Column("ma_doan_full")] public string MaDoanFull { get; set; }
        [Column("ma_doan_1")] public string MaDoan1 { get; set; }
        [Column("ma_doan_2")] public string MaDoan2 { get; set; }
        [Column("ma_doan_3")] public string MaDoan3 { get; set; }
        [Column("reback_status")] public string RebackStatus { get; set; }
        [Column("print_count")] public int PrintCount { get; set; }
        [Column("new_terminal_dispatch_code")] public string NewTerminalDispatchCode { get; set; }
        [Column("in_hoan_scan_time")] public string InHoanScanTime { get; set; }

        // --- CÁC BIẾN KIỂM SOÁT AUTO TRACKING ---
        [Column("is_active")] public bool IsActive { get; set; } = true;
        [Column("tracking_interval_mins")] public int TrackingIntervalMins { get; set; } = 30;
        [Column("last_tracked_at")] public DateTime LastTrackedAt { get; set; }
        [Column("next_track_at")] public DateTime NextTrackAt { get; set; }
    }
}