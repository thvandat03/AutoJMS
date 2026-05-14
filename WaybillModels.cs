using System;
using System.Collections.Generic;

namespace AutoJMS.Data
{
  
    public class WaybillHistoryResponse
    {
        public int code { get; set; }           // 1 = Thành công
        public string msg { get; set; }         // Thông báo lỗi nếu có
        public bool succ { get; set; }          // true/false
        public List<WaybillData> data { get; set; } // Danh sách dữ liệu
    }

    public class WaybillData
    {
        public string billCode { get; set; }
        public string keyword { get; set; }      // Từ khóa tìm kiếm
        public List<WaybillDetail> details { get; set; } // Lịch sử hành trình chi tiết
    }

    public class WaybillDetail
    {
        public string billCode { get; set; }
        public string waybillNo { get; set; }

        // Thời gian thao tác (VD: "2026-02-12 09:00:00")
        public string scanTime { get; set; }

        // Loại thao tác (VD: "Giao lại hàng", "Kiểm tra hàng tồn kho")
        public string scanTypeName { get; set; }

        // Tên bưu cục thao tác (VD: "(LCI)Kim Tân")
        public string scanNetworkName { get; set; }

        // Tên nhân viên/mã nhân viên
        public string scanByName { get; set; }
        public string staffName { get; set; }

        // Nội dung chi tiết dài dòng
        public string waybillTrackingContent { get; set; }


        // Thường chứa LÝ DO (VD: "Khách không nghe máy")
        public string remark1 { get; set; }

        // Thường chứa NGÀY HẸN (VD: "2026-02-14") hoặc Biển số xe
        public string remark2 { get; set; }

        // Ghi chú phụ
        public string remark3 { get; set; }

        // Mã lỗi hệ thống (VD: "PTWTJ...")
        public string remark5 { get; set; }

        public string uploadTime { get; set; }

        public string status { get; set; }
    }
    // ==================== MODEL CHO API getOrderDetail ====================
    public class OrderDetailResponse
    {
        public int code { get; set; }
        public string msg { get; set; }
        public bool succ { get; set; }
        public OrderDetailData data { get; set; }
    }

    public class OrderDetailData
    {
        public string waybillNo { get; set; }
        public OrderDetailInfo details { get; set; }
    }

    public class OrderDetailInfo
    {
        public string pickTime { get; set; }
        public string goodsName { get; set; }
        public string customerName { get; set; }
        public decimal? packageChargeWeight { get; set; }
        public string paymentModeName { get; set; }
        public string senderDetailedAddress { get; set; }
        public string receiverDetailedAddress { get; set; }
        public string staffName { get; set; }
        public string codMoney { get; set; }
        public string terminalDispatchCode { get; set; }
        
        public string destinationName { get; set; }
        public string status { get; set; }           // dùng cho "Dấu chuyển hoàn"
    }


    // ==================== MODEL CHO API pringListPage (In Chuyển Hoàn) ====================
    public class RebackPrintResponse
    {
        public int code { get; set; }
        public string msg { get; set; }
        public bool succ { get; set; }
        public RebackPrintData data { get; set; }
    }

    public class RebackPrintData
    {
        public List<RebackPrintItem> list { get; set; }
    }

    public class RebackPrintItem
    {
        public string waybillNo { get; set; }
        public int printCount { get; set; }
        public string newTerminalDispatchCode { get; set; }
        public string statusName { get; set; }           // tên chính
        public string status { get; set; }               // fallback
        public string auditStatus { get; set; }          // fallback
        public List<RebackScanDetail> details { get; set; }
    }

    public class RebackScanDetail
    {
        public string scanTypeName { get; set; }
        public string scanTime { get; set; }
    }

}