using System;
using System.IO;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class AppLogger
    {
        // Thư mục chứa file log (Nằm cạnh file .exe)
        private static readonly string LogDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
        
        // Khóa đồng bộ chống đụng độ khi nhiều luồng cùng ghi log một lúc
        private static readonly object _lockObj = new object();
        
        // Cấu hình số ngày giữ lại log
        private const int KeepLogsDays = 15; 

        static AppLogger()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                    Directory.CreateDirectory(LogDirectory);

                CleanupOldLogs();
            }
            catch { /* Lặng lẽ bỏ qua nếu không có quyền tạo thư mục */ }
        }

        public static void Info(string message) => WriteLog("INFO", message);
        public static void Warning(string message) => WriteLog("WARN", message);
        
        public static void Error(string message, Exception ex = null)
        {
            string fullMessage = message;
            if (ex != null)
                fullMessage += $"\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
                
            WriteLog("ERROR", fullMessage);
        }

        public static void Fatal(string message, Exception ex = null)
        {
            string fullMessage = message;
            if (ex != null)
                fullMessage += $"\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
                
            WriteLog("FATAL", fullMessage);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                // File log tự động cắt theo ngày (VD: AppLog_20260508.txt)
                string logFile = Path.Combine(LogDirectory, $"AppLog_{DateTime.Now:yyyyMMdd}.txt");
                
                // Cấu trúc chuẩn 1 dòng log
                string logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";

                lock (_lockObj) // Xếp hàng các luồng để ghi file an toàn
                {
                    File.AppendAllText(logFile, logEntry);
                }
            }
            catch
            {
                // Nguyên tắc tối thượng của Logger: Bản thân bộ ghi log lỗi thì không được làm sập App
            }
        }

        private static void CleanupOldLogs()
        {
            // Chạy dọn rác ở luồng ngầm, không làm chậm quá trình khởi động
            Task.Run(() =>
            {
                try
                {
                    var files = new DirectoryInfo(LogDirectory).GetFiles("AppLog_*.txt");
                    DateTime cutoff = DateTime.Now.Date.AddDays(-KeepLogsDays);

                    foreach (var file in files)
                    {
                        if (file.CreationTime.Date < cutoff)
                        {
                            try { file.Delete(); } catch { }
                        }
                    }
                }
                catch { }
            });
        }
    }
}