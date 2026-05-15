using System;
using System.IO;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class AppLogger
    {
        private static readonly string LogDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

        private static readonly object _lockObj = new object();
        private const int KeepLogsDays = 30;

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
                // Đã đổi định dạng file thành debug_[Ngày].log để mỗi ngày có 1 file riêng
                string logFile = Path.Combine(LogDirectory, $"debug_{DateTime.Now:yyyyMMdd}.log");

                // Cấu trúc chuẩn 1 dòng log
                string logEntry = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";

                lock (_lockObj) // Xếp hàng các luồng để ghi file an toàn
                {
                    File.AppendAllText(logFile, logEntry);
                }
            }
            catch
            {
      
            }
        }

        private static void CleanupOldLogs()
        {
            Task.Run(() =>
            {
                try
                {
                    // Lọc theo định dạng tên file debug mới
                    var files = new DirectoryInfo(LogDirectory).GetFiles("debug_*.log");
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