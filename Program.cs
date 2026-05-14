#nullable enable
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AutoJMS
{
    static class Program
    {
        [DllImport("user32.dll")] private static extern bool SetProcessDPIAware();
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true)] private static extern bool CheckRemoteDebuggerPresent(IntPtr hProcess, ref bool isDebuggerPresent);

        private const int SW_RESTORE = 9;
        private const string MUTEX_NAME = "Global\\AutoJMS_SingleInstance_Commercial";
        private static string cacheFilePath = Path.Combine(Application.StartupPath, "license.dat");

        public static string HWID { get; private set; }
        public static string ExecutableHash { get; private set; }
        private static readonly CancellationTokenSource AppCts = new CancellationTokenSource();

        [STAThread]
        static void Main()
        {
#if !DEBUG
            if (Debugger.IsAttached || IsDebuggerPresent())
                Environment.Exit(0);
#endif

            Application.SetHighDpiMode(HighDpiMode.DpiUnaware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            HWID = GetHWID();
            ExecutableHash = ComputeExecutableHash();

            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (sender, e) =>
            {
                if (e.Exception is System.Collections.Generic.KeyNotFoundException &&
                    e.Exception.StackTrace != null &&
                    e.Exception.StackTrace.Contains("KeyboardToolTipStateMachine"))
                    return;

                AppLogger.Error("Lỗi hệ thống UI chưa được xử lý", e.Exception);
                MessageBox.Show($"Lỗi hệ thống:\n{e.Exception.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };

            Application.ApplicationExit += (s, e) =>
            {
                AppCts.Cancel();
            };

            bool createdNew;
            using (Mutex mutex = new Mutex(true, MUTEX_NAME, out createdNew))
            {
                if (!createdNew)
                {
                    BringExistingInstanceToFront();
                    return;
                }

                if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();

                string myHwid = HWID;
                try { AppConfig.LoadBootstrap(myHwid); } catch { }

                // ==========================================
                // KHỞI ĐỘNG OFFLINE-FIRST THÔNG MINH
                // ==========================================
                string savedKey = ReadLocalCache(myHwid);
                bool isAuthorized = false;
                string activeToken = "";
                string activeKey = savedKey ?? "";

                if (!string.IsNullOrEmpty(savedKey))
                {
                    // Kiểm tra mạng nhanh bằng ping phần cứng
                    bool online = NetworkInterface.GetIsNetworkAvailable();

                    if (online)
                    {
                        try
                        {
                            var checkResult = Task.Run(() =>
                                LicenseApiService.VerifyLicenseSecureAsync(savedKey, myHwid))
                                .GetAwaiter().GetResult();

                            if (checkResult.success && !string.IsNullOrEmpty(checkResult.token))
                            {
                                activeToken = checkResult.token;
                                activeKey = savedKey;
                                isAuthorized = true;
                                SaveLocalCache(savedKey, myHwid);
                            }
                            else
                            {
                                // Lỗi thực sự (Key sai, hết hạn, bị thu hồi)
                                DeleteLocalCache();
                                AppLogger.Error("Key bị từ chối: " + checkResult.message);
                            }
                        }
                        catch (Exception ex)
                        {
                            // Lỗi mạng ngay khi verify, giữ key cũ, chạy offline
                            isAuthorized = true;
                            activeToken = "";
                            AppLogger.Error("Mất kết nối trong lúc xác thực.", ex);
                        }
                    }
                    else
                    {
                        // Offline hoàn toàn, dùng key đã lưu
                        isAuthorized = true;
                        activeToken = "";
                        AppLogger.Warning("Khởi động offline với key đã lưu.");
                    }
                }

                // ==========================================
                // YÊU CẦU NHẬP KEY (NẾU CHƯA CÓ HOẶC BỊ THU HỒI)
                // ==========================================
                while (!isAuthorized)
                {
                    string userInputKey = "";
                    using (frmLogin frm = new frmLogin(myHwid))
                    {
                        if (frm.ShowDialog() == DialogResult.OK) userInputKey = frm.EnteredKey;
                        else return;
                    }

                    // Khi nhập key mới, bắt buộc online
                    var activateResult = Task.Run(() =>
                        LicenseApiService.VerifyLicenseSecureAsync(userInputKey, myHwid))
                        .GetAwaiter().GetResult();

                    if (activateResult.success && !string.IsNullOrEmpty(activateResult.token))
                    {
                        MessageBox.Show(activateResult.message, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        activeToken = activateResult.token;
                        activeKey = userInputKey;
                        SaveLocalCache(userInputKey, myHwid);
                        isAuthorized = true;
                    }
                    else
                    {
                        MessageBox.Show(activateResult.message, "Từ chối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // ==========================================
                // KHỞI ĐỘNG CÁC DỊCH VỤ NỀN
                // ==========================================
                try
                {
                    GoogleSheetService.ResetService();
                    GoogleSheetService.InitService();
                }
                catch (Exception ex)
                {
                    AppLogger.Error("Không thể kết nối Google Sheet lúc khởi động.", ex);
                    MessageBox.Show("Không thể kết nối Google Sheet.\n" + ex.Message, "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var networkMonitor = new uiControlService();
                _ = networkMonitor.StartAsync(AppCts.Token);

                var heartbeat = new LicenseApiService.HeartbeatSupervisor(
                    activeKey,
                    myHwid,
                    activeToken,
                    token => { /* Cập nhật Token nội bộ nếu cần */ },
                    warning => { AppLogger.Info(warning); }
                );
                _ = heartbeat.StartAsync(AppCts.Token);

                Application.Run(new Main());
            }
        }

        // --- CÁC HÀM TIỆN ÍCH ---
        private static string ComputeExecutableHash()
        {
            try
            {
                string exePath = Application.ExecutablePath;
                using (var sha256 = SHA256.Create())
                {
                    using (var stream = new FileStream(exePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        byte[] hash = sha256.ComputeHash(stream);
                        return BitConverter.ToString(hash).Replace("-", "").ToLower();
                    }
                }
            }
            catch { return "HASH_ERROR"; }
        }

        private static bool IsDebuggerPresent()
        {
            try
            {
                bool isDebuggerPresent = false;
                CheckRemoteDebuggerPresent(Process.GetCurrentProcess().Handle, ref isDebuggerPresent);
                return isDebuggerPresent;
            }
            catch { return false; }
        }

        private static void BringExistingInstanceToFront()
        {
            var current = Process.GetCurrentProcess();
            foreach (var p in Process.GetProcessesByName(current.ProcessName))
            {
                if (p.Id != current.Id && p.MainWindowHandle != IntPtr.Zero)
                {
                    ShowWindow(p.MainWindowHandle, SW_RESTORE);
                    SetForegroundWindow(p.MainWindowHandle);
                    break;
                }
            }
        }

        // --- LOCAL CACHE DPAPI ---
        public static void SaveLocalCache(string licenseKey, string hwid)
        {
            try
            {
                string rawData = $"{licenseKey}||{hwid}";
                string encryptedData = SecureConfigCrypto.ProtectString(rawData, BuildCacheSecret(hwid));
                File.WriteAllText(cacheFilePath, encryptedData);
            }
            catch { }
        }

        private static string ReadLocalCache(string currentHwid)
        {
            if (!File.Exists(cacheFilePath)) return null;
            try
            {
                string encryptedData = File.ReadAllText(cacheFilePath);
                string rawData = SecureConfigCrypto.UnprotectString(encryptedData, BuildCacheSecret(currentHwid));
                string[] parts = rawData.Split(new[] { "||" }, StringSplitOptions.None);
                if (parts.Length == 2 && parts[1] == currentHwid) return parts[0];
                return null;
            }
            catch { return null; }
        }

        public static void DeleteLocalCache() { try { if (File.Exists(cacheFilePath)) File.Delete(cacheFilePath); } catch { } }

        private static string BuildCacheSecret(string hwid) => $"{Environment.MachineName}|{Environment.UserName}|{hwid}|AutoJMS";

        // --- HWID TAM GIÁC VÀNG ---
        private static string GetHWID()
        {
            string smbiosUUID = GetSystemUUID();
            string physicalDisk = GetPhysicalDiskSerial();
            string machineGuid = GetMachineGuid();
            return ComputeSha256($"{smbiosUUID}-{physicalDisk}-{machineGuid}");
        }

        private static string GetSystemUUID()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct"))
                {
                    foreach (ManagementObject mObject in searcher.Get())
                    {
                        string uuid = mObject["UUID"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(uuid) && uuid != "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF") return uuid;
                    }
                }
            }
            catch { }
            return "UNKNOWN_UUID";
        }

        private static string GetPhysicalDiskSerial()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_DiskDrive WHERE MediaType='Fixed hard disk media'"))
                {
                    foreach (ManagementObject mObject in searcher.Get())
                    {
                        string serial = mObject["SerialNumber"]?.ToString().Trim();
                        if (!string.IsNullOrEmpty(serial)) return serial.Replace(" ", "").Replace(".", "");
                    }
                }
            }
            catch { }
            return "UNKNOWN_PHYSICAL_DISK";
        }

        private static string GetMachineGuid()
        {
            try
            {
                using (RegistryKey localMachineX64View = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                {
                    using (RegistryKey rk = localMachineX64View.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                    {
                        if (rk != null)
                        {
                            object guid = rk.GetValue("MachineGuid");
                            if (guid != null) return guid.ToString();
                        }
                    }
                }
            }
            catch { }
            return "UNKNOWN_GUID";
        }

        private static string ComputeSha256(string rawData)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(rawData));
                return BitConverter.ToString(hash).Replace("-", "").ToLower();
            }
        }
    }

    // ==========================================
    // MÃ HÓA NỘI BỘ (Giữ nguyên)
    // ==========================================
    internal static class SecureConfigCrypto
    {
        private const int SaltSize = 16;
        private const int IvSize = 16;
        private const string AlgorithmName = "AES-CBC-HMACSHA256-MD5-SHA256";

        private sealed class ProtectedPayload
        {
            public int Version { get; set; } = 1;
            public string Algorithm { get; set; } = AlgorithmName;
            public string Salt { get; set; } = "";
            public string IV { get; set; } = "";
            public string CipherText { get; set; } = "";
            public string Hash { get; set; } = "";
        }

        public static string ProtectString(string plaintext, string secret)
        {
            if (plaintext == null) throw new ArgumentNullException(nameof(plaintext));
            if (string.IsNullOrWhiteSpace(secret)) throw new ArgumentException("Secret is required.", nameof(secret));

            byte[] salt = RandomNumberGenerator.GetBytes(SaltSize);
            byte[] iv = RandomNumberGenerator.GetBytes(IvSize);
            byte[] plainBytes = Encoding.UTF8.GetBytes(plaintext);

            byte[] cipherBytes;
            using (Aes aes = Aes.Create())
            {
                aes.Key = DeriveKey(secret, salt, "aes");
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;

                using ICryptoTransform encryptor = aes.CreateEncryptor();
                cipherBytes = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);
            }

            byte[] mac = ComputeHash(secret, salt, iv, cipherBytes);
            var payload = new ProtectedPayload
            {
                Salt = Convert.ToBase64String(salt),
                IV = Convert.ToBase64String(iv),
                CipherText = Convert.ToBase64String(cipherBytes),
                Hash = Convert.ToBase64String(mac)
            };

            return JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
        }

        public static string UnprotectString(string protectedJson, string secret)
        {
            if (string.IsNullOrWhiteSpace(protectedJson)) throw new ArgumentException("Protected payload is required.", nameof(protectedJson));
            if (string.IsNullOrWhiteSpace(secret)) throw new ArgumentException("Secret is required.", nameof(secret));

            ProtectedPayload? payload = JsonSerializer.Deserialize<ProtectedPayload>(protectedJson);
            if (payload == null || payload.Version != 1 || !string.Equals(payload.Algorithm, AlgorithmName, StringComparison.Ordinal))
                throw new InvalidDataException("Định dạng config mã hóa không hợp lệ.");

            byte[] salt = Convert.FromBase64String(payload.Salt);
            byte[] iv = Convert.FromBase64String(payload.IV);
            byte[] cipherBytes = Convert.FromBase64String(payload.CipherText);
            byte[] expectedHash = Convert.FromBase64String(payload.Hash);
            byte[] actualHash = ComputeHash(secret, salt, iv, cipherBytes);

            if (!CryptographicOperations.FixedTimeEquals(expectedHash, actualHash))
                throw new CryptographicException("Config mã hóa không còn hợp lệ hoặc sai khóa giải mã.");

            using Aes aes = Aes.Create();
            aes.Key = DeriveKey(secret, salt, "aes");
            aes.IV = iv;
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            using ICryptoTransform decryptor = aes.CreateDecryptor();
            byte[] plainBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length);
            return Encoding.UTF8.GetString(plainBytes);
        }

        private static byte[] ComputeHash(string secret, byte[] salt, byte[] iv, byte[] cipherBytes)
        {
            byte[] macKey = DeriveKey(secret, salt, "sha");
            using var hmac = new HMACSHA256(macKey);
            byte[] header = Encoding.UTF8.GetBytes(AlgorithmName);
            byte[] data = Combine(header, salt, iv, cipherBytes);
            return hmac.ComputeHash(data);
        }

        private static byte[] DeriveKey(string secret, byte[] salt, string purpose)
        {
            byte[] secretBytes = Encoding.UTF8.GetBytes(secret);
            byte[] purposeBytes = Encoding.UTF8.GetBytes(purpose);

            using MD5 md5 = MD5.Create();
            byte[] md5Hash = md5.ComputeHash(secretBytes);

            using SHA256 sha256 = SHA256.Create();
            return sha256.ComputeHash(Combine(purposeBytes, md5Hash, secretBytes, salt));
        }

        private static byte[] Combine(params byte[][] arrays)
        {
            byte[] result = new byte[arrays.Sum(a => a.Length)];
            int offset = 0;
            foreach (byte[] array in arrays)
            {
                Buffer.BlockCopy(array, 0, result, offset, array.Length);
                offset += array.Length;
            }
            return result;
        }
    }
}