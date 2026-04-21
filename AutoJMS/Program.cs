using System;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoJMS
{
    static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private const int SW_RESTORE = 9;
        private const string MUTEX_NAME = "Global\\AutoJMS_SingleInstance_v1.0_DatDzvl";

        private static string cacheFilePath = Path.Combine(Application.StartupPath, "license.dat");
        [STAThread]
        static void Main()
        {
            // Cấu hình bắt buộc phải gọi ĐẦU TIÊN
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool createdNew;
            using (Mutex mutex = new Mutex(true, MUTEX_NAME, out createdNew))
            {
                if (!createdNew)
                {
                    BringExistingInstanceToFront();
                    return;
                }

                if (Environment.OSVersion.Version.Major >= 6)
                    SetProcessDPIAware();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                try
                {
                    GoogleSheetService.InitService();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi kết nối hệ thống... " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                string myHwid = GetHWID();
                string savedKey = ReadLocalCache(myHwid);
                bool isAuthorized = false;

                if (!CheckInternetConnection())
                {
                    MessageBox.Show("Không thể kết nối !", "Lỗi kết nối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!string.IsNullOrEmpty(savedKey))
                {
                    var checkResult = Task.Run(() => FirebaseLicenseService.CheckLicenseAsync(savedKey, myHwid)).GetAwaiter().GetResult();

                    if (checkResult.success)
                    {
                        isAuthorized = true;
                        SaveLocalCache(savedKey, myHwid);
                    }
                    else
                    {
                        DeleteLocalCache();
                        MessageBox.Show("Bản quyền không hợp lệ!\n\nLý do: " + checkResult.message, "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                while (!isAuthorized)
                {
                    string userInputKey = "";

                    using (frmLogin frm = new frmLogin(myHwid))
                    {
                        DialogResult result = frm.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            userInputKey = frm.EnteredKey;
                        }
                        else
                        {
                            return;
                        }
                    }

                    var activateResult = Task.Run(() => FirebaseLicenseService.CheckLicenseAsync(userInputKey, myHwid)).GetAwaiter().GetResult();

                    if (activateResult.success)
                    {
                        MessageBox.Show("Kích hoạt thành công!\n" + activateResult.message, "Enjoy cái moment này!!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        SaveLocalCache(userInputKey, myHwid);
                        isAuthorized = true;
                    }
                    else
                    {
                        MessageBox.Show("Lỗi kích hoạt:\n" + activateResult.message, "Từ chối", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                Application.Run(new Main());
            }
        }

        private static bool CheckInternetConnection()
        {
            try
            {
                using (var client = new System.Net.WebClient())
                using (client.OpenRead("http://clients3.google.com/generate_204"))
                {
                    return true;
                }
            }
            catch { return false; }
        }

        public static void SaveLocalCache(string licenseKey, string hwid)
        {
            try
            {
                string rawData = $"{licenseKey}||{hwid}";
                string encryptedData = Convert.ToBase64String(Encoding.UTF8.GetBytes(rawData));
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
                string rawData = Encoding.UTF8.GetString(Convert.FromBase64String(encryptedData));

                string[] parts = rawData.Split(new string[] { "||" }, StringSplitOptions.None);
                if (parts.Length == 2 && parts[1] == currentHwid) return parts[0];
                return null;
            }
            catch { return null; }
        }

        public static void DeleteLocalCache()
        {
            try { if (File.Exists(cacheFilePath)) File.Delete(cacheFilePath); } catch { }
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

        private static string GetHWID()
        {
            string cpuId = GetCpuId();
            string motherboardId = GetMotherboardId();
            return GetMD5Hash($"{cpuId}-{motherboardId}");
        }

        private static string GetCpuId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("Select ProcessorId From Win32_Processor"))
                {
                    foreach (ManagementObject mObject in searcher.Get())
                    {
                        return mObject["ProcessorId"]?.ToString().Trim();
                    }
                }
            }
            catch { }
            return "UNKNOWN_CPU";
        }

        private static string GetMotherboardId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard"))
                {
                    foreach (ManagementObject mObject in searcher.Get())
                    {
                        return mObject["SerialNumber"]?.ToString().Trim();
                    }
                }
            }
            catch { }
            return "UNKNOWN_MB";
        }

        private static string GetMD5Hash(string input)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(Encoding.ASCII.GetBytes(input));
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++) sb.Append(hashBytes[i].ToString("X2"));
                return sb.ToString();
            }
        }
    }
}