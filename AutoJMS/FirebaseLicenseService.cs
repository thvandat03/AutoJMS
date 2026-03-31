using System;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace AutoJMS
{
    public class FirebaseLicenseService
    {
        private const string FIREBASE_URL = "https://keyauthjms-default-rtdb.asia-southeast1.firebasedatabase.app/";
        private const string DATABASE_SECRET = "29m37qye6O3YvtBeWuwDf26cUINV6Zyk7EZVdALm";

        private static readonly HttpClient client = new HttpClient();

        public class LicenseData
        {
            public string hwid { get; set; }
            public string status { get; set; } 
        }

        public static async Task<(bool success, string message)> CheckLicenseAsync(string key, string currentHwid)
        {
            try
            {
                string url = $"{FIREBASE_URL}Licenses/{key}.json?auth={DATABASE_SECRET}";

                HttpResponseMessage response = await client.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                    return (false, "Lỗi kết nối máy chủ");

                string jsonResponse = await response.Content.ReadAsStringAsync();

                if (jsonResponse == "null")
                    return (false, "Key không tồn tại!");

                var license = JsonSerializer.Deserialize<LicenseData>(jsonResponse);

                if (license.status != "active")
                    return (false, "Key không xác định!");

                if (string.IsNullOrEmpty(license.hwid))
                {
                    license.hwid = currentHwid;
                    string putData = JsonSerializer.Serialize(license);
                    var content = new StringContent(putData, Encoding.UTF8, "application/json");
                    await client.PutAsync(url, content); 

                    return (true, "Kích hoạt thành công!");
                }
                else if (license.hwid == currentHwid)
                {
                    return (true, "Xác thực bản quyền thành công!");
                }
                else
                {
                    return (false, "Key này đã được kích hoạt cho một máy tính khác!");
                }
            }
            catch (Exception ex)
            {
                return (false, $"Lỗi xử lý: {ex.Message}");
            }
        }
    }
}