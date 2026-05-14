using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.IdentityModel.Tokens;

namespace AutoJMS
{
    public enum HeartbeatOutcome { Continue, ServerKill, TransientFailure, Fatal }

    public class HeartbeatResult
    {
        public HeartbeatOutcome Outcome { get; }
        public string NewToken { get; }
        public string ErrorMessage { get; }

        public HeartbeatResult(HeartbeatOutcome outcome, string newToken, string errorMessage)
        {
            Outcome = outcome;
            NewToken = newToken;
            ErrorMessage = errorMessage;
        }
    }

    public static class LicenseApiService
    {
        private const string API_VERIFY = "https://autojms-api.onrender.com/api/verify-license";
        private const string API_HEARTBEAT = "https://autojms-api.onrender.com/api/heartbeat";
        private const string JWT_PUBLIC_KEY = @"-----BEGIN PUBLIC KEY-----
MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAtaK8L4eH5kvH9UQRVRsU
rJh3qoizfSmBgLLSc8dnLfICa/uVH6K9d6pxAc+iYkgqcB8LxOr7oRDnVeBKwnZm
O59Wnf/dWIYHG7bx/RZ4qa/RjU/qhTzxz4sxAnzEgH5zD2kkpXZPwisglx1naMLc
bRKz/Rmd/KYHDTgEcNDXB9QlB0vehTalCTFiwMHZCnZKHgFysIBju/4/iLmpE/7Y
ztn/m+C4k0KX03gdTbQIeqwyOX5NxDZ74TTtNiHDiMNGrOuB68+TF6SGBDbHUfc/
II8JJiIgzjDJgzNjOXB5nkyaJ6Twf0Y2TeZqX4sxdZdEWacr/RwuWRccN/NsDZI3
eQIDAQAB
-----END PUBLIC KEY-----";

        private static readonly HttpClient Http = new HttpClient() { Timeout = TimeSpan.FromSeconds(60) };
        public static string CurrentSessionId { get; private set; }

        public static async Task<(bool success, string message, string? token)> VerifyLicenseSecureAsync(
        string licenseKey, string hwid, CancellationToken ct = default)
        {
            try
            {
                var payload = new { licenseKey = licenseKey, hwid = hwid, exeHash = Program.ExecutableHash };
                string json = JsonSerializer.Serialize(payload);

                using var req = new HttpRequestMessage(HttpMethod.Post, API_VERIFY);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                using var res = await Http.SendAsync(req, ct);
                string body = await res.Content.ReadAsStringAsync(ct);

                if (!res.IsSuccessStatusCode)
                {
                    using var errDoc = JsonDocument.Parse(body);
                    string errMsg = errDoc.RootElement.TryGetProperty("error", out var err) ? err.GetString() : "Bị từ chối";
                    return (false, errMsg, null);
                }

                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;

                if (!root.TryGetProperty("payload", out var tokenProp))
                    return (false, "Dữ liệu máy chủ không hợp lệ.", null);

                string token = tokenProp.GetString();
                if (!ValidateJwtToken(token)) return (false, "Token không hợp lệ.", null);

                CurrentSessionId = root.GetProperty("sid").GetString();

                // ĐỌC CẤU HÌNH SHEET TỪ SERVER - LƯU VÀO LOG THAY VÌ HIỆN POPUP
                if (root.TryGetProperty("cfg", out var cfgProp))
                {
                    try
                    {
                        if (cfgProp.TryGetProperty("dataSpreadsheetId", out var sheetIdProp))
                        {
                            string idFromServer = sheetIdProp.GetString();
                            if (!string.IsNullOrWhiteSpace(idFromServer)) AppConfig.Current.DataSpreadsheetId = idFromServer;
                        }
                        if (cfgProp.TryGetProperty("updateXmlUrl", out var updateXmlProp))
                        {
                            string urlFromServer = updateXmlProp.GetString();
                            if (!string.IsNullOrWhiteSpace(urlFromServer)) AppConfig.Current.UpdateXmlUrl = urlFromServer;
                        }
                        AppConfig.SaveCurrent();
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Error("Lỗi đọc cấu hình từ server: " + ex.Message);
                    }
                }
                else
                {
                    AppLogger.Warning("SERVER KHÔNG GỬI TRƯỜNG 'cfg' VỀ CHO C#!");
                }

                return (true, "Kích hoạt thành công", token);
            }
            catch (HttpRequestException ex) { return (false, "Mất kết nối máy chủ.", null); }
            catch (TaskCanceledException ex) { return (false, "Máy chủ phản hồi quá chậm.", null); }
            catch (Exception ex) { return (false, "Lỗi hệ thống: " + ex.Message, null); }
        }

        public static async Task<HeartbeatResult> SendHeartbeatOnceAsync(
            string tokenToUse, string hwid, CancellationToken ct)
        {
            try
            {
                var payload = new { clientHwid = hwid, exeHash = Program.ExecutableHash };
                string json = JsonSerializer.Serialize(payload);

                using var req = new HttpRequestMessage(HttpMethod.Post, API_HEARTBEAT);
                req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", tokenToUse);
                req.Content = new StringContent(json, Encoding.UTF8, "application/json");

                using var res = await Http.SendAsync(req, ct);
                string body = await res.Content.ReadAsStringAsync(ct);

                using var doc = JsonDocument.Parse(body);
                var root = doc.RootElement;
                string action = root.TryGetProperty("action", out var act) ? act.GetString() : "";

                if (action == "kill")
                {
                    return new HeartbeatResult(HeartbeatOutcome.ServerKill, null,
                        root.TryGetProperty("reason", out var r) ? r.GetString() : "Revoked");
                }
                else if (action == "continue")
                {
                    string newToken = root.GetProperty("payload").GetString();
                    if (!ValidateJwtToken(newToken)) return new HeartbeatResult(HeartbeatOutcome.Fatal, null, "Invalid JWT");
                    return new HeartbeatResult(HeartbeatOutcome.Continue, newToken, null);
                }

                if (!res.IsSuccessStatusCode) return new HeartbeatResult(HeartbeatOutcome.Fatal, null, "Token Expired");

                return new HeartbeatResult(HeartbeatOutcome.TransientFailure, null, "Unknown action");
            }
            catch (HttpRequestException) { return new HeartbeatResult(HeartbeatOutcome.TransientFailure, null, "Network error"); }
            catch (TaskCanceledException) { return new HeartbeatResult(HeartbeatOutcome.TransientFailure, null, "Timeout"); }
            catch { return new HeartbeatResult(HeartbeatOutcome.TransientFailure, null, "Unknown error"); }
        }

        private static bool ValidateJwtToken(string token)
        {
            try
            {
                string cleanToken = token.Trim().Replace("\"", "");
                RSA rsa = RSA.Create();
                rsa.ImportFromPem(JWT_PUBLIC_KEY.ToCharArray());

                var validationParameters = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidIssuer = "autojms-license-server",
                    ValidateAudience = true,
                    ValidAudience = "autojms-desktop-client",
                    ValidateLifetime = true,
                    ClockSkew = TimeSpan.FromMinutes(2),
                    ValidateIssuerSigningKey = true,
                    IssuerSigningKey = new RsaSecurityKey(rsa) { KeyId = "accessKey" }
                };

                var handler = new JwtSecurityTokenHandler();
                handler.ValidateToken(cleanToken, validationParameters, out SecurityToken validatedToken);
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ==========================================
        // VÒNG LẶP QUẢN LÝ NHỊP TIM ĐÃ NÂNG CẤP
        // ==========================================
        public sealed class HeartbeatSupervisor
        {
            private readonly string _licenseKey;
            private readonly string _deviceId;
            private string _currentToken;
            private readonly Action<string> _onTokenUpdate;
            private readonly Action<string> _onWarning;
            private readonly TimeSpan _interval = TimeSpan.FromMinutes(2);
            private int _fatalRetryCount = 0;

            public HeartbeatSupervisor(string licenseKey, string deviceId, string initialToken, Action<string> onTokenUpdate, Action<string> onWarning)
            {
                _licenseKey = licenseKey;
                _deviceId = deviceId;
                _currentToken = initialToken;
                _onTokenUpdate = onTokenUpdate;
                _onWarning = onWarning;
            }

            public async Task StartAsync(CancellationToken ct = default)
            {
                // Nếu không có token ban đầu (offline từ lúc khởi động), thử recover ngay
                if (string.IsNullOrEmpty(_currentToken))
                {
                    _onWarning?.Invoke("Đang thử kết nối đến máy chủ...");
                    var recoverResult = await LicenseApiService.VerifyLicenseSecureAsync(_licenseKey, _deviceId, ct);
                    if (recoverResult.success && !string.IsNullOrEmpty(recoverResult.token))
                    {
                        _currentToken = recoverResult.token;
                        _onTokenUpdate(_currentToken);
                        _onWarning?.Invoke("✅ Đã kết nối!");
                    }
                    else
                    {
                        _onWarning?.Invoke("⏳ Chưa có kết nối, sẽ thử lại sau.");
                    }
                }

                await Task.Delay(_interval, ct);

                while (!ct.IsCancellationRequested)
                {
                    if (string.IsNullOrEmpty(_currentToken))
                    {
                        // Vẫn chưa có token, thử recover ngầm
                        var recoverResult = await LicenseApiService.VerifyLicenseSecureAsync(_licenseKey, _deviceId, ct);
                        if (recoverResult.success && !string.IsNullOrEmpty(recoverResult.token))
                        {
                            _currentToken = recoverResult.token;
                            _onTokenUpdate(_currentToken);
                            _onWarning?.Invoke("✅ Đã kết nối lại!");
                            _fatalRetryCount = 0;
                        }
                        else
                        {
                            _onWarning?.Invoke("⏳ Vẫn chưa có mạng, đang chờ...");
                            await Task.Delay(_interval, ct);
                            continue;
                        }
                    }

                    var result = await LicenseApiService.SendHeartbeatOnceAsync(_currentToken, _deviceId, ct);

                    switch (result.Outcome)
                    {
                        case HeartbeatOutcome.Continue:
                            _fatalRetryCount = 0;
                            if (!string.IsNullOrEmpty(result.NewToken))
                            {
                                _currentToken = result.NewToken;
                                _onTokenUpdate(result.NewToken);
                            }
                            break;

                        case HeartbeatOutcome.ServerKill:
                            _onWarning?.Invoke("⛔ Phiên bản bị khóa từ máy chủ. Ứng dụng sẽ đóng.");
                            await Task.Delay(3000, ct);
                            System.Windows.Forms.Application.Exit();
                            return;

                        case HeartbeatOutcome.TransientFailure:
                            _onWarning?.Invoke("⚠ Mất kết nối tạm thời, đang chờ...");
                            break;

                        case HeartbeatOutcome.Fatal:
                            _fatalRetryCount++;
                            if (_fatalRetryCount >= 5)
                            {
                                _onWarning?.Invoke("⛔ Đứt kết nối quá lâu. Ứng dụng vẫn hoạt động nhưng chưa xác thực.");
                                _fatalRetryCount = 0;
                            }
                            // Xóa token để vòng sau vòng lặp recover lại
                            _currentToken = null;
                            _onWarning?.Invoke($"⚠ Token hết hạn hoặc lỗi. Sẽ thử lại (lần {_fatalRetryCount})...");
                            break;
                    }

                    int jitterMs = new Random().Next(1000, 5000);
                    await Task.Delay(_interval + TimeSpan.FromMilliseconds(jitterMs), ct);
                }
            }
        }
    }
}