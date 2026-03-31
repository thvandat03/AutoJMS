using Microsoft.Web.WebView2.WinForms;
using System;
using System.Threading.Tasks;

namespace AutoJMS
{
    public static class WebViewHost
    {
        // Biến lưu trữ đối tượng WebView
        public static WebView2 WebView { get; private set; }

        public static async Task InitAsync(WebView2 webView)
        {
            WebView = webView;
            await WebView.EnsureCoreWebView2Async();
        }

        public static Task NavigateAsync(string url)
        {
            if (WebView != null && WebView.CoreWebView2 != null)
                WebView.Source = new Uri(url);
            return Task.CompletedTask;
        }

        // --- ĐÂY LÀ HÀM BẠN ĐANG THIẾU (Đã thêm lại) ---
        public static async Task<string> ExecJsAsync(string script)
        {
            if (WebView != null && WebView.CoreWebView2 != null)
            {
                return await WebView.ExecuteScriptAsync(script);
            }
            return string.Empty;
        }
    }
}