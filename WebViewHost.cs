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
            // Dù hàm này có được gọi, nó cũng sẽ dùng chung môi trường đã thiết lập ở Constructor
            await WebView.EnsureCoreWebView2Async(null);
        }

        public static Task NavigateAsync(string url)
        {
            if (WebView != null && WebView.CoreWebView2 != null)
                WebView.CoreWebView2.Navigate(url); // Lưu ý: Nên dùng Navigate thay vì đổi Source
            return Task.CompletedTask;
        }

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