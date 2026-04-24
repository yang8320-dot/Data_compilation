/*
 * 檔案功能：處理 HTTP 請求、自動判斷登入狀態、解決字元集無效報錯。
 * 對應選單名稱：網路連線
 */
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace FormCrawlerApp
{
    public class App_Network
    {
        private readonly HttpClient client;
        private readonly CookieContainer cookieContainer;

        public App_Network()
        {
            cookieContainer = new CookieContainer();
            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true
            };
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);
        }

        private async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            byte[] bytes = await response.Content.ReadAsByteArrayAsync();
            return Encoding.UTF8.GetString(bytes);
        }

        public async Task<bool> LoginAsync(string username, string password)
        {
            try
            {
                App_Settings settings = new App_Settings();
                string loginUrl = settings.LoginUrl;

                if (string.IsNullOrWhiteSpace(loginUrl)) loginUrl = "http://192.168.1.83/eipplus/login.php";

                // 💡 智慧檢查：先看目前是否已經登入
                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                // 如果畫面上沒有 login 或 passwd 輸入框，代表 Session 仍在，不需登入
                if (!checkContent.Contains("name=\"login\"") && !checkContent.Contains("name=\"passwd\""))
                {
                    return true; 
                }

                // 執行登入 POST
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password)
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                // 💡 判斷登入失敗：如果回應內容還留在登入頁面
                if (responseContent.Contains("loginfrm") || responseContent.Contains("passwd"))
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("網路連線發生異常，請檢查內網環境或網址。" + ex.Message);
            }
        }

        public async Task<string> GetHtmlAsync(string url)
        {
            HttpResponseMessage response = await client.GetAsync(url);
            response.EnsureSuccessStatusCode(); 
            return await SafeReadAsStringAsync(response);
        }
    }
}
