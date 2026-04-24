/*
 * 檔案功能：處理 HTTP 網路請求，具備「智慧判斷登入狀態」，並穿透企業 Proxy (解決 407 錯誤)。
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
            
            // 【關鍵修正】取得系統當前的 Proxy 設定 (跟隨 Windows / IE / Edge 設定)
            IWebProxy systemProxy = WebRequest.GetSystemWebProxy();
            // 自動帶入你目前登入這台電腦的 Windows 帳號權限，用來通過 Proxy 的 407 驗證
            systemProxy.Credentials = CredentialCache.DefaultCredentials;

            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true,
                Proxy = systemProxy,
                UseProxy = true,
                // 同時告訴 HttpClient 本身也使用預設權限
                UseDefaultCredentials = true 
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

                if (string.IsNullOrWhiteSpace(loginUrl))
                {
                    loginUrl = "http://192.168.1.83/eipplus/login.php"; 
                }

                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; 
                }

                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password)
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                if (responseContent.Contains("loginfrm") || responseContent.Contains("passwd"))
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("登入連線失敗：" + ex.Message);
            }
        }

        public async Task<string> GetHtmlAsync(string url)
        {
            try
            {
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode(); 
                
                return await SafeReadAsStringAsync(response);
            }
            catch (Exception ex)
            {
                throw new Exception($"抓取網頁失敗 ({url})：" + ex.Message);
            }
        }
    }
}
