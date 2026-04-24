/*
 * 檔案功能：處理 HTTP 網路請求，具備「智慧判斷登入狀態」功能，並維護 Session。
 * 修正重點：新增「企業 Proxy 代理伺服器自動認證」，解決 407 錯誤。
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
            
            // 【關鍵修正 1】抓取你電腦上 (IE/Edge) 預設的 Proxy 設定
            IWebProxy systemProxy = WebRequest.GetSystemWebProxy();
            
            // 【關鍵修正 2】自動把目前登入這台 Windows 電腦的員工帳密，交給 Proxy 進行驗證
            systemProxy.Credentials = CredentialCache.DefaultCredentials;

            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true,
                UseProxy = true,
                Proxy = systemProxy,
                // 同時也讓目標伺服器 (若需要 Windows 驗證) 自動通關
                UseDefaultCredentials = true 
            };
            
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);
        }

        // 安全讀取網頁內容，避免伺服器回傳無效字元集導致程式崩潰
        private async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            byte[] bytes = await response.Content.ReadAsByteArrayAsync();
            return Encoding.UTF8.GetString(bytes);
        }

        // 執行背景登入 (具備自動判斷機制)
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

                // 檢查是否已登入
                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; 
                }

                // 執行登入
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

        // 抓取指定網址的 HTML 原始碼
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
