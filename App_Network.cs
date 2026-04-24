/*
 * 檔案功能：處理 HTTP 網路請求，具備「智慧判斷登入狀態」功能，並維護 Session，修正伺服器無效字元集報錯。
 * 對應選單名稱：網路連線
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
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
            // 實作 CookieContainer 以便在登入後自動保留 Session
            cookieContainer = new CookieContainer();
            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true
            };
            
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);
        }

        // 💡 【關鍵修正】安全讀取網頁內容，避免伺服器回傳無效字元集導致程式崩潰
        private async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            // 改用讀取 Byte 陣列的方式，直接略過 HttpClient 對 Header 字元集的嚴格驗證
            byte[] bytes = await response.Content.ReadAsByteArrayAsync();
            
            // 根據你之前提供的 HTML 原始碼，該系統實際上使用的是 UTF-8 編碼，我們直接手動解碼
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
                    loginUrl = "http://192.168.1.83/eipplus/login.php"; // 預設防呆網址
                }

                // 【第一階段：檢查是否已登入】
                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                
                // 替換為安全讀取方法
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; 
                }

                // 【第二階段：執行登入】
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password)
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                // 替換為安全讀取方法
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

        // 抓取指定網址的 HTML 原始碼 (帶有登入後的 Cookie 狀態)
        public async Task<string> GetHtmlAsync(string url)
        {
            try
            {
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode(); 
                
                // 替換為安全讀取方法
                return await SafeReadAsStringAsync(response);
            }
            catch (Exception ex)
            {
                throw new Exception($"抓取網頁失敗 ({url})：" + ex.Message);
            }
        }
    }
}
