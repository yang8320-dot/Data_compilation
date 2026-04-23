/*
 * 檔案功能：處理 HTTP 網路請求，具備「智慧判斷登入狀態」功能，並維護 Session。
 * 對應選單名稱：網路連線
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
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

        // 執行背景登入 (具備自動判斷機制)
        public async Task<bool> LoginAsync(string username, string password)
        {
            try
            {
                // 動態讀取設定檔中的登入網址
                App_Settings settings = new App_Settings();
                string loginUrl = settings.LoginUrl;

                if (string.IsNullOrWhiteSpace(loginUrl))
                {
                    loginUrl = "http://192.168.1.83/eipplus/login.php"; // 預設防呆網址
                }

                // 【第一階段：檢查是否已登入】
                // 先對登入網址發送 GET 請求，抓取目前的網頁原始碼
                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                string checkContent = await checkResponse.Content.ReadAsStringAsync();

                // 根據網頁原始碼，登入畫面會有 name="loginfrm" 或 name="passwd"
                // 如果沒有這些特徵，代表已經在登入狀態 (沒看到登入畫面)，直接回傳 true 開始爬蟲
                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; 
                }

                // 【第二階段：執行登入】
                // 如果畫面上確實有登入表單，才設定 POST 參數送出帳密
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password)
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await response.Content.ReadAsStringAsync();

                // 再次驗證：如果登入後的回應內容仍包含登入表單，通常代表帳密錯誤
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
                response.EnsureSuccessStatusCode(); // 確保 HTTP 狀態碼為 200 OK
                
                return await response.Content.ReadAsStringAsync();
            }
            catch (Exception ex)
            {
                throw new Exception($"抓取網頁失敗 ({url})：" + ex.Message);
            }
        }
    }
}
