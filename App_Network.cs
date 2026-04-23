/*
 * 檔案功能：處理 HTTP 網路請求與自動登入，並維護 Session Cookie。
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
            cookieContainer = new CookieContainer();
            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true
            };
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);
        }

        // 執行背景登入
        public async Task<bool> LoginAsync(string username, string password)
        {
            try
            {
                // 根據網頁原始碼，登入端點為 login.php
                string loginUrl = "http://192.168.1.83/eipplus/login.php";

                // 設定 POST 參數 (推測標準 input name 為 login 與 passwd)
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password)
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await response.Content.ReadAsStringAsync();

                // 簡單驗證：若回應內容仍包含登入按鈕或表單，通常代表帳密錯誤
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
    }
}
