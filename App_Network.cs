/*
 * 檔案功能：處理 HTTP 網路請求，具備「智慧判斷登入狀態」功能，並維護 Session，修正隱藏欄位 (Hidden Fields) 造成的登入失敗問題。
 * 對應選單名稱：網路連線
 */
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack; // 💡 新增這行來解析網頁中的隱藏安全碼

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
                UseCookies = true,
                AllowAutoRedirect = true // 允許自動重導向
            };
            
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);

            // 💡 加入瀏覽器偽裝 (User-Agent)，避免被伺服器阻擋
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
        }

        // 安全讀取網頁內容，避免伺服器回傳無效字元集導致程式崩潰
        private async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            byte[] bytes = await response.Content.ReadAsByteArrayAsync();
            return Encoding.UTF8.GetString(bytes);
        }

        // 執行背景登入 (具備自動抓取隱藏驗證碼機制)
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

                // 【第一階段：檢查是否已登入，並抓取隱藏驗證碼】
                HttpResponseMessage checkResponse = await client.GetAsync(loginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                // 如果畫面上沒有密碼輸入框，代表已經在登入狀態
                if (!checkContent.Contains("passwd"))
                {
                    return true; 
                }

                // 💡 【關鍵修正：解析網頁上的隱藏欄位】
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(checkContent);

                // 從原始碼中精準抓出伺服器要求的參數
                string passwdType = doc.DocumentNode.SelectSingleNode("//input[@name='passwd_type']")?.GetAttributeValue("value", "text") ?? "text";
                string accountType = doc.DocumentNode.SelectSingleNode("//input[@name='account_type']")?.GetAttributeValue("value", "u") ?? "u";
                string encryptString = doc.DocumentNode.SelectSingleNode("//input[@name='encrypt_string']")?.GetAttributeValue("value", "") ?? "";
                string loginEncryptString = doc.DocumentNode.SelectSingleNode("//input[@name='login_encrypt_string']")?.GetAttributeValue("value", "") ?? "";

                // 【第二階段：組合完整參數並執行登入】
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("login", username),
                    new KeyValuePair<string, string>("passwd", password),
                    new KeyValuePair<string, string>("passwd_type", passwdType),
                    new KeyValuePair<string, string>("account_type", accountType),
                    new KeyValuePair<string, string>("encrypt_string", encryptString),
                    new KeyValuePair<string, string>("login_encrypt_string", loginEncryptString),
                    new KeyValuePair<string, string>("lang", "zh-tw"),
                    new KeyValuePair<string, string>("submit_button", "登入") // 模擬按下登入按鈕
                });

                HttpResponseMessage response = await client.PostAsync(loginUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                // 再次驗證：如果登入後的回應內容仍包含密碼框，代表帳密錯誤
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
                return await SafeReadAsStringAsync(response);
            }
            catch (Exception ex)
            {
                throw new Exception($"抓取網頁失敗 ({url})：" + ex.Message);
            }
        }
    }
}
