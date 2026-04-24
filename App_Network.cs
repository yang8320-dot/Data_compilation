/*
 * 檔案功能：處理 HTTP 網路請求，加入真實瀏覽器偽裝 (User-Agent) 與動態隱藏欄位抓取。
 * 對應選單名稱：網路連線
 */
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack; // 💡 引入 HtmlAgilityPack 來解析隱藏欄位

namespace FormCrawlerApp
{
    public class App_Network
    {
        private readonly HttpClient client;
        private readonly CookieContainer cookieContainer;

        public App_Network()
        {
            cookieContainer = new CookieContainer();
            
            IWebProxy systemProxy = WebRequest.GetSystemWebProxy();
            systemProxy.Credentials = CredentialCache.DefaultCredentials;

            HttpClientHandler handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true,
                Proxy = systemProxy,
                UseProxy = true,
                UseDefaultCredentials = true 
            };
            
            client = new HttpClient(handler);
            client.Timeout = TimeSpan.FromSeconds(30);

            // 【關鍵升級 1】偽裝成真實的 Google Chrome 瀏覽器，防止伺服器阻擋機器人
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            client.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8");
            client.DefaultRequestHeaders.Add("Accept-Language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7");
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

                // 【關鍵升級 2】動態抓取登入頁面上的所有隱藏欄位 (Hidden Fields)
                var formData = new List<KeyValuePair<string, string>>();
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(checkContent);

                // 掃描畫面上所有的 input 標籤
                var inputs = doc.DocumentNode.SelectNodes("//input");
                if (inputs != null)
                {
                    foreach (var input in inputs)
                    {
                        string name = input.GetAttributeValue("name", "");
                        string type = input.GetAttributeValue("type", "").ToLower();
                        string value = input.GetAttributeValue("value", "");

                        // 自動收集隱藏欄位與預設值 (排除掉我們自己要填的欄位與按鈕)
                        if (!string.IsNullOrEmpty(name) && name != "login" && name != "passwd" && type != "submit" && type != "button")
                        {
                            formData.Add(new KeyValuePair<string, string>(name, value));
                        }
                    }
                }

                // 放入我們設定的帳號與密碼
                formData.Add(new KeyValuePair<string, string>("login", username));
                formData.Add(new KeyValuePair<string, string>("passwd", password));

                var content = new FormUrlEncodedContent(formData);

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
