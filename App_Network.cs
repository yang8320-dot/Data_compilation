/*
 * 檔案功能：處理 HTTP 網路請求，具備真實瀏覽器偽裝、精準表單攔截與 Proxy 穿透。
 * 對應選單名稱：網路連線
 */
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace FormCrawlerApp
{
    public class App_Network
    {
        private readonly HttpClient client;
        private readonly CookieContainer cookieContainer;
        private string currentLoginUrl = "";

        public App_Network()
        {
            cookieContainer = new CookieContainer();
            
            // 系統代理伺服器穿透
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

            // 關閉 100-Continue (某些企業 Proxy 會因為這個 Header 擋住 POST 請求)
            client.DefaultRequestHeaders.ExpectContinue = false;

            // 真實瀏覽器偽裝
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            client.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8");
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
                currentLoginUrl = settings.LoginUrl;

                if (string.IsNullOrWhiteSpace(currentLoginUrl))
                {
                    currentLoginUrl = "http://192.168.1.83/eipplus/login.php"; 
                }

                // 加入 Referer (來源網址) 欺騙伺服器我們是在首頁操作的
                if (client.DefaultRequestHeaders.Contains("Referer"))
                    client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                // 【第一階段：獲取登入頁面與隱藏欄位】
                HttpResponseMessage checkResponse = await client.GetAsync(currentLoginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; // 已經登入，不需再送出
                }

                // 【第二階段：精準模擬表單送出】
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(checkContent);

                // 尋找登入表單的真實目標 Action URL (防呆：有時候會 POST 到不同的 php 檔案)
                string postTargetUrl = currentLoginUrl;
                HtmlNode formNode = doc.DocumentNode.SelectSingleNode("//form[@name='loginfrm']");
                if (formNode != null)
                {
                    string action = formNode.GetAttributeValue("action", "");
                    if (!string.IsNullOrWhiteSpace(action))
                    {
                        if (action.StartsWith("http")) postTargetUrl = action;
                        else postTargetUrl = new Uri(new Uri(currentLoginUrl), action).ToString();
                    }
                }

                // 使用 Dictionary 蒐集畫面上「所有」輸入欄位 (包含 submit 按鈕與隱藏 token)
                var formDict = new Dictionary<string, string>();
                var inputs = doc.DocumentNode.SelectNodes("//form[@name='loginfrm']//input") ?? doc.DocumentNode.SelectNodes("//input");
                
                if (inputs != null)
                {
                    foreach (var input in inputs)
                    {
                        string name = input.GetAttributeValue("name", "");
                        string value = input.GetAttributeValue("value", "");
                        
                        // 將欄位裝入字典 (排除純 button)
                        if (!string.IsNullOrEmpty(name) && input.GetAttributeValue("type", "").ToLower() != "button")
                        {
                            formDict[name] = value;
                        }
                    }
                }

                // 強制覆寫我們的帳號與密碼
                formDict["login"] = username;
                formDict["passwd"] = password;

                // 轉換為 HttpClient 規定的格式
                var formData = new List<KeyValuePair<string, string>>();
                foreach (var kvp in formDict)
                {
                    formData.Add(new KeyValuePair<string, string>(kvp.Key, kvp.Value));
                }

                var content = new FormUrlEncodedContent(formData);

                // 送出登入請求
                HttpResponseMessage response = await client.PostAsync(postTargetUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                // 驗證是否登入成功
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
                // 抓取爬蟲資料時，確保來源網址(Referer)維持在登入首頁，避免被踢出
                if (client.DefaultRequestHeaders.Contains("Referer"))
                    client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

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
