/*
 * 檔案功能：處理 HTTP 網路請求，優化「已登入判斷」邏輯，避免重複登入被伺服器阻擋。
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

            client.DefaultRequestHeaders.ExpectContinue = false;
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            client.DefaultRequestHeaders.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8");
            client.DefaultRequestHeaders.Add("Accept-Language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7");
            client.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");
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

                // 【關鍵修正】：改變「是否已登入」的判斷方式
                // 我們直接去摸一下「你要爬蟲的第一個目標網址」。
                // 這樣就不會被 login.php 騙，避免重複登入被伺服器阻擋！
                string testUrl = settings.CrawlUrls.Count > 0 ? settings.CrawlUrls[0] : currentLoginUrl.Replace("login.php", "index.php");
                
                HttpResponseMessage testResponse = await client.GetAsync(testUrl);
                string testContent = await SafeReadAsStringAsync(testResponse);

                // 如果目標網頁裡面「沒有」登入表單，代表我們現在其實是登入成功的狀態，直接接續爬蟲！
                if (!testContent.Contains("loginfrm") && !testContent.Contains("passwd"))
                {
                    return true; 
                }

                // ==========================================
                // 如果真的沒有登入，才正式開始執行登入流程
                // ==========================================

                Uri loginUri = new Uri(currentLoginUrl);
                string origin = $"{loginUri.Scheme}://{loginUri.Host}";
                if (client.DefaultRequestHeaders.Contains("Origin")) client.DefaultRequestHeaders.Remove("Origin");
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Origin", origin);
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                HttpResponseMessage checkResponse = await client.GetAsync(currentLoginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(checkContent);

                string postTargetUrl = currentLoginUrl;
                HtmlNode formNode = doc.DocumentNode.SelectSingleNode("//form[@name='loginfrm']") ?? doc.DocumentNode.SelectSingleNode("//form");
                
                if (formNode != null)
                {
                    string action = formNode.GetAttributeValue("action", "");
                    if (!string.IsNullOrWhiteSpace(action))
                    {
                        if (action.StartsWith("http")) postTargetUrl = action;
                        else postTargetUrl = new Uri(loginUri, action).ToString();
                    }
                }

                var formDict = new Dictionary<string, string>();

                var inputs = formNode?.SelectNodes(".//input") ?? doc.DocumentNode.SelectNodes("//input");
                if (inputs != null)
                {
                    foreach (var input in inputs)
                    {
                        string name = input.GetAttributeValue("name", "");
                        string value = input.GetAttributeValue("value", "");
                        if (!string.IsNullOrEmpty(name) && input.GetAttributeValue("type", "").ToLower() != "button")
                        {
                            formDict[name] = value;
                        }
                    }
                }

                var selects = formNode?.SelectNodes(".//select") ?? doc.DocumentNode.SelectNodes("//select");
                if (selects != null)
                {
                    foreach (var sel in selects)
                    {
                        string name = sel.GetAttributeValue("name", "");
                        var selectedOption = sel.SelectSingleNode(".//option[@selected]") ?? sel.SelectSingleNode(".//option");
                        if (selectedOption != null && !string.IsNullOrEmpty(name))
                        {
                            formDict[name] = selectedOption.GetAttributeValue("value", selectedOption.InnerText).Trim();
                        }
                    }
                }

                formDict["login"] = username;
                formDict["passwd"] = password;

                var formData = new List<KeyValuePair<string, string>>();
                foreach (var kvp in formDict)
                {
                    formData.Add(new KeyValuePair<string, string>(kvp.Key, kvp.Value));
                }

                var content = new FormUrlEncodedContent(formData);
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");

                HttpResponseMessage response = await client.PostAsync(postTargetUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                if (responseContent.Contains("loginfrm") || responseContent.Contains("passwd"))
                {
                    System.IO.File.WriteAllText("LoginError_Debug.html", responseContent, Encoding.UTF8);
                    throw new Exception("伺服器拒絕了登入請求！\n已將伺服器回傳的畫面存入程式所在資料夾下的 [LoginError_Debug.html]。\n請雙擊打開該檔案，看看伺服器顯示了什麼錯誤訊息。");
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async Task<string> GetHtmlAsync(string url)
        {
            try
            {
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
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

        public async Task DownloadFileAsync(string url, string savePath)
        {
            try
            {
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode(); 
                
                byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();
                System.IO.File.WriteAllBytes(savePath, fileBytes);
            }
            catch (Exception ex)
            {
                throw new Exception($"下載檔案失敗 ({url})：" + ex.Message);
            }
        }
    }
}
