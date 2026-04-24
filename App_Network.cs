/*
 * 檔案功能：處理 HTTP 網路請求，具備全表單攔截、老舊伺服器相容性，並修正「登入成功誤判」的問題。
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

        // 💡 輔助方法：精準判斷當前畫面是不是「登入畫面」
        private bool IsLoginPage(string htmlContent)
        {
            // 首頁可能會有 passwd 字眼(修改密碼連結)，所以改用 input 標籤特徵來判斷
            return htmlContent.Contains("type=\"password\"") || htmlContent.Contains("name=\"loginfrm\"");
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

                Uri loginUri = new Uri(currentLoginUrl);
                string origin = $"{loginUri.Scheme}://{loginUri.Host}";
                if (client.DefaultRequestHeaders.Contains("Origin")) client.DefaultRequestHeaders.Remove("Origin");
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Origin", origin);
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                // 【第一階段：獲取登入頁面】
                HttpResponseMessage checkResponse = await client.GetAsync(currentLoginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                // 如果畫面沒有密碼輸入框，代表已經登入成功了
                if (!IsLoginPage(checkContent))
                {
                    return true; 
                }

                // 【第二階段：精準模擬表單送出】
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

                // 【關鍵修正】用更嚴格的 IsLoginPage 來驗證是否卡在登入畫面
                if (IsLoginPage(responseContent))
                {
                    System.IO.File.WriteAllText("LoginError_Debug.html", responseContent, Encoding.UTF8);
                    throw new Exception("伺服器拒絕了登入請求！請確認帳號密碼是否正確。");
                }

                // 若沒有密碼輸入框，代表已經成功進入首頁！
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
    }
}
