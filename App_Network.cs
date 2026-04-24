/*
 * 檔案功能：處理 HTTP 網路請求，具備全表單攔截、代理伺服器穿透、以及實體檔案(PDF)下載功能。
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
            
            // 系統代理伺服器穿透 (解決 407 Proxy 錯誤)
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

        // 安全讀取網頁內容，避免伺服器回傳「無效字元集」導致程式崩潰
        private async Task<string> SafeReadAsStringAsync(HttpResponseMessage response)
        {
            byte[] bytes = await response.Content.ReadAsByteArrayAsync();
            return Encoding.UTF8.GetString(bytes);
        }

        // 執行背景登入 (具備自動判斷機制與隱藏欄位抓取)
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

                // 加入 Origin 與 Referer，滿足老舊系統的防偽造檢查
                Uri loginUri = new Uri(currentLoginUrl);
                string origin = $"{loginUri.Scheme}://{loginUri.Host}";
                if (client.DefaultRequestHeaders.Contains("Origin")) client.DefaultRequestHeaders.Remove("Origin");
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Origin", origin);
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                // 【第一階段：獲取登入頁面】
                HttpResponseMessage checkResponse = await client.GetAsync(currentLoginUrl);
                string checkContent = await SafeReadAsStringAsync(checkResponse);

                if (!checkContent.Contains("loginfrm") && !checkContent.Contains("passwd"))
                {
                    return true; // 已經登入，不需再送出
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

                // 1. 抓取所有 <input> (包含隱藏的 Token)
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

                // 2. 抓取所有 <select> (下拉選單，如語系、網域等)
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

                // 覆寫帳號與密碼
                formDict["login"] = username;
                formDict["passwd"] = password;

                var formData = new List<KeyValuePair<string, string>>();
                foreach (var kvp in formDict)
                {
                    formData.Add(new KeyValuePair<string, string>(kvp.Key, kvp.Value));
                }

                var content = new FormUrlEncodedContent(formData);
                
                // 【老舊系統殺手鐧】強制移除 Content-Type 的 charset=utf-8 後綴
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded");

                HttpResponseMessage response = await client.PostAsync(postTargetUrl, content);
                response.EnsureSuccessStatusCode();

                string responseContent = await SafeReadAsStringAsync(response);

                // 驗證是否登入成功
                if (responseContent.Contains("loginfrm") || responseContent.Contains("passwd"))
                {
                    // 【除錯神器】把伺服器拒絕登入的畫面存下來
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

        // 抓取指定網址的 HTML 原始碼 (帶有登入後的 Cookie 狀態)
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

        // 【新增功能】下載檔案並儲存至本機
        public async Task DownloadFileAsync(string url, string savePath)
        {
            try
            {
                // 維持登入狀態與來源驗證，防止被踢出
                if (client.DefaultRequestHeaders.Contains("Referer")) client.DefaultRequestHeaders.Remove("Referer");
                client.DefaultRequestHeaders.Add("Referer", currentLoginUrl);

                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode(); 
                
                // 直接讀取原始位元組並存成實體檔案
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
