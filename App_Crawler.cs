/*
 * 檔案功能：解析 HTML 內容，加入「安全讀取機制」，徹底避免陣列越界錯誤。
 * 對應選單名稱：網頁爬蟲
 */
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace FormCrawlerApp
{
    public class App_Crawler
    {
        public async Task<List<string[]>> ParseHtmlContentAsync(string htmlContent)
        {
            return await Task.Run(() =>
            {
                List<string[]> extractedData = new List<string[]>();
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(htmlContent);

                HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//tr");
                if (rows == null) return extractedData;

                foreach (HtmlNode row in rows)
                {
                    try 
                    {
                        HtmlNodeCollection cells = row.SelectNodes("./td");
                        if (cells == null || cells.Count < 5) continue; 

                        int offset = 0;
                        if (cells[0].InnerHtml.ToLower().Contains("checkbox"))
                        {
                            offset = 1; 
                        }

                        // 使用 SafeGetText 安全讀取，不用再擔心欄位數量不夠
                        string no = SafeGetText(cells, offset + 0);
                        string formNo = SafeGetText(cells, offset + 1);

                        if (string.IsNullOrEmpty(formNo) || formNo.Length < 5) continue;

                        string applyDate = SafeGetText(cells, offset + 2);
                        string applicant = SafeGetText(cells, offset + 3);
                        string subject = SafeGetText(cells, offset + 4);
                        
                        string stepName = "";
                        string status = "";
                        string processTime = "";

                        // 智慧判斷：如果欄位足夠多，代表有「步驟名稱」
                        if (cells.Count >= offset + 8) 
                        {
                            stepName = SafeGetText(cells, offset + 5);
                            status = SafeGetText(cells, offset + 6);
                            processTime = SafeGetText(cells, offset + 7);
                        }
                        else 
                        {
                            // 欄位較少時，代表沒有步驟名稱，直接取後面的狀態與時間
                            status = SafeGetText(cells, offset + 5);
                            processTime = SafeGetText(cells, offset + 6);
                        }

                        string link = "";
                        HtmlNode linkNode = row.SelectSingleNode(".//a[@href]");
                        if (linkNode != null)
                        {
                            link = linkNode.GetAttributeValue("href", "");
                            if (!link.StartsWith("http") && !link.StartsWith("javascript"))
                            {
                                link = "http://192.168.1.83/eipplus/" + link.TrimStart('/');
                            }
                            else if (link.StartsWith("javascript")) 
                            {
                                link = ""; 
                            }
                        }

                        extractedData.Add(new string[] { no, formNo, applyDate, applicant, subject, stepName, status, processTime, link });
                    }
                    catch 
                    {
                        continue; 
                    }
                }

                return extractedData;
            });
        }

        public async Task<List<string[]>> ParseHtmlAsync(string htmlFilePath)
        {
            return await Task.Run(() =>
            {
                HtmlDocument doc = new HtmlDocument();
                doc.Load(htmlFilePath, System.Text.Encoding.UTF8);
                return ParseHtmlContentAsync(doc.DocumentNode.OuterHtml).Result;
            });
        }

        // 💡 【新增核心機制】：安全取得陣列內容，只要索引不存在就回傳空白，絕不報錯
        private string SafeGetText(HtmlNodeCollection cells, int index)
        {
            if (cells == null || index >= cells.Count) return "";
            string input = cells[index].InnerText;
            if (string.IsNullOrWhiteSpace(input)) return "";
            return HtmlEntity.DeEntitize(input).Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
