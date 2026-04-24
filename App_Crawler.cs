/*
 * 檔案功能：解析 HTML 內容，支援動態欄位數量偵測 (解決索引超出範圍錯誤)。
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

                        if (cells.Count < offset + 5) continue;

                        string no = CleanText(cells[offset + 0].InnerText);
                        string formNo = CleanText(cells[offset + 1].InnerText);

                        if (string.IsNullOrEmpty(formNo) || formNo.Length < 5) continue;

                        string applyDate = CleanText(cells[offset + 2].InnerText);
                        string applicant = CleanText(cells[offset + 3].InnerText);
                        string subject = CleanText(cells[offset + 4].InnerText);
                        
                        // 【關鍵修正 2】動態判斷欄位數量，適應有/無「步驟名稱」的網頁
                        string stepName = "";
                        string status = "";
                        string processTime = "";

                        if (cells.Count >= offset + 8) 
                        {
                            // 9個欄位：包含步驟名稱
                            stepName = CleanText(cells[offset + 5].InnerText);
                            status = CleanText(cells[offset + 6].InnerText);
                            processTime = CleanText(cells[offset + 7].InnerText);
                        }
                        else if (cells.Count == offset + 7)
                        {
                            // 8個欄位：缺少步驟名稱，自動將對應欄位留空
                            status = CleanText(cells[offset + 5].InnerText);
                            processTime = CleanText(cells[offset + 6].InnerText);
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
                                // 過濾掉 JavaScript 超連結，防止 Excel 產生無效點擊錯誤
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

        private string CleanText(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            return HtmlEntity.DeEntitize(input).Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
