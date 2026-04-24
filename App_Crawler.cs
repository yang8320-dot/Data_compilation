/*
 * 檔案功能：解析 HTML 內容，剔除無效標題列，並修正雙重 eipplus 網址。
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

                        // 收集該行所有的 td 文字
                        List<string> cellTexts = new List<string>();
                        for (int i = offset; i < cells.Count; i++) {
                            cellTexts.Add(CleanText(cells[i].InnerText));
                        }

                        string combinedText = string.Join("", cellTexts);
                        
                        // 【需求 2】過濾掉包含標題列的資料行，一律不存入 Excel
                        if (combinedText.Contains("表單單號") || combinedText.Contains("存檔時間")) continue;

                        // 針對需求 1 的對應位置提取資料 (略過原本的存檔時間)
                        string formNo = cellTexts.Count > 0 ? cellTexts[0] : "";
                        string subject = cellTexts.Count > 1 ? cellTexts[1] : "";
                        string status1 = cellTexts.Count > 2 ? cellTexts[2] : ""; // 對應 HTML的重要性/狀態
                        string status2 = cellTexts.Count > 3 ? cellTexts[3] : ""; // 對應 HTML的狀態
                        // HTML的索引4為存檔時間，我們略過它不抓
                        string applicant = cellTexts.Count > 5 ? cellTexts[5] : "";
                        string handler = cellTexts.Count > 6 ? cellTexts[6] : "";
                        string currentProcessor = cellTexts.Count > 7 ? cellTexts[7] : "";
                        string applyTime = cellTexts.Count > 8 ? cellTexts[8] : "";

                        // 如果單號跟主題都是空的，代表這行不是有效資料，跳過
                        if (string.IsNullOrEmpty(formNo) && string.IsNullOrEmpty(subject)) continue;

                        // 處理網址
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
                                link = ""; // 略過無效的 JavaScript 點擊
                            }

                            // 【需求 3】修正雙重 eipplus 的網址問題
                            link = link.Replace("/eipplus/eipplus/", "/eipplus/");
                        }

                        extractedData.Add(new string[] { formNo, subject, status1, status2, applicant, handler, currentProcessor, applyTime, link });
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
