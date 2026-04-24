/*
 * 檔案功能：解析 HTML 內容，支援日期強制轉換為「年/月/日」格式。
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
                        if (cells[0].InnerHtml.ToLower().Contains("checkbox")) offset = 1; 

                        List<string> cellTexts = new List<string>();
                       
                        for (int i = offset; i < cells.Count; i++) {
                            cellTexts.Add(CleanText(cells[i].InnerText));
                        }

                        string combinedText = string.Join("", cellTexts);
                        // 過濾掉包含標題列的資料行
                        if (combinedText.Contains("表單單號") || combinedText.Contains("存檔時間")) continue;

                        string formNo = cellTexts.Count > 0 ? cellTexts[0] : "";
                        string subject = cellTexts.Count > 1 ? cellTexts[1] : "";
                        string status1 = cellTexts.Count > 2 ? cellTexts[2] : ""; 
                        string status2 = cellTexts.Count > 3 ? cellTexts[3] : "";
                        string applicant = cellTexts.Count > 5 ? cellTexts[5] : "";
                        string handler = cellTexts.Count > 6 ? cellTexts[6] : "";
                        string currentProcessor = cellTexts.Count > 7 ? cellTexts[7] : "";
                        // 強制將時間轉換為 YYYY/MM/DD
                        string applyTime = cellTexts.Count > 8 ?
                        FormatToDateOnly(cellTexts[8]) : "";

                        if (string.IsNullOrEmpty(formNo) && string.IsNullOrEmpty(subject)) continue;

                        string link = "";
                        HtmlNode linkNode = row.SelectSingleNode(".//a[@href]");
                        if (linkNode != null)
                        {
                            link = linkNode.GetAttributeValue("href", "");
                            if (!link.StartsWith("http") && !link.StartsWith("javascript"))
                                link = "http://192.168.1.83/eipplus/" + link.TrimStart('/');
                            else if (link.StartsWith("javascript")) 
                                link = "";

                            // 修正雙重 eipplus 的網址問題
                            link = link.Replace("/eipplus/eipplus/", "/eipplus/");
                            
                            // 【修正點 1】網址 view_ 改成 print_
                            link = link.Replace("view_formsflow", "print_frameset");
                        }

                        extractedData.Add(new string[] { formNo, subject, status1, status2, applicant, handler, currentProcessor, applyTime, link });
                    }
                    catch { continue; }
                }
                return extractedData;
            });
        }

        // 輔助方法：強制將日期字串轉換為「年/月/日」
        private string FormatToDateOnly(string datetimeStr)
        {
            if (string.IsNullOrWhiteSpace(datetimeStr)) return "";
            if (DateTime.TryParse(datetimeStr, out DateTime dt))
            {
                return dt.ToString("yyyy/MM/dd");
            }
            
            var parts = datetimeStr.Split(' ');
            if (parts.Length > 0 && parts[0].Contains("/")) return parts[0];
            
            return datetimeStr;
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
