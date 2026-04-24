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

                        // 依序重新對應：0.單號 1.主題 2.狀態 3.存檔 4.承辦人 5.目前處理者 6.申請時間 7.修改時間
                        string formNo = cellTexts.Count > 0 ? cellTexts[0] : "";
                        string subject = cellTexts.Count > 1 ? cellTexts[1] : "";
                        string status = cellTexts.Count > 2 ? cellTexts[2] : ""; 
                        string saveStatus = cellTexts.Count > 3 ? cellTexts[3] : "";
                        string handler = cellTexts.Count > 4 ? cellTexts[4] : "";
                        string currentProcessor = cellTexts.Count > 5 ? cellTexts[5] : "";
                        
                        // 強制轉換為 yyyy-MM-dd
                        string applyTime = cellTexts.Count > 6 ? FormatToDateOnly(cellTexts[6]) : "";
                        string modifyTime = cellTexts.Count > 7 ? FormatToDateOnly(cellTexts[7]) : "";

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

                            link = link.Replace("/eipplus/eipplus/", "/eipplus/");
                            link = link.Replace("view_formsflow", "print_frameset");
                        }

                        extractedData.Add(new string[] { formNo, subject, status, saveStatus, handler, currentProcessor, applyTime, modifyTime, link });
                    }
                    catch { continue; }
                }
                return extractedData;
            });
        }

        // 輔助方法：強制將文字內容中最前面的日期擷取轉換為 yyyy-MM-dd
        private string FormatToDateOnly(string datetimeStr)
        {
            if (string.IsNullOrWhiteSpace(datetimeStr)) return "";
            datetimeStr = datetimeStr.Trim();

            // 尋找字串中最前面的 yyyy/MM/dd, yyyy-MM-dd 或 yyyy.MM.dd
            var match = System.Text.RegularExpressions.Regex.Match(datetimeStr, @"(\d{4})[./-](\d{1,2})[./-](\d{1,2})");
            if (match.Success)
            {
                if (DateTime.TryParse(match.Value.Replace(".", "-").Replace("/", "-"), out DateTime dt))
                {
                    return dt.ToString("yyyy-MM-dd");
                }
            }

            // 若無明顯符號，針對前面第一個單字嘗試轉換
            var parts = datetimeStr.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0 && DateTime.TryParse(parts[0], out DateTime dtFallback))
            {
                return dtFallback.ToString("yyyy-MM-dd");
            }

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
