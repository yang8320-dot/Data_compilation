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
                            // 修正：將 <br> 標籤替換為 |，避免文字沾黏
                            string htmlReplaced = cells[i].InnerHtml.Replace("<br>", "|").Replace("<br/>", "|").Replace("<br />", "|");
                            
                            // 用暫存 Document 解析乾淨文字
                            HtmlDocument tempDoc = new HtmlDocument();
                            tempDoc.LoadHtml(htmlReplaced);
                            cellTexts.Add(CleanText(tempDoc.DocumentNode.InnerText));
                        }

                        string combinedText = string.Join("", cellTexts);
                        // 過濾掉包含標題列的資料行
                        if (combinedText.Contains("表單單號") || combinedText.Contains("存檔時間")) continue;

                        // === 真實欄位索引對應 (已扣除首欄 Checkbox) ===
                        // 0:單號+分類, 1:主題, 2:狀態, 3:存檔時間(隱藏), 4:申請者, 5:承辦人, 6:目前處理者, 7:申請時間, 8:修改時間, 9:到期時間, 10:完成時間(隱藏)
                        
                        string rawFormNo = cellTexts.Count > 0 ? cellTexts[0] : "";
                        string formNo = "";
                        string category = "";

                        // 切割單號與分類
                        if (rawFormNo.Contains("|"))
                        {
                            var parts = rawFormNo.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                            formNo = parts.Length > 0 ? parts[0].Trim() : "";
                            category = parts.Length > 1 ? parts[1].Trim() : "";
                        }
                        else
                        {
                            formNo = rawFormNo;
                        }

                        string subject = cellTexts.Count > 1 ? cellTexts[1] : "";
                        string status = cellTexts.Count > 2 ? cellTexts[2] : ""; 
                        // string saveTime = cellTexts.Count > 3 ? cellTexts[3] : ""; (隱藏的存檔時間，略過不放入結果)
                        string applicant = cellTexts.Count > 4 ? cellTexts[4] : "";
                        string handler = cellTexts.Count > 5 ? cellTexts[5] : "";
                        string currentProcessor = cellTexts.Count > 6 ? cellTexts[6] : "";
                        
                        // 強制轉換為 yyyy-MM-dd
                        string applyTime = cellTexts.Count > 7 ? FormatToDateOnly(cellTexts[7]) : "";
                        string modifyTime = cellTexts.Count > 8 ? FormatToDateOnly(cellTexts[8]) : "";
                        string expireTime = cellTexts.Count > 9 ? FormatToDateOnly(cellTexts[9]) : "";

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

                        // 回傳陣列擴增為 11 個元素
                        extractedData.Add(new string[] { formNo, category, subject, status, applicant, handler, currentProcessor, applyTime, modifyTime, expireTime, link });
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
            // 增加 \u00A0 處理，徹底去除網頁實體空白
            return HtmlEntity.DeEntitize(input).Replace("\u00A0", " ").Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
