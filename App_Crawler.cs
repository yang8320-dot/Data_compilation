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
                            string htmlReplaced = cells[i].InnerHtml.Replace("<br>", "|").Replace("<br/>", "|").Replace("<br />", "|");
                            
                            HtmlDocument tempDoc = new HtmlDocument();
                            tempDoc.LoadHtml(htmlReplaced);
                            cellTexts.Add(CleanText(tempDoc.DocumentNode.InnerText));
                        }

                        string combinedText = string.Join("", cellTexts);
                        if (combinedText.Contains("表單單號") || combinedText.Contains("存檔時間")) continue;

                        string rawFormNo = cellTexts.Count > 0 ? cellTexts[0] : "";
                        string formNo = "";
                        string category = "";

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
                        string applicant = cellTexts.Count > 4 ? cellTexts[4] : "";
                        string handler = cellTexts.Count > 5 ? cellTexts[5] : "";
                        string currentProcessor = cellTexts.Count > 6 ? cellTexts[6] : "";
                        
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
                            // 【修改點】：將網址替換為 view_frameset (顯示在 Excel 裡的網址)
                            link = link.Replace("print_frameset", "view_frameset").Replace("view_formsflow", "view_frameset");
                        }

                        extractedData.Add(new string[] { formNo, category, subject, status, applicant, handler, currentProcessor, applyTime, modifyTime, expireTime, link });
                    }
                    catch { continue; }
                }
                return extractedData;
            });
        }

        private string FormatToDateOnly(string datetimeStr)
        {
            if (string.IsNullOrWhiteSpace(datetimeStr)) return "";
            datetimeStr = datetimeStr.Trim();

            var match = System.Text.RegularExpressions.Regex.Match(datetimeStr, @"(\d{4})[./-](\d{1,2})[./-](\d{1,2})");
            if (match.Success)
            {
                if (DateTime.TryParse(match.Value.Replace(".", "-").Replace("/", "-"), out DateTime dt))
                {
                    return dt.ToString("yyyy-MM-dd");
                }
            }

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
            return HtmlEntity.DeEntitize(input).Replace("\u00A0", " ").Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
