/*
 * 檔案功能：解析 HTML 檔案，抓取表格內容與超連結 (防彈級解析，支援自動偵測欄位位移)。
 * 對應選單名稱：網頁爬蟲
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace FormCrawlerApp
{
    public class App_Crawler
    {
        public async Task<List<string[]>> ParseHtmlAsync(string htmlFilePath)
        {
            return await Task.Run(() =>
            {
                List<string[]> extractedData = new List<string[]>();
                HtmlDocument doc = new HtmlDocument();
                doc.Load(htmlFilePath, System.Text.Encoding.UTF8);

                HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//tr");
                if (rows == null) return extractedData;

                foreach (HtmlNode row in rows)
                {
                    try 
                    {
                        HtmlNodeCollection cells = row.SelectNodes("./td");
                        if (cells == null || cells.Count < 5) continue; // 小於 5 欄絕對不是我們要的資料

                        // 自動偵測偏移量：判斷第一欄是否為 Checkbox
                        int offset = 0;
                        if (cells[0].InnerHtml.ToLower().Contains("checkbox"))
                        {
                            offset = 1; // 如果第一欄是 Checkbox，資料欄位往後延遲一格
                        }

                        // 確保扣除偏移量後，仍有足夠的 8 個欄位可抓取
                        if (cells.Count < offset + 8) continue;

                        string no = CleanText(cells[offset + 0].InnerText);
                        string formNo = CleanText(cells[offset + 1].InnerText);

                        // 嚴格過濾：單號通常要有一定長度，若無則跳過 (防抓到表頭或空行)
                        if (string.IsNullOrEmpty(formNo) || formNo.Length < 5) continue;

                        string applyDate = CleanText(cells[offset + 2].InnerText);
                        string applicant = CleanText(cells[offset + 3].InnerText);
                        string subject = CleanText(cells[offset + 4].InnerText);
                        string stepName = CleanText(cells[offset + 5].InnerText);
                        string status = CleanText(cells[offset + 6].InnerText);
                        string processTime = CleanText(cells[offset + 7].InnerText);

                        string link = "";
                        HtmlNode linkNode = row.SelectSingleNode(".//a[@href]");
                        if (linkNode != null)
                        {
                            link = linkNode.GetAttributeValue("href", "");
                            if (!link.StartsWith("http"))
                            {
                                link = "http://192.168.1.83/eipplus/" + link.TrimStart('/');
                            }
                        }

                        extractedData.Add(new string[] { no, formNo, applyDate, applicant, subject, stepName, status, processTime, link });
                    }
                    catch 
                    {
                        // 若單一列發生任何不可預期錯誤直接跳過，絕不讓程式崩潰
                        continue; 
                    }
                }

                return extractedData;
            });
        }

        // 輔助方法：清除 HTML 特殊字元與換行符號
        private string CleanText(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            return HtmlEntity.DeEntitize(input).Replace("\r", "").Replace("\n", "").Trim();
        }
    }
}
