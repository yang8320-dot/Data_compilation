/*
 * 檔案功能：解析 HTML 檔案，抓取表格內容與超連結。
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
        // 非同步解析 HTML 檔案
        public async Task<List<string[]>> ParseHtmlAsync(string htmlFilePath)
        {
            return await Task.Run(() =>
            {
                List<string[]> extractedData = new List<string[]>();
                HtmlDocument doc = new HtmlDocument();
                doc.Load(htmlFilePath, System.Text.Encoding.UTF8);

                // 根據常見的網頁表格結構抓取 <tr>，需依實際 HTML 結構微調 XPath
                HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//table//tr");
                if (rows == null) return extractedData;

                foreach (HtmlNode row in rows)
                {
                    HtmlNodeCollection cells = row.SelectNodes("td");
                    if (cells == null || cells.Count < 8) continue;

                    string no = cells[0].InnerText.Trim();
                    string formNo = cells[1].InnerText.Trim();
                    string applyDate = cells[2].InnerText.Trim();
                    string applicant = cells[3].InnerText.Trim();
                    string subject = cells[4].InnerText.Trim();
                    string stepName = cells[5].InnerText.Trim();
                    string status = cells[6].InnerText.Trim();
                    string processTime = cells[7].InnerText.Trim();

                    // 尋找超連結 (通常綁定在表單單號或主旨上)
                    string link = "";
                    HtmlNode linkNode = row.SelectSingleNode(".//a[@href]");
                    if (linkNode != null)
                    {
                        link = linkNode.GetAttributeValue("href", "");
                        // 處理相對路徑
                        if (!link.StartsWith("http"))
                        {
                            link = "http://192.168.1.83/eipplus/" + link.TrimStart('/');
                        }
                    }

                    extractedData.Add(new string[] { no, formNo, applyDate, applicant, subject, stepName, status, processTime, link });
                }

                return extractedData;
            });
        }
    }
}
