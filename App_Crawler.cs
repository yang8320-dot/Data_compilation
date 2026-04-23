/*
 * 檔案功能：解析 HTML 檔案，抓取表格內容與超連結 (具備嚴格防呆與抗干擾機制)。
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

                // 抓取所有的 <tr> 標籤 (不限定在特定的 table 內以增加相容性)
                HtmlNodeCollection rows = doc.DocumentNode.SelectNodes("//tr");
                if (rows == null) return extractedData;

                foreach (HtmlNode row in rows)
                {
                    // 修正 1：使用 "./td" 確保只抓取「直屬」的 td，避免抓到巢狀表格導致欄位暴增或錯亂
                    HtmlNodeCollection cells = row.SelectNodes("./td");
                    
                    // 修正 2：【關鍵防呆】依照系統截圖，真實資料列有 9 個欄位 (包含第 1 欄的 Checkbox)
                    // 若小於 9，代表這是網頁的其他排版表格，直接略過，這能徹底解決 IndexOutOfRange！
                    if (cells == null || cells.Count < 9) continue;

                    // 修正 3：略過 Checkbox (索引 0)，從索引 1 開始對應實際資料
                    string no = cells[1].InnerText.Trim();
                    string formNo = cells[2].InnerText.Trim();
                    
                    // 修正 4：進階防呆，確認單號不為空且長度合理，確保這真的是資料列而非標題列
                    if (string.IsNullOrEmpty(formNo) || formNo.Length < 5) continue;

                    string applyDate = cells[3].InnerText.Trim();
                    string applicant = cells[4].InnerText.Trim();
                    string subject = cells[5].InnerText.Trim();
                    string stepName = cells[6].InnerText.Trim();
                    string status = cells[7].InnerText.Trim();
                    string processTime = cells[8].InnerText.Trim();

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

                    // 寫入陣列 (維持 9 個欄位的輸出格式，無須修改其他模組)
                    extractedData.Add(new string[] { no, formNo, applyDate, applicant, subject, stepName, status, processTime, link });
                }

                return extractedData;
            });
        }
    }
}
