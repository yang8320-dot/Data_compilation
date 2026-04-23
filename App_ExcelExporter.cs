/*
 * 檔案功能：將文字檔資料讀出並匯出為帶有超連結的 Excel 檔案。
 * 對應選單名稱：報表匯出
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace FormCrawlerApp
{
    public class App_ExcelExporter
    {
        public async Task ExportAsync(string outputPath, List<string[]> data)
        {
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("經手表單資料");

                    // 設定標頭
                    string[] headers = { "序號", "表單單號", "申請日期", "申請人", "主旨", "步驟名稱", "狀態", "處理時間" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                    }

                    // 填入資料
                    for (int r = 0; r < data.Count; r++)
                    {
                        int excelRow = r + 2;
                        for (int c = 0; c < 8; c++) // 前 8 個欄位是可視資料
                        {
                            worksheet.Cell(excelRow, c + 1).Value = data[r][c];
                        }

                        // 處理超連結 (第 9 個欄位 data[r][8] 為 URL)
                        string url = data[r][8];
                        if (!string.IsNullOrEmpty(url))
                        {
                            // 將主旨欄位設定為超連結
                            var subjectCell = worksheet.Cell(excelRow, 5);
                            subjectCell.SetHyperlink(new XLHyperlink(url));
                            subjectCell.Style.Font.FontColor = XLColor.Blue;
                            subjectCell.Style.Font.Underline = XLFontUnderlineValues.Single;
                        }
                    }

                    // 自動調整欄寬
                    worksheet.Columns().AdjustToContents();
                    workbook.SaveAs(outputPath);
                }
            });
        }
    }
}
