using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace FormCrawlerApp
{
    public class App_ExcelExporter
    {
        public async Task ExportAsync(string filePath, List<string[]> dataList)
        {
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("表單資料");

                    // 依序更新標題名稱
                    string[] headers = new string[] { "表單單號", "表單主題", "狀態", "存檔", "承辦人", "目前處理者", "申請時間", "修改時間", "網址" };
                  
                    for (int i = 0; i < headers.Length; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = headers[i];
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                        worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }

                    int rowIdx = 2;
                    foreach (var row in dataList)
                    {
                        for (int colIdx = 0; colIdx < row.Length; colIdx++)
                        {
                            if (colIdx == 8) // 第9欄是網址欄位
                            {
                                string link = row[colIdx];
                                if (!string.IsNullOrWhiteSpace(link))
                                {
                                    worksheet.Cell(rowIdx, 9).Value = "超連結";
                                    worksheet.Cell(rowIdx, 9).SetHyperlink(new XLHyperlink(link));
                                    worksheet.Cell(rowIdx, 9).Style.Font.FontColor = XLColor.Blue;
                                    worksheet.Cell(rowIdx, 9).Style.Font.Underline = XLFontUnderlineValues.Single;
                                }
                            }
                            else
                            {
                                worksheet.Cell(rowIdx, colIdx + 1).Value = row[colIdx];
                            }
                        }
                        rowIdx++;
                    }

                    worksheet.Rows().Height = 25;
                    
                    // 9 個欄位的自訂欄寬
                    double[] colWidths = { 50, 65, 9, 9, 15, 15, 15, 15, 10 };
                    for (int i = 0; i < colWidths.Length; i++)
                    {
                        worksheet.Column(i + 1).Width = colWidths[i];
                    }

                    workbook.SaveAs(filePath);
                }
            });
        }
    }
}
