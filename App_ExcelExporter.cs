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

                    // 更新 11 個欄位標題
                    string[] headers = new string[] { "表單單號", "分類", "表單主題", "狀態", "申請者", "承辦人", "目前處理者", "申請時間", "修改時間", "到期時間", "網址" };
                  
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
                            if (colIdx == 10) // 索引 10 為網址
                            {
                                string link = row[colIdx];
                                if (!string.IsNullOrWhiteSpace(link))
                                {
                                    worksheet.Cell(rowIdx, 11).Value = "超連結";
                                    worksheet.Cell(rowIdx, 11).SetHyperlink(new XLHyperlink(link));
                                    worksheet.Cell(rowIdx, 11).Style.Font.FontColor = XLColor.Blue;
                                    worksheet.Cell(rowIdx, 11).Style.Font.Underline = XLFontUnderlineValues.Single;
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
                    
                    // 11 個欄位的寬度
                    double[] colWidths = { 25, 20, 50, 10, 15, 15, 15, 15, 15, 15, 12 };
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
