/*
 * 檔案功能：實作文字檔讀寫功能，以 '|' 分隔欄位，作為資料中繼層。
 * 對應選單名稱：資料處理
 * 對應資料庫名稱：無 (純文字儲存)
 * 對應資料表名稱：無 (Data.txt)
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FormCrawlerApp
{
    public class App_TxtStorage
    {
        private readonly string filePath = "Data.txt";

        // 將爬取的資料寫入 txt 檔案
        public void SaveData(List<string[]> records)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // 寫入標頭
                sw.WriteLine("序號|表單單號|申請日期|申請人|主旨|步驟名稱|狀態|處理時間|超連結");
                
                foreach (var record in records)
                {
                    // 強制處理日期格式 (假設索引 2 為申請日期，索引 7 為處理時間)
                    FormatDate(record, 2);
                    FormatDate(record, 7);
                    
                    sw.WriteLine(string.Join("|", record));
                }
            }
        }

        // 從 txt 檔案讀取資料
        public List<string[]> LoadData()
        {
            List<string[]> records = new List<string[]>();
            if (!File.Exists(filePath)) return records;

            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            for (int i = 1; i < lines.Length; i++) // 跳過標頭
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                records.Add(lines[i].Split('|'));
            }
            return records;
        }

        // 強制轉換日期為一致格式 yyyy/MM/dd HH:mm:ss
        private void FormatDate(string[] record, int index)
        {
            if (index < record.Length && DateTime.TryParse(record[index], out DateTime parsedDate))
            {
                record[index] = parsedDate.ToString("yyyy/MM/dd HH:mm:ss");
            }
        }
    }
}
