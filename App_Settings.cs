/*
 * 檔案功能：讀寫本機設定檔，支援特殊字元密碼的安全儲存 (Base64)。
 * 對應選單名稱：系統設定
 * 對應資料庫名稱：無 (Settings.txt)
 * 對應資料表名稱：無
 */
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace FormCrawlerApp
{
    public class App_Settings
    {
        private readonly string filePath = "Settings.txt";

        public string Username { get; set; } = "";
        public string Password { get; set; } = "";
        public string LoginUrl { get; set; } = "http://192.168.1.83/eipplus/login.php";
        public List<string> CrawlUrls { get; set; } = new List<string>();

        public App_Settings() { Load(); }

        public void Load()
        {
            if (File.Exists(filePath))
            {
                string[] parts = File.ReadAllText(filePath, Encoding.UTF8).Split('|');
                if (parts.Length >= 4)
                {
                    Username = parts[0];
                    // 讀取時將 Base64 解碼還原為包含特殊符號的真實密碼
                    Password = DecodeBase64(parts[1]); 
                    LoginUrl = parts[2];
                    CrawlUrls = new List<string>(parts[3].Split(new[] { "^" }, StringSplitOptions.RemoveEmptyEntries));
                }
            }
        }

        public void Save()
        {
            // 將密碼轉換為 Base64，避免密碼內的特殊符號 (如 | 或 ^) 破壞 txt 檔案的分隔結構
            string encodedPassword = EncodeBase64(Password);
            string urls = string.Join("^", CrawlUrls);
            
            string content = $"{Username}|{encodedPassword}|{LoginUrl}|{urls}";
            File.WriteAllText(filePath, content, Encoding.UTF8);
        }

        // --- 輔助方法：Base64 編碼 ---
        private string EncodeBase64(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            return Convert.ToBase64String(plainTextBytes);
        }

        // --- 輔助方法：Base64 解碼 ---
        private string DecodeBase64(string encodedText)
        {
            if (string.IsNullOrEmpty(encodedText)) return "";
            try 
            {
                byte[] base64EncodedBytes = Convert.FromBase64String(encodedText);
                return Encoding.UTF8.GetString(base64EncodedBytes);
            }
            catch 
            {
                // 若解碼失敗 (例如吃到舊版尚未加密的明文)，則直接回傳原字串防呆
                return encodedText; 
            }
        }
    }
}
