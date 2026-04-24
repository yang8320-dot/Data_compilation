/*
 * 檔案功能：讀寫本機設定檔，支援 Base64 加密密碼，並新增自訂存檔路徑功能。
 * 對應選單名稱：系統設定
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
        // 新增：自訂存檔路徑 (若為空，則預設為程式執行目錄)
        public string ExportPath { get; set; } = "";

        public App_Settings() { Load(); }

        public void Load()
        {
            if (File.Exists(filePath))
            {
                string[] parts = File.ReadAllText(filePath, Encoding.UTF8).Split('|');
                if (parts.Length >= 4)
                {
                    Username = parts[0];
                    Password = DecodeBase64(parts[1]); 
                    LoginUrl = parts[2];
                    CrawlUrls = new List<string>(parts[3].Split(new[] { "^" }, StringSplitOptions.RemoveEmptyEntries));
                }
                // 相容舊版設定檔，若有第5個參數則讀取存檔路徑
                if (parts.Length >= 5)
                {
                    ExportPath = parts[4];
                }
            }
        }

        public void Save()
        {
            string encodedPassword = EncodeBase64(Password);
            string urls = string.Join("^", CrawlUrls);
            
            // 加入 ExportPath 寫入設定檔
            string content = $"{Username}|{encodedPassword}|{LoginUrl}|{urls}|{ExportPath}";
            File.WriteAllText(filePath, content, Encoding.UTF8);
        }

        private string EncodeBase64(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return "";
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(plainText));
        }

        private string DecodeBase64(string encodedText)
        {
            if (string.IsNullOrEmpty(encodedText)) return "";
            try { return Encoding.UTF8.GetString(Convert.FromBase64String(encodedText)); }
            catch { return encodedText; }
        }
    }
}
