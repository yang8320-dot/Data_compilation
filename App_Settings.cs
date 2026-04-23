/*
 * 檔案功能：讀寫本機設定檔，儲存帳密、登入網址及多個爬蟲網址。
 * 對應選單名稱：系統設定
 * 對應資料庫名稱：無 (Settings.txt)
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
                    Password = parts[1];
                    LoginUrl = parts[2];
                    // 將儲存的換行標記換回 List
                    CrawlUrls = new List<string>(parts[3].Split(new[] { "^" }, StringSplitOptions.RemoveEmptyEntries));
                }
            }
        }

        public void Save()
        {
            // 使用 ^ 作為多個網址間的子分隔符，使用 | 作為大項分隔符
            string urls = string.Join("^", CrawlUrls);
            string content = $"{Username}|{Password}|{LoginUrl}|{urls}";
            File.WriteAllText(filePath, content, Encoding.UTF8);
        }
    }
}
