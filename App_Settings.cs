using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace FormCrawlerApp
{
    public class App_Settings
    {
        // 強制綁定絕對路徑，確保不管怎麼開啟都不會讀錯位置
        private readonly string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Settings.txt");

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
                    Password = DecodeBase64(parts[1]); 
                    LoginUrl = parts[2];
                    CrawlUrls = new List<string>(parts[3].Split(new[] { "^" }, StringSplitOptions.RemoveEmptyEntries));
                }
            }
        }

        public void Save()
        {
            string encodedPassword = EncodeBase64(Password);
            string urls = string.Join("^", CrawlUrls);
            string content = $"{Username}|{encodedPassword}|{LoginUrl}|{urls}";
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
