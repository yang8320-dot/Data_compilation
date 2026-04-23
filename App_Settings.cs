/*
 * 檔案功能：讀寫本機設定檔，負責儲存系統登入帳號與密碼。
 * 對應選單名稱：系統設定
 * 對應資料庫名稱：無 (純文字儲存)
 * 對應資料表名稱：無 (Settings.txt)
 */
using System;
using System.IO;
using System.Text;

namespace FormCrawlerApp
{
    public class App_Settings
    {
        private readonly string filePath = "Settings.txt";

        public string Username { get; private set; } = "";
        public string Password { get; private set; } = "";

        public App_Settings()
        {
            Load();
        }

        public void Load()
        {
            if (File.Exists(filePath))
            {
                string content = File.ReadAllText(filePath, Encoding.UTF8);
                string[] parts = content.Split('|');
                if (parts.Length >= 2)
                {
                    Username = parts[0];
                    Password = parts[1];
                }
            }
        }

        public void Save(string user, string pass)
        {
            Username = user;
            Password = pass;
            // 遵守純文字且以 '|' 分隔的資料規範
            File.WriteAllText(filePath, $"{user}|{pass}", Encoding.UTF8);
        }

        public bool HasCredentials()
        {
            return !string.IsNullOrWhiteSpace(Username) && !string.IsNullOrWhiteSpace(Password);
        }
    }
}
