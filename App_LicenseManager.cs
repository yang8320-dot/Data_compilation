using System;
using System.Data.SQLite;
using System.IO;

namespace FormCrawlerApp
{
    public static class App_LicenseManager
    {
        private const string DbName = "SystemConfig.sqlite";
        private const string TableName = "AllowedUsers";
        private static readonly DateTime ExpirationDate = new DateTime(2050, 12, 31);

        // 預設授權使用者清單
        public static readonly string[] DefaultUsers = { "黃忠揚", "TJ700657", "TJ700228", "TJ700533", "TJ204159" };

        public static bool VerifyLicense()
        {
            // 1. 檢查軟體使用期限
            if (DateTime.Today > ExpirationDate)
            {
                return false;
            }

            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DbName);
            string connectionString = $"Data Source={dbPath};Version=3;Read Write=True;Pooling=False;";

            // 2. 初始化資料庫與資料表
            if (!File.Exists(dbPath))
            {
                SQLiteConnection.CreateFile(dbPath);
            }

            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();

                string createSql = $"CREATE TABLE IF NOT EXISTS [{TableName}] (Id INTEGER PRIMARY KEY AUTOINCREMENT, [使用者帳號] TEXT);";
                using (SQLiteCommand cmd = new SQLiteCommand(createSql, conn))
                {
                    cmd.ExecuteNonQuery();
                }

                // 3. 檢查表內是否有資料，若無則寫入預設名單
                long count = 0;
                using (SQLiteCommand cmd = new SQLiteCommand($"SELECT COUNT(1) FROM [{TableName}]", conn))
                {
                    count = (long)cmd.ExecuteScalar();
                }

                if (count == 0)
                {
                    using (var transaction = conn.BeginTransaction())
                    {
                        foreach (string user in DefaultUsers)
                        {
                            using (SQLiteCommand cmd = new SQLiteCommand($"INSERT INTO [{TableName}] ([使用者帳號]) VALUES (@user)", conn))
                            {
                                cmd.Parameters.AddWithValue("@user", user);
                                cmd.ExecuteNonQuery();
                            }
                        }
                        transaction.Commit();
                    }
                }

                // 4. 驗證當前電腦登入的帳號 (不分大小寫)
                string currentComputerUser = Environment.UserName.Trim();
                
                using (SQLiteCommand cmd = new SQLiteCommand($"SELECT [使用者帳號] FROM [{TableName}]", conn))
                using (SQLiteDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string allowedUser = reader["使用者帳號"]?.ToString().Trim();
                        if (string.Equals(currentComputerUser, allowedUser, StringComparison.OrdinalIgnoreCase))
                        {
                            return true; // 驗證通過
                        }
                    }
                }
            }

            // 驗證失敗
            return false;
        }
    }
}
