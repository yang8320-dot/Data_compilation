using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace FormCrawlerApp
{
    public class App_Database
    {
        // 取得 SQLite 所有 Table (拿掉空的 catch，讓錯誤可以顯示出來)
        public static List<string> GetTables(string dbPath)
        {
            var tables = new List<string>();
            if (!File.Exists(dbPath)) throw new Exception("找不到指定的 SQLite 檔案！\n路徑：" + dbPath);
            
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", conn))
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) tables.Add(reader["name"].ToString());
                }
            }
            return tables;
        }

        // 取得 SQLite 指定 Table 內的所有欄位名稱
        public static List<string> GetColumns(string dbPath, string tableName)
        {
            var cols = new List<string>();
            if (!File.Exists(dbPath) || string.IsNullOrEmpty(tableName)) return cols;
            
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) {
                conn.Open();
                using (var cmd = new SQLiteCommand($"PRAGMA table_info({tableName});", conn))
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) cols.Add(reader["name"].ToString());
                }
            }
            return cols;
        }

        public void ProcessCategoryData(CategoryDbSetting config, List<string[]> records)
        {
            if (!config.IsEnabled || string.IsNullOrEmpty(config.DbFilePath) || !File.Exists(config.DbFilePath) || string.IsNullOrEmpty(config.TargetTable))
                return;

            string[] scrapeHeaders = { "表單單號", "表單主題", "狀態", "存檔", "承辦人", "目前處理者", "申請時間", "修改時間", "網址" };

            string keyDbColumn = config.Mappings.FirstOrDefault(m => m.ScrapedField == "表單單號")?.DbColumn;
            if (string.IsNullOrEmpty(keyDbColumn)) return;

            using (var conn = new SQLiteConnection($"Data Source={config.DbFilePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    foreach (var row in records)
                    {
                        string formNo = row[0]; 
                        if (string.IsNullOrEmpty(formNo)) continue;

                        // 檢查自訂黑名單，若清單內有此表單單號，則略過不寫入
                        if (config.ExcludeFormNumbers != null && config.ExcludeFormNumbers.Contains(formNo))
                        {
                            continue;
                        }

                        var insertCols = new List<string>();
                        var insertParams = new List<string>();
                        var updateSets = new List<string>();
                        var parameters = new Dictionary<string, object>();

                        for (int i = 0; i < scrapeHeaders.Length; i++)
                        {
                            var mapping = config.Mappings.FirstOrDefault(m => m.ScrapedField == scrapeHeaders[i]);
                            if (mapping != null && !string.IsNullOrEmpty(mapping.DbColumn))
                            {
                                string pName = "@p" + i;
                                string dbCol = mapping.DbColumn;
                                
                                insertCols.Add(dbCol);
                                insertParams.Add(pName);
                                if (dbCol != keyDbColumn) updateSets.Add($"{dbCol} = {pName}");
                                
                                parameters.Add(pName, row[i]);
                            }
                        }

                        if (insertCols.Count == 0) continue;

                        bool exists = false;
                        using (var cmdExist = new SQLiteCommand($"SELECT COUNT(1) FROM {config.TargetTable} WHERE {keyDbColumn} = @key", conn))
                        {
                            cmdExist.Parameters.AddWithValue("@key", formNo);
                            exists = (long)cmdExist.ExecuteScalar() > 0;
                        }

                        string sql = "";
                        if (exists && updateSets.Count > 0)
                        {
                            sql = $"UPDATE {config.TargetTable} SET {string.Join(", ", updateSets)} WHERE {keyDbColumn} = @key";
                        }
                        else if (!exists)
                        {
                            sql = $"INSERT INTO {config.TargetTable} ({string.Join(", ", insertCols)}) VALUES ({string.Join(", ", insertParams)})";
                        }

                        if (!string.IsNullOrEmpty(sql))
                        {
                            using (var cmdExec = new SQLiteCommand(sql, conn))
                            {
                                foreach (var kvp in parameters) cmdExec.Parameters.AddWithValue(kvp.Key, kvp.Value);
                                if (exists) cmdExec.Parameters.AddWithValue("@key", formNo);
                                cmdExec.ExecuteNonQuery();
                            }
                        }
                    }
                    transaction.Commit();
                }
            }
        }
    }
}
