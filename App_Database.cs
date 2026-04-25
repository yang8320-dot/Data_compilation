using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace FormCrawlerApp
{
    public class App_Database
    {
        // 取得 SQLite 指定檔案內的所有 Table 名稱
        public static List<string> GetTables(string dbPath)
        {
            var tables = new List<string>();
            if (!File.Exists(dbPath)) return tables;
            try {
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) tables.Add(reader["name"].ToString());
                    }
                }
            } catch {}
            return tables;
        }

        // 取得 SQLite 指定 Table 內的所有欄位名稱
        public static List<string> GetColumns(string dbPath, string tableName)
        {
            var cols = new List<string>();
            if (!File.Exists(dbPath) || string.IsNullOrEmpty(tableName)) return cols;
            try {
                using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;")) {
                    conn.Open();
                    using (var cmd = new SQLiteCommand($"PRAGMA table_info({tableName});", conn))
                    using (var reader = cmd.ExecuteReader()) {
                        while (reader.Read()) cols.Add(reader["name"].ToString());
                    }
                }
            } catch {}
            return cols;
        }

        // 核心邏輯：執行寫入資料庫(包含排除檢查、取代與新增)
        public void ProcessCategoryData(CategoryDbSetting config, List<string[]> records)
        {
            if (!config.IsEnabled || string.IsNullOrEmpty(config.DbFilePath) || !File.Exists(config.DbFilePath) || string.IsNullOrEmpty(config.TargetTable))
                return;

            string[] scrapeHeaders = { "表單單號", "表單主題", "狀態", "存檔", "承辦人", "目前處理者", "申請時間", "修改時間", "網址" };

            // 找出 [表單單號] 在資料庫設定中對應的欄位名 (作為辨識鍵值 Key)
            string keyDbColumn = config.Mappings.FirstOrDefault(m => m.ScrapedField == "表單單號")?.DbColumn;
            if (string.IsNullOrEmpty(keyDbColumn)) return; // 沒對應單號無法進行 Update/Exclude，跳過

            using (var conn = new SQLiteConnection($"Data Source={config.DbFilePath};Version=3;"))
            {
                conn.Open();
                using (var transaction = conn.BeginTransaction())
                {
                    foreach (var row in records)
                    {
                        string formNo = row[0]; // row[0] 必定是表單單號
                        if (string.IsNullOrEmpty(formNo)) continue;

                        // 1. 檢查排除清單 (如果有設定)
                        if (!string.IsNullOrEmpty(config.ExcludeTable) && !string.IsNullOrEmpty(config.ExcludeColumn))
                        {
                            using (var cmdCheck = new SQLiteCommand($"SELECT COUNT(1) FROM {config.ExcludeTable} WHERE {config.ExcludeColumn} = @no", conn))
                            {
                                cmdCheck.Parameters.AddWithValue("@no", formNo);
                                long exCount = (long)cmdCheck.ExecuteScalar();
                                if (exCount > 0) continue; // 存在於排除清單，跳過此筆
                            }
                        }

                        // 準備可寫入的對應欄位資料
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
                                // Key值(表單單號)不更新自己
                                if (dbCol != keyDbColumn) updateSets.Add($"{dbCol} = {pName}");
                                
                                parameters.Add(pName, row[i]);
                            }
                        }

                        if (insertCols.Count == 0) continue; // 什麼都沒設定

                        // 2. 檢查主表是否已存在該筆單號
                        bool exists = false;
                        using (var cmdExist = new SQLiteCommand($"SELECT COUNT(1) FROM {config.TargetTable} WHERE {keyDbColumn} = @key", conn))
                        {
                            cmdExist.Parameters.AddWithValue("@key", formNo);
                            exists = (long)cmdExist.ExecuteScalar() > 0;
                        }

                        // 3. 執行 Update (取代) 或 Insert (新增)
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
