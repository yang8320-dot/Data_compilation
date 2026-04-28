using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;

namespace FormCrawlerApp
{
    public class App_Database
    {
        public static List<string> GetTables(string dbPath)
        {
            var tables = new List<string>();
            if (!File.Exists(dbPath)) throw new Exception("找不到指定的 SQLite 檔案！\n路徑：" + dbPath);
            
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;Read Write=True;Pooling=False;")) {
                conn.Open();
                using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table';", conn))
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) tables.Add(reader["name"].ToString());
                }
            }
            return tables;
        }

        public static List<string> GetColumns(string dbPath, string tableName)
        {
            var cols = new List<string>();
            if (!File.Exists(dbPath) || string.IsNullOrEmpty(tableName)) return cols;
            
            using (var conn = new SQLiteConnection($"Data Source={dbPath};Version=3;Read Write=True;Pooling=False;")) {
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

            string[] scrapeHeaders = { "表單單號", "分類", "表單主題", "狀態", "申請者", "承辦人", "目前處理者", "申請時間", "修改時間", "到期時間", "網址" };
            string[] vFields = { "v_1", "v_2", "v_3", "v_4" }; 
            string customTextFieldName = "CustomText";

            string keyDbColumn = config.Mappings.FirstOrDefault(m => m.ScrapedField == "表單單號")?.DbColumn;

            using (var conn = new SQLiteConnection($"Data Source={config.DbFilePath};Version=3;Read Write=True;Pooling=False;"))
            {
                try 
                {
                    conn.Open();
                }
                catch (SQLiteException ex)
                {
                    if (ex.ResultCode == SQLiteErrorCode.ReadOnly || ex.Message.ToLower().Contains("readonly"))
                    {
                        throw new Exception($"\n請確認以下兩點：\n1. 該資料庫檔案是否被設定為「唯讀」。\n2. 是否有其他軟體 (例如 DB Browser) 正在開啟並鎖定該檔案，請先關閉它！\n\n系統原始錯誤：{ex.Message}");
                    }
                    throw;
                }

                using (var transaction = conn.BeginTransaction())
                {
                    foreach (var row in records)
                    {
                        // 確保表單單號沒有任何隱藏空白
                        string formNo = row[0]?.Trim(); 
                        if (string.IsNullOrEmpty(formNo)) continue;

                        // 【強化防呆】無死角黑名單比對：忽略大小寫、過濾任何看不見的隱形符號 (BOM, 零寬空白等)
                        if (config.ExcludeFormNumbers != null && config.ExcludeFormNumbers.Count > 0)
                        {
                            bool isBlacklisted = config.ExcludeFormNumbers.Any(blackItem => 
                            {
                                string cleanBlackItem = blackItem?.Trim('\uFEFF', '\u200B', ' ', '\t', '\r', '\n') ?? "";
                                string cleanFormNo = formNo?.Trim('\uFEFF', '\u200B', ' ', '\t', '\r', '\n') ?? "";
                                return string.Equals(cleanBlackItem, cleanFormNo, StringComparison.OrdinalIgnoreCase);
                            });

                            if (isBlacklisted) continue; // 如果在黑名單內，直接跳過這筆不處理
                        }

                        var insertCols = new List<string>();
                        var insertParams = new List<string>();
                        var updateSets = new List<string>();
                        var parameters = new Dictionary<string, object>();

                        // 1. 處理 11 個標準爬蟲欄位
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
                                
                                parameters.Add(pName, row[i]?.Trim());
                            }
                        }

                        // 2. 處理 4 個 [v] 寫入欄位
                        for (int i = 0; i < vFields.Length; i++)
                        {
                            var mapping = config.Mappings.FirstOrDefault(m => m.ScrapedField == vFields[i]);
                            if (mapping != null && !string.IsNullOrEmpty(mapping.DbColumn))
                            {
                                string pName = "@v_param_" + i;
                                string dbCol = mapping.DbColumn;
                                
                                insertCols.Add(dbCol);
                                insertParams.Add(pName);
                                if (dbCol != keyDbColumn) updateSets.Add($"{dbCol} = {pName}");
                                
                                parameters.Add(pName, "v");
                            }
                        }

                        // 3. 處理 [自訂文字] 寫入欄位
                        var customMapping = config.Mappings.FirstOrDefault(m => m.ScrapedField == customTextFieldName);
                        if (customMapping != null && !string.IsNullOrEmpty(customMapping.DbColumn) && !string.IsNullOrEmpty(config.CustomTextValue))
                        {
                            string pName = "@custom_text_param";
                            string dbCol = customMapping.DbColumn;

                            insertCols.Add(dbCol);
                            insertParams.Add(pName);
                            if (dbCol != keyDbColumn) updateSets.Add($"{dbCol} = {pName}");

                            parameters.Add(pName, config.CustomTextValue.Trim());
                        }

                        if (insertCols.Count == 0) continue;

                        bool exists = false;
                        if (!string.IsNullOrEmpty(keyDbColumn))
                        {
                            using (var cmdExist = new SQLiteCommand($"SELECT COUNT(1) FROM {config.TargetTable} WHERE {keyDbColumn} = @key", conn))
                            {
                                cmdExist.Parameters.AddWithValue("@key", formNo);
                                exists = (long)cmdExist.ExecuteScalar() > 0;
                            }
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
