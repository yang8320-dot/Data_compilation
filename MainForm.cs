/*
 * 檔案功能：應用程式主視窗，新增資料分類拆檔與多檔案匯出邏輯。
 * 對應選單名稱：主選單
 */
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class MainForm : Form
    {
        private Button btnExecute, btnSettings, btnOpenFolder;
        private Label lblStatus;
        private App_Settings settings;
        private App_Network network;
        private App_Crawler crawler;
        private App_ExcelExporter excelExporter;

        public MainForm()
        {
            settings = new App_Settings();
            network = new App_Network();
            crawler = new App_Crawler();
            excelExporter = new App_ExcelExporter();
            InitializeUI();
        }

        private string GetExportPath()
        {
            if (!string.IsNullOrWhiteSpace(settings.ExportPath)) return settings.ExportPath;
            return Application.StartupPath;
        }

        private void InitializeUI()
        {
            this.Text = "經手表單自動化工具";
            this.Size = new Size(800, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 系統與路徑設定", Location = new Point(15, 15), Size = new Size(180, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => {
                if (new SettingsForm(settings).ShowDialog() == DialogResult.OK) { settings.Load(); }
            };
            menuPanel.Controls.Add(btnSettings);

            btnExecute = new Button { Text = "🚀 執行登入並批次匯入", Location = new Point(30, 100), Size = new Size(250, 60), Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            btnOpenFolder = new Button { Text = "📁 開啟下載資料夾", Location = new Point(300, 100), Size = new Size(200, 60), Cursor = Cursors.Hand };
            btnOpenFolder.Click += (s, e) => {
                string currentPath = GetExportPath();
                if (Directory.Exists(currentPath)) Process.Start("explorer.exe", currentPath);
                else MessageBox.Show($"資料夾不存在！\n路徑：{currentPath}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            };

            lblStatus = new Label { 
                Text = "系統就緒。請確認網址清單與存檔路徑後執行。", 
                Location = new Point(30, 200), Size = new Size(700, 150), AutoSize = false, 
                ForeColor = Color.DarkSlateGray, Font = new Font("Microsoft JhengHei", 11F)
            };

            this.Controls.Add(btnExecute); this.Controls.Add(btnOpenFolder);
            this.Controls.Add(lblStatus); this.Controls.Add(menuPanel);
        }

        private async void BtnExecute_Click(object sender, EventArgs e)
        {
            if (settings.CrawlUrls.Count == 0) { MessageBox.Show("請先設定爬蟲網址清單。"); return; }

            try {
                UIState(false, "系統連線中：正在自動檢查登入狀態...");
                bool login = await network.LoginAsync(settings.Username, settings.Password);
                if (!login) { UIState(true, "登入失敗！請確認帳號密碼。"); return; }

                List<string[]> allData = new List<string[]>();
                Random rnd = new Random(); 

                for (int i = 0; i < settings.CrawlUrls.Count; i++)
                {
                    string url = settings.CrawlUrls[i];
                    UIState(false, $"正在抓取第 {i + 1}/{settings.CrawlUrls.Count} 個網頁...\n目標網址：{url}");
                    
                    string html = await network.GetHtmlAsync(url);
                    var data = await crawler.ParseHtmlContentAsync(html);
                    allData.AddRange(data);

                    if (i < settings.CrawlUrls.Count - 1)
                    {
                        int waitTime = rnd.Next(2000, 4001); 
                        UIState(false, $"第 {i + 1} 頁成功！隨機等待 {waitTime / 1000.0:F1} 秒...");
                        await Task.Delay(waitTime);
                    }
                }

                string exportDir = GetExportPath();
                if (!Directory.Exists(exportDir)) Directory.CreateDirectory(exportDir);

                UIState(false, "網頁抓取完畢，正在依照分類拆分為多個 Excel 檔案...");

                // 【需求 4】定義 6 個要抓取的關鍵字
                string[] targetKeywords = new string[] {
                    "彰濱廠異常改善單", "彰濱聯絡書", "台玻內文", "彰濱廠郵件收文", "彰濱廠虛驚事件及輕度傷害記錄表"
                };

                // 準備分類容器
                Dictionary<string, List<string[]>> categorizedData = new Dictionary<string, List<string[]>>();
                categorizedData["未分類其他表單"] = new List<string[]>(); // 裝沒對應到的表單
                foreach (var kw in targetKeywords) categorizedData[kw] = new List<string[]>();

                // 掃描每一筆資料並分類
                foreach (var row in allData) {
                    string formNo = row[0];
                    string subject = row[1];
                    bool matched = false;
                    
                    // 只要單號或主題包含關鍵字，就放入該分類
                    foreach (var kw in targetKeywords) {
                        if (formNo.Contains(kw) || subject.Contains(kw)) {
                            categorizedData[kw].Add(row);
                            matched = true;
                            break;
                        }
                    }
                    if (!matched) categorizedData["未分類其他表單"].Add(row);
                }

                // 針對有資料的分類，個別匯出 Excel 檔
                int fileCount = 0;
                foreach (var kvp in categorizedData) {
                    if (kvp.Value.Count > 0) {
                        string fileName = Path.Combine(exportDir, $"{kvp.Key}_匯出_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
                        await excelExporter.ExportAsync(fileName, kvp.Value);
                        fileCount++;
                    }
                }
                
                UIState(true, $"✅ 作業完成！\n總共抓取 {allData.Count} 筆資料。\n已自動拆分成 {fileCount} 個 Excel 檔案！\n請點擊「開啟下載資料夾」查看成果。");
            }
            catch (Exception ex) { UIState(true, $"❌ 發生錯誤：\n{ex.Message}"); }
        }

        private void UIState(bool enable, string msg) {
            if (this.InvokeRequired) { this.Invoke(new Action(() => UIState(enable, msg))); return; }
            btnExecute.Enabled = enable; btnSettings.Enabled = enable; lblStatus.Text = msg;
        }
    }
}
