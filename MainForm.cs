/*
 * 檔案功能：應用程式主視窗，新增 TXT 暫存與批次 PDF 下載功能。
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
        private Button btnExecute, btnSettings, btnOpenFolder, btnDownloadPdf;
        private ComboBox cmbCategories;
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
            this.Size = new Size(800, 500); // 加高視窗容納新按鈕
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 系統與路徑設定", Location = new Point(15, 15), Size = new Size(180, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => {
                if (new SettingsForm(settings).ShowDialog() == DialogResult.OK) { settings.Load(); }
            };
            menuPanel.Controls.Add(btnSettings);

            // 第一排：爬蟲與開啟資料夾
            btnExecute = new Button { Text = "🚀 1. 執行登入並批次匯入", Location = new Point(30, 90), Size = new Size(250, 50), Font = new Font("Microsoft JhengHei", 11F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            btnOpenFolder = new Button { Text = "📁 開啟下載資料夾", Location = new Point(300, 90), Size = new Size(200, 50), Cursor = Cursors.Hand };
            btnOpenFolder.Click += (s, e) => {
                string currentPath = GetExportPath();
                if (Directory.Exists(currentPath)) Process.Start("explorer.exe", currentPath);
                else MessageBox.Show($"資料夾不存在！\n路徑：{currentPath}");
            };

            // 第二排：下拉選單與下載 PDF
            cmbCategories = new ComboBox { Location = new Point(30, 160), Size = new Size(250, 30), DropDownStyle = ComboBoxStyle.DropDownList };
            
            btnDownloadPdf = new Button { Text = "📄 2. 依選單下載 PDF", Location = new Point(300, 155), Size = new Size(200, 40), Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold), BackColor = Color.LightSkyBlue, Cursor = Cursors.Hand };
            btnDownloadPdf.Click += BtnDownloadPdf_Click;

            lblStatus = new Label { 
                Text = "系統就緒。請確認網址清單與存檔路徑後執行。", 
                Location = new Point(30, 220), Size = new Size(700, 150), AutoSize = false, 
                ForeColor = Color.DarkSlateGray, Font = new Font("Microsoft JhengHei", 11F)
            };

            this.Controls.Add(btnExecute); this.Controls.Add(btnOpenFolder);
            this.Controls.Add(cmbCategories); this.Controls.Add(btnDownloadPdf);
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
                        await Task.Delay(waitTime);
                    }
                }

                string exportDir = GetExportPath();
                if (!Directory.Exists(exportDir)) Directory.CreateDirectory(exportDir);

                string[] targetKeywords = new string[] {
                    "彰濱廠異常改善單", "彰濱聯絡書", "台玻內文", "彰濱廠郵件收文", "彰濱廠虛驚事件", "輕度傷害記錄表"
                };

                Dictionary<string, List<string[]>> categorizedData = new Dictionary<string, List<string[]>>();
                categorizedData["未分類其他表單"] = new List<string[]>(); 
                foreach (var kw in targetKeywords) categorizedData[kw] = new List<string[]>();

                foreach (var row in allData) {
                    string formNo = row[0], subject = row[1];
                    bool matched = false;
                    foreach (var kw in targetKeywords) {
                        if (formNo.Contains(kw) || subject.Contains(kw)) {
                            categorizedData[kw].Add(row);
                            matched = true; break;
                        }
                    }
                    if (!matched) categorizedData["未分類其他表單"].Add(row);
                }

                // 清空選單並準備匯出與存入 TXT
                cmbCategories.Items.Clear();
                int fileCount = 0;

                foreach (var kvp in categorizedData) {
                    if (kvp.Value.Count > 0) {
                        // 1. 匯出 Excel
                        string fileName = Path.Combine(exportDir, $"{kvp.Key}_匯出_{DateTime.Now:yyyyMMdd}.xlsx");
                        await excelExporter.ExportAsync(fileName, kvp.Value);
                        fileCount++;

                        // 2. 存入 TXT 供 PDF 下載使用 (格式: 檔名單號,網址)
                        cmbCategories.Items.Add(kvp.Key);
                        string txtPath = Path.Combine(exportDir, $"Links_{kvp.Key}.txt");
                        List<string> urlsToSave = new List<string>();
                        foreach(var row in kvp.Value)
                        {
                            urlsToSave.Add($"{row[0]},{row[8]}"); // row[0]是單號，row[8]是超連結
                        }
                        File.WriteAllLines(txtPath, urlsToSave);
                    }
                }
                
                if (cmbCategories.Items.Count > 0) cmbCategories.SelectedIndex = 0;
                
                UIState(true, $"✅ 作業完成！已拆分成 {fileCount} 個 Excel。\n您可以透過下拉選單，點擊「下載 PDF」來批次儲存表單。");
            }
            catch (Exception ex) { UIState(true, $"❌ 發生錯誤：\n{ex.Message}"); }
        }

        // 【需求 2】依照 TXT 紀錄批次下載檔案
        private async void BtnDownloadPdf_Click(object sender, EventArgs e)
        {
            if (cmbCategories.SelectedItem == null)
            {
                MessageBox.Show("請先執行「批次匯入」，系統才會產生分類選單！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedCat = cmbCategories.SelectedItem.ToString();
            string txtPath = Path.Combine(GetExportPath(), $"Links_{selectedCat}.txt");

            if (!File.Exists(txtPath))
            {
                MessageBox.Show("找不到該分類的暫存網址檔，請重新執行爬蟲！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(txtPath);
                string pdfDir = Path.Combine(GetExportPath(), selectedCat + "_檔案下載");
                if (!Directory.Exists(pdfDir)) Directory.CreateDirectory(pdfDir);

                UIState(false, $"準備下載 {lines.Length} 份文件...");
                Random rnd = new Random();

                for (int i = 0; i < lines.Length; i++)
                {
                    var parts = lines[i].Split(',');
                    if (parts.Length < 2) continue;

                    string formNo = parts[0].Trim();
                    string url = parts[1].Trim();
                    if (string.IsNullOrWhiteSpace(url)) continue;

                    UIState(false, $"正在下載 ({i + 1}/{lines.Length}): 單號 {formNo}");
                    
                    // 強制附檔名為 .pdf
                    string savePath = Path.Combine(pdfDir, $"{formNo}.pdf");
                    await network.DownloadFileAsync(url, savePath);

                    if (i < lines.Length - 1)
                    {
                        await Task.Delay(rnd.Next(1000, 2000)); // 隨機等待保護伺服器
                    }
                }

                UIState(true, $"✅ {selectedCat} 所有的表單下載完成！\n已儲存於：{pdfDir}");
            }
            catch (Exception ex)
            {
                UIState(true, $"❌ 下載發生錯誤：\n{ex.Message}");
            }
        }

        private void UIState(bool enable, string msg) {
            if (this.InvokeRequired) { this.Invoke(new Action(() => UIState(enable, msg))); return; }
            btnExecute.Enabled = enable; btnSettings.Enabled = enable; btnDownloadPdf.Enabled = enable; lblStatus.Text = msg;
        }
    }
}
