/*
 * 檔案功能：應用程式主視窗。支援視窗自動縮放、滑鼠選取狀態列、自訂 PDF 下載網址參數。
 */
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class MainForm : Form
    {
        private Button btnExecute, btnSettings, btnDbSettings, btnOpenFolder, btnDownloadPdf;
        private ComboBox cmbCategories;
        private TextBox txtStatus; // 改用 TextBox 讓文字可以被選取
        private App_Settings settings;
        private App_DbSettings dbSettings;
        private App_Network network;
        private App_Crawler crawler;
        private App_ExcelExporter excelExporter;
        private App_Database appDatabase;

        public MainForm()
        {
            settings = new App_Settings();
            dbSettings = App_DbSettings.Load();
            network = new App_Network();
            crawler = new App_Crawler();
            excelExporter = new App_ExcelExporter();
            appDatabase = new App_Database();
            InitializeUI();
        }

        private string GetExportPath()
        {
            return Path.Combine(Application.StartupPath, "ExcelExport");
        }

        private void CleanExcelFolder(string targetPath)
        {
            if (Directory.Exists(targetPath))
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(targetPath);
                    foreach (FileInfo file in di.GetFiles()) file.Delete();
                }
                catch { /* 忽略鎖定的檔案 */ }
            }
            else
            {
                Directory.CreateDirectory(targetPath);
            }
        }

        private void InitializeUI()
        {
            this.Text = "經手表單自動化工具";
            this.Size = new Size(550, 500); 
            this.MinimumSize = new Size(550, 500); // 避免視窗縮太小
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 網址與帳密設定", Location = new Point(15, 15), Size = new Size(160, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => {
                if (new SettingsForm(settings).ShowDialog() == DialogResult.OK) { settings.Load(); }
            };
            
            btnDbSettings = new Button { Text = "🗄️ 資料庫寫入設定", Location = new Point(190, 15), Size = new Size(160, 40), Cursor = Cursors.Hand };
            btnDbSettings.Click += (s, e) => {
                if (new DbSettingsForm(dbSettings).ShowDialog() == DialogResult.OK) { dbSettings = App_DbSettings.Load(); }
            };

            menuPanel.Controls.Add(btnSettings);
            menuPanel.Controls.Add(btnDbSettings);

            btnExecute = new Button { Text = "🚀 1. 執行登入並批次匯入", Location = new Point(30, 90), Size = new Size(250, 50), Font = new Font("Microsoft JhengHei", 11F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            btnOpenFolder = new Button { Text = "📁 開啟 Excel 資料夾", Location = new Point(300, 90), Size = new Size(200, 50), Cursor = Cursors.Hand, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            btnOpenFolder.Click += (s, e) => {
                string currentPath = GetExportPath();
                if (!Directory.Exists(currentPath)) Directory.CreateDirectory(currentPath);
                Process.Start("explorer.exe", currentPath);
            };

            cmbCategories = new ComboBox { Location = new Point(30, 160), Size = new Size(250, 30), DropDownStyle = ComboBoxStyle.DropDownList };
            
            btnDownloadPdf = new Button { Text = "📄 2. 依選單下載 PDF", Location = new Point(300, 155), Size = new Size(200, 40), Font = new Font("Microsoft JhengHei", 10F, FontStyle.Bold), BackColor = Color.LightSkyBlue, Cursor = Cursors.Hand, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            btnDownloadPdf.Click += BtnDownloadPdf_Click;

            // 【修改點】：改為唯讀 TextBox，可以滑鼠選取複製，且會隨視窗縮放
            txtStatus = new TextBox { 
                Text = "系統就緒。執行時將會先清空 Excel 資料夾。\n注意：請先確認各類別是否開啟「資料庫寫入設定」。", 
                Location = new Point(30, 220), 
                Size = new Size(470, 210), 
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = true,
                BackColor = Color.WhiteSmoke,
                ForeColor = Color.DarkSlateGray, 
                Font = new Font("Microsoft JhengHei", 11F),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            
            this.Controls.Add(btnExecute); 
            this.Controls.Add(btnOpenFolder);
            this.Controls.Add(cmbCategories); 
            this.Controls.Add(btnDownloadPdf);
            this.Controls.Add(txtStatus); 
            this.Controls.Add(menuPanel);
        }

        private async void BtnExecute_Click(object sender, EventArgs e)
        {
            if (settings.CrawlUrls.Count == 0) { MessageBox.Show("請先設定爬蟲網址清單。"); return; }

            string exportDir = GetExportPath();
            CleanExcelFolder(exportDir); 

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
                        await Task.Delay(rnd.Next(2000, 4001));
                    }
                }

                string[] targetKeywords = new string[] {
                    "彰濱廠異常改善單", "彰濱聯絡書", "台玻內文", "彰濱廠郵件收文", "彰濱廠虛驚事件輕度傷害記錄表"
                };
                Dictionary<string, List<string[]>> categorizedData = new Dictionary<string, List<string[]>>();
                categorizedData["未分類其他表單"] = new List<string[]>(); 
                foreach (var kw in targetKeywords) categorizedData[kw] = new List<string[]>();
                
                foreach (var row in allData) {
                    string formNo = row[0], subject = row[2];
                    bool matched = false;
                    foreach (var kw in targetKeywords) {
                        if (formNo.Contains(kw) || subject.Contains(kw)) {
                            categorizedData[kw].Add(row);
                            matched = true; break;
                        }
                    }
                    if (!matched) categorizedData["未分類其他表單"].Add(row);
                }

                cmbCategories.Items.Clear();
                int fileCount = 0;
                int dbWriteCount = 0;

                UIState(false, "正在匯出 Excel 並寫入資料庫(若有設定)...");

                foreach (var kvp in categorizedData) {
                    if (kvp.Value.Count > 0) {
                        string fileName = Path.Combine(exportDir, $"{kvp.Key}_匯出_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                        await excelExporter.ExportAsync(fileName, kvp.Value);
                        fileCount++;

                        cmbCategories.Items.Add(kvp.Key);
                        string txtPath = Path.Combine(exportDir, $"Links_{kvp.Key}.txt");
                        List<string> urlsToSave = new List<string>();
                        foreach(var row in kvp.Value)
                        {
                            // 儲存格式: 單號(0) | 主題(2) | 網址(10)
                            urlsToSave.Add($"{row[0]}|{row[2]}|{row[10]}");
                        }
                        File.WriteAllLines(txtPath, urlsToSave);

                        var dbConfig = dbSettings.Categories.FirstOrDefault(c => c.CategoryName == kvp.Key);
                        if (dbConfig != null && dbConfig.IsEnabled)
                        {
                            try {
                                appDatabase.ProcessCategoryData(dbConfig, kvp.Value);
                                dbWriteCount++;
                            }
                            catch (Exception dbEx) {
                                MessageBox.Show($"寫入資料庫失敗 [{kvp.Key}]：\n{dbEx.Message}");
                            }
                        }
                    }
                }
                
                if (cmbCategories.Items.Count > 0) cmbCategories.SelectedIndex = 0;
                UIState(true, $"✅ 作業完成！已拆分成 {fileCount} 個 Excel。\n已連動寫入 {dbWriteCount} 個類別至資料庫。");
            }
            catch (Exception ex) { UIState(true, $"❌ 發生錯誤：\n{ex.Message}"); }
        }

        private async void BtnDownloadPdf_Click(object sender, EventArgs e)
        {
            if (cmbCategories.SelectedItem == null)
            {
                MessageBox.Show("請先執行「批次匯入」！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedCat = cmbCategories.SelectedItem.ToString();
            string txtPath = Path.Combine(GetExportPath(), $"Links_{selectedCat}.txt");

            if (!File.Exists(txtPath))
            {
                MessageBox.Show("找不到暫存網址檔！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                string[] lines = File.ReadAllLines(txtPath);
                string pdfDir = Path.Combine(GetExportPath(), selectedCat + "_檔案下載");
                if (!Directory.Exists(pdfDir)) Directory.CreateDirectory(pdfDir);

                UIState(false, $"準備下載 {lines.Length} 份文件...");
                Random rnd = new Random();
                int successCount = 0;

                for (int i = 0; i < lines.Length; i++)
                {
                    var parts = lines[i].Split('|');
                    if (parts.Length < 3) continue;

                    string formNo = parts[0].Trim();
                    string subject = parts[1].Trim();
                    string url = parts[2].Trim();
                    if (string.IsNullOrWhiteSpace(url)) continue;

                    // 【修改點】：將下載路徑替換成 print_version
                    string realPdfUrl = url.Replace("view_frameset", "print_version")
                                           .Replace("print_frameset", "print_version")
                                           .Replace("view_formsflow", "print_version")
                                           .Replace("show_formsflow_list", "print_version");

                    // 【修改點】：自動補上列印版 PDF 參數
                    if (!realPdfUrl.Contains("print_style=0"))
                    {
                        realPdfUrl += "&print_style=0&portrait=1&display_qrcode=";
                    }

                    if (realPdfUrl.StartsWith("/")) 
                    {
                        realPdfUrl = "http://192.168.1.83" + realPdfUrl;
                    } 
                    else if (!realPdfUrl.StartsWith("http")) 
                    {
                        realPdfUrl = "http://192.168.1.83/eipplus/" + realPdfUrl.TrimStart('/');
                    }

                    string safeSubject = string.Concat(subject.Split(Path.GetInvalidFileNameChars()));
                    string savePath = Path.Combine(pdfDir, $"{formNo}_{safeSubject}.pdf");
                    
                    UIState(false, $"正在下載表單 PDF ({i + 1}/{lines.Length}):\n{formNo} - {subject}");
                    
                    try 
                    {
                        await network.DownloadFileAsync(realPdfUrl, savePath);

                        // 驗證下載檔案 (如果 print_version API 產生的是網頁 HTML 而非二進制 PDF，會自動標記為 .html)
                        if (File.Exists(savePath))
                        {
                            byte[] buffer = new byte[4];
                            using (FileStream fs = new FileStream(savePath, FileMode.Open, FileAccess.Read))
                            {
                                fs.Read(buffer, 0, 4);
                            }
                            string header = System.Text.Encoding.ASCII.GetString(buffer);
                            
                            if (header != "%PDF")
                            {
                                File.Move(savePath, savePath + "_列印版.html"); // 若伺服器輸出 HTML 就保存為 HTML
                            }
                            successCount++;
                        }
                    }
                    catch (Exception dlEx)
                    {
                        System.Diagnostics.Debug.WriteLine($"下載失敗 {formNo}: {dlEx.Message}");
                    }

                    // 避免被伺服器阻擋，加入隨機延遲
                    if (i < lines.Length - 1)
                    {
                        await Task.Delay(rnd.Next(1500, 3000));
                    }
                }

                UIState(true, $"✅ {selectedCat} 下載完成！共成功下載 {successCount}/{lines.Length} 份。\n已儲存於：{pdfDir}");
            }
            catch (Exception ex)
            {
                UIState(true, $"❌ 下載發生錯誤：\n{ex.Message}");
            }
        }

        private void UIState(bool enable, string msg) {
            if (this.InvokeRequired) { this.Invoke(new Action(() => UIState(enable, msg))); return; }
            btnExecute.Enabled = enable; 
            btnSettings.Enabled = enable;
            btnDbSettings.Enabled = enable;
            btnDownloadPdf.Enabled = enable; 
            
            // UI 更新 (改為 AppendText + 捲動到最底，方便觀察進度)
            if (enable) txtStatus.Text = msg; 
            else {
                txtStatus.AppendText(Environment.NewLine + ">>> " + msg);
                txtStatus.SelectionStart = txtStatus.Text.Length;
                txtStatus.ScrollToCaret();
            }
        }
    }
}
