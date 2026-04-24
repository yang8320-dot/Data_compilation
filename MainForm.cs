/*
 * 檔案功能：應用程式主視窗，支援執行狀況總結彈窗與登入失敗警告。
 * 對應選單名稱：主選單
 */
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
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
        private string lastExportPath = @"D:\Tgeoffice"; 

        public MainForm()
        {
            settings = new App_Settings();
            network = new App_Network();
            crawler = new App_Crawler();
            excelExporter = new App_ExcelExporter();
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "經手表單自動化工具 - 專業版";
            this.Size = new Size(850, 500); // 稍微加寬
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);
            this.BackColor = Color.WhiteSmoke;

            // 上方選單區
            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 帳密與網址設定", Location = new Point(15, 15), Size = new Size(180, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => new SettingsForm(settings).ShowDialog();
            menuPanel.Controls.Add(btnSettings);

            // 執行按鈕
            btnExecute = new Button { Text = "🚀 開始執行批次作業", Location = new Point(30, 90), Size = new Size(250, 60), Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            // 資料夾按鈕
            btnOpenFolder = new Button { Text = "📁 開啟下載資料夾", Location = new Point(300, 90), Size = new Size(200, 60), Cursor = Cursors.Hand };
            btnOpenFolder.Click += (s, e) => {
                if (Directory.Exists(lastExportPath)) Process.Start("explorer.exe", lastExportPath);
                else MessageBox.Show("資料夾不存在！", "提示");
            };

            // 💡 狀態標籤：支援多行、固定範圍、字體清晰
            lblStatus = new Label { 
                Text = "系統就緒。\n請點擊上方設定帳密與網址清單後，點擊「執行」按鈕。", 
                Location = new Point(30, 180), 
                Size = new Size(780, 250), 
                AutoSize = false, 
                ForeColor = Color.DarkSlateGray,
                Font = new Font("Microsoft JhengHei", 11F),
                TextAlign = ContentAlignment.TopLeft
            };

            this.Controls.Add(btnExecute);
            this.Controls.Add(btnOpenFolder);
            this.Controls.Add(lblStatus);
            this.Controls.Add(menuPanel);
        }

        private async void BtnExecute_Click(object sender, EventArgs e)
        {
            if (settings.CrawlUrls.Count == 0) {
                MessageBox.Show("請先在設定中輸入爬蟲網址清單。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            StringBuilder summary = new StringBuilder();
            int successPages = 0;
            int totalRecords = 0;

            try {
                UIState(false, "【第一步】\n系統連線中：正在檢查登入狀態...");
                
                bool login = await network.LoginAsync(settings.Username, settings.Password);
                
                // 💡 登入失敗處理：跳出彈窗警告
                if (!login) { 
                    UIState(true, "❌ 登入失敗！\n請確認您的帳號與密碼是否正確。\n(建議點擊右上角「帳密設定」重新檢查)");
                    MessageBox.Show("登入失敗！\n請確認您的 EIP 帳號與密碼是否輸入正確。", "安全驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return; 
                }

                List<string[]> allData = new List<string[]>();
                Random rnd = new Random();

                UIState(false, "【第二步】\n開始批次抓取網頁資料...");

                for (int i = 0; i < settings.CrawlUrls.Count; i++)
                {
                    string url = settings.CrawlUrls[i];
                    UIState(false, $"正在處理 (第 {i + 1}/{settings.CrawlUrls.Count} 頁)...\n網址：{url}");
                    
                    try {
                        string html = await network.GetHtmlAsync(url);
                        var data = await crawler.ParseHtmlContentAsync(html);
                        allData.AddRange(data);
                        
                        successPages++;
                        totalRecords += data.Count;
                        summary.AppendLine($"- 第 {i+1} 頁：成功 (抓取 {data.Count} 筆)");
                    } catch (Exception ex) {
                        summary.AppendLine($"- 第 {i+1} 頁：失敗 ({ex.Message})");
                    }

                    if (i < settings.CrawlUrls.Count - 1)
                    {
                        int waitTime = rnd.Next(2000, 4001); 
                        UIState(false, $"第 {i + 1} 頁處理成功，共 {allData.Count} 筆。\n模擬人工操作中，等待 {waitTime / 1000.0:F1} 秒後切換網頁...");
                        await Task.Delay(waitTime);
                    }
                }

                // 匯出 Excel
                string fileName = Path.Combine(lastExportPath, $"批次匯出_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
                if (!Directory.Exists(lastExportPath)) Directory.CreateDirectory(lastExportPath);

                UIState(false, "【第三步】\n抓取完畢，正在封裝 Excel 報表...");
                await excelExporter.ExportAsync(fileName, allData);
                
                string finalMsg = $"✅ 作業完全結束！\n總共處理：{settings.CrawlUrls.Count} 個網址\n成功抓取：{successPages} 頁\n累積筆數：{totalRecords} 筆\n存檔位置：{fileName}";
                UIState(true, finalMsg);

                // 💡 執行狀況彈窗
                MessageBox.Show(finalMsg, "網頁執行狀況總結", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { 
                UIState(true, $"❌ 發生非預期錯誤：\n{ex.Message}");
                MessageBox.Show($"程式執行中斷：\n{ex.Message}", "異常錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UIState(bool enable, string msg) {
            if (this.InvokeRequired) {
                this.Invoke(new Action(() => UIState(enable, msg)));
                return;
            }
            btnExecute.Enabled = enable;
            btnSettings.Enabled = enable;
            lblStatus.Text = msg;
            // 發生失敗時將文字改為紅色
            lblStatus.ForeColor = msg.Contains("❌") ? Color.Red : Color.DarkSlateGray;
        }
    }
}
