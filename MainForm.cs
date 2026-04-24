/*
 * 檔案功能：應用程式主視窗，將提示文字改為可選取的唯讀文字框，並支援自動捲動。
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
        private TextBox txtStatus; // 💡 將 Label 改為 TextBox 支援游標選取
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
            this.Text = "經手表單自動化工具";
            this.Size = new Size(800, 450);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 帳密與網址設定", Location = new Point(15, 15), Size = new Size(180, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => new SettingsForm(settings).ShowDialog();
            
            menuPanel.Controls.Add(btnSettings);

            btnExecute = new Button { Text = "🚀 執行登入並批次匯入", Location = new Point(30, 100), Size = new Size(250, 60), Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            btnOpenFolder = new Button { Text = "📁 開啟下載資料夾", Location = new Point(300, 100), Size = new Size(200, 60), Cursor = Cursors.Hand };
            btnOpenFolder.Click += (s, e) => {
                if (Directory.Exists(lastExportPath)) Process.Start("explorer.exe", lastExportPath);
                else MessageBox.Show("資料夾不存在！");
            };

            // 💡 【關鍵修正】使用 TextBox 偽裝成 Label，允許滑鼠選取複製文字
            txtStatus = new TextBox { 
                Text = "系統就緒。請確認網址清單後執行。", 
                Location = new Point(30, 180), 
                Size = new Size(720, 180), 
                Multiline = true, 
                ReadOnly = true,                // 不可編輯，但可游標選取複製
                BorderStyle = BorderStyle.None, // 隱藏邊框，視覺上就像一般文字
                BackColor = this.BackColor,     // 顏色與背景融合
                ForeColor = Color.DarkSlateGray,
                Font = new Font("Microsoft JhengHei", 11F),
                ScrollBars = ScrollBars.Vertical// 若內容過多自動顯示捲軸
            };

            this.Controls.Add(btnExecute);
            this.Controls.Add(btnOpenFolder);
            this.Controls.Add(txtStatus);
            this.Controls.Add(menuPanel);
        }

        private async void BtnExecute_Click(object sender, EventArgs e)
        {
            if (settings.CrawlUrls.Count == 0) {
                MessageBox.Show("請先在設定中輸入爬蟲網址清單。");
                return;
            }

            try {
                UIState(false, "系統連線中：正在自動檢查登入狀態...\r\n");
                bool login = await network.LoginAsync(settings.Username, settings.Password);
                if (!login) { UIState(true, "登入失敗！請確認帳號密碼。"); return; }

                List<string[]> allData = new List<string[]>();
                Random rnd = new Random(); 

                for (int i = 0; i < settings.CrawlUrls.Count; i++)
                {
                    string url = settings.CrawlUrls[i];
                    UIState(false, txtStatus.Text + $"正在抓取第 {i + 1}/{settings.CrawlUrls.Count} 個網頁...\r\n目標網址：{url}\r\n");
                    
                    string html = await network.GetHtmlAsync(url);
                    var data = await crawler.ParseHtmlContentAsync(html);
                    allData.AddRange(data);

                    if (i < settings.CrawlUrls.Count - 1)
                    {
                        int waitTime = rnd.Next(2000, 4001); 
                        UIState(false, txtStatus.Text + $"第 {i + 1} 頁成功！隨機等待 {waitTime / 1000.0:F1} 秒...\r\n\r\n");
                        await Task.Delay(waitTime);
                    }
                }

                string fileName = Path.Combine(lastExportPath, $"批次匯出_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
                if (!Directory.Exists(lastExportPath)) Directory.CreateDirectory(lastExportPath);

                UIState(false, txtStatus.Text + "\r\n網頁抓取完畢，正在生成 Excel 檔案...\r\n");
                await excelExporter.ExportAsync(fileName, allData);
                
                UIState(true, txtStatus.Text + $"✅ 作業完成！\r\n共成功匯出 {allData.Count} 筆資料至 Excel。\r\n檔案路徑：{fileName}\r\n");
            }
            catch (Exception ex) { UIState(true, txtStatus.Text + $"\r\n❌ 發生錯誤：\r\n{ex.Message}\r\n"); }
        }

        private void UIState(bool enable, string msg) {
            if (this.InvokeRequired) {
                this.Invoke(new Action(() => UIState(enable, msg)));
                return;
            }
            btnExecute.Enabled = enable;
            btnSettings.Enabled = enable;
            txtStatus.Text = msg;

            // 💡 讓文字框自動捲動到最底下，隨時顯示最新進度
            txtStatus.SelectionStart = txtStatus.Text.Length;
            txtStatus.ScrollToCaret();
        }
    }
}
