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
        private string lastExportPath = @"D:\Tgeoffice"; // 預設資料夾

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

            // 上方選單區
            Panel menuPanel = new Panel { Dock = DockStyle.Top, Height = 70, BackColor = Color.LightSteelBlue };
            btnSettings = new Button { Text = "⚙️ 帳密與網址設定", Location = new Point(15, 15), Size = new Size(180, 40), Cursor = Cursors.Hand };
            btnSettings.Click += (s, e) => new SettingsForm(settings).ShowDialog();
            
            menuPanel.Controls.Add(btnSettings);

            // 主內容區
            btnExecute = new Button { Text = "🚀 執行登入並批次匯入", Location = new Point(30, 100), Size = new Size(250, 60), Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold), BackColor = Color.PaleGreen, Cursor = Cursors.Hand };
            btnExecute.Click += BtnExecute_Click;

            btnOpenFolder = new Button { Text = "📁 開啟下載資料夾", Location = new Point(300, 100), Size = new Size(200, 60), Cursor = Cursors.Hand };
            btnOpenFolder.Click += (s, e) => {
                if (Directory.Exists(lastExportPath)) Process.Start("explorer.exe", lastExportPath);
                else MessageBox.Show("資料夾不存在！");
            };

            lblStatus = new Label { Text = "系統就緒。請確認網址清單後執行。", Location = new Point(30, 200), AutoSize = true, ForeColor = Color.Gray };

            this.Controls.Add(btnExecute);
            this.Controls.Add(btnOpenFolder);
            this.Controls.Add(lblStatus);
            this.Controls.Add(menuPanel);
        }

        private async void BtnExecute_Click(object sender, EventArgs e)
        {
            if (settings.CrawlUrls.Count == 0) {
                MessageBox.Show("請先在設定中輸入爬蟲網址清單。");
                return;
            }

            try {
                UIState(false, "正在登入系統...");
                bool login = await network.LoginAsync(settings.Username, settings.Password); // 內部會讀取 settings.LoginUrl
                if (!login) { UIState(true, "登入失敗！"); return; }

                List<string[]> allData = new List<string[]>();
                foreach (string url in settings.CrawlUrls)
                {
                    UIState(false, $"正在抓取：{url}");
                    // 這裡模擬網路抓取內容，App_Network 需要新增 GetHtmlAsync 方法
                    string html = await network.GetHtmlAsync(url);
                    var data = await crawler.ParseHtmlContentAsync(html);
                    allData.AddRange(data);
                }

                string fileName = Path.Combine(lastExportPath, $"批次匯出_{DateTime.Now:yyyyMMdd_HHmm}.xlsx");
                if (!Directory.Exists(lastExportPath)) Directory.CreateDirectory(lastExportPath);

                await excelExporter.ExportAsync(fileName, allData);
                UIState(true, $"完成！已匯出 {allData.Count} 筆資料至 Excel。");
            }
            catch (Exception ex) { UIState(true, $"發生錯誤：{ex.Message}"); }
        }

        private void UIState(bool enable, string msg) {
            btnExecute.Enabled = enable;
            btnSettings.Enabled = enable;
            lblStatus.Text = msg;
        }
    }
}
