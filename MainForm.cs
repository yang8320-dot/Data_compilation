/*
 * 檔案功能：應用程式主視窗，採用 Code-First 動態生成控制項，並整合網路登入流程。
 * 對應選單名稱：主選單
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class MainForm : Form
    {
        private Panel menuPanel;
        private Panel contentPanel;
        private Button btnSelectHtml;
        private Button btnProcessAndExport;
        private Button btnSettings;
        private Label lblStatus;
        private TextBox txtSelectedFile;

        private App_Crawler crawler;
        private App_TxtStorage txtStorage;
        private App_ExcelExporter excelExporter;
        private App_Settings settings;
        private App_Network network;

        public MainForm()
        {
            crawler = new App_Crawler();
            txtStorage = new App_TxtStorage();
            excelExporter = new App_ExcelExporter();
            settings = new App_Settings();
            network = new App_Network();

            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "經手表單解析工具";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScaleMode = AutoScaleMode.Dpi; 
            this.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));
            this.BackColor = Color.WhiteSmoke;

            // 主選單面板
            menuPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = Color.LightSteelBlue,
                Padding = new Padding(10)
            };

            Label titleLabel = new Label
            {
                Text = "系統選單",
                Font = new Font("Microsoft JhengHei", 12F, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(15, 20)
            };
            
            btnSettings = new Button
            {
                Text = "⚙️ 登入設定",
                Size = new Size(120, 35),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnSettings.Location = new Point(this.ClientSize.Width - 135, 12);
            btnSettings.Click += (s, e) => { new SettingsForm(settings).ShowDialog(); };

            menuPanel.Controls.Add(titleLabel);
            menuPanel.Controls.Add(btnSettings);

            // 內容區塊面板
            contentPanel = new Panel
            {
                Location = new Point(15, menuPanel.Height + 15),
                Width = this.ClientSize.Width - 30,
                Height = this.ClientSize.Height - menuPanel.Height - 30,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            btnSelectHtml = new Button
            {
                Text = "選擇 HTML 檔案",
                Location = new Point(15, 15),
                Size = new Size(150, 40),
                Cursor = Cursors.Hand
            };
            btnSelectHtml.Click += BtnSelectHtml_Click;

            txtSelectedFile = new TextBox
            {
                Location = new Point(175, 25),
                Width = 400,
                ReadOnly = true,
                Text = @"D:\Tgeoffice\台灣玻璃工業股份有限公司-經手表單.html"
            };

            btnProcessAndExport = new Button
            {
                Text = "執行登入並匯出 Excel",
                Location = new Point(15, 70),
                Size = new Size(200, 40),
                Cursor = Cursors.Hand,
                Enabled = true
            };
            btnProcessAndExport.Click += BtnProcessAndExport_Click;

            lblStatus = new Label
            {
                Text = "準備就緒...",
                Location = new Point(15, 130),
                AutoSize = true,
                ForeColor = Color.DarkGreen
            };

            contentPanel.Controls.Add(btnSelectHtml);
            contentPanel.Controls.Add(txtSelectedFile);
            contentPanel.Controls.Add(btnProcessAndExport);
            contentPanel.Controls.Add(lblStatus);

            this.Controls.Add(contentPanel);
            this.Controls.Add(menuPanel);
            
            this.Resize += (s, e) => 
            {
                contentPanel.Width = this.ClientSize.Width - 30;
                contentPanel.Height = this.ClientSize.Height - menuPanel.Height - 30;
                btnSettings.Location = new Point(this.ClientSize.Width - 135, 12);
            };
        }

        private void BtnSelectHtml_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "HTML 檔案 (*.html)|*.html|所有檔案 (*.*)|*.*";
                if (System.IO.File.Exists(txtSelectedFile.Text))
                {
                    ofd.InitialDirectory = System.IO.Path.GetDirectoryName(txtSelectedFile.Text);
                }
                
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtSelectedFile.Text = ofd.FileName;
                    btnProcessAndExport.Enabled = true;
                }
            }
        }

        private async void BtnProcessAndExport_Click(object sender, EventArgs e)
        {
            // 步驟 0：檢查是否已設定帳密
            if (!settings.HasCredentials())
            {
                MessageBox.Show("請先點擊右上角「登入設定」設定 EIP 帳號與密碼。", "尚未設定帳密", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                new SettingsForm(settings).ShowDialog();
                if (!settings.HasCredentials()) return; // 若使用者還是沒設定就取消
            }

            string sourceFile = txtSelectedFile.Text;
            if (string.IsNullOrEmpty(sourceFile) || !System.IO.File.Exists(sourceFile))
            {
                MessageBox.Show($"找不到指定的檔案：\n{sourceFile}\n請確認路徑是否正確。", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel 檔案 (*.xlsx)|*.xlsx";
                sfd.FileName = "表單匯出資料.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    await ProcessWorkflowAsync(sourceFile, sfd.FileName);
                }
            }
        }

        private async Task ProcessWorkflowAsync(string htmlPath, string excelPath)
        {
            try
            {
                // 步驟 1：執行 HTTP 自動登入
                UIState(false, "系統連線中：正在登入 EIP 系統...");
                bool isLoginSuccess = await network.LoginAsync(settings.Username, settings.Password);
                
                if (!isLoginSuccess)
                {
                    UIState(true, "登入失敗：帳號或密碼錯誤，請重新設定。");
                    MessageBox.Show("登入 EIP 系統失敗，請確認帳號與密碼是否正確。", "登入失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 步驟 2：爬蟲解析 (解析本地的 HTML 檔案)
                UIState(false, "登入成功！正在解析 HTML 檔案...");
                List<string[]> parsedData = await crawler.ParseHtmlAsync(htmlPath);
                
                if (parsedData.Count == 0)
                {
                    UIState(true, "未找到符合的資料，請確認 HTML 結構。");
                    return;
                }

                // 步驟 3：存入 Txt 作為中繼
                UIState(false, "資料寫入中繼文字檔 (.txt)...");
                txtStorage.SaveData(parsedData);

                // 步驟 4：從 Txt 讀取並匯出 Excel
                UIState(false, "正在生成 Excel 檔案...");
                List<string[]> storedData = txtStorage.LoadData();
                await excelExporter.ExportAsync(excelPath, storedData);

                UIState(true, "作業完成！已成功匯出至：" + excelPath);
            }
            catch (Exception ex)
            {
                UIState(true, "發生錯誤：" + ex.Message);
            }
        }

        private void UIState(bool isEnable, string statusMessage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UIState(isEnable, statusMessage)));
                return;
            }
            
            btnSelectHtml.Enabled = isEnable;
            btnProcessAndExport.Enabled = isEnable && !string.IsNullOrEmpty(txtSelectedFile.Text);
            btnSettings.Enabled = isEnable;
            lblStatus.Text = statusMessage;
        }
    }
}
