/*
 * 檔案功能：系統設定視窗，新增自訂存檔路徑的 UI 介面。
 * 對應選單名稱：設定視窗
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class SettingsForm : Form
    {
        private TextBox txtUser, txtPass, txtLoginUrl, txtCrawlUrls, txtExportPath;
        private App_Settings settings;

        public SettingsForm(App_Settings currentSettings)
        {
            settings = currentSettings;
            InitializeUI();
            
            txtUser.Text = settings.Username;
            txtPass.Text = settings.Password;
            txtLoginUrl.Text = settings.LoginUrl;
            txtCrawlUrls.Text = string.Join(Environment.NewLine, settings.CrawlUrls);
            txtExportPath.Text = settings.ExportPath;
        }

        private void InitializeUI()
        {
            this.Text = "系統設定 - 網址與帳密管理";
            this.ClientSize = new Size(550, 620); // 視窗加高，容納新設定
            this.StartPosition = FormStartPosition.CenterParent;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);
            this.BackColor = Color.WhiteSmoke;

            CreateLabel("EIP 帳號：", 20, 20);
            txtUser = new TextBox { Location = new Point(130, 17), Width = 380 };

            CreateLabel("EIP 密碼：", 20, 60);
            txtPass = new TextBox { Location = new Point(130, 57), Width = 380, PasswordChar = '*' };

            CreateLabel("登入首頁：", 20, 100);
            txtLoginUrl = new TextBox { Location = new Point(130, 97), Width = 380 };

            CreateLabel("爬蟲網址清單：\n(一行一個網址)", 20, 140);
            txtCrawlUrls = new TextBox { 
                Location = new Point(130, 140), 
                Width = 380, 
                Height = 300, 
                Multiline = true, 
                ScrollBars = ScrollBars.Both, 
                WordWrap = false 
            };

            // 新增：存檔路徑設定區塊
            CreateLabel("存檔路徑：\n(留空為程式所在目錄)", 20, 460);
            txtExportPath = new TextBox { Location = new Point(130, 460), Width = 280 };
            
            Button btnBrowse = new Button { Text = "選擇資料夾", Location = new Point(420, 458), Size = new Size(90, 28), Cursor = Cursors.Hand };
            btnBrowse.Click += (s, e) => {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog()) {
                    fbd.Description = "請選擇 Excel 檔案要匯出的資料夾：";
                    if (!string.IsNullOrWhiteSpace(txtExportPath.Text) && System.IO.Directory.Exists(txtExportPath.Text))
                        fbd.SelectedPath = txtExportPath.Text;

                    if (fbd.ShowDialog() == DialogResult.OK) {
                        txtExportPath.Text = fbd.SelectedPath;
                    }
                }
            };

            // 儲存按鈕位置往下移
            Button btnSave = new Button { 
                Text = "儲存設定", 
                Location = new Point(130, 520), 
                Size = new Size(380, 45), 
                BackColor = Color.LightSteelBlue, 
                Cursor = Cursors.Hand 
            };
            
            btnSave.Click += (s, e) => {
                settings.Username = txtUser.Text.Trim();
                settings.Password = txtPass.Text.Trim();
                settings.LoginUrl = txtLoginUrl.Text.Trim();
                settings.CrawlUrls = new List<string>(txtCrawlUrls.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                settings.ExportPath = txtExportPath.Text.Trim(); // 儲存路徑
                
                settings.Save();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] { txtUser, txtPass, txtLoginUrl, txtCrawlUrls, txtExportPath, btnBrowse, btnSave });
        }

        private void CreateLabel(string text, int x, int y)
        {
            Label lbl = new Label { 
                Text = text, 
                Location = new Point(x, y), 
                Size = new Size(110, 40), 
                TextAlign = ContentAlignment.TopLeft 
            };
            this.Controls.Add(lbl);
        }
    }
}
