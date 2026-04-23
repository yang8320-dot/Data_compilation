/*
 * 檔案功能：帳號密碼設定視窗，修正網址清單折行問題，確保視覺明確區分各行網址。
 * 對應選單名稱：設定視窗
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class SettingsForm : Form
    {
        private TextBox txtUser, txtPass, txtLoginUrl, txtCrawlUrls;
        private App_Settings settings;

        public SettingsForm(App_Settings currentSettings)
        {
            settings = currentSettings;
            InitializeUI();
            
            // 載入現有設定
            txtUser.Text = settings.Username;
            txtPass.Text = settings.Password;
            txtLoginUrl.Text = settings.LoginUrl;
            txtCrawlUrls.Text = string.Join(Environment.NewLine, settings.CrawlUrls);
        }

        private void InitializeUI()
        {
            this.Text = "系統設定 - 網址與帳密管理";
            this.ClientSize = new Size(550, 550); // 加寬視窗
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
            
            // 【關鍵修正】WordWrap = false 關閉自動換行，ScrollBars.Both 開啟雙向卷軸
            // 這樣無論網址多長，都不會被折到下一行，明確分出不同網址！
            txtCrawlUrls = new TextBox { 
                Location = new Point(130, 140), 
                Width = 380, 
                Height = 300, 
                Multiline = true, 
                ScrollBars = ScrollBars.Both, 
                WordWrap = false 
            };

            Button btnSave = new Button { 
                Text = "儲存設定", 
                Location = new Point(130, 460), 
                Size = new Size(380, 45), 
                BackColor = Color.LightSteelBlue, 
                Cursor = Cursors.Hand 
            };
            
            btnSave.Click += (s, e) => {
                settings.Username = txtUser.Text.Trim();
                settings.Password = txtPass.Text.Trim();
                settings.LoginUrl = txtLoginUrl.Text.Trim();
                
                // 確保使用 \r\n 或 \n 都能正確切割出每一行網址
                settings.CrawlUrls = new List<string>(txtCrawlUrls.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                settings.Save();
                
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] { txtUser, txtPass, txtLoginUrl, txtCrawlUrls, btnSave });
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
