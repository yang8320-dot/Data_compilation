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
            txtUser.Text = settings.Username;
            txtPass.Text = settings.Password;
            txtLoginUrl.Text = settings.LoginUrl;
            txtCrawlUrls.Text = string.Join(Environment.NewLine, settings.CrawlUrls);
        }

        private void InitializeUI()
        {
            this.Text = "系統設定 - 網址與帳密管理";
            this.ClientSize = new Size(500, 550);
            this.StartPosition = FormStartPosition.CenterParent;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            CreateLabel("EIP 帳號：", 20, 20);
            txtUser = new TextBox { Location = new Point(130, 17), Width = 330 };

            CreateLabel("EIP 密碼：", 20, 60);
            txtPass = new TextBox { Location = new Point(130, 57), Width = 330, PasswordChar = '*' };

            CreateLabel("登入首頁：", 20, 100);
            txtLoginUrl = new TextBox { Location = new Point(130, 97), Width = 330 };

            CreateLabel("爬蟲網址清單：\n(一行一個)", 20, 140);
            txtCrawlUrls = new TextBox { Location = new Point(130, 140), Width = 330, Height = 300, Multiline = true, ScrollBars = ScrollBars.Vertical };

            Button btnSave = new Button { Text = "儲存設定", Location = new Point(130, 460), Size = new Size(330, 45), BackColor = Color.LightSteelBlue, Cursor = Cursors.Hand };
            btnSave.Click += (s, e) => {
                settings.Username = txtUser.Text.Trim();
                settings.Password = txtPass.Text.Trim();
                settings.LoginUrl = txtLoginUrl.Text.Trim();
                settings.CrawlUrls = new List<string>(txtCrawlUrls.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries));
                settings.Save();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.AddRange(new Control[] { txtUser, txtPass, txtLoginUrl, txtCrawlUrls, btnSave });
        }

        private void CreateLabel(string text, int x, int y)
        {
            Label lbl = new Label { Text = text, Location = new Point(x, y), Size = new Size(110, 40), TextAlign = ContentAlignment.TopLeft };
            this.Controls.Add(lbl);
        }
    }
}
