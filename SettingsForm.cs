/*
 * 檔案功能：帳號密碼設定視窗，修正 DPI 縮放導致的字體遮蔽問題。
 * 對應選單名稱：設定視窗
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class SettingsForm : Form
    {
        private TextBox txtUsername;
        private TextBox txtPassword;
        private Button btnSave;
        private App_Settings settings;

        public SettingsForm(App_Settings currentSettings)
        {
            settings = currentSettings;
            InitializeUI();
            
            // 載入現有設定
            txtUsername.Text = settings.Username;
            txtPassword.Text = settings.Password;
        }

        private void InitializeUI()
        {
            this.Text = "登入設定 (可隨時更新)";
            // 加大視窗內部尺寸，預留 DPI 放大空間
            this.ClientSize = new Size(380, 220); 
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            
            // 先設定字體，確保後續排版以此為基準
            this.Font = new Font("Microsoft JhengHei", 11F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));
            this.BackColor = Color.WhiteSmoke;

            // 帳號標籤與輸入框 (取消 AutoSize，改用固定寬高避免遮蔽)
            Label lblUser = new Label 
            { 
                Text = "EIP 帳號：", 
                Location = new Point(30, 40), 
                Size = new Size(90, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            txtUsername = new TextBox 
            { 
                Location = new Point(120, 40), 
                Width = 200 
            };

            // 密碼標籤與輸入框
            Label lblPass = new Label 
            { 
                Text = "EIP 密碼：", 
                Location = new Point(30, 90), 
                Size = new Size(90, 25),
                TextAlign = ContentAlignment.MiddleLeft
            };
            txtPassword = new TextBox 
            { 
                Location = new Point(120, 90), 
                Width = 200, 
                PasswordChar = '*' 
            };

            // 儲存按鈕
            btnSave = new Button
            {
                Text = "儲存設定",
                Location = new Point(120, 145),
                Size = new Size(200, 40),
                Cursor = Cursors.Hand,
                BackColor = Color.LightSteelBlue
            };
            btnSave.Click += BtnSave_Click;

            this.Controls.Add(lblUser);
            this.Controls.Add(txtUsername);
            this.Controls.Add(lblPass);
            this.Controls.Add(txtPassword);
            this.Controls.Add(btnSave);
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtUsername.Text) || string.IsNullOrWhiteSpace(txtPassword.Text))
            {
                MessageBox.Show("請輸入完整的帳號與密碼。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 儲存至文字檔
            settings.Save(txtUsername.Text.Trim(), txtPassword.Text.Trim());
            
            MessageBox.Show("設定已成功儲存！\n未來若密碼有變更，請隨時點擊主畫面的「帳密設定」進行更新。", "儲存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
