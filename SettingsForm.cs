/*
 * 檔案功能：帳號密碼設定視窗，採用 Code-First 動態生成控制項。
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
            this.Text = "登入設定";
            this.Size = new Size(320, 220);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);
            this.BackColor = Color.WhiteSmoke;

            Label lblUser = new Label { Text = "EIP 帳號：", Location = new Point(20, 30), AutoSize = true };
            txtUsername = new TextBox { Location = new Point(100, 27), Width = 180 };

            Label lblPass = new Label { Text = "EIP 密碼：", Location = new Point(20, 75), AutoSize = true };
            txtPassword = new TextBox { Location = new Point(100, 72), Width = 180, PasswordChar = '*' };

            btnSave = new Button
            {
                Text = "儲存設定",
                Location = new Point(100, 120),
                Size = new Size(180, 35),
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

            settings.Save(txtUsername.Text.Trim(), txtPassword.Text.Trim());
            MessageBox.Show("設定已儲存！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
