/*
 * 檔案功能：應用程式主視窗，採用 Code-First 動態生成控制項。
 * 對應選單名稱：主選單
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class MainForm : Form
    {
        private Panel menuPanel;
        private Panel contentPanel;
        private Button btnSelectHtml;
        private Button btnProcessAndExport;
        private Label lblStatus;
        private TextBox txtSelectedFile;

        private App_Crawler crawler;
        private App_TxtStorage txtStorage;
        private App_ExcelExporter excelExporter;

        public MainForm()
        {
            crawler = new App_Crawler();
            txtStorage = new App_TxtStorage();
            excelExporter = new App_ExcelExporter();

            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "經手表單解析工具";
            this.Size = new Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Font = new Font("Microsoft JhengHei", 10F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(136)));
            this.BackColor = Color.WhiteSmoke;

            // 主選單面板 (間距 15，框內與文字間距 10)
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
            menuPanel.Controls.Add(titleLabel);

            // 內容區塊面板 (與主選單間隔 15)
            contentPanel = new Panel
            {
                Location = new Point(15, menuPanel.Height + 15),
                Width = this.ClientSize.Width - 30,
                Height = this.ClientSize.Height - menuPanel.Height - 30,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            // 控制項排版
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
                ReadOnly = true
            };

            btnProcessAndExport = new Button
            {
                Text = "執行解析並匯出 Excel",
                Location = new Point(15, 70),
                Size = new Size(200, 40),
                Cursor = Cursors.Hand,
                Enabled = false
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
            
            // 視窗縮放事件以維持動態排版
            this.Resize += (s, e) => 
            {
                contentPanel.Width = this.ClientSize.Width - 30;
                contentPanel.Height = this.ClientSize.Height - menuPanel.Height - 30;
            };
        }

        private void BtnSelectHtml_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "HTML 檔案 (*.html)|*.html|所有檔案 (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    txtSelectedFile.Text = ofd.FileName;
                    btnProcessAndExport.Enabled = true;
                }
            }
        }

        private async void BtnProcessAndExport_Click(object sender, EventArgs e)
        {
            string sourceFile = txtSelectedFile.Text;
            if (string.IsNullOrEmpty(sourceFile)) return;

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
                UIState(false, "正在解析 HTML 檔案...");

                // 1. 爬蟲解析
                List<string[]> parsedData = await crawler.ParseHtmlAsync(htmlPath);
                
                if (parsedData.Count == 0)
                {
                    UIState(true, "未找到符合的資料，請確認 HTML 結構。");
                    return;
                }

                // 2. 存入 Txt 作為中繼 (遵守無資料庫規範)
                UIState(false, "資料寫入中繼文字檔 (.txt)...");
                txtStorage.SaveData(parsedData);

                // 3. 從 Txt 讀取並匯出 Excel
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

        // 執行緒安全的 UI 更新方法
        private void UIState(bool isEnable, string statusMessage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => UIState(isEnable, statusMessage)));
                return;
            }
            
            btnSelectHtml.Enabled = isEnable;
            btnProcessAndExport.Enabled = isEnable && !string.IsNullOrEmpty(txtSelectedFile.Text);
            lblStatus.Text = statusMessage;
        }
    }
}
