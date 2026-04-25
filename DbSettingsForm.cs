using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    public class DbSettingsForm : Form
    {
        private App_DbSettings dbSettings;
        private TabControl tabControl;
        private string[] scrapeHeaders = { "表單單號", "表單主題", "狀態", "存檔", "承辦人", "目前處理者", "申請時間", "修改時間", "網址" };

        // 記錄每個 Tab 內的黑名單 TextBox，儲存時再一次抓取
        private Dictionary<CategoryDbSetting, TextBox> excludeTextBoxes = new Dictionary<CategoryDbSetting, TextBox>();

        public DbSettingsForm(App_DbSettings settings)
        {
            dbSettings = settings;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "資料庫寫入設定";
            // 1. 視窗加大加寬
            this.Size = new Size(900, 750); 
            this.StartPosition = FormStartPosition.CenterParent;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            tabControl = new TabControl { Dock = DockStyle.Top, Height = 630 };

            foreach (var cat in dbSettings.Categories)
            {
                TabPage page = new TabPage { Text = cat.CategoryName, BackColor = Color.White };
                BuildCategoryPanel(page, cat);
                tabControl.TabPages.Add(page);
            }

            Button btnSave = new Button {
                Text = "💾 儲存所有資料庫設定",
                Location = new Point(260, 650), Size = new Size(380, 45),
                BackColor = Color.LightSteelBlue, Cursor = Cursors.Hand
            };
            btnSave.Click += (s, e) => {
                // 儲存時將右側文字框內容回寫到設定檔
                foreach (var kvp in excludeTextBoxes)
                {
                    kvp.Key.ExcludeFormNumbers = kvp.Value.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                                                               .Select(txt => txt.Trim())
                                                               .Where(txt => !string.IsNullOrEmpty(txt))
                                                               .ToList();
                }
                dbSettings.Save();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.Add(tabControl);
            this.Controls.Add(btnSave);
        }

        private void BuildCategoryPanel(TabPage page, CategoryDbSetting config)
        {
            int y = 20;
            int labelX = 20;
            int controlX = 180; 

            // 是否寫入
            Label lblEnable = new Label { Text = "是否寫入此類別：", Location = new Point(labelX, y), AutoSize = true };
            ComboBox cmbEnable = new ComboBox { Location = new Point(controlX, y), DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cmbEnable.Items.AddRange(new[] { "寫入", "不寫入" });
            cmbEnable.SelectedIndex = config.IsEnabled ? 0 : 1;
            cmbEnable.SelectedIndexChanged += (s, e) => { config.IsEnabled = cmbEnable.SelectedIndex == 0; };
            page.Controls.AddRange(new Control[] { lblEnable, cmbEnable });
            y += 45;

            // 寫入主資料庫
            Label lblDb = new Label { Text = "寫入主庫(SQLite)：", Location = new Point(labelX, y+5), AutoSize = true };
            TextBox txtDb = new TextBox { Text = config.DbFilePath, Location = new Point(controlX, y), Width = 300, ReadOnly = true };
            
            // 2. 按鈕加大加寬 (Size 加大)
            Button btnDbBrowse = new Button { Text = "瀏覽", Location = new Point(490, y - 3), Size = new Size(80, 32), Cursor = Cursors.Hand };
            Button btnDbLoad = new Button { Text = "讀取", Location = new Point(580, y - 3), Size = new Size(80, 32), Cursor = Cursors.Hand, BackColor = Color.PaleGreen };
            
            page.Controls.AddRange(new Control[] { lblDb, txtDb, btnDbBrowse, btnDbLoad });
            y += 45;

            // 主資料表
            ComboBox cmbTable = new ComboBox { Location = new Point(controlX, y), Width = 400, DropDownStyle = ComboBoxStyle.DropDownList };
            Label lblTable = new Label { Text = "寫入主資料表：", Location = new Point(labelX, y+3), AutoSize = true };
            page.Controls.AddRange(new Control[] { lblTable, cmbTable });
            y += 50;

            // 左側：爬蟲欄位對應 Panel (面板加寬)
            List<ComboBox> colMappingCmbs = new List<ComboBox>();
            Panel mappingPanel = new Panel { Location = new Point(labelX, y), Size = new Size(600, 330), BorderStyle = BorderStyle.FixedSingle };
            int my = 15;
            foreach (var field in scrapeHeaders)
            {
                Label lblF = new Label { Text = $"爬蟲 [{field}] 寫入：", Location = new Point(15, my+4), AutoSize = true, ForeColor = Color.DarkBlue };
                
                // 3. 解決字體被遮到：將下拉選單的 X 座標從 160 往右推到 220，寬度微調
                ComboBox cmbF = new ComboBox { Location = new Point(220, my), Width = 350, DropDownStyle = ComboBoxStyle.DropDownList };
                
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == field);
                if (existMap != null) cmbF.Tag = existMap.DbColumn; 

                colMappingCmbs.Add(cmbF);
                mappingPanel.Controls.AddRange(new Control[] { lblF, cmbF });
                my += 34; // 增加垂直間距
            }
            page.Controls.Add(mappingPanel);

            // 右側：排除單號清單 TextBox
            Label lblExclude = new Label { Text = "排除寫入清單\n(每行輸入一筆表單單號)：", Location = new Point(640, 145), AutoSize = true, ForeColor = Color.Brown };
            TextBox txtExclude = new TextBox {
                Location = new Point(640, 190),
                Size = new Size(220, 275), // 高度與左方面板對齊
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false
            };
            
            // 載入舊設定
            if (config.ExcludeFormNumbers != null && config.ExcludeFormNumbers.Count > 0)
            {
                txtExclude.Text = string.Join(Environment.NewLine, config.ExcludeFormNumbers);
            }
            excludeTextBoxes[config] = txtExclude;
            
            page.Controls.Add(lblExclude);
            page.Controls.Add(txtExclude);

            // ====== 邏輯事件綁定 ======
            Action<string> UpdateColumnLists = (tableName) => {
                var cols = App_Database.GetColumns(txtDb.Text, tableName);
                cols.Insert(0, ""); 
                for (int i = 0; i < scrapeHeaders.Length; i++) {
                    var cb = colMappingCmbs[i];
                    cb.Items.Clear();
                    cb.Items.AddRange(cols.ToArray());
                    string targetStr = cb.Tag?.ToString() ?? "";
                    cb.SelectedIndex = cols.Contains(targetStr) ? cols.IndexOf(targetStr) : 0;
                    
                    int indexCopy = i;
                    cb.SelectedIndexChanged += (s, e) => {
                        var map = config.Mappings.FirstOrDefault(m => m.ScrapedField == scrapeHeaders[indexCopy]);
                        if (map == null) { map = new FieldMapping { ScrapedField = scrapeHeaders[indexCopy] }; config.Mappings.Add(map); }
                        map.DbColumn = cb.SelectedItem?.ToString() ?? "";
                        cb.Tag = map.DbColumn;
                    };
                }
            };

            Action LoadMainTables = () => {
                var tables = App_Database.GetTables(txtDb.Text);
                tables.Insert(0, "");
                cmbTable.Items.Clear(); cmbTable.Items.AddRange(tables.ToArray());
                if (tables.Contains(config.TargetTable)) cmbTable.SelectedIndex = tables.IndexOf(config.TargetTable);
            };

            btnDbBrowse.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "SQLite|*.sqlite;*.db3;*.db|所有|*.*" }) {
                    if (ofd.ShowDialog() == DialogResult.OK)
