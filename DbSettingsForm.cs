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

        public DbSettingsForm(App_DbSettings settings)
        {
            dbSettings = settings;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "資料庫寫入設定";
            this.Size = new Size(750, 750); // 寬度加大 100，高度微調
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
                Location = new Point(180, 650), Size = new Size(380, 45),
                BackColor = Color.LightSteelBlue, Cursor = Cursors.Hand
            };
            btnSave.Click += (s, e) => {
                dbSettings.Save();
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.Add(tabControl);
            this.Controls.Add(btnSave);
        }

        private void BuildCategoryPanel(TabPage page, CategoryDbSetting config)
        {
            int y = 15;
            int labelX = 20;
            int controlX = 180; // 統一將控制項往右推，避免蓋到文字

            // 1. 是否寫入
            Label lblEnable = new Label { Text = "是否寫入此類別：", Location = new Point(labelX, y), AutoSize = true };
            ComboBox cmbEnable = new ComboBox { Location = new Point(controlX, y), DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cmbEnable.Items.AddRange(new[] { "寫入", "不寫入" });
            cmbEnable.SelectedIndex = config.IsEnabled ? 0 : 1;
            cmbEnable.SelectedIndexChanged += (s, e) => { config.IsEnabled = cmbEnable.SelectedIndex == 0; };
            page.Controls.AddRange(new Control[] { lblEnable, cmbEnable });
            y += 40;

            // 2. 寫入主資料庫
            Label lblDb = new Label { Text = "寫入主庫(SQLite)：", Location = new Point(labelX, y), AutoSize = true };
            TextBox txtDb = new TextBox { Text = config.DbFilePath, Location = new Point(controlX, y), Width = 380, ReadOnly = true };
            Button btnDbBrowse = new Button { Text = "瀏覽", Location = new Point(570, y - 1), Width = 60, Cursor = Cursors.Hand };
            Button btnDbLoad = new Button { Text = "讀取", Location = new Point(640, y - 1), Width = 60, Cursor = Cursors.Hand, BackColor = Color.PaleGreen };
            page.Controls.AddRange(new Control[] { lblDb, txtDb, btnDbBrowse, btnDbLoad });
            y += 40;

            // 3. 主資料表
            ComboBox cmbTable = new ComboBox { Location = new Point(controlX, y), Width = 380, DropDownStyle = ComboBoxStyle.DropDownList };
            Label lblTable = new Label { Text = "寫入主資料表：", Location = new Point(labelX, y), AutoSize = true };
            page.Controls.AddRange(new Control[] { lblTable, cmbTable });
            y += 40;

            // 4. 獨立排除資料庫
            Label lblExDb = new Label { Text = "排除清單(SQLite)：", Location = new Point(labelX, y), AutoSize = true, ForeColor = Color.Brown };
            TextBox txtExDb = new TextBox { Text = config.ExcludeDbFilePath, Location = new Point(controlX, y), Width = 380, ReadOnly = true };
            Button btnExDbBrowse = new Button { Text = "瀏覽", Location = new Point(570, y - 1), Width = 60, Cursor = Cursors.Hand };
            Button btnExDbLoad = new Button { Text = "讀取", Location = new Point(640, y - 1), Width = 60, Cursor = Cursors.Hand, BackColor = Color.LightYellow };
            page.Controls.AddRange(new Control[] { lblExDb, txtExDb, btnExDbBrowse, btnExDbLoad });
            y += 40;

            // 5. 排除資料表與欄位
            Label lblEx = new Label { Text = "排除資料表/欄位：", Location = new Point(labelX, y), AutoSize = true, ForeColor = Color.Brown };
            ComboBox cmbExTable = new ComboBox { Location = new Point(controlX, y), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            ComboBox cmbExCol = new ComboBox { Location = new Point(380, y), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            page.Controls.AddRange(new Control[] { lblEx, cmbExTable, cmbExCol });
            y += 50;

            // 6. 爬蟲欄位對應 Panel
            List<ComboBox> colMappingCmbs = new List<ComboBox>();
            Panel mappingPanel = new Panel { Location = new Point(labelX, y), Size = new Size(680, 320), BorderStyle = BorderStyle.FixedSingle };
            int my = 10;
            foreach (var field in scrapeHeaders)
            {
                Label lblF = new Label { Text = $"爬蟲 [{field}] 寫入：", Location = new Point(10, my+3), AutoSize = true, ForeColor = Color.DarkBlue };
                ComboBox cmbF = new ComboBox { Location = new Point(180, my), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
                
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == field);
                if (existMap != null) cmbF.Tag = existMap.DbColumn; 

                colMappingCmbs.Add(cmbF);
                mappingPanel.Controls.AddRange(new Control[] { lblF, cmbF });
                my += 33;
            }
            page.Controls.Add(mappingPanel);

            // ====== 邏輯事件綁定 ======

            // 主資料庫更新欄位下拉選項
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

            // 讀取主資料庫表單
            Action LoadMainTables = () => {
                var tables = App_Database.GetTables(txtDb.Text);
                tables.Insert(0, "");
                cmbTable.Items.Clear(); cmbTable.Items.AddRange(tables.ToArray());
                if (tables.Contains(config.TargetTable)) cmbTable.SelectedIndex = tables.IndexOf(config.TargetTable);
            };

            // 讀取排除資料庫表單
            Action LoadExcludeTables = () => {
                var tables = App_Database.GetTables(txtExDb.Text);
                tables.Insert(0, "");
                cmbExTable.Items.Clear(); cmbExTable.Items.AddRange(tables.ToArray());
                if (tables.Contains(config.ExcludeTable)) cmbExTable.SelectedIndex = tables.IndexOf(config.ExcludeTable);
            };

            btnDbBrowse.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "SQLite|*.sqlite;*.db3;*.db|所有|*.*" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) { txtDb.Text = config.DbFilePath = ofd.FileName; }
                }
            };

            btnDbLoad.Click += (s, e) => { LoadMainTables(); MessageBox.Show("主資料庫架構讀取完成。"); };

            cmbTable.SelectedIndexChanged += (s, e) => {
                config.TargetTable = cmbTable.SelectedItem?.ToString() ?? "";
                UpdateColumnLists(config.TargetTable);
            };

            btnExDbBrowse.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "SQLite|*.sqlite;*.db3;*.db|所有|*.*" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) { txtExDb.Text = config.ExcludeDbFilePath = ofd.FileName; }
                }
            };

            btnExDbLoad.Click += (s, e) => { LoadExcludeTables(); MessageBox.Show("排除清單庫架構讀取完成。"); };

            cmbExTable.SelectedIndexChanged += (s, e) => {
                config.ExcludeTable = cmbExTable.SelectedItem?.ToString() ?? "";
                var exCols = App_Database.GetColumns(txtExDb.Text, config.ExcludeTable);
                exCols.Insert(0, "");
                cmbExCol.Items.Clear(); cmbExCol.Items.AddRange(exCols.ToArray());
                if (exCols.Contains(config.ExcludeColumn)) cmbExCol.SelectedIndex = exCols.IndexOf(config.ExcludeColumn);
            };
            
            cmbExCol.SelectedIndexChanged += (s, e) => { config.ExcludeColumn = cmbExCol.SelectedItem?.ToString() ?? ""; };

            // 初始載入（如果有預設檔）
            if (!string.IsNullOrEmpty(txtDb.Text)) LoadMainTables();
            if (!string.IsNullOrEmpty(txtExDb.Text)) LoadExcludeTables();
        }
    }
}
