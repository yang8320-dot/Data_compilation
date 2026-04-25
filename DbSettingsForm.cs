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
            this.Size = new Size(650, 700);
            this.StartPosition = FormStartPosition.CenterParent;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            tabControl = new TabControl { Dock = DockStyle.Top, Height = 580 };

            foreach (var cat in dbSettings.Categories)
            {
                TabPage page = new TabPage { Text = cat.CategoryName, BackColor = Color.White };
                BuildCategoryPanel(page, cat);
                tabControl.TabPages.Add(page);
            }

            Button btnSave = new Button {
                Text = "💾 儲存所有資料庫設定",
                Location = new Point(130, 600), Size = new Size(380, 45),
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

            // 1. 下拉選單：是否寫入
            Label lblEnable = new Label { Text = "是否寫入此類別：", Location = new Point(20, y), AutoSize = true };
            ComboBox cmbEnable = new ComboBox { Location = new Point(150, y), DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cmbEnable.Items.AddRange(new[] { "寫入", "不寫入" });
            cmbEnable.SelectedIndex = config.IsEnabled ? 0 : 1;
            cmbEnable.SelectedIndexChanged += (s, e) => { config.IsEnabled = cmbEnable.SelectedIndex == 0; };
            page.Controls.AddRange(new Control[] { lblEnable, cmbEnable });
            y += 40;

            // 2. 選擇 SQLite 檔案
            Label lblDb = new Label { Text = "SQLite資料庫：", Location = new Point(20, y), AutoSize = true };
            TextBox txtDb = new TextBox { Text = config.DbFilePath, Location = new Point(150, y), Width = 300, ReadOnly = true };
            Button btnDb = new Button { Text = "瀏覽", Location = new Point(460, y-1), Width = 70 };
            page.Controls.AddRange(new Control[] { lblDb, txtDb, btnDb });
            y += 40;

            // UI 控制項宣告
            ComboBox cmbTable = new ComboBox { Location = new Point(150, y), Width = 300, DropDownStyle = ComboBoxStyle.DropDownList };
            ComboBox cmbExTable = new ComboBox { Location = new Point(150, y + 40), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            ComboBox cmbExCol = new ComboBox { Location = new Point(300, y + 40), Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            List<ComboBox> colMappingCmbs = new List<ComboBox>(); // 存放9個欄位的下拉選單

            // 建立 9 個對應欄位的 UI
            Panel mappingPanel = new Panel { Location = new Point(20, y + 80), Size = new Size(580, 320), BorderStyle = BorderStyle.FixedSingle };
            int my = 10;
            foreach (var field in scrapeHeaders)
            {
                Label lblF = new Label { Text = $"爬蟲 [{field}] 寫入：", Location = new Point(10, my+3), AutoSize = true, ForeColor = Color.DarkBlue };
                ComboBox cmbF = new ComboBox { Location = new Point(180, my), Width = 200, DropDownStyle = ComboBoxStyle.DropDownList };
                
                // 恢復已儲存的設定值
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == field);
                if (existMap != null) cmbF.Tag = existMap.DbColumn; // 暫存起來等有選項時綁定

                colMappingCmbs.Add(cmbF);
                mappingPanel.Controls.AddRange(new Control[] { lblF, cmbF });
                my += 33;
            }

            // 更新所有欄位下拉選單的副程式
            Action<string> UpdateColumnLists = (tableName) => {
                var cols = App_Database.GetColumns(txtDb.Text, tableName);
                cols.Insert(0, ""); // 第一個為空代表不對應
                for (int i = 0; i < scrapeHeaders.Length; i++) {
                    var cb = colMappingCmbs[i];
                    cb.Items.Clear();
                    cb.Items.AddRange(cols.ToArray());
                    string targetStr = cb.Tag?.ToString() ?? "";
                    cb.SelectedIndex = cols.Contains(targetStr) ? cols.IndexOf(targetStr) : 0;
                    
                    // 綁定變更事件即時存入 config
                    int indexCopy = i;
                    cb.SelectedIndexChanged += (s, e) => {
                        var map = config.Mappings.FirstOrDefault(m => m.ScrapedField == scrapeHeaders[indexCopy]);
                        if (map == null) { map = new FieldMapping { ScrapedField = scrapeHeaders[indexCopy] }; config.Mappings.Add(map); }
                        map.DbColumn = cb.SelectedItem?.ToString() ?? "";
                        cb.Tag = map.DbColumn;
                    };
                }
            };

            // 載入資料表的副程式
            Action LoadTables = () => {
                var tables = App_Database.GetTables(txtDb.Text);
                tables.Insert(0, "");
                
                cmbTable.Items.Clear(); cmbTable.Items.AddRange(tables.ToArray());
                cmbExTable.Items.Clear(); cmbExTable.Items.AddRange(tables.ToArray());

                if (tables.Contains(config.TargetTable)) cmbTable.SelectedIndex = tables.IndexOf(config.TargetTable);
                if (tables.Contains(config.ExcludeTable)) cmbExTable.SelectedIndex = tables.IndexOf(config.ExcludeTable);
            };

            btnDb.Click += (s, e) => {
                using (OpenFileDialog ofd = new OpenFileDialog { Filter = "SQLite 檔案|*.sqlite;*.db3;*.db|所有檔案|*.*" }) {
                    if (ofd.ShowDialog() == DialogResult.OK) {
                        txtDb.Text = config.DbFilePath = ofd.FileName;
                        LoadTables();
                    }
                }
            };

            // 3. 主資料表選擇
            Label lblTable = new Label { Text = "寫入主資料表：", Location = new Point(20, y), AutoSize = true };
            cmbTable.SelectedIndexChanged += (s, e) => {
                config.TargetTable = cmbTable.SelectedItem?.ToString() ?? "";
                UpdateColumnLists(config.TargetTable);
            };
            page.Controls.AddRange(new Control[] { lblTable, cmbTable });
            y += 40;

            // 4. 排除清單設定
            Label lblEx = new Label { Text = "排除寫入比對：", Location = new Point(20, y), AutoSize = true };
            cmbExTable.SelectedIndexChanged += (s, e) => {
                config.ExcludeTable = cmbExTable.SelectedItem?.ToString() ?? "";
                var exCols = App_Database.GetColumns(txtDb.Text, config.ExcludeTable);
                exCols.Insert(0, "");
                cmbExCol.Items.Clear(); cmbExCol.Items.AddRange(exCols.ToArray());
                if (exCols.Contains(config.ExcludeColumn)) cmbExCol.SelectedIndex = exCols.IndexOf(config.ExcludeColumn);
            };
            cmbExCol.SelectedIndexChanged += (s, e) => { config.ExcludeColumn = cmbExCol.SelectedItem?.ToString() ?? ""; };
            page.Controls.AddRange(new Control[] { lblEx, cmbExTable, cmbExCol });

            // 載入初始畫面資料
            if (!string.IsNullOrEmpty(txtDb.Text)) LoadTables();

            page.Controls.Add(mappingPanel);
        }
    }
}
