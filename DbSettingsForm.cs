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
        private string[] scrapeHeaders = { "表單單號", "分類", "表單主題", "狀態", "申請者", "承辦人", "目前處理者", "申請時間", "修改時間", "到期時間", "網址" };
        
        // 內部用來記錄 4 個 v 欄位和自訂文字的假名
        private string[] vFields = { "v_1", "v_2", "v_3", "v_4" };
        private string customTextFieldName = "CustomText";

        private Dictionary<CategoryDbSetting, TextBox> excludeTextBoxes = new Dictionary<CategoryDbSetting, TextBox>();
        private Dictionary<CategoryDbSetting, TextBox> customTextBoxes = new Dictionary<CategoryDbSetting, TextBox>();

        public DbSettingsForm(App_DbSettings settings)
        {
            dbSettings = settings;
            InitializeUI();
        }

        private void InitializeUI()
        {
            this.Text = "資料庫寫入設定";
            this.Size = new Size(970, 820); 
            this.StartPosition = FormStartPosition.CenterParent;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Font = new Font("Microsoft JhengHei", 10F);

            tabControl = new TabControl { Dock = DockStyle.Top, Height = 700 };

            foreach (var cat in dbSettings.Categories)
            {
                TabPage page = new TabPage { Text = cat.CategoryName, BackColor = Color.White };
                BuildCategoryPanel(page, cat);
                tabControl.TabPages.Add(page);
            }

            Button btnSave = new Button {
                Text = "💾 儲存所有資料庫設定",
                Location = new Point(295, 720), Size = new Size(380, 45),
                BackColor = Color.LightSteelBlue, Cursor = Cursors.Hand
            };
            btnSave.Click += (s, e) => {
                // 儲存黑名單
                foreach (var kvp in excludeTextBoxes)
                {
                    kvp.Key.ExcludeFormNumbers = kvp.Value.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                                                               .Select(txt => txt.Trim())
                                                               .Where(txt => !string.IsNullOrEmpty(txt))
                                                               .ToList();
                }
                // 儲存自訂文字
                foreach (var kvp in customTextBoxes)
                {
                    kvp.Key.CustomTextValue = kvp.Value.Text;
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

            Label lblEnable = new Label { Text = "是否寫入此類別：", Location = new Point(labelX, y), AutoSize = true };
            ComboBox cmbEnable = new ComboBox { Location = new Point(controlX, y), DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            cmbEnable.Items.AddRange(new[] { "寫入", "不寫入" });
            cmbEnable.SelectedIndex = config.IsEnabled ? 0 : 1;
            cmbEnable.SelectedIndexChanged += (s, e) => { config.IsEnabled = cmbEnable.SelectedIndex == 0; };
            page.Controls.AddRange(new Control[] { lblEnable, cmbEnable });
            y += 45;

            Label lblDb = new Label { Text = "寫入主庫(SQLite)：", Location = new Point(labelX, y+5), AutoSize = true };
            TextBox txtDb = new TextBox { Text = config.DbFilePath, Location = new Point(controlX, y), Width = 380, ReadOnly = true };
            
            Button btnDbBrowse = new Button { Text = "選擇資料庫", Location = new Point(570, y - 4), Size = new Size(100, 35), Cursor = Cursors.Hand };
            Button btnDbLoad = new Button { Text = "讀取資料庫", Location = new Point(680, y - 4), Size = new Size(100, 35), Cursor = Cursors.Hand, BackColor = Color.PaleGreen };
            
            page.Controls.AddRange(new Control[] { lblDb, txtDb, btnDbBrowse, btnDbLoad });
            y += 45;

            ComboBox cmbTable = new ComboBox { Location = new Point(controlX, y), Width = 400, DropDownStyle = ComboBoxStyle.DropDownList };
            Label lblTable = new Label { Text = "寫入主資料表：", Location = new Point(labelX, y+3), AutoSize = true };
            page.Controls.AddRange(new Control[] { lblTable, cmbTable });
            y += 50;

            List<ComboBox> colMappingCmbs = new List<ComboBox>();
            List<ComboBox> vMappingCmbs = new List<ComboBox>();
            ComboBox cbCustom = new ComboBox(); // 自訂文字用的下拉選單
            
            // 為了塞入 2 行 v 以及 1 行自訂文字，高度增加到 500
            Panel mappingPanel = new Panel { Location = new Point(labelX, y), Size = new Size(570, 500), BorderStyle = BorderStyle.FixedSingle };
            int my = 15;
            
            // 11 個標準欄位
            foreach (var field in scrapeHeaders)
            {
                Label lblF = new Label { Text = $"[{field}] 寫入：", Location = new Point(15, my+4), AutoSize = true, ForeColor = Color.DarkBlue };
                // 【修改點】再往右移 10px (變成 X=180)，寬度維持 260
                ComboBox cmbF = new ComboBox { Location = new Point(180, my), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
                
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == field);
                if (existMap != null) cmbF.Tag = existMap.DbColumn; 

                colMappingCmbs.Add(cmbF);
                mappingPanel.Controls.AddRange(new Control[] { lblF, cmbF });
                my += 33; 
            }

            // 【新增】第 1 行：v 欄位 1, 2
            Label lblV1 = new Label { Text = "[v] 寫入(1,2)：", Location = new Point(15, my + 4), AutoSize = true, ForeColor = Color.DarkMagenta };
            mappingPanel.Controls.Add(lblV1);
            for (int i = 0; i < 2; i++) {
                ComboBox cmbV = new ComboBox { Location = new Point(180 + (i * 135), my), Width = 125, DropDownStyle = ComboBoxStyle.DropDownList };
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == vFields[i]);
                if (existMap != null) cmbV.Tag = existMap.DbColumn; 
                vMappingCmbs.Add(cmbV);
                mappingPanel.Controls.Add(cmbV);
            }
            my += 33;

            // 【新增】第 2 行：v 欄位 3, 4
            Label lblV2 = new Label { Text = "[v] 寫入(3,4)：", Location = new Point(15, my + 4), AutoSize = true, ForeColor = Color.DarkMagenta };
            mappingPanel.Controls.Add(lblV2);
            for (int i = 2; i < 4; i++) {
                ComboBox cmbV = new ComboBox { Location = new Point(180 + ((i - 2) * 135), my), Width = 125, DropDownStyle = ComboBoxStyle.DropDownList };
                var existMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == vFields[i]);
                if (existMap != null) cmbV.Tag = existMap.DbColumn; 
                vMappingCmbs.Add(cmbV);
                mappingPanel.Controls.Add(cmbV);
            }
            my += 33;

            // 【新增】第 3 行：自行 Key-in 文字與對應欄位
            Label lblCustom = new Label { Text = "[自訂字] 寫入：", Location = new Point(15, my + 4), AutoSize = true, ForeColor = Color.DarkGreen };
            TextBox txtCustom = new TextBox { Location = new Point(125, my), Width = 100, Text = config.CustomTextValue };
            Label lblArrow = new Label { Text = "➔", Location = new Point(228, my + 4), AutoSize = true, ForeColor = Color.DarkGray };
            cbCustom = new ComboBox { Location = new Point(250, my), Width = 190, DropDownStyle = ComboBoxStyle.DropDownList };
            
            var existCustomMap = config.Mappings.FirstOrDefault(m => m.ScrapedField == customTextFieldName);
            if (existCustomMap != null) cbCustom.Tag = existCustomMap.DbColumn;
            
            customTextBoxes[config] = txtCustom;
            mappingPanel.Controls.AddRange(new Control[] { lblCustom, txtCustom, lblArrow, cbCustom });

            page.Controls.Add(mappingPanel);

            Label lblExclude = new Label { Text = "排除寫入清單\n(每行輸入一筆表單單號)：", Location = new Point(600, 145), AutoSize = true, ForeColor = Color.Brown };
            
            // 配合左方面板加長，黑名單文字框也同步加長
            TextBox txtExclude = new TextBox {
                Location = new Point(600, 190),
                Size = new Size(300, 470), 
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = false
            };
            
            if (config.ExcludeFormNumbers != null && config.ExcludeFormNumbers.Count > 0)
            {
                txtExclude.Text = string.Join(Environment.NewLine, config.ExcludeFormNumbers);
            }
            excludeTextBoxes[config] = txtExclude;
            
            page.Controls.Add(lblExclude);
            page.Controls.Add(txtExclude);

            Action<string> UpdateColumnLists = (tableName) => {
                try {
                    var cols = App_Database.GetColumns(txtDb.Text, tableName);
                    cols.Insert(0, ""); 
                    
                    // 1. 更新 11 個標準欄位
                    for (int i = 0; i < scrapeHeaders.Length; i++) {
                        var cb = colMappingCmbs[i];
                        cb.Items.Clear(); cb.Items.AddRange(cols.ToArray());
                        string targetStr = cb.Tag?.ToString() ?? "";
                        cb.SelectedIndex = cols.Contains(targetStr) ? cols.IndexOf(targetStr) : 0;
                        
                        int indexCopy = i;
                        cb.SelectedIndexChanged -= Cb_SelectedIndexChanged;
                        cb.SelectedIndexChanged += Cb_SelectedIndexChanged;
                        void Cb_SelectedIndexChanged(object sender, EventArgs e) {
                            var map = config.Mappings.FirstOrDefault(m => m.ScrapedField == scrapeHeaders[indexCopy]);
                            if (map == null) { map = new FieldMapping { ScrapedField = scrapeHeaders[indexCopy] }; config.Mappings.Add(map); }
                            map.DbColumn = cb.SelectedItem?.ToString() ?? "";
                            cb.Tag = map.DbColumn;
                        }
                    }

                    // 2. 更新 4 個 [v] 寫入欄位
                    for (int i = 0; i < vFields.Length; i++) {
                        var cbV = vMappingCmbs[i];
                        cbV.Items.Clear(); cbV.Items.AddRange(cols.ToArray());
                        string targetStr = cbV.Tag?.ToString() ?? "";
                        cbV.SelectedIndex = cols.Contains(targetStr) ? cols.IndexOf(targetStr) : 0;
                        
                        int indexCopy = i;
                        cbV.SelectedIndexChanged -= CbV_SelectedIndexChanged;
                        cbV.SelectedIndexChanged += CbV_SelectedIndexChanged;
                        void CbV_SelectedIndexChanged(object sender, EventArgs e) {
                            var map = config.Mappings.FirstOrDefault(m => m.ScrapedField == vFields[indexCopy]);
                            if (map == null) { map = new FieldMapping { ScrapedField = vFields[indexCopy] }; config.Mappings.Add(map); }
                            map.DbColumn = cbV.SelectedItem?.ToString() ?? "";
                            cbV.Tag = map.DbColumn;
                        }
                    }

                    // 3. 更新 [自訂文字] 下拉選單
                    cbCustom.Items.Clear(); cbCustom.Items.AddRange(cols.ToArray());
                    string targetCustomStr = cbCustom.Tag?.ToString() ?? "";
                    cbCustom.SelectedIndex = cols.Contains(targetCustomStr) ? cols.IndexOf(targetCustomStr) : 0;
                    
                    cbCustom.SelectedIndexChanged -= CbCustom_SelectedIndexChanged;
                    cbCustom.SelectedIndexChanged += CbCustom_SelectedIndexChanged;
                    void CbCustom_SelectedIndexChanged(object sender, EventArgs e) {
                        var map = config.Mappings.FirstOrDefault(m => m.ScrapedField == customTextFieldName);
                        if (map == null) { map = new FieldMapping { ScrapedField = customTextFieldName }; config.Mappings.Add(map); }
                        map.DbColumn = cbCustom.SelectedItem?.ToString() ?? "";
                        cbCustom.Tag = map.DbColumn;
                    }

                } catch (Exception ex) {
                    MessageBox.Show("讀取欄位失敗：\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    if (ofd.ShowDialog() == DialogResult.OK) { txtDb.Text = config.DbFilePath = ofd.FileName; }
                }
            };

            btnDbLoad.Click += (s, e) => { 
                try {
                    LoadMainTables(); 
                    if (cmbTable.Items.Count <= 1) {
                        MessageBox.Show("讀取成功，但該資料庫內沒有任何資料表(Table)！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    } else {
                        MessageBox.Show($"資料庫讀取成功！\n共找到 {cmbTable.Items.Count - 1} 個資料表。", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                } catch (Exception ex) {
                    MessageBox.Show("資料庫讀取失敗，詳細原因：\n\n" + ex.Message, "讀取錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            cmbTable.SelectedIndexChanged += (s, e) => {
                config.TargetTable = cmbTable.SelectedItem?.ToString() ?? "";
                if (!string.IsNullOrEmpty(config.TargetTable)) {
                    UpdateColumnLists(config.TargetTable);
                }
            };

            if (!string.IsNullOrEmpty(txtDb.Text)) {
                try { LoadMainTables(); } catch { }
            }
        }
    }
}
