using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace FormCrawlerApp
{
    public class App_DbSettings
    {
        public List<CategoryDbSetting> Categories { get; set; } = new List<CategoryDbSetting>();

        public static App_DbSettings Load()
        {
            // 強制綁定絕對路徑
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DbMappingSettings.xml");
            if (!File.Exists(path)) return DefaultSettings();
            try {
                using (FileStream fs = new FileStream(path, FileMode.Open)) {
                    XmlSerializer xs = new XmlSerializer(typeof(App_DbSettings));
                    return (App_DbSettings)xs.Deserialize(fs);
                }
            } catch { return DefaultSettings(); }
        }

        public void Save()
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DbMappingSettings.xml");
            using (FileStream fs = new FileStream(path, FileMode.Create)) {
                XmlSerializer xs = new XmlSerializer(typeof(App_DbSettings));
                xs.Serialize(fs, this);
            }
        }

        private static App_DbSettings DefaultSettings()
        {
            var s = new App_DbSettings();
            string[] cats = { "彰濱廠異常改善單", "彰濱聯絡書", "台玻內文", "彰濱廠郵件收文", "彰濱廠虛驚事件輕度傷害記錄表" };
            foreach (var c in cats) {
                s.Categories.Add(new CategoryDbSetting { CategoryName = c });
            }
            return s;
        }
    }

    public class CategoryDbSetting
    {
        public string CategoryName { get; set; } = "";
        public bool IsEnabled { get; set; } = false;
        
        public string DbFilePath { get; set; } = "";
        public string TargetTable { get; set; } = "";
        public List<FieldMapping> Mappings { get; set; } = new List<FieldMapping>();
        
        public List<string> ExcludeFormNumbers { get; set; } = new List<string>();
    }

    public class FieldMapping
    {
        public string ScrapedField { get; set; } = "";
        public string DbColumn { get; set; } = "";
    }
}
