/*
 * 檔案功能：應用程式進入點，負責啟動主視窗與全域例外處理。
 * 對應選單名稱：無 (系統核心)
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Windows.Forms;

namespace FormCrawlerApp
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
