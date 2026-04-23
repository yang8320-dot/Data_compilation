/*
 * 檔案功能：應用程式進入點，負責啟動主視窗與全域例外處理 (支援高解析度 DPI 防模糊)。
 * 對應選單名稱：無 (系統核心)
 * 對應資料庫名稱：無
 * 對應資料表名稱：無
 */
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace FormCrawlerApp
{
    internal static class Program
    {
        // 匯入 Windows 底層 API 以啟用 DPI 感知，防止畫面模糊
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main()
        {
            // 在 Windows Vista 以上版本啟用 DPI 感知
            if (Environment.OSVersion.Version.Major >= 6)
            {
                SetProcessDPIAware();
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}
