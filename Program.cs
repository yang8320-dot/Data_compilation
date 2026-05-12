/*
 * 檔案功能：應用程式進入點，支援高 DPI 防模糊，並防止程式重複開啟 (單一執行個體)。
 * 對應選單名稱：無 (系統核心)
 */
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;

namespace FormCrawlerApp
{
    internal static class Program
    {
        // --- Windows API 宣告 ---
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SW_RESTORE = 9; // 還原視窗的代碼
        private static Mutex mutex = null;

        [STAThread]
        static void Main()
        {
            // 設定一個獨一無二的應用程式識別碼
            const string appName = "FormCrawlerApp_Unique_Instance";
            bool createdNew;

            // 嘗試取得 Mutex，若 createdNew 為 false，代表程式已經在執行了
            mutex = new Mutex(true, appName, out createdNew);

            if (!createdNew)
            {
                // 尋找名為 "經手表單自動化工具" 的主視窗 (名稱必須與 MainForm.Text 一致)
                IntPtr hWnd = FindWindow(null, "經手表單自動化工具");
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, SW_RESTORE); // 如果被最小化了，將它還原
                    SetForegroundWindow(hWnd);    // 將視窗拉到最上層
                }
                return; // 直接結束這次的開啟，達成防重複啟動
            }

            // 啟用 DPI 感知防模糊 (確保 MsgBox 也是清晰的)
            if (Environment.OSVersion.Version.Major >= 6)
            {
                SetProcessDPIAware();
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 【新增】執行 Windows 使用者權限驗證
            if (!App_LicenseManager.VerifyLicense())
            {
                MessageBox.Show("驗證失敗，非授權使用者！\n\n您沒有權限執行此程式，系統即將關閉。", "系統提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // 直接結束程式
            }

            Application.Run(new MainForm());
        }
    }
}
