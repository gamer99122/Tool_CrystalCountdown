using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;

namespace DesktopAnnouncement
{
    /// <summary>
    /// App.xaml 的互動邏輯
    /// </summary>
    public partial class App : Application
    {
        /// <summary>
        /// 互斥鎖，用於確保單一實例
        /// </summary>
        private static Mutex? _mutex = null;

        /// <summary>
        /// 應用程式啟動事件
        /// </summary>
        protected override void OnStartup(StartupEventArgs e)
        {
            // 建立全域互斥鎖
            const string mutexName = "Global\\DesktopAnnouncement_SingleInstance_Mutex";
            bool createdNew;

            _mutex = new Mutex(true, mutexName, out createdNew);

            if (!createdNew)
            {
                // 已有實例在執行，關閉舊實例並啟動新的
                TerminateExistingInstance();

                // 等待舊實例完全關閉
                Thread.Sleep(500);

                // 重新建立互斥鎖
                _mutex?.Dispose();
                _mutex = new Mutex(true, mutexName, out createdNew);
            }

            base.OnStartup(e);

            // 設定未處理的例外處理器（增強穩定性）
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            DispatcherUnhandledException += OnDispatcherUnhandledException;
        }

        /// <summary>
        /// 處理非 UI 執行緒的未處理例外
        /// </summary>
        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[CRITICAL] 非 UI 執行緒未處理的例外：{ex.GetType().Name}");
                System.Diagnostics.Debug.WriteLine($"[CRITICAL] 訊息：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[CRITICAL] 堆疊追蹤：{ex.StackTrace}");

                // 如果是致命錯誤，記錄完整堆疊
                if (e.IsTerminating)
                {
                    System.Diagnostics.Debug.WriteLine("[CRITICAL] 應用程式即將終止");
                }
            }
        }

        /// <summary>
        /// 處理 UI 執行緒的未處理例外
        /// </summary>
        private void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[ERROR] UI 執行緒未處理的例外：{e.Exception.GetType().Name}");
            System.Diagnostics.Debug.WriteLine($"[ERROR] 訊息：{e.Exception.Message}");
            System.Diagnostics.Debug.WriteLine($"[ERROR] 堆疊追蹤：{e.Exception.StackTrace}");

            e.Handled = true; // 防止應用程式崩潰
        }

        /// <summary>
        /// 終止已存在的實例
        /// </summary>
        private void TerminateExistingInstance()
        {
            Process? currentProcess = null;
            try
            {
                // 取得當前處理程序資訊
                currentProcess = Process.GetCurrentProcess();
                var currentProcessName = currentProcess.ProcessName;
                var currentProcessId = currentProcess.Id;

                // 尋找所有同名的處理程序，使用 using 確保 Process 物件被正確釋放
                Process[] allProcesses = Process.GetProcessesByName(currentProcessName);
                try
                {
                    // 終止所有舊實例（先優雅關閉，必要時才強制終止）
                    foreach (var process in allProcesses)
                    {
                        // 跳過當前處理程序
                        if (process.Id == currentProcessId)
                        {
                            process.Dispose();
                            continue;
                        }

                        try
                        {
                            Debug.WriteLine($"[DesktopAnnouncement] 嘗試關閉舊實例 (PID: {process.Id})");

                            // 先嘗試優雅關閉（關閉主視窗）
                            bool closedGracefully = process.CloseMainWindow();

                            if (closedGracefully)
                            {
                                // 等待進程優雅關閉
                                if (process.WaitForExit(2000))
                                {
                                    Debug.WriteLine($"[DesktopAnnouncement] 舊實例 (PID: {process.Id}) 已優雅關閉");
                                    process.Dispose();
                                    continue;
                                }
                                else
                                {
                                    Debug.WriteLine($"[DesktopAnnouncement] 舊實例 (PID: {process.Id}) 未在時間限制內關閉，進行強制終止");
                                }
                            }

                            // 如果優雅關閉失敗或超時，進行強制終止
                            process.Kill();
                            process.WaitForExit(1000);
                            Debug.WriteLine($"[DesktopAnnouncement] 舊實例 (PID: {process.Id}) 已強制終止");
                            process.Dispose();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[DesktopAnnouncement] 無法終止處理程序 {process.Id}: {ex.Message}");
                            process.Dispose();
                        }
                    }
                }
                finally
                {
                    // 確保所有 Process 物件都被釋放
                    foreach (var process in allProcesses)
                    {
                        process?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[DesktopAnnouncement] 終止舊實例時發生錯誤: {ex.Message}");
            }
            finally
            {
                // 確保當前進程物件被釋放
                currentProcess?.Dispose();
            }
        }

        /// <summary>
        /// 應用程式退出事件
        /// </summary>
        protected override void OnExit(ExitEventArgs e)
        {
            try
            {
                // 取消訂閱全域異常處理事件，防止記憶體洩漏
                AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
                DispatcherUnhandledException -= OnDispatcherUnhandledException;

                // 釋放互斥鎖
                _mutex?.ReleaseMutex();
                _mutex?.Dispose();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 應用程式退出時發生異常：{ex.Message}");
            }
            finally
            {
                base.OnExit(e);
            }
        }
    }
}
