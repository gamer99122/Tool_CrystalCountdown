using System;
using System.Runtime.InteropServices;
using System.Text;

namespace DesktopAnnouncement
{
    /// <summary>
    /// 封裝 Win32 API 調用，用於監控前景視窗狀態
    /// </summary>
    internal static class NativeMethods
    {
        #region Win32 API 宣告

        /// <summary>
        /// 取得當前前景視窗的 Handle
        /// </summary>
        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        /// <summary>
        /// 取得視窗的類別名稱
        /// </summary>
        /// <param name="hWnd">視窗 Handle</param>
        /// <param name="lpClassName">接收類別名稱的緩衝區</param>
        /// <param name="nMaxCount">緩衝區大小</param>
        /// <returns>成功時返回字串長度，失敗返回 0</returns>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>
        /// 取得視窗所屬的執行緒 ID 與處理程序 ID
        /// </summary>
        /// <param name="hWnd">視窗 Handle</param>
        /// <param name="lpdwProcessId">接收處理程序 ID 的變數指標</param>
        /// <returns>執行緒 ID</returns>
        [DllImport("user32.dll")]
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary>
        /// 開啟處理程序並取得 Handle
        /// </summary>
        /// <param name="dwDesiredAccess">存取權限</param>
        /// <param name="bInheritHandle">是否繼承 Handle</param>
        /// <param name="dwProcessId">處理程序 ID</param>
        /// <returns>處理程序 Handle</returns>
        [DllImport("kernel32.dll")]
        internal static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        /// <summary>
        /// 取得處理程序的執行檔路徑
        /// </summary>
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        internal static extern uint GetModuleFileNameEx(IntPtr hProcess, IntPtr hModule, StringBuilder lpFilename, uint nSize);

        /// <summary>
        /// 關閉 Handle 以釋放資源
        /// </summary>
        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr hObject);

        /// <summary>
        /// 設定視窗的位置和層級
        /// </summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        /// <summary>
        /// 檢查視窗 Handle 是否有效
        /// </summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindow(IntPtr hWnd);

        /// <summary>
        /// 查找指定類別名稱的視窗
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        /// <summary>
        /// 設置視窗的父視窗
        /// </summary>
        [DllImport("user32.dll")]
        internal static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        /// <summary>
        /// 枚舉所有視窗
        /// </summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        /// <summary>
        /// 枚舉視窗的回調函數委託
        /// </summary>
        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        #region GetWindowLong / SetWindowLong (32/64 bit compatible)

        // 32-bit signature
        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern IntPtr GetWindowLong32(IntPtr hWnd, int nIndex);

        // 64-bit signature
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        /// <summary>
        /// 設定視窗樣式 (Compatible with 32/64 bit)
        /// </summary>
        public static IntPtr GetWindowLong(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size == 8)
                return GetWindowLongPtr64(hWnd, nIndex);
            else
                return GetWindowLong32(hWnd, nIndex);
        }

        // 32-bit signature
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        // 64-bit signature
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        /// <summary>
        /// 修改視窗樣式 (Compatible with 32/64 bit)
        /// </summary>
        internal static IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8)
                return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            else
                return SetWindowLong32(hWnd, nIndex, dwNewLong);
        }

        #endregion

        /// <summary>
        /// 取得視窗文字標題
        /// </summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        #endregion

        #region 常數定義

        /// <summary>
        /// 處理程序存取權限：查詢資訊
        /// </summary>
        internal const uint PROCESS_QUERY_INFORMATION = 0x0400;

        /// <summary>
        /// 處理程序存取權限：讀取虛擬記憶體
        /// </summary>
        internal const uint PROCESS_VM_READ = 0x0010;

        /// <summary>
        /// SetWindowPos：將視窗置於所有非頂層視窗的下方
        /// </summary>
        internal static readonly IntPtr HWND_BOTTOM = new IntPtr(1);

        /// <summary>
        /// SetWindowPos：將視窗置於最上層(總是在其他視窗之上)
        /// </summary>
        internal static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        /// <summary>
        /// SetWindowPos 旗標：保持當前大小
        /// </summary>
        internal const uint SWP_NOSIZE = 0x0001;

        /// <summary>
        /// SetWindowPos 旗標：保持當前位置
        /// </summary>
        internal const uint SWP_NOMOVE = 0x0002;

        /// <summary>
        /// SetWindowPos 旗標：不啟動視窗
        /// </summary>
        internal const uint SWP_NOACTIVATE = 0x0010;

        /// <summary>
        /// 視窗樣式：WS_CHILD
        /// </summary>
        internal const int WS_CHILD = 0x40000000;

        /// <summary>
        /// 視窗樣式：WS_VISIBLE
        /// </summary>
        internal const int WS_VISIBLE = 0x10000000;

        /// <summary>
        /// GetWindowLong 索引：視窗樣式
        /// </summary>
        internal const int GWL_STYLE = -16;

        /// <summary>
        /// GetWindowLong 索引:擴展視窗樣式
        /// </summary>
        internal const int GWL_EXSTYLE = -20;

        /// <summary>
        /// 擴展視窗樣式:視窗不會被激活(不搶奪焦點)
        /// </summary>
        internal const int WS_EX_NOACTIVATE = 0x08000000;

        #endregion

        #region 輔助方法

        /// <summary>
        /// 取得前景視窗的類別名稱
        /// </summary>
        /// <returns>類別名稱，失敗時返回空字串</returns>
        internal static string GetForegroundWindowClassName()
        {
            try
            {
                IntPtr hWnd = GetForegroundWindow();
                if (hWnd == IntPtr.Zero)
                    return string.Empty;

                // 檢查 Handle 是否有效（防止競態條件：視窗可能在取得 Handle 後被銷毀）
                if (!IsWindow(hWnd))
                {
                    Logger.Warn("[WARN] 前景視窗 Handle 無效（視窗可能已被銷毀）");
                    return string.Empty;
                }

                // 使用足夠大的緩衝區（Win32 API 類別名稱通常不超過 256 字元，但使用 1024 以防不測）
                const int initialSize = 256;
                StringBuilder className = new StringBuilder(initialSize);
                int result = GetClassName(hWnd, className, className.Capacity);

                // 檢查是否被截斷（回傳值 = 緩衝區容量 - 1 表示可能被截斷）
                if (result > 0 && result == className.Capacity - 1)
                {
                    Logger.Warn($"[WARN] 類別名稱可能被截斷，重試使用更大的緩衝區");
                    className = new StringBuilder(1024);
                    result = GetClassName(hWnd, className, className.Capacity);
                }

                return result > 0 ? className.ToString() : string.Empty;
            }
            catch (DllNotFoundException ex)
            {
                Logger.Error($"[WARN] Win32 DLL 未找到: {ex.Message}", ex);
                return string.Empty;
            }
            catch (EntryPointNotFoundException ex)
            {
                Logger.Error($"[WARN] Win32 API 進入點未找到: {ex.Message}", ex);
                return string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 取得前景視窗類別名稱時發生意外異常", ex);
                return string.Empty;
            }
        }

        /// <summary>
        /// 取得前景視窗的執行檔路徑
        /// </summary>
        /// <returns>執行檔路徑，失敗時返回空字串</returns>
        internal static string GetForegroundWindowProcessPath()
        {
            IntPtr processHandle = IntPtr.Zero;
            try
            {
                IntPtr hWnd = GetForegroundWindow();
                if (hWnd == IntPtr.Zero)
                    return string.Empty;

                // 檢查 Handle 是否有效（防止競態條件：視窗可能在取得 Handle 後被銷毀）
                if (!IsWindow(hWnd))
                {
                    Logger.Warn("[WARN] 前景視窗 Handle 無效（視窗可能已被銷毀）");
                    return string.Empty;
                }

                // 取得處理程序 ID
                GetWindowThreadProcessId(hWnd, out uint processId);
                if (processId == 0)
                    return string.Empty;

                // 開啟處理程序
                processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, processId);
                if (processHandle == IntPtr.Zero)
                {
                    // 權限不足是正常的，只記錄一般日誌即可
                    // Logger.Warn($"[WARN] 無法開啟處理程序 (PID: {processId})，可能權限不足");
                    return string.Empty;
                }

                // 取得執行檔路徑（使用足夠的緩衝區並檢查截斷）
                StringBuilder processPath = new StringBuilder(1024);
                uint pathLength = GetModuleFileNameEx(processHandle, IntPtr.Zero, processPath, (uint)processPath.Capacity);

                // 檢查是否被截斷（回傳值 >= 緩衝區容量表示可能被截斷）
                if (pathLength > 0 && pathLength >= processPath.Capacity)
                {
                    Logger.Warn($"[WARN] 執行檔路徑可能被截斷，重試使用更大的緩衝區");
                    processPath = new StringBuilder(4096);
                    pathLength = GetModuleFileNameEx(processHandle, IntPtr.Zero, processPath, (uint)processPath.Capacity);
                }

                return pathLength > 0 ? processPath.ToString() : string.Empty;
            }
            catch (DllNotFoundException ex)
            {
                Logger.Error($"[WARN] Win32 DLL 未找到: {ex.Message}", ex);
                return string.Empty;
            }
            catch (EntryPointNotFoundException ex)
            {
                Logger.Error($"[WARN] Win32 API 進入點未找到: {ex.Message}", ex);
                return string.Empty;
            }
            catch (UnauthorizedAccessException ex)
            {
                 // 權限不足是正常的
                return string.Empty;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 取得前景視窗執行檔路徑時發生意外異常", ex);
                return string.Empty;
            }
            finally
            {
                // 確保資源正確釋放
                if (processHandle != IntPtr.Zero)
                {
                    try
                    {
                        CloseHandle(processHandle);
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"[ERROR] 關閉處理程序句柄時發生異常", ex);
                    }
                }
            }
        }

        /// <summary>
        /// 檢查當前前景視窗是否為桌面
        /// </summary>
        /// <returns>true 表示為桌面，false 表示為其他應用程式</returns>
        internal static bool IsForegroundWindowDesktop()
        {
            try
            {
                // 取得類別名稱
                string className = GetForegroundWindowClassName();

                // 檢查已知的桌面類別名稱
                if (className.Equals("Progman", StringComparison.OrdinalIgnoreCase) ||
                    className.Equals("WorkerW", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                // 取得執行檔路徑並檢查是否為 Explorer.exe
                string processPath = GetForegroundWindowProcessPath();
                if (!string.IsNullOrEmpty(processPath))
                {
                    string fileName = System.IO.Path.GetFileName(processPath);
                    if (fileName.Equals("explorer.exe", StringComparison.OrdinalIgnoreCase))
                    {
                        // 進一步確認是否為桌面視窗（而非檔案總管視窗）
                        return className.Equals("Progman", StringComparison.OrdinalIgnoreCase) ||
                               className.Equals("WorkerW", StringComparison.OrdinalIgnoreCase) ||
                               className.Equals("Shell_TrayWnd", StringComparison.OrdinalIgnoreCase);
                    }
                }

                return false;
            }
            catch (DllNotFoundException ex)
            {
                Logger.Error($"[WARN] Win32 DLL 未找到 (檢查桌面): {ex.Message}", ex);
                return false;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 檢查前景視窗時發生異常", ex);
                return false; // 發生錯誤時，預設不顯示
            }
        }

        /// <summary>
        /// 查找桌面視窗（Progman 或 WorkerW）
        /// </summary>
        /// <returns>桌面視窗 Handle，失敗返回 IntPtr.Zero</returns>
        internal static IntPtr FindDesktopWindow()
        {
            try
            {
                // 首先嘗試查找 Progman（舊版 Windows）
                IntPtr desktopWindow = FindWindow("Progman", null);
                if (desktopWindow != IntPtr.Zero)
                {
                    return desktopWindow;
                }

                // 如果找不到 Progman，查找 WorkerW（Windows 7+）
                desktopWindow = FindWorkerWindow();
                if (desktopWindow != IntPtr.Zero)
                {
                    return desktopWindow;
                }

                Logger.Warn("[WARN] 無法找到桌面窗口");
                return IntPtr.Zero;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 查找桌面窗口時發生異常", ex);
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 查找 WorkerW 窗口（Windows 7+）
        /// </summary>
        private static IntPtr FindWorkerWindow()
        {
            try
            {
                IntPtr progmanHandle = FindWindow("Progman", null);
                if (progmanHandle == IntPtr.Zero)
                    return IntPtr.Zero;

                // WorkerW 是 Progman 的子窗口，通過枚舉查找
                IntPtr workerWindow = IntPtr.Zero;

                EnumWindowsProc enumProc = (hWnd, lParam) =>
                {
                    StringBuilder className = new StringBuilder(256);
                    GetClassName(hWnd, className, className.Capacity);

                    if (className.ToString() == "WorkerW")
                    {
                        // 檢查這個 WorkerW 是否是我們需要的（應該在 Progman 下）
                        workerWindow = hWnd;
                        return false; // 停止枚舉
                    }

                    return true; // 繼續枚舉
                };

                EnumWindows(enumProc, IntPtr.Zero);

                return workerWindow;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 查找 WorkerW 窗口時發生異常", ex);
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 將視窗與桌面整合（成為桌面的子視窗）
        /// </summary>
        /// <param name="hWnd">要與桌面整合的視窗 Handle</param>
        /// <returns>成功返回 true，失敗返回 false</returns>
        internal static bool IntegrateWithDesktop(IntPtr hWnd)
        {
            try
            {
                if (hWnd == IntPtr.Zero)
                {
                    Logger.Error("[ERROR] 無效的視窗 Handle");
                    return false;
                }

                // 查找桌面視窗
                IntPtr desktopWindow = FindDesktopWindow();
                if (desktopWindow == IntPtr.Zero)
                {
                    Logger.Warn("[WARN] 無法找到桌面窗口，跳過桌面整合");
                    return false;
                }

                // 將當前視窗設置為桌面的子視窗
                IntPtr result = SetParent(hWnd, desktopWindow);
                if (result == IntPtr.Zero)
                {
                    Logger.Error("[ERROR] SetParent 失敗");
                    return false;
                }

                Logger.Info("[INFO] 視窗已與桌面整合");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error($"[ERROR] 視窗與桌面整合時發生異常", ex);
                return false;
            }
        }

        #endregion
    }
}
