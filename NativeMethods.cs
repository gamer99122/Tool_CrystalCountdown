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

                StringBuilder className = new StringBuilder(256);
                int result = GetClassName(hWnd, className, className.Capacity);

                return result > 0 ? className.ToString() : string.Empty;
            }
            catch (DllNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 DLL 未找到: {ex.Message}");
                return string.Empty;
            }
            catch (EntryPointNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 API 進入點未找到: {ex.Message}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 取得前景視窗類別名稱時發生意外異常: {ex.Message}");
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

                // 取得處理程序 ID
                GetWindowThreadProcessId(hWnd, out uint processId);
                if (processId == 0)
                    return string.Empty;

                // 開啟處理程序
                processHandle = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, false, processId);
                if (processHandle == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine($"[WARN] 無法開啟處理程序 (PID: {processId})，可能權限不足");
                    return string.Empty;
                }

                // 取得執行檔路徑
                StringBuilder processPath = new StringBuilder(1024);
                uint pathLength = GetModuleFileNameEx(processHandle, IntPtr.Zero, processPath, (uint)processPath.Capacity);

                return pathLength > 0 ? processPath.ToString() : string.Empty;
            }
            catch (DllNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 DLL 未找到: {ex.Message}");
                return string.Empty;
            }
            catch (EntryPointNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 API 進入點未找到: {ex.Message}");
                return string.Empty;
            }
            catch (UnauthorizedAccessException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 訪問被拒絕（權限不足）: {ex.Message}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 取得前景視窗執行檔路徑時發生意外異常: {ex.Message}");
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
                        System.Diagnostics.Debug.WriteLine($"[ERROR] 關閉處理程序句柄時發生異常: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 DLL 未找到 (檢查桌面): {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 檢查前景視窗時發生異常: {ex.Message}");
                return false; // 發生錯誤時，預設不顯示
            }
        }

        #endregion
    }
}
