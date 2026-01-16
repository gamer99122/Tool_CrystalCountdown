using System;
using System.IO;
using System.Text;

namespace DesktopAnnouncement
{
    /// <summary>
    /// 提供簡易的檔案日誌功能
    /// </summary>
    public static class Logger
    {
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");
        private static readonly object LockObj = new object();

        /// <summary>
        /// 記錄一般訊息
        /// </summary>
        public static void Info(string message)
        {
            WriteLog("INFO", message);
        }

        /// <summary>
        /// 記錄警告訊息
        /// </summary>
        public static void Warn(string message)
        {
            WriteLog("WARN", message);
        }

        /// <summary>
        /// 記錄錯誤訊息
        /// </summary>
        public static void Error(string message)
        {
            WriteLog("ERROR", message);
        }

        /// <summary>
        /// 記錄錯誤訊息與例外
        /// </summary>
        public static void Error(string message, Exception ex)
        {
            WriteLog("ERROR", $"{message}\nException: {ex.GetType().Name}\nMessage: {ex.Message}\nStackTrace: {ex.StackTrace}");
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                lock (LockObj)
                {
                    string logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
                    File.AppendAllText(LogPath, logLine, Encoding.UTF8);
                }
            }
            catch (Exception)
            {
                // 日誌寫入失敗不應導致程式崩潰，忽略之或嘗試寫入 Debug
                System.Diagnostics.Debug.WriteLine($"[Logger Failure] {message}");
            }
        }
    }
}
