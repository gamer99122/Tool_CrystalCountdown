using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;

namespace DesktopAnnouncement
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        #region 私有欄位

        /// <summary>
        /// 設定檔路徑
        /// </summary>
        private readonly string _configFilePath;

        /// <summary>
        /// 位置設定檔路徑
        /// </summary>
        private readonly string _positionFilePath;

        /// <summary>
        /// 目標日期（從 config.txt 讀取）
        /// </summary>
        private DateTime _targetDate;

        /// <summary>
        /// 公告標題文字
        /// </summary>
        private string _announcementTitle = string.Empty;

        /// <summary>
        /// 定時器，用於監控日期變化（每天凌晨 00:00 更新）
        /// </summary>
        private readonly DispatcherTimer _dateCheckTimer;

        /// <summary>
        /// 定時器，用於持續監控視窗層級（確保視窗保持在最底層）
        /// </summary>
        private readonly DispatcherTimer _windowLevelCheckTimer;

        /// <summary>
        /// 視窗位置變化防抖定時器（避免頻繁寫入磁盤）
        /// </summary>
        private DispatcherTimer _savePositionDebounceTimer;

        /// <summary>
        /// 標記視窗位置是否已改變
        /// </summary>
        private bool _hasPositionChanged = false;

        #endregion

        #region 建構函式

        public MainWindow()
        {
            InitializeComponent();

            // 設定視窗初始狀態
            this.Opacity = 1.0; // 總是顯示
            this.Loaded += MainWindow_Loaded;

            // 設定設定檔路徑（與執行檔同目錄）
            _configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.txt");
            _positionFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "position.txt");

            // 初始化日期檢查定時器（將在第一次載入時設定觸發時間）
            _dateCheckTimer = new DispatcherTimer();
            _dateCheckTimer.Tick += DateCheckTimer_Tick;

            // 初始化視窗層級監控定時器（每 2000ms 檢查一次，降低系統開銷）
            _windowLevelCheckTimer = new DispatcherTimer();
            _windowLevelCheckTimer.Interval = TimeSpan.FromMilliseconds(2000);
            _windowLevelCheckTimer.Tick += WindowLevelCheckTimer_Tick;

            // 初始化視窗位置保存防抖定時器（避免頻繁I/O）
            _savePositionDebounceTimer = new DispatcherTimer();
            _savePositionDebounceTimer.Interval = TimeSpan.FromMilliseconds(1000);  // 位置改變後延遲1秒再保存
            _savePositionDebounceTimer.Tick += SavePositionDebounceTimer_Tick;

            // 載入設定檔
            LoadConfiguration();
        }

        #endregion

        #region 事件處理

        /// <summary>
        /// 視窗載入完成事件
        /// </summary>
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // 載入儲存的位置，如果沒有則使用右下角
            if (!LoadWindowPosition())
            {
                SetWindowPositionToBottomRight();
            }

            // 監聽位置變化事件（用於儲存位置）
            this.LocationChanged += MainWindow_LocationChanged;

            // 設定視窗層級為底層（在所有應用程式視窗下方）
            SetWindowToBottom();

            // 啟動視窗層級監控定時器
            _windowLevelCheckTimer.Start();

            // 啟動日期檢查定時器
            ScheduleNextMidnightUpdate();
        }

        /// <summary>
        /// 日期檢查定時器觸發事件（每天凌晨 00:00）
        /// </summary>
        private void DateCheckTimer_Tick(object? sender, EventArgs e)
        {
            // 更新顯示
            UpdateAnnouncementDisplay();

            // 設定下一次觸發時間（下一個凌晨 00:00）
            ScheduleNextMidnightUpdate();
        }

        /// <summary>
        /// 視窗層級監控定時器觸發事件（每 500ms 檢查一次，確保視窗保持在最底層）
        /// </summary>
        private void WindowLevelCheckTimer_Tick(object? sender, EventArgs e)
        {
            // 確保視窗保持在最底層
            SetWindowToBottom();
        }

        /// <summary>
        /// 關閉按鈕點擊事件
        /// </summary>
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            // 關閉應用程式
            Application.Current.Shutdown();
        }

        /// <summary>
        /// 主邊框滑鼠左鍵按下事件（用於拖拽視窗）
        /// </summary>
        private void MainBorder_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                // 拖拽視窗
                this.DragMove();

                // 拖拽結束後重新設定視窗層級為底層
                SetWindowToBottom();
            }
            catch (InvalidOperationException ex)
            {
                // 滑鼠已釋放或其他無效操作
                System.Diagnostics.Debug.WriteLine($"[WARN] 拖拽視窗時發生無效操作：{ex.Message}");
            }
            catch (Exception ex)
            {
                // 其他拖拽相關錯誤
                System.Diagnostics.Debug.WriteLine($"[ERROR] 拖拽視窗時發生錯誤：{ex.GetType().Name} - {ex.Message}");
            }
        }

        /// <summary>
        /// 視窗位置變化事件（用於標記位置已改變，延遲儲存以避免頻繁I/O）
        /// </summary>
        private void MainWindow_LocationChanged(object? sender, EventArgs e)
        {
            // 標記位置已改變
            _hasPositionChanged = true;

            // 重新啟動防抖定時器（每次位置改變都重新計時）
            _savePositionDebounceTimer.Stop();
            _savePositionDebounceTimer.Start();
        }

        /// <summary>
        /// 視窗位置保存防抖定時器事件（延遲儲存位置以減少I/O）
        /// </summary>
        private void SavePositionDebounceTimer_Tick(object? sender, EventArgs e)
        {
            // 停止定時器
            _savePositionDebounceTimer.Stop();

            // 如果位置已改變，才儲存（避免不必要的磁盤寫入）
            if (_hasPositionChanged)
            {
                SaveWindowPosition();
                _hasPositionChanged = false;
            }
        }

        #endregion

        #region 核心邏輯

        /// <summary>
        /// 設定視窗位置到螢幕右下角（考慮工作列）
        /// </summary>
        private void SetWindowPositionToBottomRight()
        {
            try
            {
                // 取得工作區域（排除工作列後的可用區域）
                var workArea = SystemParameters.WorkArea;

                // 設定邊距（距離螢幕邊緣的距離）
                const double marginRight = 20;
                const double marginBottom = 20;

                // 計算右下角位置
                this.Left = workArea.Right - this.Width - marginRight;
                this.Top = workArea.Bottom - this.Height - marginBottom;
            }
            catch (InvalidOperationException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 無法取得螢幕工作區域：{ex.Message}");
                // 使用預設位置（中央）
                this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
            }
            catch (Exception ex)
            {
                // 如果定位失敗，使用預設位置（中央）
                System.Diagnostics.Debug.WriteLine($"[ERROR] 設定視窗位置失敗：{ex.GetType().Name} - {ex.Message}");
                try
                {
                    this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
                    this.Top = (SystemParameters.PrimaryScreenHeight - this.Height) / 2;
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine("[ERROR] 無法設定預設位置");
                }
            }
        }

        /// <summary>
        /// 載入儲存的視窗位置
        /// </summary>
        /// <returns>true 表示成功載入，false 表示使用預設位置</returns>
        private bool LoadWindowPosition()
        {
            try
            {
                // 檢查位置檔案是否存在
                if (!File.Exists(_positionFilePath))
                {
                    return false;
                }

                // 讀取位置資料
                string[] lines = File.ReadAllLines(_positionFilePath);
                if (lines.Length < 2)
                {
                    System.Diagnostics.Debug.WriteLine("[WARN] 位置檔案格式不正確");
                    return false;
                }

                // 解析位置
                if (double.TryParse(lines[0], out double left) &&
                    double.TryParse(lines[1], out double top))
                {
                    // 確保位置在螢幕範圍內
                    if (IsPositionValid(left, top))
                    {
                        this.Left = left;
                        this.Top = top;
                        return true;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[WARN] 載入的視窗位置無效：({left}, {top})");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[WARN] 位置值無法解析");
                }

                return false;
            }
            catch (FileNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 位置檔案未找到：{ex.Message}");
                return false;
            }
            catch (UnauthorizedAccessException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 無法訪問位置檔案（權限不足）：{ex.Message}");
                return false;
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 讀取位置檔案時發生 I/O 錯誤：{ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 載入位置失敗：{ex.GetType().Name} - {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 儲存當前視窗位置
        /// </summary>
        private void SaveWindowPosition()
        {
            try
            {
                // 只在視窗位置有效時儲存
                if (!double.IsNaN(this.Left) && !double.IsNaN(this.Top))
                {
                    // 寫入位置資料
                    string[] lines = new string[]
                    {
                        this.Left.ToString("F0"),
                        this.Top.ToString("F0")
                    };
                    File.WriteAllLines(_positionFilePath, lines);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 無法儲存位置檔案（權限不足）：{ex.Message}");
            }
            catch (DirectoryNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] 位置檔案目錄不存在：{ex.Message}");
            }
            catch (IOException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 儲存位置檔案時發生 I/O 錯誤：{ex.Message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 儲存位置失敗：{ex.GetType().Name} - {ex.Message}");
            }
        }

        /// <summary>
        /// 檢查位置是否在螢幕範圍內
        /// </summary>
        private bool IsPositionValid(double left, double top)
        {
            try
            {
                // 取得主螢幕工作區域
                var workArea = SystemParameters.WorkArea;

                // 確保視窗至少有一部分在螢幕內（考慮視窗大小）
                bool horizontalValid = left + this.Width > workArea.Left &&
                                      left < workArea.Right;
                bool verticalValid = top + this.Height > workArea.Top &&
                                    top < workArea.Bottom;

                return horizontalValid && verticalValid;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 載入 config.txt 設定檔
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                // 檢查檔案是否存在
                if (!File.Exists(_configFilePath))
                {
                    SetErrorState("找不到 config.txt 檔案");
                    return;
                }

                // 檢查檔案大小（防止 OOM 攻擊，限制最大 10KB）
                const long MAX_CONFIG_SIZE = 10 * 1024; // 10KB
                FileInfo fileInfo = new FileInfo(_configFilePath);
                if (fileInfo.Length > MAX_CONFIG_SIZE)
                {
                    SetErrorState($"config.txt 檔案過大（超過 {MAX_CONFIG_SIZE / 1024}KB 限制）");
                    return;
                }

                // 讀取檔案內容
                string content = File.ReadAllText(_configFilePath).Trim();
                if (string.IsNullOrWhiteSpace(content))
                {
                    SetErrorState("config.txt 檔案為空");
                    return;
                }

                // 解析格式：YYYY/MM/DD,標題文字
                string[] parts = content.Split(new[] { ',' }, 2);
                if (parts.Length != 2)
                {
                    SetErrorState("config.txt 格式錯誤（應為：YYYY/MM/DD,標題文字）");
                    return;
                }

                // 解析日期
                string dateString = parts[0].Trim();
                if (!DateTime.TryParseExact(dateString, "yyyy/MM/dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out _targetDate))
                {
                    SetErrorState($"日期格式錯誤：{dateString}（應為：YYYY/MM/DD）");
                    return;
                }

                // 取得標題
                _announcementTitle = parts[1].Trim();
                if (string.IsNullOrWhiteSpace(_announcementTitle))
                {
                    SetErrorState("標題文字不可為空");
                    return;
                }

                // 計算並更新顯示
                UpdateAnnouncementDisplay();
            }
            catch (FileNotFoundException ex)
            {
                SetErrorState($"設定檔未找到：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex}");
            }
            catch (UnauthorizedAccessException ex)
            {
                SetErrorState("無法訪問設定檔（權限不足）");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex}");
            }
            catch (IOException ex)
            {
                SetErrorState("讀取設定檔時發生 I/O 錯誤");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex}");
            }
            catch (OutOfMemoryException ex)
            {
                SetErrorState("記憶體不足，無法讀取設定檔");
                System.Diagnostics.Debug.WriteLine($"[CRITICAL] {ex}");
            }
            catch (Exception ex)
            {
                SetErrorState($"讀取設定檔時發生意外錯誤：{ex.GetType().Name}");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex}");
            }
        }

        /// <summary>
        /// 更新公告顯示內容（計算日期差）
        /// </summary>
        private void UpdateAnnouncementDisplay()
        {
            try
            {
                // 計算日期差（今日日期 - 目標日期）
                DateTime today = DateTime.Today;
                TimeSpan difference = today - _targetDate;
                int daysPassed = (int)difference.TotalDays;

                // 確保 UI 元素可用
                if (TitleTextBlock == null || DaysTextBlock == null || SuffixTextBlock == null)
                {
                    System.Diagnostics.Debug.WriteLine("[ERROR] UI 元素未初始化");
                    return;
                }

                // 更新 UI（透過資料繫結）
                TitleTextBlock.Text = _announcementTitle;
                DaysTextBlock.Text = daysPassed.ToString("N0", System.Globalization.CultureInfo.InvariantCulture);

                // 根據天數顯示不同訊息
                if (daysPassed < 0)
                {
                    SuffixTextBlock.Text = "天後到來";
                }
                else if (daysPassed == 0)
                {
                    SuffixTextBlock.Text = "就是今天！";
                }
                else
                {
                    SuffixTextBlock.Text = "天";
                }
            }
            catch (OverflowException ex)
            {
                SetErrorState("日期計算溢位");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex.Message}");
            }
            catch (FormatException ex)
            {
                SetErrorState("日期格式轉換失敗");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex.Message}");
            }
            catch (Exception ex)
            {
                SetErrorState($"計算日期時發生錯誤");
                System.Diagnostics.Debug.WriteLine($"[ERROR] {ex.GetType().Name} - {ex.Message}");
            }
        }

        /// <summary>
        /// 設定錯誤狀態顯示
        /// </summary>
        private void SetErrorState(string errorMessage)
        {
            try
            {
                // 安全地更新每個 UI 元素
                if (TitleTextBlock != null) TitleTextBlock.Text = "設定錯誤";
                if (DaysTextBlock != null) DaysTextBlock.Text = "!";
                if (SuffixTextBlock != null) SuffixTextBlock.Text = "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 設定錯誤狀態時發生異常：{ex.Message}");
            }

            // 記錄錯誤
            System.Diagnostics.Debug.WriteLine($"[ERROR] {errorMessage}");
        }

        /// <summary>
        /// 設定視窗層級為底層（在所有應用程式視窗的下方）
        /// </summary>
        private void SetWindowToBottom()
        {
            try
            {
                // 取得視窗 Handle
                IntPtr hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd == IntPtr.Zero)
                {
                    return;
                }

                // 設定視窗層級為 HWND_BOTTOM
                // SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE：保持當前大小、位置，且不啟動視窗
                bool success = NativeMethods.SetWindowPos(
                    hwnd,
                    NativeMethods.HWND_BOTTOM,
                    0, 0, 0, 0,
                    NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOMOVE | NativeMethods.SWP_NOACTIVATE
                );

                if (!success)
                {
                    System.Diagnostics.Debug.WriteLine("[WARN] SetWindowPos 返回失敗");
                }
            }
            catch (DllNotFoundException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WARN] Win32 DLL 未找到：{ex.Message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 設定視窗層級時發生錯誤：{ex.GetType().Name} - {ex.Message}");
            }
        }

        /// <summary>
        /// 計算並設定下一次凌晨 00:00 的更新時間
        /// </summary>
        private void ScheduleNextMidnightUpdate()
        {
            try
            {
                // 停止現有定時器
                _dateCheckTimer.Stop();

                // 取得當前時間
                DateTime now = DateTime.Now;

                // 計算下一個凌晨 00:00
                DateTime nextMidnight = now.Date.AddDays(1); // 明天的 00:00

                // 計算時間差
                TimeSpan timeUntilMidnight = nextMidnight - now;

                // 檢查時間間隔是否有效（應大於 0，小於 25 小時）
                if (timeUntilMidnight.TotalMilliseconds <= 0 || timeUntilMidnight.TotalHours > 25)
                {
                    System.Diagnostics.Debug.WriteLine("[WARN] 計算的下一次更新時間無效，使用 1 小時作為檢查間隔");
                    _dateCheckTimer.Interval = TimeSpan.FromHours(1);
                }
                else
                {
                    // 設定定時器間隔
                    _dateCheckTimer.Interval = timeUntilMidnight;
                }

                // 啟動定時器
                _dateCheckTimer.Start();

                // 記錄日誌（方便除錯）
                System.Diagnostics.Debug.WriteLine(
                    $"[INFO] 已排程下一次更新時間：{nextMidnight:yyyy-MM-dd HH:mm:ss}（{timeUntilMidnight.TotalHours:F2} 小時後）"
                );
            }
            catch (OverflowException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 時間計算溢位：{ex.Message}");
                // 使用預設的 1 小時檢查間隔
                _dateCheckTimer.Interval = TimeSpan.FromHours(1);
                _dateCheckTimer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 排程更新時間時發生錯誤：{ex.GetType().Name} - {ex.Message}");
            }
        }

        #endregion

        #region 資源清理

        /// <summary>
        /// 視窗關閉時清理資源
        /// </summary>
        protected override void OnClosed(EventArgs e)
        {
            try
            {
                // 停止並清理日期檢查定時器（DispatcherTimer 不實現 IDisposable）
                if (_dateCheckTimer != null)
                {
                    _dateCheckTimer.Stop();
                    _dateCheckTimer.Tick -= DateCheckTimer_Tick;
                }

                // 停止並清理視窗層級監控定時器
                if (_windowLevelCheckTimer != null)
                {
                    _windowLevelCheckTimer.Stop();
                    _windowLevelCheckTimer.Tick -= WindowLevelCheckTimer_Tick;
                }

                // 停止並清理視窗位置保存防抖定時器
                if (_savePositionDebounceTimer != null)
                {
                    _savePositionDebounceTimer.Stop();
                    _savePositionDebounceTimer.Tick -= SavePositionDebounceTimer_Tick;
                }

                // 最後保存一次位置（確保最終位置被保存）
                if (_hasPositionChanged)
                {
                    SaveWindowPosition();
                }

                // 取消事件訂閱
                this.Loaded -= MainWindow_Loaded;
                this.LocationChanged -= MainWindow_LocationChanged;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ERROR] 視窗關閉時發生異常：{ex.GetType().Name} - {ex.Message}");
            }
            finally
            {
                base.OnClosed(e);
            }
        }

        #endregion
    }
}
