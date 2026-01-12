# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**DesktopAnnouncement** is a WPF desktop announcement utility written in C# for .NET 8. It displays a countdown timer on the desktop background, showing days elapsed from a target date. The app automatically hides when switching away from the desktop and shows when returning.

### Key Features
- Win32 API monitoring for desktop state detection (GetForegroundWindow, GetClassName)
- Crystalline glass visual effect with transparency
- Single instance protection using Mutex
- Window position persistence in `position.txt`
- Configuration via `config.txt` (format: `YYYY/MM/DD,title text`)

## Build Commands

```bash
# Development build (Debug mode)
dotnet build

# Release build
dotnet build -c Release

# Publish as standalone executable (single-file, 64-bit)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# Windows batch script (uses dotnet internally)
./build.bat
```

## Architecture Overview

### Core Components

**NativeMethods.cs** - Win32 API wrapper layer
- Encapsulates P/Invoke declarations for user32.dll and kernel32.dll
- Key functions: `GetForegroundWindow()`, `GetClassName()`, `GetWindowThreadProcessId()`, `OpenProcess()`, `SetWindowPos()`
- Helper methods: `IsForegroundWindowDesktop()` checks if current foreground window is desktop
- All Win32 API calls wrapped with exception handling and proper resource cleanup (CloseHandle)

**App.xaml.cs** - Application lifecycle and single-instance management
- Mutex-based single instance enforcement in `OnStartup()`
- `TerminateExistingInstance()` gracefully closes old instances before starting new ones
- Unhandled exception handlers for both UI and background threads

**MainWindow.xaml.cs** - Core business logic
- Three DispatcherTimers:
  - `_dateCheckTimer`: Triggers at midnight daily to update countdown
  - `_windowLevelCheckTimer`: Every 2 seconds, keeps window at bottom Z-order
  - `_savePositionDebounceTimer`: Debounces position saves to disk (1 second delay)
- Desktop state monitoring: calls `NativeMethods.IsForegroundWindowDesktop()` every 500ms
- Configuration loading: `LoadConfiguration()` reads `config.txt`, validates date format
- Position persistence: `SaveWindowPosition()` and `LoadWindowPosition()` manage `position.txt`

**MainWindow.xaml** - UI with crystal glass effect
- Uses StaticResource for gradient and shadow effects
- Close button at top-right; main display area shows title + large countdown number
- Window properties: `WindowStyle="None"`, `AllowsTransparency="True"`, `ShowInTaskbar="False"`

### Data Files

- **config.txt**: Configuration file with format `YYYY/MM/DD,display title`. Required for app to function.
- **position.txt**: Auto-generated on first run, stores window X,Y coordinates. Delete to reset to default position (bottom-right).

## Key Design Patterns & Implementation Details

### Win32 API Usage
- Desktop detection checks for window class names "Progman" or "WorkerW" (desktop background windows)
- Also detects "explorer.exe" with Shell_TrayWnd class (taskbar scenario)
- Avoid holding references to window handles; re-query on each check
- Always use CloseHandle() in finally blocks to prevent handle leaks

### File I/O & Encoding
- **Current limitation**: File operations (ReadAllLines, WriteAllLines, ReadAllText) don't specify encoding. Should use `Encoding.UTF8` explicitly for consistency across platforms.
- Config file size limited to 10KB to prevent OOM attacks
- Position file format: two lines, first is X coordinate, second is Y coordinate

### Resource Management & Memory Leaks
All timers must be stopped and event handlers unsubscribed in `OnClosed()`:
- DispatcherTimer does NOT implement IDisposable; only Stop() and unsubscribe from Tick event
- LocationChanged event must be unsubscribed to prevent dangling handlers
- AppDomain exception events must be unsubscribed in OnExit()

### Thread Safety
- DispatcherTimer callbacks run on UI thread (no cross-thread issues)
- Win32 API calls are non-blocking; handle QueryProcesses with timeouts where applicable
- `_hasPositionChanged` bool is marked volatile for visibility across threads

## Common Development Tasks

### Modifying Countdown Display
Edit `MainWindow.xaml.cs`:
- `UpdateAnnouncementDisplay()` (line ~465): Controls what text displays (logic for past/future dates)
- DaysTextBlock in xaml shows the large number; SuffixTextBlock shows "天" (days) or custom text

### Changing Desktop Detection Logic
Edit `NativeMethods.cs`:
- `IsForegroundWindowDesktop()` (line ~214): Modify window class checks or add new desktop scenarios
- Current checks: Progman, WorkerW, explorer.exe + Shell_TrayWnd

### Adjusting Timer Intervals
Edit `MainWindow.xaml.cs`:
- `_windowLevelCheckTimer.Interval = TimeSpan.FromMilliseconds(2000)` (line ~79): Controls Z-order refresh rate
- `_savePositionDebounceTimer.Interval = TimeSpan.FromMilliseconds(1000)` (line ~88): Debounce delay for position saves
- `ScheduleNextMidnightUpdate()` (line ~578): Calculates next midnight for date refresh

### Configuration Format Validation
Edit `LoadConfiguration()` (line ~379):
- Format parsing happens at line 408 with `Split(',')`
- Date validation uses `DateTime.TryParseExact()` with format "yyyy/MM/dd"
- Add additional validation rules in this method

1. Win32 與 P/Invoke 安全
Handle 管理：所有回傳 IntPtr (Handle) 的 P/Invoke 呼叫，必須使用 try-finally 確保呼叫 CloseHandle()，或考慮重構成 SafeHandle 以防止資源洩漏。

字串編碼：Win32 API 必須統一標註 CharSet = CharSet.Unicode。

緩衝區防禦：使用 StringBuilder 接收 Win32 字串時（如 GetClassName），應先檢查回傳長度，避免固定長度（256 字符）造成的截斷或緩衝區溢位。

2. WPF 執行緒與記憶體安全
UI 執行緒隔離：雖然 DispatcherTimer 在 UI 執行緒執行，但任何涉及外部行程掃描（QueryProcesses）的邏輯若可能導致 UI 凍結，應改用 Task.Run 並透過 Dispatcher.InvokeAsync 回傳結果。

事件退訂：修復任何 UI 邏輯時，必須檢查 OnClosed 或 OnExit 是否已正確針對該物件進行 -= 事件退訂。

異步規範：所有的異步方法必須明確處理 Task（不可使用 async void，除非是事件處理常式），並考慮傳入 CancellationToken。

3. 檔案 I/O 強化
明確編碼：所有檔案讀寫（config.txt, position.txt）禁止使用預設編碼，必須明確傳入 System.Text.Encoding.UTF8。

併發控制：考慮到 TerminateExistingInstance 可能與新實體產生衝突，檔案寫入應加入簡單的 try-catch 或 FileShare 鎖定處理。

## Known Issues & High-Risk Areas (Prioritized)

### CRITICAL (Already Fixed)
- ✅ AppDomain exception events not unsubscribed in OnExit
- ✅ LocationChanged event subscription memory leak
- ✅ _savePositionDebounceTimer not properly cleared
- ✅ Process.GetCurrentProcess() not disposed

### HIGH (Should Fix)
- File encoding not specified (should use UTF8 explicitly)
- GetClassName StringBuilder fixed at 256 chars (may truncate long class names)
- GetForegroundWindow has race condition (window may be destroyed between call and next operation)
- SetWindowPos called every 2 seconds (high frequency, consider event-driven approach)

### MEDIUM
- Position validation allows window 95% off-screen (should require 50% on-screen)
- Concurrent file writes possible between TerminateExistingInstance and SaveWindowPosition
- DateTime calculation doesn't handle timezone changes
- WPF Effect resources use StaticResource (creates permanent global references)

## Testing

No automated tests exist. Manual testing checklist:
- Desktop detection: Switch between desktop and applications, verify app shows/hides
- Single instance: Launch app twice, verify old instance closes
- Position persistence: Move window, close, reopen - window should be in same position
- Configuration: Edit config.txt with valid/invalid dates, verify error handling
- Midnight update: Manually advance system clock to test daily refresh

## Build Output

```
bin/Release/net8.0-windows/win-x64/publish/DesktopAnnouncement.exe
```

This is a self-contained, single-file executable (~100MB with runtime). config.txt must be in the same directory.

## Important Notes for Future Work

1. **Nullable Reference Types**: Project has `<Nullable>enable</Nullable>`. All new code should handle nullability properly.
2. **P/Invoke Marshaling**: NativeMethods uses CharSet.Unicode for string marshaling. Be consistent with encoding.
3. **WPF Event Subscriptions**: Any new event subscriptions must be explicitly unsubscribed in OnClosed to prevent memory leaks.
4. **Win32 Handle Management**: Every Win32 handle obtained must have corresponding CloseHandle call in a finally block.
5. **Timer Safety**: DispatcherTimer callbacks are UI-thread-safe but ensure no blocking operations in tick handlers.
